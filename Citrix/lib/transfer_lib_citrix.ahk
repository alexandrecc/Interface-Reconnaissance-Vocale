#Requires AutoHotkey v2

global g_ctxBusy := false
global g_ctxVerboseLogs := false
global g_ctxPerfLogs := false
global g_ctxInitLogged := false
global g_ctxPerfState := Map()
global g_ctxCache := Map(
    "tabCtrl", "",
    "tabIndex", 0,
    "reqCtrl", "",
    "lotButton", "",
    "fenDict", "",
    "editCtrls", [],
    "winClass", ""
)
global g_ctxLastIdleLogTick := 0

CtxInitRuntime() {
    global g_ctxCache, g_ctxVerboseLogs, g_ctxPerfLogs, g_ctxPerfState

    ; Fast-mode defaults for UI automation in Citrix.
    SendMode("Input")
    SetWinDelay(0)
    SetControlDelay(-1)
    SetKeyDelay(-1, -1)
    SetMouseDelay(-1)
    SetTitleMatchMode(2)

    g_ctxVerboseLogs := CtxFlagFromEnv("CITRIX_VERBOSE_LOGS", false)
    g_ctxPerfLogs := CtxFlagFromEnv("CITRIX_PERF_LOGS", false)

    g_ctxCache["tabCtrl"] := ""
    g_ctxCache["tabIndex"] := 0
    g_ctxCache["reqCtrl"] := ""
    g_ctxCache["lotButton"] := ""
    g_ctxCache["fenDict"] := ""
    g_ctxCache["editCtrls"] := []
    g_ctxCache["winClass"] := ""
    g_ctxPerfState := Map()
}

CtxFlagFromEnv(name, defaultValue := false) {
    raw := StrLower(Trim(EnvGet(name)))
    if (raw = "")
        return defaultValue
    return (raw = "1" || raw = "true" || raw = "yes" || raw = "on")
}

CtxShouldRunPersistentWorker() {
    raw := StrLower(Trim(EnvGet("CITRIX_PERSISTENT_WORKER")))
    if (raw = "")
        return true
    return (raw = "1" || raw = "true" || raw = "yes")
}

CtxPollIntervalMs() {
    raw := Trim(EnvGet("CITRIX_POLL_MS"))
    if RegExMatch(raw, "^\d+$") {
        n := Integer(raw)
        if (n >= 100 && n <= 5000)
            return n
    }
    return 250
}

CtxFindRadImageHwnd(waitMs := 0) {
    deadline := A_TickCount + waitMs
    loop {
        hwnd := CtxFindRadImageHwndOnce()
        if hwnd
            return hwnd
        if (waitMs <= 0 || A_TickCount >= deadline)
            return 0
        Sleep 100
    }
}

CtxFindRadImageHwndOnce() {
    hint := Trim(EnvGet("CITRIX_RADIMAGE_HINT"))
    if (hint != "") {
        if hwnd := WinExist(hint)
            return hwnd
    }

    if hwnd := WinExist("ahk_exe RadImage.exe")
        return hwnd

    for hwnd in WinGetList() {
        ws := "ahk_id " hwnd
        try cls := WinGetClass(ws)
        catch
            continue
        if !InStr(cls, "WindowsForms10.Window.8.app.")
            continue

        title := ""
        proc := ""
        try title := WinGetTitle(ws)
        try proc := WinGetProcessName(ws)

        if (InStr(StrLower(title), "radimage") || StrLower(proc) = "radimage.exe" || InStr(StrLower(proc), "radimage")) {
            return hwnd
        }
    }

    return 0
}

Transfer_Citrix(invocation := "manual") {
    global g_ctxBusy, g_ctxLastIdleLogTick, g_ctxInitLogged

    if !g_ctxInitLogged {
        g_ctxInitLogged := true
        CtxLog("Citrix worker active. invocation=" invocation)
    }

    if g_ctxBusy
        return "busy"
    g_ctxBusy := true
    try {
        outcome := CtxProcessNextJob(invocation)
        if (outcome = "idle" && invocation != "poll")
            CtxLogVerbose("No queued Citrix transfer job.")
        else if (outcome = "idle" && invocation = "poll") {
            nowTick := A_TickCount
            if (nowTick - g_ctxLastIdleLogTick > 300000) {
                g_ctxLastIdleLogTick := nowTick
                CtxLogVerbose("Worker polling. No queued Citrix transfer job.")
            }
        }
        return outcome
    } catch as err {
        CtxLog("Unhandled Citrix transfer error: " err.Message)
        if (invocation != "poll")
            MsgBox "Erreur transfert Citrix:`n" err.Message, "Transfert Citrix", 0x10
        return "error"
    } finally {
        g_ctxBusy := false
    }
}

CtxProcessNextJob(invocation := "manual") {
    processStartTick := A_TickCount
    checkpoints := Map()
    dirs := CtxBridgeEnsureStructure()
    CtxPerfMark(checkpoints, "cp1", processStartTick)
    queueDir := dirs["queue"]
    hintDiag := CtxFindHintedJobDiag(queueDir)
    jobDir := hintDiag["jobDir"]
    CtxPerfMark(checkpoints, "cp2", processStartTick)
    if (jobDir = "" && invocation != "poll") {
        jobDir := CtxFindNextReadyJob(queueDir)
        CtxPerfMark(checkpoints, "queue_scan_done", processStartTick)
    } else {
        CtxPerfMark(checkpoints, "queue_scan_done", processStartTick)
    }
    if (jobDir = "")
        return "idle"

    jobId := CtxGetJobIdFromPath(jobDir)
    CtxPerfStartJob(jobId, invocation, processStartTick)
    CtxPerfSetHintDiag(hintDiag)
    CtxPerfMark(checkpoints, "cp3", processStartTick)

    ; Defer job (don't fail it) until RadImage is available in session.
    ; This avoids false TRANSFER_FAILED when hotkey/poller wakes before app publish is ready.
    if !CtxFindRadImageHwnd() {
        CtxLogVerbose("Job deferred (RadImage not available yet): " CtxGetJobIdFromPath(jobDir))
        CtxPerfMark(checkpoints, "find_radimage_done", processStartTick)
        CtxPerfFlushJob("deferred_radimage_missing", checkpoints)
        return "idle"
    }
    CtxPerfMark(checkpoints, "find_radimage_done", processStartTick)

    CtxClearJobHintIfMatches(queueDir, jobDir, hintDiag)
    CtxPerfMark(checkpoints, "clear_hint_done", processStartTick)

    lockPath := jobDir "\processing.lock"
    if !CtxTryCreateLock(lockPath) {
        CtxPerfMark(checkpoints, "lock_create_done", processStartTick)
        CtxPerfFlushJob("busy", checkpoints)
        return "busy"
    }
    CtxPerfMark(checkpoints, "lock_create_done", processStartTick)

    meta := Map()
    try {
        metaPath := jobDir "\meta.json"
        if !FileExist(metaPath)
            throw Error("meta.json introuvable pour job: " jobDir)

        meta := CtxBridgeReadFlatJson(metaPath)
        CtxPerfMark(checkpoints, "meta_read_done", processStartTick)
        if !meta.Has("jobId")
            meta["jobId"] := CtxGetJobIdFromPath(jobDir)
        jobId := meta["jobId"]
        CtxPerfSetJobId(jobId)
        CtxPerfMark(checkpoints, "cp4", processStartTick)
        CtxTransferJobToRadImage(jobDir, meta, invocation)
        CtxPerfMark(checkpoints, "cp5", processStartTick)
        CtxBridgeWriteDone(jobDir, Map(
            "jobId", meta["jobId"],
            "worker", A_ComputerName
        ))
        CtxLog("Job completed: " meta["jobId"])
        CtxPerfMark(checkpoints, "cp6", processStartTick)
        CtxPerfFlushJob("done", checkpoints)
        return "done"
    } catch as err {
        jobId := meta.Has("jobId") ? meta["jobId"] : CtxGetJobIdFromPath(jobDir)
        CtxBridgeWriteError(jobDir, Map(
            "jobId", jobId,
            "worker", A_ComputerName,
            "errorCode", "TRANSFER_FAILED",
            "message", err.Message
        ))
        CtxLog("Job failed: " jobId " - " err.Message)
        CtxPerfSetJobId(jobId)
        CtxPerfFlushJob("error", checkpoints)
        return "error"
    } finally {
        try {
            if FileExist(lockPath)
                FileDelete(lockPath)
        }
    }
}

CtxTransferJobToRadImage(jobDir, meta, invocation := "manual") {
    jobId := meta.Has("jobId") ? meta["jobId"] : CtxGetJobIdFromPath(jobDir)
    transferStart := A_TickCount

    reportPath := jobDir "\report.rtf"
    if !FileExist(reportPath)
        throw Error("report.rtf introuvable pour job: " jobDir)

    stage := A_TickCount
    localJob := CtxPrepareLocalJobFiles(jobDir, meta)
    effectiveJobDir := localJob["jobDir"]
    effectiveReportPath := localJob["reportPath"]
    CtxLogStage(jobId, "prepare_local_job", stage)

    stage := A_TickCount
    radHwnd := CtxFindRadImageHwnd(15000)
    if !radHwnd
        throw Error("RadImage.exe introuvable dans la session Citrix.")

    radWin := "ahk_id " radHwnd
    WinActivate(radWin)
    if !WinWaitActive(radWin, , 2.5)
        throw Error("Impossible d'activer RadImage.")
    CtxRefreshCacheForWindow()
    CtxLogStage(jobId, "activate_radimage", stage)

    tabIndex := CtxMetaTrue(meta, "casexterne") ? 2 : 3
    stage := A_TickCount
    ctx := CtxSelectTranscriptionTab(tabIndex, 3500)
    if !IsObject(ctx)
        throw Error("Impossible de selectionner l'onglet Transcription.")
    CtxLogStage(jobId, "select_tab", stage)

    stage := A_TickCount
    CtxValidateTarget(ctx, meta)
    CtxLogStage(jobId, "validate_target", stage)

    stage := A_TickCount
    if !CtxClickLotCourant(ctx)
        throw Error("Bouton Lot Courant introuvable.")
    CtxLogStage(jobId, "click_lot_courant", stage)

    Sleep 20
    Send "{F2}"

    stage := A_TickCount
    if !WinWaitActive("ahk_exe WINWORD.EXE", , 6)
        throw Error("Word n'est pas devenu actif apres F2.")
    CtxLogStage(jobId, "open_word", stage)

    stage := A_TickCount
    try word := ComObjActive("Word.Application")
    catch
        throw Error("Word.Application inaccessible.")

    doc := word.ActiveDocument
    oldSU := ""
    try oldSU := word.ScreenUpdating

    try {
        word.ScreenUpdating := False
        examCheck := CtxCheckExamList(doc, effectiveReportPath, effectiveJobDir)
        if (examCheck.Has("cancelled") && examCheck["cancelled"])
            throw Error("Transfert annulé par l'utilisateur lors de la vérification des titres.")

        doc.Content.Delete()
        sel := word.Selection
        sel.InsertFile(effectiveReportPath)
        if (examCheck.Has("missing") && examCheck["missing"].Length > 0) {
            if !CtxInsertMissingTitles(doc, examCheck["titres"], examCheck["missing"])
                throw Error("Transfert annulé lors de l'insertion des titres manquants.")
        }
        CtxRensMaj(doc)
        CtxCapitalizeParagraphStarts(doc)
        CtxLogStage(jobId, "insert_report", stage)

        stage := A_TickCount
        CtxAddFinalText(doc, meta, effectiveJobDir)
        CtxLogStage(jobId, "insert_signature", stage)
    } finally {
        try word.ScreenUpdating := oldSU
    }

    stage := A_TickCount
    Send "!{F4}"
    ; In Citrix, RadImage can regain focus before Word has fully closed.
    ; Wait for Word to actually disappear to avoid premature F8.
    WinWaitClose("ahk_exe WINWORD.EXE", , 4)
    Sleep 30

    if !WinWait(radWin, , 3)
        throw Error("RadImage n'est pas revenu apres fermeture de Word.")

    WinActivate(radWin)
    if !WinWaitActive(radWin, , 2.5)
        throw Error("RadImage inactif apres fermeture de Word.")
    CtxLogStage(jobId, "return_radimage", stage)

    stage := A_TickCount
    Send "{F8}"
    CtxLogStage(jobId, "send_f8", stage)
    CtxLogStage(jobId, "total", transferStart)
}

CtxPrepareLocalJobFiles(jobDir, meta := unset) {
    reportPath := jobDir "\report.rtf"
    CtxLogVerbose("Using mounted artifacts directly: " jobDir)
    return Map("jobDir", jobDir, "reportPath", reportPath)
}

CtxValidateTarget(ctx, meta) {
    global g_ctxCache

    expected := CtxNormalizeReq(meta.Has("reqnb") ? meta["reqnb"] : "")
    ; Always force a fresh req-control lookup per transfer (same window class can persist across exams).
    CtxClearReqReadCache()
    ; RadImage fields can lag briefly after a manual exam change in Citrix.
    current := CtxReadRadReqNormWithRetry(ctx, expected, 2200, 120)

    if (expected != "" && current != "" && expected != current) {
        throw Error(
            "Requête active RadImage différente du job.`n"
            . "Job: " expected "`n"
            . "RadImage: " current
        )
    }

    if !CtxMetaTrue(meta, "allowOverwrite") {
        fenDict := CtxResolveFenDictControl(ctx)
        if (fenDict != "")
            g_ctxCache["fenDict"] := fenDict
        existingText := CtxReadControlTextSafe(fenDict)
        if (Trim(existingText) != "")
            throw Error("L'examen RadImage contient déjà un rapport.")
    }
}

CtxClearReqReadCache() {
    global g_ctxCache
    g_ctxCache["reqCtrl"] := ""
}

CtxReadRadReqNormWithRetry(ctx, expectedNorm := "", maxMs := 2200, intervalMs := 120) {
    deadline := A_TickCount + (maxMs < 0 ? 0 : maxMs)
    lastNorm := ""

    loop {
        ; When we know the expected request, bypass cache to avoid stale control reuse.
        useCache := (expectedNorm = "")
        norm := CtxReadRadReqNorm(ctx, expectedNorm, useCache)
        if (norm != "") {
            lastNorm := norm
            if (expectedNorm = "" || norm = expectedNorm)
                return norm
        }

        if (A_TickCount >= deadline)
            break
        Sleep intervalMs
    }

    return lastNorm
}

CtxSelectTranscriptionTab(tabIndex, maxMs := 5000) {
    global g_ctxCache
    deadline := A_TickCount + maxMs
    lastDiagTick := 0

    ; Fast path: cached control/index from previous successful transfer.
    cachedCtrl := g_ctxCache["tabCtrl"]
    cachedIdx := g_ctxCache["tabIndex"]
    if (cachedCtrl != "") {
        for idx in CtxCandidateTabIndices(cachedIdx > 0 ? cachedIdx : tabIndex) {
            if CtxTrySelectTabControl(cachedCtrl, idx) {
                ctxFast := CtxGetRadImageContext()
                if !IsObject(ctxFast)
                    ctxFast := Map("tabCtrl", cachedCtrl, "fenDict", "", "editCtrls", [])
                ctxFast["tabCtrl"] := cachedCtrl
                ctxFast["usedTabIndex"] := idx
                g_ctxCache["tabCtrl"] := cachedCtrl
                g_ctxCache["tabIndex"] := idx
                CtxLogVerbose("Tab selected (cache). ctrl=" cachedCtrl " index=" idx)
                return ctxFast
            }
        }
    }

    while (A_TickCount < deadline) {
        ctx := CtxGetRadImageContext()
        if IsObject(ctx) {
            tabCtrls := CtxGetTabControls()
            if (tabCtrls.Length = 0 && ctx.Has("tabCtrl") && ctx["tabCtrl"] != "")
                tabCtrls.Push(ctx["tabCtrl"])

            indices := CtxCandidateTabIndices(tabIndex)

            for tabCtrl in tabCtrls {
                for idx in indices {
                    if CtxTrySelectTabControl(tabCtrl, idx) {
                        ctx["tabCtrl"] := tabCtrl
                        ctx["usedTabIndex"] := idx
                        g_ctxCache["tabCtrl"] := tabCtrl
                        g_ctxCache["tabIndex"] := idx
                        CtxLogVerbose("Tab selected. ctrl=" tabCtrl " index=" idx)
                        return ctx
                    }
                }
            }

            if (A_TickCount - lastDiagTick > 1200) {
                lastDiagTick := A_TickCount
                cls := ""
                if hwnd := CtxFindRadImageHwnd()
                    cls := WinGetClass("ahk_id " hwnd)
                CtxLogVerbose(
                    "Tab select retry. preferred=" tabIndex
                    . " controls=" tabCtrls.Length
                    . " class=" cls
                )
            }
        }
        Sleep 60
    }

    CtxLog("Tab select failed. " CtxBuildControlSummary())
    return 0
}

CtxGetRadImageContext() {
    global g_ctxCache

    ctx := CtxGetRadImageContextByPattern()
    if IsObject(ctx)
        return CtxApplyCacheToContext(ctx)
    ctx := CtxGetRadImageContextByScan()
    if IsObject(ctx)
        return CtxApplyCacheToContext(ctx)
    return 0
}

CtxGetRadImageContextByPattern() {
    hwnd := CtxFindRadImageHwnd()
    if !hwnd
        return 0

    winSpec := "ahk_id " hwnd
    winClass := WinGetClass(winSpec)
    appseg := RegExReplace(winClass, ".*\b(app\.[^_]+).*", "$1")
    if (appseg = "" || appseg = winClass)
        return 0

    return Map(
        "tabCtrl", "WindowsForms10.SysTabControl32." appseg "_r8_ad11",
        "fenDict", "WindowsForms10.RichEdit20W." appseg "_r8_ad11",
        "numReq27", "WindowsForms10.EDIT." appseg "_r8_ad127",
        "numReq30", "WindowsForms10.EDIT." appseg "_r8_ad130",
        "numReq48", "WindowsForms10.EDIT." appseg "_r8_ad148",
        "editCtrls", []
    )
}

CtxGetRadImageContextByScan() {
    hwnd := CtxFindRadImageHwnd()
    if !hwnd
        return 0
    ctrls := WinGetControls("ahk_id " hwnd)
    if !IsObject(ctrls) || ctrls.Length = 0
        return 0

    tabCtrl := ""
    fenDict := ""
    editCtrls := []

    for ctrl in ctrls {
        if (tabCtrl = "" && InStr(ctrl, "WindowsForms10.SysTabControl32."))
            tabCtrl := ctrl
        if (fenDict = "" && InStr(ctrl, "WindowsForms10.RichEdit20W."))
            fenDict := ctrl
        if InStr(ctrl, "WindowsForms10.EDIT.")
            editCtrls.Push(ctrl)
    }

    if (tabCtrl = "")
        return 0

    return Map(
        "tabCtrl", tabCtrl,
        "fenDict", fenDict,
        "numReq27", "",
        "numReq30", "",
        "numReq48", "",
        "editCtrls", editCtrls
    )
}

CtxApplyCacheToContext(ctx) {
    global g_ctxCache

    if (g_ctxCache["fenDict"] != "")
        ctx["fenDict"] := g_ctxCache["fenDict"]
    if (g_ctxCache["tabCtrl"] != "" && !ctx.Has("tabCtrl"))
        ctx["tabCtrl"] := g_ctxCache["tabCtrl"]

    return ctx
}

CtxTrySelectTabControl(tabCtrl, idx) {
    hwnd := CtxFindRadImageHwnd()
    if !hwnd
        return false
    win := "ahk_id " hwnd
    try {
        _ := ControlGetHwnd(tabCtrl, win)
        ControlChooseIndex(idx, tabCtrl, win)
        return true
    } catch {
        return false
    }
}

CtxResolveFenDictControl(ctx) {
    global g_ctxCache

    if IsObject(ctx) && ctx.Has("fenDict") && ctx["fenDict"] != ""
        return ctx["fenDict"]

    if (g_ctxCache["fenDict"] != "")
        return g_ctxCache["fenDict"]

    hwnd := CtxFindRadImageHwnd()
    if !hwnd
        return ""
    try ctrls := WinGetControls("ahk_id " hwnd)
    catch
        return ""

    for ctrl in ctrls {
        if InStr(ctrl, "WindowsForms10.RichEdit20W.")
            return ctrl
    }
    return ""
}

CtxReadRadReqNorm(ctx, expectedNorm := "", allowCached := true) {
    global g_ctxCache

    expectedNorm := CtxNormalizeReq(expectedNorm)

    ; Fast path: cached request control from previous success.
    cachedReqCtrl := g_ctxCache["reqCtrl"]
    if (allowCached && cachedReqCtrl != "") {
        valCached := CtxReadControlTextSafe(cachedReqCtrl)
        candCached := CtxReqCandidate(valCached)
        if (candCached["score"] >= 6)
            return candCached["norm"]
    }

    bestNorm := ""
    bestScore := -1
    bestCtrl := ""

    for key in ["numReq48", "numReq30", "numReq27"] {
        if !ctx.Has(key)
            continue
        ctrl := ctx[key]
        if (ctrl = "")
            continue
        val := CtxReadControlTextSafe(ctrl)
        cand := CtxReqCandidate(val)
        if (expectedNorm != "" && cand["norm"] = expectedNorm) {
            g_ctxCache["reqCtrl"] := ctrl
            return cand["norm"]
        }
        if (cand["score"] > bestScore) {
            bestScore := cand["score"]
            bestNorm := cand["norm"]
            bestCtrl := ctrl
            if (bestScore >= 6)
            {
                g_ctxCache["reqCtrl"] := bestCtrl
                return bestNorm
            }
        }
    }

    editCtrls := []
    if (ctx.Has("editCtrls") && IsObject(ctx["editCtrls"]) && ctx["editCtrls"].Length > 0)
        editCtrls := ctx["editCtrls"]
    else
        editCtrls := CtxGetEditControls()

    if (IsObject(editCtrls)) {
        for ctrl in editCtrls {
            val := CtxReadControlTextSafe(ctrl)
            cand := CtxReqCandidate(val)
            if (expectedNorm != "" && cand["norm"] = expectedNorm) {
                g_ctxCache["reqCtrl"] := ctrl
                return cand["norm"]
            }
            if (cand["score"] > bestScore) {
                bestScore := cand["score"]
                bestNorm := cand["norm"]
                bestCtrl := ctrl
            }
        }
    }

    if (bestScore >= 2) {
        if (bestCtrl != "")
            g_ctxCache["reqCtrl"] := bestCtrl
        return bestNorm
    }
    return ""
}

CtxBuildControlSummary() {
    hwnd := CtxFindRadImageHwnd()
    if !hwnd
        return "RadImage window missing."

    win := "ahk_id " hwnd
    className := WinGetClass(win)
    ctrls := WinGetControls(win)
    if !IsObject(ctrls)
        return "class=" className " controls=0"

    tabs := 0
    edits := 0
    rich := 0
    buttons := 0

    for ctrl in ctrls {
        if InStr(ctrl, "WindowsForms10.SysTabControl32.")
            tabs++
        if InStr(ctrl, "WindowsForms10.EDIT.")
            edits++
        if InStr(ctrl, "WindowsForms10.RichEdit20W.")
            rich++
        if InStr(ctrl, "WindowsForms10.BUTTON.")
            buttons++
    }

    return "class=" className " controls=" ctrls.Length " tabs=" tabs " edits=" edits " rich=" rich " buttons=" buttons
}

CtxGetTabControls() {
    global g_ctxCache

    tabs := []
    cached := g_ctxCache["tabCtrl"]
    if (cached != "")
        tabs.Push(cached)

    hwnd := CtxFindRadImageHwnd()
    if !hwnd
        return tabs
    try ctrls := WinGetControls("ahk_id " hwnd)
    catch
        return tabs

    for ctrl in ctrls {
        if InStr(ctrl, "WindowsForms10.SysTabControl32.") {
            if (tabs.Length = 0 || !CtxArrayHas(tabs, ctrl))
            tabs.Push(ctrl)
        }
    }

    return tabs
}

CtxGetEditControls() {
    global g_ctxCache

    cached := g_ctxCache["editCtrls"]
    if (IsObject(cached) && cached.Length > 0)
        return cached

    edits := []
    hwnd := CtxFindRadImageHwnd()
    if !hwnd {
        g_ctxCache["editCtrls"] := edits
        return edits
    }
    try ctrls := WinGetControls("ahk_id " hwnd)
    catch {
        g_ctxCache["editCtrls"] := edits
        return edits
    }

    for ctrl in ctrls {
        if InStr(ctrl, "WindowsForms10.EDIT.")
            edits.Push(ctrl)
    }
    g_ctxCache["editCtrls"] := edits
    return edits
}

CtxRefreshCacheForWindow() {
    global g_ctxCache

    hwnd := CtxFindRadImageHwnd()
    if !hwnd
        return

    cls := WinGetClass("ahk_id " hwnd)
    if (g_ctxCache["winClass"] = cls)
        return

    g_ctxCache["winClass"] := cls
    g_ctxCache["tabCtrl"] := ""
    g_ctxCache["tabIndex"] := 0
    g_ctxCache["reqCtrl"] := ""
    g_ctxCache["lotButton"] := ""
    g_ctxCache["fenDict"] := ""
    g_ctxCache["editCtrls"] := []
    CtxLogVerbose("Cache invalidated for new RadImage class: " cls)
}

CtxCandidateTabIndices(primary) {
    out := []
    seen := Map()

    for idx in [primary, 3, 2, 1, 4, 5] {
        if (idx < 1)
            continue
        key := "" idx
        if seen.Has(key)
            continue
        seen[key] := true
        out.Push(idx)
    }

    return out
}

CtxArrayHas(arr, needle) {
    for v in arr
        if (v = needle)
            return true
    return false
}

CtxNormalizeReq(value) {
    s := RegExReplace("" value, "[^\d]")
    n := StrLen(s)
    if (n < 8)
        return ""
    if (n > 12)
        return ""
    if (n > 8)
        s := SubStr(s, -8)
    return s
}

CtxReqCandidate(raw) {
    norm := CtxNormalizeReq(raw)
    if (norm = "")
        return Map("norm", "", "score", -1)

    score := 1
    if RegExMatch(raw, "\d{2}\s*-\s*\d{6}")
        score += 5
    else if RegExMatch(raw, "\d{4}\s*-\s*\d{4}")
        score += 4
    else if InStr(raw, "-")
        score += 2
    if RegExMatch(raw, "(^|[^\d])\d{8}([^\d]|$)")
        score += 1

    return Map("norm", norm, "score", score)
}

CtxReadControlTextSafe(ctrl) {
    hwnd := CtxFindRadImageHwnd()
    if !hwnd
        return ""
    try return ControlGetText(ctrl, "ahk_id " hwnd)
    catch
        return ""
}

CtxClickLotCourant(ctx := unset) {
    global g_ctxCache

    hwnd := CtxFindRadImageHwnd()
    if !hwnd
        return false
    win := "ahk_id " hwnd
    cachedBtn := g_ctxCache["lotButton"]
    if (cachedBtn != "") {
        try {
            ControlClick(cachedBtn, win)
            return true
        } catch {
        }
    }

    ctrls := WinGetControls(win)
    for ctrl in ctrls {
        if InStr(ctrl, "WindowsForms10.BUTTON.") {
            txt := ""
            try txt := ControlGetText(ctrl, win)
            if InStr(txt, "Lot Courant") {
                ControlClick(ctrl, win)
                g_ctxCache["lotButton"] := ctrl
                return true
            }
        }
    }
    return false
}

CtxFindNextReadyJob(queueDir) {
    if !DirExist(queueDir)
        return ""

    ; Keep scans bounded; old completed jobs should be archived/cleaned.
    maxDirs := 200
    seen := 0

    Loop Files, queueDir "\*", "D" {
        seen++
        if (seen > maxDirs)
            break

        jobDir := A_LoopFileFullPath
        ready := jobDir "\ready.flag"
        done := jobDir "\done.json"
        err := jobDir "\error.json"
        lock := jobDir "\processing.lock"

        if !FileExist(ready)
            continue
        if FileExist(done) || FileExist(err) || FileExist(lock)
            continue

        return jobDir
    }

    return ""
}

CtxFindHintedJob(queueDir) {
    info := CtxFindHintedJobDiag(queueDir)
    return info["jobDir"]
}

CtxFindHintedJobDiag(queueDir) {
    info := Map(
        "jobDir", "",
        "reason", "hint_unknown",
        "hintRaw", ""
    )

    hintPath := CtxHintFilePath(queueDir)
    if !FileExist(hintPath) {
        info["reason"] := "hint_file_missing"
        return info
    }

    try hintRaw := CtxNormalizeHintRaw(FileRead(hintPath, "UTF-8"))
    catch {
        info["reason"] := "hint_read_error"
        return info
    }

    info["hintRaw"] := hintRaw
    if (hintRaw = "") {
        info["reason"] := "hint_empty"
        return info
    }

    if (InStr(hintRaw, "\") || InStr(hintRaw, "/"))
        jobDir := hintRaw
    else
        jobDir := queueDir "\" hintRaw

    if !DirExist(jobDir) {
        info["reason"] := "hint_dir_missing"
        return info
    }

    ready := jobDir "\ready.flag"
    done := jobDir "\done.json"
    err := jobDir "\error.json"
    lock := jobDir "\processing.lock"

    if !FileExist(ready) {
        info["reason"] := "hint_ready_missing"
        return info
    }
    if FileExist(done) {
        info["reason"] := "hint_done_exists"
        return info
    }
    if FileExist(err) {
        info["reason"] := "hint_error_exists"
        return info
    }
    if FileExist(lock) {
        info["reason"] := "hint_lock_exists"
        return info
    }

    info["jobDir"] := jobDir
    info["reason"] := "hint_ok"
    return info
}

CtxClearJobHintIfMatches(queueDir, jobDir, hintDiag := unset) {
    hintPath := CtxHintFilePath(queueDir)
    if !FileExist(hintPath)
        return

    jobId := CtxGetJobIdFromPath(jobDir)

    ; Fast path: avoid an extra network read when current job came from valid hint.
    if IsSet(hintDiag) && IsObject(hintDiag) {
        if (hintDiag.Has("reason") && hintDiag["reason"] = "hint_ok" && hintDiag.Has("hintRaw")) {
            hinted := hintDiag["hintRaw"]
            if (hinted = jobId || hinted = jobDir) {
                try FileDelete(hintPath)
                return
            }
        }
    }

    try hintRaw := CtxNormalizeHintRaw(FileRead(hintPath, "UTF-8"))
    catch
        return

    if (hintRaw = "")
        return

    if (hintRaw = jobId || hintRaw = jobDir) {
        try FileDelete(hintPath)
    }
}

CtxHintFilePath(queueDir) {
    return queueDir "\_next_job.txt"
}

CtxNormalizeHintRaw(raw) {
    s := "" raw
    ; Trim spaces, tabs and line breaks; remove UTF-8 BOM if present.
    s := Trim(s, " `t`r`n")
    if (SubStr(s, 1, 1) = Chr(65279))
        s := SubStr(s, 2)
    return Trim(s, " `t`r`n")
}

CtxTryCreateLock(lockPath) {
    if FileExist(lockPath)
        return false
    try {
        FileAppend(CtxBridgeUtcNowIso() "`r`n", lockPath, "UTF-8")
        return true
    } catch {
        return false
    }
}

CtxResolveBridgeRoot() {
    envPath := Trim(EnvGet("RADEDITSYNC_DIR"))
    if (envPath != "")
        return envPath

    c1 := A_ScriptDir "\..\RadEditSync"
    if DirExist(c1)
        return c1

    c2 := A_ScriptDir "\RadEditSync"
    if DirExist(c2)
        return c2

    c3 := A_WorkingDir "\RadEditSync"
    if DirExist(c3)
        return c3

    return c1
}

CtxMetaTrue(meta, key) {
    if !meta.Has(key)
        return false
    return (StrLower(Trim(meta[key])) = "true")
}

CtxGetJobIdFromPath(jobDir) {
    return RegExReplace(jobDir, ".*[\\/]")
}

CtxLog(msg) {
    logPath := CtxLogPath()
    stamp := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
    try {
        FileAppend("[" stamp "] " msg "`r`n", logPath, "UTF-8")
    } catch {
        FileAppend("[" stamp "] " msg "`r`n", A_Temp "\citrix_transfer.log", "UTF-8")
    }
}

CtxLogVerbose(msg) {
    global g_ctxVerboseLogs
    if !g_ctxVerboseLogs
        return
    CtxLog(msg)
}

CtxLogStage(jobId, stageName, stageStartTick) {
    global g_ctxPerfLogs
    if !g_ctxPerfLogs
        return
    elapsed := A_TickCount - stageStartTick
    CtxPerfAppendStage(stageName, elapsed)
}

CtxPerfStartJob(jobId, invocation, processStartTick := 0) {
    global g_ctxPerfLogs, g_ctxPerfState
    if !g_ctxPerfLogs
        return

    if (processStartTick <= 0)
        processStartTick := A_TickCount

    g_ctxPerfState := Map(
        "jobId", jobId,
        "invocation", invocation,
        "processStartTick", processStartTick,
        "stages", [],
        "hintReason", "",
        "hintRaw", ""
    )
}

CtxPerfSetJobId(jobId) {
    global g_ctxPerfLogs, g_ctxPerfState
    if !g_ctxPerfLogs
        return
    if !IsObject(g_ctxPerfState) || !g_ctxPerfState.Has("jobId")
        return
    g_ctxPerfState["jobId"] := jobId
}

CtxPerfSetHintDiag(hintDiag) {
    global g_ctxPerfLogs, g_ctxPerfState
    if !g_ctxPerfLogs
        return
    if !IsObject(g_ctxPerfState) || !g_ctxPerfState.Has("jobId")
        return
    if !IsObject(hintDiag)
        return

    if (hintDiag.Has("reason"))
        g_ctxPerfState["hintReason"] := CtxPerfToken(hintDiag["reason"], 48)
    if (hintDiag.Has("hintRaw"))
        g_ctxPerfState["hintRaw"] := CtxPerfToken(hintDiag["hintRaw"], 80)
}

CtxPerfMark(checkpoints, checkpointName, processStartTick) {
    global g_ctxPerfLogs
    if !g_ctxPerfLogs || !IsObject(checkpoints)
        return
    checkpoints[checkpointName] := A_TickCount - processStartTick
}

CtxPerfAppendStage(stageName, elapsedMs) {
    global g_ctxPerfLogs, g_ctxPerfState
    if !g_ctxPerfLogs
        return
    if !IsObject(g_ctxPerfState) || !g_ctxPerfState.Has("stages")
        return
    g_ctxPerfState["stages"].Push(stageName "_ms=" elapsedMs)
}

CtxPerfToken(value, maxLen := 64) {
    s := Trim("" value)
    if (s = "")
        return ""
    s := RegExReplace(s, "[=\s]+", "_")
    if (StrLen(s) > maxLen)
        s := SubStr(s, 1, maxLen)
    return s
}

CtxPerfAddDelta(parts, checkpoints, fromKey, toKey, label) {
    if !IsObject(parts) || !IsObject(checkpoints)
        return
    if !checkpoints.Has(fromKey) || !checkpoints.Has(toKey)
        return
    delta := checkpoints[toKey] - checkpoints[fromKey]
    if (delta < 0)
        return
    parts.Push(label "=" delta)
}

CtxPerfFlushJob(status, checkpoints := unset) {
    global g_ctxPerfLogs, g_ctxPerfState
    if !g_ctxPerfLogs
        return
    if !IsObject(g_ctxPerfState) || !g_ctxPerfState.Has("jobId")
        return

    parts := []
    if IsSet(checkpoints) && IsObject(checkpoints) {
        for cp in ["cp1", "cp2", "cp3", "cp4", "cp5", "cp6"] {
            if checkpoints.Has(cp)
                parts.Push(cp "_ms=" checkpoints[cp])
        }
        CtxPerfAddDelta(parts, checkpoints, "cp1", "cp2", "cp1_to_cp2_ms")
        CtxPerfAddDelta(parts, checkpoints, "cp2", "cp3", "cp2_to_cp3_ms")
        CtxPerfAddDelta(parts, checkpoints, "cp3", "find_radimage_done", "find_radimage_ms")
        CtxPerfAddDelta(parts, checkpoints, "find_radimage_done", "clear_hint_done", "clear_hint_ms")
        CtxPerfAddDelta(parts, checkpoints, "clear_hint_done", "lock_create_done", "lock_create_ms")
        CtxPerfAddDelta(parts, checkpoints, "lock_create_done", "meta_read_done", "read_meta_ms")
        CtxPerfAddDelta(parts, checkpoints, "cp3", "cp4", "cp3_to_cp4_ms")
        CtxPerfAddDelta(parts, checkpoints, "cp4", "cp5", "cp4_to_cp5_ms")
        CtxPerfAddDelta(parts, checkpoints, "cp5", "cp6", "cp5_to_cp6_ms")
    }

    if (g_ctxPerfState.Has("hintReason") && g_ctxPerfState["hintReason"] != "")
        parts.Push("hint_reason=" g_ctxPerfState["hintReason"])
    if (g_ctxPerfState.Has("hintRaw") && g_ctxPerfState["hintRaw"] != "")
        parts.Push("hint_raw=" g_ctxPerfState["hintRaw"])

    stages := g_ctxPerfState["stages"]
    if (stages.Length > 0)
        parts.Push(CtxJoin(stages, " "))

    total := A_TickCount - g_ctxPerfState["processStartTick"]
    msg := "Perf[" g_ctxPerfState["jobId"] "] invocation=" g_ctxPerfState["invocation"] " status=" status " total_ms=" total
    if (parts.Length > 0)
        msg .= " " CtxJoin(parts, " ")
    CtxLog(msg)

    g_ctxPerfState := Map()
}

CtxLogPath() {
    try {
        root := CtxBridgeRoot()
        logsDir := root "\logs"
        if !DirExist(logsDir)
            DirCreate(logsDir)
        return logsDir "\citrix_transfer.log"
    } catch {
        return A_Temp "\citrix_transfer.log"
    }
}

CtxRensMaj(doc) {
    try {
        p1  := doc.Paragraphs.Item(1).Range
        txt := p1.Text

        if !RegExMatch(txt, "i)^\s*renseignements cliniques\s*:")
            return

        posColon := InStr(txt, ":")
        if (posColon = 0)
            return

        i := posColon + 1
        while (i <= StrLen(txt)) {
            ch := SubStr(txt, i, 1)
            if (ch = " " || ch = "`t" || Ord(ch) = 160) {
                i++
                continue
            }
            if (ch = "[" && SubStr(txt, i, 2) = "[]") {
                i += 2
                continue
            }
            break
        }
        if (i > StrLen(txt))
            return

        absStart := p1.Start + i - 1
        absEnd := absStart + 1

        r := doc.Range(absStart, absEnd)
        r.Case := 1

        one := r.Text
        up := StrUpper(one)
        if (up != one)
            r.Text := up
    } catch {
    }
}

CtxCapitalizeParagraphStarts(doc, maxScan := 40) {
    wdUpperCase := 1
    paras := doc.Paragraphs
    count := paras.Count

    Loop count {
        para := paras.Item(A_Index)
        r := para.Range.Duplicate

        if (r.End - r.Start <= 1)
            continue

        r.End -= 1
        start := r.Start
        end := r.End
        if (end <= start)
            continue

        sampleEnd := start + maxScan
        if (sampleEnd > end)
            sampleEnd := end

        sample := doc.Range(start, sampleEnd).Text
        if (sample = "")
            continue

        if !RegExMatch(sample, "\p{L}", &m)
            continue

        pos := start + (m.Pos[0] - 1)
        chR := doc.Range(pos, pos + 1)
        ch := chR.Text
        if RegExMatch(ch, "^\p{Ll}$")
            chR.Case := wdUpperCase
    }
}

CtxCheckExamList(doc, reportPath, jobDir := "") {
    textFile := CtxJobTextesPath(jobDir)
    if (textFile = "")
        textFile := CtxInsertionsPath("textesRapport.txt")

    defaultExcl := ["renseignements", "manipulation", "rapport", "sein", "reconstruction", "inj."]
    exclusions := CtxGetWordsByKey(textFile, "exclusions", defaultExcl)
    titres := []

    paras := doc.Paragraphs
    count := paras.Count
    Loop count {
        para := paras.Item(A_Index)
        txt := Trim(para.Range.Text, "`r`n `t")
        if (SubStr(txt, -1) = ":")
            txt := RTrim(txt, ":")
        if (txt = "")
            continue

        lower := StrLower(txt)
        excluded := false
        for excl in exclusions {
            if InStr(lower, excl) = 1 {
                excluded := true
                break
            }
        }
        if !excluded
            titres.Push(txt)
    }

    try {
        rawText := FileRead(reportPath, "CP1252")
    } catch as err {
        throw Error("Impossible de lire le fichier temporaire: " reportPath "`nErreur: " err.Message)
    }

    rawText := CtxDecodeRTFHex(rawText)

    rawText := StrReplace(rawText, "{\par", "`n")
    rawText := RegExReplace(rawText, "\\[a-z0-9]+(?:-?\d+)?[ ]?", "")
    rawText := RegExReplace(rawText, "\{\\[^}]+\}", "")
    rawText := RegExReplace(rawText, "[{}]", "")
    rawText := RegExReplace(rawText, "[\x00-\x08\x0B-\x1F]", "")

    lines := []
    for l in StrSplit(rawText, ["`r`n", "`n", "`r"]) {
        l := Trim(l, "`r`n `t")
        if (l != "")
            lines.Push(l)
    }

    needAttVer := false
    if (lines.Length > 0) {
        last := lines[lines.Length]
        last := StrReplace(last, Chr(160), " ")
        last := Trim(last, "`r`n `t")
        if RegExMatch(last, "i)^\s*Attention\s+(?:à|a)\s+v(?:é|e)rifier\s*$")
            needAttVer := true
    }

    missing := []
    for titre in titres {
        found := false
        titreNorm := StrLower(Trim(titre))
        for line in lines {
            lineTrim := StrLower(Trim(line))
            if (lineTrim = "")
                continue
            if (InStr(lineTrim, titreNorm) = 1) {
                found := true
                break
            }
        }
        if !found
            missing.Push(titre)
    }

    if (missing.Length > 0) {
        choice := CtxPromptMissingTitles(missing)
        if (choice = "continue")
            return Map("titres", titres, "missing", [], "needAttVer", needAttVer, "cancelled", false)
        if (choice = "insert")
            return Map("titres", titres, "missing", missing, "needAttVer", needAttVer, "cancelled", false)
        return Map("titres", titres, "missing", [], "needAttVer", needAttVer, "cancelled", true)
    }

    return Map("titres", titres, "missing", [], "needAttVer", needAttVer, "cancelled", false)
}

CtxPromptMissingTitles(missing) {
    msg := "Ces examens sont absents de votre dictée :`n`n"
    for titre in missing
        msg .= "- " titre "`n"
    msg .= "`nChoisissez l'option appropriée :`n`n"
    msg .= "1 - Continuer (ignorer et poursuivre)`n`n"
    msg .= "2 - Retour à la dictée (annuler le transfert)`n`n"
    msg .= "3 - Insérer le(s) titre(s) manquant(s)"

    dlg := Gui("+AlwaysOnTop", "Rapport incomplet")
    dlg.SetFont("s11", "Segoe UI")
    dlg.Add("Text", "w560 h420", msg)
    dlg.Show()

    ih := InputHook("L1")
    ih.Start()
    ih.Wait()
    choice := ih.Input
    dlg.Destroy()

    if (choice = "1")
        return "continue"
    if (choice = "3")
        return "insert"
    return "cancel"
}

CtxInsertMissingTitles(doc, titres, missing) {
    titresEligibles := []
    for t in titres {
        skip := false
        for m in missing {
            if (StrLower(Trim(t)) = StrLower(Trim(m))) {
                skip := true
                break
            }
        }
        if !skip
            titresEligibles.Push(t)
    }

    titresEligibles.Push("%TITRE%")

    for titre in missing {
        msg := "Titre manquant: " titre "`n`n"
        msg .= "Choisissez après quel titre l'insérer (appuyez sur le chiffre):`n`n"

        for i, t in titresEligibles {
            if (t = "%TITRE%")
                msg .= i ". Insérer à la place de: %TITRE%`n"
            else
                msg .= i ". " t "`n"
        }

        dlg := Gui("+AlwaysOnTop", "Insertion de titre manquant")
        dlg.SetFont("s11", "Segoe UI")
        dlg.Add("Text", "w760 h560", msg)
        dlg.Show()

        ih := InputHook("L1")
        ih.Start()
        ih.Wait()
        choice := ih.Input
        dlg.Destroy()

        if !RegExMatch(choice, "^\d+$")
            return false
        idx := choice + 0
        if (idx < 1 || idx > titresEligibles.Length)
            return false

        anchor := titresEligibles[idx]
        if (anchor = "%TITRE%") {
            rng := doc.Content
            f := rng.Find
            f.ClearFormatting()
            f.MatchWildcards := false
            f.MatchCase := false
            f.MatchWholeWord := false
            f.Text := "%TITRE%"

            if (f.Execute()) {
                rng.Text := StrUpper(titre)
                rng.Font.Bold := True
                rng.Font.Underline := 1
            } else {
                MsgBox "Aucun marqueur %TITRE% trouvé dans le document. Le titre n'a pas pu être inséré.", "Erreur", 0x10
                return false
            }
            continue
        }

        anchorNorm := StrLower(RTrim(Trim(anchor), ": "))
        pos := 0
        paras := doc.Paragraphs
        count := paras.Count
        Loop count {
            ptxt := paras.Item(A_Index).Range.Text
            pnorm := StrLower(RTrim(Trim(ptxt, "`r`n `t"), ": "))
            if (pnorm = anchorNorm) {
                pos := paras.Item(A_Index).Range.End
                break
            }
        }

        if (pos > 0) {
            r := doc.Range(pos, pos)
            r.Text := StrUpper(titre) . "`r" . r.Text
            r.Font.Bold := True
            r.Font.Underline := 1
        }
    }

    return true
}

CtxAddFinalText(doc, meta := unset, jobDir := "") {
    try word := doc.Application
    catch
        return false

    mode := ""
    if (IsSet(meta) && IsObject(meta) && meta.Has("mode"))
        mode := meta["mode"]

    sel := word.Selection
    sel.SetRange(doc.Content.End, doc.Content.End)
    sel.ParagraphFormat.Alignment := 0

    isSUG := InStr(StrUpper(mode), "SUG")

    chosen := CtxJobSignaturePath(jobDir, meta)
    if (chosen = "") {
        finalStd := CtxInsertionsPath("textefinal.rtf")
        finalSUG := CtxInsertionsPath("textefinalSUG.rtf")
        chosen := isSUG ? finalSUG : finalStd
    }

    if FileExist(chosen) {
        beforeEnd := doc.Content.End
        sel.InsertFile(chosen)
        if (doc.Content.End > beforeEnd) {
            CtxLogVerbose("Final text inserted from RTF: " chosen)
            sel.SetRange(doc.Content.End, doc.Content.End)
            sel.ParagraphFormat.Alignment := 0
            sel.Font.Name := "Arial"
            sel.Font.Size := 10
            return true
        }
    }

    CtxLogVerbose("Final RTF missing/unreadable. Using fallback text block.")
    return CtxAddFinalTextFallback(doc, sel, isSUG, jobDir)
}

CtxAddFinalTextFallback(doc, sel, isSUG := false, jobDir := "") {
    textFile := CtxJobTextesPath(jobDir)
    if (textFile = "")
        textFile := CtxInsertionsPath("textesRapport.txt")
    textSUG := CtxGetTextByKey(textFile, "textSUG"
        , "Examen demande, realise et interprete en urgence sur l'horaire de garde a la demande du medecin prescripteur.")
    authorLine := CtxGetTextByKey(textFile, "authorLine", "Dictee faite par Dr")
    sent1 := CtxGetTextByKey(textFile, "sent1", "Ce rapport a ete genere a l'aide d'un logiciel de reconnaissance vocale.")
    sent2 := CtxGetTextByKey(textFile, "sent2", "Pour cette raison il pourrait contenir des erreurs grammaticales, orthographiques ou syntaxiques.")

    if isSUG {
        fullText := "`r`n`r`n" textSUG "`r`n`r`n" authorLine " " CtxGetFullName_API() "`r`n`r`n" sent1 " " sent2
    } else {
        fullText := "`r`n`r`n" authorLine " " CtxGetFullName_API() "`r`n`r`n" sent1 " " sent2
    }

    sel.EndKey(6)
    startPos := sel.Start
    sel.TypeText(fullText)
    endPos := sel.Start

    blockRng := doc.Range(startPos, endPos)
    blockRng.Font.Name := "Arial"
    blockRng.Font.Size := 10
    blockRng.Font.Italic := True

    fr := blockRng.Duplicate
    fr.Find.ClearFormatting
    fr.Find.Text := sent1
    if fr.Find.Execute()
        fr.Font.Bold := True

    sel.SetRange(doc.Content.End, doc.Content.End)
    sel.Font.Name := "Arial"
    sel.Font.Size := 10
    return true
}

CtxJobSignaturePath(jobDir, meta := unset) {
    if (jobDir = "")
        return ""

    fileName := "textefinal.rtf"
    if (IsSet(meta) && IsObject(meta) && meta.Has("signatureFile") && Trim(meta["signatureFile"]) != "")
        fileName := Trim(meta["signatureFile"])

    candidate := jobDir "\" fileName
    if FileExist(candidate)
        return candidate
    return ""
}

CtxJobTextesPath(jobDir) {
    if (jobDir = "")
        return ""

    candidate := jobDir "\textesRapport.txt"
    if FileExist(candidate)
        return candidate
    return ""
}

CtxInsertionsPath(fileName) {
    ; Priority 1: exactly .\Textes\Insertions from current working directory.
    p1 := A_WorkingDir "\Textes\Insertions\" fileName
    if FileExist(p1)
        return p1

    ; Priority 2: repository root when script is in Citrix\lib.
    p2 := A_ScriptDir "\..\..\Textes\Insertions\" fileName
    if FileExist(p2)
        return p2

    ; Priority 3: alongside Citrix directory layout if mirrored.
    p3 := A_ScriptDir "\..\Textes\Insertions\" fileName
    if FileExist(p3)
        return p3

    root := ""
    try root := CtxBridgeRoot()
    catch
        root := ""
    if (root != "") {
        p4 := root "\..\Textes\Insertions\" fileName
        if FileExist(p4)
            return p4
    }

    return p1
}

CtxGetTextByKey(filePath, key, default := "") {
    if !FileExist(filePath)
        return default

    try text := FileRead(filePath, "UTF-8")
    catch
        return default

    for line in StrSplit(text, "`n", "`r") {
        parts := StrSplit(line, "=")
        if (parts.Length >= 2 && parts[1] = key)
            return Trim(parts[2])
    }
    return default
}

CtxGetWordsByKey(filePath, key, defaults := unset) {
    if !IsSet(defaults)
        defaults := []
    if !FileExist(filePath)
        return defaults

    try text := FileRead(filePath, "UTF-8")
    catch
        return defaults

    for line in StrSplit(text, "`n", "`r") {
        parts := StrSplit(line, "=")
        if (parts.Length >= 2 && parts[1] = key) {
            cleaned := Trim(parts[2])
            if (cleaned = "")
                return defaults
            words := []
            for w in StrSplit(cleaned, ",") {
                w := StrLower(Trim(w))
                if (w != "")
                    words.Push(w)
            }
            return words.Length > 0 ? words : defaults
        }
    }
    return defaults
}

CtxDecodeRTFHex(s) {
    out := ""
    pos := 1
    while RegExMatch(s, "\\'([0-9A-Fa-f]{2})", &m, pos) {
        out .= SubStr(s, pos, m.Pos(0) - pos)
        out .= Chr("0x" m[1])
        pos := m.Pos(0) + m.Len(0)
    }
    return out . SubStr(s, pos)
}

CtxGetFullName_API() {
    NameDisplay := 3
    len := 256
    buf := Buffer(len * 2)

    ok := DllCall("Secur32\GetUserNameExW"
        , "Int", NameDisplay
        , "Ptr", buf.Ptr
        , "UInt*", len)

    if ok {
        name := StrGet(buf.Ptr, "UTF-16")
        name := RegExReplace(name, "\d+")
        parts := StrSplit(name, " ")
        if (parts.Length = 2)
            return parts[2] " " parts[1]
        return name
    }

    return A_UserName
}


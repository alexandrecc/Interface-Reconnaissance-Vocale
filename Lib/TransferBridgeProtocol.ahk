#Requires AutoHotkey v2

BridgeProtocolVersion() {
    return "1.0"
}

BridgeUtcNowIso() {
    return FormatTime(A_NowUTC, "yyyy-MM-dd'T'HH:mm:ss'Z'")
}

BridgeRoot(rootDir := "") {
    if (rootDir = "")
        rootDir := A_ScriptDir "\RadEditSync"
    if RegExMatch(rootDir, "^[A-Za-z]:[\\/]?$")
        return SubStr(rootDir, 1, 2) "\"
    return RegExReplace(rootDir, "[\\/]+$")
}

BridgeQueueDir(rootDir := "") {
    return BridgeRoot(rootDir) "\queue"
}

BridgeArchiveDir(rootDir := "") {
    return BridgeRoot(rootDir) "\archive"
}

BridgeUseSingleSlotMode() {
    raw := StrLower(Trim(EnvGet("CITRIX_SINGLE_SLOT")))
    if (raw = "")
        return true
    return !(raw = "0" || raw = "false" || raw = "no" || raw = "off")
}

BridgeSingleSlotName() {
    raw := Trim(EnvGet("CITRIX_SLOT_NAME"))
    return (raw != "" ? raw : "slot_current")
}

BridgeSingleSlotDir(rootDir := "") {
    return BridgeQueueDir(rootDir) "\" BridgeSingleSlotName()
}

BridgeEnsureStructure(rootDir := "") {
    root := BridgeRoot(rootDir)
    queue := BridgeQueueDir(root)
    archive := BridgeArchiveDir(root)

    try DirCreate(root)
    try DirCreate(queue)
    try DirCreate(archive)
    if BridgeUseSingleSlotMode()
        try DirCreate(BridgeSingleSlotDir(root))

    return Map("root", root, "queue", queue, "archive", archive)
}

BridgeNewJobId(prefix := "job") {
    stamp := FormatTime(A_NowUTC, "yyyyMMdd_HHmmss")
    suffix := Format("{:06}", Random(0, 999999))
    return prefix "_" stamp "_" suffix
}

BridgeJobDir(jobId, rootDir := "") {
    return BridgeQueueDir(rootDir) "\" jobId
}

BridgeCreateJobFromFile(rootDir, meta, sourceReportPath, ioFlow := "") {
    if !IsObject(meta)
        throw Error("meta must be an object map.")
    if (sourceReportPath = "" || !FileExist(sourceReportPath))
        throw Error("source report file not found: " sourceReportPath)

    ioStart := A_TickCount
    ioLast := ioStart

    ; Structure is expected to be primed at startup (batch init).
    ; Keep a fallback ensure on failure when creating the job directory.
    BridgeLogJobIoTiming(ioFlow, "ensure_structure", 0, 0)

    jobId := meta.Has("jobId") && meta["jobId"] != "" ? meta["jobId"] : BridgeNewJobId()
    singleSlot := BridgeUseSingleSlotMode()
    jobDir := singleSlot ? BridgeSingleSlotDir(rootDir) : BridgeJobDir(jobId, rootDir)

    if singleSlot {
        try DirCreate(jobDir)
        catch {
            ensureStart := A_TickCount
            BridgeEnsureStructure(rootDir)
            ioNow := A_TickCount
            BridgeLogJobIoTiming(ioFlow, "ensure_structure_fallback", ioNow - ensureStart, ioNow - ioStart)
            DirCreate(jobDir)
        }
    } else {
        if DirExist(jobDir)
            throw Error("job directory already exists: " jobDir)
        try DirCreate(jobDir)
        catch {
            ensureStart := A_TickCount
            BridgeEnsureStructure(rootDir)
            ioNow := A_TickCount
            BridgeLogJobIoTiming(ioFlow, "ensure_structure_fallback", ioNow - ensureStart, ioNow - ioStart)
            DirCreate(jobDir)
        }
    }
    ioNow := A_TickCount
    BridgeLogJobIoTiming(ioFlow, "create_job_dir", ioNow - ioLast, ioNow - ioStart)
    ioLast := ioNow

    if !meta.Has("protocolVersion")
        meta["protocolVersion"] := BridgeProtocolVersion()
    if !meta.Has("createdAtUtc")
        meta["createdAtUtc"] := BridgeUtcNowIso()
    meta["jobId"] := jobId

    metaPath := jobDir "\meta.json"
    reportPath := jobDir "\report.rtf"
    readyPath := ""

    if singleSlot {
        lockPath := jobDir "\processing.lock"
        if FileExist(lockPath)
            throw Error("single-slot busy (processing.lock exists): " jobDir)

        BridgeDeleteIfExists(jobDir "\done.json")
        BridgeDeleteIfExists(jobDir "\error.json")
        BridgeDeleteIfExists(metaPath)
        BridgeDeleteIfExists(jobDir "\meta.tmp.json")

        ioNow := A_TickCount
        BridgeLogJobIoTiming(ioFlow, "clear_slot_files", ioNow - ioLast, ioNow - ioStart)
        ioLast := ioNow

        ; Write report first, then publish meta last as the ready/commit signal.
        FileCopy(sourceReportPath, reportPath, true)
        ioNow := A_TickCount
        BridgeLogJobIoTiming(ioFlow, "copy_report", ioNow - ioLast, ioNow - ioStart)
        ioLast := ioNow

        ; Direct write avoids UNC rename latency; Citrix side uses guarded read retries.
        BridgeWriteFlatJson(metaPath, meta)
        ioNow := A_TickCount
        BridgeLogJobIoTiming(ioFlow, "write_meta_json", ioNow - ioLast, ioNow - ioStart)
        BridgeLogJobIoTiming(ioFlow, "write_ready_flag", 0, ioNow - ioStart)
    } else {
        BridgeWriteFlatJson(metaPath, meta)
        ioNow := A_TickCount
        BridgeLogJobIoTiming(ioFlow, "write_meta_json", ioNow - ioLast, ioNow - ioStart)
        ioLast := ioNow

        FileCopy(sourceReportPath, reportPath, true)
        ioNow := A_TickCount
        BridgeLogJobIoTiming(ioFlow, "copy_report", ioNow - ioLast, ioNow - ioStart)
        ioLast := ioNow

        readyPath := jobDir "\ready.flag"
        readyFile := FileOpen(readyPath, "w", "UTF-8")
        if !IsObject(readyFile)
            throw Error("unable to open ready flag for write: " readyPath)
        readyFile.Write(BridgeUtcNowIso() "`r`n")
        readyFile.Close()
        ioNow := A_TickCount
        BridgeLogJobIoTiming(ioFlow, "write_ready_flag", ioNow - ioLast, ioNow - ioStart)
    }

    BridgeLogJobIoTiming(ioFlow, "job_create_total", ioNow - ioStart, ioNow - ioStart)

    return Map(
        "jobId", jobId,
        "jobDir", jobDir,
        "metaPath", metaPath,
        "reportPath", reportPath,
        "readyPath", readyPath
    )
}

BridgeDeleteIfExists(path) {
    if FileExist(path)
        FileDelete(path)
}

BridgeCleanupSlotArtifacts(jobDir) {
    if !BridgeUseSingleSlotMode()
        return
    try BridgeDeleteIfExists(jobDir "\meta.json")
    try BridgeDeleteIfExists(jobDir "\meta.tmp.json")
    try BridgeDeleteIfExists(jobDir "\report.rtf")
    try BridgeDeleteIfExists(jobDir "\done.json")
    try BridgeDeleteIfExists(jobDir "\error.json")
    try BridgeDeleteIfExists(jobDir "\ready.flag")
}

BridgeLogJobIoTiming(flow, stage, deltaMs, elapsedMs) {
    if (Trim(flow) = "")
        return

    stamp := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
    logPath := A_ScriptDir "\transfert_citrix_timing_io.log"
    line := stamp
        . ";flow=" flow
        . ";stage=" stage
        . ";delta_ms=" deltaMs
        . ";elapsed_ms=" elapsedMs
        . "`r`n"
    try FileAppend(line, logPath, "UTF-8")
}

BridgeWriteDone(jobDir, payload := unset) {
    if !DirExist(jobDir)
        throw Error("job directory not found: " jobDir)

    if !IsSet(payload)
        payload := Map()
    if !IsObject(payload)
        throw Error("payload must be an object map.")

    if !payload.Has("protocolVersion")
        payload["protocolVersion"] := BridgeProtocolVersion()
    if !payload.Has("completedAtUtc")
        payload["completedAtUtc"] := BridgeUtcNowIso()
    if !payload.Has("status")
        payload["status"] := "done"
    if !payload.Has("worker")
        payload["worker"] := "citrix"

    donePath := jobDir "\done.json"
    BridgeWriteFlatJson(donePath, payload)
    return donePath
}

BridgeWriteError(jobDir, payload) {
    if !DirExist(jobDir)
        throw Error("job directory not found: " jobDir)
    if !IsObject(payload)
        throw Error("payload must be an object map.")

    if !payload.Has("protocolVersion")
        payload["protocolVersion"] := BridgeProtocolVersion()
    if !payload.Has("completedAtUtc")
        payload["completedAtUtc"] := BridgeUtcNowIso()
    if !payload.Has("status")
        payload["status"] := "error"
    if !payload.Has("worker")
        payload["worker"] := "citrix"

    errPath := jobDir "\error.json"
    BridgeWriteFlatJson(errPath, payload)
    return errPath
}

BridgeWriteFlatJson(path, obj) {
    if !IsObject(obj)
        throw Error("obj must be an object map.")

    parts := []
    for key, value in obj {
        parts.Push('  "' BridgeJsonEscape("" key) '": "' BridgeJsonEscape(BridgeToString(value)) '"')
    }
    body := "{`r`n" . (parts.Length ? BridgeStrJoin(parts, ",`r`n") : "") . "`r`n}`r`n"
    file := FileOpen(path, "w", "UTF-8")
    if !IsObject(file)
        throw Error("unable to open json file for write: " path)
    file.Write(body)
    file.Close()
}

BridgeReadFlatJson(path) {
    if !FileExist(path)
        throw Error("json file not found: " path)

    text := FileRead(path, "UTF-8")
    out := Map()

    pos := 1
    while RegExMatch(text, '"((?:\\.|[^"\\])*)"\s*:\s*"((?:\\.|[^"\\])*)"', &m, pos) {
        key := BridgeJsonUnescape(m[1])
        val := BridgeJsonUnescape(m[2])
        out[key] := val
        pos := m.Pos(0) + m.Len(0)
    }
    return out
}

BridgeJsonEscape(s) {
    s := "" s
    s := StrReplace(s, "\", "\\")
    s := StrReplace(s, '"', '\"')
    s := StrReplace(s, "`r", "\r")
    s := StrReplace(s, "`n", "\n")
    s := StrReplace(s, "`t", "\t")
    return s
}

BridgeJsonUnescape(s) {
    s := StrReplace(s, "\r", "`r")
    s := StrReplace(s, "\n", "`n")
    s := StrReplace(s, "\t", "`t")
    s := StrReplace(s, '\"', '"')
    s := StrReplace(s, "\\", "\")
    return s
}

BridgeToString(value) {
    try {
        return "" value
    } catch {
        return ""
    }
}

BridgeStrJoin(arr, sep := "") {
    out := ""
    for idx, val in arr {
        if (idx > 1)
            out .= sep
        out .= val
    }
    return out
}

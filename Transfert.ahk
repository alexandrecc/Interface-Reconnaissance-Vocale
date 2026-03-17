#Requires AutoHotkey v2

#Include <AmenderRapport>
#Include <FinalTextInserts>
#Include <CheckExam>
#Include <AppendLogFiles>
#Include <GetData>
#Include <TransferShared>
#Include <TransferBridgeProtocol>
#Include <TransferArgumentRouting>

WM_COPYDATA := 0x004A
CMD := Map(
  "SetText",1, "InsertText",2, "SetFile",3, "InsertFile",4, "RequestTemp",5,
  "Response",6, "Error",7, "SetTitle",8, "GetTitle",9, "TitleResponse",10,
  "SetName",11, "GetName",12, "NameResponse",13, "GotoEnd",14,
  "FixFont",15, "CleanUpEnd",16, "SetHtmlFile", 17, "RequestHtmlFile", 18, "HtmlFileResponse", 19,
  "SetDataContext", 20, "GetDataContext", 21, "DataContextResponse", 22
)
global report_file := ""
global reqnb_name := ""
global reqnb_norm := ""
global reqnb_full := ""
global data_context_json := ""
global reqnb_title := ""

global g_dcRaw := ""          ; JSON brut en mÃ©moire
global g_dcCache := Map()     ; Map(key -> value)
global g_dcHwnd := 0          ; hwnd associÃ© au cache

global g_scriptStartTick := A_TickCount
global g_scriptStartIso := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
global g_transferArg := ""
OnMessage(WM_COPYDATA, CopyDataHandler)
textFile := ".\Textes\Insertions\textesRapport.txt"
finalRtf := ".\Textes\Insertions\textefinal.rtf"
finalRtfSUG := ".\Textes\Insertions\textefinalSUG.rtf"
flagFile := EnvGet("USERPROFILE") "\pause.flag"
global target
global downloadsDir := EnvGet("USERPROFILE") "\Downloads"
SetTitleMatchMode 2

;if A_Args.Length = 0 {
;    MsgBox "Aucun Argument fourni. Utiliser Ouvrir, Fermer, Tout, ToutCitrix, Signer ou Reculer."
;    ExitApp
;}

; Si aucun argument, on agit comme "Tout"
arg := (A_Args.Length >= 1) ? A_Args[1] : "Tout"
arg := ResolveTransferArgument(arg)
g_transferArg := arg


switch arg {
    case "Ouvrir":
        Ouvrir()
    case "OuvrirSUG":
        Ouvrir("OuvrirSUG")
    case "Fermer":
        WinActivate("ahk_exe WINWORD.EXE")
        WinWaitActive("ahk_exe WINWORD.EXE",,5)
        Sleep 50
        Fermer()
    case "Tout":
	Ouvrir()
	Fermer()
    case "ToutSUG":
	Ouvrir("ToutSUG")
	Fermer()
    case "ToutSUGCitrix":
        ToutCitrix("ToutSUG")
    case "ToutCitrix":
        ToutCitrix()
    case "Signer":
        Signer()
    case "SignerCitrix":
        SignerCitrix()
    case "Reculer":
        Reculer()
    case "Pathologie":
	AmenderRapport("Pathologie")
    case "Amender":
        AmenderRapport("Amender")
    case "Modification":
   	AmenderRapport("Modification")
    case "EffacerRadEdit":
	EraseRadEdit()
    case "GotoEnd":
	GotoEndofText()
    case "Basculer":
        Basculer()
    case "PauseSynchro":
        PauseSynchro()
    case "ToutSuivant":
	Ouvrir()
	Fermer()
	Signer()
    default:
        MsgBox "Invalid argument: " arg "`nUse Ouvrir, Fermer, Tout, ToutCitrix, Signer or Reculer."
}

ExitApp

Ouvrir(mode := "") {
    global reqnb_norm, reqnb_full, target  

; === PrÃ© requis:FenÃªtre de protocolage fermÃ©e
bloquante := "Module Protocolage"  

; VÃ©rifier si ouverte
if WinExist(bloquante) {
    WinClose bloquante                        
    if !WinWaitClose(bloquante, , 2) {        
        MsgBox "La fenÃªtre bloquante est encore ouverte. Le script va s'arrÃªter.", 262144
        ExitApp
    }
}

; === Step 1: Activer la fenÃªtre Radimage ===
if !WinExist("ahk_exe RadImage.exe") {
    MsgBox "FenÃªtre Radimage introuvable ou inactive. Le script va se fermer.", 262144
    ExitApp
}

WinActivate("ahk_exe RadImage.exe")
WinWaitActive("ahk_exe RadImage.exe",,5)

maxMs    := 5000    ; dÃ©lai total max en ms (augmente si Radimage est lent)
interval := 150      ; pause entre essais
deadline := A_TickCount + maxMs
ok := false
lastErr := ""

while (A_TickCount < deadline) {
    winClass := WinGetClass(WinExist("ahk_exe RadImage.exe"))
    appseg   := RegExReplace(winClass, ".*\b(app\.[^_]+).*", "$1")
    tabCtrl  := "WindowsForms10.SysTabControl32." appseg "_r8_ad11"
    fenDict  := "WindowsForms10.RichEdit20W." appseg "_r8_ad11"
    numReq27   := "WindowsForms10.EDIT." appseg "_r8_ad127"    
    numReq30   := "WindowsForms10.EDIT." appseg "_r8_ad130"
    numReq48   := "WindowsForms10.EDIT." appseg "_r8_ad148"
    tabIndex := HasArg("CasExterne") ? 2 : 3
    radText := ""

found := false
    Loop 15 {
        Try {
            radText := ControlGetText(numReq48, "ahk_exe RadImage.exe")
            if (radText != "") {
                found := true
                break
            }
        } catch as err {
	}
        Try {
            radText := ControlGetText(numReq30, "ahk_exe RadImage.exe")
            if (radText != "") {
                found := true
                break
            }
        } catch as err {
	}
        Try {
            radText := ControlGetText(numReq27, "ahk_exe RadImage.exe")
            if (radText != "") {
                found := true
                break
            }
        } catch as err {
	}
        Sleep 200  ; attendre avant de rÃ©essayer
    }
    if !found {
        MsgBox "Impossible de trouver le numÃ©ro de requÃªte car RadImage est trop lent Ã  rÃ©agir. SVP rÃ©essayer dans quelques secondes."
	        ExitApp
    }

    rad_norm := StrReplace(radText, "-") 

; === VÃ©rification concordance ReqNb ===
target := WinExist("RadEdit ahk_exe RadEdit.exe")
if !target {
    MsgBox "RadEdit introuvable."
    ExitApp
}

PrimeDataContextCached(target)

formcomplete := Trim(GetDataContextVarCached(target, "formcomplete", ""))
if (formcomplete = "false") {
    MsgBox "Le rapport n'a pas Ã©tÃ© gÃ©nÃ©rÃ© par le formulaire."
    ExitApp
}

try {
    ;reqnb_full := GetTitleFromRadEdit(target)
    reqnb_full := Trim(GetDataContextVarCached(target, "reqnb", ""))
    if (reqnb_full = "") {
        if ProcessExist("FusionDictate.exe") {
	    SendCopyData(target, CMD["SetText"], "{\rtf1\ansi}")
            ExitApp
        } else {
            MsgBox "Impossible de rÃ©cupÃ©rer la requÃªte depuis RadEdit.`nImpossible de vÃ©rifier la concordance."
            ExitApp
        }
    }
    reqnb_norm := SubStr(reqnb_full, -8)  

    if (reqnb_norm != rad_norm) {
        result := MsgBox(
            "La requÃªte de l'examen actif ne correspond pas Ã  la requÃªte de l'examen dictÃ©.`n"
            . "Synapse â†’ " reqnb_norm "`n"
            . "RadImage â†’ " rad_norm "`n`n"
            . "Voulez-vous continuer quand mÃªme ?",
            "VÃ©rification", 0x40024) ; Yes/No + AlwaysOnTop

        if (result = "No") {
            ExitApp
        } else {
            ; === VÃ©rification supplÃ©mentaire : rapport dÃ©jÃ  prÃ©sent ? ===
            text := CheckFenDict(fenDict)
            if (Trim(text) != "") {
                SoundPlay "*16"
                MsgBox "L'examen sÃ©lectionnÃ© contient dÃ©jÃ  un rapport.`n`n"
                    . "Le script s'arrÃªtera.`n`n"
                    . "SVP sÃ©lectionnez le bon examen.",
                    "ATTENTION", 262144
                ExitApp
            }
        }
    }
}
catch as err {
    MsgBox "Erreur lors de la vÃ©rification ReqNb:`n" err.Message
    ExitApp
}

    try {
        _ := ControlGetHwnd(tabCtrl, "ahk_exe RadImage.exe")
        ControlChooseIndex(tabIndex, tabCtrl, "ahk_exe RadImage.exe")  ;tabIndex (default=3, CasExterne=2)
        ok := true
        break
    }
    catch as err {
        lastErr := err.Message
        Sleep interval
    }
}

if !ok {
    MsgBox "Impossible de sÃ©lectionner l'onglet 'Transcription' aprÃ¨s " maxMs " ms.`nDerniÃ¨re erreur: " lastErr
    ExitApp
}
Sleep 20

if !ClickLotCourant() {
    MsgBox "Bouton 'Lot Courant' introuvable. Le script s'arrÃªte."
    ExitApp
}

Sleep 20
Send "{F2}"



; === Step 2: Wait for Word window to appear ===
if !WinWaitActive("ahk_exe WINWORD.EXE", , 10) {
    MsgBox "Word did not open or become active in time. Script will exit."
    ExitApp
}


; === Step 3: Check Exam List ===

source := WinExist("RadEdit ahk_exe RadEdit.exe")
if !source {
    MsgBox "RadEdit window not found."
    ExitApp
}

SendCopyData(source, CMD["CleanUpEnd"], "")

keepfont := Trim(GetDataContextVarCached(target, "keepfont", ""))
if (keepfont != "true") {
    SendCopyData(source, CMD["FixFont"], "Arial;10")
}


SendCopyData(source, CMD["RequestTemp"], '{ "stripHiddenMarkers": true }')

; Attendre que report_file soit rempli
maxWait := 1000   ; attendre max 1 seconde
waited := 0
while (report_file == "" && waited < maxWait) {
    Sleep 50
    waited += 50
}
if (report_file = "") {
    MsgBox "RadEdit nâ€™a pas retournÃ© de fichier temporaire."
    ExitApp
}

try word := ComObjActive("Word.Application")
catch {
    MsgBox "Word n'est pas ouvert."
    ExitApp
}

doc := word.ActiveDocument
result := CheckExamList(doc, report_file)  

undo := word.UndoRecord
oldSU := word.ScreenUpdating
word.ScreenUpdating := False
undo.StartCustomRecord("TransfertRapport")

TransfertStandard(word, doc, report_file, result, mode)

undo.EndCustomRecord()
word.ScreenUpdating := oldSU

ResetRadEdit(target, report_file, reqnb_full)

;fin ouvrir
}


TransfertStandard(word, doc, report_file, result, mode) {
    ; --- contenu identique Ã  ce que tu avais dans le try de Ouvrir() ---

    doc.Content.Delete()
    sel := word.Selection
    sel.InsertFile(report_file)

    if (result.missing.Length > 0) {
        InsertMissingTitles(doc, report_file, result.titres, result.missing)
    }

    RensMaj(doc)
    CapitalizeParagraphStarts(doc)

    attv := HasArg("AttV")

    didAttVer := false
    if (attv) {
        InsertManualAttVer(doc)
        didAttVer := true
    } else if (result.needAttVer) {
        AttVer(doc)
        didAttVer := true
    }

    LogFile(target, result, didAttVer)
    AddText(mode)
}

ToutCitrix(mode := "ToutCitrix") {
    global CMD, report_file, reqnb_full
    perf := CitrixPerfStart("ToutCitrix:" mode)

    source := WinExist("RadEdit ahk_exe RadEdit.exe")
    if !source {
        MsgBox "RadEdit introuvable."
        ExitApp
    }

    CitrixPerfMark(perf, "source_found")
    PrimeDataContextCached(source)
    CitrixPerfMark(perf, "data_context_primed")
    formcomplete := Trim(GetDataContextVarCached(source, "formcomplete", ""))
    if (formcomplete = "false") {
        MsgBox "Le rapport n'a pas Ã©tÃ© gÃ©nÃ©rÃ© par le formulaire."
        ExitApp
    }

    reqnb_full := Trim(GetDataContextVarCached(source, "reqnb", ""))
    if (reqnb_full = "") {
        MsgBox "Impossible de rÃ©cupÃ©rer la requÃªte depuis RadEdit."
        ExitApp
    }

    CitrixPerfMark(perf, "precheck_done")
    SendCopyData(source, CMD["CleanUpEnd"], "")
    keepfont := Trim(GetDataContextVarCached(source, "keepfont", ""))
    if (keepfont != "true")
        SendCopyData(source, CMD["FixFont"], "Arial;10")

    CitrixPerfMark(perf, "cleanup_font_done")
    report_file := ""
    SendCopyData(source, CMD["RequestTemp"], '{ "stripHiddenMarkers": true }')
    CitrixPerfMark(perf, "request_temp_sent")
    if !WaitForRadEditTempFile(2000) {
        MsgBox "RadEdit n'a pas retournÃ© de fichier temporaire pour le transfert Citrix."
        ExitApp
    }

    CitrixPerfMark(perf, "temp_file_ready")
    bridgeRoot := ResolveBridgeRoot()
    meta := BuildCitrixTransferMeta(source, mode)
    CitrixPerfMark(perf, "bridge_meta_ready")
    primedCitrixHwnd := PrimeCitrixHostWindowFast()

    try job := BridgeCreateJobFromFile(bridgeRoot, meta, report_file, "ToutCitrix:" mode)
    catch as err {
        MsgBox "Impossible de crÃ©er la tÃ¢che Citrix.`n`nErreur: " err.Message
        ExitApp
    }

    CitrixPerfMark(perf, "job_created")
    if CitrixShouldAttachJobArtifacts(bridgeRoot, mode) {
        AttachCitrixJobArtifacts(job["jobDir"], mode, meta)
        CitrixPerfMark(perf, "artifacts_attached")
    } else {
        CitrixPerfMark(perf, "artifacts_skipped")
    }
    hintWritten := WriteCitrixNextJobHint(bridgeRoot, job["jobId"])
    CitrixPerfMark(perf, hintWritten ? "hint_written" : "hint_skipped")

    TriggerCitrixTransferHotkey(perf, primedCitrixHwnd)

    result := WaitForBridgeTerminal(job["jobDir"], 90000, job["jobId"])
    status := result.Has("status") ? result["status"] : "timeout"

    if (status = "done") {
        wasCancelled := result.Has("cancelled") && (StrLower(Trim(result["cancelled"])) = "true")
        if wasCancelled {
            BridgeCleanupSlotArtifacts(job["jobDir"])
            HandleCitrixCancelledTransferLocally(source)
            return
        }

        LogCitrixAttVerIfNeeded(source, result)
        LogFileCitrix(source)
        ResetRadEdit(source, report_file, reqnb_full)
        BridgeCleanupSlotArtifacts(job["jobDir"])
        return
    }

    if (status = "error") {
        msg := result.Has("message") ? result["message"] : "Erreur Citrix inconnue."
        code := result.Has("errorCode") ? result["errorCode"] : "UNKNOWN"
        BridgeCleanupSlotArtifacts(job["jobDir"])
        MsgBox "Le transfert Citrix a Ã©chouÃ©.`nCode: " code "`nMessage: " msg
        ExitApp
    }

    MsgBox "Aucune rÃ©ponse Citrix aprÃ¨s 90 secondes.`nLa tÃ¢che est conservÃ©e dans:`n" job["jobDir"]
    ExitApp
}

HandleCitrixCancelledTransferLocally(source) {
    MsgBox "Transfert annule. Vous pouvez maintenant continuer votre dictee.", "TERMINE", 262144

    win := "RadEdit ahk_exe RadEdit.exe"
    if WinExist(win)
        WinActivate(win)
    if WinWaitActive(win, , 5)
        Send "{PgDn}"
}

LogCitrixAttVerIfNeeded(source, bridgeResult) {
    global target
    if !IsObject(bridgeResult)
        return

    didAttVer := false
    if bridgeResult.Has("didAttVer") {
        didAttVer := (StrLower(Trim(bridgeResult["didAttVer"])) = "true")
    } else if HasArg("AttV") {
        ; Compatibilite: ancien worker Citrix sans didAttVer dans done.json.
        didAttVer := true
    }

    if !didAttVer
        return

    ; LogAttVer lit le data context via la variable globale target.
    target := source

    loc := StrLower(Trim(GetDataContextVarCached(source, "loc", "")))
    modal := StrUpper(Trim(GetDataContextVarCached(source, "modal", "")))
    if !(loc = "urgence" && modal = "CR")
        return

    titres := []
    if bridgeResult.Has("attVerTitres") {
        raw := Trim(bridgeResult["attVerTitres"])
        if (raw != "") {
            for t in StrSplit(raw, "||") {
                t := Trim(t)
                if (t != "")
                    titres.Push(t)
            }
        }
    }

    if !LogAttVer(titres) {
        MsgBox "ATTV appliquÃ©, mais impossible d'Ã©crire dans le fichier de suivi urgence (`R:\...\ATTENTION Ã€ VÃ‰RIFIER.txt`)."
        return
    }
}

SignerCitrix() {
    perf := CitrixPerfStart("SignerCitrix")
    primedCitrixHwnd := PrimeCitrixHostWindowFast()
    if CitrixUseDirectSignerMode() {
        CitrixPerfMark(perf, "signer_mode_direct")
        TriggerCitrixSignerHotkey(perf, primedCitrixHwnd)
        if !SendF8ToLocalSynapse() {
            MsgBox "Commande Signer envoyÃ©e dans Citrix, mais Synapse local est introuvable pour envoyer F8."
            ExitApp
        }
        return
    }

    CitrixPerfMark(perf, "signer_mode_bridge")
    bridgeRoot := ResolveBridgeRoot()
    meta := Map(
        "mode", "SignerCitrix",
        "stripHiddenMarkers", "false",
        "keepfont", "false",
        "attv", "false",
        "casexterne", "false",
        "signatureFile", ""
    )
    CitrixPerfMark(perf, "bridge_meta_ready")

    dummyRtf := A_Temp "\radedit_signer_" FormatTime(A_NowUTC, "yyyyMMdd_HHmmss") "_" Format("{:06}", Random(0, 999999)) ".rtf"
    try {
        FileAppend("{\rtf1\ansi}", dummyRtf, "UTF-8")
        job := BridgeCreateJobFromFile(bridgeRoot, meta, dummyRtf, "SignerCitrix")
    } catch as err {
        MsgBox "Impossible de crÃ©er la tÃ¢che Citrix Signer.`n`nErreur: " err.Message
        ExitApp
    } finally {
        try {
            if FileExist(dummyRtf)
                FileDelete(dummyRtf)
        }
    }
    CitrixPerfMark(perf, "job_created")

    hintWritten := WriteCitrixNextJobHint(bridgeRoot, job["jobId"])
    CitrixPerfMark(perf, hintWritten ? "hint_written" : "hint_skipped")
    TriggerCitrixTransferHotkey(perf, primedCitrixHwnd)

    result := WaitForBridgeTerminal(job["jobDir"], 30000, job["jobId"])
    status := result.Has("status") ? result["status"] : "timeout"

    if (status = "done") {
        BridgeCleanupSlotArtifacts(job["jobDir"])
        if !SendF8ToLocalSynapse() {
            MsgBox "Commande Signer envoyÃ©e dans Citrix, mais Synapse local est introuvable pour envoyer F8."
            ExitApp
        }
        return
    }

    if (status = "error") {
        msg := result.Has("message") ? result["message"] : "Erreur Citrix inconnue."
        code := result.Has("errorCode") ? result["errorCode"] : "UNKNOWN"
        BridgeCleanupSlotArtifacts(job["jobDir"])
        MsgBox "Le mode Signer Citrix a Ã©chouÃ©.`nCode: " code "`nMessage: " msg
        ExitApp
    }

    MsgBox "Aucune rÃ©ponse Citrix aprÃ¨s 30 secondes.`nLa tÃ¢che est conservÃ©e dans:`n" job["jobDir"]
    ExitApp
}

CitrixSignerMode() {
    raw := StrLower(Trim(EnvGet("CITRIX_SIGNER_MODE")))
    if (raw = "")
        return "direct"
    if (raw = "bridge" || raw = "job")
        return "bridge"
    return "direct"
}

CitrixUseDirectSignerMode() {
    return (CitrixSignerMode() != "bridge")
}

WaitForRadEditTempFile(timeoutMs := 1200) {
    global report_file
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if (report_file != "" && FileExist(report_file))
            return true
        Sleep 40
    }
    return false
}

ResolveBridgeRoot() {
    envPath := Trim(EnvGet("RADEDITSYNC_DIR"))
    if (envPath != "")
        return envPath

    citrixDataRoot := "\\regional.reg14.rtss.qc.ca\app\DragonMedicalOne\Radiologie\Citrix Data"
    userSubdir := Trim(A_UserName)
    return (userSubdir != "" ? citrixDataRoot "\" userSubdir : citrixDataRoot)
}

BuildCitrixTransferMeta(source, mode := "ToutCitrix") {
    global reqnb_full

    keepfont := Trim(GetDataContextVarCached(source, "keepfont", ""))
    meta := Map(
        "mode", mode,
        "reqnb", reqnb_full,
        "patdos", Trim(GetDataContextVarCached(source, "patdos", "")),
        "proc", Trim(GetDataContextVarCached(source, "proc", "")),
        "modal", Trim(GetDataContextVarCached(source, "modal", "")),
        "loc", Trim(GetDataContextVarCached(source, "loc", "")),
        "studydate", Trim(GetDataContextVarCached(source, "studydate", "")),
        "stripHiddenMarkers", "true",
        "keepfont", (keepfont = "true" ? "true" : "false"),
        "attv", (HasArg("AttV") ? "true" : "false"),
        "casexterne", (HasArg("CasExterne") ? "true" : "false"),
        "formcomplete", Trim(GetDataContextVarCached(source, "formcomplete", "")),
        "signatureFile", "textefinal.rtf"
    )
    return meta
}

AttachCitrixJobArtifacts(jobDir, mode, meta := unset) {
    global finalRtf, finalRtfSUG, textFile

    isSUG := InStr(StrUpper(mode), "SUG")
    sigSrc := ResolveScriptPath(isSUG ? finalRtfSUG : finalRtf)
    sigDestName := "textefinal.rtf"

    if IsSet(meta) && IsObject(meta) && meta.Has("signatureFile") && Trim(meta["signatureFile"]) != ""
        sigDestName := Trim(meta["signatureFile"])

    if FileExist(sigSrc) {
        try FileCopy(sigSrc, jobDir "\" sigDestName, false)
    }

    txtSrc := ResolveScriptPath(textFile)
    if FileExist(txtSrc) {
        try FileCopy(txtSrc, jobDir "\textesRapport.txt", false)
    }
}

CitrixShouldAttachJobArtifacts(bridgeRoot := "", mode := "") {
    raw := StrLower(Trim(EnvGet("CITRIX_ATTACH_JOB_ARTIFACTS")))
    if (raw = "1" || raw = "true" || raw = "yes" || raw = "on")
        return true
    if (raw = "0" || raw = "false" || raw = "no" || raw = "off")
        return false

    ; Auto mode: skip per-job artifact copy when shared cache already has required files.
    return !CitrixArtifactsCachedForMode(bridgeRoot, mode)
}

CitrixArtifactsCachedForMode(bridgeRoot, mode := "") {
    cacheDir := CitrixCachedInsertionsDir(bridgeRoot)
    if (cacheDir = "")
        return false

    textPath := cacheDir "\textesRapport.txt"
    if !FileExist(textPath)
        return false

    isSUG := InStr(StrUpper(mode), "SUG")
    sigFile := isSUG ? "textefinalSUG.rtf" : "textefinal.rtf"
    return (FileExist(cacheDir "\" sigFile) != "")
}

CitrixCachedInsertionsDir(bridgeRoot := "") {
    root := Trim(bridgeRoot)
    if (root = "")
        return ""
    ; bridgeRoot is ...\Citrix Data\<username>; cached artifacts are user-specific in ...\<username>\Textes\Insertions
    return root "\Textes\Insertions"
}

ResolveScriptPath(path) {
    if (path = "")
        return ""
    if InStr(path, ":")
        return path
    cleaned := RegExReplace(path, "^\.(\\|/)", "")
    return A_ScriptDir "\" cleaned
}

WriteCitrixNextJobHint(bridgeRoot, jobId) {
    if (jobId = "")
        return false
    if BridgeUseSingleSlotMode()
        return false
    try {
        queueDir := BridgeQueueDir(bridgeRoot)
        if !DirExist(queueDir)
            return false
        hintPath := queueDir "\_next_job.txt"
        file := FileOpen(hintPath, "w", "UTF-8")
        if !IsObject(file)
            return false
        file.Write(jobId)
        file.Close()
        return true
    } catch {
        return false
    }
}

WaitForBridgeTerminal(jobDir, timeoutMs := 90000, expectedJobId := "") {
    donePath := jobDir "\done.json"
    errPath := jobDir "\error.json"
    deadline := A_TickCount + timeoutMs

    while (A_TickCount < deadline) {
        if FileExist(donePath) {
            try {
                out := BridgeReadFlatJson(donePath)
                if (expectedJobId != "" && (!out.Has("jobId") || Trim(out["jobId"]) != expectedJobId)) {
                    Sleep 50
                    continue
                }
                out["status"] := "done"
                return out
            } catch {
            }
        }
        if FileExist(errPath) {
            try {
                out := BridgeReadFlatJson(errPath)
                if (expectedJobId != "" && (!out.Has("jobId") || Trim(out["jobId"]) != expectedJobId)) {
                    Sleep 50
                    continue
                }
                out["status"] := "error"
                return out
            } catch {
            }
        }
        Sleep 200
    }

    return Map("status", "timeout")
}

PrimeCitrixHostWindowFast() {
    hwnd := FindCitrixHostWindow()
    if hwnd && !WinActive("ahk_id " hwnd) {
        try WinActivate("ahk_id " hwnd)
    }
    return hwnd
}

TriggerCitrixTransferHotkey(perf := unset, primedHwnd := 0) {
    TriggerCitrixMappedHotkey("^!+t", perf, primedHwnd, "hotkey_sent")
}

TriggerCitrixSignerHotkey(perf := unset, primedHwnd := 0) {
    TriggerCitrixMappedHotkey("^!+s", perf, primedHwnd, "hotkey_sent")
}

TriggerCitrixMappedHotkey(keyCombo, perf := unset, primedHwnd := 0, perfStage := "hotkey_sent") {
    global g_scriptStartTick
    hwnd := primedHwnd ? primedHwnd : FindCitrixHostWindow()
    if hwnd {
        primedReady := (primedHwnd && WinActive("ahk_id " hwnd))
        if !primedReady && !WinActive("ahk_id " hwnd) {
            WinActivate("ahk_id " hwnd)
            WinWaitActive("ahk_id " hwnd, , 2)
        }
        if !primedReady {
            settleMs := CitrixPostActivateSleepMs()
            if (settleMs > 0)
                Sleep settleMs
        }
        if IsSet(perf)
            CitrixPerfMark(perf, "citrix_window_ready")
    }

    elapsedBeforeSendMs := A_TickCount - g_scriptStartTick
    Send keyCombo
    elapsedAfterSendMs := A_TickCount - g_scriptStartTick
    LogCitrixHotkeyDispatchTiming(elapsedBeforeSendMs, elapsedAfterSendMs, hwnd)
    if IsSet(perf)
        CitrixPerfMark(perf, perfStage)
}

CitrixPostActivateSleepMs() {
    raw := Trim(EnvGet("CITRIX_POST_ACTIVATE_SLEEP_MS"))
    if (raw = "")
        return 30
    ms := Round(raw + 0)
    return (ms < 0 ? 0 : ms)
}

LogCitrixHotkeyDispatchTiming(elapsedBeforeSendMs, elapsedAfterSendMs, hwnd) {
    global g_scriptStartIso, g_transferArg
    logPath := A_ScriptDir "\transfert_citrix_timing.log"
    stamp := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
    hwndText := (hwnd ? Format("0x{:X}", hwnd) : "0")
    line := stamp
        . ";arg=" g_transferArg
        . ";script_start=" g_scriptStartIso
        . ";elapsed_before_send_ms=" elapsedBeforeSendMs
        . ";elapsed_after_send_ms=" elapsedAfterSendMs
        . ";citrix_hwnd=" hwndText
        . "`r`n"
    try FileAppend(line, logPath, "UTF-8")
}

CitrixPerfStart(flowName) {
    tick := A_TickCount
    return Map("flow", flowName, "startTick", tick, "lastTick", tick)
}

CitrixPerfMark(perf, stageName) {
    global g_scriptStartTick, g_transferArg
    if !IsObject(perf)
        return

    now := A_TickCount
    deltaMs := now - perf["lastTick"]
    flowElapsedMs := now - perf["startTick"]
    scriptElapsedMs := now - g_scriptStartTick
    perf["lastTick"] := now

    stamp := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
    logPath := A_ScriptDir "\transfert_citrix_timing_stages.log"
    line := stamp
        . ";arg=" g_transferArg
        . ";flow=" perf["flow"]
        . ";stage=" stageName
        . ";delta_ms=" deltaMs
        . ";flow_elapsed_ms=" flowElapsedMs
        . ";script_elapsed_ms=" scriptElapsedMs
        . "`r`n"
    try FileAppend(line, logPath, "UTF-8")
}

FindCitrixHostWindow() {
    hint := Trim(EnvGet("CITRIX_WINDOW_HINT"))
    if (hint != "") {
        if hwnd := WinExist(hint)
            return hwnd
    }

    if hwnd := WinExist("ahk_exe wfica32.exe")
        return hwnd
    if hwnd := WinExist("ahk_class Transparent Windows Client")
        return hwnd

    return 0
}


Fermer() {
Send "!{F4}"    ; Alt + F4 (close Word)
Sleep 50       ; short pause

; === Step 2: Wait max 5s for Radimage window to exist ===
if !WinWait("ahk_exe RadImage.exe", , 5) {
    MsgBox "Radimage window did not appear in time. Script will exit."
    ExitApp
}

; === Step 3: Activate Radimage and send F8 ===
WinActivate      ; bring Radimage to front
WinWaitActive    ; wait until itâ€™s actually active
Send "{F8}"
; fin Fermer
}

Signer() {
if WinExist("ahk_exe RadImage.exe") {
    WinActivate
}
if !WinWaitActive("ahk_exe RadImage.exe", , 5) {
    MsgBox "Radimage window not found or not active. Script will exit."
    ExitApp
}
    Sleep 20
    Send "^g"

    if !SendF8ToLocalSynapse()
        ExitApp
    ;fin Signer
    }

SendF8ToLocalSynapse() {
    title := "v5.7"
    hwnd := WinExist(title " ahk_exe msedge.exe")
    if !hwnd
        return false

    WinActivate("ahk_id " hwnd)
    if !WinWaitActive("ahk_id " hwnd, , 5)
        return false

    Sleep 20
    Send "{F8}"
    return true
}

Reculer() {
if WinExist("ahk_exe RadImage.exe") {
    WinActivate
}
if !WinWaitActive("ahk_exe RadImage.exe", , 5) {
    MsgBox "Radimage window not found or not active. Script will exit."
    ExitApp
}
Send "{F2}"
}


ClickLotCourant() {
    win := "ahk_exe RadImage.exe"
    ctrls := WinGetControls(win)

    for ctrl in ctrls {
        if InStr(ctrl, "WindowsForms10.BUTTON.") {
            txt := ""
            try txt := ControlGetText(ctrl, win)
            if InStr(txt, "Lot Courant") {
                ControlClick ctrl, win
                return true
            }
        }
    }
    return false
}

CheckFenDict(fenDict, win := "ahk_exe Radimage.exe") {
    maxTry := 5
    Loop maxTry {
        try {
            return ControlGetText(fenDict, win)
        }
        catch {
            Sleep 200
        }
    }
    SoundPlay "*16"
    MsgBox "Impossible de vÃ©rifier si le rapport de l'examen est vide.`n`nVous Ãªtes probablement dans le mauvais examen.`n`nLe script va s'arrÃªter.", "ERREUR", 262144
    ExitApp
}

FixFontsInWord() {
    try {
        word := ComObjActive("Word.Application")  ; Get running Word instance
    } catch {
        MsgBox "Word not found."
        return false
    }

    doc := word.ActiveDocument
    rng := doc.Content
    rng.Font.Name := "Arial"
    rng.Font.Size   := 10
    rng.Font.ColorIndex := 1   ; 1 = noir
    return true
}

CopyDataHandler(wParam, lParam, msg, hwnd) {
    static CMD_RESPONSE := 6
    static CMD_ERROR := 7
    static CMD_TITLERESPONSE := 10
    static CMD_NAMERESPONSE := 13
    static CMD_HTMLFILERESPONSE := 19
    static CMD_DATACONTEXTRESPONSE := 22

    global report_file, html_file, reqnb_title, reqnb_name, last_radedit_error, data_context_json

    cmd  := NumGet(lParam, 0, "UPtr")
    size := NumGet(lParam, A_PtrSize, "UInt")
    text := StrGet(NumGet(lParam, 2*A_PtrSize, "Ptr"), size/2, "UTF-16")
    text := Trim(text)

    if (cmd = CMD_RESPONSE) {
        report_file := text
    } else if (cmd = CMD_HTMLFILERESPONSE) {
        html_file := text
    } else if (cmd = CMD_TITLERESPONSE) {
        reqnb_title := text
    } else if (cmd = CMD_NAMERESPONSE) {
        reqnb_name := text
    } else if (cmd = CMD_DATACONTEXTRESPONSE) {
        data_context_json := text
    } else if (cmd = CMD_ERROR) {
        last_radedit_error := text
        ; MsgBox "RadEdit error: " text  ; (optionnel) je le laisse commentÃ© pour Ã©viter les popups
    }
    return true
}

SendCopyData(hwnd, command, text := "") {
    text := text . Chr(0)
    buf := Buffer(StrLen(text) * 2, 0)
    StrPut(text, buf, "UTF-16")
    cds := Buffer(3 * A_PtrSize, 0)
    NumPut("UPtr", command, cds, 0)
    NumPut("UInt", buf.Size, cds, A_PtrSize)
    NumPut("Ptr", buf.Ptr, cds, 2 * A_PtrSize)
    return DllCall("user32\SendMessageW", "Ptr", hwnd, "UInt", WM_COPYDATA, "Ptr", A_ScriptHwnd, "Ptr", cds.Ptr, "Ptr")
}


SendToRadImage(key) {
    if !WinWaitActive("ahk_exe RadImage.exe", , 1) {
        WinActivate "ahk_exe RadImage.exe"
        WinWaitActive "ahk_exe RadImage.exe", , 1
    }
    Send key
}

CancelTransfer() {
    try WinClose("ahk_exe WINWORD.EXE")
    SendToRadImage("{Enter}")
    SendToRadImage("{F4}")
    if WinWait("Confirmation", , 1) {
        WinActivate "Confirmation"
        if WinWaitActive("Confirmation", , 1) {
            Send "{o}"   ; confirme avec la touche "o"
        }
    }

    MsgBox "Vous pouvez maintenant aller continuer votre dictÃ©e.", "TERMINÃ‰", 262144

    if WinExist("RadEdit ahk_exe RadEdit.exe") {
	WinActivate
	}
    if !WinWaitActive("RadEdit ahk_exe RadEdit.exe", , 5) {
	MsgBox "RadEdit window not found or not active. Script will exit."
	ExitApp
    }
    Send "{PgDn}"
    ExitApp
}

ResetRadEdit(target, tempFile, reqnb_full) {
    global CMD
    ; 1) Vider RadEdit
    SendCopyData(target, CMD["SetText"], "{\rtf1\ansi}")
    SendCopyData(target, CMD["SetTitle"], "")
    SendCopyData(target, CMD["SetName"], "")
    SendCopyData(target, CMD["SetDataContext"], "")
    SendCopyData(target, CMD["SetHtmlFile"], "")

    ; 2) Supprimer le fichier temporaire
    try {
        if (tempFile != "" && FileExist(tempFile)) {
            FileDelete(tempFile)
        }
    } catch as err {
        MsgBox "Impossible d'effacer le fichier temporaire :`n" tempFile "`n`nErreur : " err.Message, "ATTENTION", 262144
    }
    ; 3) Supprimer les HTML temporaires associÃ©s (RadEdit_<safeReqnb>_*.html)
    if (reqnb_full != "") {
        safeReqnb := RegExReplace(reqnb_full, "[^\w-]", "_")
        pattern := A_ScriptDir "\Textes\HTML\RadEdit_" safeReqnb "*.html"

        try {
            Loop Files, pattern {
                try FileDelete(A_LoopFileFullPath)
            }
        } catch as err {
            ; optionnel: MsgBox si tu veux Ãªtre averti d'un problÃ¨me global de loop
            ; MsgBox "Erreur nettoyage HTML temp:`n" err.Message, "ATTENTION", 262144
        }
    }
}

EraseRadEdit() {
    global CMD, reqnb_title
    target := WinExist("RadEdit ahk_exe RadEdit.exe")
    if !target {
        MsgBox "RadEdit introuvable.", "Erreur", 262144
        return
    }
    WinActivate "RadEdit ahk_exe RadEdit.exe"
    WinWaitActive "RadEdit ahk_exe RadEdit.exe", , 1

    reqnb_title := GetTitleFromRadEdit(target)

    ; 1) Effacer contenu RadEdit
    SendCopyData(target, CMD["SetText"], "{\rtf1\ansi}")
    SendCopyData(target, CMD["SetTitle"], "")
    SendCopyData(target, CMD["SetName"], "")
    SendCopyData(target, CMD["SetDataContext"], "")
    SendCopyData(target, CMD["SetHtmlFile"], "")

    ; 2) Supprimer le fichier temporaire associÃ© si prÃ©sent
    if (reqnb_title != "") {
        tempFile := A_Temp "\RadEdit_" reqnb_title ".rtf"
        try {
            if FileExist(tempFile) {
                FileDelete(tempFile)
                }
        } catch as err {
            MsgBox "Impossible d'effacer le fichier temporaire :`n" tempFile "`n`nErreur : " err.Message, "ATTENTION", 262144
        }
    }
    FileAppend("", EnvGet("USERPROFILE") "\tcp_listener.reset")
}


GotoEndofText(){
	target := WinExist("RadEdit ahk_exe RadEdit.exe")
	if !target {
	    MsgBox("âŒ RadEdit non trouvÃ©.")
 	   ExitApp()
	}

	SendCopyData(target, CMD["GotoEnd"], "")
	ExitApp()
}

Basculer() {
    if FileExist(flagFile) {
        FileDelete(flagFile)
        Resume()
    }
    else {
        FileAppend("", flagFile) 
        Pause()
    }
}

PauseSynchro() {
    if FileExist(flagFile) {
        FileDelete(flagFile)
        target := WinExist("RadEdit ahk_exe RadEdit.exe")
        ;SendCopyData(target, CMD["SetText"], "{\rtf1\ansi}")	
        SendCopyData(target, CMD["SetName"], "")
        ;SendCopyData(target, CMD["SetTitle"], "")
    }
    else {
        FileAppend("", flagFile) 
        target := WinExist("RadEdit ahk_exe RadEdit.exe")
        SendCopyData(target, CMD["SetName"], "PAUSE SYNCHRO RADEDIT")
        ;SendCopyData(target, CMD["SetText"], "{\rtf1\ansi}")	
        ;SendCopyData(target, CMD["SetTitle"], "")
        }
}

Pause() {
    target := WinExist("RadEdit ahk_exe RadEdit.exe")
    if !ProcessExist("FusionDictate.exe")
        Run '"C:\APP\Fusion_StartUp\startup_imagerie.cmd"', , "Hide"
    SendCopyData(target, CMD["SetName"], "PAUSE SYNCHRO RADEDIT - SYNCHRO FUSION ACTIF")
    SendCopyData(target, CMD["SetText"], "{\rtf1\ansi}")	
    SendCopyData(target, CMD["SetTitle"], "")
    SendCopyData(target, CMD["SetDataContext"], "")
}

Resume() {
    target := WinExist("RadEdit ahk_exe RadEdit.exe")
    SendCopyData(target, CMD["SetText"], "{\rtf1\ansi}")	
    SendCopyData(target, CMD["SetName"], "")
    SendCopyData(target, CMD["SetTitle"], "")
    if ProcessExist("FusionDictate.exe")
       ProcessClose("FusionDictate.exe")
}

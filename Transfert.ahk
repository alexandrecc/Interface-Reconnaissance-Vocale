#Requires AutoHotkey v2

#Include <AmenderRapport>
#Include <FinalTextInserts>
#Include <CheckExam>
#Include <AppendLogFiles>
#Include <GetData>

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

global g_dcRaw := ""          ; JSON brut en mémoire
global g_dcCache := Map()     ; Map(key -> value)
global g_dcHwnd := 0          ; hwnd associé au cache

OnMessage(WM_COPYDATA, CopyDataHandler)
textFile := ".\Textes\Insertions\textesRapport.txt"
finalRtf := ".\Textes\Insertions\textefinal.rtf"
finalRtfSUG := ".\Textes\Insertions\textefinalSUG.rtf"
flagFile := EnvGet("USERPROFILE") "\pause.flag"
global target
global downloadsDir := EnvGet("USERPROFILE") "\Downloads"
SetTitleMatchMode 2

if A_Args.Length = 0 {
    MsgBox "Aucun Argument fourni. Utiliser Ouvrir, Fermer, Tout, Signer ou Reculer."
    ExitApp
}

arg := A_Args[1]

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
    case "Signer":
        Signer()
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
        MsgBox "Invalid argument: " arg "`nUse Ouvrir, Fermer, Tout, Signer or Reculer."
}

ExitApp

Ouvrir(mode := "") {
    global reqnb_norm, reqnb_full, target  

; === Pré requis:Fenêtre de protocolage fermée
bloquante := "Module Protocolage"  

; Vérifier si ouverte
if WinExist(bloquante) {
    WinClose bloquante                        
    if !WinWaitClose(bloquante, , 2) {        
        MsgBox "La fenêtre bloquante est encore ouverte. Le script va s'arrêter.", 262144
        ExitApp
    }
}

; === Step 1: Activer la fenêtre Radimage ===
if !WinExist("ahk_exe RadImage.exe") {
    MsgBox "Fenêtre Radimage introuvable ou inactive. Le script va se fermer.", 262144
    ExitApp
}

WinActivate("ahk_exe RadImage.exe")
WinWaitActive("ahk_exe RadImage.exe",,5)

maxMs    := 5000    ; délai total max en ms (augmente si Radimage est lent)
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
        Sleep 200  ; attendre avant de réessayer
    }
    if !found {
        MsgBox "Impossible de trouver le numéro de requête car RadImage est trop lent à réagir. SVP réessayer dans quelques secondes."
	        ExitApp
    }

    rad_norm := StrReplace(radText, "-") 

; === Vérification concordance ReqNb ===
target := WinExist("RadEdit ahk_exe RadEdit.exe")
if !target {
    MsgBox "RadEdit introuvable."
    ExitApp
}

PrimeDataContextCached(target)

formcomplete := Trim(GetDataContextVarCached(target, "formcomplete", ""))
if (formcomplete = "false") {
    MsgBox "Le rapport n'a pas été généré par le formulaire."
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
            MsgBox "Impossible de récupérer la requête depuis RadEdit.`nImpossible de vérifier la concordance."
            ExitApp
        }
    }
    reqnb_norm := SubStr(reqnb_full, -8)  

    if (reqnb_norm != rad_norm) {
        result := MsgBox(
            "La requête de l'examen actif ne correspond pas à la requête de l'examen dicté.`n"
            . "Synapse → " reqnb_norm "`n"
            . "RadImage → " rad_norm "`n`n"
            . "Voulez-vous continuer quand même ?",
            "Vérification", 0x40024) ; Yes/No + AlwaysOnTop

        if (result = "No") {
            ExitApp
        } else {
            ; === Vérification supplémentaire : rapport déjà présent ? ===
            text := CheckFenDict(fenDict)
            if (Trim(text) != "") {
                SoundPlay "*16"
                MsgBox "L'examen sélectionné contient déjà un rapport.`n`n"
                    . "Le script s'arrêtera.`n`n"
                    . "SVP sélectionnez le bon examen.",
                    "ATTENTION", 262144
                ExitApp
            }
        }
    }
}
catch as err {
    MsgBox "Erreur lors de la vérification ReqNb:`n" err.Message
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
    MsgBox "Impossible de sélectionner l'onglet 'Transcription' après " maxMs " ms.`nDernière erreur: " lastErr
    ExitApp
}
Sleep 20

if !ClickLotCourant() {
    MsgBox "Bouton 'Lot Courant' introuvable. Le script s'arrête."
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
    MsgBox "RadEdit n’a pas retourné de fichier temporaire."
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
    ; --- contenu identique à ce que tu avais dans le try de Ouvrir() ---

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
WinWaitActive    ; wait until it’s actually active
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

;Sleep 20

Title := "v5.7"   ; stable part
if hwnd := WinExist(Title " ahk_exe msedge.exe")
{
    WinActivate("ahk_id " hwnd)
}
if !WinWaitActive("ahk_id " hwnd, , 5)
{
    ;MsgBox "Synapse Viewer window not found or not active. Script will exit."
    ExitApp
}
    Sleep 20
    Send "{F8}"
    ;fin Signer
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
    MsgBox "Impossible de vérifier si le rapport de l'examen est vide.`n`nVous êtes probablement dans le mauvais examen.`n`nLe script va s'arrêter.", "ERREUR", 262144
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
        ; MsgBox "RadEdit error: " text  ; (optionnel) je le laisse commenté pour éviter les popups
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

    MsgBox "Vous pouvez maintenant aller continuer votre dictée.", "TERMINÉ", 262144

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
    ; 3) Supprimer les HTML temporaires associés (RadEdit_<safeReqnb>_*.html)
    if (reqnb_full != "") {
        safeReqnb := RegExReplace(reqnb_full, "[^\w-]", "_")
        pattern := A_ScriptDir "\Textes\HTML\RadEdit_" safeReqnb "*.html"

        try {
            Loop Files, pattern {
                try FileDelete(A_LoopFileFullPath)
            }
        } catch as err {
            ; optionnel: MsgBox si tu veux être averti d'un problème global de loop
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

    ; 2) Supprimer le fichier temporaire associé si présent
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


; Convertit toutes les séquences RTF \'xx (hex) en caractères réels
DecodeRTFHex(s) {
    out := ""
    pos := 1
    while RegExMatch(s, "\\'([0-9A-Fa-f]{2})", &m, pos) {
        out .= SubStr(s, pos, m.Pos(0) - pos)
        out .= Chr("0x" m[1])          ; ex: 'c9 -> É
        pos := m.Pos(0) + m.Len(0)
    }
    return out . SubStr(s, pos)
}

; Décode \u#### (y compris valeurs négatives) -> caractère Unicode
DecodeRTFUnicode(s) {
    out := ""
    pos := 1
    while RegExMatch(s, "\\u(-?\d+)\??", &m, pos) {
        out .= SubStr(s, pos, m.Pos(0) - pos)
        code := m[1] + 0               ; coercition numérique
        if (code < 0)
            code += 65536
        out .= Chr(code)
        pos := m.Pos(0) + m.Len(0)
    }
    return out . SubStr(s, pos)
}

RensMaj(doc) {
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
        absEnd   := absStart + 1

        ; --- majuscule sans sélection visible ---
        r := doc.Range(absStart, absEnd)
        r.Case := 1

        ; fallback si Case ne fonctionne pas
        one := r.Text
        up  := StrUpper(one)
        if (up != one)
            r.Text := up

    } catch as err {
        ; MsgBox "RensMaj err:`n" err.Message
    }
}

CapitalizeParagraphStarts(doc, maxScan := 40) {
    wdUpperCase := 1
    paras := doc.Paragraphs
    count := paras.Count

    Loop count {
        para := paras.Item(A_Index)
        r := para.Range.Duplicate

        ; Paragraphe vide / juste un retour
        if (r.End - r.Start <= 1)
            continue

        ; Retirer le marqueur de paragraphe final (`r)
        r.End -= 1

        start := r.Start
        end   := r.End
        if (end <= start)
            continue

        sampleEnd := start + maxScan
        if (sampleEnd > end)
            sampleEnd := end

        ; Lire juste le début du paragraphe (1 call COM)
        sample := doc.Range(start, sampleEnd).Text
        if (sample = "")
            continue

        ; Trouver la première LETTRE (pas chiffre) dans le sample
	if !RegExMatch(sample, "\p{L}", &m)
    	    continue

        ; Position absolue de cette lettre dans le document Word
        pos := start + (m.Pos[0] - 1)

        chR := doc.Range(pos, pos + 1)
        ch  := chR.Text

        ; Si minuscule -> majuscule (1 seule lettre)
        if RegExMatch(ch, "^\p{Ll}$")
            chR.Case := wdUpperCase
    }
}



GotoEndofText(){
	target := WinExist("RadEdit ahk_exe RadEdit.exe")
	if !target {
	    MsgBox("❌ RadEdit non trouvé.")
 	   ExitApp()
	}

	SendCopyData(target, CMD["GotoEnd"], "")
	ExitApp()
}

HasArg(name) {
    for v in A_Args
        if (StrLower(v) = StrLower(name))
            return true
    return false
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


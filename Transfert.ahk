#Requires AutoHotkey v2

WM_COPYDATA := 0x004A
CMD := Map(
  "SetText",1, "InsertText",2, "SetFile",3, "InsertFile",4, "RequestTemp",5,
  "Response",6, "Error",7, "SetTitle",8, "GetTitle",9, "TitleResponse",10,
  "SetName",11, "GetName",12, "NameResponse",13, "GotoEnd",14,
  "FixFont",15, "CleanUpEnd",16,
)
global report_file := ""
global reqnb_name := ""
OnMessage(WM_COPYDATA, CopyDataHandler)
textFile := ".\Textes\Insertions\textesRapport.txt"
finalRtf := ".\Textes\Insertions\textefinal.rtf"
finalRtfSUG := ".\Textes\Insertions\textefinalSUG.rtf"
flagFile := EnvGet("USERPROFILE") "\pause.flag"

Esc:: {  
     try WinClose("ahk_exe WINWORD.EXE")
     MsgBox "Le script est arrêté", "Info", "T1 262144"
     ExitApp
}

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
    default:
        MsgBox "Invalid argument: " arg "`nUse Ouvrir, Fermer, Tout, Signer or Reculer."
}

ExitApp

Ouvrir(mode := "") {

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
target := WinExist("ahk_exe RadEdit.exe")
if !target {
    MsgBox "Fenêtre RadEdit introuvable."
    ExitApp
}

try {
    reqnb_full := GetTitleFromRadEdit(target)
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
            "Vérification", 0x24) ; Yes/No

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
        ControlChooseIndex(3, tabCtrl, "ahk_exe RadImage.exe")   ; 3 = Transcription
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

source := WinExist("ahk_exe RadEdit.exe")
if !source {
    MsgBox "RadEdit window not found."
    ExitApp
}

SendCopyData(source, CMD["CleanUpEnd"], "")
SendCopyData(source, CMD["FixFont"], "Arial;10")

SendCopyData(source, CMD["RequestTemp"], "")

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

try {
    doc.Content.Delete()
    sel := word.Selection
    sel.InsertFile(report_file)

    if (result.missing.Length > 0) {
        InsertMissingTitles(doc, report_file, result.titres, result.missing)
    }

    RensMaj(doc)

attv := HasArg("AttV")
if (attv) {
    InsertManualAttVer(doc)
    LogAttVer(result.titres)
} else if (result.needAttVer) {
    AttVer(doc)
    LogAttVer(result.titres)
}

AddText(mode)
ResetRadEdit(target, report_file)


}finally {
    undo.EndCustomRecord()
    word.ScreenUpdating := oldSU
}
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
    MsgBox "Synapse Viewer window not found or not active. Script will exit."
    ExitApp
}
Sleep 20
Send "{F8}"
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

GetTextByKey(filePath, key, default:="") {
    if !FileExist(filePath)
        return default
    For line in StrSplit(FileRead(filePath, "UTF-8"), "`n", "`r") {
        parts := StrSplit(line, "=")
        if (parts.Length >= 2 && parts[1] = key)
            return Trim(parts[2])
    }
    return default
}

AddText(mode := "") {
    try word := ComObjActive("Word.Application")
    catch {
        MsgBox "Word n'est pas ouvert.", "Erreur", "T3"
        return
    }
    doc := word.ActiveDocument
    sel := word.Selection

    chosenRtf := (mode = "ToutSUG" or mode = "OuvrirSUG") ? finalRtfSUG : finalRtf
    if AddFinalRtf(doc, sel, chosenRtf) {
        return
    }


    ; --- Textes fixes ---
    textSUG    := GetTextByKey(textFile, "textSUG",	"Examen demandé, réalisé et interprété en urgence sur l'horaire de garde à la demande du médecin prescripteur.")
    authorLine := GetTextByKey(textFile, "authorLine",	"Dictée faite par Dr")
    sent1      := GetTextByKey(textFile, "sent1",	"Ce rapport a été généré à l'aide d'un logiciel de reconnaissance vocale.")
    sent2      := GetTextByKey(textFile, "sent2",	"Pour cette raison il pourrait contenir des erreurs grammaticales, orthographiques ou syntaxiques.")

    ; Bloc en fonction du mode
    if (mode = "ToutSUG" or mode = "OuvrirSUG") {
        fullText := "`r`n`r`n" textSUG "`r`n`r`n" authorLine " " GetFullName_API() "`r`n`r`n" sent1 " " sent2
    } else {
        fullText := "`r`n`r`n" authorLine " " GetFullName_API() "`r`n`r`n" sent1 " " sent2
    }
    ; --- Insertion ---
    sel.EndKey(6)          ; aller à la fin du doc (wdStory = 6)
    startPos := sel.Start  ; mémoriser début de l’insertion
    sel.ParagraphFormat.Alignment := 0
    sel.TypeText(fullText) ; insérer tout le bloc

    ; --- Mise en forme ---
    endPos := sel.Start
    blockRng := doc.Range(startPos, endPos)

    ; Globale : Arial 10 italique
    blockRng.Font.Name   := "Arial"
    blockRng.Font.Size   := 10
    blockRng.Font.Italic := True

    ; Spécifique : première phrase en gras + souligné
    fr := blockRng.Duplicate
    fr.Find.ClearFormatting
    fr.Find.Text := sent1
    if fr.Find.Execute() {
        fr.Font.Bold := True
        ;fr.Font.Underline := 1
    }

    ; --- Fin ---
    sel.SetRange(doc.Content.End, doc.Content.End)
}

GetFullName_API() {
    NameDisplay := 3
    len := 256
    buf := Buffer(len * 2)  ; UTF-16

    ok := DllCall("Secur32\GetUserNameExW"
                , "Int",  NameDisplay
                , "Ptr",  buf.Ptr
                , "UInt*", len)

    if ok {
        name := StrGet(buf.Ptr, "UTF-16")
        parts := StrSplit(name, " ")
        if parts.Length = 2
            return parts[2] " " parts[1]  ; Inverse : Prénom Nom
        return name
    }

    return ""
}

AddFinalRtf(doc, sel, finalRtfPath := "") {
    global finalRtf, finalRtfSUG, textFile
    try {
        ; 1) Résoudre le chemin demandé
        if (finalRtfPath = "")
            finalRtfPath := finalRtf

        abs := finalRtfPath
        if !InStr(abs, ":")
            abs := A_ScriptDir "\" RegExReplace(finalRtfPath, "^\.(\\|/)", "")

        ; 2) Si le fichier demandé existe → insertion directe
        if FileExist(abs) {
            sel.SetRange(doc.Content.End, doc.Content.End)
            beforeEnd := doc.Content.End
            sel.InsertFile(abs)
            return (doc.Content.End > beforeEnd)
        }

        ; 3) Sinon, si on a demandé le SUG mais qu'il n'existe pas…
        ;    On tente: insérer textSUG + insérer le RTF standard si présent
        ;    (Si le standard n'existe pas non plus → return false, AddText fera son fallback.)
        ;    On compare avec le chemin SUG global.
        sugAbs := finalRtfSUG
        if !InStr(sugAbs, ":")
            sugAbs := A_ScriptDir "\" RegExReplace(finalRtfSUG, "^\.(\\|/)", "")

        isSUGRequest := (StrLower(abs) = StrLower(sugAbs))
        if isSUGRequest {
            ; Standard absolu
            stdAbs := finalRtf
            if !InStr(stdAbs, ":")
                stdAbs := A_ScriptDir "\" RegExReplace(finalRtf, "^\.(\\|/)", "")

            if FileExist(stdAbs) {
                ; 3a) Insérer la ligne textSUG (style: Arial 10 italique), puis le RTF standard
                textSUG := GetTextByKey(textFile, "textSUG"
                    , "Examen demandé, réalisé et interprété en urgence sur l'horaire de garde à la demande du médecin prescripteur.")

                sel.EndKey(6)
                startSUG := sel.Start
                sel.TypeText("`r`n`r`n" textSUG "")
                rngSUG := doc.Range(startSUG, sel.Start)
                rngSUG.Font.Name   := "Arial"
                rngSUG.Font.Size   := 10
                rngSUG.Font.Italic := True

                ; Insérer le RTF standard
                sel.SetRange(doc.Content.End, doc.Content.End)
                beforeEnd := doc.Content.End
                sel.InsertFile(stdAbs)
                return (doc.Content.End > beforeEnd)
            }
        }

        ; 4) Rien inséré ici → AddText fera le fallback (incluant textSUG si mode SUG)
        return false
    } catch as err {
        return false
    }
}



CleanupEnd(doc) {
    ; Nettoie les paragraphes vides en fin de document V1
    while (doc.Paragraphs.Count > 1) {
        lastPara := doc.Paragraphs.Item(doc.Paragraphs.Count)
        t := Trim(lastPara.Range.Text, "`r`n `t")
        if (t = "") {
            lastPara.Range.Delete()
        } else {
            break
        }
    }
}

AmenderRapport(typeAmendement) {
; === Step 1: Activer la fenêtre Radimage ===
if !WinExist("ahk_exe RadImage.exe") {
    MsgBox "Fenêtre Radimage introuvable ou inactive. Le script va se fermer."
    ExitApp
}
WinActivate("ahk_exe RadImage.exe")
WinWaitActive("ahk_exe RadImage.exe",,5)

Send "!d"

; === Step 2: Attendre et activer la fenêtre Amendement ===
if !WinWait("Amendement", , 5) {
    MsgBox "La fenêtre 'Amendement' n'a pas été détectée dans le délai imparti. Le script va s'arrêter."
    ExitApp
}
if !WinWaitActive("Amendement", , 5) {
    MsgBox "La fenêtre 'Amendement' n'est pas devenue active dans le délai imparti. Le script va s'arrêter."
    ExitApp
}

; === Step 3: Écrire Pathologie ou Amender et valider ===
if (typeAmendement = "Pathologie") {
    Send "PATHOLOGIE{Enter}"
} else if (typeAmendement = "Amender" || typeAmendement = "Modification") {
    Send "Correction{Enter}"   ; mot tapé reste Correction
} else {
    MsgBox "Argument invalide : " typeAmendement
    ExitApp
}

; === Attendre Word ===
if !WinWaitActive("ahk_exe WINWORD.EXE", , 10) {
    MsgBox "Word did not open or become active in time. Script will exit."
    ExitApp
}

; === Step 4: Insérer texte formaté dans Word ===
try {
    wdApp := ComObjActive("Word.Application")
    if !wdApp.ActiveDocument {
        throw Error("Aucun document actif dans Word.")
    }
    wdDoc := wdApp.ActiveDocument
    rng := wdDoc.Range()
    rng.Collapse(0)

    ; Police par défaut
    rng.Font.Name := "Arial"
    rng.Font.Size := 10

    ; Contenu spécifique
    if (typeAmendement = "Pathologie") {
	rng.InsertParagraphAfter()
	rng.InsertParagraphAfter()
        rng.Collapse(0)
        rng.Text := "RAPPORT COMPLÉMENTAIRE"
        rng.Font.Bold := True
	rng.Font.Italic := False
        rng.InsertParagraphAfter()

        rng.Collapse(0)
        rng.Text := "RAPPORT DE PATHOLOGIE"
        rng.Font.Bold := True
	rng.Font.Italic := False
        rng.InsertParagraphAfter()
        rng.InsertParagraphAfter()

        rng.Collapse(0)
        rng.Text := "[]"
        rng.Font.Bold := False
        rng.Font.Italic := False
        rng.InsertParagraphAfter()
        rng.InsertParagraphAfter()

        rng.Collapse(0)
        rng.Text := "Se référer au rapport de pathologie complet pour les détails."
        rng.InsertParagraphAfter()
        rng.InsertParagraphAfter()

    } else if (typeAmendement = "Amender") {
 	rng.InsertParagraphAfter()
	rng.InsertParagraphAfter()
        rng.Collapse(0)
        rng.Text := "RAPPORT COMPLÉMENTAIRE"
        rng.Font.Bold := True
	rng.Font.Italic := False
        rng.InsertParagraphAfter()
        rng.InsertParagraphAfter()
        
        rng.Collapse(0)
        rng.Text := "[]"
        rng.Font.Bold := False
        rng.Font.Italic := False
        rng.InsertParagraphAfter()
        rng.InsertParagraphAfter()

    } else if (typeAmendement = "Modification") {
        rng.InsertParagraphAfter()
	rng.InsertParagraphAfter()
    }

    ; Phrase italique + gras avec date et nom
    today := FormatTime(, "d MMMM yyyy")
    rng.Collapse(0)
    rng.Text := "Rapport amendé le " today " par Dr " GetFullName_API()    
    rng.Font.Bold := True
    rng.Font.Italic := True
    rng.InsertParagraphAfter()

}
catch as err {
    MsgBox "Erreur COM Word: " err.Message
    ExitApp
}
Send "^!{Right}"
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
    static CMD_RESPONSE := 6, CMD_ERROR := 7, CMD_TITLERESPONSE := 10, CMD_NAMERESPONSE := 13
    global report_file, reqnb_title, reqnb_name

    cmd := NumGet(lParam, 0, "UPtr")
    size := NumGet(lParam, A_PtrSize, "UInt")
    text := StrGet(NumGet(lParam, 2*A_PtrSize, "Ptr"), size/2, "UTF-16")

    if (cmd = CMD_RESPONSE) {
        report_file := text
    } else if (cmd = CMD_TITLERESPONSE) {
        reqnb_title := Trim(text)
    } else if (cmd = CMD_NAMERESPONSE) {
        reqnb_name := Trim(text)
    } else if (cmd = CMD_ERROR) {
        MsgBox "RadEdit error: " text
    }
    return true
}


GetTitleFromRadEdit(hwnd) {
    global reqnb_title
    reqnb_title := ""   ; reset

    SendCopyData(hwnd, 9, "")  ; 9 = GetTitle

    t0 := A_TickCount
    while (reqnb_title = "" && A_TickCount - t0 < 1000)
        Sleep 20

    return reqnb_title
}

GetNameFromRadEdit(hwnd) {
    global reqnb_name
    reqnb_name := ""                  ; reset
    SendCopyData(hwnd, 12, "")        ; 12 = GetName
    t0 := A_TickCount
    while (reqnb_name = "" && A_TickCount - t0 < 1000)
        Sleep 20
    return reqnb_name
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

GetWordsByKey(filePath, key, defaults := []) {
    ; --- Si le fichier n’existe pas : retourner la liste par défaut ---
    if !FileExist(filePath)
        return defaults

    ; --- Lire et chercher la clé ---
    For line in StrSplit(FileRead(filePath, "UTF-8"), "`n", "`r") {
        parts := StrSplit(line, "=")
        if (parts.Length >= 2 && parts[1] = key) {
            ; Séparer les mots par virgule, point-virgule ou espace
            cleaned := Trim(parts[2])
            if (cleaned = "")
                return defaults
            words := []
            For w in StrSplit(cleaned, ",")
                if (Trim(w) != "")
                    words.Push(StrLower(Trim(w)))
            return words
        }
    }
    ; --- Si la clé n’a pas été trouvée ---
    return defaults
}


CheckExamList(doc, report_file) {
    app := doc.Application
    defaultExcl := ["renseignements", "manipulation", "rapport", "sein", "reconstruction", "inj."]
    exclusions := GetWordsByKey(textFile, "exclusions", defaultExcl)
    titres := []

    ; --- Étape 1 : Extraire titres de WordRadimage ---
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
        exclu := false
        for excl in exclusions {
            if InStr(lower, excl) = 1 {
                exclu := true
                break
            }
        }
        if !exclu
            titres.Push(txt)
    }

    ; --- Étape 2 : Lire directement le contenu RTF (rapide et sans Word) ---
try {
    rawText := FileRead(report_file, "CP1252")
} catch as err {
    MsgBox "Impossible de lire le fichier temporaire :`n" report_file "`n`nErreur : " err.Message, "Erreur critique", 262144
    CancelTransfer()
    return { titres: titres, missing: [], needAttVer: false }
}

; --- Décodage des séquences RTF \'xx -> caractère ---
rawText := DecodeRTFHex(rawText)

; (optionnel, si un jour tu vois des \u2019, \u00E9, etc.)
; rawText := DecodeRTFUnicode(rawText)

; --- Nettoyage des balises RTF (tes lignes existantes) ---
rawText := StrReplace(rawText, "{\par", "`n")
rawText := RegExReplace(rawText, "\\[a-z0-9]+(?:-?\d+)?[ ]?", "")
rawText := RegExReplace(rawText, "\{\\[^}]+\}", "")
rawText := RegExReplace(rawText, "[{}]", "")
rawText := RegExReplace(rawText, "[\x00-\x08\x0B-\x1F]", "")

; Séparer en lignes (paragraphes)
lines := []
for l in StrSplit(rawText, ["`r`n", "`n", "`r"]) {
    l := Trim(l, "`r`n `t")
    if (l != "")
        lines.Push(l)
}

; --- Détection stricte "Attention à vérifier" (insensible à la casse) ---
needAttVer := false
if (lines.Length > 0) {
    last := lines[lines.Length]
    last := StrReplace(last, Chr(160), " ")   ; NBSP -> espace, au cas où
    last := Trim(last, "`r`n `t")             ; on tolère juste espaces/retours
    if RegExMatch(last, "i)^\s*Attention à vérifier\s*$")
        needAttVer := true
}


    ; --- Étape 3 : Vérification ---
    missing := []
    for titre in titres {
        found := false
        titreNorm := StrLower(Trim(titre))

        for line in lines {
            lineTrim := StrLower(Trim(line))
            if (lineTrim = "")
                continue
            ; Comparaison : le titre Word doit être au début de la ligne du fichier temp
            if (InStr(lineTrim, titreNorm) = 1) {
                found := true
                break
            }
        }

        if !found
            missing.Push(titre)
    }

   ; --- Étape 4 : Si manquants, proposer options (fenêtre + touche 1/2/3) ---
if (missing.Length > 0) {
    total := titres.Length
    nbMiss := missing.Length

    msg := "Ces examens sont absents de votre dictée :" . "`n`n"
    for titre in missing
        msg .= "- " titre . "`n"

    msg .= "`n" . "Choisissez l'option appropriée :" . "`n`n"
    msg .= "1 - Continuer (ignorer et poursuivre)" . "`n`n"
    msg .= "2 - Retour à la dictée (annuler le transfert)" . "`n`n"
    msg .= "3 - Insérer le(s) titre(s) manquant(s)"

    dlg := Gui("+AlwaysOnTop", "Rapport incomplet")
    dlg.SetFont("s12 bold", "Segoe UI")    
    dlg.Add("Text", "w500 h500", msg)
    dlg.Show()

    ih := InputHook("L1")
    ih.Start()
    ih.Wait()
    choice := ih.Input

    dlg.Destroy()

        if (choice = "1") {
        return { titres: titres, missing: []            , needAttVer: needAttVer }
    } else if (choice = "2") {
        CancelTransfer()
        return { titres: titres, missing: []            , needAttVer: needAttVer }
    } else if (choice = "3") {
        return { titres: titres, missing: missing       , needAttVer: needAttVer }
    } else {
        CancelTransfer()
        return { titres: titres, missing: []            , needAttVer: needAttVer }
    }


}

return { titres: titres, missing: [], needAttVer: needAttVer }

}


InsertMissingTitles(doc, report_file, titres, missing) {
    ; --- 1) Construire la liste des titres éligibles (titres - missing) ---
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

    ; --- Ajouter l’option spéciale %TITRE% ---
    titresEligibles.Push("%TITRE%")

    ; --- 2) Pour chaque titre manquant, demander où insérer ---
    for titre in missing {
        msg := "⚠️  Titre manquant : " titre . "`n`n"
        msg .= "Choisissez après quel titre l’insérer (appuyez sur le chiffre) :" . "`n`n"

        for i, t in titresEligibles {
    if (t = "%TITRE%")
        msg .= i ". Insérer à la place de: %TITRE%`n"
    else
        msg .= i ". " . t . "`n"
}

        dlg := Gui("+AlwaysOnTop", "Insertion de titre manquant")
        dlg.SetFont("s12")
        dlg.Add("Text", "w800 h600", msg)
        dlg.Show()

        ih := InputHook("L1")
        ih.Start(), ih.Wait()
        choice := ih.Input
        dlg.Destroy()

        if (!RegExMatch(choice, "^\d+$")) {
            CancelTransfer()
            return
        }

        idx := choice + 0
        if (idx < 1 || idx > titresEligibles.Length) {
            CancelTransfer()
            return
        }

        anchor := titresEligibles[idx]

        ; --- 3) Cas spécial : insertion à la place de %TITRE% ---
        if (anchor = "%TITRE%") {
    ; Rechercher le littéral %TITRE% dans tout le document,
    ; puis remplacer directement dans le Range trouvé (sans passer par Selection)
    rng := doc.Content
    f := rng.Find
    f.ClearFormatting()
    f.MatchWildcards := false
    f.MatchCase := false
    f.MatchWholeWord := false
    f.Text := "%TITRE%"

    if (f.Execute()) {
        rng.Text := StrUpper(titre)  ; remplace le marqueur par le vrai titre
        rng.Font.Bold := True
        rng.Font.Underline := 1
    } else {
        MsgBox "Aucun marqueur %TITRE% trouvé dans le document.`nLe titre n’a pas pu être inséré.", "Erreur", 262144
        CancelTransfer()
    }
    continue  ; passer au titre manquant suivant
}

        ; --- 4) Sinon, insertion après un vrai titre existant ---
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

    if WinExist("ahk_exe RadEdit.exe") {
	WinActivate
	}
    if !WinWaitActive("ahk_exe RadEdit.exe", , 5) {
	MsgBox "RadEdit window not found or not active. Script will exit."
	ExitApp
    }
    Send "{PgDn}"
    ExitApp
}

ResetRadEdit(target, tempFile) {
    global CMD
    ; 1) Vider RadEdit
    SendCopyData(target, CMD["SetText"], "{\rtf1\ansi}")
    SendCopyData(target, CMD["SetTitle"], "")
    SendCopyData(target, CMD["SetName"], "")

    ; 2) Supprimer le fichier temporaire
    try {
        if FileExist(tempFile) {
            FileDelete(tempFile)
        }
    } catch as err {
        MsgBox "Impossible d'effacer le fichier temporaire :`n" tempFile "`n`nErreur : " err.Message, "ATTENTION", 262144
    }
}

EraseRadEdit() {
    global CMD, reqnb_title
    target := WinExist("ahk_exe RadEdit.exe")
    if !target {
        MsgBox "RadEdit introuvable.", "Erreur", 262144
        return
    }
    reqnb_title := GetTitleFromRadEdit(target)

    ; 1) Effacer contenu RadEdit
    SendCopyData(target, CMD["SetText"], "{\rtf1\ansi}")
    SendCopyData(target, CMD["SetTitle"], "")
    SendCopyData(target, CMD["SetName"], "")

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

AttVer(doc) {
    AlignLeft := 0, AlignCenter := 1

    paras := doc.Paragraphs
    cnt := paras.Count
    if (cnt < 1)
        return

    ; Dernier paragraphe non vide
    idx := 0
    Loop cnt {
      i := cnt - A_Index + 1
      t := Trim(paras.Item(i).Range.Text, "`r`n `t")
      if (t != "") {
          idx := i
          break
      }
    }
    if (idx = 0)
      return

    p := paras.Item(idx)

; --- normalisation ultra-légère et tolérante ---
txt := p.Range.Text
txt := StrReplace(txt, Chr(160), " ")       ; remplace NBSP par espace
txt := Trim(txt, "`r`n `t")
txt := RTrim(txt, " :")                     ; tolère un ":" final + espace

; --- match tolérant accents / casse ---
if !RegExMatch(txt, "i)^\s*attention\s+(?:à|a)\s+v(?:é|e)rifier\s*$")
    return


    ; --- Ajouter un Enter AVANT ---
    p.Range.InsertParagraphBefore()
    paras := doc.Paragraphs          ; refresh
    idx := idx + 1
    p := paras.Item(idx)
    if (idx > 1)
        paras.Item(idx - 1).Alignment := AlignLeft

    ; --- Mise en forme de CE paragraphe ---
    r := p.Range
    txtRange := doc.Range(r.Start, Max(r.Start, r.End - 1))
    try txtRange.Case := 1            ; wdUpperCase
    catch as err {
        up := StrUpper(txtRange.Text)
        txtRange.Text := up
    }
    txtRange.Font.Name      := "Arial"
    txtRange.Font.Size      := 12
    txtRange.Font.Bold      := True
    txtRange.Font.Underline := 1
    p.Alignment := AlignCenter

    ; Paragraphe suivant à gauche (pour ne pas polluer AddText)
    if (idx < paras.Count)
        paras.Item(idx + 1).Alignment := AlignLeft
    else {
        r2 := doc.Range(doc.Content.End, doc.Content.End)
        r2.InsertParagraphAfter()
        paras := doc.Paragraphs
        paras.Item(paras.Count).Alignment := AlignLeft
    }
}


GotoEndofText(){
	target := WinExist("ahk_exe RadEdit.exe")
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

LogAttVer(titres := [], rootPath := "\\bureautique.chrdl.qc.ca\CHRDL\Imagerie Medicale\Partage\Suivi Urgence") {
    
    ; 1) S’assurer du dossier
    try {
        if !DirExist(rootPath)
            DirCreate(rootPath)
    } catch as err {
        return false
    }

    ; 2) Récupérer le NAME du bandeau RadEdit
    name := ""
    try {
        hwnd := WinExist("ahk_exe RadEdit.exe")
        if (hwnd)
            name := GetNameFromRadEdit(hwnd)
    }

    if (name = "")
        name := "(NAME inconnu)"

    ; 3) Construire le bloc texte
    ts := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
    buf := name "  —  Date de lecture: " ts "`r`n"
    if (IsObject(titres) && titres.Length > 0) {
        for t in titres
            buf .= "- " t "`r`n"
    }
    buf .= "----------------------------------------`r`n"

    ; 4) Append dans le fichier log
    logFile := rootPath "\ATTENTION À VÉRIFIER.txt"   ; nom exact demandé
    try {
        FileAppend(buf, logFile, "UTF-8")
        return true
    } catch as err {
        return false
    }
}

InsertManualAttVer(doc) {
    ; Ancre sûre à la fin du document, même s’il est vide
    endPos := doc.Content.End
    startPos := (endPos > 0) ? (endPos - 1) : 0
    r := doc.Range(startPos, startPos)

    ; 2 lignes vides AVANT
    r.InsertParagraphAfter()
    r.Collapse(0)
    r.InsertParagraphAfter()
    r.Collapse(0)

    ; Insérer le texte et le formater
    r.Text := "ATTENTION À VÉRIFIER"
    r.Font.Name      := "Arial"
    r.Font.Size      := 12
    r.Font.Bold      := True
    r.Font.Underline := 1
    r.ParagraphFormat.Alignment := 1
    
    r.InsertParagraphAfter()

    ; Revenir en Arial 10 pour la suite (à la nouvelle fin)
    endPos2 := doc.Content.End
    tailPos := (endPos2 > 0) ? (endPos2 - 1) : 0
    tail := doc.Range(tailPos, tailPos)
    tail.Font.Name := "Arial"
    tail.Font.Size := 10
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

Pause() {
    target := WinExist("ahk_exe RadEdit.exe")
    if !ProcessExist("FusionDictate.exe")
    Run '"C:\APP\Fusion_StartUp\startup_imagerie.cmd"', , "Hide"
    SendCopyData(target, CMD["SetName"], "PAUSE SYNCHRO RADEDIT - SYNCHRO FUSION ACTIF")
    SendCopyData(target, CMD["SetText"], "{\rtf1\ansi}")	
    SendCopyData(target, CMD["SetTitle"], "")
}

Resume() {
    target := WinExist("ahk_exe RadEdit.exe")
    SendCopyData(target, CMD["SetText"], "{\rtf1\ansi}")	
    SendCopyData(target, CMD["SetName"], "")
    SendCopyData(target, CMD["SetTitle"], "")
    if ProcessExist("FusionDictate.exe")
       ProcessClose("FusionDictate.exe")
}
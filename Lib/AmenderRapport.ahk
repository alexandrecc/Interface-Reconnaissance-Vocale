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

    rng := wdDoc.Content
    rng.Font.Name := "Arial"
    rng.Font.ColorIndex := 1 
    return true
}
catch as err {
    MsgBox "Erreur COM Word: " err.Message
    ExitApp
}
Send "^!{Right}"
}
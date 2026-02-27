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

        if !exclu {
            titres.Push(txt)
        }
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
            return { titres: titres, missing: [],      needAttVer: needAttVer }
        } else if (choice = "2") {
            CancelTransfer()
            return { titres: titres, missing: [],      needAttVer: needAttVer }
        } else if (choice = "3") {
            return { titres: titres, missing: missing, needAttVer: needAttVer }
        } else {
            CancelTransfer()
            return { titres: titres, missing: [],      needAttVer: needAttVer }
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


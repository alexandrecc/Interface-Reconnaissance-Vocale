BasculerMammo() {
    global CMD, mammoflag, mammodictflag

    target := WinExist("RadEdit ahk_exe RadEdit.exe")

    if (FileExist(mammoflag) || FileExist(mammodictflag)) {
        if FileExist(mammoflag)
            FileDelete(mammoflag)
        if FileExist(mammodictflag)
            FileDelete(mammodictflag)
    SendCopyData(target, CMD["SetName"], "Mode Lecture Mammo Dépistage INACTIF")
    }
    else {
        FileAppend("", mammoflag) 
    SendCopyData(target, CMD["SetName"], "Mode Lecture Mammo Dépistage ACTIF")
    }
    ExitApp
}

SendMAMMOReporttoRadEdit(reqnb_full) {
    global CMD, downloadsDir, target, mammoflag

    target := WinExist("RadEdit ahk_exe RadEdit.exe")
    if !target {
        MsgBox "Impossible de trouver la fenêtre cible RadEdit.`nAjuste WinExist(...) dans le script de test."
        ExitApp
    }

    mammoTxtPath := downloadsDir "\Mammo_" reqnb_full ".txt"

    vars := ReadMammoVarsFromTxt(mammoTxtPath)
    if (vars.Count = 0) {
        MsgBox "Aucune variable mammo trouvée dans:`n" mammoTxtPath
        ExitApp
    }

    SendCopyData(target, CMD["GotoEnd"], "")
    block1 := BuildBlock_TypeExam(vars)  ; <TypeExam>
    block2 := BuildBlock_Proc(vars)      ; <proc>
    block3 := BuildBlock_Reste(vars)     ; reste du rapport

    if (block1 != "")
        SendCopyData(target, CMD["InsertText"], block1)

    if (block2 != "")
        SendCopyData(target, CMD["InsertText"], block2)

    if (block3 != "")
        SendCopyData(target, CMD["InsertText"], block3)

    if FileExist(mammoTxtPath) {
        FileDelete(mammoTxtPath)
    }

    RequeteLOC(vars)
    FinalizeMAMMODictation(vars)
}

TransfertMAMMO(source, target, mode := "") {
    global CMD, report_file, mammoflag, mammodictflag

    SendCopyData(source, CMD["CleanUpEnd"], "")
    SendCopyData(source, CMD["FixFont"], "Arial;10")

    SendCopyData(source, CMD["RequestTemp"], "")

    ; Attendre que report_file soit rempli
    maxWait := 1000
    waited := 0
    while (report_file == "" && waited < maxWait) {
        Sleep 50
        waited += 50
    }
    if (report_file == "") {
        MsgBox "RadEdit n’a pas retourné de fichier temporaire."
        ExitApp
    }

    try word := ComObjActive("Word.Application")
    catch as err {
        MsgBox "Word n'est pas ouvert."
        ExitApp
    }

    doc := word.ActiveDocument
    doc.Content.Delete()

    sel := word.Selection
    sel.SetRange(doc.Content.Start, doc.Content.Start)
    sel.InsertFile(report_file)

    AddText(mode)

    ResetRadEdit(target, report_file, reqnb_full)

    if FileExist(mammodictflag) {
        FileDelete(mammodictflag)
        FileAppend("", mammoflag)
    }
}

ReadMammoVarsFromTxt(filePath) {
    vars := Map()

    if !FileExist(filePath) {
        MsgBox "Fichier mammo introuvable:`n" filePath
        CancelTransfer()
    }

    try file := FileOpen(filePath, "r", "UTF-8")
    catch as err {
        MsgBox "Erreur à l'ouverture du fichier mammo:`n" filePath "`n`n" err.Message
        return vars
    }

    while !file.AtEOF {
        line := Trim(file.ReadLine())
        if (line = "" || RegExMatch(line, "^(;|#)"))
            continue

        ; NomVar = [valeur]
        if !RegExMatch(line, "^\s*([^=]+?)\s*=\s*\[(.*)]\s*$", &m)
            continue

        key   := Trim(m[1])
        value := m[2]
        vars[key] := value
    }

    file.Close()
    return vars
}

BuildBlock_TypeExam(vars) {
    typeExam := GetVar(vars, "TypeExam")
    if (typeExam = "")
        return ""

    label := "Renseignements cliniques"

    rtf := "{\rtf1\ansi "
    ; \b\ul label \ul0:\b0  texte
    rtf .= "\b\ul " RtfEscape(label) "\ul0:\b0  "
    if (typeExam != "")
        rtf .= RtfEscape(typeExam)

    rtf .= "\par\par\par}"
    return rtf
}


BuildBlock_Proc(vars) {
    ; adapte "proc" au nom exact de ta variable
    proc := GetVar(vars, "proc")
    if (proc = "")
        return ""

    txt := RtfEscape(proc)

    ; b = bold, ul = underline, puis \par\par pour une ligne blanche après
    return "{\rtf1\ansi\b\ul " txt "\b0\ul0\par\par}"
}

BuildBlock_Reste(vars) {

    if (GetVar(vars, "Dictee") = "Oui") {
        txt := "[]"
        return "{\rtf1\ansi\b0\i0\ul0 " RtfEscape(txt) "}"
    }

    lines := []

    ; --- Infos de base ---
    AddIfNotEmpty(lines, vars, "Mammo#1")
    AddIfNotEmpty(lines, vars, "Tomo")
    AddIfNotEmpty(lines, vars, "AspectPar")
    AddIfNotEmpty(lines, vars, "Protheses")
    AddIfNotEmpty(lines, vars, "ATCDfam")
    AddIfNotEmpty(lines, vars, "SuiviHR")
    AddIfNotEmpty(lines, vars, "ATCDneo")

    postOp1 := GetVar(vars, "PostOp1")
    if (postOp1 = "Oui")
        lines.Push("Première mammographie post chirurgie")

    result := GetVar(vars, "Result")

    ; ========================
    ; CAS 1 : Résultat NON anormal
    ; ========================
    if (result != "" && result != "Anormal") {
        ; Ligne vide entre les infos de base et "Résultat"
        if (lines.Length > 0)
            lines.Push("")

        ; --- Résultat (sans ligne vide avant, juste après les infos) ---
        if (result = "Normal" && postOp1 = "Oui")
            lines.Push("Résultat: Normal pour une première mammographie post-chirurgie.")
        else
            lines.Push("Résultat: " result)

        ; --- Tout ce qui vient APRÈS Résultat: avec 1 ligne vide entre chaque ---

        ; Recommandation
        suiviResult := GetVar(vars, "SuiviResult")
        if (suiviResult != "")
            AddBlockAfterResult(lines, "Recommandation: " suiviResult)

        ; CS, C2... C6, avec bloc "Remarques supplémentaires:"
        hasExtra := false
        for , key in ["CS","C2","C3","C4","C5","C6"] {
            val := GetVar(vars, key)
            if (val != "") {
                ; première fois qu'on rencontre un texte supplémentaire
                if (!hasExtra) {
                    AddBlockAfterResult(lines, "Remarques supplémentaires:")
                    hasExtra := true
                }
                AddBlockAfterResult(lines, val)
            }
        }

        TrimEmptyEdges(lines)
        return LinesToRtf(lines)
    }


    ; ========================
    ; CAS 2 : Résultat Anormal
    ; ========================
    if (result = "Anormal") {
        ; Ligne vide entre infos de base et “Résultat”
        if (lines.Length > 0)
            lines.Push("")

        ; Construire les patterns (LesionUni/Multi/Diffus D/G)
        patterns := []
        for , key in ["LesionUniD","LesionUniG"
                    ,"LesionMultiD","LesionMultiG"
                    ,"LesionDiffusD","LesionDiffusG"] {
            val := GetVar(vars, key)
            if (val != "")
                patterns.Push(val)
        }

        ; --- Résultat anormal (bloc de base) ---
        if (patterns.Length > 0)
            lines.Push("Résultat: Anormal: " JoinLines(patterns, " et "))
        else
            lines.Push("Résultat: Anormal.")

        ; --- Lésions : chacune séparée par une ligne vide ---
        maxLesions := 5
        Loop maxLesions {
            idx    := A_Index
            lat    := GetVar(vars, "LatLesion" idx)
            lesion := GetVar(vars, "Lesion" idx)
            loc    := GetVar(vars, "LocLesion" idx)
            taille := GetVar(vars, "TailleLesion" idx)

            if (lat = "" && lesion = "" && loc = "" && taille = "")
                continue

            desc := ""
            if (lat != "")
                desc .= lat ": "
            if (lesion != "")
                desc .= lesion
            if (loc != "") {
                if (desc != "")
                    desc .= " en " loc
                else
                    desc .= loc
            }
            if (taille != "") {
                if (desc != "")
                    desc .= " de " taille
                else
                    desc .= taille
            }

            if (desc = "")
                continue

            AddBlockAfterResult(lines, "    " idx ". " desc)
        }

        ; --- Examens complémentaires : encore un bloc espacé ---
        examSuppl := GetVar(vars, "ExamSuppl")
        if (examSuppl != "")
            AddBlockAfterResult(lines, "Patiente sera rappelée pour des examens complémentaires: " examSuppl)

        TrimEmptyEdges(lines)
        return LinesToRtf(lines)

    }

    ; ========================
    ; CAS 3 : Result vide -> on renvoie juste les infos
    ; ========================
    TrimEmptyEdges(lines)
    return LinesToRtf(lines)

}

RequeteLOC(vars) {
    ; Seulement si résultat anormal
    if (GetVar(vars, "Result") != "Anormal")
        return

    ; Nom du requérant (API) — appelé ici de façon indépendante
    fullName := ""
    try fullName := GetFullName_API()
    catch as err
        fullName := ""

    if (fullName = "")
        fullName := "Inconnu"

    ; Date: une pour le nom de fichier, une pour l’affichage
    fileDate    := FormatTime(, "yyyy-MM-dd")
    displayDate := FormatTime(, "dd/MM/yyyy")

    outPath := A_ScriptDir "\LOC - " fullName " - " fileDate ".txt"

    ; En-tête + infos demandées
    numDos   := GetVar(vars, "NumDos")
    patName  := GetVar(vars, "PatName")
    proc     := GetVar(vars, "proc")
    dateExam := GetVar(vars, "DateExam")
    examens := GetVar(vars, "ExamSuppl")

    entry := ""

    ; Séparateur si le fichier existe déjà et n'est pas vide
    if FileExist(outPath) {
        try {
            if (FileGetSize(outPath) > 0)
                entry .= "`r`n########################################`r`n"
        } catch as err {
            entry .= "`r`n########################################`r`n"
        }
    }

    entry .= displayDate "`r`n"
    if (numDos != "")
        entry .= numDos "`r`n"
    if (patName != "")
        entry .= patName "`r`n"

    entry .= "`r`n"
    if (proc != "")
        entry .= "Examen complémentaire de: " proc "`r`n"
    if (dateExam != "")
        entry .= "du: " dateExam "`r`n"

    entry .= "`r`n"

    ; Texte CAS 2 (Anormal) SANS la ligne ExamSuppl
    anormalTxt := BuildLOC_AnormalText(vars)
    if (anormalTxt != "")
        entry .= anormalTxt "`r`n`r`n"

    entry .= "EXAMENS COMPLÉMENTAIRES: " examens "`r`n`r`n"

    entry .= "Requête faite par: " fullName "`r`n"

    try FileAppend(entry, outPath, "UTF-8")
    catch as err {
        MsgBox "Erreur lors de l'écriture du fichier LOC:`n" outPath "`n`n" err.Message, "LOC", 262144
    }
}

BuildLOC_AnormalText(vars) {
    lines := []

    ; Patterns (LesionUni/Multi/Diffus D/G)
    patterns := []
    for , key in ["LesionUniD","LesionUniG"
                ,"LesionMultiD","LesionMultiG"
                ,"LesionDiffusD","LesionDiffusG"] {
        val := GetVar(vars, key)
        if (val != "")
            patterns.Push(val)
    }

    if (patterns.Length > 0)
        lines.Push("RÉSULTAT: Anormal: " JoinLines(patterns, " et "))
    else
        lines.Push("RÉSULTAT: Anormal.")

    ; Lésions (avec une ligne vide entre chacune)
    maxLesions := 5
    Loop maxLesions {
        idx    := A_Index
        lat    := GetVar(vars, "LatLesion" idx)
        lesion := GetVar(vars, "Lesion" idx)

        if (lat = "" && lesion = "")
            continue

        desc := ""
        if (lat != "")
            desc .= lat ": "
        if (lesion != "")
            desc .= lesion

        if (desc = "")
            continue

        lines.Push("")
        lines.Push(idx ". " desc)
    }

    TrimEmptyEdges(lines)

    ; Join CRLF
    out := ""
    for i, line in lines {
        if (i > 1)
            out .= "`r`n"
        out .= line
    }
    return out
}

GetVar(vars, key, default := "") {
    return vars.Has(key) ? vars[key] : default
}

AddIfNotEmpty(linesArr, vars, key, prefix := "") {
    val := GetVar(vars, key)
    if (val = "")
        return
    if (prefix != "")
        linesArr.Push(prefix val)
    else
        linesArr.Push(val)
}

TrimEmptyEdges(linesArr) {
    while (linesArr.Length > 0 && Trim(linesArr[1]) = "")
        linesArr.RemoveAt(1)
    while (linesArr.Length > 0 && Trim(linesArr[linesArr.Length]) = "")
        linesArr.RemoveAt(linesArr.Length)
}

JoinLines(arr, sep := "`r`n") {
    out := ""
    for i, line in arr {
        if (i > 1)
            out .= sep
        out .= line
    }
    return out
}

RtfEscape(text) {
    ; protéger les caractères RTF sensibles
    text := StrReplace(text, "\", "\\")
    text := StrReplace(text, "{", "\{")
    text := StrReplace(text, "}", "\}")

    ; convertir les retours à la ligne en \par
    text := StrReplace(text, "`r`n", "\par ")
    text := StrReplace(text, "`n",   "\par ")
    text := StrReplace(text, "`r",   "\par ")

    return text
}

AddBlockAfterResult(linesArr, text) {
    if (text = "")
        return
    ; ligne vide pour espacer
    linesArr.Push("")
    linesArr.Push(text)
}

LinesToRtf(lines) {
    if (lines.Length = 0)
        return ""

    ; reset de tout style au début du bloc
    rtfBody := "\b0\i0\ul0 "

    for i, line in lines {
        if (i > 1)
            rtfBody .= "\par "

        if (line = "")
            continue

        ; --- Résultat: ... ---
        if (SubStr(line, 1, 9) = "Résultat:") {
            label := "Résultat"
            rest  := SubStr(line, StrLen(label . ": ") + 1)
            rtfBody .= "\b\ul " RtfEscape(label) "\ul0:\b0  "
            if (rest != "")
                rtfBody .= RtfEscape(rest)
            continue
        }

        ; --- Recommandation: ... ---
        if (SubStr(line, 1, 15) = "Recommandation:") {
            label := "Recommandation"
            rest  := SubStr(line, StrLen(label . ": ") + 1)
            rtfBody .= "\b\ul " RtfEscape(label) "\ul0:\b0  "
            if (rest != "")
                rtfBody .= RtfEscape(rest)
            continue
        }

        ; --- Remarques supplémentaires: (gras simple) ---
        if (line = "Remarques supplémentaires:") {
            rtfBody .= "\b " RtfEscape(line) "\b0"
            continue
        }

        ; --- lignes normales ---
        rtfBody .= RtfEscape(line)
    }

    if (rtfBody = "")
        return ""

    return "{\rtf1\ansi " rtfBody "}"
}

FinalizeMAMMODictation(vars) {
    global mammoflag, mammodictflag

    dictee := GetVar(vars, "Dictee")
    c5     := GetVar(vars, "C5")

    if !(dictee = "Oui" || c5 = "[]")
        return

    PrepareForDictation()

    if FileExist(mammoflag) {
        FileDelete(mammoflag)
    }

    FileAppend("", mammodictflag, "UTF-8")

    CancelTransfer()
}

PrepareForDictation() {
    target := 0
    maxTries := 10        ; nombre maximum de tentatives
    interval := 200       ; délai entre chaque tentative (ms)

    ; === Boucle de détection de la fenêtre RadEdit ===
    Loop maxTries {
        target := WinExist("RadEdit ahk_exe RadEdit.exe")
        if (target)
            break
        Sleep interval
    }

    if !target {
        MsgBox "Fenêtre RadEdit introuvable après " (maxTries * interval) " ms."
        ExitApp
    }

    ; === Activation de la fenêtre ===
    WinActivate(target)
    if WinWaitActive("RadEdit ahk_exe RadEdit.exe", , 3)
    {
        Send "^!{Right}"
        return
    }

    ; --- Si la première tentative échoue, on temporise et recommence ---
    Sleep 200
    WinActivate(target)
    if !WinWaitActive("RadEdit ahk_exe RadEdit.exe", , 3) {
        MsgBox "RadEdit introuvable ou inactive après double tentative."
        ExitApp
    }

    Send "^!{Right}"
}
LogFile(target, result, didAttVer) {

    if (didAttVer) {
        loc   := Trim(GetDataContextVarCached(target, "loc", ""))
        modal := Trim(GetDataContextVarCached(target, "modal", ""))

        if (StrLower(loc) = "urgence" && StrUpper(modal) = "CR")
            LogAttVer(result.titres)
    }

    if (HasArg("ToutSUG"))
        LogSUG()

    LogFromDataContext(target)
}


LogAttVer(titres := [], rootPath := "R:\CSSSNL\Bureautique\Imagerie Medicale\Partage\Suivi Urgence") {

    global reqnb_full
    
    ; 1) S’assurer du dossier
    try {
        if !DirExist(rootPath)
            DirCreate(rootPath)
    } catch as err {
        return false
    }

    name    := Trim(GetDataContextVarCached(target, "patnom", ""))

    if (name = "")
        name := "(NAME inconnu)"

    reqTxt := (reqnb_full != "" ? reqnb_full : "(ReqNb inconnu)")

    ; 3) Construire le bloc texte
    ts := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
    buf := reqTxt " - " name "  —  Date de lecture: " ts "`r`n"
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

LogSUG() {
    global target

    ; --- Lire les infos ---
    proc      := Trim(GetDataContextVarCached(target, "proc", ""))
    patdos    := Trim(GetDataContextVarCached(target, "patdos", ""))
    patnom    := Trim(GetDataContextVarCached(target, "patnom", ""))
    studydate := Trim(GetDataContextVarCached(target, "studydate", ""))
    modal     := Trim(GetDataContextVarCached(target, "modal", ""))
    radname   := GetFullName_API()

    ; Si rien de pertinent, sortir silencieusement
    if (proc = "" && patdos = "" && patnom = "" && studydate = "")
        return ""

    sep := ";"  ; Excel FR/Québec

    ; Date/heure au moment de l'entrée
    dictDate := FormatTime(A_Now, "yyyy-MM-dd")
    dictTime := FormatTime(A_Now, "HH:mm:ss")

    ; --- Nom de fichier: SUG_log_radname_date.csv ---
    ;radSafe := Trim(radname)
    ;radSafe := RegExReplace(radSafe, "[\\/:*?`"<>|]", "")
    ;radSafe := RegExReplace(radSafe, "\s+", "_")         
    ;if (radSafe = "")
    ;    radSafe := "Rad"

    fileName := "SUG_log_" radname "_" dictDate ".csv"
    logPath  := A_ScriptDir "\" fileName

    header := "Dossier;Nom patient;Examen;;Date d'examen;Date de dictée;Heure de dictée;Radiologue`r`n"

    ; Helper local CSV
    CsvField := (val) => '"' StrReplace(StrReplace(StrReplace("" (val ?? ""), "`r", " "), "`n", " "), '"', '""') '"'

    try {
        ; Créer / écrire l'en-tête si fichier absent ou vide
        if !FileExist(logPath) || (FileGetSize(logPath) = 0) {
            FileAppend header, logPath, "UTF-8"
        }

        ; Append 1 ligne
        line :=
            CsvField(patdos)    sep
          . CsvField(patnom)    sep
          . CsvField(proc)      sep
          . CsvField(modal)     sep
          . CsvField(studydate) sep
          . CsvField(dictDate)  sep
          . CsvField(dictTime)  sep
          . CsvField(radname)   "`r`n"

        FileAppend line, logPath, "UTF-8"
        return logPath
    } catch as err {
        MsgBox "Impossible d'écrire le SUG log :`n" logPath "`n`nErreur : " err.Message, "ATTENTION", 262144
        return ""
    }
}

LogFromDataContext(target) {

    log_required := Trim(GetDataContextVarCached(target, "log_required", false))

    if (log_required != "true")
        return

    log_type := Trim(GetDataContextVarCached(target, "log_type", ""))

    if (log_type = "")
        return  ; éventuellement: handler "default" 

    ; Optionnel mais utile: rendre le dispatch insensible à la casse
    log_type := StrLower(log_type)

    ; Table de dispatch : type -> fonction
    static handlers := Map(
        "mammoloc", Log_Mammo
        ; plus tard: "sug", Log_SUG
    )

    if !handlers.Has(log_type)
        return  ; ou debug si tu veux (type inconnu)

    try {
        handlers[log_type].Call(target)
    } catch as err {
        ; Optionnel: journaliser l'erreur de dispatch
        ; FileAppend("Log dispatch error: " err.Message "`r`n", A_ScriptDir "\log_dispatch_errors.txt", "UTF-8")
    }
}

Log_Mammo(target) {

    CRLF := "`r`n"
    today := FormatTime(A_Now, "yyyy-MM-dd")  ; <date du jour> pour header + filename

    ; ---- Variables de base ----
    fullName  := GetFullName_API()
    patdos    := Trim(GetDataContextVarCached(target, "patdos", ""))
    patnom    := Trim(GetDataContextVarCached(target, "patnom", ""))
    proc      := Trim(GetDataContextVarCached(target, "proc", ""))
    studydate := Trim(GetDataContextVarCached(target, "studydate", ""))

    ; ---- Nom de fichier (dans le répertoire du script) ----
    ; (Si un jour tu veux sécuriser fullName pour le filename, on pourra le faire.
    ;  Pour l’instant tu dis que c’est safe dans ton environnement.)
    filePath := A_ScriptDir "\LOC - " fullName " - " today ".txt"

    ; ---- Délimiteur si on append dans un fichier existant non-vide ----
    if FileExist(filePath) {
        try {
            if (FileGetSize(filePath) > 0)
                FileAppend(CRLF "###############" CRLF CRLF, filePath, "UTF-8")
        }
    }

    ; ---- Construction du texte (ton format) ----
    entry := ""
    entry .= today CRLF CRLF
    entry .= patdos CRLF CRLF
    entry .= patnom CRLF CRLF
    entry .= "Examen complémentaire: " proc CRLF
    entry .= "du: " studydate CRLF CRLF

    ; ---- log_v1..log_vN : stop au premier vide ----
    ; Ajuste le max si tu veux (50, 100, etc.)
    Loop 50 {
        key := "log_v" A_Index
        val := Trim(GetDataContextVarCached(target, key, ""))

        if (val = "")
            break  ; fin de la liste au premier vide

        entry .= val CRLF CRLF
    }

    entry .= "Requête faite par: " fullName CRLF

    ; ---- Écriture (append) ----
    FileAppend(entry, filePath, "UTF-8")
}


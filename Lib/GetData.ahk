GetDataContextRaw(hwnd, timeout := 800) {
    global CMD, data_context_json

    data_context_json := ""  ; évite un vieux contexte
    SendCopyData(hwnd, CMD["GetDataContext"], "")

    t0 := A_TickCount
    while (data_context_json = "" && A_TickCount - t0 < timeout)
        Sleep 20

    return data_context_json
}

GetDataContextVar(hwnd, key, default := "", timeout := 800) {
    raw := GetDataContextRaw(hwnd, timeout)
    if (raw = "")
        return default

    ; capture "key":"value" (en respectant les \" si jamais)
    patt := '"\Q' key '\E"\s*:\s*"((?:\\.|[^"\\])*)"'
    if RegExMatch(raw, patt, &m)
        return m[1]   ; <-- plus d’unescape

    ; fallback key=value (si jamais)
    patt2 := "im)^\s*\Q" key "\E\s*=\s*(.+?)\s*$"
    if RegExMatch(raw, patt2, &m2)
        return m2[1]

    return default
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
	name := RegExReplace(name, "\d+")    ; <-- AJOUT: enlève les chiffres
        parts := StrSplit(name, " ")
        if parts.Length = 2
            return parts[2] " " parts[1]  ; Inverse : Prénom Nom
        return name
    }

    return ""
}

PrimeDataContextCached(hwnd, timeout := 800) {
    global g_dcRaw, g_dcCache, g_dcHwnd

    g_dcCache := Map()
    g_dcHwnd := hwnd
    g_dcRaw := GetDataContextRaw(hwnd, timeout)  ; <-- utilise TA fonction existante (1 seule fois)
    return g_dcRaw
}

RefreshDataContextCached(hwnd, timeout := 800) {
    ; alias clair si tu veux forcer une relecture en cours de route
    return PrimeDataContextCached(hwnd, timeout)
}

GetDataContextVarCached(hwnd, key, default := "", timeout := 800) {
    global g_dcRaw, g_dcCache, g_dcHwnd

    ; si on change de fenêtre cible, ou si pas encore primé -> on prime
    if (g_dcHwnd != hwnd || g_dcRaw = "") {
        PrimeDataContextCached(hwnd, timeout)
    }

    ; déjà en cache ?
    if g_dcCache.Has(key)
        return g_dcCache[key]

    raw := g_dcRaw
    if (raw = "") {
        g_dcCache[key] := default
        return default
    }

    ; 1) JSON string: "key":"value"
    patt := '"\Q' key '\E"\s*:\s*"((?:\\.|[^"\\])*)"'
    if RegExMatch(raw, patt, &m) {
        g_dcCache[key] := m[1]   ; pas d’unescape, comme ton code actuel
        return m[1]
    }

    ; 2) JSON non-quoted: "key": true/false/null/123 (au cas où)
    patt3 := '"\Q' key '\E"\s*:\s*(true|false|null|-?\d+(?:\.\d+)?)'
    if RegExMatch(raw, patt3, &m3) {
        g_dcCache[key] := m3[1]
        return m3[1]
    }

    ; 3) fallback key=value (si jamais)
    patt2 := "im)^\s*\Q" key "\E\s*=\s*(.+?)\s*$"
    if RegExMatch(raw, patt2, &m2) {
        g_dcCache[key] := m2[1]
        return m2[1]
    }

    g_dcCache[key] := default
    return default
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

GetStudyDate(dateexam) {
    ; Retourne "JJ/MM/AAAA" ou "" si format non reconnu

    dateexam := Trim(dateexam)

    ; 1) Format ISO: YYYY-MM-DD (ex: 2025-10-22)
    if RegExMatch(dateexam, "^\s*(\d{4})-(\d{2})-(\d{2})", &m) {
        y  := m[1], mo := m[2], d := m[3]
        return Format("{:02}/{:02}/{}", d, mo, y)
    }

    ; 2) Format US: MM/DD/YYYY (ex: 10/22/2025)
    if RegExMatch(dateexam, "^\s*(\d{1,2})/(\d{1,2})/(\d{4})", &m) {
        mo := m[1], d := m[2], y := m[3]
        return Format("{:02}/{:02}/{}", d, mo, y)
    }

    return ""
}
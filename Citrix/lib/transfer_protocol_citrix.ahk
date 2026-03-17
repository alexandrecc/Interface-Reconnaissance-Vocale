#Requires AutoHotkey v2

CtxBridgeProtocolVersion() {
    return "1.0"
}

CtxBridgeUtcNowIso() {
    return FormatTime(A_NowUTC, "yyyy-MM-dd'T'HH:mm:ss'Z'")
}

CtxBridgeRoot(rootDir := "") {
    if (rootDir = "")
        rootDir := CtxResolveBridgeRoot()
    if RegExMatch(rootDir, "^[A-Za-z]:[\\/]?$")
        return SubStr(rootDir, 1, 2) "\"
    return RegExReplace(rootDir, "[\\/]+$")
}

CtxBridgeQueueDir(rootDir := "") {
    return CtxBridgeRoot(rootDir) "\queue"
}

CtxBridgeEnsureStructure(rootDir := "") {
    root := CtxBridgeRoot(rootDir)
    queue := CtxBridgeQueueDir(root)
    archive := root "\archive"

    if !DirExist(root)
        DirCreate(root)
    if !DirExist(queue)
        DirCreate(queue)
    if !DirExist(archive)
        DirCreate(archive)
    raw := StrLower(Trim(EnvGet("CITRIX_SINGLE_SLOT")))
    singleSlot := (raw = "" || !(raw = "0" || raw = "false" || raw = "no" || raw = "off"))
    if singleSlot {
        slotName := Trim(EnvGet("CITRIX_SLOT_NAME"))
        if (slotName = "")
            slotName := "slot_current"
        slotDir := queue "\" slotName
        if !DirExist(slotDir)
            DirCreate(slotDir)
    }

    return Map("root", root, "queue", queue, "archive", archive)
}

CtxBridgeReadFlatJson(path) {
    if !FileExist(path)
        throw Error("json file not found: " path)

    text := FileRead(path, "UTF-8")
    out := Map()
    pos := 1

    while RegExMatch(text, '"((?:\\.|[^"\\])*)"\s*:\s*"((?:\\.|[^"\\])*)"', &m, pos) {
        key := CtxJsonUnescape(m[1])
        val := CtxJsonUnescape(m[2])
        out[key] := val
        pos := m.Pos(0) + m.Len(0)
    }
    return out
}

CtxBridgeWriteDone(jobDir, payload := unset) {
    if !IsSet(payload)
        payload := Map()
    if !payload.Has("protocolVersion")
        payload["protocolVersion"] := CtxBridgeProtocolVersion()
    if !payload.Has("completedAtUtc")
        payload["completedAtUtc"] := CtxBridgeUtcNowIso()
    if !payload.Has("status")
        payload["status"] := "done"
    if !payload.Has("worker")
        payload["worker"] := "citrix"

    path := jobDir "\done.json"
    CtxWriteFlatJson(path, payload)
    return path
}

CtxBridgeWriteError(jobDir, payload) {
    if !IsObject(payload)
        payload := Map("message", "" payload)
    if !payload.Has("protocolVersion")
        payload["protocolVersion"] := CtxBridgeProtocolVersion()
    if !payload.Has("completedAtUtc")
        payload["completedAtUtc"] := CtxBridgeUtcNowIso()
    if !payload.Has("status")
        payload["status"] := "error"
    if !payload.Has("worker")
        payload["worker"] := "citrix"
    if !payload.Has("errorCode")
        payload["errorCode"] := "TRANSFER_FAILED"

    path := jobDir "\error.json"
    CtxWriteFlatJson(path, payload)
    return path
}

CtxWriteFlatJson(path, obj) {
    if !IsObject(obj)
        throw Error("obj must be an object map.")

    lines := []
    for key, value in obj
        lines.Push('  "' CtxJsonEscape("" key) '": "' CtxJsonEscape("" value) '"')

    body := "{`r`n" . (lines.Length ? CtxJoin(lines, ",`r`n") : "") . "`r`n}`r`n"
    if FileExist(path)
        FileDelete(path)
    FileAppend(body, path, "UTF-8")
}

CtxJsonEscape(s) {
    s := "" s
    s := StrReplace(s, "\", "\\")
    s := StrReplace(s, '"', '\"')
    s := StrReplace(s, "`r", "\r")
    s := StrReplace(s, "`n", "\n")
    s := StrReplace(s, "`t", "\t")
    return s
}

CtxJsonUnescape(s) {
    s := StrReplace(s, "\r", "`r")
    s := StrReplace(s, "\n", "`n")
    s := StrReplace(s, "\t", "`t")
    s := StrReplace(s, '\"', '"')
    s := StrReplace(s, "\\", "\")
    return s
}

CtxJoin(arr, sep := "") {
    out := ""
    for idx, val in arr {
        if (idx > 1)
            out .= sep
        out .= val
    }
    return out
}

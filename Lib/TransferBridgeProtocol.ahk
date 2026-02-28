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

BridgeEnsureStructure(rootDir := "") {
    root := BridgeRoot(rootDir)
    queue := BridgeQueueDir(root)
    archive := BridgeArchiveDir(root)

    if !DirExist(root)
        DirCreate(root)
    if !DirExist(queue)
        DirCreate(queue)
    if !DirExist(archive)
        DirCreate(archive)

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

BridgeCreateJobFromFile(rootDir, meta, sourceReportPath) {
    if !IsObject(meta)
        throw Error("meta must be an object map.")
    if (sourceReportPath = "" || !FileExist(sourceReportPath))
        throw Error("source report file not found: " sourceReportPath)

    BridgeEnsureStructure(rootDir)

    jobId := meta.Has("jobId") && meta["jobId"] != "" ? meta["jobId"] : BridgeNewJobId()
    jobDir := BridgeJobDir(jobId, rootDir)
    if DirExist(jobDir)
        throw Error("job directory already exists: " jobDir)

    DirCreate(jobDir)

    if !meta.Has("protocolVersion")
        meta["protocolVersion"] := BridgeProtocolVersion()
    if !meta.Has("createdAtUtc")
        meta["createdAtUtc"] := BridgeUtcNowIso()
    meta["jobId"] := jobId

    metaPath := jobDir "\meta.json"
    reportPath := jobDir "\report.rtf"
    readyPath := jobDir "\ready.flag"

    BridgeWriteFlatJson(metaPath, meta)
    FileCopy(sourceReportPath, reportPath, true)
    FileAppend(BridgeUtcNowIso() "`r`n", readyPath, "UTF-8")

    return Map(
        "jobId", jobId,
        "jobDir", jobDir,
        "metaPath", metaPath,
        "reportPath", reportPath,
        "readyPath", readyPath
    )
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
    if FileExist(path)
        FileDelete(path)
    FileAppend(body, path, "UTF-8")
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

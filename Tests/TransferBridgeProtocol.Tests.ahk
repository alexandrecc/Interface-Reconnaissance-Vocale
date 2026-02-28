#Requires AutoHotkey v2
#Include ..\Lib\TransferBridgeProtocol.ahk

Assert(condition, message) {
    if !condition
        throw Error(message)
}

RunTests() {
    testRoot := A_Temp "\\RadEditSync_Test_" BridgeNewJobId("suite")
    srcReport := testRoot "\\source_report.rtf"

    try {
        DirCreate(testRoot)
        FileAppend("{\\rtf1\\ansi Bridge Test}", srcReport, "UTF-8")

        dirs := BridgeEnsureStructure(testRoot)
        Assert(DirExist(dirs["root"]), "Missing root dir")
        Assert(DirExist(dirs["queue"]), "Missing queue dir")
        Assert(DirExist(dirs["archive"]), "Missing archive dir")

        meta := Map(
            "mode", "ToutCitrix",
            "reqnb", "RA202600000001",
            "patdos", "1234567",
            "patnom", "DOE, JOHN",
            "proc", "CT THORAX",
            "studydate", "2026-02-27"
        )

        job := BridgeCreateJobFromFile(testRoot, meta, srcReport)

        Assert(FileExist(job["metaPath"]), "meta.json was not created")
        Assert(FileExist(job["reportPath"]), "report.rtf was not created")
        Assert(FileExist(job["readyPath"]), "ready.flag was not created")

        loadedMeta := BridgeReadFlatJson(job["metaPath"])
        Assert(loadedMeta["protocolVersion"] = BridgeProtocolVersion(), "protocol version mismatch")
        Assert(loadedMeta["jobId"] = job["jobId"], "jobId mismatch")
        Assert(loadedMeta["reqnb"] = "RA202600000001", "reqnb mismatch")

        donePath := BridgeWriteDone(job["jobDir"], Map("jobId", job["jobId"]))
        Assert(FileExist(donePath), "done.json was not created")
        doneMeta := BridgeReadFlatJson(donePath)
        Assert(doneMeta["status"] = "done", "done status mismatch")
        Assert(doneMeta["worker"] = "citrix", "done worker mismatch")

        errPath := BridgeWriteError(job["jobDir"], Map(
            "jobId", job["jobId"],
            "errorCode", "RADIMAGE_TIMEOUT",
            "message", "Window activation timeout"
        ))
        Assert(FileExist(errPath), "error.json was not created")
        errMeta := BridgeReadFlatJson(errPath)
        Assert(errMeta["status"] = "error", "error status mismatch")
        Assert(errMeta["errorCode"] = "RADIMAGE_TIMEOUT", "errorCode mismatch")
    }
    finally {
        try DirDelete(testRoot, 1)
    }
}

try {
    RunTests()
    FileAppend("PASS TransferBridgeProtocol.Tests`n", "*", "UTF-8")
    ExitApp(0)
} catch as err {
    FileAppend("FAIL TransferBridgeProtocol.Tests: " err.Message "`n", "*", "UTF-8")
    ExitApp(1)
}


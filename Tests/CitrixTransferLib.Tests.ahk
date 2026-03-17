#Requires AutoHotkey v2
#Include ..\Citrix\lib\transfer_protocol_citrix.ahk
#Include ..\Citrix\lib\transfer_lib_citrix.ahk

Assert(condition, message) {
    if !condition
        throw Error(message)
}

RunTests() {
    oldRoot := EnvGet("RADEDITSYNC_DIR")
    oldSingleSlot := EnvGet("CITRIX_SINGLE_SLOT")
    testRoot := A_Temp "\\CitrixBridge_Test_" FormatTime(A_NowUTC, "yyyyMMdd_HHmmss") "_" Format("{:06}", Random(0, 999999))

    try {
        EnvSet("RADEDITSYNC_DIR", testRoot)
        EnvSet("CITRIX_SINGLE_SLOT", "0")
        dirs := CtxBridgeEnsureStructure()
        queue := dirs["queue"]

        jobDone := queue "\\job_done"
        jobPick := queue "\\job_pick"
        jobLock := queue "\\job_lock"

        DirCreate(jobDone)
        DirCreate(jobPick)
        DirCreate(jobLock)

        FileAppend("", jobDone "\\ready.flag", "UTF-8")
        FileAppend("", jobDone "\\done.json", "UTF-8")

        FileAppend("", jobPick "\\ready.flag", "UTF-8")
        FileAppend("", jobLock "\\ready.flag", "UTF-8")
        FileAppend("", jobLock "\\processing.lock", "UTF-8")

        FileSetTime("20260101010101", jobDone "\\ready.flag", "M")
        FileSetTime("20260101010102", jobPick "\\ready.flag", "M")
        FileSetTime("20260101010100", jobLock "\\ready.flag", "M")

        nextJob := CtxFindNextReadyJob(queue)
        Assert(CtxGetJobIdFromPath(nextJob) = "job_pick", "CtxFindNextReadyJob mismatch: " nextJob)

        lockPath := jobPick "\\processing.lock"
        Assert(CtxTryCreateLock(lockPath), "CtxTryCreateLock should create lock")
        Assert(!CtxTryCreateLock(lockPath), "CtxTryCreateLock should fail on existing lock")

        donePath := CtxBridgeWriteDone(jobPick, Map("jobId", "job_pick"))
        Assert(FileExist(donePath), "done.json not created")
        done := CtxBridgeReadFlatJson(donePath)
        Assert(done["status"] = "done", "done status mismatch")

        errPath := CtxBridgeWriteError(jobPick, Map("jobId", "job_pick", "errorCode", "X", "message", "boom"))
        Assert(FileExist(errPath), "error.json not created")
        err := CtxBridgeReadFlatJson(errPath)
        Assert(err["status"] = "error", "error status mismatch")
        Assert(err["errorCode"] = "X", "error code mismatch")

        Assert(CtxNormalizeReq("2600-0041") = "26000041", "CtxNormalizeReq mismatch (hyphen format)")
        cand := CtxReqCandidate("26-000041")
        Assert(cand["norm"] = "26000041", "CtxReqCandidate norm mismatch (2-6 format)")
        Assert(cand["score"] >= 6, "CtxReqCandidate score too low (2-6 format)")
        Assert(CtxNormalizeReq("11611") = "", "CtxNormalizeReq should ignore short numeric fields")
    }
    finally {
        EnvSet("RADEDITSYNC_DIR", oldRoot)
        EnvSet("CITRIX_SINGLE_SLOT", oldSingleSlot)
        try DirDelete(testRoot, 1)
    }
}

try {
    RunTests()
    FileAppend("PASS CitrixTransferLib.Tests`n", "*", "UTF-8")
    ExitApp(0)
} catch as err {
    FileAppend("FAIL CitrixTransferLib.Tests: " err.Message "`n", "*", "UTF-8")
    ExitApp(1)
}

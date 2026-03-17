#Requires AutoHotkey v2

#Include ".\lib\transfer_protocol_citrix.ahk"
#Include ".\lib\transfer_lib_citrix.ahk"

; Force bridge root in Citrix session to the mounted share path.
citrixDataRoot := "\\regional.reg14.rtss.qc.ca\app\DragonMedicalOne\Radiologie\Citrix Data"
userSubdir := Trim(A_UserName)
EnvSet("RADEDITSYNC_DIR", (userSubdir != "" ? citrixDataRoot "\" userSubdir : citrixDataRoot))
EnvSet("CITRIX_PERSISTENT_WORKER", "0")
EnvSet("CITRIX_PERF_LOGS", "0")
if (Trim(EnvGet("CITRIX_POLL_MS")) = "")
    EnvSet("CITRIX_POLL_MS", "250")


CtxInitRuntime()
CtxLog("Citrix identity check: A_UserName=" A_UserName " RADEDITSYNC_DIR=" EnvGet("RADEDITSYNC_DIR"))
CtxLog("Citrix worker started. poll=" (CtxShouldRunPersistentWorker() ? "on" : "off") " poll_ms=" CtxPollIntervalMs())

;if CtxShouldRunPersistentWorker() {
;    SetTimer(CtxPollTick, CtxPollIntervalMs())
;}

^!+t:: {
    CtxLog("Manual hotkey trigger received.")
    Transfer_Citrix("manual")
}

^!+s:: {
    CtxLog("Manual signer hotkey trigger received.")
    Transfer_CitrixSignerDirect("manual_signer")
}


;CtxPollTick() {
;    Transfer_Citrix("poll")
;}

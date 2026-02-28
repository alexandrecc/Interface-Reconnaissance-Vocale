#Requires AutoHotkey v2

#Include ".\lib\transfer_protocol_citrix.ahk"
#Include ".\lib\transfer_lib_citrix.ahk"

; Force bridge root in Citrix session to the mounted share path.
EnvSet("RADEDITSYNC_DIR", "\\regional.reg14.rtss.qc.ca\app\DragonMedicalOne\Radiologie\Test Citrix")
EnvSet("CITRIX_PERSISTENT_WORKER", "0")
EnvSet("CITRIX_PERF_LOGS", "1")
if (Trim(EnvGet("CITRIX_POLL_MS")) = "")
    EnvSet("CITRIX_POLL_MS", "250")


CtxInitRuntime()
CtxLog("Citrix worker started. poll=" (CtxShouldRunPersistentWorker() ? "on" : "off") " poll_ms=" CtxPollIntervalMs())

;if CtxShouldRunPersistentWorker() {
;    SetTimer(CtxPollTick, CtxPollIntervalMs())
;}

^!+t:: {
    CtxLog("Manual hotkey trigger received.")
    Transfer_Citrix("manual")
}


;CtxPollTick() {
;    Transfer_Citrix("poll")
;}

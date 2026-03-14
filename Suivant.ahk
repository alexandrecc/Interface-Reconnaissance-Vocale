#Requires AutoHotkey v2

; Ancien Signer()

if WinExist("ahk_exe RadImage.exe") {
    WinActivate
}
if !WinWaitActive("ahk_exe RadImage.exe", , 5) {
    MsgBox "Radimage window not found or not active. Script will exit."
    ExitApp
}
    Sleep 20
    Send "^g"

;Sleep 20

Title := "v5.7"   ; stable part
if hwnd := WinExist(Title " ahk_exe msedge.exe")
{
    WinActivate("ahk_id " hwnd)
}
if !WinWaitActive("ahk_id " hwnd, , 5)
{
    ;MsgBox "Synapse Viewer window not found or not active. Script will exit."
    ExitApp
}
    Sleep 20
    Send "{F8}"


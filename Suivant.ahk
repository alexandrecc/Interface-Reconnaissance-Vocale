#Requires AutoHotkey v2

; Delegate to Transfert.ahk Signer so Suivant always follows
; the same local/Citrix auto-routing and Synapse F8 behavior.
transfertScript := A_ScriptDir "\Transfert.ahk"

if !FileExist(transfertScript) {
    MsgBox "Transfert.ahk introuvable:`n" transfertScript
    ExitApp(1)
}

try {
    exitCode := RunWait('"' A_AhkPath '" "' transfertScript '" Signer', A_ScriptDir)
    ExitApp(exitCode)
} catch as err {
    MsgBox "Impossible d'executer Transfert.ahk Signer.`n`nErreur: " err.Message
    ExitApp(1)
}

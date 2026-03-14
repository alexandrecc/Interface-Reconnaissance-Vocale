#Requires AutoHotkey v2

ResolveTransferArgument(arg, localRadImageRunning := unset) {
    if IsSet(localRadImageRunning)
        localRunning := !!localRadImageRunning
    else
        localRunning := IsLocalRadImageRunning()

    if (arg = "Tout")
        return localRunning ? "Tout" : "ToutCitrix"
    if (arg = "ToutSUG")
        return localRunning ? "ToutSUG" : "ToutSUGCitrix"
    if (arg = "Signer")
        return localRunning ? "Signer" : "SignerCitrix"

    return arg
}

IsLocalRadImageRunning() {
    return WinExist("ahk_exe RadImage.exe") != 0
}

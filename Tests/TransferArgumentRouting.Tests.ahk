#Requires AutoHotkey v2
#Include ..\Lib\TransferArgumentRouting.ahk

Assert(condition, message) {
    if !condition
        throw Error(message)
}

RunTests() {
    Assert(ResolveTransferArgument("Tout", true) = "Tout", "Tout should stay local when RadImage is local")
    Assert(ResolveTransferArgument("Tout", false) = "ToutCitrix", "Tout should route to Citrix when local RadImage is absent")

    Assert(ResolveTransferArgument("ToutSUG", true) = "ToutSUG", "ToutSUG should stay local when RadImage is local")
    Assert(ResolveTransferArgument("ToutSUG", false) = "ToutSUGCitrix", "ToutSUG should route to Citrix when local RadImage is absent")

    Assert(ResolveTransferArgument("Signer", true) = "Signer", "Signer should stay local when RadImage is local")
    Assert(ResolveTransferArgument("Signer", false) = "SignerCitrix", "Signer should route to Citrix when local RadImage is absent")

    Assert(ResolveTransferArgument("ToutCitrix", true) = "ToutCitrix", "Explicit ToutCitrix should remain unchanged")
    Assert(ResolveTransferArgument("SignerCitrix", false) = "SignerCitrix", "Explicit SignerCitrix should remain unchanged")
    Assert(ResolveTransferArgument("Pathologie", false) = "Pathologie", "Non-routed args should remain unchanged")

    ; Omitted optional arg should not throw and should return one of the valid routes.
    routed := ResolveTransferArgument("Tout")
    Assert(routed = "Tout" || routed = "ToutCitrix", "Tout with omitted optional flag returned invalid route")
}

try {
    RunTests()
    FileAppend("PASS TransferArgumentRouting.Tests`n", "*", "UTF-8")
    ExitApp(0)
} catch as err {
    FileAppend("FAIL TransferArgumentRouting.Tests: " err.Message "`n", "*", "UTF-8")
    ExitApp(1)
}

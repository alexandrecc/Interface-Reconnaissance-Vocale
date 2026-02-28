#Requires AutoHotkey v2
#Include ..\Lib\TransferShared.ahk

Assert(condition, message) {
    if !condition
        throw Error(message)
}

RunTests() {
    outHex := DecodeRTFHex("A\'c9B")
    Assert(StrLen(outHex) = 3, "DecodeRTFHex length mismatch")
    Assert(Ord(SubStr(outHex, 2, 1)) = 0xC9, "DecodeRTFHex codepoint mismatch")

    outUni := DecodeRTFUnicode("x\u233?y")
    Assert(StrLen(outUni) = 3, "DecodeRTFUnicode length mismatch")
    Assert(Ord(SubStr(outUni, 2, 1)) = 233, "DecodeRTFUnicode codepoint mismatch")

    Assert(DecodeRTFHex("plain") = "plain", "DecodeRTFHex should preserve plain text")
    Assert(DecodeRTFUnicode("plain") = "plain", "DecodeRTFUnicode should preserve plain text")

    Assert(HasArg("alpha"), "HasArg should match alpha")
    Assert(HasArg("BETA"), "HasArg should be case-insensitive")
    Assert(!HasArg("gamma"), "HasArg should not match missing arg")
}

try {
    RunTests()
    FileAppend("PASS TransferShared.Tests`n", "*", "UTF-8")
    ExitApp(0)
} catch as err {
    FileAppend("FAIL TransferShared.Tests: " err.Message "`n", "*", "UTF-8")
    ExitApp(1)
}


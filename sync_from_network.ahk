#Requires AutoHotkey v2.0
; --- Settings ----------------------------------------------------
; Absolute path to your config.csv
configFile := ".\sync_config.csv"

textesFolderName := "Textes"  ; name of the subfolder to use
; ----------------------------------------------------------------

; Get current Windows username
userName := A_UserName
userKey  := StrLower(userName)

; Base directory = current working directory (where script is run from)
baseDir := A_WorkingDir
targetPath := ""
syncExitCode := 1

; Make sure the config file exists
if !FileExist(configFile) {
    MsgBox "Config file not found:`n" configFile
    ExitApp syncExitCode
}

ShowSyncStatus(title, message) {
    statusGui := Gui("+AlwaysOnTop -SysMenu +ToolWindow", title)
    statusGui.MarginX := 14
    statusGui.MarginY := 12
    statusGui.SetFont("s10", "Segoe UI")
    statusGui.Add("Text", "w420", message)
    statusGui.Show("AutoSize Center")
    Sleep 75
    return statusGui
}

HideSyncStatus(statusGui) {
    if !IsObject(statusGui)
        return
    try statusGui.Destroy()
}

PowerShellLiteral(value) {
    return "'" StrReplace(value, "'", "''") "'"
}

NewTempFilePath(prefix, extension := ".txt") {
    return A_Temp "\" prefix "_" A_TickCount "_" Random(1000, 9999) extension
}

CopyDirectoryResponsive(sourceDir, destDir) {
    resultFile := NewTempFilePath("sync_from_network")
    powershellExe := A_WinDir "\System32\WindowsPowerShell\v1.0\powershell.exe"
    psCommand := "$ErrorActionPreference='Stop';"
        . "try {"
        . "if (-not (Test-Path -LiteralPath " PowerShellLiteral(destDir) ")) { "
        . "New-Item -ItemType Directory -Path " PowerShellLiteral(destDir) " -Force | Out-Null "
        . "};"
        . "Get-ChildItem -LiteralPath " PowerShellLiteral(sourceDir) " -Force | ForEach-Object { "
        . "Copy-Item -LiteralPath $_.FullName -Destination " PowerShellLiteral(destDir) " -Recurse -Force "
        . "};"
        . "Set-Content -LiteralPath " PowerShellLiteral(resultFile) " -Value 'OK' -Encoding UTF8"
        . "} catch {"
        . "Set-Content -LiteralPath " PowerShellLiteral(resultFile) " -Value $_.Exception.Message -Encoding UTF8;"
        . "exit 1"
        . "}"
    commandLine := '"' powershellExe '" -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -Command "' psCommand '"'

    try Run(commandLine, , "Hide", &copyPid)
    catch Error as e {
        throw Error("Unable to start the copy worker.`n`n" e.Message)
    }

    while ProcessExist(copyPid)
        Sleep 100

    if !FileExist(resultFile)
        throw Error("The copy worker ended without returning a result.")

    try {
        resultText := Trim(FileRead(resultFile, "UTF-8"), "`r`n`t ")
    } finally {
        try FileDelete(resultFile)
    }

    if (resultText != "OK")
        throw Error(resultText = "" ? "The copy worker failed without an error message." : resultText)
}

; --- Read CSV and find matching user ---------------------------------
Loop Read, configFile {
    line := A_LoopReadLine

    ; Skip header (first line) if it contains "Username"
    if (A_Index = 1 && InStr(StrLower(line), "username"))
        continue

    if !Trim(line)
        continue

    ; Split on first comma: Username, TargetPath
    fields := StrSplit(line, ",", , 2)
    if (fields.Length < 2)
        continue

    name := StrLower(Trim(fields[1], " `t`""))  ; trim spaces, tabs, quotes

    if (name = userKey) {
        rawPath := Trim(fields[2], " `t")

        ; Remove surrounding quotes around path, if any
        if (SubStr(rawPath, 1, 1) = '"' && SubStr(rawPath, -1) = '"')
            rawPath := SubStr(rawPath, 2, StrLen(rawPath) - 2)

        targetPath := rawPath
        break
    }
}

if (targetPath = "") {
    MsgBox "No matching entry found in the config for user: " userName
    ExitApp syncExitCode
}

; --- Determine source & destination ----------------------------------

; Source: <TargetPath>\Textes
textesSource := targetPath "\" textesFolderName

; Destination in current working directory: .\Textes
destTextesDir := baseDir "\" textesFolderName

; Optional fallback: .\Default (inside current working directory)
defaultDir := baseDir "\Default"

; --- Step 1: erase existing .\Textes if it exists --------------------
if DirExist(destTextesDir) {
    try {
        DirDelete(destTextesDir, 1)  ; 1 = recursive
    } catch Error as e {
        MsgBox "Error deleting existing Textes directory:`n" destTextesDir "`n`n" e.Message
        ExitApp syncExitCode
    }
}

; --- Step 2: choose source (Textes from TargetPath OR Default) -------

if DirExist(textesSource) {
    ; Copy from user's TargetPath\Textes
    statusGui := ShowSyncStatus(
        "Synchronisation",
        "Copie des fichiers Textes en cours...`n`n"
        . "Source:`n" textesSource "`n`n"
        . "Destination:`n" destTextesDir "`n`n"
        . "Veuillez patienter."
    )
    try {
        CopyDirectoryResponsive(textesSource, destTextesDir)
        HideSyncStatus(statusGui)
        syncExitCode := 0
        MsgBox "Copied '" textesSource "' to:`n" destTextesDir
    } catch Error as e {
        HideSyncStatus(statusGui)
        MsgBox "Error copying from source Textes:`n" textesSource "`n`n" e.Message
    }
} else if DirExist(defaultDir) {
    ; No Textes in TargetPath → use ./Default
    statusGui := ShowSyncStatus(
        "Synchronisation",
        "Copie du dossier par défaut en cours...`n`n"
        . "Source:`n" defaultDir "`n`n"
        . "Destination:`n" destTextesDir "`n`n"
        . "Veuillez patienter."
    )
    try {
        CopyDirectoryResponsive(defaultDir, destTextesDir)
        HideSyncStatus(statusGui)
        syncExitCode := 0
        MsgBox "No '" textesFolderName "' in target path.`nUsed default folder instead:`n" defaultDir "`n→ `n" destTextesDir
    } catch Error as e {
        HideSyncStatus(statusGui)
        MsgBox "Error copying from Default directory:`n" defaultDir "`n`n" e.Message
    }
} else {
    MsgBox "No '" textesFolderName "' directory found at:`n" textesSource "`n`nAnd no 'Default' directory found at:`n" defaultDir
}

ExitApp syncExitCode

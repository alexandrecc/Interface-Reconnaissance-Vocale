#Requires AutoHotkey v2.0
; --- Settings ----------------------------------------------------
configFile := ".\sync_config.csv"
textesFolderName := "Textes"
; -----------------------------------------------------------------

userName := A_UserName
userKey  := StrLower(userName)

; Base directory = where the script is RUN FROM
baseDir := A_WorkingDir
localTextesDir := baseDir "\" textesFolderName

; Check local Textes
if !DirExist(localTextesDir) {
    MsgBox "Local '" textesFolderName "' directory not found in:`n" baseDir
    ExitApp
}

; Check config file
if !FileExist(configFile) {
    MsgBox "Config file not found:`n" configFile
    ExitApp
}

DriveLetter := "R:"
SharePath   := "\\regional.reg14.rtss.qc.ca"
if !FileExist(DriveLetter . "\")
{
    Cmd := "net use " . DriveLetter . " " . SharePath . " /persistent:no"
    RunWait(A_ComSpec " /c " Cmd, , "Hide")
}


targetPath := ""

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
    resultFile := NewTempFilePath("sync_to_network")
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
    ExitApp
}

if !DirExist(targetPath) {
    MsgBox "TargetPath does not exist or is not accessible:`n" targetPath
    ExitApp
}

; -------- Destination Textes dir ------------------------------------
targetTextesDir := targetPath "\" textesFolderName

; -------- Copy local .\Textes -> TargetPath\Textes ------------------
statusGui := ShowSyncStatus(
    "Synchronisation",
    "Copie des fichiers Textes vers le dossier réseau en cours...`n`n"
    . "Source:`n" localTextesDir "`n`n"
    . "Destination:`n" targetTextesDir "`n`n"
    . "Veuillez patienter."
)
try {
    CopyDirectoryResponsive(localTextesDir, targetTextesDir)
    HideSyncStatus(statusGui)
    MsgBox "Copied local '" localTextesDir "' to:`n" targetTextesDir
} catch Error as e {
    HideSyncStatus(statusGui)
    MsgBox "Error copying to target Textes directory:`n" targetTextesDir "`n`n" e.Message
}

ExitApp

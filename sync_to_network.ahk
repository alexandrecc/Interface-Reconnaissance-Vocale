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
try {
    ; 1 = overwrite existing files/dirs
    DirCopy(localTextesDir, targetTextesDir, 1)
    MsgBox "Copied local '" localTextesDir "' to:`n" targetTextesDir
} catch Error as e {
    MsgBox "Error copying to target Textes directory:`n" targetTextesDir "`n`n" e.Message
}

ExitApp

@echo off
setlocal

rem Usage:
rem   cache_insertions_to_bridge.bat [bridge_root]
rem Example bridge_root:
rem   \\regional.reg14.rtss.qc.ca\app\DragonMedicalOne\Radiologie\Citrix Data\jdoe

if "%~1"=="" (
    set "BRIDGE_ROOT=\\regional.reg14.rtss.qc.ca\app\DragonMedicalOne\Radiologie\Citrix Data\%USERNAME%"
) else (
    set "BRIDGE_ROOT=%~1"
)

set "SCRIPT_DIR=%~dp0"
set "CONFIG_FILE=%SCRIPT_DIR%..\sync_config.csv"
set "DEFAULT_DIR=%SCRIPT_DIR%..\Default\Insertions"
set "CACHE_DIR=%BRIDGE_ROOT%\Textes\Insertions"
set "SLOT_NAME=%CITRIX_SLOT_NAME%"
if "%SLOT_NAME%"=="" set "SLOT_NAME=slot_current"

if not exist "%BRIDGE_ROOT%" mkdir "%BRIDGE_ROOT%"
if not exist "%BRIDGE_ROOT%\queue" mkdir "%BRIDGE_ROOT%\queue"
if not exist "%BRIDGE_ROOT%\archive" mkdir "%BRIDGE_ROOT%\archive"
if not exist "%BRIDGE_ROOT%\queue\%SLOT_NAME%" mkdir "%BRIDGE_ROOT%\queue\%SLOT_NAME%"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference = 'Stop';" ^
  "try {" ^
  "  $configFile = [System.IO.Path]::GetFullPath($env:CONFIG_FILE);" ^
  "  $defaultDir = [System.IO.Path]::GetFullPath($env:DEFAULT_DIR);" ^
  "  $userKey = $env:USERNAME.Trim().ToLowerInvariant();" ^
  "  $targetPath = '';" ^
  "  if (Test-Path -LiteralPath $configFile) {" ^
  "    $rows = Get-Content -LiteralPath $configFile;" ^
  "    $row = $rows | Select-Object -Skip 1 | ConvertFrom-Csv -Header 'Username', 'TargetPath' | Where-Object { $_.Username -and $_.Username.Trim().ToLower() -eq $userKey } | Select-Object -First 1;" ^
  "    if ($row -and $row.TargetPath) { $targetPath = ([string]$row.TargetPath).Trim() }" ^
  "  }" ^
  "  $networkDir = if ($targetPath) { Join-Path $targetPath 'Textes\Insertions' } else { '' };" ^
  "  if ($networkDir -and (Test-Path -LiteralPath $networkDir)) {" ^
  "    $srcDir = $networkDir" ^
  "  } elseif (Test-Path -LiteralPath $defaultDir) {" ^
  "    $srcDir = $defaultDir" ^
  "  } else {" ^
  "    throw 'No source Insertions directory found on the user network path and no Default\Insertions fallback is available.'" ^
  "  }" ^
  "  $required = @('textesRapport.txt', 'textefinal.rtf', 'textefinalSUG.rtf');" ^
  "  $missing = @($required | Where-Object { -not (Test-Path -LiteralPath (Join-Path $srcDir $_)) });" ^
  "  if ($missing.Count -gt 0) {" ^
  "    throw ('Missing source file(s): ' + ($missing -join ', '))" ^
  "  }" ^
  "  New-Item -ItemType Directory -Path $env:CACHE_DIR -Force | Out-Null;" ^
  "  $required | ForEach-Object { Copy-Item -LiteralPath (Join-Path $srcDir $_) -Destination $env:CACHE_DIR -Force };" ^
  "  Write-Output ('[cache] Insertion artifacts cached from ' + $srcDir + ' to ' + $env:CACHE_DIR);" ^
  "} catch {" ^
  "  Write-Output ('[cache] ' + $_.Exception.Message);" ^
  "  exit 1" ^
  "}"

if errorlevel 1 exit /b 1

exit /b 0

@echo off

set "SCRIPT=listener_forwarder.ps1"

set "SRC=\\regional.reg14.rtss.qc.ca\app\DragonMedicalOne\Radiologie\Interface Reconnaissance Vocale"


set "DST=C:\APP\Interface Reconnaissance Vocale"
set "ROBO_OPTS=/E /XO /R:1 /W:1"
robocopy "%SRC%" "%DST%" %ROBO_OPTS%

pushd "%DST%"

set "AHK_EXE=%ProgramFiles%\AutoHotkey\v2\AutoHotkey64.exe"
if not exist "%AHK_EXE%" set "AHK_EXE=%ProgramFiles%\AutoHotkey\v2\AutoHotkey.exe"
if not exist "%AHK_EXE%" set "AHK_EXE="
if not defined AHK_EXE (
  for %%I in (AutoHotkey64.exe AutoHotkey.exe) do (
    if not defined AHK_EXE for /f "delims=" %%J in ('where %%I 2^>nul') do set "AHK_EXE=%%J"
  )
)

if exist ".\sync_from_network.ahk" (
  if defined AHK_EXE (
    start "" "%AHK_EXE%" ".\sync_from_network.ahk"
  ) else (
    start "" ".\sync_from_network.ahk"
  )
)

if exist ".\Script\cache_insertions_to_bridge.bat" (
  call ".\Script\cache_insertions_to_bridge.bat" >nul 2>&1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Unblock-File -LiteralPath '.\RadEdit.exe'"
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Unblock-File -LiteralPath '.\listener_forwarder.ps1'"

if exist "%USERPROFILE%\pause.flag" del /f /q "%USERPROFILE%\pause.flag"
taskkill /F /T /IM "FusionDictate.exe" >nul 2>&1
start "" "C:\Nuance\DMO\SoD.exe"
start "" ".\RadEdit.exe"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$procs = Get-CimInstance Win32_Process -Filter \"Name='powershell.exe'\" |" ^
  " Where-Object { $_.CommandLine -match [regex]::Escape('%SCRIPT%') -and" ^
  "                (Invoke-CimMethod -InputObject $_ -MethodName GetOwner).User -eq $env:USERNAME };" ^
  " $procs | ForEach-Object { try { Stop-Process -Id $_.ProcessId -Force -ErrorAction Stop } catch {} }"


powershell -NoProfile -WindowStyle Hidden -Command "Start-Process -WindowStyle Hidden -FilePath 'powershell.exe' -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File','.\listener_forwarder.ps1'"

popd

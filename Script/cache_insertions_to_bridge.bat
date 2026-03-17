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

set "SRC_DIR=%~dp0..\Textes\Insertions"
set "CACHE_DIR=%BRIDGE_ROOT%\Textes\Insertions"
set "SLOT_NAME=%CITRIX_SLOT_NAME%"
if "%SLOT_NAME%"=="" set "SLOT_NAME=slot_current"

if not exist "%BRIDGE_ROOT%" mkdir "%BRIDGE_ROOT%"
if not exist "%BRIDGE_ROOT%\queue" mkdir "%BRIDGE_ROOT%\queue"
if not exist "%BRIDGE_ROOT%\archive" mkdir "%BRIDGE_ROOT%\archive"
if not exist "%BRIDGE_ROOT%\queue\%SLOT_NAME%" mkdir "%BRIDGE_ROOT%\queue\%SLOT_NAME%"

if not exist "%SRC_DIR%\textesRapport.txt" (
    echo [cache] Missing source file: "%SRC_DIR%\textesRapport.txt"
    exit /b 2
)
if not exist "%SRC_DIR%\textefinal.rtf" (
    echo [cache] Missing source file: "%SRC_DIR%\textefinal.rtf"
    exit /b 2
)
if not exist "%SRC_DIR%\textefinalSUG.rtf" (
    echo [cache] Missing source file: "%SRC_DIR%\textefinalSUG.rtf"
    exit /b 2
)

if not exist "%CACHE_DIR%" mkdir "%CACHE_DIR%"

robocopy "%SRC_DIR%" "%CACHE_DIR%" "textesRapport.txt" "textefinal.rtf" "textefinalSUG.rtf" /R:1 /W:1 /NP /NFL /NDL >nul
if errorlevel 8 (
    echo [cache] Failed to copy insertion artifacts to "%CACHE_DIR%"
    exit /b 1
)

echo [cache] Insertion artifacts cached in "%CACHE_DIR%"
exit /b 0

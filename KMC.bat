@echo off
cd /d %~dp0

powershell -NoProfile  -executionpolicy RemoteSigned -File "%~dp0KarutaMovieClip.ps1"
call :isSuccess

:isSuccess
if not %errorlevel% == 0 (
    exit 1
)
exit /b 0

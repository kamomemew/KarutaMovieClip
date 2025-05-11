@echo on
cd /d %~dp0
powershell -executionpolicy RemoteSigned -File "%~dp0KarutaMovieClip.ps1"
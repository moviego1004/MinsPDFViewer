@echo off
rem Publish + installer in one step. Output: dist\MinsPDFViewer_Setup_v<version>.exe
setlocal
call "%~dp0publish.bat" || exit /b 1
call "%~dp0makeSetup.bat" || exit /b 1

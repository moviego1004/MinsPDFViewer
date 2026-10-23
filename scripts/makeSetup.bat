@echo off
rem Build the installer from build\publish (run publish.bat first, or use release.bat).
rem Output: dist\MinsPDFViewer_Setup_v<version>.exe
setlocal
cd /d "%~dp0"
set "MAKENSIS=C:\Program Files (x86)\NSIS\Bin\makensis.exe"
if not exist "%MAKENSIS%" (
  echo makensis.exe not found: "%MAKENSIS%"
  exit /b 1
)
if not exist "..\dist" mkdir "..\dist"
"%MAKENSIS%" installer.nsi || exit /b 1

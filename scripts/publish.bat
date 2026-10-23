@echo off
rem Publish Release (self-contained win-x64) to build\publish (used by installer.nsi).
setlocal
cd /d "%~dp0.."
if exist "build\publish" rmdir /s /q "build\publish"
dotnet publish src\MinsPDFViewer\MinsPDFViewer.csproj -c Release -o build\publish || exit /b 1

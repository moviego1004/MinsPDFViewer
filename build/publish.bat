@echo off
rem Publish Release (self-contained win-x64).
rem Output: src\MinsPDFViewer\bin\Release\net8.0-windows10.0.19041.0\win-x64\publish
setlocal
cd /d "%~dp0.."
dotnet publish src\MinsPDFViewer\MinsPDFViewer.csproj -c Release || exit /b 1

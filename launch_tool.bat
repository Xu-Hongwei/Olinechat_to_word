@echo off
setlocal

cd /d "%~dp0"

set "PYTHON_EXE=D:\APP\miniconda3\python.exe"
if not exist "%PYTHON_EXE%" set "PYTHON_EXE=python"

if exist ".\dist\ChatMarkdownToWord\ChatMarkdownToWord.exe" (
  start "" ".\dist\ChatMarkdownToWord\ChatMarkdownToWord.exe"
  exit /b 0
)

if exist "..\software\dist\ChatMarkdownToWord\ChatMarkdownToWord.exe" (
  start "" "..\software\dist\ChatMarkdownToWord\ChatMarkdownToWord.exe"
  exit /b 0
)

"%PYTHON_EXE%" app.py
endlocal

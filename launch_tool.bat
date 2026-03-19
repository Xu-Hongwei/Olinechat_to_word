@echo off
cd /d "%~dp0"
if exist ".\dist\ChatMarkdownToWord\ChatMarkdownToWord.exe" (
  start "" ".\dist\ChatMarkdownToWord\ChatMarkdownToWord.exe"
) else (
  D:\APP\miniconda3\python.exe app.py
)

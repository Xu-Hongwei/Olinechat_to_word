@echo off
cd /d "%~dp0"
D:\APP\miniconda3\python.exe -m PyInstaller --noconfirm --clean --windowed --name ChatMarkdownToWord --collect-data docx --hidden-import lxml._elementpath app.py

@echo off
setlocal

cd /d "%~dp0"

set "PYTHON_EXE=D:\APP\miniconda3\python.exe"
if not exist "%PYTHON_EXE%" set "PYTHON_EXE=python"

set "ROOT_DIR=%~dp0.."
set "SOFTWARE_DIST=%ROOT_DIR%\software\dist"
set "APP_DIR=.\dist\ChatMarkdownToWord"
set "APP_EXE=%APP_DIR%\ChatMarkdownToWord.exe"
set "PORTABLE_ZIP=.\dist\ChatMarkdownToWord-portable.zip"

echo [1/5] Cleaning old build outputs...
if exist ".\build" rmdir /s /q ".\build"
if exist ".\dist" rmdir /s /q ".\dist"

echo [2/5] Building executable with PyInstaller...
"%PYTHON_EXE%" -m PyInstaller --noconfirm --clean --windowed --name ChatMarkdownToWord --collect-data docx --hidden-import lxml._elementpath app.py
if errorlevel 1 (
  echo Build failed.
  exit /b 1
)

if not exist "%APP_EXE%" (
  echo Build output not found: %APP_EXE%
  exit /b 1
)

echo [3/5] Creating portable zip...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Compress-Archive -Path '.\dist\ChatMarkdownToWord\*' -DestinationPath '.\dist\ChatMarkdownToWord-portable.zip' -Force"
if errorlevel 1 (
  echo Zip creation failed.
  exit /b 1
)

echo [4/5] Syncing artifacts to software\dist...
if not exist "%SOFTWARE_DIST%" mkdir "%SOFTWARE_DIST%"
if exist "%SOFTWARE_DIST%\ChatMarkdownToWord" rmdir /s /q "%SOFTWARE_DIST%\ChatMarkdownToWord"
xcopy "%APP_DIR%" "%SOFTWARE_DIST%\ChatMarkdownToWord\" /E /I /Y >nul
copy /y "%PORTABLE_ZIP%" "%SOFTWARE_DIST%\ChatMarkdownToWord-portable.zip" >nul

echo [5/5] Done.
echo EXE: %APP_EXE%
echo ZIP: %PORTABLE_ZIP%
echo SOFTWARE DIST: %SOFTWARE_DIST%

endlocal

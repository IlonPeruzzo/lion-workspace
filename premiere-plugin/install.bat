@echo off
echo ============================================
echo   Lion Workspace - Premiere Pro Plugin
echo   Instalador
echo ============================================
echo.

:: Enable unsigned CEP extensions (dev mode)
echo Habilitando extensoes CEP...
reg add "HKCU\SOFTWARE\Adobe\CSXS.11" /v PlayerDebugMode /t REG_SZ /d 1 /f >nul 2>&1
reg add "HKCU\SOFTWARE\Adobe\CSXS.12" /v PlayerDebugMode /t REG_SZ /d 1 /f >nul 2>&1

:: Copy plugin to Adobe CEP extensions folder
set "DEST=%APPDATA%\Adobe\CEP\extensions\com.lionworkspace.premiere"

echo Instalando plugin em: %DEST%
if exist "%DEST%" rmdir /s /q "%DEST%"
mkdir "%DEST%\CSXS" 2>nul
mkdir "%DEST%\client" 2>nul

xcopy /e /y "%~dp0CSXS\*" "%DEST%\CSXS\" >nul
xcopy /e /y "%~dp0client\*" "%DEST%\client\" >nul
copy /y "%~dp0.debug" "%DEST%\" >nul

echo.
echo ============================================
echo   Instalado com sucesso!
echo.
echo   Para usar:
echo   1. Abra o Lion Workspace
echo   2. Abra o Premiere Pro
echo   3. Janela ^> Extensoes ^> Lion Workspace
echo ============================================
echo.
pause

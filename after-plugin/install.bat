@echo off
setlocal
echo ============================================
echo  Lion Workspace - After Effects Plugin
echo ============================================
echo.
set CEP_DIR=%APPDATA%\Adobe\CEP\extensions
set TARGET=%CEP_DIR%\com.lionworkspace.after

if not exist "%CEP_DIR%" mkdir "%CEP_DIR%"
if exist "%TARGET%" rmdir /S /Q "%TARGET%"
mkdir "%TARGET%"

xcopy /E /I /Y /Q "%~dp0CSXS"   "%TARGET%\CSXS"   >nul
xcopy /E /I /Y /Q "%~dp0client" "%TARGET%\client" >nul
xcopy /E /I /Y /Q "%~dp0host"   "%TARGET%\host"   >nul

REM Habilitar carregamento de extensoes nao assinadas (CEP 11)
reg add "HKCU\Software\Adobe\CSXS.11" /v PlayerDebugMode /t REG_SZ /d 1 /f >nul

echo.
echo  Plugin instalado em: %TARGET%
echo  Abra o After Effects e va em: Window ^> Extensions ^> Lion Workspace
echo.
pause
endlocal

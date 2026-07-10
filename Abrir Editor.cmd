@echo off
REM Lanza el editor de hojas (busca el .exe compilado o abre el visor web).
setlocal
cd /d "%~dp0"

set "EXE=app\bin\Release\net8.0-windows\XlsViewApp.exe"
if exist "%EXE%" (
  start "" "%EXE%" %*
  exit /b 0
)

echo No se encontro el ejecutable compilado.
echo Compila primero con:  app\build.cmd
pause

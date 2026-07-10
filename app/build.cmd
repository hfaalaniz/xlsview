@echo off
REM ============================================================
REM  Compila XlsViewApp.exe (browser propietario WebView2).
REM  Requiere el SDK de .NET y el runtime WebView2 (presente en
REM  Windows 10/11 modernos).
REM ============================================================
setlocal
cd /d "%~dp0"

where dotnet >nul 2>nul
if errorlevel 1 (
  echo No se encontro 'dotnet'. Instala el SDK de .NET 8 o superior.
  exit /b 1
)

echo Compilando (Release)...
dotnet build -c Release
if errorlevel 1 (
  echo *** Fallo la compilacion ***
  exit /b 1
)

echo.
echo OK. Ejecutable en:
echo   app\bin\Release\net8.0-windows\XlsViewApp.exe
echo.
echo Para asociar las extensiones y anadir al PATH:
echo   app\bin\Release\net8.0-windows\XlsViewApp.exe --install

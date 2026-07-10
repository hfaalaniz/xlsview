@echo off
setlocal enabledelayedexpansion
rem ============================================================================
rem  deploy.cmd  -  Publica XlsView a una carpeta ESTABLE y registra la
rem                 asociacion desde alli (no desde app\bin\Release, que se
rem                 destruye al recompilar y rompe "Abrir con").
rem
rem  Destino:  %LOCALAPPDATA%\Programs\XlsView
rem  Copia el exe self-contained + todos los assets web (index.html, *.js, lib\)
rem  y ejecuta --install para asociar .xlsx/.xlsm/.xls/.csv y anadir al PATH.
rem
rem  Uso:
rem     deploy.cmd            publica, copia y registra la asociacion
rem     deploy.cmd /noinstall publica y copia, sin tocar el registro
rem ============================================================================

set "ROOT=%~dp0"
set "ROOT=%ROOT:~0,-1%"
set "DEST=%LOCALAPPDATA%\Programs\XlsView"
set "PUBTMP=%ROOT%\app\bin\publish"

echo === XlsView deploy ===
echo   Origen : %ROOT%
echo   Destino: %DEST%
echo.

echo [1/4] Publicando el ejecutable (self-contained win-x64)...
dotnet publish "%ROOT%\app\XlsViewApp.csproj" -c Release -r win-x64 --self-contained true ^
    -p:PublishSingleFile=false -o "%PUBTMP%" -v q
if errorlevel 1 (
    echo   ERROR: fallo dotnet publish.
    exit /b 1
)

echo [2/4] Preparando carpeta estable...
if not exist "%DEST%" mkdir "%DEST%"
rem Limpiar solo binarios previos (conserva WebView2 user data si existiera).
del /q "%DEST%\*.dll" 2>nul
del /q "%DEST%\*.exe" 2>nul
del /q "%DEST%\*.json" 2>nul
del /q "%DEST%\*.pdb" 2>nul

echo [3/4] Copiando binarios y assets web...
rem Binarios publicados (exe + runtime .NET + WebView2).
xcopy "%PUBTMP%\*" "%DEST%\" /E /I /Y /Q >nul
rem Assets web del visor (deben quedar junto al exe: ResolveWebRoot los busca ahi).
copy /Y "%ROOT%\index.html"     "%DEST%\" >nul
copy /Y "%ROOT%\app.js"         "%DEST%\" >nul
copy /Y "%ROOT%\xlsx-styles.js" "%DEST%\" >nul
copy /Y "%ROOT%\print.js"       "%DEST%\" >nul
if exist "%DEST%\lib" rmdir /s /q "%DEST%\lib"
xcopy "%ROOT%\lib" "%DEST%\lib\" /E /I /Y /Q >nul

if /I "%~1"=="/noinstall" (
    echo [4/4] Omitido --install ^(/noinstall^).
    echo.
    echo Listo. Ejecutable estable en: %DEST%\XlsViewApp.exe
    exit /b 0
)

echo [4/4] Registrando asociacion desde la carpeta estable...
"%DEST%\XlsViewApp.exe" --install

echo.
echo Listo. XlsView quedo instalado en: %DEST%
echo Si Windows no lo usa por defecto: clic derecho en un .xlsx -^> Abrir con
echo   -^> Elegir otra aplicacion -^> XlsView -^> Siempre.
endlocal

@echo off
rem Abre el Convertidor Markdown. Tambien se pueden soltar archivos sobre este .bat.
rem No depende de una version concreta de Python: usa el que haya instalado,
rem y la propia app se encarga de encontrar (o instalar) MarkItDown.
setlocal
set "APP=%~dp0convertidor_markdown.pyw"

where pyw >nul 2>nul
if not errorlevel 1 goto con_pyw
where pythonw >nul 2>nul
if not errorlevel 1 goto con_pythonw

set "PYW="
for /d %%D in ("%LOCALAPPDATA%\Python\pythoncore-*" "%LOCALAPPDATA%\Programs\Python\Python3*") do if exist "%%D\pythonw.exe" set "PYW=%%D\pythonw.exe"
if defined PYW goto con_ruta

echo No se encontro Python en este equipo.
echo Instalalo desde https://www.python.org/downloads/ y vuelve a abrir este archivo.
pause
exit /b 1

:con_pyw
start "" pyw "%APP%" %*
exit /b 0

:con_pythonw
start "" pythonw "%APP%" %*
exit /b 0

:con_ruta
start "" "%PYW%" "%APP%" %*
exit /b 0

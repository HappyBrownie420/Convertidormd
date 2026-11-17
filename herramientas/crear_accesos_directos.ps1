# Crea accesos directos del Convertidor Markdown en el Escritorio y en el menu Inicio.
#
# Apuntan al lanzador de Python (pyw.exe), no a una version concreta, asi que
# siguen funcionando aunque se actualice Python. Al apuntar a un .exe, Windows
# permite anclarlos a Inicio o a la barra de tareas (con un .bat no deja).
#
# Uso:  powershell -ExecutionPolicy Bypass -File herramientas\crear_accesos_directos.ps1
#
# (Este archivo va sin tildes a proposito: PowerShell 5.1 las lee mal.)

$ErrorActionPreference = 'Stop'

$raiz  = Split-Path -Parent $PSScriptRoot
$app   = Join-Path $raiz 'convertidor_markdown.pyw'
$icono = Join-Path $raiz 'icono.ico'

$lanzador = (Get-Command pyw.exe -ErrorAction SilentlyContinue | Select-Object -First 1).Source
if (-not $lanzador) {
    $lanzador = (Get-Command pythonw.exe -ErrorAction SilentlyContinue | Select-Object -First 1).Source
}
if (-not $lanzador) {
    throw 'No se encontro Python (ni pyw.exe ni pythonw.exe). Instalalo desde https://www.python.org/downloads/'
}

$shell = New-Object -ComObject WScript.Shell
$carpetas = @(
    [Environment]::GetFolderPath('Desktop'),
    [Environment]::GetFolderPath('Programs')      # menu Inicio > Todas las aplicaciones
)
foreach ($carpeta in $carpetas) {
    $ruta = Join-Path $carpeta 'Convertidor Markdown.lnk'
    $acceso = $shell.CreateShortcut($ruta)
    $acceso.TargetPath = $lanzador
    $acceso.Arguments = '"' + $app + '"'
    $acceso.WorkingDirectory = $raiz
    $acceso.IconLocation = "$icono,0"
    $acceso.Description = 'Convierte documentos a Markdown y Markdown a Word, PDF, Excel o HTML'
    $acceso.Save()
    Write-Output "Creado: $ruta"
}
Write-Output "Lanzador usado: $lanzador"

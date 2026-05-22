#requires -Version 5.1
<#
.SYNOPSIS
    Compila ost2pst.py a un .exe portable con PyInstaller (Python 3.9).

.DESCRIPTION
    libpff-python solo publica wheel precompilado para Windows en cp39.
    Auto-genera version_info.txt desde __version__ en ost2pst.py para evitar
    drift. Prefiere el wheel local en vendor/ y solo recurre a PyPI si falta.
    Empaqueta el .py como recurso (para que --debug pueda mostrar el código)
    y publica SHA-256 del .exe junto al binario.

.PARAMETER Clean
    Elimina build/, dist/ y ost2pst.spec antes de compilar.

.PARAMETER Test
    Ejecuta pytest antes de compilar; aborta si falla.

.PARAMETER Upx
    Comprime el .exe con UPX si está en PATH (reduce ~50% el tamaño).
    Aviso: algunos antivirus marcan binarios UPX como sospechosos.

.EXAMPLE
    .\build.ps1
    .\build.ps1 -Clean -Test
    .\build.ps1 -Clean -Test -Upx
#>

[CmdletBinding()]
param(
    [switch]$Clean,
    [switch]$Test,
    [switch]$Upx
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot

function Read-Version([string]$pyFile) {
    $content = Get-Content -Raw $pyFile
    if ($content -match '__version__\s*=\s*"([^"]+)"') {
        return $matches[1]
    }
    throw "no se encontró __version__ en $pyFile"
}

function Write-VersionInfo([string]$version, [string]$outFile) {
    $parts = $version -split '\.'
    while ($parts.Count -lt 4) { $parts += '0' }
    $tuple = "($($parts[0]), $($parts[1]), $($parts[2]), $($parts[3]))"
    $verStr = ($parts -join '.')

    $content = @"
# Auto-generado por build.ps1 desde __version__. NO EDITAR A MANO.
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=$tuple,
    prodvers=$tuple,
    mask=0x3f, flags=0x0, OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0),
  ),
  kids=[
    StringFileInfo([
      StringTable('040904B0', [
        StringStruct('CompanyName', 'ost2pst project'),
        StringStruct('FileDescription',
                     'Procesador de ficheros Outlook OST '
                     '(exportacion a PST / extraccion a EML)'),
        StringStruct('FileVersion', '$verStr'),
        StringStruct('InternalName', 'ost2pst'),
        StringStruct('LegalCopyright', ''),
        StringStruct('OriginalFilename', 'ost2pst.exe'),
        StringStruct('ProductName', 'ost2pst'),
        StringStruct('ProductVersion', '$verStr'),
      ]),
    ]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])]),
  ],
)
"@
    Set-Content -Path $outFile -Value $content -Encoding UTF8
}

Push-Location $root
try {
    if ($Clean) {
        Write-Host "[build] limpiando artefactos previos..." -ForegroundColor Yellow
        Remove-Item -Recurse -Force build, dist, ost2pst.spec -ErrorAction SilentlyContinue
    }

    Write-Host "[build] verificando Python 3.9..." -ForegroundColor Cyan
    & py -3.9 --version
    if ($LASTEXITCODE -ne 0) {
        throw "Python 3.9 no encontrado. Instálalo con: winget install Python.Python.3.9"
    }

    $version = Read-Version (Join-Path $root 'ost2pst.py')
    Write-Host "[build] versión detectada: $version" -ForegroundColor Cyan
    Write-VersionInfo -version $version -outFile (Join-Path $root 'version_info.txt')

    Write-Host "[build] instalando dependencias (prefiere vendor/ local)..." -ForegroundColor Cyan
    $vendorArgs = @()
    if (Test-Path (Join-Path $root 'vendor')) {
        $vendorArgs = @('--find-links', 'vendor')
    }
    & py -3.9 -m pip install --quiet @vendorArgs -r requirements.txt
    if ($LASTEXITCODE -ne 0) { throw "fallo instalando dependencias" }

    if ($Test) {
        Write-Host "[build] ejecutando tests..." -ForegroundColor Cyan
        & py -3.9 -m pip install --quiet pytest
        & py -3.9 -m pytest tests/ -q
        if ($LASTEXITCODE -ne 0) { throw "tests fallaron - aborto el build" }
        Write-Host "[build] tests OK" -ForegroundColor Green
    }

    $piArgs = @(
        '-m', 'PyInstaller',
        '--onefile',
        '--console',
        '--name', 'ost2pst',
        '--clean',
        '--noconfirm',
        '--add-data', 'ost2pst.py;.',
        '--version-file', 'version_info.txt'
    )
    if (Test-Path (Join-Path $root 'ost2pst.ico')) {
        $piArgs += @('--icon', 'ost2pst.ico')
    }
    if ($Upx) {
        $upxPath = (Get-Command upx -ErrorAction SilentlyContinue)?.Source
        if ($upxPath) {
            $piArgs += @('--upx-dir', (Split-Path $upxPath))
            Write-Host "[build] UPX habilitado: $upxPath" -ForegroundColor Cyan
        } else {
            Write-Host "[build] UPX no encontrado en PATH, se ignora -Upx" -ForegroundColor Yellow
        }
    }
    $piArgs += 'ost2pst.py'

    Write-Host "[build] compilando con PyInstaller..." -ForegroundColor Cyan
    & py -3.9 @piArgs
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller devolvió error" }

    $exe = Join-Path $root 'dist\ost2pst.exe'
    if (-not (Test-Path $exe)) { throw "no se generó dist\ost2pst.exe" }

    Write-Host "[build] calculando SHA-256 del .exe..." -ForegroundColor Cyan
    $hash = (Get-FileHash $exe -Algorithm SHA256).Hash.ToLower()
    $hashFile = "$exe.sha256"
    Set-Content -Path $hashFile -Value "$hash  ost2pst.exe" -Encoding ASCII

    $size = '{0:N1} MB' -f ((Get-Item $exe).Length / 1MB)
    Write-Host ""
    Write-Host "[build] OK -> $exe  ($size)" -ForegroundColor Green
    Write-Host "[build] SHA-256: $hash" -ForegroundColor Green
    & $exe --version
}
finally {
    Pop-Location
}

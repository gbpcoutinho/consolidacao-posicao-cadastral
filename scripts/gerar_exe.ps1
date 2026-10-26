# Gera o ConsolidadorPC (PyInstaller --onedir), roda o fumo do .exe nos
# exemplos ficticios e compacta em ConsolidadorPC-<versao>.zip.
#
# Uso (na raiz do projeto, PowerShell):
#   .\scripts\gerar_exe.ps1
#   .\scripts\gerar_exe.ps1 -Dist C:\temp\dist -Build C:\temp\build   # fora do OneDrive
param(
    [string]$Dist = "dist",
    [string]$Build = "build"
)
$ErrorActionPreference = "Stop"
$raiz = Split-Path -Parent $PSScriptRoot
Set-Location $raiz
$py = Join-Path $raiz ".venv\Scripts\python.exe"

& $py -m pip install -q -r requirements.txt -r requirements-app.txt
if ($LASTEXITCODE -ne 0) { throw "Falha ao instalar as dependencias." }

& $py -m PyInstaller --noconfirm --clean --distpath $Dist --workpath $Build ConsolidadorPC.spec
if ($LASTEXITCODE -ne 0) { throw "PyInstaller falhou." }

$versao = (& $py -c "import app; print(app.VERSAO)").Trim()
$exe = Join-Path $Dist "ConsolidadorPC\ConsolidadorPC.exe"

# Fumo: o .exe processa os exemplos sem interface e tem de sair com codigo 0.
$saida = Join-Path ([System.IO.Path]::GetTempPath()) "fumo_ConsolidadorPC.xlsx"
Remove-Item $saida -ErrorAction SilentlyContinue
$argumentos = @(
    "--fumo",
    "`"$(Join-Path $raiz 'exemplos\entrada')`"",
    "`"$(Join-Path $raiz 'exemplos\COD_EMPRESAS_EXEMPLO.csv')`"",
    "`"$saida`""
)
$processo = Start-Process -FilePath $exe -ArgumentList $argumentos -Wait -PassThru
if ($processo.ExitCode -ne 0 -or -not (Test-Path $saida)) {
    throw "Fumo do executavel falhou (codigo $($processo.ExitCode)). Veja %APPDATA%\ConsolidadorPC\logs\app.log."
}
Write-Host "Fumo OK: $saida"

$zip = Join-Path $Dist "ConsolidadorPC-$versao.zip"
Compress-Archive -Path (Join-Path $Dist "ConsolidadorPC") -DestinationPath $zip -Force
Write-Host "Gerado: $zip"

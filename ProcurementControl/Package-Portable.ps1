[CmdletBinding()]
param(
    [switch]$Clean
)

$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSCommandPath
$projectFile = Join-Path $projectRoot 'src\ProcurementControl.App\ProcurementControl.App.csproj'
$releaseRoot = Join-Path $projectRoot 'release'
$packageRoot = Join-Path $releaseRoot 'ProcurementControl-Portable'
$archivePath = Join-Path $releaseRoot 'ProcurementControl-Portable-win-x64.zip'

if ((Test-Path -LiteralPath $packageRoot) -or (Test-Path -LiteralPath $archivePath)) {
    if (-not $Clean) {
        throw "Папка или архив релиза уже существуют. Для их пересоздания запустите: .\Package-Portable.ps1 -Clean"
    }

    Remove-Item -LiteralPath $packageRoot -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $archivePath -Force -ErrorAction SilentlyContinue
}

New-Item -ItemType Directory -Path $packageRoot -Force | Out-Null

# Self-contained: на компьютере получателя не требуется установленный .NET.
dotnet publish $projectFile `
    --configuration Release `
    --runtime win-x64 `
    --self-contained true `
    --output $packageRoot

if ($LASTEXITCODE -ne 0) {
    throw "Сборка завершилась с кодом $LASTEXITCODE."
}

# Не копируем пользовательские БД. Метки сохраняют пустые каталоги и в папке,
# и в ZIP-архиве. При первом старте DatabaseInitializer создаст чистую SQLite БД.
$emptyDirectories = @(
    'data\purchase_control\files',
    'data\notes',
    'data\components\files'
)
foreach ($relativePath in $emptyDirectories) {
    $directory = Join-Path $packageRoot $relativePath
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $directory '.keep') `
        -Value 'Эта папка создана для переносимой версии. Пользовательские данные не входят в поставку.' `
        -Encoding utf8
}

@'
Procurement Control — переносимая версия

Запустите ProcurementControl.App.exe.
Папка data поставляется пустой: при первом запуске приложение автоматически
создаст чистую локальную базу SQLite и остальные рабочие каталоги.

Не запускайте программу из ZIP: сначала распакуйте архив в обычную папку,
где у пользователя есть право записи (например, Документы).
'@ | Set-Content -LiteralPath (Join-Path $packageRoot 'README.txt') -Encoding utf8

Compress-Archive -Path (Join-Path $packageRoot '*') -DestinationPath $archivePath -CompressionLevel Optimal

Write-Host "Готово: $archivePath"

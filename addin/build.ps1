<#
  Сборка надстройки и MSI-установщика без Visual Studio.
  Требуется только .NET SDK 8+ (https://dotnet.microsoft.com/download) и доступ к nuget.org.

  Запуск:  powershell -ExecutionPolicy Bypass -File .\build.ps1
  Результат: .\dist\OutlookAutosave-<версия>.msi
#>
param(
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    throw "Не найден .NET SDK. Установите его: winget install Microsoft.DotNet.SDK.8"
}

Write-Host "==> Сборка надстройки" -ForegroundColor Cyan
dotnet build "$root\src\OutlookAutosave\OutlookAutosave.vbproj" -c $Configuration
if ($LASTEXITCODE -ne 0) { throw "Сборка надстройки завершилась с ошибкой." }

Write-Host "==> Сборка MSI" -ForegroundColor Cyan
dotnet build "$root\installer\OutlookAutosave.Installer.wixproj" -c $Configuration -p:Platform=x64
if ($LASTEXITCODE -ne 0) { throw "Сборка установщика завершилась с ошибкой." }

$dist = Join-Path $root "dist"
New-Item -ItemType Directory -Force -Path $dist | Out-Null
Get-ChildItem -Path "$root\installer\bin" -Recurse -Filter *.msi |
    Where-Object { $_.FullName -like "*\$Configuration\*" } |
    Copy-Item -Destination $dist -Force

Write-Host "==> Готово:" -ForegroundColor Green
Get-ChildItem -Path $dist -Filter *.msi | ForEach-Object { Write-Host "    $($_.FullName)" }

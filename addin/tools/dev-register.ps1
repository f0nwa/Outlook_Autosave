<#
  Регистрация собранной надстройки для отладки без MSI (только текущий пользователь, без прав администратора).
  Outlook должен быть закрыт.

  Запуск:  powershell -ExecutionPolicy Bypass -File .\tools\dev-register.ps1
  Отмена:  powershell -ExecutionPolicy Bypass -File .\tools\dev-register.ps1 -Unregister
#>
param(
    [string]$Configuration = "Release",
    [switch]$Unregister
)

$ErrorActionPreference = "Stop"

$progId = "OutlookAutosave.Connect"
$clsid = "{6F0E8E54-3B7C-4E0B-9C51-2B7A3C1D5E10}"
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$dll = Join-Path $root "src\OutlookAutosave\bin\$Configuration\net48\OutlookAutosave.dll"

$classRoots = @(
    "HKCU:\Software\Classes\CLSID\$clsid",
    "HKCU:\Software\Classes\Wow6432Node\CLSID\$clsid"
)
$progIdKey = "HKCU:\Software\Classes\$progId"
$addinKey = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$progId"

if ($Unregister) {
    foreach ($key in $classRoots + @($progIdKey, $addinKey)) {
        if (Test-Path $key) { Remove-Item -Path $key -Recurse -Force }
    }
    Write-Host "Регистрация удалена."
    return
}

if (-not (Test-Path $dll)) {
    throw "Не найден $dll. Сначала выполните build.ps1."
}

$assemblyName = [System.Reflection.AssemblyName]::GetAssemblyName($dll)
$version = $assemblyName.Version.ToString()
$codeBase = "file:///" + ($dll -replace "\\", "/")

function Set-Default([string]$path, [string]$value) {
    New-Item -Path $path -Force | Out-Null
    Set-ItemProperty -Path $path -Name "(default)" -Value $value
}

Set-Default $progIdKey $progId
Set-Default "$progIdKey\CLSID" $clsid

foreach ($classRoot in $classRoots) {
    Set-Default $classRoot $progId
    Set-Default "$classRoot\ProgId" $progId
    New-Item -Path "$classRoot\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}" -Force | Out-Null

    foreach ($inproc in @("$classRoot\InprocServer32", "$classRoot\InprocServer32\$version")) {
        New-Item -Path $inproc -Force | Out-Null
        Set-ItemProperty -Path $inproc -Name "Class" -Value $progId
        Set-ItemProperty -Path $inproc -Name "Assembly" -Value $assemblyName.FullName
        Set-ItemProperty -Path $inproc -Name "RuntimeVersion" -Value "v4.0.30319"
        Set-ItemProperty -Path $inproc -Name "CodeBase" -Value $codeBase
    }

    Set-ItemProperty -Path "$classRoot\InprocServer32" -Name "(default)" -Value "mscoree.dll"
    Set-ItemProperty -Path "$classRoot\InprocServer32" -Name "ThreadingModel" -Value "Both"
}

New-Item -Path $addinKey -Force | Out-Null
Set-ItemProperty -Path $addinKey -Name "FriendlyName" -Value "Outlook Autosave (dev)"
Set-ItemProperty -Path $addinKey -Name "Description" -Value "Outlook Autosave, отладочная регистрация"
New-ItemProperty -Path $addinKey -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null

Write-Host "Надстройка зарегистрирована: $dll"
Write-Host "Запустите Outlook. Вкладка 'Главная' -> группа 'Outlook Autosave'."

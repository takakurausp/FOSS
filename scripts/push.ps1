<#
.SYNOPSIS
    現在 active な環境（または指定環境）にローカルのソースを反映する。

.DESCRIPTION
    -Name を指定すれば switch-env.ps1 で切り替えてから push。
    指定しない場合は現在の .clasp.json で push。

.PARAMETER Name
    環境を指定して push したい場合に使用。

.EXAMPLE
    .\scripts\push.ps1
    # → 現在 active な環境に push

.EXAMPLE
    .\scripts\push.ps1 -Name test
    # → test 環境に切り替えてから push
#>
param([string]$Name = "")

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

if ($Name -ne "") {
    & (Join-Path $PSScriptRoot 'switch-env.ps1') -Name $Name
}

if (-not (Test-Path .\.clasp.json)) {
    Write-Host "No active .clasp.json." -ForegroundColor Red
    Write-Host "Use: .\scripts\switch-env.ps1 -Name <Name>" -ForegroundColor Yellow
    exit 1
}

# 現在の環境を確認
try {
    $cfg = Get-Content .\.clasp.json | ConvertFrom-Json
    $shortId = $cfg.scriptId.Substring(0, [Math]::Min(20, $cfg.scriptId.Length)) + '...'
    Write-Host "Pushing to: $shortId" -ForegroundColor Gray
} catch {}

clasp push --force
if ($LASTEXITCODE -ne 0) {
    Write-Host "clasp push failed." -ForegroundColor Red
    exit $LASTEXITCODE
}
Write-Host "Pushed successfully." -ForegroundColor Green

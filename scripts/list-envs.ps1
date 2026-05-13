<#
.SYNOPSIS
    保存されているデプロイ環境の一覧を表示し、現在 active な環境を示す。
#>

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$dDir = Join-Path $root 'deployments'

# 現在 active な scriptId を取得
$activeId = $null
if (Test-Path (Join-Path $root '.clasp.json')) {
    try {
        $cfg = Get-Content (Join-Path $root '.clasp.json') | ConvertFrom-Json
        $activeId = $cfg.scriptId
    } catch {}
}

Write-Host "Deployments:" -ForegroundColor Cyan
$deployments = Get-ChildItem $dDir -Filter '*.clasp.json' -ErrorAction SilentlyContinue
if (-not $deployments) {
    Write-Host "  (none — create one with: .\scripts\deploy.ps1 -Name <Name>)" -ForegroundColor Gray
    exit 0
}

$deployments | ForEach-Object {
    $envName = $_.BaseName.Replace('.clasp', '')
    try {
        $cfg = Get-Content $_.FullName | ConvertFrom-Json
        $shortId = $cfg.scriptId.Substring(0, [Math]::Min(20, $cfg.scriptId.Length)) + '...'
        $marker = if ($cfg.scriptId -eq $activeId) { '* (active)' } else { '  ' }
        Write-Host ("  {0}{1,-15} {2}" -f $marker, $envName, $shortId)
    } catch {
        Write-Host ("    {0,-15} (parse error)" -f $envName) -ForegroundColor Yellow
    }
}

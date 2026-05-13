<#
.SYNOPSIS
    環境（deployment）を切り替える。

.DESCRIPTION
    deployments/<Name>.clasp.json を .clasp.json にコピーして active にする。
    以降の clasp push / open はその環境に対して実行される。

.PARAMETER Name
    環境識別子（deploy.ps1 で作成したもの）。

.EXAMPLE
    .\scripts\switch-env.ps1 -Name test
    # → test 環境を active にする

.EXAMPLE
    .\scripts\switch-env.ps1 -Name prod
#>
param([Parameter(Mandatory=$true)][string]$Name)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$src  = Join-Path $root "deployments\$Name.clasp.json"

if (-not (Test-Path $src)) {
    Write-Host "Deployment '$Name' not found at: $src" -ForegroundColor Red
    Write-Host ""
    Write-Host "Available deployments:" -ForegroundColor Yellow
    $deployments = Get-ChildItem (Join-Path $root 'deployments') -Filter '*.clasp.json' -ErrorAction SilentlyContinue
    if ($deployments) {
        $deployments | ForEach-Object {
            $envName = $_.BaseName.Replace('.clasp', '')
            Write-Host "  - $envName" -ForegroundColor Gray
        }
    } else {
        Write-Host "  (none — create one with: .\scripts\deploy.ps1 -Name <Name>)" -ForegroundColor Gray
    }
    exit 1
}

# 切替前に現行 .clasp.json を退避（不慮の上書き防止）
if (Test-Path (Join-Path $root '.clasp.json')) {
    Copy-Item (Join-Path $root '.clasp.json') (Join-Path $root '.clasp.json.backup') -Force
}

Copy-Item $src (Join-Path $root '.clasp.json') -Force
Write-Host "Switched to deployment: $Name" -ForegroundColor Green

# scriptId の最初の数文字を表示して切替確認
try {
    $cfg = Get-Content (Join-Path $root '.clasp.json') | ConvertFrom-Json
    $shortId = $cfg.scriptId.Substring(0, [Math]::Min(20, $cfg.scriptId.Length)) + '...'
    Write-Host "  scriptId: $shortId" -ForegroundColor Gray
} catch {
    # parse 失敗は致命的でないので無視
}

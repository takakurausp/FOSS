<#
.SYNOPSIS
    新規 GAS 環境（Sheets + Bound Script）を作成し、ソースを bundle する。

.DESCRIPTION
    内部処理:
      1. 既存 .clasp.json を .clasp.json.backup に退避
      2. clasp create で新規 Spreadsheet + bound Apps Script を作成
      3. clasp push で全ソース（.js / .html）を反映
      4. 生成された .clasp.json を deployments/<Name>.clasp.json に保存
      5. ブラウザで script editor を開き、bootstrap 実行手順を案内

.PARAMETER Name
    環境識別子（例: test, prod, dev-koichi）。
    deployments/<Name>.clasp.json として保存される。

.PARAMETER Title
    Google Sheets のタイトル。省略時は "FOSS Journal - <Name>"。

.EXAMPLE
    .\scripts\deploy.ps1 -Name test
    # → "FOSS Journal - test" という名前の新規 Sheets を作成

.EXAMPLE
    .\scripts\deploy.ps1 -Name prod -Title "投稿査読システム本番"
#>
param(
    [Parameter(Mandatory=$true)][string]$Name,
    [string]$Title = ""
)

if ([string]::IsNullOrWhiteSpace($Title)) { $Title = "FOSS Journal - $Name" }

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

# clasp 存在チェック
$claspPath = Get-Command clasp -ErrorAction SilentlyContinue
if (-not $claspPath) {
    Write-Host "clasp not found. Install via: npm install -g @google/clasp" -ForegroundColor Red
    exit 1
}

# 既存 deployment との衝突チェック
$dest = Join-Path $root "deployments\$Name.clasp.json"
if (Test-Path $dest) {
    Write-Host "Deployment '$Name' already exists at: $dest" -ForegroundColor Yellow
    Write-Host "  - To activate it:  .\scripts\switch-env.ps1 -Name $Name" -ForegroundColor Yellow
    Write-Host "  - To recreate:     delete the file first, then re-run this script." -ForegroundColor Yellow
    exit 1
}

# 既存 .clasp.json を退避（現在使用中の環境を壊さないように）
if (Test-Path .\.clasp.json) {
    Move-Item .\.clasp.json .\.clasp.json.backup -Force
    Write-Host "Existing .clasp.json backed up to .clasp.json.backup" -ForegroundColor Gray
}

try {
    # 1. Sheets + Bound Script を新規作成
    Write-Host "[1/3] Creating new Spreadsheet + bound Apps Script..." -ForegroundColor Cyan
    clasp create --type sheets --title $Title --rootDir .
    if ($LASTEXITCODE -ne 0) { throw "clasp create failed (exit $LASTEXITCODE)" }

    # 2. ソース push
    Write-Host "[2/3] Pushing source files..." -ForegroundColor Cyan
    clasp push --force
    if ($LASTEXITCODE -ne 0) { throw "clasp push failed (exit $LASTEXITCODE)" }

    # 3. deployment 設定を保管
    New-Item -ItemType Directory -Force -Path .\deployments | Out-Null
    Copy-Item .\.clasp.json $dest
    Write-Host "[3/3] Deployment config saved to: $dest" -ForegroundColor Green

    # 4. bootstrap 実行手順を案内
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host " Deployment '$Name' created successfully" -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps (one-time, manual):" -ForegroundColor Yellow
    Write-Host "  1. The script editor will open in your browser." -ForegroundColor Yellow
    Write-Host "  2. Select function 'bootstrap' from the dropdown."  -ForegroundColor Yellow
    Write-Host "  3. Click Run, authorize when prompted." -ForegroundColor Yellow
    Write-Host "  4. Check Logger output for '=== bootstrap complete ==='." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "After bootstrap, switch back to this deployment any time with:" -ForegroundColor Gray
    Write-Host "  .\scripts\switch-env.ps1 -Name $Name" -ForegroundColor Gray
    Write-Host ""

    # clasp v3: open-script (Apps Script IDE) / open-container (Sheets)
    # v2 互換のため、まず open-script を試して失敗したら従来の open にフォールバック
    & clasp open-script 2>$null
    if ($LASTEXITCODE -ne 0) {
        & clasp open 2>$null
        if ($LASTEXITCODE -ne 0) {
            # 最終フォールバック: scriptId から直接 URL 構築
            try {
                $cfg = Get-Content .\.clasp.json | ConvertFrom-Json
                $url = "https://script.google.com/d/$($cfg.scriptId)/edit"
                Write-Host "Opening script editor: $url" -ForegroundColor Gray
                Start-Process $url
            } catch {
                Write-Host "Could not open script editor automatically." -ForegroundColor Yellow
                Write-Host "Manually open: clasp open-script" -ForegroundColor Yellow
            }
        }
    }
}
catch {
    Write-Host ""
    Write-Host "[ERROR] $_" -ForegroundColor Red
    # エラー時は退避した .clasp.json を戻す（ロールバック）
    if (Test-Path .\.clasp.json.backup) {
        Move-Item .\.clasp.json.backup .\.clasp.json -Force
        Write-Host "Restored previous .clasp.json from backup." -ForegroundColor Gray
    }
    exit 1
}

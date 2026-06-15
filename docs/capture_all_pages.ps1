# ========================================================
# 모든 페이지 자동 로그인 + 캡쳐 스크립트
# ========================================================
# LoginController 가 ?redirect= 파라미터를 지원하므로
# 각 페이지마다 redirect 값을 다르게 해서 한 번 호출로 그 페이지 캡쳐
# ========================================================

param(
    [string]$UserId = "lysabb",
    [string]$Password = "1234",
    [string]$BaseUrl = "https://alamis-pj.onrender.com",
    [string]$OutDir = "D:\testPro\docs\screenshots"
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

$edge = "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
if (-not (Test-Path $edge)) { $edge = "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe" }
if (-not (Test-Path $edge)) { throw "Edge 못 찾음" }

# 캡쳐할 페이지 목록
$pages = @(
    @{ name = "메인 (대시보드)"; redirect = "/main";           out = "02_main.png" },
    @{ name = "입고 목록";        redirect = "/inbound/list";    out = "03_inbound_list.png" },
    @{ name = "입고 등록 폼";      redirect = "/inbound/form";    out = "04_inbound_form.png" },
    @{ name = "출고 목록";        redirect = "/outbound/list";   out = "05_outbound_list.png" },
    @{ name = "출고 등록 폼";      redirect = "/outbound/form";   out = "06_outbound_form.png" },
    @{ name = "현재 재고 조회";    redirect = "/stock/list";      out = "07_stock_list.png" }
)

# 워밍업 한 번
Write-Host "Render 서버 워밍업..." -ForegroundColor Cyan
try {
    Invoke-WebRequest "$BaseUrl/login" -UseBasicParsing -TimeoutSec 120 | Out-Null
    Write-Host "  OK" -ForegroundColor Green
} catch {
    Write-Host "  실패하지만 계속: $($_.Exception.Message)" -ForegroundColor Yellow
}

$pageNum = 0
$total = $pages.Count
foreach ($p in $pages) {
    $pageNum++
    Write-Host ""
    Write-Host "[$pageNum/$total] $($p.name) 캡쳐 중..." -ForegroundColor Cyan
    Write-Host "  → redirect=$($p.redirect)" -ForegroundColor Gray

    # 임시 HTML (자동 로그인 + 원하는 페이지로 redirect)
    $temp = "$env:TEMP\autocap_$([Guid]::NewGuid().ToString('N')).html"
    $html = @"
<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>auto</title></head>
<body>
<form id="f" action="$BaseUrl/login" method="post">
<input type="hidden" name="userId" value="$UserId">
<input type="hidden" name="userPw" value="$Password">
<input type="hidden" name="redirect" value="$($p.redirect)">
</form>
<script>document.getElementById('f').submit();</script>
</body></html>
"@
    [System.IO.File]::WriteAllText($temp, $html, [System.Text.UTF8Encoding]::new($false))

    $outPath = Join-Path $OutDir $p.out
    Remove-Item $outPath -ErrorAction SilentlyContinue

    $userDataDir = "$env:TEMP\edge_cap_$([Guid]::NewGuid().ToString('N'))"
    $fileUrl = "file:///" + $temp.Replace('\','/')

    $args = @(
        "--headless=new",
        "--disable-gpu",
        "--hide-scrollbars",
        "--window-size=1600,1000",
        "--virtual-time-budget=15000",
        "--user-data-dir=$userDataDir",
        "--no-sandbox",
        "--screenshot=$outPath",
        $fileUrl
    )
    & $edge @args | Out-Null
    Start-Sleep -Seconds 2

    # 정리
    Remove-Item $temp -Force -ErrorAction SilentlyContinue
    if (Test-Path $userDataDir) {
        Remove-Item $userDataDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    # 결과
    if (Test-Path $outPath) {
        $size = [math]::Round((Get-Item $outPath).Length/1KB, 2)
        $bytes = [System.IO.File]::ReadAllBytes($outPath)[0..3]
        $isPng = $bytes[0] -eq 0x89 -and $bytes[1] -eq 0x50 -and $bytes[2] -eq 0x4E -and $bytes[3] -eq 0x47
        Write-Host "  ✓ $($p.out) ($size KB, PNG=$isPng)" -ForegroundColor Green
    } else {
        Write-Host "  ✗ 실패" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "=========================================" -ForegroundColor Green
Write-Host "  모든 페이지 캡쳐 완료!" -ForegroundColor Green
Write-Host "  폴더: $OutDir" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Green

# ========================================================
# 자동 로그인 + 메인화면 캡쳐 스크립트
# ========================================================
# 1) 임시 HTML 파일에 auto-submit form 작성
# 2) Edge headless 로 그 HTML 열기 → 자동 로그인 → 메인 페이지 리다이렉트
# 3) virtual-time-budget 으로 JS 차트 렌더링 기다린 후 스크린샷
# 4) 임시 HTML 파일 삭제 (비밀번호 정리)
# ========================================================

param(
    [string]$UserId = "admin",
    [string]$Password = "admin",
    [string]$BaseUrl = "https://alamis-pj.onrender.com",
    [string]$OutPath = "D:\testPro\docs\screenshots\02_main.png"
)

$ErrorActionPreference = 'Stop'

# 출력 디렉터리 보장
$outDir = Split-Path $OutPath -Parent
if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

# Edge 경로 찾기
$edge = "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
if (-not (Test-Path $edge)) {
    $edge = "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
}
if (-not (Test-Path $edge)) {
    throw "Edge 를 찾을 수 없습니다"
}

Write-Host "1) Render 서버 워밍업 (콜드 스타트 방지)..." -ForegroundColor Cyan
try {
    Invoke-WebRequest "$BaseUrl/login" -UseBasicParsing -TimeoutSec 90 | Out-Null
    Write-Host "   서버 응답 OK" -ForegroundColor Green
} catch {
    Write-Host "   서버 응답 실패하지만 계속 진행: $($_.Exception.Message)" -ForegroundColor Yellow
}

# 임시 HTML 파일 (auto-submit form)
$temp = "$env:TEMP\autologin_$([Guid]::NewGuid().ToString('N')).html"
$html = @"
<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>auto</title></head>
<body>
<form id="loginForm" action="$BaseUrl/login" method="post">
<input type="hidden" name="userId" value="$UserId">
<input type="hidden" name="userPw" value="$Password">
</form>
<script>document.getElementById('loginForm').submit();</script>
</body></html>
"@
[System.IO.File]::WriteAllText($temp, $html, [System.Text.UTF8Encoding]::new($false))
Write-Host "2) 임시 자동 로그인 HTML 생성: $temp" -ForegroundColor Cyan

# Edge headless 실행 (file:// → POST → redirect → /main)
$fileUrl = "file:///" + $temp.Replace('\','/')
Write-Host "3) Edge headless 로 자동 로그인 시도 중..." -ForegroundColor Cyan

# user-data-dir 도 임시로 (캐시 격리)
$userDataDir = "$env:TEMP\edge_capture_$([Guid]::NewGuid().ToString('N'))"

$args = @(
    "--headless=new",
    "--disable-gpu",
    "--hide-scrollbars",
    "--window-size=1600,1000",
    "--virtual-time-budget=15000",   # JS 차트 렌더링 충분히 대기
    "--user-data-dir=$userDataDir",
    "--no-sandbox",
    "--screenshot=$OutPath",
    $fileUrl
)
# stderr 리다이렉트 하지 않음 (PS 5.1 native command 이슈)
& $edge @args | Out-Null

Start-Sleep -Seconds 3

# 정리
Remove-Item $temp -Force -ErrorAction SilentlyContinue
if (Test-Path $userDataDir) {
    Remove-Item $userDataDir -Recurse -Force -ErrorAction SilentlyContinue
}

# 결과 확인
if (Test-Path $OutPath) {
    $size = (Get-Item $OutPath).Length
    $bytes = [System.IO.File]::ReadAllBytes($OutPath)[0..3]
    $isPng = $bytes[0] -eq 0x89 -and $bytes[1] -eq 0x50 -and $bytes[2] -eq 0x4E -and $bytes[3] -eq 0x47
    Write-Host ""
    Write-Host "=========================================" -ForegroundColor Green
    Write-Host "  캡쳐 완료!" -ForegroundColor Green
    Write-Host "  파일: $OutPath" -ForegroundColor Green
    Write-Host "  크기: $([math]::Round($size/1KB,2)) KB" -ForegroundColor Green
    Write-Host "  PNG 정상: $isPng" -ForegroundColor Green
    Write-Host "=========================================" -ForegroundColor Green
} else {
    throw "캡쳐 실패: 파일이 생성되지 않음"
}

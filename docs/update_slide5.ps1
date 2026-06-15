# ========================================================
# 슬라이드 5 (사이드바 메뉴) 교체 스크립트
# 메인 화면 캡쳐 + 사이드바에 번호 마크업 + 우측 설명
# ========================================================

$ErrorActionPreference = 'Stop'

# 색상
function RGB([int]$r, [int]$g, [int]$b) { return ($r + ($g * 256) + ($b * 65536)) }

$C_PRIMARY = RGB 49  130 246
$C_NAVY    = RGB 30  39  97
$C_SUCCESS = RGB 0   200 150
$C_DANGER  = RGB 240 68  82
$C_WARNING = RGB 255 153 0
$C_PURPLE  = RGB 138 43  226
$C_GRAY900 = RGB 25  31  40
$C_GRAY700 = RGB 78  89  104
$C_GRAY500 = RGB 140 150 165
$C_GRAY200 = RGB 226 230 236
$C_GRAY100 = RGB 242 244 246
$C_WHITE   = RGB 255 255 255

$FONT_KO = "맑은 고딕"

$pptPath = "D:\testPro\docs\Alamis_재고관리_사용설명서.pptx"
$imgPath = "D:\testPro\docs\screenshots\02_main.png"

if (-not (Test-Path $imgPath)) { throw "캡쳐 이미지가 없습니다: $imgPath" }
if (-not (Test-Path $pptPath)) { throw "PPT 파일이 없습니다: $pptPath" }

Write-Host "PowerPoint 열기..." -ForegroundColor Cyan
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pres = $ppt.Presentations.Open($pptPath)

$SW = $pres.PageSetup.SlideWidth
$SH = $pres.PageSetup.SlideHeight
Write-Host "슬라이드 크기: $SW x $SH" -ForegroundColor Gray

# 슬라이드 5의 모든 Shape 제거 (배경 제외)
$slide = $pres.Slides.Item(5)
Write-Host "슬라이드 5의 기존 Shape 제거 중..." -ForegroundColor Yellow
while ($slide.Shapes.Count -gt 0) {
    $slide.Shapes.Item(1).Delete()
}

# 배경색 흰색
$slide.Background.Fill.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$slide.Background.Fill.ForeColor.RGB = $C_WHITE
$slide.FollowMasterBackground = [Microsoft.Office.Core.MsoTriState]::msoFalse

# ===== 헬퍼 함수 =====
function Add-Rect {
    param($Slide, [single]$X, [single]$Y, [single]$W, [single]$H, [int]$FillColor, [bool]$NoLine = $true)
    $shape = $Slide.Shapes.AddShape(1, $X, $Y, $W, $H)
    $shape.Fill.ForeColor.RGB = $FillColor
    if ($NoLine) { $shape.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse }
    return $shape
}

function Add-RoundRect {
    param($Slide, [single]$X, [single]$Y, [single]$W, [single]$H, [int]$FillColor, [bool]$NoLine = $true, [single]$Radius = 0.1)
    $shape = $Slide.Shapes.AddShape(5, $X, $Y, $W, $H)
    $shape.Adjustments.Item(1) = $Radius
    $shape.Fill.ForeColor.RGB = $FillColor
    if ($NoLine) { $shape.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse }
    return $shape
}

function Add-TextBox {
    param($Slide, [single]$X, [single]$Y, [single]$W, [single]$H, [string]$Text,
          [int]$FontSize = 18, [int]$Color = 4868682, [bool]$Bold = $false,
          [string]$Align = "left", [string]$VAlign = "top", [string]$Font = "맑은 고딕")
    $tb = $Slide.Shapes.AddTextbox(1, $X, $Y, $W, $H)
    $tb.TextFrame.MarginLeft = 0; $tb.TextFrame.MarginRight = 0
    $tb.TextFrame.MarginTop = 0;  $tb.TextFrame.MarginBottom = 0
    $tb.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $Text
    $tr.Font.Name = $Font; $tr.Font.NameFarEast = $Font
    $tr.Font.Size = $FontSize; $tr.Font.Color.RGB = $Color
    if ($Bold) { $tr.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue }
    switch ($Align) {
        "left"   { $tr.ParagraphFormat.Alignment = 1 }
        "center" { $tr.ParagraphFormat.Alignment = 2 }
        "right"  { $tr.ParagraphFormat.Alignment = 3 }
    }
    switch ($VAlign) {
        "top"    { $tb.TextFrame.VerticalAnchor = 1 }
        "middle" { $tb.TextFrame.VerticalAnchor = 3 }
        "bottom" { $tb.TextFrame.VerticalAnchor = 4 }
    }
    return $tb
}

function Add-NumberCircle {
    # 번호 마크업용 동그라미 (강한 색상으로)
    param($Slide, [single]$X, [single]$Y, [single]$Size, [string]$Num, [int]$BgColor, [int]$FontSize = 0)
    $oval = $Slide.Shapes.AddShape(9, $X, $Y, $Size, $Size)
    $oval.Fill.ForeColor.RGB = $BgColor
    # 흰색 외곽선으로 강조
    $oval.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $oval.Line.ForeColor.RGB = $C_WHITE
    $oval.Line.Weight = 1.5
    # 그림자
    $oval.Shadow.Type = 1  # msoShadow1 (simple)
    $oval.Shadow.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

    $oval.TextFrame.MarginLeft = 0; $oval.TextFrame.MarginRight = 0
    $oval.TextFrame.MarginTop = 0;  $oval.TextFrame.MarginBottom = 0
    $oval.TextFrame.VerticalAnchor = 3
    $tr = $oval.TextFrame.TextRange
    $tr.Text = $Num
    $tr.Font.Name = $FONT_KO; $tr.Font.NameFarEast = $FONT_KO
    if ($FontSize -eq 0) { $FontSize = [int]($Size * 0.6) }
    $tr.Font.Size = $FontSize
    $tr.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $tr.Font.Color.RGB = $C_WHITE
    $tr.ParagraphFormat.Alignment = 2
    return $oval
}

# ===== 1) 챕터 헤더 =====
Write-Host "챕터 헤더 그리는 중..." -ForegroundColor Yellow

$badge = Add-RoundRect -Slide $slide -X 40 -Y 35 -W 65 -H 24 -FillColor $C_PRIMARY -Radius 0.5
$badge.TextFrame.VerticalAnchor = 3
$badge.TextFrame.MarginTop = 0; $badge.TextFrame.MarginBottom = 0
$bt = $badge.TextFrame.TextRange
$bt.Text = "02"
$bt.Font.Name = $FONT_KO; $bt.Font.NameFarEast = $FONT_KO
$bt.Font.Size = 11; $bt.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
$bt.Font.Color.RGB = $C_WHITE
$bt.ParagraphFormat.Alignment = 2

Add-TextBox -Slide $slide -X 115 -Y 36 -W ($SW - 150) -H 30 `
    -Text "사이드바 메뉴 — 4가지 기능" `
    -FontSize 20 -Color $C_GRAY900 -Bold $true -Align "left" -VAlign "middle" | Out-Null

Add-Rect -Slide $slide -X 40 -Y 75 -W ($SW - 80) -H 1 -FillColor $C_GRAY200 | Out-Null

# ===== 2) 메인 화면 캡쳐 이미지 삽입 =====
Write-Host "캡쳐 이미지 삽입 중..." -ForegroundColor Yellow

# 이미지 위치 & 크기
$imgX = 20
$imgY = 95
$imgW = 600   # 캡쳐가 1600x1000 비율 (16:10) → 600 x 375
$imgH = 375

# AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
$pic = $slide.Shapes.AddPicture($imgPath, 0, -1, $imgX, $imgY, $imgW, $imgH)

# 이미지 테두리 (얇은 회색)
$pic.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pic.Line.ForeColor.RGB = $C_GRAY200
$pic.Line.Weight = 0.75

# ===== 3) 사이드바에 번호 마크업 =====
Write-Host "사이드바 번호 마크업 그리는 중..." -ForegroundColor Yellow

# 캡쳐 좌표(1600x1000 기준) → 슬라이드 좌표 변환
# 스케일: 600/1600 = 0.375
$scale = $imgW / 1600.0

# 사이드바 메뉴의 캡쳐 내 Y좌표 (각 메뉴 텍스트 중심)
# 입고 관리:  y≈183
# 출고 관리:  y≈228
# 현재 재고:  y≈273
# 사용자 관리: y≈354
$menuPositions = @(
    @{ no = "1"; capY = 183; color = $C_PRIMARY },
    @{ no = "2"; capY = 228; color = $C_DANGER },
    @{ no = "3"; capY = 273; color = $C_SUCCESS }
)

# 번호 동그라미를 사이드바 메뉴 텍스트 바로 옆에 배치
# 사이드바 폭이 약 240px 캡쳐 → 슬라이드에서 90px
# 메뉴 텍스트 시작 캡쳐X≈30, 텍스트 끝 약 캡쳐X≈170
# 번호를 사이드바 우측(메뉴 텍스트 옆)에 배치
$numSize = 16
$numCapX = 195  # 사이드바 메뉴 텍스트 오른쪽 위치 (캡쳐 내)

foreach ($mp in $menuPositions) {
    $cx = $imgX + ($numCapX * $scale)
    $cy = $imgY + ($mp.capY * $scale) - ($numSize / 2)
    Add-NumberCircle -Slide $slide -X $cx -Y $cy -Size $numSize -Num $mp.no -BgColor $mp.color -FontSize 9 | Out-Null
}

# ===== 4) 우측 설명 패널 =====
Write-Host "우측 설명 패널 그리는 중..." -ForegroundColor Yellow

$rightX = 640
$rightW = 295

Add-TextBox -Slide $slide -X $rightX -Y 95 -W $rightW -H 25 `
    -Text "📋 메뉴별 기능 안내" `
    -FontSize 14 -Color $C_GRAY900 -Bold $true | Out-Null

$menus = @(
    @{ no = "1"; title = "입고 관리"; desc = "제품이 들어왔을 때 등록합니다.`r`n생산일자 · 포장단위 · 수량 등 입력`r`n제품 라벨(스티커) 인쇄도 여기서!"; color = $C_PRIMARY },
    @{ no = "2"; title = "출고 관리"; desc = "제품을 출고할 때 사용합니다.`r`nCOA(성적서) 코드는 필수 입력`r`n한 출고에 여러 입고 묶음 가능"; color = $C_DANGER },
    @{ no = "3"; title = "현재 재고 조회"; desc = "지금 남아있는 재고를 확인합니다.`r`n입고량 - 출고량 = 잔여재고`r`n엑셀(.xlsx) 다운로드 지원"; color = $C_SUCCESS }
)

$cardY = 130
$cardH = 105
$cardGap = 14
foreach ($m in $menus) {
    Add-RoundRect -Slide $slide -X $rightX -Y $cardY -W $rightW -H $cardH -FillColor $C_GRAY100 -Radius 0.08 | Out-Null
    # 좌측 색상 막대
    Add-Rect -Slide $slide -X $rightX -Y $cardY -W 4 -H $cardH -FillColor $m.color | Out-Null
    # 번호 (좌측 큰 동그라미)
    Add-NumberCircle -Slide $slide -X ($rightX + 14) -Y ($cardY + 16) -Size 22 -Num $m.no -BgColor $m.color -FontSize 12 | Out-Null
    # 제목
    Add-TextBox -Slide $slide -X ($rightX + 44) -Y ($cardY + 12) -W ($rightW - 55) -H 22 `
        -Text $m.title -FontSize 13 -Color $C_GRAY900 -Bold $true | Out-Null
    # 설명 (여러 줄)
    Add-TextBox -Slide $slide -X ($rightX + 44) -Y ($cardY + 35) -W ($rightW - 55) -H 65 `
        -Text $m.desc -FontSize 10 -Color $C_GRAY700 | Out-Null
    $cardY += $cardH + $cardGap
}

# ===== 5) 하단 설명 박스 (캡쳐 이미지 아래) =====
Add-RoundRect -Slide $slide -X 20 -Y 480 -W 600 -H 35 -FillColor (RGB 234 244 255) -Radius 0.2 | Out-Null
Add-TextBox -Slide $slide -X 35 -Y 489 -W 580 -H 22 `
    -Text "💡 화면 왼쪽 사이드바의 ①②③④ 번호 메뉴를 클릭하면 해당 기능으로 이동합니다." `
    -FontSize 10 -Color (RGB 30 90 180) -Bold $true | Out-Null

# ===== 6) 푸터 =====
Add-TextBox -Slide $slide -X 30 -Y ($SH - 25) -W 200 -H 18 `
    -Text "아라미스 펄 · 재고관리 시스템" `
    -FontSize 9 -Color $C_GRAY500 -Align "left" | Out-Null
Add-TextBox -Slide $slide -X ($SW - 100) -Y ($SH - 25) -W 70 -H 18 `
    -Text "5 / 12" `
    -FontSize 9 -Color $C_GRAY500 -Align "right" | Out-Null

# ===== 저장 =====
Write-Host ""
Write-Host "저장 중..." -ForegroundColor Cyan
$pres.Save()
$pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host "=========================================" -ForegroundColor Green
Write-Host "  슬라이드 5 업데이트 완료!" -ForegroundColor Green
Write-Host "  파일: $pptPath" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Green

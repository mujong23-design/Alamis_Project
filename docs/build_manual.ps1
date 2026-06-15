# ========================================================
# 아라미스 펄 재고관리 시스템 - 사용설명서 PPT 생성 스크립트
# ========================================================
# PowerPoint COM Automation 사용
# 16:9 와이드 슬라이드, 한국어, Toss 스타일 디자인
# ========================================================

$ErrorActionPreference = 'Stop'

# ---- 색상 정의 (Toss 디자인 시스템 매칭) ----
function RGB([int]$r, [int]$g, [int]$b) { return ($r + ($g * 256) + ($b * 65536)) }

$C_PRIMARY = RGB 49  130 246   # #3182F6 Toss Blue
$C_NAVY    = RGB 30  39  97    # #1E2761 다크 네이비
$C_SUCCESS = RGB 0   200 150   # #00C896 성공 그린
$C_DANGER  = RGB 240 68  82    # #F04452 위험 레드
$C_WARNING = RGB 255 153 0     # #FF9900 경고 오렌지
$C_GRAY900 = RGB 25  31  40    # 가장 진한 텍스트
$C_GRAY700 = RGB 78  89  104   # 본문 텍스트
$C_GRAY500 = RGB 140 150 165   # 보조 텍스트
$C_GRAY200 = RGB 226 230 236   # 구분선
$C_GRAY100 = RGB 242 244 246   # 카드 배경
$C_WHITE   = RGB 255 255 255

$FONT_KO = "맑은 고딕"

# ---- PowerPoint 시작 ----
Write-Host "PowerPoint 시작..." -ForegroundColor Cyan
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue  # 백그라운드 모드 지원 안되는 버전 대응

$pres = $ppt.Presentations.Add()
# 16:9 와이드 (13.333" × 7.5") → 사용자 지정 크기로 직접 설정
$pres.PageSetup.SlideWidth = 960
$pres.PageSetup.SlideHeight = 540
$SW = $pres.PageSetup.SlideWidth   # 960
$SH = $pres.PageSetup.SlideHeight  # 540

Write-Host "슬라이드 크기: $SW x $SH points" -ForegroundColor Gray

# ===== 헬퍼 함수 =====

function Add-BlankSlide {
    param($Presentation)
    # ppLayoutBlank = 12
    return $Presentation.Slides.Add($Presentation.Slides.Count + 1, 12)
}

function Set-SlideBackground {
    param($Slide, [int]$RGBColor)
    $Slide.Background.Fill.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $Slide.Background.Fill.ForeColor.RGB = $RGBColor
    $Slide.FollowMasterBackground = [Microsoft.Office.Core.MsoTriState]::msoFalse
}

function Add-Rect {
    # msoShapeRectangle = 1
    param($Slide, [single]$X, [single]$Y, [single]$W, [single]$H, [int]$FillColor, [bool]$NoLine = $true)
    $shape = $Slide.Shapes.AddShape(1, $X, $Y, $W, $H)
    $shape.Fill.ForeColor.RGB = $FillColor
    if ($NoLine) {
        $shape.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    }
    return $shape
}

function Add-RoundRect {
    # msoShapeRoundedRectangle = 5
    param($Slide, [single]$X, [single]$Y, [single]$W, [single]$H, [int]$FillColor, [bool]$NoLine = $true, [single]$Radius = 0.1)
    $shape = $Slide.Shapes.AddShape(5, $X, $Y, $W, $H)
    $shape.Adjustments.Item(1) = $Radius
    $shape.Fill.ForeColor.RGB = $FillColor
    if ($NoLine) {
        $shape.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    }
    return $shape
}

function Add-TextBox {
    param(
        $Slide,
        [single]$X, [single]$Y, [single]$W, [single]$H,
        [string]$Text,
        [int]$FontSize = 18,
        [int]$Color = 4868682,  # 기본 회색
        [bool]$Bold = $false,
        [string]$Align = "left",
        [string]$VAlign = "top",
        [string]$Font = "맑은 고딕"
    )
    # AddTextbox(Orientation, Left, Top, Width, Height)
    # msoTextOrientationHorizontal = 1
    $tb = $Slide.Shapes.AddTextbox(1, $X, $Y, $W, $H)
    $tb.TextFrame.MarginLeft = 0
    $tb.TextFrame.MarginRight = 0
    $tb.TextFrame.MarginTop = 0
    $tb.TextFrame.MarginBottom = 0
    $tb.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue

    $tr = $tb.TextFrame.TextRange
    $tr.Text = $Text
    $tr.Font.Name = $Font
    $tr.Font.NameFarEast = $Font
    $tr.Font.Size = $FontSize
    $tr.Font.Color.RGB = $Color
    if ($Bold) {
        $tr.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    }

    # 정렬
    switch ($Align) {
        "left"   { $tr.ParagraphFormat.Alignment = 1 }  # ppAlignLeft
        "center" { $tr.ParagraphFormat.Alignment = 2 }  # ppAlignCenter
        "right"  { $tr.ParagraphFormat.Alignment = 3 }  # ppAlignRight
    }

    # 수직 정렬
    # msoAnchorTop=1, msoAnchorMiddle=3, msoAnchorBottom=4
    switch ($VAlign) {
        "top"    { $tb.TextFrame.VerticalAnchor = 1 }
        "middle" { $tb.TextFrame.VerticalAnchor = 3 }
        "bottom" { $tb.TextFrame.VerticalAnchor = 4 }
    }

    return $tb
}

function Add-StepBadge {
    # 동그란 단계 번호 배지 (예: "1", "2", "3")
    param($Slide, [single]$X, [single]$Y, [single]$Size, [string]$Num, [int]$BgColor)
    # msoShapeOval = 9
    $oval = $Slide.Shapes.AddShape(9, $X, $Y, $Size, $Size)
    $oval.Fill.ForeColor.RGB = $BgColor
    $oval.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    $oval.TextFrame.MarginLeft = 0
    $oval.TextFrame.MarginRight = 0
    $oval.TextFrame.MarginTop = 0
    $oval.TextFrame.MarginBottom = 0
    $oval.TextFrame.VerticalAnchor = 3
    $tr = $oval.TextFrame.TextRange
    $tr.Text = $Num
    $tr.Font.Name = $FONT_KO
    $tr.Font.NameFarEast = $FONT_KO
    $tr.Font.Size = [int]($Size * 0.55)
    $tr.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $tr.Font.Color.RGB = $C_WHITE
    $tr.ParagraphFormat.Alignment = 2  # center
    return $oval
}

# 슬라이드 푸터 (페이지 번호 & 브랜드)
function Add-Footer {
    param($Slide, [int]$Page, [int]$Total)
    # 좌측: 브랜드명
    Add-TextBox -Slide $Slide -X 30 -Y ($SH - 25) -W 200 -H 18 `
        -Text "아라미스 펄 · 재고관리 시스템" `
        -FontSize 9 -Color $C_GRAY500 -Align "left" | Out-Null
    # 우측: 페이지
    Add-TextBox -Slide $Slide -X ($SW - 100) -Y ($SH - 25) -W 70 -H 18 `
        -Text "$Page / $Total" `
        -FontSize 9 -Color $C_GRAY500 -Align "right" | Out-Null
}

# 챕터 헤더 (모든 본문 슬라이드 상단 공통)
function Add-ChapterHeader {
    param($Slide, [string]$ChapterNo, [string]$ChapterTitle)
    # 좌측 배지
    $badge = Add-RoundRect -Slide $Slide -X 40 -Y 35 -W 65 -H 24 -FillColor $C_PRIMARY -Radius 0.5
    $badge.TextFrame.VerticalAnchor = 3
    $badge.TextFrame.MarginTop = 0
    $badge.TextFrame.MarginBottom = 0
    $bt = $badge.TextFrame.TextRange
    $bt.Text = $ChapterNo
    $bt.Font.Name = $FONT_KO
    $bt.Font.NameFarEast = $FONT_KO
    $bt.Font.Size = 11
    $bt.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $bt.Font.Color.RGB = $C_WHITE
    $bt.ParagraphFormat.Alignment = 2

    Add-TextBox -Slide $Slide -X 115 -Y 36 -W ($SW - 150) -H 30 `
        -Text $ChapterTitle `
        -FontSize 20 -Color $C_GRAY900 -Bold $true -Align "left" -VAlign "middle" | Out-Null

    # 얇은 구분선
    Add-Rect -Slide $Slide -X 40 -Y 75 -W ($SW - 80) -H 1 -FillColor $C_GRAY200 | Out-Null
}

# ============================================================
# 슬라이드 1 — 표지
# ============================================================
Write-Host "[1/12] 표지 슬라이드..." -ForegroundColor Yellow
$s1 = Add-BlankSlide $pres
Set-SlideBackground $s1 $C_NAVY

# 배경 장식 (큰 흐릿한 사각형)
$dec = Add-Rect -Slide $s1 -X 600 -Y -100 -W 500 -H 700 -FillColor $C_PRIMARY
$dec.Fill.Transparency = 0.85
$dec.Rotation = 25

# 상단 작은 라벨
Add-TextBox -Slide $s1 -X 70 -Y 130 -W 400 -H 25 `
    -Text "USER MANUAL · 2026" `
    -FontSize 12 -Color (RGB 130 160 220) -Bold $true -Align "left" | Out-Null

# 메인 타이틀
Add-TextBox -Slide $s1 -X 70 -Y 175 -W 850 -H 80 `
    -Text "아라미스 펄" `
    -FontSize 56 -Color $C_WHITE -Bold $true -Align "left" | Out-Null

Add-TextBox -Slide $s1 -X 70 -Y 250 -W 850 -H 60 `
    -Text "재고관리 시스템 사용설명서" `
    -FontSize 32 -Color $C_WHITE -Align "left" | Out-Null

# 부제목
Add-TextBox -Slide $s1 -X 70 -Y 340 -W 700 -H 25 `
    -Text "입고 · 출고 · 재고 · 라벨 인쇄까지 한 번에" `
    -FontSize 15 -Color (RGB 180 200 235) -Align "left" | Out-Null

# 하단 강조 박스
Add-Rect -Slide $s1 -X 70 -Y 430 -W 5 -H 55 -FillColor $C_SUCCESS | Out-Null
Add-TextBox -Slide $s1 -X 85 -Y 430 -W 500 -H 25 `
    -Text "https://alamis-pj.onrender.com" `
    -FontSize 13 -Color $C_WHITE -Bold $true -Align "left" | Out-Null
Add-TextBox -Slide $s1 -X 85 -Y 458 -W 500 -H 25 `
    -Text "언제 어디서나 인터넷만 있으면 접속 가능" `
    -FontSize 11 -Color (RGB 180 200 235) -Align "left" | Out-Null

# ============================================================
# 슬라이드 2 — 목차
# ============================================================
Write-Host "[2/12] 목차 슬라이드..." -ForegroundColor Yellow
$s2 = Add-BlankSlide $pres
Set-SlideBackground $s2 $C_WHITE

# 타이틀
Add-TextBox -Slide $s2 -X 40 -Y 50 -W 700 -H 50 `
    -Text "목차" `
    -FontSize 36 -Color $C_GRAY900 -Bold $true -Align "left" | Out-Null

Add-TextBox -Slide $s2 -X 40 -Y 100 -W 700 -H 25 `
    -Text "CONTENTS" `
    -FontSize 12 -Color $C_PRIMARY -Bold $true -Align "left" | Out-Null

# 목차 항목들 (2열 그리드)
$tocItems = @(
    @{ no = "01"; title = "접속 & 로그인"; desc = "사이트에 들어가고 로그인하기" },
    @{ no = "02"; title = "메인 화면 안내"; desc = "대시보드와 메뉴 구성 살펴보기" },
    @{ no = "03"; title = "입고 등록"; desc = "새로운 제품 입고 정보 등록" },
    @{ no = "04"; title = "라벨 인쇄"; desc = "제품 부착용 스티커 출력하기" },
    @{ no = "05"; title = "출고 등록"; desc = "출고 정보 + COA 등록" },
    @{ no = "06"; title = "재고 조회"; desc = "현재 재고 확인 + 엑셀 다운로드" }
)

$cardW = 420
$cardH = 90
$gapX = 20
$gapY = 18
$startX = 40
$startY = 155

for ($i = 0; $i -lt 6; $i++) {
    $col = $i % 2
    $row = [Math]::Floor($i / 2)
    $x = $startX + $col * ($cardW + $gapX)
    $y = $startY + $row * ($cardH + $gapY)

    # 카드 배경 (RoundRect)
    Add-RoundRect -Slide $s2 -X $x -Y $y -W $cardW -H $cardH -FillColor $C_GRAY100 -Radius 0.08 | Out-Null

    # 번호 (파란색 크게)
    Add-TextBox -Slide $s2 -X ($x + 18) -Y ($y + 18) -W 60 -H 50 `
        -Text $tocItems[$i].no `
        -FontSize 32 -Color $C_PRIMARY -Bold $true -Align "left" | Out-Null

    # 타이틀
    Add-TextBox -Slide $s2 -X ($x + 90) -Y ($y + 20) -W 310 -H 28 `
        -Text $tocItems[$i].title `
        -FontSize 16 -Color $C_GRAY900 -Bold $true -Align "left" | Out-Null

    # 설명
    Add-TextBox -Slide $s2 -X ($x + 90) -Y ($y + 48) -W 310 -H 22 `
        -Text $tocItems[$i].desc `
        -FontSize 11 -Color $C_GRAY500 -Align "left" | Out-Null
}

Add-Footer $s2 2 12

# ============================================================
# 슬라이드 3 — 챕터 1: 접속 & 로그인
# ============================================================
Write-Host "[3/12] 챕터1: 접속 & 로그인..." -ForegroundColor Yellow
$s3 = Add-BlankSlide $pres
Set-SlideBackground $s3 $C_WHITE
Add-ChapterHeader $s3 "01" "접속 & 로그인"

# 좌측 STEP 1
Add-StepBadge -Slide $s3 -X 50 -Y 110 -Size 38 -Num "1" -BgColor $C_PRIMARY | Out-Null
Add-TextBox -Slide $s3 -X 100 -Y 115 -W 400 -H 25 `
    -Text "인터넷 브라우저에서 주소 입력" `
    -FontSize 17 -Color $C_GRAY900 -Bold $true -Align "left" | Out-Null
Add-TextBox -Slide $s3 -X 100 -Y 142 -W 500 -H 22 `
    -Text "크롬, 엣지, 사파리 등 아무 브라우저나 사용 가능합니다." `
    -FontSize 12 -Color $C_GRAY700 -Align "left" | Out-Null

# URL 박스
Add-RoundRect -Slide $s3 -X 100 -Y 175 -W 440 -H 55 -FillColor $C_NAVY -Radius 0.15 | Out-Null
Add-TextBox -Slide $s3 -X 115 -Y 183 -W 410 -H 18 `
    -Text "🌐 주소창에 입력" `
    -FontSize 10 -Color (RGB 180 200 235) -Align "left" | Out-Null
Add-TextBox -Slide $s3 -X 115 -Y 202 -W 410 -H 23 `
    -Text "https://alamis-pj.onrender.com" `
    -FontSize 15 -Color $C_WHITE -Bold $true -Align "left" | Out-Null

# STEP 2
Add-StepBadge -Slide $s3 -X 50 -Y 260 -Size 38 -Num "2" -BgColor $C_PRIMARY | Out-Null
Add-TextBox -Slide $s3 -X 100 -Y 265 -W 400 -H 25 `
    -Text "아이디와 비밀번호 입력" `
    -FontSize 17 -Color $C_GRAY900 -Bold $true -Align "left" | Out-Null
Add-TextBox -Slide $s3 -X 100 -Y 292 -W 500 -H 22 `
    -Text "관리자에게 발급받은 계정 정보를 입력하세요." `
    -FontSize 12 -Color $C_GRAY700 -Align "left" | Out-Null

# STEP 3
Add-StepBadge -Slide $s3 -X 50 -Y 340 -Size 38 -Num "3" -BgColor $C_PRIMARY | Out-Null
Add-TextBox -Slide $s3 -X 100 -Y 345 -W 400 -H 25 `
    -Text "[로그인] 버튼 클릭" `
    -FontSize 17 -Color $C_GRAY900 -Bold $true -Align "left" | Out-Null
Add-TextBox -Slide $s3 -X 100 -Y 372 -W 500 -H 22 `
    -Text "성공하면 대시보드(메인 화면)로 이동합니다." `
    -FontSize 12 -Color $C_GRAY700 -Align "left" | Out-Null

# 우측 TIP 박스
Add-RoundRect -Slide $s3 -X 610 -Y 110 -W 305 -H 320 -FillColor $C_GRAY100 -Radius 0.05 | Out-Null
Add-Rect -Slide $s3 -X 610 -Y 110 -W 4 -H 320 -FillColor $C_WARNING | Out-Null

Add-TextBox -Slide $s3 -X 630 -Y 125 -W 280 -H 25 `
    -Text "💡 알아두면 좋은 점" `
    -FontSize 13 -Color $C_WARNING -Bold $true -Align "left" | Out-Null

Add-TextBox -Slide $s3 -X 630 -Y 160 -W 280 -H 250 `
    -Text "• 처음 접속하면 약 30초~1분 정도 로딩이 길 수 있습니다.`r`n  (서버가 깨어나는 시간)`r`n`r`n• 두 번째부터는 1~2초 안에 빠르게 뜹니다.`r`n`r`n• 즐겨찾기에 추가해 두면 매번 입력하지 않아도 됩니다.`r`n`r`n• 비밀번호를 잊으셨다면 관리자에게 문의하세요." `
    -FontSize 11 -Color $C_GRAY700 -Align "left" | Out-Null

Add-Footer $s3 3 12

# ============================================================
# 슬라이드 4 — 챕터 2: 메인 화면 안내
# ============================================================
Write-Host "[4/12] 챕터2: 메인 화면..." -ForegroundColor Yellow
$s4 = Add-BlankSlide $pres
Set-SlideBackground $s4 $C_WHITE
Add-ChapterHeader $s4 "02" "메인 화면 안내 (대시보드)"

# 상단 4개 통계 카드
$stats = @(
    @{ label = "전체 입고건수"; val = "1,234"; unit = "건"; color = $C_GRAY900 },
    @{ label = "현재 재고";     val = "892";   unit = "포장"; color = $C_PRIMARY },
    @{ label = "금일 입고";     val = "12";    unit = "건"; color = $C_SUCCESS },
    @{ label = "금일 출고";     val = "8";     unit = "건"; color = $C_DANGER }
)

$cardW2 = 200
$cardH2 = 85
$gapX2 = 15
$startX2 = 40
$startY2 = 105

for ($i = 0; $i -lt 4; $i++) {
    $x = $startX2 + $i * ($cardW2 + $gapX2)
    $y = $startY2

    Add-RoundRect -Slide $s4 -X $x -Y $y -W $cardW2 -H $cardH2 -FillColor $C_WHITE -Radius 0.08 | Out-Null
    # 카드 테두리 효과를 위한 얇은 라인
    Add-Rect -Slide $s4 -X $x -Y ($y + $cardH2 - 1) -W $cardW2 -H 1 -FillColor $C_GRAY200 | Out-Null

    Add-TextBox -Slide $s4 -X ($x + 15) -Y ($y + 15) -W ($cardW2 - 30) -H 18 `
        -Text $stats[$i].label `
        -FontSize 11 -Color $C_GRAY500 -Align "left" | Out-Null

    Add-TextBox -Slide $s4 -X ($x + 15) -Y ($y + 35) -W ($cardW2 - 30) -H 42 `
        -Text "$($stats[$i].val) $($stats[$i].unit)" `
        -FontSize 24 -Color $stats[$i].color -Bold $true -Align "left" | Out-Null
}

# 카드 설명
Add-TextBox -Slide $s4 -X 40 -Y 205 -W 880 -H 25 `
    -Text "📊 화면 상단의 4개 카드 — 한눈에 보는 핵심 지표" `
    -FontSize 14 -Color $C_GRAY900 -Bold $true -Align "left" | Out-Null

# 차트 영역 (가짜 시각화)
Add-RoundRect -Slide $s4 -X 40 -Y 245 -W 430 -H 200 -FillColor $C_WHITE -Radius 0.05 | Out-Null
Add-Rect -Slide $s4 -X 40 -Y 444 -W 430 -H 1 -FillColor $C_GRAY200 | Out-Null
Add-TextBox -Slide $s4 -X 55 -Y 260 -W 400 -H 22 `
    -Text "📈 월별 입출고 추이 (최근 6개월)" `
    -FontSize 12 -Color $C_GRAY900 -Bold $true -Align "left" | Out-Null

# 간단한 라인 차트 흉내 (사각형 막대로)
$lineY = 380
$months = @("Dec","Jan","Feb","Mar","Apr","May")
$inVals = @(60, 90, 75, 110, 85, 130)
$outVals = @(40, 70, 55, 80, 100, 95)
$barW = 18
$baseX = 90
$step = 60
for ($i = 0; $i -lt 6; $i++) {
    $x = $baseX + $i * $step
    Add-Rect -Slide $s4 -X $x -Y ($lineY - $inVals[$i] * 0.6) -W $barW -H ($inVals[$i] * 0.6) -FillColor $C_PRIMARY | Out-Null
    Add-Rect -Slide $s4 -X ($x + $barW + 2) -Y ($lineY - $outVals[$i] * 0.6) -W $barW -H ($outVals[$i] * 0.6) -FillColor $C_DANGER | Out-Null
    Add-TextBox -Slide $s4 -X ($x - 5) -Y ($lineY + 5) -W 50 -H 15 `
        -Text $months[$i] -FontSize 9 -Color $C_GRAY500 -Align "center" | Out-Null
}
# 범례
Add-Rect -Slide $s4 -X 340 -Y 263 -W 12 -H 12 -FillColor $C_PRIMARY | Out-Null
Add-TextBox -Slide $s4 -X 355 -Y 262 -W 50 -H 15 -Text "입고" -FontSize 10 -Color $C_GRAY700 | Out-Null
Add-Rect -Slide $s4 -X 395 -Y 263 -W 12 -H 12 -FillColor $C_DANGER | Out-Null
Add-TextBox -Slide $s4 -X 410 -Y 262 -W 50 -H 15 -Text "출고" -FontSize 10 -Color $C_GRAY700 | Out-Null

# 우측 차트
Add-RoundRect -Slide $s4 -X 490 -Y 245 -W 430 -H 200 -FillColor $C_WHITE -Radius 0.05 | Out-Null
Add-Rect -Slide $s4 -X 490 -Y 444 -W 430 -H 1 -FillColor $C_GRAY200 | Out-Null
Add-TextBox -Slide $s4 -X 505 -Y 260 -W 400 -H 22 `
    -Text "📦 제품별 현재 재고 TOP 10" `
    -FontSize 12 -Color $C_GRAY900 -Bold $true -Align "left" | Out-Null

# 가짜 막대 차트
$topVals = @(180, 150, 130, 110, 95, 80, 70, 55, 45, 30)
$baseX2 = 510
$step2 = 39
for ($i = 0; $i -lt 10; $i++) {
    $x = $baseX2 + $i * $step2
    Add-Rect -Slide $s4 -X $x -Y ($lineY - $topVals[$i] * 0.55) -W 25 -H ($topVals[$i] * 0.55) -FillColor $C_SUCCESS | Out-Null
}

Add-Footer $s4 4 12

# ============================================================
# 슬라이드 5 — 사이드바 메뉴 소개
# ============================================================
Write-Host "[5/12] 사이드바 메뉴..." -ForegroundColor Yellow
$s5 = Add-BlankSlide $pres
Set-SlideBackground $s5 $C_WHITE
Add-ChapterHeader $s5 "02" "사이드바 메뉴 — 4가지 기능"

# 메뉴 설명 4개 (2x2 그리드)
$menus = @(
    @{ icon = "📥"; title = "입고 관리"; desc = "새로운 제품 입고 등록`r`n입고 목록 조회`r`n제품 라벨 인쇄"; color = $C_PRIMARY },
    @{ icon = "📤"; title = "출고 관리"; desc = "출고 등록 (COA 필수)`r`n출고 목록 조회`r`n출고 디테일 확인"; color = $C_DANGER },
    @{ icon = "📊"; title = "현재 재고 조회"; desc = "제품별 잔여 재고 확인`r`n조건별 필터링`r`n엑셀(.xlsx) 다운로드"; color = $C_SUCCESS },
    @{ icon = "👤"; title = "사용자 관리"; desc = "사용자 추가/수정/삭제`r`n비밀번호 변경`r`n(관리자 전용)"; color = $C_WARNING }
)

$mw = 420
$mh = 145
$mgapX = 20
$mgapY = 15
$msx = 40
$msy = 100

for ($i = 0; $i -lt 4; $i++) {
    $col = $i % 2
    $row = [Math]::Floor($i / 2)
    $x = $msx + $col * ($mw + $mgapX)
    $y = $msy + $row * ($mh + $mgapY)

    Add-RoundRect -Slide $s5 -X $x -Y $y -W $mw -H $mh -FillColor $C_GRAY100 -Radius 0.06 | Out-Null
    # 좌측 색상 막대
    Add-Rect -Slide $s5 -X $x -Y $y -W 5 -H $mh -FillColor $menus[$i].color | Out-Null

    # 아이콘 + 제목
    Add-TextBox -Slide $s5 -X ($x + 22) -Y ($y + 18) -W 50 -H 35 `
        -Text $menus[$i].icon -FontSize 24 -Align "left" | Out-Null

    Add-TextBox -Slide $s5 -X ($x + 70) -Y ($y + 22) -W 320 -H 30 `
        -Text $menus[$i].title `
        -FontSize 18 -Color $C_GRAY900 -Bold $true -Align "left" | Out-Null

    # 설명
    Add-TextBox -Slide $s5 -X ($x + 22) -Y ($y + 62) -W ($mw - 40) -H 80 `
        -Text $menus[$i].desc `
        -FontSize 11 -Color $C_GRAY700 -Align "left" | Out-Null
}

Add-Footer $s5 5 12

# ============================================================
# 슬라이드 6 — 챕터 3: 입고 등록
# ============================================================
Write-Host "[6/12] 챕터3: 입고 등록..." -ForegroundColor Yellow
$s6 = Add-BlankSlide $pres
Set-SlideBackground $s6 $C_WHITE
Add-ChapterHeader $s6 "03" "입고 등록 방법"

# 좌측: 단계
$steps = @(
    @{ no = "1"; title = "입고 관리 메뉴 클릭"; desc = "사이드바에서 [📥 입고 관리] 클릭 → 입고 목록 화면 진입" },
    @{ no = "2"; title = "[+ 입고 등록] 버튼 클릭"; desc = "화면 우측 상단의 파란색 버튼을 누릅니다" },
    @{ no = "3"; title = "정보 입력"; desc = "생산일자 · 포장단위(kg) · 포장수량 · 크기(μm) · 비고 입력" },
    @{ no = "4"; title = "[저장] 버튼 클릭"; desc = "입고번호는 자동으로 생성됩니다 (P-YYYYMMDD-NNN)" }
)

$sy = 100
foreach ($st in $steps) {
    Add-StepBadge -Slide $s6 -X 50 -Y $sy -Size 34 -Num $st.no -BgColor $C_PRIMARY | Out-Null
    Add-TextBox -Slide $s6 -X 95 -Y ($sy + 2) -W 460 -H 25 `
        -Text $st.title `
        -FontSize 14 -Color $C_GRAY900 -Bold $true -Align "left" | Out-Null
    Add-TextBox -Slide $s6 -X 95 -Y ($sy + 26) -W 460 -H 22 `
        -Text $st.desc `
        -FontSize 11 -Color $C_GRAY700 -Align "left" | Out-Null
    $sy += 70
}

# 우측: 입력 항목 표
Add-RoundRect -Slide $s6 -X 590 -Y 100 -W 340 -H 340 -FillColor $C_GRAY100 -Radius 0.05 | Out-Null
Add-TextBox -Slide $s6 -X 610 -Y 115 -W 300 -H 25 `
    -Text "📝 입력 항목 안내" `
    -FontSize 13 -Color $C_GRAY900 -Bold $true -Align "left" | Out-Null

$fields = @(
    @{ label = "생산일자"; req = "필수"; desc = "제품이 생산된 날짜" },
    @{ label = "포장단위(kg)"; req = "필수"; desc = "한 포장의 무게 (소수 가능)" },
    @{ label = "포장수량"; req = "필수"; desc = "총 몇 포장인지 (정수)" },
    @{ label = "크기(μm)"; req = "선택"; desc = "제품 입자 크기" },
    @{ label = "비고"; req = "선택"; desc = "기타 메모 사항" }
)

$fy = 155
foreach ($f in $fields) {
    Add-TextBox -Slide $s6 -X 610 -Y $fy -W 130 -H 22 `
        -Text $f.label `
        -FontSize 11 -Color $C_GRAY900 -Bold $true -Align "left" | Out-Null
    # 필수/선택 뱃지
    $bcolor = if ($f.req -eq "필수") { $C_DANGER } else { $C_GRAY500 }
    $badge = Add-RoundRect -Slide $s6 -X 745 -Y ($fy + 2) -W 32 -H 16 -FillColor $bcolor -Radius 0.4
    $badge.TextFrame.VerticalAnchor = 3
    $badge.TextFrame.MarginTop = 0
    $badge.TextFrame.MarginBottom = 0
    $bt = $badge.TextFrame.TextRange
    $bt.Text = $f.req
    $bt.Font.Name = $FONT_KO
    $bt.Font.NameFarEast = $FONT_KO
    $bt.Font.Size = 8
    $bt.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $bt.Font.Color.RGB = $C_WHITE
    $bt.ParagraphFormat.Alignment = 2

    Add-TextBox -Slide $s6 -X 610 -Y ($fy + 22) -W 300 -H 20 `
        -Text $f.desc `
        -FontSize 10 -Color $C_GRAY500 -Align "left" | Out-Null
    $fy += 50
}

Add-Footer $s6 6 12

# ============================================================
# 슬라이드 7 — 챕터 4: 라벨 인쇄
# ============================================================
Write-Host "[7/12] 챕터4: 라벨 인쇄..." -ForegroundColor Yellow
$s7 = Add-BlankSlide $pres
Set-SlideBackground $s7 $C_WHITE
Add-ChapterHeader $s7 "04" "제품 라벨(스티커) 인쇄"

# 좌측 단계
$lsteps = @(
    @{ no = "1"; title = "입고 목록 화면으로 이동"; desc = "[📥 입고 관리] 메뉴를 클릭합니다" },
    @{ no = "2"; title = "인쇄할 항목 체크박스 선택"; desc = "각 행 왼쪽의 □ 를 클릭하여 여러 개 선택 가능" },
    @{ no = "3"; title = "[라벨 인쇄] 버튼 클릭"; desc = "이미 인쇄한 항목이면 재인쇄 확인 창이 뜹니다" },
    @{ no = "4"; title = "프린터로 출력"; desc = "A4 한 장에 라벨 10개 (2x5 배열)로 인쇄됩니다" }
)

$sy = 100
foreach ($st in $lsteps) {
    Add-StepBadge -Slide $s7 -X 50 -Y $sy -Size 34 -Num $st.no -BgColor $C_SUCCESS | Out-Null
    Add-TextBox -Slide $s7 -X 95 -Y ($sy + 2) -W 430 -H 25 `
        -Text $st.title `
        -FontSize 14 -Color $C_GRAY900 -Bold $true -Align "left" | Out-Null
    Add-TextBox -Slide $s7 -X 95 -Y ($sy + 26) -W 430 -H 22 `
        -Text $st.desc `
        -FontSize 11 -Color $C_GRAY700 -Align "left" | Out-Null
    $sy += 70
}

# 우측: 라벨 미리보기
Add-TextBox -Slide $s7 -X 560 -Y 100 -W 360 -H 22 `
    -Text "🏷️ 라벨 인쇄 미리보기" `
    -FontSize 13 -Color $C_GRAY900 -Bold $true -Align "left" | Out-Null

# A4 페이지 흉내
Add-RoundRect -Slide $s7 -X 560 -Y 128 -W 360 -H 320 -FillColor $C_WHITE -Radius 0.02 | Out-Null
Add-Rect -Slide $s7 -X 560 -Y 447 -W 360 -H 1 -FillColor $C_GRAY200 | Out-Null
Add-Rect -Slide $s7 -X 919 -Y 128 -W 1 -H 320 -FillColor $C_GRAY200 | Out-Null
Add-Rect -Slide $s7 -X 560 -Y 128 -W 1 -H 320 -FillColor $C_GRAY200 | Out-Null
Add-Rect -Slide $s7 -X 560 -Y 128 -W 360 -H 1 -FillColor $C_GRAY200 | Out-Null

# 라벨 2x5 = 10개
$lw = 165
$lh = 56
$lstartX = 570
$lstartY = 142
for ($r = 0; $r -lt 5; $r++) {
    for ($c = 0; $c -lt 2; $c++) {
        $lx = $lstartX + $c * ($lw + 10)
        $ly = $lstartY + $r * ($lh + 4)
        Add-Rect -Slide $s7 -X $lx -Y $ly -W $lw -H $lh -FillColor (RGB 250 251 252) | Out-Null
        # 라벨 내용 흉내
        Add-TextBox -Slide $s7 -X ($lx + 6) -Y ($ly + 5) -W ($lw - 12) -H 14 `
            -Text "아라미스 펄" -FontSize 7 -Color $C_PRIMARY -Bold $true | Out-Null
        Add-TextBox -Slide $s7 -X ($lx + 6) -Y ($ly + 19) -W ($lw - 12) -H 14 `
            -Text "P-20260528-001" -FontSize 8 -Color $C_GRAY900 -Bold $true | Out-Null
        Add-TextBox -Slide $s7 -X ($lx + 6) -Y ($ly + 35) -W ($lw - 12) -H 14 `
            -Text "20kg · 5.300μm" -FontSize 7 -Color $C_GRAY500 | Out-Null
    }
}

# 하단 안내 박스
Add-RoundRect -Slide $s7 -X 40 -Y 395 -W 500 -H 55 -FillColor (RGB 255 249 230) -Radius 0.1 | Out-Null
Add-TextBox -Slide $s7 -X 55 -Y 405 -W 485 -H 20 `
    -Text "💡 인쇄가 끝나면 '출력 여부' 가 자동으로 [완료] 로 바뀝니다" `
    -FontSize 11 -Color (RGB 153 102 0) -Bold $true | Out-Null
Add-TextBox -Slide $s7 -X 55 -Y 425 -W 485 -H 20 `
    -Text "한 번 더 인쇄하려고 하면 '이미 출력한 이력이 있습니다' 확인창이 뜹니다." `
    -FontSize 10 -Color (RGB 153 102 0) | Out-Null

Add-Footer $s7 7 12

# ============================================================
# 슬라이드 8 — 챕터 5: 출고 등록
# ============================================================
Write-Host "[8/12] 챕터5: 출고 등록..." -ForegroundColor Yellow
$s8 = Add-BlankSlide $pres
Set-SlideBackground $s8 $C_WHITE
Add-ChapterHeader $s8 "05" "출고 등록 방법"

# 좌측 단계
$osteps = @(
    @{ no = "1"; title = "출고 관리 → [+ 출고 등록]"; desc = "사이드바 [📤 출고 관리] → 우측 상단 버튼" },
    @{ no = "2"; title = "출고 정보 입력"; desc = "출고일자 · COA(필수) · 출고장소 · 비고" },
    @{ no = "3"; title = "출고할 입고 제품 선택"; desc = "[+ 제품 추가] 클릭 → 입고 목록에서 선택 후 수량 입력" },
    @{ no = "4"; title = "여러 입고 묶음 가능"; desc = "한 출고건에 여러 입고 제품을 같이 등록할 수 있습니다" },
    @{ no = "5"; title = "[저장] 클릭"; desc = "출고번호 자동 생성 (O-YYYYMMDD-NNN) · 재고 자동 차감" }
)

$sy = 95
foreach ($st in $osteps) {
    Add-StepBadge -Slide $s8 -X 50 -Y $sy -Size 32 -Num $st.no -BgColor $C_DANGER | Out-Null
    Add-TextBox -Slide $s8 -X 92 -Y ($sy + 2) -W 460 -H 22 `
        -Text $st.title `
        -FontSize 13 -Color $C_GRAY900 -Bold $true -Align "left" | Out-Null
    Add-TextBox -Slide $s8 -X 92 -Y ($sy + 23) -W 460 -H 20 `
        -Text $st.desc `
        -FontSize 11 -Color $C_GRAY700 -Align "left" | Out-Null
    $sy += 62
}

# 우측: COA 강조 박스
Add-RoundRect -Slide $s8 -X 590 -Y 95 -W 340 -H 160 -FillColor (RGB 254 235 238) -Radius 0.06 | Out-Null
Add-Rect -Slide $s8 -X 590 -Y 95 -W 4 -H 160 -FillColor $C_DANGER | Out-Null
Add-TextBox -Slide $s8 -X 610 -Y 110 -W 310 -H 25 `
    -Text "⚠️ COA 는 반드시 입력!" `
    -FontSize 14 -Color $C_DANGER -Bold $true | Out-Null
Add-TextBox -Slide $s8 -X 610 -Y 140 -W 310 -H 100 `
    -Text "COA(성적서 코드)는 출고 등록 시 필수 입력 항목입니다.`r`n`r`n• 비워두면 저장이 안 됩니다`r`n• 추후 추적 관리를 위해 정확히 입력하세요" `
    -FontSize 10 -Color (RGB 120 30 40) | Out-Null

# 우측: 자동 처리 안내
Add-RoundRect -Slide $s8 -X 590 -Y 270 -W 340 -H 175 -FillColor (RGB 230 248 240) -Radius 0.06 | Out-Null
Add-Rect -Slide $s8 -X 590 -Y 270 -W 4 -H 175 -FillColor $C_SUCCESS | Out-Null
Add-TextBox -Slide $s8 -X 610 -Y 285 -W 310 -H 25 `
    -Text "✅ 자동으로 처리되는 것" `
    -FontSize 13 -Color $C_SUCCESS -Bold $true | Out-Null
Add-TextBox -Slide $s8 -X 610 -Y 315 -W 310 -H 130 `
    -Text "• 출고번호 자동 생성`r`n  (O-YYYYMMDD-NNN)`r`n`r`n• 선택한 입고의 '출고량' 자동 증가`r`n`r`n• 현재 재고에서 자동 차감`r`n  (잔여 = 입고 - 출고)" `
    -FontSize 10 -Color (RGB 0 100 75) | Out-Null

Add-Footer $s8 8 12

# ============================================================
# 슬라이드 9 — 챕터 5 추가: 출고 디테일 보기
# ============================================================
Write-Host "[9/12] 출고 디테일 조회..." -ForegroundColor Yellow
$s9 = Add-BlankSlide $pres
Set-SlideBackground $s9 $C_WHITE
Add-ChapterHeader $s9 "05" "출고 목록 & 디테일 확인"

# 상단 설명
Add-TextBox -Slide $s9 -X 40 -Y 95 -W 880 -H 25 `
    -Text "출고 목록에서 행을 클릭하면 → 어떤 입고에서 몇 개씩 출고했는지 디테일을 볼 수 있습니다." `
    -FontSize 13 -Color $C_GRAY700 -Align "left" | Out-Null

# 좌측: 출고 목록 화면 흉내
Add-TextBox -Slide $s9 -X 40 -Y 135 -W 400 -H 22 `
    -Text "📋 출고 목록 화면" `
    -FontSize 12 -Color $C_GRAY900 -Bold $true | Out-Null

Add-RoundRect -Slide $s9 -X 40 -Y 165 -W 430 -H 280 -FillColor $C_WHITE -Radius 0.03 | Out-Null
Add-Rect -Slide $s9 -X 40 -Y 444 -W 430 -H 1 -FillColor $C_GRAY200 | Out-Null
Add-Rect -Slide $s9 -X 40 -Y 165 -W 430 -H 1 -FillColor $C_GRAY200 | Out-Null
Add-Rect -Slide $s9 -X 40 -Y 165 -W 1 -H 280 -FillColor $C_GRAY200 | Out-Null
Add-Rect -Slide $s9 -X 469 -Y 165 -W 1 -H 280 -FillColor $C_GRAY200 | Out-Null

# 헤더 행
Add-Rect -Slide $s9 -X 41 -Y 166 -W 428 -H 28 -FillColor $C_GRAY100 | Out-Null
Add-TextBox -Slide $s9 -X 50 -Y 172 -W 60 -H 18 -Text "출고번호" -FontSize 9 -Color $C_GRAY700 -Bold $true | Out-Null
Add-TextBox -Slide $s9 -X 165 -Y 172 -W 60 -H 18 -Text "일자" -FontSize 9 -Color $C_GRAY700 -Bold $true | Out-Null
Add-TextBox -Slide $s9 -X 240 -Y 172 -W 80 -H 18 -Text "COA" -FontSize 9 -Color $C_GRAY700 -Bold $true | Out-Null
Add-TextBox -Slide $s9 -X 350 -Y 172 -W 80 -H 18 -Text "총 수량" -FontSize 9 -Color $C_GRAY700 -Bold $true -Align "right" | Out-Null

# 데이터 행 (5개)
$rows = @(
    @{ id="O-20260528-003"; date="2026-05-28"; coa="COA-2026-K12"; qty="50" },
    @{ id="O-20260528-002"; date="2026-05-28"; coa="COA-2026-K11"; qty="30" },
    @{ id="O-20260527-001"; date="2026-05-27"; coa="COA-2026-K10"; qty="80" },
    @{ id="O-20260526-002"; date="2026-05-26"; coa="COA-2026-K09"; qty="120" },
    @{ id="O-20260526-001"; date="2026-05-26"; coa="COA-2026-K08"; qty="45" }
)
$ry = 200
for ($i = 0; $i -lt $rows.Count; $i++) {
    $bg = if ($i -eq 0) { (RGB 234 244 255) } else { $C_WHITE }  # 첫 행 강조
    Add-Rect -Slide $s9 -X 41 -Y $ry -W 428 -H 36 -FillColor $bg | Out-Null
    Add-TextBox -Slide $s9 -X 50 -Y ($ry + 9) -W 110 -H 18 -Text $rows[$i].id -FontSize 9 -Color $C_GRAY900 | Out-Null
    Add-TextBox -Slide $s9 -X 165 -Y ($ry + 9) -W 75 -H 18 -Text $rows[$i].date -FontSize 9 -Color $C_GRAY700 | Out-Null
    Add-TextBox -Slide $s9 -X 240 -Y ($ry + 9) -W 100 -H 18 -Text $rows[$i].coa -FontSize 9 -Color $C_GRAY700 | Out-Null
    Add-TextBox -Slide $s9 -X 350 -Y ($ry + 9) -W 80 -H 18 -Text $rows[$i].qty -FontSize 10 -Color $C_PRIMARY -Bold $true -Align "right" | Out-Null
    Add-Rect -Slide $s9 -X 41 -Y ($ry + 36) -W 428 -H 1 -FillColor $C_GRAY200 | Out-Null
    $ry += 37
}

# 화살표
Add-TextBox -Slide $s9 -X 478 -Y 270 -W 30 -H 30 -Text "→" -FontSize 24 -Color $C_PRIMARY -Bold $true -Align "center" | Out-Null
Add-TextBox -Slide $s9 -X 478 -Y 305 -W 30 -H 20 -Text "클릭" -FontSize 8 -Color $C_PRIMARY -Bold $true -Align "center" | Out-Null

# 우측: 모달 흉내
Add-TextBox -Slide $s9 -X 520 -Y 135 -W 400 -H 22 `
    -Text "🗂 출고 디테일 모달" `
    -FontSize 12 -Color $C_GRAY900 -Bold $true | Out-Null

Add-RoundRect -Slide $s9 -X 520 -Y 165 -W 410 -H 280 -FillColor $C_WHITE -Radius 0.04 | Out-Null
Add-Rect -Slide $s9 -X 520 -Y 165 -W 410 -H 1 -FillColor $C_GRAY200 | Out-Null
Add-Rect -Slide $s9 -X 520 -Y 444 -W 410 -H 1 -FillColor $C_GRAY200 | Out-Null
Add-Rect -Slide $s9 -X 520 -Y 165 -W 1 -H 280 -FillColor $C_GRAY200 | Out-Null
Add-Rect -Slide $s9 -X 929 -Y 165 -W 1 -H 280 -FillColor $C_GRAY200 | Out-Null

# 모달 헤더
Add-Rect -Slide $s9 -X 521 -Y 166 -W 408 -H 40 -FillColor $C_NAVY | Out-Null
Add-TextBox -Slide $s9 -X 535 -Y 175 -W 380 -H 22 `
    -Text "📤 출고 디테일 - O-20260528-003" `
    -FontSize 12 -Color $C_WHITE -Bold $true | Out-Null

# 모달 본문 - 출고된 입고 리스트
Add-TextBox -Slide $s9 -X 535 -Y 220 -W 380 -H 20 -Text "출고된 입고 제품 리스트" -FontSize 10 -Color $C_GRAY700 -Bold $true | Out-Null

$details = @(
    @{ id="P-20260520-001"; pdate="2026-05-20"; pkg="20kg"; qty="30" },
    @{ id="P-20260521-002"; pdate="2026-05-21"; pkg="25kg"; qty="20" }
)

# 디테일 테이블 헤더
Add-Rect -Slide $s9 -X 535 -Y 250 -W 380 -H 24 -FillColor $C_GRAY100 | Out-Null
Add-TextBox -Slide $s9 -X 545 -Y 254 -W 120 -H 18 -Text "제품번호" -FontSize 9 -Color $C_GRAY700 -Bold $true | Out-Null
Add-TextBox -Slide $s9 -X 670 -Y 254 -W 80 -H 18 -Text "생산일자" -FontSize 9 -Color $C_GRAY700 -Bold $true | Out-Null
Add-TextBox -Slide $s9 -X 770 -Y 254 -W 50 -H 18 -Text "포장" -FontSize 9 -Color $C_GRAY700 -Bold $true | Out-Null
Add-TextBox -Slide $s9 -X 840 -Y 254 -W 65 -H 18 -Text "수량" -FontSize 9 -Color $C_GRAY700 -Bold $true -Align "right" | Out-Null

$dy = 280
foreach ($d in $details) {
    Add-TextBox -Slide $s9 -X 545 -Y $dy -W 130 -H 18 -Text $d.id -FontSize 9 -Color $C_GRAY900 | Out-Null
    Add-TextBox -Slide $s9 -X 670 -Y $dy -W 100 -H 18 -Text $d.pdate -FontSize 9 -Color $C_GRAY700 | Out-Null
    Add-TextBox -Slide $s9 -X 770 -Y $dy -W 50 -H 18 -Text $d.pkg -FontSize 9 -Color $C_GRAY700 | Out-Null
    Add-TextBox -Slide $s9 -X 840 -Y $dy -W 65 -H 18 -Text $d.qty -FontSize 10 -Color $C_PRIMARY -Bold $true -Align "right" | Out-Null
    Add-Rect -Slide $s9 -X 535 -Y ($dy + 22) -W 380 -H 1 -FillColor $C_GRAY200 | Out-Null
    $dy += 30
}

# 합계
Add-TextBox -Slide $s9 -X 545 -Y 360 -W 200 -H 22 -Text "합계" -FontSize 11 -Color $C_GRAY900 -Bold $true | Out-Null
Add-TextBox -Slide $s9 -X 800 -Y 360 -W 105 -H 22 -Text "50 포장" -FontSize 13 -Color $C_PRIMARY -Bold $true -Align "right" | Out-Null

Add-Footer $s9 9 12

# ============================================================
# 슬라이드 10 — 챕터 6: 현재 재고 조회 + 엑셀
# ============================================================
Write-Host "[10/12] 챕터6: 현재 재고 조회..." -ForegroundColor Yellow
$s10 = Add-BlankSlide $pres
Set-SlideBackground $s10 $C_WHITE
Add-ChapterHeader $s10 "06" "현재 재고 조회 & 엑셀 다운로드"

# 좌측 단계
Add-TextBox -Slide $s10 -X 40 -Y 100 -W 500 -H 25 `
    -Text "📊 재고 확인 방법" `
    -FontSize 14 -Color $C_GRAY900 -Bold $true | Out-Null

$rsteps = @(
    "사이드바 [📊 현재 재고 조회] 메뉴 클릭",
    "전체 제품의 잔여 재고 한눈에 확인",
    "체크박스 [재고 있는 것만 보기] 로 필터",
    "[엑셀 다운로드] 버튼으로 .xlsx 파일 저장"
)
$sy = 130
for ($i = 0; $i -lt 4; $i++) {
    Add-StepBadge -Slide $s10 -X 50 -Y $sy -Size 28 -Num ([string]($i + 1)) -BgColor $C_SUCCESS | Out-Null
    Add-TextBox -Slide $s10 -X 90 -Y ($sy + 4) -W 480 -H 22 `
        -Text $rsteps[$i] `
        -FontSize 12 -Color $C_GRAY700 -Align "left" | Out-Null
    $sy += 38
}

# 컬럼 안내
Add-TextBox -Slide $s10 -X 40 -Y 295 -W 500 -H 22 `
    -Text "📋 화면에 표시되는 정보" `
    -FontSize 13 -Color $C_GRAY900 -Bold $true | Out-Null

$cols = @(
    @{ name = "입고번호"; desc = "P-YYYYMMDD-NNN 형식" },
    @{ name = "크기 (μm)"; desc = "제품 입자 크기" },
    @{ name = "초기재고 / 총입고량"; desc = "처음 입고된 포장 수" },
    @{ name = "총출고량"; desc = "지금까지 출고된 누적량" },
    @{ name = "현재재고"; desc = "잔여 = 입고량 - 출고량" }
)
$cy = 325
foreach ($c in $cols) {
    Add-Rect -Slide $s10 -X 55 -Y ($cy + 8) -W 6 -H 6 -FillColor $C_PRIMARY | Out-Null
    Add-TextBox -Slide $s10 -X 70 -Y $cy -W 200 -H 20 `
        -Text $c.name -FontSize 11 -Color $C_GRAY900 -Bold $true | Out-Null
    Add-TextBox -Slide $s10 -X 240 -Y $cy -W 320 -H 20 `
        -Text "— $($c.desc)" -FontSize 10 -Color $C_GRAY500 | Out-Null
    $cy += 23
}

# 우측: 엑셀 다운로드 강조 박스
Add-RoundRect -Slide $s10 -X 600 -Y 100 -W 330 -H 345 -FillColor (RGB 230 248 240) -Radius 0.05 | Out-Null
Add-Rect -Slide $s10 -X 600 -Y 100 -W 4 -H 345 -FillColor $C_SUCCESS | Out-Null

Add-TextBox -Slide $s10 -X 620 -Y 120 -W 290 -H 30 `
    -Text "📥 엑셀 다운로드" `
    -FontSize 17 -Color $C_SUCCESS -Bold $true | Out-Null

Add-TextBox -Slide $s10 -X 620 -Y 155 -W 290 -H 22 `
    -Text "현재 화면 그대로 엑셀로 받기" `
    -FontSize 11 -Color (RGB 0 100 75) | Out-Null

# 가짜 엑셀 아이콘 / 파일명 박스
Add-RoundRect -Slide $s10 -X 620 -Y 195 -W 290 -H 75 -FillColor $C_WHITE -Radius 0.08 | Out-Null
Add-TextBox -Slide $s10 -X 635 -Y 208 -W 30 -H 50 -Text "📄" -FontSize 28 | Out-Null
Add-TextBox -Slide $s10 -X 675 -Y 210 -W 230 -H 22 `
    -Text "현재재고_2026-05-28.xlsx" -FontSize 11 -Color $C_GRAY900 -Bold $true | Out-Null
Add-TextBox -Slide $s10 -X 675 -Y 232 -W 230 -H 18 `
    -Text "엑셀에서 바로 열기 가능" -FontSize 10 -Color $C_GRAY500 | Out-Null
Add-TextBox -Slide $s10 -X 675 -Y 250 -W 230 -H 18 `
    -Text "한글 파일명 자동 처리" -FontSize 10 -Color $C_GRAY500 | Out-Null

# 활용 예시
Add-TextBox -Slide $s10 -X 620 -Y 290 -W 290 -H 22 `
    -Text "✨ 이런 식으로 활용하세요" `
    -FontSize 11 -Color (RGB 0 100 75) -Bold $true | Out-Null
Add-TextBox -Slide $s10 -X 620 -Y 315 -W 290 -H 120 `
    -Text "• 거래처 보낼 재고 리스트 작성`r`n`r`n• 월말 재고 실사 자료 백업`r`n`r`n• 회계 자료로 활용`r`n`r`n• 인쇄 후 현장 비치" `
    -FontSize 10 -Color (RGB 0 90 65) | Out-Null

Add-Footer $s10 10 12

# ============================================================
# 슬라이드 11 — 자주 묻는 질문 & 주의사항
# ============================================================
Write-Host "[11/12] FAQ & 주의사항..." -ForegroundColor Yellow
$s11 = Add-BlankSlide $pres
Set-SlideBackground $s11 $C_WHITE
Add-ChapterHeader $s11 "TIP" "자주 묻는 질문 & 주의사항"

# 좌측: FAQ
Add-TextBox -Slide $s11 -X 40 -Y 100 -W 400 -H 25 `
    -Text "❓ 자주 묻는 질문" `
    -FontSize 14 -Color $C_GRAY900 -Bold $true | Out-Null

$faqs = @(
    @{ q = "Q. 화면이 안 떠요"; a = "A. 30초~1분 기다려보세요. 서버가 깨어나는 중입니다." },
    @{ q = "Q. 잘못 등록했어요"; a = "A. 입고 목록 / 출고 목록의 [삭제] 버튼으로 지울 수 있어요." },
    @{ q = "Q. 출고 삭제하면 재고는?"; a = "A. 자동으로 원복됩니다. 입고의 출고량이 그만큼 줄어들어요." },
    @{ q = "Q. 비밀번호를 잊었어요"; a = "A. 관리자(admin) 계정으로 사용자 관리에서 초기화 가능합니다." },
    @{ q = "Q. 모바일에서도 되나요?"; a = "A. 네! 핸드폰 브라우저로도 모든 기능 사용 가능합니다." }
)

$fy = 130
foreach ($faq in $faqs) {
    Add-TextBox -Slide $s11 -X 50 -Y $fy -W 420 -H 22 `
        -Text $faq.q -FontSize 11 -Color $C_PRIMARY -Bold $true | Out-Null
    Add-TextBox -Slide $s11 -X 50 -Y ($fy + 22) -W 420 -H 35 `
        -Text $faq.a -FontSize 10 -Color $C_GRAY700 | Out-Null
    Add-Rect -Slide $s11 -X 50 -Y ($fy + 58) -W 420 -H 1 -FillColor $C_GRAY200 | Out-Null
    $fy += 65
}

# 우측: 주의사항
Add-TextBox -Slide $s11 -X 510 -Y 100 -W 400 -H 25 `
    -Text "⚠️ 꼭 기억해주세요" `
    -FontSize 14 -Color $C_DANGER -Bold $true | Out-Null

$warnings = @(
    @{ icon = "🔒"; title = "비밀번호 관리"; desc = "다른 사람과 공유하지 마세요. 직원마다 별도 계정 발급 권장." },
    @{ icon = "💾"; title = "데이터는 자동 저장"; desc = "[저장] 버튼만 누르면 클라우드에 즉시 저장됩니다." },
    @{ icon = "🗑️"; title = "삭제는 신중히"; desc = "삭제 시 확인창이 뜹니다. 실수로 누르지 않도록 주의." },
    @{ icon = "🌐"; title = "인터넷 필수"; desc = "인터넷이 끊기면 화면이 안 뜹니다. 와이파이 확인하세요." },
    @{ icon = "📞"; title = "문제 발생 시"; desc = "캡쳐 후 개발자에게 문의 (구체적인 화면이 빠른 해결의 열쇠)" }
)

$wy = 130
foreach ($w in $warnings) {
    Add-RoundRect -Slide $s11 -X 510 -Y $wy -W 410 -H 50 -FillColor (RGB 254 245 246) -Radius 0.1 | Out-Null
    Add-TextBox -Slide $s11 -X 522 -Y ($wy + 12) -W 30 -H 30 `
        -Text $w.icon -FontSize 18 | Out-Null
    Add-TextBox -Slide $s11 -X 555 -Y ($wy + 7) -W 360 -H 18 `
        -Text $w.title -FontSize 11 -Color $C_GRAY900 -Bold $true | Out-Null
    Add-TextBox -Slide $s11 -X 555 -Y ($wy + 25) -W 360 -H 22 `
        -Text $w.desc -FontSize 9 -Color $C_GRAY700 | Out-Null
    $wy += 60
}

Add-Footer $s11 11 12

# ============================================================
# 슬라이드 12 — 마무리
# ============================================================
Write-Host "[12/12] 마무리 슬라이드..." -ForegroundColor Yellow
$s12 = Add-BlankSlide $pres
Set-SlideBackground $s12 $C_NAVY

# 배경 장식
$dec2 = Add-Rect -Slide $s12 -X -100 -Y 350 -W 600 -H 400 -FillColor $C_PRIMARY
$dec2.Fill.Transparency = 0.85
$dec2.Rotation = 15

# 큰 이모지 / 심볼
Add-TextBox -Slide $s12 -X 40 -Y 100 -W 880 -H 80 `
    -Text "🎉" -FontSize 56 -Align "center" | Out-Null

# 메인 메시지
Add-TextBox -Slide $s12 -X 40 -Y 195 -W 880 -H 55 `
    -Text "이제 시작해 보세요!" `
    -FontSize 42 -Color $C_WHITE -Bold $true -Align "center" | Out-Null

Add-TextBox -Slide $s12 -X 40 -Y 255 -W 880 -H 30 `
    -Text "복잡한 종이 장부는 그만, 클릭 몇 번이면 끝." `
    -FontSize 16 -Color (RGB 180 200 235) -Align "center" | Out-Null

# 하단 정보 박스 (3개)
$infos = @(
    @{ icon = "🌐"; title = "접속 주소"; val = "alamis-pj.onrender.com" },
    @{ icon = "🕐"; title = "운영 시간"; val = "24시간 · 연중무휴" },
    @{ icon = "💬"; title = "문의"; val = "개발자에게 연락" }
)
$ix = 100
for ($i = 0; $i -lt 3; $i++) {
    $x = $ix + $i * 250
    Add-RoundRect -Slide $s12 -X $x -Y 340 -W 220 -H 90 -FillColor (RGB 50 60 120) -Radius 0.08 | Out-Null
    Add-TextBox -Slide $s12 -X ($x + 15) -Y 355 -W 30 -H 25 `
        -Text $infos[$i].icon -FontSize 18 | Out-Null
    Add-TextBox -Slide $s12 -X ($x + 50) -Y 358 -W 160 -H 20 `
        -Text $infos[$i].title -FontSize 10 -Color (RGB 180 200 235) | Out-Null
    Add-TextBox -Slide $s12 -X ($x + 50) -Y 378 -W 160 -H 25 `
        -Text $infos[$i].val -FontSize 12 -Color $C_WHITE -Bold $true | Out-Null
}

# 푸터
Add-TextBox -Slide $s12 -X 40 -Y 480 -W 880 -H 25 `
    -Text "아라미스 펄 · 재고관리 시스템 · v1.0" `
    -FontSize 10 -Color (RGB 130 150 200) -Align "center" | Out-Null

# ============================================================
# 저장
# ============================================================
Write-Host ""
Write-Host "PPT 파일 저장 중..." -ForegroundColor Cyan
$outPath = "D:\testPro\docs\Alamis_재고관리_사용설명서.pptx"
$pres.SaveAs($outPath, 24)  # ppSaveAsOpenXMLPresentation = 24
$pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host "==========================================" -ForegroundColor Green
Write-Host "  완료!" -ForegroundColor Green
Write-Host "  경로: $outPath" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green

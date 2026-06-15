# ========================================================
# PPT 슬라이드 업데이트용 공통 헬퍼 함수 라이브러리
# (다른 스크립트에서 dot-source 해서 사용)
# ========================================================

# 색상 (Toss 디자인 시스템)
function RGB([int]$r, [int]$g, [int]$b) { return ($r + ($g * 256) + ($b * 65536)) }

$global:C_PRIMARY = RGB 49  130 246
$global:C_NAVY    = RGB 30  39  97
$global:C_SUCCESS = RGB 0   200 150
$global:C_DANGER  = RGB 240 68  82
$global:C_WARNING = RGB 255 153 0
$global:C_PURPLE  = RGB 138 43  226
$global:C_GRAY900 = RGB 25  31  40
$global:C_GRAY700 = RGB 78  89  104
$global:C_GRAY500 = RGB 140 150 165
$global:C_GRAY200 = RGB 226 230 236
$global:C_GRAY100 = RGB 242 244 246
$global:C_WHITE   = RGB 255 255 255

$global:FONT_KO = "맑은 고딕"

# === 슬라이드 클리어 & 배경 ===
function Clear-Slide {
    param($Slide, [int]$BgColor = -1)
    while ($Slide.Shapes.Count -gt 0) {
        $Slide.Shapes.Item(1).Delete()
    }
    if ($BgColor -lt 0) { $BgColor = $global:C_WHITE }
    $Slide.Background.Fill.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $Slide.Background.Fill.ForeColor.RGB = $BgColor
    $Slide.FollowMasterBackground = [Microsoft.Office.Core.MsoTriState]::msoFalse
}

# === 도형 ===
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

# === 텍스트박스 ===
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

# === 번호 마크업 동그라미 ===
function Add-NumberCircle {
    param($Slide, [single]$X, [single]$Y, [single]$Size, [string]$Num, [int]$BgColor, [int]$FontSize = 0)
    $oval = $Slide.Shapes.AddShape(9, $X, $Y, $Size, $Size)
    $oval.Fill.ForeColor.RGB = $BgColor
    $oval.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $oval.Line.ForeColor.RGB = $global:C_WHITE
    $oval.Line.Weight = 1.5
    $oval.Shadow.Type = 1
    $oval.Shadow.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $oval.TextFrame.MarginLeft = 0; $oval.TextFrame.MarginRight = 0
    $oval.TextFrame.MarginTop = 0;  $oval.TextFrame.MarginBottom = 0
    $oval.TextFrame.VerticalAnchor = 3
    $tr = $oval.TextFrame.TextRange
    $tr.Text = $Num
    $tr.Font.Name = $global:FONT_KO; $tr.Font.NameFarEast = $global:FONT_KO
    if ($FontSize -eq 0) { $FontSize = [int]($Size * 0.6) }
    $tr.Font.Size = $FontSize
    $tr.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $tr.Font.Color.RGB = $global:C_WHITE
    $tr.ParagraphFormat.Alignment = 2
    return $oval
}

# === 챕터 헤더 ===
function Add-ChapterHeader {
    param($Slide, [single]$SlideWidth, [string]$ChapterNo, [string]$ChapterTitle)
    $badge = Add-RoundRect -Slide $Slide -X 40 -Y 35 -W 65 -H 24 -FillColor $global:C_PRIMARY -Radius 0.5
    $badge.TextFrame.VerticalAnchor = 3
    $badge.TextFrame.MarginTop = 0; $badge.TextFrame.MarginBottom = 0
    $bt = $badge.TextFrame.TextRange
    $bt.Text = $ChapterNo
    $bt.Font.Name = $global:FONT_KO; $bt.Font.NameFarEast = $global:FONT_KO
    $bt.Font.Size = 11; $bt.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $bt.Font.Color.RGB = $global:C_WHITE
    $bt.ParagraphFormat.Alignment = 2

    Add-TextBox -Slide $Slide -X 115 -Y 36 -W ($SlideWidth - 150) -H 30 `
        -Text $ChapterTitle -FontSize 20 -Color $global:C_GRAY900 -Bold $true `
        -Align "left" -VAlign "middle" | Out-Null

    Add-Rect -Slide $Slide -X 40 -Y 75 -W ($SlideWidth - 80) -H 1 -FillColor $global:C_GRAY200 | Out-Null
}

# === 푸터 ===
function Add-Footer {
    param($Slide, [single]$SlideWidth, [single]$SlideHeight, [int]$Page, [int]$Total)
    Add-TextBox -Slide $Slide -X 30 -Y ($SlideHeight - 25) -W 200 -H 18 `
        -Text "아라미스 펄 · 재고관리 시스템" `
        -FontSize 9 -Color $global:C_GRAY500 -Align "left" | Out-Null
    Add-TextBox -Slide $Slide -X ($SlideWidth - 100) -Y ($SlideHeight - 25) -W 70 -H 18 `
        -Text "$Page / $Total" `
        -FontSize 9 -Color $global:C_GRAY500 -Align "right" | Out-Null
}

# === 캡쳐 이미지 삽입 ===
function Add-ScreenshotImage {
    param($Slide, [string]$ImagePath, [single]$X, [single]$Y, [single]$W, [single]$H)
    if (-not (Test-Path $ImagePath)) { throw "이미지 없음: $ImagePath" }
    # AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
    $pic = $Slide.Shapes.AddPicture($ImagePath, 0, -1, $X, $Y, $W, $H)
    $pic.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $pic.Line.ForeColor.RGB = $global:C_GRAY200
    $pic.Line.Weight = 0.75
    return $pic
}

# === 화살표 (point to something) ===
function Add-Arrow {
    # msoConnectorLine = 1 / msoShapeRightArrow = 33
    # 간단한 직선 + 화살표 헤드
    param($Slide, [single]$X1, [single]$Y1, [single]$X2, [single]$Y2, [int]$Color, [single]$Weight = 2)
    # AddLine(BeginX, BeginY, EndX, EndY)
    $line = $Slide.Shapes.AddLine($X1, $Y1, $X2, $Y2)
    $line.Line.ForeColor.RGB = $Color
    $line.Line.Weight = $Weight
    # 화살표 헤드 (BeginArrowheadStyle=1, EndArrowheadStyle=2 (triangle))
    $line.Line.EndArrowheadStyle = 2
    $line.Line.EndArrowheadWidth = 2
    $line.Line.EndArrowheadLength = 2
    return $line
}

Write-Host "[helpers] 공통 함수 라이브러리 로드 완료" -ForegroundColor DarkGray

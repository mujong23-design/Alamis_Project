# ========================================================
# 슬라이드 6 ~ 10 일괄 업데이트
# 실제 화면 캡쳐 + 번호 마크업 + 우측 설명
# ========================================================

$ErrorActionPreference = 'Stop'

# 공통 헬퍼 로드
. "D:\testPro\docs\_ppt_helpers.ps1"

$pptPath = "D:\testPro\docs\Alamis_재고관리_사용설명서.pptx"
$shotDir = "D:\testPro\docs\screenshots"

Write-Host "PowerPoint 열기..." -ForegroundColor Cyan
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pres = $ppt.Presentations.Open($pptPath)

$SW = $pres.PageSetup.SlideWidth   # 960
$SH = $pres.PageSetup.SlideHeight  # 540

# 캡쳐 → 슬라이드 좌표 변환 (1600x1000 → 600x375 위치 (20,95))
$imgX = 20; $imgY = 95; $imgW = 600; $imgH = 375
$scale = $imgW / 1600.0

function Cap-X { param([single]$cx) return $imgX + $cx * $scale }
function Cap-Y { param([single]$cy) return $imgY + $cy * $scale }

# 번호 마크업 그릴 때 사이즈 (캡쳐 위라서 작게)
$markSize = 18

# 우측 설명 패널
$rightX = 640; $rightW = 295

# ===================================================================
# 슬라이드 6 — 입고 등록
# ===================================================================
Write-Host "[6/10] 슬라이드 6: 입고 등록..." -ForegroundColor Yellow
$slide = $pres.Slides.Item(6)
Clear-Slide $slide
Add-ChapterHeader -Slide $slide -SlideWidth $SW -ChapterNo "03" -ChapterTitle "입고 등록 방법"
Add-ScreenshotImage -Slide $slide -ImagePath "$shotDir\04_inbound_form.png" -X $imgX -Y $imgY -W $imgW -H $imgH | Out-Null

# 캡쳐 내 좌표 (1600x1000):
# 생산일자: y=323
# 포장 단위(kg): x=440, y=410
# 포장 개수: x=740, y=410
# 크기(μm): y=498
# 등록 버튼: x=330, y=695
$marks6 = @(
    @{ no = "1"; cx = 600;  cy = 323; color = $global:C_PRIMARY },
    @{ no = "2"; cx = 350;  cy = 411; color = $global:C_PRIMARY },
    @{ no = "3"; cx = 645;  cy = 411; color = $global:C_PRIMARY },
    @{ no = "4"; cx = 600;  cy = 498; color = $global:C_PRIMARY },
    @{ no = "5"; cx = 335;  cy = 695; color = $global:C_SUCCESS }
)
foreach ($m in $marks6) {
    $cx = (Cap-X $m.cx) - ($markSize / 2)
    $cy = (Cap-Y $m.cy) - ($markSize / 2)
    Add-NumberCircle -Slide $slide -X $cx -Y $cy -Size $markSize -Num $m.no -BgColor $m.color -FontSize 10 | Out-Null
}

# 우측 설명
Add-TextBox -Slide $slide -X $rightX -Y 95 -W $rightW -H 25 `
    -Text "📝 입력 순서" -FontSize 14 -Color $global:C_GRAY900 -Bold $true | Out-Null

$steps6 = @(
    @{ no = "1"; t = "생산일자 선택"; d = "달력에서 클릭"; c = $global:C_PRIMARY },
    @{ no = "2"; t = "포장 단위 (kg)"; d = "한 포장의 무게 (예: 10)"; c = $global:C_PRIMARY },
    @{ no = "3"; t = "포장 개수"; d = "총 몇 포장인지 (예: 30)"; c = $global:C_PRIMARY },
    @{ no = "4"; t = "크기 (μm)"; d = "선택 — 정수2 + 소수3 (예: 66.110)"; c = $global:C_PRIMARY },
    @{ no = "5"; t = "[등록] 클릭"; d = "P-YYYYMMDD-NNN 자동 채번"; c = $global:C_SUCCESS }
)
$sy = 130
foreach ($s in $steps6) {
    Add-NumberCircle -Slide $slide -X ($rightX + 6) -Y ($sy + 4) -Size 22 -Num $s.no -BgColor $s.c -FontSize 12 | Out-Null
    Add-TextBox -Slide $slide -X ($rightX + 36) -Y $sy -W ($rightW - 40) -H 18 `
        -Text $s.t -FontSize 11 -Color $global:C_GRAY900 -Bold $true | Out-Null
    Add-TextBox -Slide $slide -X ($rightX + 36) -Y ($sy + 17) -W ($rightW - 40) -H 30 `
        -Text $s.d -FontSize 9 -Color $global:C_GRAY700 | Out-Null
    $sy += 55
}

# 하단 안내
Add-RoundRect -Slide $slide -X $imgX -Y 480 -W $imgW -H 35 -FillColor (RGB 234 244 255) -Radius 0.2 | Out-Null
Add-TextBox -Slide $slide -X ($imgX + 15) -Y 489 -W ($imgW - 30) -H 22 `
    -Text "💡 제품번호는 [등록] 누르면 자동으로 생성돼요. 직접 입력 안 해도 됩니다." `
    -FontSize 10 -Color (RGB 30 90 180) -Bold $true | Out-Null

Add-Footer -Slide $slide -SlideWidth $SW -SlideHeight $SH -Page 6 -Total 12

# ===================================================================
# 슬라이드 7 — 라벨 인쇄
# ===================================================================
Write-Host "[7/10] 슬라이드 7: 라벨 인쇄..." -ForegroundColor Yellow
$slide = $pres.Slides.Item(7)
Clear-Slide $slide
Add-ChapterHeader -Slide $slide -SlideWidth $SW -ChapterNo "04" -ChapterTitle "제품 라벨(스티커) 인쇄"
Add-ScreenshotImage -Slide $slide -ImagePath "$shotDir\03_inbound_list.png" -X $imgX -Y $imgY -W $imgW -H $imgH | Out-Null

# 캡쳐 내 좌표:
# 체크박스 (첫 행): x=302, y=310
# [선택 라벨 인쇄] 버튼: x=1390, y=108
# "미출력" 배지: x=1170, y=310
$marks7 = @(
    @{ no = "1"; cx = 302;  cy = 310; color = $global:C_PRIMARY },
    @{ no = "2"; cx = 1390; cy = 108; color = $global:C_DANGER },
    @{ no = "3"; cx = 1175; cy = 310; color = $global:C_WARNING }
)
foreach ($m in $marks7) {
    $cx = (Cap-X $m.cx) - ($markSize / 2)
    $cy = (Cap-Y $m.cy) - ($markSize / 2)
    Add-NumberCircle -Slide $slide -X $cx -Y $cy -Size $markSize -Num $m.no -BgColor $m.color -FontSize 10 | Out-Null
}

Add-TextBox -Slide $slide -X $rightX -Y 95 -W $rightW -H 25 `
    -Text "🏷️ 라벨 인쇄 순서" -FontSize 14 -Color $global:C_GRAY900 -Bold $true | Out-Null

$steps7 = @(
    @{ no = "1"; t = "체크박스 선택"; d = "라벨 인쇄할 항목 □ 클릭 (여러 개 선택 가능)"; c = $global:C_PRIMARY },
    @{ no = "2"; t = "[선택 라벨 인쇄]"; d = "우측 상단 버튼 클릭 → 새 창 열림"; c = $global:C_DANGER },
    @{ no = "3"; t = "출력 여부 확인"; d = "이미 인쇄한 항목은 '출력완료' 배지 표시"; c = $global:C_WARNING }
)
$sy = 130
foreach ($s in $steps7) {
    Add-NumberCircle -Slide $slide -X ($rightX + 6) -Y ($sy + 4) -Size 22 -Num $s.no -BgColor $s.c -FontSize 12 | Out-Null
    Add-TextBox -Slide $slide -X ($rightX + 36) -Y $sy -W ($rightW - 40) -H 18 `
        -Text $s.t -FontSize 11 -Color $global:C_GRAY900 -Bold $true | Out-Null
    Add-TextBox -Slide $slide -X ($rightX + 36) -Y ($sy + 17) -W ($rightW - 40) -H 35 `
        -Text $s.d -FontSize 9 -Color $global:C_GRAY700 | Out-Null
    $sy += 60
}

# A4 라벨 미리보기 (작게)
Add-TextBox -Slide $slide -X $rightX -Y 315 -W $rightW -H 20 `
    -Text "📄 한 장에 라벨 10개 (2×5)" -FontSize 10 -Color $global:C_GRAY700 -Bold $true | Out-Null
Add-RoundRect -Slide $slide -X $rightX -Y 340 -W $rightW -H 100 -FillColor (RGB 250 251 252) -Radius 0.02 | Out-Null
for ($r = 0; $r -lt 5; $r++) {
    for ($c = 0; $c -lt 2; $c++) {
        $lx = $rightX + 8 + $c * 142
        $ly = 348 + $r * 18
        Add-Rect -Slide $slide -X $lx -Y $ly -W 136 -H 14 -FillColor $global:C_WHITE | Out-Null
    }
}

# 하단 안내
Add-RoundRect -Slide $slide -X $imgX -Y 480 -W $imgW -H 35 -FillColor (RGB 255 249 230) -Radius 0.2 | Out-Null
Add-TextBox -Slide $slide -X ($imgX + 15) -Y 489 -W ($imgW - 30) -H 22 `
    -Text "💡 이미 인쇄한 것을 다시 누르면 '재출력 확인' 창이 떠요." `
    -FontSize 10 -Color (RGB 153 102 0) -Bold $true | Out-Null

Add-Footer -Slide $slide -SlideWidth $SW -SlideHeight $SH -Page 7 -Total 12

# ===================================================================
# 슬라이드 8 — 출고 등록
# ===================================================================
Write-Host "[8/10] 슬라이드 8: 출고 등록..." -ForegroundColor Yellow
$slide = $pres.Slides.Item(8)
Clear-Slide $slide
Add-ChapterHeader -Slide $slide -SlideWidth $SW -ChapterNo "05" -ChapterTitle "출고 등록 방법"
Add-ScreenshotImage -Slide $slide -ImagePath "$shotDir\06_outbound_form.png" -X $imgX -Y $imgY -W $imgW -H $imgH | Out-Null

# 캡쳐 내 좌표:
# ① 입고 제품 선택 영역 안내: x=315, y=170 (헤더 옆)
# ② 체크박스 + 수량 입력: x=335, y=280 (첫 행)
# ③ 출고일자: x=600, y=573
# ④ COA (필수): x=1235, y=573
# ⑤ [출고 등록] 버튼: x=337, y=858
$marks8 = @(
    @{ no = "1"; cx = 320;  cy = 170; color = $global:C_PRIMARY },
    @{ no = "2"; cx = 335;  cy = 280; color = $global:C_PRIMARY },
    @{ no = "3"; cx = 1235; cy = 573; color = $global:C_DANGER },
    @{ no = "4"; cx = 337;  cy = 858; color = $global:C_SUCCESS }
)
foreach ($m in $marks8) {
    $cx = (Cap-X $m.cx) - ($markSize / 2)
    $cy = (Cap-Y $m.cy) - ($markSize / 2)
    Add-NumberCircle -Slide $slide -X $cx -Y $cy -Size $markSize -Num $m.no -BgColor $m.color -FontSize 10 | Out-Null
}

Add-TextBox -Slide $slide -X $rightX -Y 95 -W $rightW -H 25 `
    -Text "📤 출고 등록 순서" -FontSize 14 -Color $global:C_GRAY900 -Bold $true | Out-Null

$steps8 = @(
    @{ no = "1"; t = "출고할 입고 제품 선택"; d = "리스트에서 출고할 제품을 □ 체크"; c = $global:C_PRIMARY },
    @{ no = "2"; t = "출고 수량 입력"; d = "각 행 우측에 '몇 개 나갈지' 입력 (잔여 이내)"; c = $global:C_PRIMARY },
    @{ no = "3"; t = "COA 입력 (필수!)"; d = "성적서 코드 — 비우면 저장 안 됨"; c = $global:C_DANGER },
    @{ no = "4"; t = "[출고 등록] 클릭"; d = "O-YYYYMMDD-NNN 자동 생성 + 재고 자동 차감"; c = $global:C_SUCCESS }
)
$sy = 130
foreach ($s in $steps8) {
    Add-NumberCircle -Slide $slide -X ($rightX + 6) -Y ($sy + 4) -Size 22 -Num $s.no -BgColor $s.c -FontSize 12 | Out-Null
    Add-TextBox -Slide $slide -X ($rightX + 36) -Y $sy -W ($rightW - 40) -H 18 `
        -Text $s.t -FontSize 11 -Color $global:C_GRAY900 -Bold $true | Out-Null
    Add-TextBox -Slide $slide -X ($rightX + 36) -Y ($sy + 17) -W ($rightW - 40) -H 35 `
        -Text $s.d -FontSize 9 -Color $global:C_GRAY700 | Out-Null
    $sy += 60
}

# 하단 COA 강조
Add-RoundRect -Slide $slide -X $imgX -Y 480 -W $imgW -H 35 -FillColor (RGB 254 235 238) -Radius 0.2 | Out-Null
Add-TextBox -Slide $slide -X ($imgX + 15) -Y 489 -W ($imgW - 30) -H 22 `
    -Text "⚠️ COA(성적서 코드)는 반드시 입력! 비어있으면 저장이 안 됩니다." `
    -FontSize 10 -Color $global:C_DANGER -Bold $true | Out-Null

Add-Footer -Slide $slide -SlideWidth $SW -SlideHeight $SH -Page 8 -Total 12

# ===================================================================
# 슬라이드 9 — 출고 목록 & 디테일
# ===================================================================
Write-Host "[9/10] 슬라이드 9: 출고 목록 디테일..." -ForegroundColor Yellow
$slide = $pres.Slides.Item(9)
Clear-Slide $slide
Add-ChapterHeader -Slide $slide -SlideWidth $SW -ChapterNo "05" -ChapterTitle "출고 목록 & 디테일 확인"
Add-ScreenshotImage -Slide $slide -ImagePath "$shotDir\05_outbound_list.png" -X $imgX -Y $imgY -W $imgW -H $imgH | Out-Null

# 캡쳐 내 좌표:
# 행 (클릭하면 모달): x=600, y=335
# [+ 출고 등록]: x=1517, y=108
# [삭제]: x=1517, y=335
$marks9 = @(
    @{ no = "1"; cx = 1517; cy = 108; color = $global:C_PRIMARY },
    @{ no = "2"; cx = 600;  cy = 335; color = $global:C_SUCCESS },
    @{ no = "3"; cx = 1517; cy = 335; color = $global:C_DANGER }
)
foreach ($m in $marks9) {
    $cx = (Cap-X $m.cx) - ($markSize / 2)
    $cy = (Cap-Y $m.cy) - ($markSize / 2)
    Add-NumberCircle -Slide $slide -X $cx -Y $cy -Size $markSize -Num $m.no -BgColor $m.color -FontSize 10 | Out-Null
}

Add-TextBox -Slide $slide -X $rightX -Y 95 -W $rightW -H 25 `
    -Text "🗂 출고 목록 사용법" -FontSize 14 -Color $global:C_GRAY900 -Bold $true | Out-Null

$steps9 = @(
    @{ no = "1"; t = "[+ 출고 등록]"; d = "새 출고를 등록하려면 우측 상단 버튼"; c = $global:C_PRIMARY },
    @{ no = "2"; t = "행 클릭 → 디테일"; d = "어떤 입고에서 몇 개씩 나갔는지 팝업으로 확인"; c = $global:C_SUCCESS },
    @{ no = "3"; t = "[삭제] 버튼"; d = "출고를 삭제하면 입고 재고가 자동으로 원복됩니다"; c = $global:C_DANGER }
)
$sy = 130
foreach ($s in $steps9) {
    Add-NumberCircle -Slide $slide -X ($rightX + 6) -Y ($sy + 4) -Size 22 -Num $s.no -BgColor $s.c -FontSize 12 | Out-Null
    Add-TextBox -Slide $slide -X ($rightX + 36) -Y $sy -W ($rightW - 40) -H 18 `
        -Text $s.t -FontSize 11 -Color $global:C_GRAY900 -Bold $true | Out-Null
    Add-TextBox -Slide $slide -X ($rightX + 36) -Y ($sy + 17) -W ($rightW - 40) -H 35 `
        -Text $s.d -FontSize 9 -Color $global:C_GRAY700 | Out-Null
    $sy += 60
}

# 디테일 모달 미리보기
Add-RoundRect -Slide $slide -X $rightX -Y 320 -W $rightW -H 130 -FillColor $global:C_GRAY100 -Radius 0.05 | Out-Null
Add-Rect -Slide $slide -X $rightX -Y 320 -W $rightW -H 28 -FillColor $global:C_NAVY | Out-Null
Add-TextBox -Slide $slide -X ($rightX + 10) -Y 325 -W ($rightW - 20) -H 18 `
    -Text "📤 출고 디테일 (모달 예시)" -FontSize 10 -Color $global:C_WHITE -Bold $true | Out-Null
Add-TextBox -Slide $slide -X ($rightX + 10) -Y 358 -W ($rightW - 20) -H 90 `
    -Text "제품번호       수량`r`nP-20260526-001   3`r`nP-20260521-002   2`r`n──────────────────`r`n합계              5 포장" `
    -FontSize 9 -Color $global:C_GRAY700 | Out-Null

# 하단 안내
Add-RoundRect -Slide $slide -X $imgX -Y 480 -W $imgW -H 35 -FillColor (RGB 230 248 240) -Radius 0.2 | Out-Null
Add-TextBox -Slide $slide -X ($imgX + 15) -Y 489 -W ($imgW - 30) -H 22 `
    -Text "💡 출고 행을 클릭만 해도 상세 내역이 팝업으로 떠요. 더블클릭 X !" `
    -FontSize 10 -Color (RGB 0 100 75) -Bold $true | Out-Null

Add-Footer -Slide $slide -SlideWidth $SW -SlideHeight $SH -Page 9 -Total 12

# ===================================================================
# 슬라이드 10 — 현재 재고 조회 & 엑셀
# ===================================================================
Write-Host "[10/10] 슬라이드 10: 재고 조회 & 엑셀..." -ForegroundColor Yellow
$slide = $pres.Slides.Item(10)
Clear-Slide $slide
Add-ChapterHeader -Slide $slide -SlideWidth $SW -ChapterNo "06" -ChapterTitle "현재 재고 조회 & 엑셀 다운로드"
Add-ScreenshotImage -Slide $slide -ImagePath "$shotDir\07_stock_list.png" -X $imgX -Y $imgY -W $imgW -H $imgH | Out-Null

# 캡쳐 내 좌표:
# "현재재고가 있는 것만" 체크박스: x=302, y=178
# [엑셀 다운로드] 버튼: x=1500, y=108
# "현재재고" 컬럼 값: x=1547, y=315
$marks10 = @(
    @{ no = "1"; cx = 302;  cy = 178; color = $global:C_PRIMARY },
    @{ no = "2"; cx = 1500; cy = 108; color = $global:C_SUCCESS },
    @{ no = "3"; cx = 1547; cy = 315; color = $global:C_WARNING }
)
foreach ($m in $marks10) {
    $cx = (Cap-X $m.cx) - ($markSize / 2)
    $cy = (Cap-Y $m.cy) - ($markSize / 2)
    Add-NumberCircle -Slide $slide -X $cx -Y $cy -Size $markSize -Num $m.no -BgColor $m.color -FontSize 10 | Out-Null
}

Add-TextBox -Slide $slide -X $rightX -Y 95 -W $rightW -H 25 `
    -Text "📊 재고 확인 방법" -FontSize 14 -Color $global:C_GRAY900 -Bold $true | Out-Null

$steps10 = @(
    @{ no = "1"; t = "재고 있는 것만 필터"; d = "체크하면 0인 항목은 숨김 처리"; c = $global:C_PRIMARY },
    @{ no = "2"; t = "[엑셀 다운로드]"; d = "현재 화면 그대로 .xlsx 파일로 저장"; c = $global:C_SUCCESS },
    @{ no = "3"; t = "현재재고 컬럼"; d = "= 총입고량 - 총출고량 (실제 남은 수량)"; c = $global:C_WARNING }
)
$sy = 130
foreach ($s in $steps10) {
    Add-NumberCircle -Slide $slide -X ($rightX + 6) -Y ($sy + 4) -Size 22 -Num $s.no -BgColor $s.c -FontSize 12 | Out-Null
    Add-TextBox -Slide $slide -X ($rightX + 36) -Y $sy -W ($rightW - 40) -H 18 `
        -Text $s.t -FontSize 11 -Color $global:C_GRAY900 -Bold $true | Out-Null
    Add-TextBox -Slide $slide -X ($rightX + 36) -Y ($sy + 17) -W ($rightW - 40) -H 35 `
        -Text $s.d -FontSize 9 -Color $global:C_GRAY700 | Out-Null
    $sy += 60
}

# 엑셀 다운로드 박스
Add-RoundRect -Slide $slide -X $rightX -Y 320 -W $rightW -H 125 -FillColor (RGB 230 248 240) -Radius 0.06 | Out-Null
Add-Rect -Slide $slide -X $rightX -Y 320 -W 4 -H 125 -FillColor $global:C_SUCCESS | Out-Null
Add-TextBox -Slide $slide -X ($rightX + 15) -Y 330 -W ($rightW - 25) -H 22 `
    -Text "📄 현재재고_YYYY-MM-DD.xlsx" -FontSize 11 -Color $global:C_SUCCESS -Bold $true | Out-Null
Add-TextBox -Slide $slide -X ($rightX + 15) -Y 355 -W ($rightW - 25) -H 85 `
    -Text "• 한글 파일명 자동 처리`r`n• 엑셀에서 바로 열기 가능`r`n• 월말 재고 실사 자료로 활용`r`n• 거래처 보낼 리스트 작성" `
    -FontSize 9 -Color (RGB 0 100 75) | Out-Null

# 하단 안내
Add-RoundRect -Slide $slide -X $imgX -Y 480 -W $imgW -H 35 -FillColor (RGB 230 248 240) -Radius 0.2 | Out-Null
Add-TextBox -Slide $slide -X ($imgX + 15) -Y 489 -W ($imgW - 30) -H 22 `
    -Text "💡 엑셀 파일은 다운로드 폴더에 저장됩니다." `
    -FontSize 10 -Color (RGB 0 100 75) -Bold $true | Out-Null

Add-Footer -Slide $slide -SlideWidth $SW -SlideHeight $SH -Page 10 -Total 12

# ===================================================================
# 저장
# ===================================================================
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
Write-Host "  슬라이드 6~10 모두 업데이트 완료!" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Green

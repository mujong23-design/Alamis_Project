# ========================================================
# 슬라이드 6 ~ 10 v2 — 순서 재정렬 (목록 → 등록)
# 슬라이드 6: 입고 목록 (라벨 인쇄 + 입고등록 버튼 마크업 통합)
# 슬라이드 7: 입고 등록 방법
# 슬라이드 8: 출고 목록
# 슬라이드 9: 출고 등록 방법
# 슬라이드 10: 현재 재고 조회 (챕터번호 06→07)
# ========================================================

$ErrorActionPreference = 'Stop'
. "D:\testPro\docs\_ppt_helpers.ps1"

$pptPath = "D:\testPro\docs\Alamis_재고관리_사용설명서.pptx"
$shotDir = "D:\testPro\docs\screenshots"

Write-Host "PowerPoint 열기..." -ForegroundColor Cyan
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pres = $ppt.Presentations.Open($pptPath)

$SW = $pres.PageSetup.SlideWidth
$SH = $pres.PageSetup.SlideHeight

# 좌표 변환 (1600x1000 캡쳐 → 600x375 위치 (20,95))
$imgX = 20; $imgY = 95; $imgW = 600; $imgH = 375
$scale = $imgW / 1600.0
function Cap-X { param([single]$cx) return $imgX + $cx * $scale }
function Cap-Y { param([single]$cy) return $imgY + $cy * $scale }

$markSize = 18
$rightX = 640; $rightW = 295

function Draw-Marks {
    param($Slide, $Marks)
    foreach ($m in $Marks) {
        $cx = (Cap-X $m.cx) - ($markSize / 2)
        $cy = (Cap-Y $m.cy) - ($markSize / 2)
        Add-NumberCircle -Slide $Slide -X $cx -Y $cy -Size $markSize -Num $m.no -BgColor $m.color -FontSize 10 | Out-Null
    }
}

function Draw-RightCards {
    param($Slide, $Title, $Steps)
    Add-TextBox -Slide $Slide -X $rightX -Y 95 -W $rightW -H 25 `
        -Text $Title -FontSize 14 -Color $global:C_GRAY900 -Bold $true | Out-Null
    $sy = 130
    foreach ($s in $Steps) {
        Add-NumberCircle -Slide $Slide -X ($rightX + 6) -Y ($sy + 4) -Size 22 -Num $s.no -BgColor $s.c -FontSize 12 | Out-Null
        Add-TextBox -Slide $Slide -X ($rightX + 36) -Y $sy -W ($rightW - 40) -H 18 `
            -Text $s.t -FontSize 11 -Color $global:C_GRAY900 -Bold $true | Out-Null
        Add-TextBox -Slide $Slide -X ($rightX + 36) -Y ($sy + 17) -W ($rightW - 40) -H 35 `
            -Text $s.d -FontSize 9 -Color $global:C_GRAY700 | Out-Null
        $sy += 55
    }
    return $sy
}

function Draw-BottomTip {
    param($Slide, $Text, $BgColor, $TextColor)
    Add-RoundRect -Slide $Slide -X $imgX -Y 480 -W $imgW -H 35 -FillColor $BgColor -Radius 0.2 | Out-Null
    Add-TextBox -Slide $Slide -X ($imgX + 15) -Y 489 -W ($imgW - 30) -H 22 `
        -Text $Text -FontSize 10 -Color $TextColor -Bold $true | Out-Null
}

# ===================================================================
# 슬라이드 6 — 입고 목록 (라벨 인쇄 통합)
# ===================================================================
Write-Host "[6] 입고 목록..." -ForegroundColor Yellow
$slide = $pres.Slides.Item(6)
Clear-Slide $slide
Add-ChapterHeader -Slide $slide -SlideWidth $SW -ChapterNo "03" -ChapterTitle "입고 목록"
Add-ScreenshotImage -Slide $slide -ImagePath "$shotDir\03_inbound_list.png" -X $imgX -Y $imgY -W $imgW -H $imgH | Out-Null

# 마크업
Draw-Marks -Slide $slide -Marks @(
    @{ no = "1"; cx = 302;  cy = 310; color = $global:C_PRIMARY },
    @{ no = "2"; cx = 1390; cy = 108; color = $global:C_DANGER  },
    @{ no = "3"; cx = 1175; cy = 310; color = $global:C_WARNING },
    @{ no = "4"; cx = 1517; cy = 108; color = $global:C_SUCCESS }
)

# 우측 설명
Draw-RightCards -Slide $slide -Title "📋 입고 목록 기능" -Steps @(
    @{ no = "1"; t = "행 체크박스";          d = "라벨 인쇄할 항목을 □ 클릭 (여러 개 가능)"; c = $global:C_PRIMARY },
    @{ no = "2"; t = "[선택 라벨 인쇄]";    d = "체크한 항목의 제품 스티커 출력 (A4에 10장)"; c = $global:C_DANGER  },
    @{ no = "3"; t = "출력 여부 확인";       d = "출력완료(초록) / 미출력(회색) 배지 표시"; c = $global:C_WARNING },
    @{ no = "4"; t = "[+ 입고 등록]";        d = "새 제품을 등록하려면 클릭 → 다음 페이지"; c = $global:C_SUCCESS }
) | Out-Null

Draw-BottomTip -Slide $slide `
    -Text "💡 이미 인쇄한 항목을 다시 누르면 '재출력 확인' 창이 떠요." `
    -BgColor (RGB 255 249 230) -TextColor (RGB 153 102 0)

Add-Footer -Slide $slide -SlideWidth $SW -SlideHeight $SH -Page 6 -Total 12

# ===================================================================
# 슬라이드 7 — 입고 등록 방법
# ===================================================================
Write-Host "[7] 입고 등록 방법..." -ForegroundColor Yellow
$slide = $pres.Slides.Item(7)
Clear-Slide $slide
Add-ChapterHeader -Slide $slide -SlideWidth $SW -ChapterNo "04" -ChapterTitle "입고 등록 방법"
Add-ScreenshotImage -Slide $slide -ImagePath "$shotDir\04_inbound_form.png" -X $imgX -Y $imgY -W $imgW -H $imgH | Out-Null

Draw-Marks -Slide $slide -Marks @(
    @{ no = "1"; cx = 600;  cy = 323; color = $global:C_PRIMARY },
    @{ no = "2"; cx = 350;  cy = 411; color = $global:C_PRIMARY },
    @{ no = "3"; cx = 645;  cy = 411; color = $global:C_PRIMARY },
    @{ no = "4"; cx = 600;  cy = 498; color = $global:C_PRIMARY },
    @{ no = "5"; cx = 335;  cy = 695; color = $global:C_SUCCESS }
)

Draw-RightCards -Slide $slide -Title "📝 입력 순서" -Steps @(
    @{ no = "1"; t = "생산일자 선택";     d = "달력에서 클릭";                       c = $global:C_PRIMARY },
    @{ no = "2"; t = "포장 단위 (kg)";    d = "한 포장의 무게 (예: 10)";              c = $global:C_PRIMARY },
    @{ no = "3"; t = "포장 개수";          d = "총 몇 포장인지 (예: 30)";              c = $global:C_PRIMARY },
    @{ no = "4"; t = "크기 (μm)";          d = "선택 — 정수2 + 소수3 (예: 66.110)";   c = $global:C_PRIMARY },
    @{ no = "5"; t = "[등록] 클릭";        d = "P-YYYYMMDD-NNN 자동 채번";              c = $global:C_SUCCESS }
) | Out-Null

Draw-BottomTip -Slide $slide `
    -Text "💡 제품번호는 [등록] 누르면 자동으로 생성돼요. 직접 입력 X" `
    -BgColor (RGB 234 244 255) -TextColor (RGB 30 90 180)

Add-Footer -Slide $slide -SlideWidth $SW -SlideHeight $SH -Page 7 -Total 12

# ===================================================================
# 슬라이드 8 — 출고 목록
# ===================================================================
Write-Host "[8] 출고 목록..." -ForegroundColor Yellow
$slide = $pres.Slides.Item(8)
Clear-Slide $slide
Add-ChapterHeader -Slide $slide -SlideWidth $SW -ChapterNo "05" -ChapterTitle "출고 목록"
Add-ScreenshotImage -Slide $slide -ImagePath "$shotDir\05_outbound_list.png" -X $imgX -Y $imgY -W $imgW -H $imgH | Out-Null

Draw-Marks -Slide $slide -Marks @(
    @{ no = "1"; cx = 1517; cy = 108; color = $global:C_SUCCESS },
    @{ no = "2"; cx = 600;  cy = 335; color = $global:C_PRIMARY },
    @{ no = "3"; cx = 1517; cy = 335; color = $global:C_DANGER  }
)

Draw-RightCards -Slide $slide -Title "🗂 출고 목록 사용법" -Steps @(
    @{ no = "1"; t = "[+ 출고 등록]";       d = "새 출고를 등록하려면 클릭 → 다음 페이지";       c = $global:C_SUCCESS },
    @{ no = "2"; t = "행 클릭 → 디테일";   d = "어떤 입고에서 몇 개씩 나갔는지 팝업으로 확인"; c = $global:C_PRIMARY },
    @{ no = "3"; t = "[삭제] 버튼";          d = "출고 삭제 시 입고 재고가 자동으로 원복";       c = $global:C_DANGER  }
) | Out-Null

# 디테일 모달 미리보기
Add-RoundRect -Slide $slide -X $rightX -Y 305 -W $rightW -H 140 -FillColor $global:C_GRAY100 -Radius 0.05 | Out-Null
Add-Rect -Slide $slide -X $rightX -Y 305 -W $rightW -H 28 -FillColor $global:C_NAVY | Out-Null
Add-TextBox -Slide $slide -X ($rightX + 10) -Y 310 -W ($rightW - 20) -H 18 `
    -Text "📤 출고 디테일 (모달 예시)" -FontSize 10 -Color $global:C_WHITE -Bold $true | Out-Null
Add-TextBox -Slide $slide -X ($rightX + 10) -Y 343 -W ($rightW - 20) -H 95 `
    -Text "제품번호           수량`r`nP-20260526-001       3`r`nP-20260521-002       2`r`n────────────────────────`r`n합계                    5 포장" `
    -FontSize 9 -Color $global:C_GRAY700 | Out-Null

Draw-BottomTip -Slide $slide `
    -Text "💡 행을 한 번 클릭만 해도 상세 내역 팝업이 떠요. 더블클릭 X" `
    -BgColor (RGB 230 248 240) -TextColor (RGB 0 100 75)

Add-Footer -Slide $slide -SlideWidth $SW -SlideHeight $SH -Page 8 -Total 12

# ===================================================================
# 슬라이드 9 — 출고 등록 방법
# ===================================================================
Write-Host "[9] 출고 등록 방법..." -ForegroundColor Yellow
$slide = $pres.Slides.Item(9)
Clear-Slide $slide
Add-ChapterHeader -Slide $slide -SlideWidth $SW -ChapterNo "06" -ChapterTitle "출고 등록 방법"
Add-ScreenshotImage -Slide $slide -ImagePath "$shotDir\06_outbound_form.png" -X $imgX -Y $imgY -W $imgW -H $imgH | Out-Null

Draw-Marks -Slide $slide -Marks @(
    @{ no = "1"; cx = 335;  cy = 280; color = $global:C_PRIMARY },
    @{ no = "2"; cx = 1444; cy = 280; color = $global:C_PRIMARY },
    @{ no = "3"; cx = 1235; cy = 573; color = $global:C_DANGER  },
    @{ no = "4"; cx = 337;  cy = 858; color = $global:C_SUCCESS }
)

Draw-RightCards -Slide $slide -Title "📤 출고 등록 순서" -Steps @(
    @{ no = "1"; t = "제품 선택 (체크박스)";  d = "출고할 입고 제품을 □ 체크 (여러 개 가능)";       c = $global:C_PRIMARY },
    @{ no = "2"; t = "출고 수량 입력";          d = "각 행 우측에 '몇 개 나갈지' 입력 (잔여 이내)";  c = $global:C_PRIMARY },
    @{ no = "3"; t = "COA 입력 (필수!)";         d = "성적서 코드 — 비우면 저장 안 됨";              c = $global:C_DANGER  },
    @{ no = "4"; t = "[출고 등록] 클릭";        d = "O-YYYYMMDD-NNN 자동 생성 + 재고 자동 차감";    c = $global:C_SUCCESS }
) | Out-Null

Draw-BottomTip -Slide $slide `
    -Text "⚠️ COA(성적서 코드)는 반드시 입력! 비어있으면 저장이 안 됩니다." `
    -BgColor (RGB 254 235 238) -TextColor $global:C_DANGER

Add-Footer -Slide $slide -SlideWidth $SW -SlideHeight $SH -Page 9 -Total 12

# ===================================================================
# 슬라이드 10 — 현재 재고 조회 (챕터번호만 07로 변경)
# ===================================================================
Write-Host "[10] 현재 재고 조회 (챕터 07)..." -ForegroundColor Yellow
$slide = $pres.Slides.Item(10)
Clear-Slide $slide
Add-ChapterHeader -Slide $slide -SlideWidth $SW -ChapterNo "07" -ChapterTitle "현재 재고 조회 & 엑셀 다운로드"
Add-ScreenshotImage -Slide $slide -ImagePath "$shotDir\07_stock_list.png" -X $imgX -Y $imgY -W $imgW -H $imgH | Out-Null

Draw-Marks -Slide $slide -Marks @(
    @{ no = "1"; cx = 302;  cy = 178; color = $global:C_PRIMARY },
    @{ no = "2"; cx = 1500; cy = 108; color = $global:C_SUCCESS },
    @{ no = "3"; cx = 1547; cy = 315; color = $global:C_WARNING }
)

Draw-RightCards -Slide $slide -Title "📊 재고 확인 방법" -Steps @(
    @{ no = "1"; t = "재고 있는 것만 필터";  d = "체크하면 0인 항목은 숨김 처리";                  c = $global:C_PRIMARY },
    @{ no = "2"; t = "[엑셀 다운로드]";       d = "현재 화면 그대로 .xlsx 파일로 저장";              c = $global:C_SUCCESS },
    @{ no = "3"; t = "현재재고 컬럼";          d = "= 총입고량 - 총출고량 (실제 남은 수량)";          c = $global:C_WARNING }
) | Out-Null

# 엑셀 박스
Add-RoundRect -Slide $slide -X $rightX -Y 300 -W $rightW -H 145 -FillColor (RGB 230 248 240) -Radius 0.06 | Out-Null
Add-Rect -Slide $slide -X $rightX -Y 300 -W 4 -H 145 -FillColor $global:C_SUCCESS | Out-Null
Add-TextBox -Slide $slide -X ($rightX + 15) -Y 312 -W ($rightW - 25) -H 22 `
    -Text "📄 현재재고_YYYY-MM-DD.xlsx" -FontSize 11 -Color $global:C_SUCCESS -Bold $true | Out-Null
Add-TextBox -Slide $slide -X ($rightX + 15) -Y 338 -W ($rightW - 25) -H 100 `
    -Text "• 한글 파일명 자동 처리`r`n• 엑셀에서 바로 열기 가능`r`n• 월말 재고 실사 자료로 활용`r`n• 거래처 보낼 리스트 작성" `
    -FontSize 9 -Color (RGB 0 100 75) | Out-Null

Draw-BottomTip -Slide $slide `
    -Text "💡 엑셀 파일은 다운로드 폴더에 저장됩니다." `
    -BgColor (RGB 230 248 240) -TextColor (RGB 0 100 75)

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
Write-Host "  슬라이드 6~10 v2 완료!" -ForegroundColor Green
Write-Host "  순서: 입고목록 → 입고등록 → 출고목록 → 출고등록 → 재고조회" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Green

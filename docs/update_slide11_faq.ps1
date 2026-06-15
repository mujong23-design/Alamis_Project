# 슬라이드 11 (FAQ) 에서 "사용자 관리" 관련 답변 정리
# - 비밀번호 잊었어요 답변에 "관리자/사용자 관리" 언급을 부드럽게 변경

$ErrorActionPreference = 'Stop'

$pptPath = "D:\testPro\docs\Alamis_재고관리_사용설명서.pptx"

Write-Host "PowerPoint 열기..." -ForegroundColor Cyan
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pres = $ppt.Presentations.Open($pptPath)

$slide = $pres.Slides.Item(11)

# 모든 텍스트박스를 순회하며 특정 문구를 찾아 교체
$replacements = @{
    "A. 관리자(admin) 계정으로 사용자 관리에서 초기화 가능합니다." = "A. 관리자(또는 사장님)께 문의해 주세요. 새 비밀번호를 발급해 드립니다."
    "다른 사람과 공유하지 마세요. 직원마다 별도 계정 발급 권장." = "다른 사람과 공유하지 마세요. 본인만 사용하세요."
}

$changedCount = 0
foreach ($shape in $slide.Shapes) {
    if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
        $tr = $shape.TextFrame.TextRange
        $currentText = $tr.Text
        foreach ($key in $replacements.Keys) {
            if ($currentText -like "*$key*") {
                # 텍스트 통째로 교체 (서식 보존을 위해 부분 텍스트만 변경하기보다 전체 교체)
                $newText = $currentText.Replace($key, $replacements[$key])
                $tr.Text = $newText
                Write-Host "교체됨: '$key' → '$($replacements[$key])'" -ForegroundColor Green
                $changedCount++
            }
        }
    }
}

Write-Host ""
Write-Host "총 $changedCount 개 텍스트 변경됨" -ForegroundColor Yellow

# 저장
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
Write-Host "  FAQ 슬라이드 정리 완료!" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Green

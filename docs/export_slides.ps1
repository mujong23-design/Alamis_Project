# 슬라이드를 PNG로 추출 (QA용)
$ErrorActionPreference = 'Stop'
$pptPath = "D:\testPro\docs\Alamis_재고관리_사용설명서.pptx"
$outDir = "D:\testPro\docs\slides_preview"

if (Test-Path $outDir) {
    Remove-Item "$outDir\*.png" -Force -ErrorAction SilentlyContinue
} else {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pres = $ppt.Presentations.Open($pptPath, $false, $false, $false)

foreach ($slide in $pres.Slides) {
    $n = "{0:D2}" -f $slide.SlideNumber
    $outPath = "$outDir\slide_$n.png"
    # Export(filename, filterName, ScaleWidth, ScaleHeight) - 1920x1080 for HD
    $slide.Export($outPath, "PNG", 1920, 1080)
    Write-Host "Exported: $outPath"
}

$pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host "총 $($pres.Slides.Count)개 슬라이드 추출 완료"

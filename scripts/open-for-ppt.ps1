# HTML to PowerPoint Converter (Alternative Method)
# Opens HTML in browser and provides instructions for manual conversion

param(
    [string]$HtmlPath = "slides\Customer Systems Presentation_v4.html"
)

$HtmlFullPath = (Resolve-Path $HtmlPath).Path
$HtmlUri = "file:///$($HtmlFullPath.Replace('\', '/'))"

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "HTML to PowerPoint Converter" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Host "Opening HTML file in browser..." -ForegroundColor Yellow
Start-Process $HtmlUri

Write-Host "`nFollow these steps to create PowerPoint slide:" -ForegroundColor Green
Write-Host "`nMETHOD 1: Screenshot (Recommended)" -ForegroundColor Cyan
Write-Host "1. Press F11 for fullscreen mode" -ForegroundColor White
Write-Host "2. Press Windows + Shift + S to open Snipping Tool" -ForegroundColor White
Write-Host "3. Select 'Full Screen Snip' or drag to capture entire page" -ForegroundColor White
Write-Host "4. Save the screenshot (Ctrl+S)" -ForegroundColor White
Write-Host "5. Open PowerPoint > New Slide > Blank" -ForegroundColor White
Write-Host "6. Insert > Pictures > This Device > Select your screenshot" -ForegroundColor White
Write-Host "7. Right-click image > Format Picture > Crop to fit slide" -ForegroundColor White

Write-Host "`nMETHOD 2: Print to PDF then Insert" -ForegroundColor Cyan
Write-Host "1. In browser, press Ctrl+P (Print)" -ForegroundColor White
Write-Host "2. Select 'Save as PDF' as destination" -ForegroundColor White
Write-Host "3. Set margins to 'None' and scale to 'Fit to page'" -ForegroundColor White
Write-Host "4. Save PDF" -ForegroundColor White
Write-Host "5. In PowerPoint: Insert > Object > Create from File > Select PDF" -ForegroundColor White

Write-Host "`nMETHOD 3: Copy-Paste (Quick)" -ForegroundColor Cyan
Write-Host "1. In browser, press Ctrl+A (Select All)" -ForegroundColor White
Write-Host "2. Press Ctrl+C (Copy)" -ForegroundColor White
Write-Host "3. In PowerPoint, create new slide and press Ctrl+V" -ForegroundColor White
Write-Host "4. Format as needed" -ForegroundColor White

Write-Host "`n========================================`n" -ForegroundColor Cyan


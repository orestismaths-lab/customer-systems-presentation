# HTML to PowerPoint Converter
# Takes a screenshot of the HTML file and inserts it into a PowerPoint slide

param(
    [string]$HtmlPath = "slides\Customer Systems Presentation_v4.html",
    [string]$OutputPath = "slides\Customer Systems Presentation_v4.pptx"
)

$ErrorActionPreference = "Stop"

Write-Host "Converting HTML to PowerPoint..." -ForegroundColor Cyan

# Get absolute paths
$HtmlFullPath = (Resolve-Path $HtmlPath).Path
$OutputFullPath = Join-Path (Get-Location) $OutputPath
$ScreenshotPath = Join-Path $env:TEMP "presentation_screenshot.png"

Write-Host "HTML file: $HtmlFullPath" -ForegroundColor Gray
Write-Host "Output PPTX: $OutputFullPath" -ForegroundColor Gray

# Step 1: Take screenshot using Edge
Write-Host "`nTaking screenshot..." -ForegroundColor Yellow

$edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if (-not (Test-Path $edgePath)) {
    $edgePath = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
}

if (-not (Test-Path $edgePath)) {
    Write-Host "ERROR: Microsoft Edge not found!" -ForegroundColor Red
    Write-Host "Please take a screenshot manually:" -ForegroundColor Yellow
    Write-Host "1. Open the HTML file in your browser" -ForegroundColor Yellow
    Write-Host "2. Press F11 for fullscreen" -ForegroundColor Yellow
    Write-Host "3. Press Windows+Shift+S for screenshot tool" -ForegroundColor Yellow
    Write-Host "4. Save the screenshot" -ForegroundColor Yellow
    Write-Host "5. Open PowerPoint and insert the image" -ForegroundColor Yellow
    exit 1
}

# Use Edge headless to take screenshot
$htmlUri = "file:///$($HtmlFullPath.Replace('\', '/'))"
Write-Host "Opening HTML in Edge..." -ForegroundColor Gray

Start-Process $edgePath -ArgumentList "--headless", "--disable-gpu", "--window-size=1920,1080", "--screenshot=$ScreenshotPath", $htmlUri -Wait -NoNewWindow

if (-not (Test-Path $ScreenshotPath)) {
    Write-Host "ERROR: Screenshot failed!" -ForegroundColor Red
    Write-Host "Please take a screenshot manually (see instructions above)" -ForegroundColor Yellow
    exit 1
}

Write-Host "Screenshot saved: $ScreenshotPath" -ForegroundColor Green

# Step 2: Create PowerPoint presentation
Write-Host "`nCreating PowerPoint presentation..." -ForegroundColor Yellow

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    
    $presentation = $ppt.Presentations.Add()
    $slide = $presentation.Slides.Add(1, 11) # 11 = Blank layout
    
    # Set slide size to 16:9 (widescreen)
    $presentation.PageSetup.SlideWidth = 960  # 10 inches * 96 DPI
    $presentation.PageSetup.SlideHeight = 540  # 5.625 inches * 96 DPI
    
    # Insert screenshot
    $left = 0
    $top = 0
    $width = $presentation.PageSetup.SlideWidth
    $height = $presentation.PageSetup.SlideHeight
    
    $pic = $slide.Shapes.AddPicture($ScreenshotPath, $false, $true, $left, $top, $width, $height)
    
    # Fit image to slide
    $pic.LockAspectRatio = $true
    $pic.Width = $width
    if ($pic.Height -gt $height) {
        $pic.Height = $height
        $pic.Left = ($width - $pic.Width) / 2
    }
    $pic.Top = ($height - $pic.Height) / 2
    
    # Save presentation
    $presentation.SaveAs($OutputFullPath)
    Write-Host "`nPowerPoint saved: $OutputFullPath" -ForegroundColor Green
    
    # Cleanup
    $presentation.Close()
    $ppt.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
    
    # Clean up screenshot
    Remove-Item $ScreenshotPath -ErrorAction SilentlyContinue
    
    Write-Host "`nDone! PowerPoint file created successfully." -ForegroundColor Green
    
} catch {
    Write-Host "`nERROR creating PowerPoint: $_" -ForegroundColor Red
    Write-Host "`nManual steps:" -ForegroundColor Yellow
    Write-Host "1. Open PowerPoint" -ForegroundColor Yellow
    Write-Host "2. Create a new blank slide" -ForegroundColor Yellow
    Write-Host "3. Insert > Pictures > This Device" -ForegroundColor Yellow
    Write-Host "4. Select: $ScreenshotPath" -ForegroundColor Yellow
    Write-Host "5. Resize image to fit slide" -ForegroundColor Yellow
    exit 1
}


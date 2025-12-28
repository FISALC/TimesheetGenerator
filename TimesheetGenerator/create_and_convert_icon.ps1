Add-Type -AssemblyName System.Drawing

$width = 256
$height = 256
$maroonColor = [System.Drawing.ColorTranslator]::FromHtml("#963634")
$whiteColor = [System.Drawing.Color]::White

$bmp = New-Object System.Drawing.Bitmap $width, $height
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.Clear($whiteColor)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias

# Draw Maroon Circle
$brush = New-Object System.Drawing.SolidBrush $maroonColor
$g.FillEllipse($brush, 10, 10, 236, 236)

# Create Font - using float for size to resolve ambiguity
$fontSize = [float]100
$fontStyle = [System.Drawing.FontStyle]::Bold
$font = New-Object System.Drawing.Font("Segoe UI", $fontSize, $fontStyle)

$textBrush = New-Object System.Drawing.SolidBrush $whiteColor
$stringFormat = New-Object System.Drawing.StringFormat
$stringFormat.Alignment = [System.Drawing.StringAlignment]::Center
$stringFormat.LineAlignment = [System.Drawing.StringAlignment]::Center

# Draw "VG"
$g.DrawString("VG", $font, $textBrush, 128, 138, $stringFormat)

# Save as PNG
$pngPath = "app_icon.png"
if (Test-Path $pngPath) { Remove-Item $pngPath }
$bmp.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)

# Convert to ICO
$icoPath = "logo.ico"
if (Test-Path $icoPath) { Remove-Item $icoPath }

$handle = $bmp.GetHicon()
$icon = [System.Drawing.Icon]::FromHandle($handle)
$fs = [System.IO.File]::OpenWrite($icoPath)
$icon.Save($fs)
$fs.Close()

$g.Dispose()
$bmp.Dispose()
$icon.Dispose()

Write-Host "Success: Created modern $pngPath and converted to $icoPath"

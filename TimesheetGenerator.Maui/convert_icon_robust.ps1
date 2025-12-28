Add-Type -AssemblyName System.Drawing

$inputFile = "c:\MCIT\Projects\Experiment\TimesheetGenerator\TimesheetGenerator.Maui\logo.png"
$outputFile = "c:\MCIT\Projects\Experiment\TimesheetGenerator\TimesheetGenerator.Maui\logo.ico"

if (-not (Test-Path $inputFile)) {
    Write-Error "Input file not found: $inputFile"
    exit 1
}

$bitmap = [System.Drawing.Bitmap]::FromFile($inputFile)

# Ensure 256x256 max
if ($bitmap.Width -gt 256 -or $bitmap.Height -gt 256) {
    $resized = new-object System.Drawing.Bitmap 256, 256
    $g = [System.Drawing.Graphics]::FromImage($resized)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.DrawImage($bitmap, 0, 0, 256, 256)
    $g.Dispose()
    $bitmap.Dispose()
    $bitmap = $resized
}

$ms = New-Object System.IO.MemoryStream
$bitmap.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
$pngBytes = $ms.ToArray()
$ms.Dispose()

$fs = [System.IO.File]::Create($outputFile)
$bw = New-Object System.IO.BinaryWriter($fs)

# ICO Header
$bw.Write([int16]0)  # Reserved
$bw.Write([int16]1)  # Type (1=Icon)
$bw.Write([int16]1)  # Count (1 image)

# Image Directory Entry
$w = if ($bitmap.Width -ge 256) { 0 } else { [byte]$bitmap.Width }
$h = if ($bitmap.Height -ge 256) { 0 } else { [byte]$bitmap.Height }
$bw.Write([byte]$w)
$bw.Write([byte]$h)
$bw.Write([byte]0)   # ColorCount
$bw.Write([byte]0)   # Reserved
$bw.Write([int16]0)  # Planes
$bw.Write([int16]32) # BitCount
$bw.Write([int]$pngBytes.Length) # SizeInBytes
$bw.Write([int](6 + 16)) # ImageOffset (Header 6 + DirEntry 16)

# Image Data
$bw.Write($pngBytes)

$bw.Close()
$fs.Close()
$bitmap.Dispose()

Write-Host "Successfully created logo.ico ($($pngBytes.Length + 22) bytes)"

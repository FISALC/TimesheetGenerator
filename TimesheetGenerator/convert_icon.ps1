Add-Type -AssemblyName System.Drawing
$inputFile = "logo.png"
$outputFile = "logo.ico"

if (Test-Path $inputFile) {
    try {
        $img = [System.Drawing.Bitmap]::FromFile($inputFile)
        $handle = $img.GetHicon()
        $icon = [System.Drawing.Icon]::FromHandle($handle)
        
        $fs = [System.IO.File]::OpenWrite($outputFile)
        $icon.Save($fs)
        $fs.Close()
        
        $icon.Dispose()
        $img.Dispose()
        Write-Host "Success: Created $outputFile"
    }
    catch {
        Write-Host "Error: $_"
        exit 1
    }
} else {
    Write-Host "Error: $inputFile not found"
    exit 1
}

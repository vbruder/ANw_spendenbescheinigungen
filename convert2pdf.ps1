# Convert-DocxToPDF.ps1

# Get the script directory and construct paths
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$outputDir = Join-Path -Path $scriptDir -ChildPath "..\output"
$outputDir = (Resolve-Path $outputDir).Path
$pdfDir = Join-Path -Path $outputDir -ChildPath "pdf"

# Create PDF subfolder if it doesn't exist
if (-not (Test-Path $pdfDir)) {
    New-Item -ItemType Directory -Path $pdfDir
    Write-Host "Created PDF directory: $pdfDir"
}

# Create Word application object
$word = New-Object -ComObject Word.Application
$word.Visible = $false

# Counter for progress tracking
$totalFiles = (Get-ChildItem -Path $outputDir -Filter *.docx).Count
$converted = 0
$failed = @()

Write-Host "Starting conversion of $totalFiles files from $outputDir..."

# Process each DOCX file in the output directory
Get-ChildItem -Path $outputDir -Filter *.docx | ForEach-Object {
    $docPath = $_.FullName
    $pdfPath = Join-Path $pdfDir ($_.BaseName + ".pdf")
    
    Write-Host "Converting $($_.Name)..."
    
    try {
        $doc = $word.Documents.Open($docPath)
        Write-Host "Document opened successfully: $docPath"

        Write-Host "Saving as PDF: $pdfPath"
        $doc.SaveAs($pdfPath, 17) # 17 is the value for PDF format
        Write-Host "Saved successfully."

        # Close document
        $doc.Close()
        
        $converted++
        Write-Host "Successfully converted $($_.Name)" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to convert $($_.Name): $_" -ForegroundColor Red
        $failed += $_.Name
    }
}

# Clean up
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
Remove-Variable word

# Print summary
Write-Host "`nConversion complete!"
Write-Host "Successfully converted: $converted files"
Write-Host "Failed conversions: $($failed.Count) files"

if ($failed.Count -gt 0) {
    Write-Host "`nFailed files:"
    $failed | ForEach-Object { Write-Host "- $_" -ForegroundColor Red }
}

Write-Host "`nPDF files are saved in: $pdfDir"

# Wait for user input before closing
Write-Host "`nPress any key to exit..."
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
param(
    [string]$TestDocPath
)

$ErrorActionPreference = "Stop"

Write-Host "=== Word to PDF Conversion Test ==="
Write-Host ""

if (-not $TestDocPath) {
    Write-Host "Usage: test-word-conversion.ps1 -TestDocPath 'C:\path\to\test.docx'"
    exit 1
}

Write-Host "Input file: $TestDocPath"

if (-not (Test-Path $TestDocPath)) {
    Write-Host "ERROR: Input file not found!"
    exit 1
}

Write-Host "File exists: OK"

$OutputFolder = Join-Path $env:TEMP "word-test-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
Write-Host "Output folder: $OutputFolder"
Write-Host ""

try {
    Write-Host "Step 1: Creating Word COM object..."
    $word = New-Object -ComObject Word.Application
    Write-Host "  SUCCESS - Word version: $($word.Version)"

    Write-Host "Step 2: Setting Word properties..."
    $word.Visible = $false
    $word.DisplayAlerts = 0
    Write-Host "  SUCCESS"

    Write-Host "Step 3: Opening document..."
    $InputFile = Resolve-Path $TestDocPath
    Write-Host "  Absolute path: $InputFile"
    $doc = $word.Documents.Open($InputFile, $false, $true)
    Write-Host "  SUCCESS - Document opened"
    Write-Host "  Pages: $($doc.ComputeStatistics(2))"

    Write-Host "Step 4: Generating PDF path..."
    $pdfName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile) + ".pdf"
    $pdf = Join-Path $OutputFolder $pdfName
    Write-Host "  PDF will be saved to: $pdf"

    Write-Host "Step 5: Exporting to PDF..."
    $doc.ExportAsFixedFormat($pdf, 17)
    Write-Host "  SUCCESS - PDF exported"

    Write-Host "Step 6: Closing document..."
    $doc.Close($false)
    Write-Host "  SUCCESS"

    Write-Host "Step 7: Quitting Word..."
    $word.Quit()
    Write-Host "  SUCCESS"

    Write-Host "Step 8: Cleanup..."
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-Host "  SUCCESS"

    Write-Host ""
    Write-Host "Step 9: Verifying output..."
    if (Test-Path $pdf) {
        $pdfSize = (Get-Item $pdf).Length
        Write-Host "  SUCCESS - PDF created ($pdfSize bytes)"
        Write-Host ""
        Write-Host "=== CONVERSION SUCCESSFUL ==="
        Write-Host "PDF location: $pdf"
        Write-Output $pdf
    } else {
        Write-Host "  ERROR - PDF file not found!"
        exit 1
    }

} catch {
    Write-Host ""
    Write-Host "=== ERROR OCCURRED ==="
    Write-Host "Error message: $($_.Exception.Message)"
    Write-Host "Error type: $($_.Exception.GetType().FullName)"
    Write-Host ""
    Write-Host "Stack trace:"
    Write-Host $_.ScriptStackTrace

    # Try to cleanup
    try {
        if ($doc) { $doc.Close($false) }
        if ($word) { $word.Quit() }
    } catch {}

    exit 1
}

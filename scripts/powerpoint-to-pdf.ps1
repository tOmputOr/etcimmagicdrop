param(
  [string]$inputPath,
  [string]$outputPath
)

try {
    $inputPath = (Resolve-Path $inputPath).Path

    $powerpoint = New-Object -ComObject PowerPoint.Application
    $powerpoint.DisplayAlerts = 2  # ppAlertsNone

    $presentation = $powerpoint.Presentations.Open($inputPath, $false, $false, $false)

    # 32 = ppSaveAsPDF
    $presentation.SaveAs($outputPath, 32)

    $presentation.Close()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null

    $powerpoint.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint) | Out-Null

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Output "Success: PDF created at $outputPath"
}
catch {
    Write-Error "PowerPoint conversion failed: $_"
    exit 1
}

$ErrorActionPreference = "Stop"

Write-Host "Testing Word COM object availability..."

try {
    $word = New-Object -ComObject Word.Application
    Write-Host "SUCCESS: Word COM object created"
    Write-Host "Word Version: $($word.Version)"
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    Write-Host "Word is available and working"
} catch {
    Write-Host "ERROR: Word COM object not available"
    Write-Host "Error message: $($_.Exception.Message)"
    Write-Host ""
    Write-Host "This means Microsoft Word is either:"
    Write-Host "  1. Not installed on this system"
    Write-Host "  2. Not properly registered for COM automation"
    Write-Host ""
    Write-Host "To fix this, you need to install Microsoft Word/Office"
}

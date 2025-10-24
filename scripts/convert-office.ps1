param(
  [string]$InputFile,
  [string]$OutputFolder
)

$ErrorActionPreference = "Stop"

try {
  # Validate input
  if (-not (Test-Path $InputFile)) {
    throw "Input file not found: $InputFile"
  }

  $ext = [System.IO.Path]::GetExtension($InputFile).ToLower()
  if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
  }

  # Get absolute paths as strings
  $InputFile = (Resolve-Path $InputFile).Path
  $OutputFolder = (Resolve-Path $OutputFolder).Path

  switch ($ext) {
    ".doc" {
      $word = New-Object -ComObject Word.Application
      $word.Visible = $false
      $word.DisplayAlerts = 0
      $doc = $word.Documents.Open($InputFile, $false, $true)
      $pdf = Join-Path $OutputFolder ([System.IO.Path]::GetFileNameWithoutExtension($InputFile) + ".pdf")
      $doc.ExportAsFixedFormat($pdf, 17)
      $doc.Close($false)
      $word.Quit()
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
      [System.GC]::Collect()
      [System.GC]::WaitForPendingFinalizers()
      Write-Output $pdf
    }
    ".docx" {
      $word = $null
      $doc = $null
      try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $word.ScreenUpdating = $false
        $doc = $word.Documents.Open($InputFile, $false, $true)
        $pdf = Join-Path $OutputFolder ([System.IO.Path]::GetFileNameWithoutExtension($InputFile) + ".pdf")
        # Use ExportAsFixedFormat instead of SaveAs - no dialog
        $doc.ExportAsFixedFormat($pdf, 17)
        $doc.Close($false)
        $word.Quit($false)
        Write-Output $pdf
      } finally {
        if ($doc) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null }
        if ($word) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
      }
    }
    ".xls" {
      $excel = New-Object -ComObject Excel.Application
      $excel.Visible = $false
      $excel.DisplayAlerts = $false
      $wb = $excel.Workbooks.Open($InputFile)
      $pdf = Join-Path $OutputFolder ([System.IO.Path]::GetFileNameWithoutExtension($InputFile) + ".pdf")
      $wb.ExportAsFixedFormat(0, $pdf)
      $wb.Close($false)
      $excel.Quit()
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
      [System.GC]::Collect()
      [System.GC]::WaitForPendingFinalizers()
      Write-Output $pdf
    }
    ".xlsx" {
      $excel = New-Object -ComObject Excel.Application
      $excel.Visible = $false
      $excel.DisplayAlerts = $false
      $wb = $excel.Workbooks.Open($InputFile)
      $pdf = Join-Path $OutputFolder ([System.IO.Path]::GetFileNameWithoutExtension($InputFile) + ".pdf")
      $wb.ExportAsFixedFormat(0, $pdf)
      $wb.Close($false)
      $excel.Quit()
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
      [System.GC]::Collect()
      [System.GC]::WaitForPendingFinalizers()
      Write-Output $pdf
    }
    ".ppt" {
      $ppt = New-Object -ComObject PowerPoint.Application
      $pres = $ppt.Presentations.Open($InputFile, $false, $false, $false)
      $pngDir = Join-Path $OutputFolder ([System.IO.Path]::GetFileNameWithoutExtension($InputFile))
      $pres.SaveAs($pngDir, 18)
      $pres.Close()
      $ppt.Quit()
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) | Out-Null
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
      [System.GC]::Collect()
      [System.GC]::WaitForPendingFinalizers()
      Write-Output $pngDir
    }
    ".pptx" {
      $ppt = New-Object -ComObject PowerPoint.Application
      $pres = $ppt.Presentations.Open($InputFile, $false, $false, $false)
      $pngDir = Join-Path $OutputFolder ([System.IO.Path]::GetFileNameWithoutExtension($InputFile))
      $pres.SaveAs($pngDir, 18)
      $pres.Close()
      $ppt.Quit()
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) | Out-Null
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
      [System.GC]::Collect()
      [System.GC]::WaitForPendingFinalizers()
      Write-Output $pngDir
    }
    default {
      throw "Unsupported file type: $ext"
    }
  }
} catch {
  Write-Error $_.Exception.Message
  exit 1
}

param(
  [string]$inputPath,
  [string]$outputPath
)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($inputPath, $false, $true)
$doc.ExportAsFixedFormat($outputPath, 17)  # 17 = wdExportFormatPDF
$doc.Close()
$word.Quit()

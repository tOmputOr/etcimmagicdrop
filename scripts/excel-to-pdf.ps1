param(
  [string]$inputPath,
  [string]$outputPath
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Open($inputPath)
$workbook.ExportAsFixedFormat(0, $outputPath)  # 0 = xlTypePDF
$workbook.Close($false)
$excel.Quit()

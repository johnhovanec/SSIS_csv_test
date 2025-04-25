$excelPath = "C:\Users\jhovanec\Downloads\Asthma ED and IP 2010_2021.xlsx"

# param (
#    [string]$excelPath
#)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$workbook = $excel.Workbooks.Open($excelPath)

$outputFolder = Split-Path $excelPath
$filenameBase = [System.IO.Path]::GetFileNameWithoutExtension($excelPath)

for ($i = 1; $i -le $workbook.Sheets.Count; $i++) {
    $sheet = $workbook.Sheets.Item($i)
    $sheetName = $sheet.Name.Replace(" ", "_")  # Clean up for filename

    $outputPath = Join-Path $outputFolder "$filenameBase-$sheetName.csv"
    $sheet.SaveAs($outputPath, 6)  # 6 = xlCSV
}

$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel
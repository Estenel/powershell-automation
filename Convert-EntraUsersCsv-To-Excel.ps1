$CsvFile = "C:\Temp\users.csv"
$XlsxFile = "C:\Temp\users.xlsx"

$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false

$Workbook = $Excel.Workbooks.Add()
$Worksheet = $Workbook.Worksheets.Item(1)

# vytvoření importu CSV
$Connection = "TEXT;$CsvFile"

$Query = $Worksheet.QueryTables.Add(
    $Connection,
    $Worksheet.Range("A1")
)

$Query.TextFileParseType = 1          # delimited
$Query.TextFileCommaDelimiter = $true # oddělovač čárka
$Query.TextFilePlatform = 65001       # UTF-8
$Query.TextFileColumnDataTypes = @(2)

$Query.Refresh()

# odstranění dotazu
$Query.Delete()

# šířka sloupců
$Worksheet.Columns.AutoFit()

# uložení XLSX
$Workbook.SaveAs($XlsxFile,51)

$Workbook.Close($false)
$Excel.Quit()

# uvolnění COM
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Worksheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null

Write-Host "Hotovo: $XlsxFile"

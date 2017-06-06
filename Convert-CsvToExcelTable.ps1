<#
.SYNOPSIS
Converts a CSV to a single-sheet Excel workbook formatted as a table

.DESCRIPTION
Takes a CSV file and returns an Excel spreadsheet formatted as a table in the style specified.

.EXAMPLE
.\Convert-CsvToExcelTable.ps1 -csvFilePath 'mycsv.csv' -excelFilePath 'myspreadsheet.xlsx'

.LINK
https://flamingkeys.com/BLOG-POST-SLUG

.NOTES
Written by: Chris Brown

Find me on:
* My blog: https://flamingkeys.com/
* Github: https://github.com/chrisbrownie

#>

Param(
    [Parameter(Mandatory)][string]$csvFilePath,
    [Parameter(Mandatory)][string]$excelFilePath,
    [Parameter][string]$TableStyle = "TableStyleLight9"
    )

# Convert CSV File Path to the absolute path
$csvFilePath = Convert-Path $csvFilePath
# Convert XLSX path to the absolute path if it's relative
if (-not [System.IO.Path]::IsPathRooted($excelFilePath)) {
    $excelFilePath = Join-Path -Path (Get-Location) -ChildPath $excelFilePath
}

$excelApplication = New-Object -ComObject Excel.Application -Verbose:$false
$excelApplication.DisplayAlerts = $false
$Workbook = $excelApplication.Workbooks.Open($csvFilePath)

# It was a CSV, so only one sheet
$sheet = $Workbook.Worksheets[1]

# Create a range that refers to all cells from A1 to the end in both row and columns
$A1range = $sheet.Range("A1")
$range = $sheet.Range($A1range,$A1range.SpecialCells([Microsoft.Office.Interop.Excel.XlCellType]::xlCellTypeLastCell))

$table = $sheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,$range, $null , [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)

$table.TableStyle = $TableStyle


# Save the file out as a spreadsheet
$Workbook.SaveAs($excelFilePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook)

"Saved '$excelFilePath'"

try { 
    $ExcelApplication.Workbooks.Close() 
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApplication)
} catch { }


& {
    $sourceFile = "C:\Users\mimor\Downloads\C331-A13_XChem_Blue_Thumb_Drive_Formatted (files highlighted).xlsx.xlsx"

    $outputFile = "C:\Users\mimor\Downloads\Output.txt"

    $startRow = 1

    $startColumn = 5

    $usedCellType = 11
    
    $excelApp = New-Object -ComObject Excel.Application 

    try {
        $excelApp.visible = $false
        $excelApp.DisplayAlerts = $false 

        $workbook = $excelApp.Workbooks.Open($sourceFile) 
        $worksheet = $workbook.WorkSheets.item("Properties")
        $endRow = $worksheet.UsedRange.SpecialCells($usedCellType).Row

        $rangeAddress = $worksheet.Cells.Item($startRow, $startColumn).Address() + ":" + $worksheet.Cells.Item($endRow, $startColumn).Address()
        Write-Host "Using range $($rangeAddress)"

        $worksheet.Range($rangeAddress).Value2 | Out-File -FilePath $outputFile
        $workbook.Close($false) 
    }
    finally {
        $excelApp.Quit()
    }
}



    

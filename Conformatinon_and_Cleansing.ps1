& {
    ################################################################################################
    <#
    .NOTES
        This project was managed by Matthew Morales with LCG, LLC. Documentation can be found at 
        https://www.notion.so/5dab2cb8ac194d64b657183dd40e1c36?v=128e61587acd4430b2910a783f2e8e07
    .SYNOPSIS
        Utilize the excel spreadsheet containing the Paths for deletion to extract a usable .txt file.
    .DESCRIPTION
        PART 1.
        The User creates variables for the Source File & Output file location. The user can specify 
        which Row, or Column to extract the data from. The script will then call Excel to open as an 
        Object within powershell. The script will then iterate through a try block to handle extraction
        within the background. The user specifies which column that is needed to extract from within
        the worksheet function, promptly naming it after the necessary sheet name. The rangeAddress
        function then returns a range object that represents all the cells on the worksheet, not just 
        the current one in use. The worksheet.Range function returns object that represents a cell or
        a range of cells. The script will the close excel after proper extraction.
        
        PART 2.
        Following the close of the excel file,the script will then utilize the powershell terminal to 
        check if the files listed within the excel sheet created are present within the host device. 
        This is done by first confirming if the user would like to continue with the script then
        creating a foreach loop to run a get-content function against the path where text file from 
        extraction is. The loop will list if the path could be found or not. From this output, the 
        user can determine which files were not discovered, and in turn would not be accessible to Part 
        3, implementation for uniform script will be researched at a later date.
        
        PART 3.

    #>
    ################################################################################################# 

    # BE SURE TO CHANGE THE SOURCE FILE
    $sourceFile = "C:\Users\mmorales\Documents\C331-A13_XChem_Blue_Thumb_Drive.xlsx"

    # BE SURE TO CHANGE THE OUTPUT FILE
    $outputFile = "C:\Users\mmorales\Output.txt"

    $startRow = 2

    $startColumn = 5

    $usedCellType = 11
    
    $excelApp = New-Object -ComObject Excel.Application 

    try {
        $excelApp.visible = $false
        $excelApp.DisplayAlerts = $false 

        $workbook = $excelApp.Workbooks.Open($sourceFile) 
        $worksheet = $workbook.WorkSheets("Sheet")
        $endRow = $worksheet.UsedRange.SpecialCells($usedCellType).Row

        $rangeAddress = $worksheet.Cells.Item($startRow, $startColumn).Address() + ":" + $worksheet.Cells.Item($endRow, $startColumn).Address()
        Write-Host "Using range $($rangeAddress)"

        $worksheet.Range($rangeAddress).Value2 | Out-File -FilePath $outputFile
        $workbook.Close($false) 
    }
    finally {
        $excelApp.Quit()
    }

<#
  Shelved function to extract paths to export in csv. Shelved for later production.

{
    $path = 'C:\Users\mmorales\Documents\TestingPaths.txt'

    $files = Get-Content $path *.* -Recurse

    $data = foreach ($file in $files) {
        foreach ($pathinfo in (get-content $file.fullname)) {
            [PSCustomObject]@{
                Foldername = $file.directory
                FileName   = $file.FileName
                Reference  = $pathinfo
                Exists     = (test-path $pathinfo)
           }
            
        }
        }
    }

$data | export-csv "C:\Users\mmorales\Github Projects\PathCleansing.csv" -NoTypeInformation
}
#>
    #Ensure the Path is correct 
    Foreach ($path in Get-Content C:\Users\mmorales\TestingPaths.txt) {
    [PSCustomObject]@{
         Path   = $path
         Exists = Test-Path $path
    }
}

    write-host -nonewline "Continue? (Y/N) "
    $response = read-host
    if ( $response -ne "Y" ) { exit }

    Get-Content "C:\Users\mmorales\TestingPaths.txt" | ForEach-Object { Move-Item -Path $_ -Destination "C:\Users\mmorales\Purgeable Folder" -Verbose}
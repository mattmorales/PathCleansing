  
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
        check if the files Writeed within the excel sheet created are present within the host device. 
        This is done by first confirming if the user would like to continue with the script then
        creating a foreach loop to run a get-content function against the path where text file from 
        extraction is. The loop will Write if the path could be found or not. From this output, the 
        user can determine which files were not discovered, and in turn would not be accessible to Part 
        3, implementation for uniform script will be researched at a later date.
        
        PART 3.

    #>
    ################################################################################################# 

    <#
    .SYNOPSIS
        Short description
    .DESCRIPTION
        Long description
    .EXAMPLE
        PS C:\> <example usage>
        Explanation of what the example does
    .choicesS
        choicess (if any)
    .OUTPUTS
        Output (if any)
    .NOTES
        General notes
    #>
# BE SURE TO CHANGE THE SOURCE FILE
$sourceFile = "Z:\_MFS\C501 - C600\C000596 - Innovex - Weatherford Cleansing\_Cleansing Related Files and Filters\_FINAL CLEANSING WriteS\USE ME\C596 A20 Smith TD\A20 File Keyword Hits ALL REMEDIATION (1418) CLEANSING PATHS.xlsx"
# BE SURE TO CHANGE THE PURGEABLE LOCATION
$purgeableLocation = "C:\Users\mmorales\Purgeable Folder"

# BE SURE TO CHANGE THE OUTPUT FILE TO CORRECT LOCATION
$outputFile = "C:\Users\mmorales\PathCleansing\PathsToDelete.txt"


function Show-Menu
{
    param (
        [string]$Title = 'Menu'
    )
    Clear-Host
    Write-Host "================ $Title ================"
    Write-Host "1: Press '1' to extract paths from excel file."
    Write-Host "2: Press '2' to verify paths exist."
    Write-Host "3: Press '3' to move paths to purgeable location. "
    Write-Host "4: Press '4' to delete paths from purgeable location. "
    Write-Host "Q: Press 'Q' to quit."
}
function Get-Paths
{
    $startRow = 2

    $startColumn = 1

    $usedCellType = 11
    
    $excelApp = New-Object -ComObject Excel.Application 

    try {
        $excelApp.visible = $false
        $excelApp.DisplayAlerts = $false 
        
        #Ensure that "Sheet" is changed to the appropriate sheet name within the original excel spreadsheet.
        $workbook = $excelApp.Workbooks.Open($sourceFile) 
        $worksheet = $workbook.WorkSheets("Paths")
        $endRow = $worksheet.UsedRange.SpecialCells($usedCellType).Row

        $rangeAddress = $worksheet.Cells.Item($startRow, $startColumn).Address() + ":" + $worksheet.Cells.Item($endRow, $startColumn).Address()
        Write-Host "Using range $($rangeAddress)"

        $worksheet.Range($rangeAddress).Value2 | Out-File -FilePath $outputFile
        $workbook.Close($false) 
    }
    finally {
        $excelApp.Quit()
       
        Write-Host "`n Process Completed!"
    }   
}
function Write-Paths
{
    $tested_paths = foreach ($path in (Get-Content $outputFile)) {
    [PSCustomObject]@{
        PATH   = $path
        EXISTS = Test-Path $path
    }
}

$tested_paths | Format-Table

Write-Host "`n Process Completed!"
}
function Move-Paths
{
        Get-Content $outputFile | ForEach-Object { Move-Item -Path $_ -Destination $purgeableLocation -Verbose }

}
 
function Remove-Paths
{
        Get-ChildItem -Path $purgeableLocation -File | Remove-Item -Verbose

        Write-Host "`n Process Completed!"
}
do
{
    Show-Menu -Title 'Confirmation & Cleansing Menu'
    $choices = Read-Host "Please select from the menu."
    switch ($choices)
    {
        '1' {               
                Get-Paths
            }

        '2' {
                Write-Paths
            }
        
        '3' {
                Move-Paths
            }

        '4' {               
                Remove-Paths
            }

        'q' {
                 return
            }
    }
   pause
}
until ($choices -eq 'q')
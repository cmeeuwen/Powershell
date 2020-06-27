clear

$erroractionpreference = "silentlycontinue"

# Load Excel Comobject
#=====================#
$xlDotNet = New-Object -ComObject excel.application

# Path to spreadsheet
#====================#
$XLXSFile = "C:\SuperscriptProd\OnboardTemplates\CMO.csv"


# Open Workbook
#==============#
$XLOpenWB =  $xlDotNet.Workbooks.Open($XLXSFile)
$XLOpenWS = $xlDotNet.Worksheets.Item("CMO")

# Generates an array containing the Alphabetical Characters for column reference
#===============================================================================#
$ColumnletterArray = New-Object System.Collections.ArrayList

for ($test = 0; $test -lt 26; $test++)
{
    [char](65 + $test) | ForEach-Object {$ColumnletterArray.Add($_) | Out-Null}
}

# Generates Cell count into an array
#===================================#
$CellCountArray = New-Object System.Collections.ArrayList

           $xlCellCount = [int32]$XLOpenWS.UsedRange.Cells.Count
           $xlCellCount++
                                                         
           Do {

               $xlCellCount-- | Out-Null
               #Write-Host "$($XLColumnCount) is the column Count"
               $CellCountArray.add($xlCellCount) | Out-Null


              } until ($xlCellCount -eq 0)

           ForEach ($Column in $ColumnletterArray){
                                                         
                   $Column | Out-Null

                   # For Each Loop to store cell values into a variable
                   foreach ($Cell in $CellCountArray) {
                                                                                             
                           $findCellValue = $XLOpenWS.Range("$($Column)$($Cell)").Text

                           $ColumnCell = $Column + $Cell
                                                      
                           # If cell value contains specific string and cell number id D14 output to host
                           If ($findCellValue -eq "Operations" -and $ColumnCell -eq "D14") {
                                                                                                                                   
                              Write-Host "Cell: $($Column)$($Cell) contains the value of $($findCellValue)!" -ForegroundColor Green
                                                         
                                                                #Insert code here
                                                                #================
                                                                                                                                   
                                                                                                                                   
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                               }

                                                                                                                                   
                                                                                             
                                                         
                                                      }
                                                         
                                                         
                                                         

                                                         
          
                                                  }


$XLOpenWB.close()
$xlDotNet.Quit()
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($XLOpenWS)
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($XLOpenWB)
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xlDotNet)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
$xlDotNet = $null
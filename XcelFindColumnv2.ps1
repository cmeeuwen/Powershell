clear

Function CleanUP {

$XLOpenWB.close()
$xlDotNet.Quit()
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($XLOpenWS)
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($XLOpenWB)
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xlDotNet)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
$xlDotNet = $null

}

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

for ($num = 0; $num -lt 26; $num++)
{
    [char](65 + $num) | ForEach-Object {$ColumnletterArray.Add($_) | Out-Null}
}


# Generates Array Object for cell count
#======================================#
$CellCountArray = New-Object System.Collections.ArrayList


# Stores Cell Count into variable
#================================#
$xlCellCount = [int32]$XLOpenWS.UsedRange.Cells.Count
$xlCellCount++

# Parses out Cell count out and adds to array
#============================================#                       
Do {

    $xlCellCount-- | Out-Null
    #Write-Host "$($XLColumnCount) is the column Count"
    $CellCountArray.add($xlCellCount) | Out-Null


   } until ($xlCellCount -eq 0)

   ForEach ($Column in $ColumnletterArray){
                                                         
           $Column | Out-Null

           # For Each Loop to look up cell values 
           #=====================================#
           ForEach ($Cell in $CellCountArray) {
                                                                                             
                   $findCellValue = $XLOpenWS.Range("$($Column)$($Cell)").Text

                   $ColumnCell = $Column + $Cell
                                                      
                   # If cell value contains specific string and cell number is D14 output to host
                   If ($findCellValue -eq "Operations" -and ($ColumnCell -eq "D14" -or $ColumnCell -eq "L14")) {
                                                                                                                                   
                      Write-Host "Cell: $($Column)$($Cell) contains the value of $($findCellValue)!" -ForegroundColor Green
                                                         
                                                                                    #Insert code here
                                                                                    #================
                                                                                    
                      Write-Host "Replacing value in cell $($ColumnCell) with Test"                                               
                      [void]$XLOpenWS.Range("$ColumnCell", "$ColumnCell").Replace("Operations", "Test", [Microsoft.Office.Interop.Excel.XlLookAt]::xlPart)
                                                                                    
                      $XLOpenWB.save()                                             
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                           
                                                                                   CleanUP        
                                                                                   Break Script        
                                                                                   }

                                                                                                                                   
                                                                                             
                                                         
                                                }
                                                         
                                                         
                                                         

                                                         
          
                                            }

CleanUP

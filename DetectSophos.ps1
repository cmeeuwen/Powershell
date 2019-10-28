Write-Host "
             Check if Sophos is installed on remote systems
==========================================================================
| Synopsis: This script will check all systems within a domain to verify |
| that sophos antivirus is installed.                                    |
|                                                                        |
| If the system is offline or does not have sophos installed the script  |
| will append to a text file containing the hostname and whether that    |
| system was online during the scan or if it doesnt have Sophos          |
| installed at all.                                                      |
|                                                                        |
| If the system has sophos installed the script will append to another   |
| text file containing the hostname and the version of sophos that is    |
| installed                                                              |
|                                                                        |
| Path to text files: \\prlfs01\iswizards\superscript\results\           |
|                                                                        |
| Result text file names: SophosNotInstalled, SophosInstalled,           |
|                         NotOnlineDuringScan                            |
|                                                                        |
==========================================================================
" -ForegroundColor Yellow

# Query a list of all computers under the AD OU OU=WSUS_Desktops,OU=Ruby Computers,DC=ruby,DC=local and store in the variable $getComps
$getComps = Get-ADComputer -SearchBase 'OU=WSUS_Desktops,OU=Ruby Computers,DC=ruby,DC=local' -Filter 'ObjectClass -eq "Computer"' | 
            Select-Object name | Sort-Object -Property name

$DoResultsExist = Get-ChildItem -Path \\prlfs01\ISWizards\SuperScript\Results\*.txt

    if ($DoResultsExist -ne $null) {
                                    Write-Host "Removing previous scan results.`n" -ForegroundColor Yellow
                                    Remove-Item -Path \\prlfs01\ISWizards\SuperScript\Results\*.txt -Force -Verbose
                                    }
                                    else {
                                          Write-Host "No Previous Results Exist.`n" -ForegroundColor Yellow 
                                          }



# foreach loop for all computers found in AD under the specified OU
foreach ($comp in $getComps) {

    $ErrorActionPreference = "0"
    
                              
    Write-Host "`nChecking connection for $($comp.name)" -ForegroundColor Yellow
                              
    # Test to see if system is online / on the network
    $issystemonline = Test-Connection -ComputerName $comp.name | Select-Object -Property IPV4Address

    # Statement (If systems is online check to see if sophos is installed if not output error to text file with name of that system)
    if ($issystemonline -ne $null) {
       $issophosinstallted = Get-WmiObject -Class win32_product -ComputerName $comp.name | 
       Select-Object name, version | Where-Object -FilterScript {$_.Name -like "Sophos Anti*"} |
       ft -AutoSize
       
                   # Nested If Else statement for if sophos is found / not found on system 
                   if ($issophosinstallted -ne $null) {
                       Write-Host "Sophos is installed!" -ForegroundColor Green
                       $issophosinstallted
                       "Sophos is installed on $($comp.name)`n" |
                       Out-File -FilePath \\prlfs01\ISWizards\SuperScript\Results\SophosInstalled.txt -Append
                       } 
                       else {
                             Write-Host "Sophos not found on $($comp.name)`n" -ForegroundColor Red
                             "Sophos not found on $($comp.name)" | 
                             Out-File -FilePath \\prlfs01\ISWizards\SuperScript\Results\SophosNotInstalled.txt -Append
                             }
                                    }
                                     else {
                                           Write-Host "$($comp.name) is not online! `n" -ForegroundColor Red
                                           "$($comp.name) is not online!" | Out-File -FilePath \\prlfs01\ISWizards\SuperScript\Results\NotOnlineDuringScan.txt -Append
                                           }
                                 
}

Write-Host "Scan Completed!" -ForegroundColor Green
pause
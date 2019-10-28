function Assigno365Licenses {

Write-Host "
                Assign O365 licenses to new Receptionists
============================================================================
 
 Synopsis: This script will assign the following
 o365 licenses to receptionists

 Microsoft 365 F1                              
 Office 365 Advanced Threat Protection (Plan1)

     SkuPartNumber                            Translation
+====================+   +================================================+
|    VISIOCLIENT     |   |               Visio Online Plan 2              |
| ENTERPRISEPREMIUM  |   |                  Office 365 E5                 |
|   WINDOWS_STORE    |   |           Windows Store for Business           |
|  ENTERPRISEPACK    |   |                  Office 365 E3                 |
|     FLOW_FREE      |   |               Microsoft Flow Free              |
|  EXCHANGESTANDARD  |   |             Exchange Online (Plan 1)           |
|    MS_TEAMS_IW     |   |              Microsoft Teams Trial             |
| POWER_BI_STANDARD  |   |                 Power BI (free)                |
| OFFICESUBSCRIPTION |   |               Office 365 ProPlus               |
|     MCOPSTNC       | = |              Communications Credits            |
|        EMS         |   |         Enterprise Mobility + Security E3      |
|    MCOMEETADV      |   |               Audio Conferencing               |
|    AAD_PREMIUM     |   |         Azure Active Directory Premium P1      |
|      SPE_E3        |   |                Microsoft 365 E3                |
|PROJECTPROFESSIONAL |   |           Project Online Professional          |
|      SPE_F1        |   |                Microsoft 365 F1                |
|   ATP_ENTERPRISE   |   | Office 365 Advanced Threat Protection (Plan 1) |
|   O365_BUSINESS    |   |               Office 365 Business              |
|   STANDARDPACK     |   |                 Office 365 E1                  |
+====================+   +================================================+

============================================================================
" -ForegroundColor Yellow

$o365upn = Read-Host -Prompt "Please provide Users UPN (ex: firstname.lastname@callruby.com)"

# Verifies that the AzureAD and MsolService Modules are installed
$isAZUREADthere = Get-Module -Name AzureAD
    
    # If Modules exist, Connect to AzureAD and MsolService
    If ($isAZUREADthere -ne $null) {
                                    Write-Host "`nAzureAD and MSonline Modules installed already!" -ForegroundColor Green
                                    Write-Host "`nConnecting to AzureAD...." -ForegroundColor Yellow
                                    Connect-AzureAD 
                                    Write-Host "`nConnecting to MSOonline...." -ForegroundColor Yellow
                                    Connect-MsolService 

                                    $doesO365userExist = Get-MsolUser -UserPrincipalName $o365upn

                                    # If user exists build .net form and query available licenses to assign
                                    if ($doesO365userExist -ne $null) {
                                                                       
                                    Add-Type -AssemblyName System.Windows.Forms
                                    Add-Type -AssemblyName System.Drawing

                
                                    $form = New-Object System.Windows.Forms.Form
                                    $form.Text = 'o365 Licenses'
                                    $form.Size = New-Object System.Drawing.Size(400,300)
                                    $form.StartPosition = 'CenterScreen'
                

                                    $OKButton = New-Object System.Windows.Forms.Button
                                    $OKButton.Location = New-Object System.Drawing.Point(120,220)
                                    $OKButton.Size = New-Object System.Drawing.Size(75,23)
                                    $OKButton.Text = 'OK'
                                    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
                                    $form.AcceptButton = $OKButton
                                    $form.AutoSize = $true
                                    $form.Controls.Add($OKButton)

                                    $CancelButton = New-Object System.Windows.Forms.Button
                                    $CancelButton.Location = New-Object System.Drawing.Point(200,220)
                                    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
                                    $CancelButton.Text = 'Cancel'
                                    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                                    $form.CancelButton = $CancelButton
                                    $form.Controls.Add($CancelButton)

                                    $label = New-Object System.Windows.Forms.Label
                                    $label.Location = New-Object System.Drawing.Point(10,20)
                                    $label.Size = New-Object System.Drawing.Size(280,20)
                                    $label.Text = 'Please Select licenses:'
                                    $form.Controls.Add($label)

                                    $listBox = New-Object System.Windows.Forms.Listbox
                                    $listBox.Location = New-Object System.Drawing.Point(15,60)
                                    $listBox.Size = New-Object System.Drawing.Size(350,500)
                

                                    $listBox.SelectionMode = 'MultiSimple'


                                    $geto365licenses = Get-AzureADSubscribedSku | select skupartnumber
                

                                    foreach ($license in $geto365licenses) {
                                                [void] $listBox.Items.Add($license.skupartnumber)
                                                }
                

                                    $listBox.Height = 150
                                    $form.Controls.Add($listBox)
                                    $form.Topmost = $true

                                    $result = $form.ShowDialog()

                                        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {Write-Host "Selected licenses assigned!" -ForegroundColor Green} else {Write-Host "Cancelled! No licenses assigned!" -ForegroundColor Red}
                                            ForEach-Object {

                                                                foreach ($license in $listBox.SelectedItems)
                                                           
                                                                {
                                                                    Try {
                                                                         Set-MsolUserLicense -UserPrincipalName $o365upn -AddLicenses ('callruby:' + $license)
                                                                         }
                                                                         catch {Write-Host "$($o365upn) is already assigned the following license: $($license)" -ForegroundColor Red}
                                             
                                                                }
                                                            }

Write-Host "
=================================================


Licenses have been assigned to $($o365upn)                               
                                               
                                                                                     
=================================================
" -ForegroundColor Yellow


                                                                       }
                                                                       else {
                                                                             Write-Host "$($o365upn) does not exist. No licenses assigned" -ForegroundColor Red
                                                                             }

                                    
                                    }
                                     # If modules do not exist, Install then connect, Query licenses and assign
                                     else {
                                           Write-Host "`nInstalling AzureAD Module...." -ForegroundColor Yellow
                                           Install-Module AzureAD -Force
                                           Write-Host "`nInstalling MSOonline Module...." -ForegroundColor Yellow
                                           Install-Module MSOnline -Force

                                           Write-Host "`nConnecting to AzureAD...." -ForegroundColor Yellow
                                           Connect-AzureAD
                                           Write-Host "`nConnecting to MSOonline...." -ForegroundColor Yellow
                                           Connect-MsolService 

                                           $doesO365userExist = Get-MsolUser -UserPrincipalName $o365upn
                                           
                                           if ($doesO365userExist -ne $null) {
                                                                      
                                           Add-Type -AssemblyName System.Windows.Forms
                                           Add-Type -AssemblyName System.Drawing

                
                                    $form = New-Object System.Windows.Forms.Form
                                    $form.Text = 'o365 Licenses'
                                    $form.Size = New-Object System.Drawing.Size(400,300)
                                    $form.StartPosition = 'CenterScreen'
                

                                    $OKButton = New-Object System.Windows.Forms.Button
                                    $OKButton.Location = New-Object System.Drawing.Point(120,220)
                                    $OKButton.Size = New-Object System.Drawing.Size(75,23)
                                    $OKButton.Text = 'OK'
                                    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
                                    $form.AcceptButton = $OKButton
                                    $form.AutoSize = $true
                                    $form.Controls.Add($OKButton)

                                    $CancelButton = New-Object System.Windows.Forms.Button
                                    $CancelButton.Location = New-Object System.Drawing.Point(200,220)
                                    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
                                    $CancelButton.Text = 'Cancel'
                                    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                                    $form.CancelButton = $CancelButton
                                    $form.Controls.Add($CancelButton)

                                    $label = New-Object System.Windows.Forms.Label
                                    $label.Location = New-Object System.Drawing.Point(10,20)
                                    $label.Size = New-Object System.Drawing.Size(280,20)
                                    $label.Text = 'Please Select licenses:'
                                    $form.Controls.Add($label)

                                    $listBox = New-Object System.Windows.Forms.Listbox
                                    $listBox.Location = New-Object System.Drawing.Point(15,60)
                                    $listBox.Size = New-Object System.Drawing.Size(350,500)
                

                                    $listBox.SelectionMode = 'MultiSimple'


                                    $geto365licenses = Get-AzureADSubscribedSku | select skupartnumber
                

                                    foreach ($license in $geto365licenses) {
                                                [void] $listBox.Items.Add($license.skupartnumber)
                                                }
                

                                    $listBox.Height = 150
                                    $form.Controls.Add($listBox)
                                    $form.Topmost = $true

                                    $result = $form.ShowDialog()

                                        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {Write-Host "Selected licenses assigned!" -ForegroundColor Green} else {Write-Host "Cancelled! No licenses assigned!" -ForegroundColor Red}
                                            ForEach-Object {

                                                                foreach ($license in $listBox.SelectedItems)
                                                           
                                                                {
                                                                 try {
                                                                      Set-MsolUserLicense -UserPrincipalName $o365upn -AddLicenses ('callruby:' + $license) 
                                                                     }
                                                                      catch {Write-Host "$($o365upn) is already assigned the following license: $($license)" -ForegroundColor Red}
                                                                }
                                                            }

Write-Host "
=================================================


Licenses have been assigned to $($o365upn)                               
                                               
                                                                                     
=================================================
" -ForegroundColor Yellow


                                                                       }
                                                                       else {
                                                                             Write-Host "$($o365upn) does not exist. No licenses assigned" -ForegroundColor Red
                                                                             }
                                           }
}
Assigno365Licenses
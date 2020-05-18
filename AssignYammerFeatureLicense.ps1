Function AssignYammerFeatureLicense {

$ErrorActionPreference = 'SilentlyContinue'


[void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
[void][reflection.assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
[void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][reflection.assembly]::Load('System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][reflection.assembly]::Load('System.ServiceProcess, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')



# Connect to MSOnline - Check if module exists, Installs Module if missing
#=========================================================================
$DoesMSOlineModuleExist = Get-Module -Name MSOnline

if ($DoesMSOlineModuleExist -ne $null) {
                                        Connect-MsolService
                                       }

   Else {
         Write-Host "MSOline Module missing, Installing......" -ForegroundColor Yellow
         Install-Module -Name MSOline -Verbose -Force -ErrorAction SilentlyContinue
         Connect-MsolService
        }

#Import CSV for receps being updated
#=================================== 
$importcsvFile = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
           InitialDirectory = '\\prlfs01\ISWizards\SuperScript\Tools\MassAssignYammerLicenses' 
           Filter = 'CSV (*.csv)|*.csv|SpreadSheet (*.xlsx)|*.xlsx'
           }
           $importcsvFile.ShowDialog()

Import-Csv $importcsvFile.FileName | ForEach-Object {


$UPN = $_.UserPrincipalName
$LicenseDetails = (Get-MsolUser -UserPrincipalName $UPN).Licenses

ForEach ($License in $LicenseDetails) {
                                       $DisabledOptions = @()
                                       $License.ServiceStatus | ForEach {
     
     If ($_.ProvisioningStatus -eq "Disabled" -and $_.ServicePlan.ServiceName -notlike "*YAMMER*") { 
                                                                                                    
                                                                                                    $DisabledOptions += "$($_.ServicePlan.ServiceName)" 
                                                                                                    
                                                                                                    } 
                                                                        }
   
   $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $License.AccountSkuId -DisabledPlans $DisabledOptions
   
   Set-MsolUserLicense -UserPrincipalName $UPN -LicenseOptions $LicenseOptions

   

}
Write-Host "Yammer Feature License Assigned for $($UPN)!" -ForegroundColor Green
}

}

AssignYammerFeatureLicense
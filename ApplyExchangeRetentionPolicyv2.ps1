Function AssignRetentionPolicies {

$ErrorActionPreference = 'silentlycontinue'

clear

Write-Host """
+====================================================+
|                                                    |
|       Retention Policy Assignment Script           |
|                                                    |
| Steps:                                             |
|                                                    |
|      1. Provide UPN credentials at prompt          |
|                                                    |
| Notes: This script will apply retention policies   |
|        for Exchange and Teams. This script should  |
|        be run every 30 days.                       |
|                                                    |
|                                                    |
+====================================================+
""" -ForegroundColor Green


# Provide Admin Credentials
#==========================
$LoginCredentials = Get-Credential


# Connection URI variables
#=========================
$EXOnlineURI = 'https://outlook.office365.com/powershell-liveid/'
$ComplianceURI = 'https://ps.compliance.protection.outlook.com/powershell-liveid/'


# Retention Policy Variables
#===========================
$MailboxRetentionPolicy = "Ruby - F3 User Remove/Delete emails that are 1 year and older"
$TeamsRetentionPolicy = "Ruby - Teams - 1:1 Chat F1  Standard Retention Policy"



# Query AD users based on Criteria
#=================================
$QueryUsers = get-aduser -Filter * -Properties title,department,enabled,employeeid,emailaddress,manager,PasswordLastSet,PasswordExpired,DistinguishedName,Created,UserPrincipalName,SamAccountName |
              Where {$_.title -notlike $NUll -and $_.department -notlike "$null" -and $_.manager -ne $null -and $_.employeeid -ne $Null -and $_.Enabled -eq $true}

$FrontLineEmployees = ($QueryUsers | where {$_.title -like "Chat Specialists (*)" -or $_.title -like "* Receptionist" })


# Verify AzureAD Module exists and connect
#=======================================================
$DoesAzureADModuleExist = Get-Module -Name AzureAD

if ($DoesAzureADModuleExist -ne $null) {
                                        Connect-AzureAD
                                       }

   Else {
         Write-Host "AzureAD Module missing, Installing......" -ForegroundColor Yellow
         Install-Module -Name AzureAD -Verbose -Force
         Connect-AzureAD
        }



# Exchange Online PS Session initializaion
#=========================================
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $EXOnlineURI -Credential $LoginCredentials -Authentication Basic -AllowRedirection
Import-PSSession $Session

$ErrorActionPreference = 'silentlycontinue'

foreach ($FrontEmp in $FrontLineEmployees){

# Font Color for template
#========================
$RedFontWarning = "color:red"

# Email HTMLBody
#===============
$emailBody = "
<p style=$($RedFontWarning)>Retention policy applied for  $($FrontEmp.SamAccountName)
<br>
<p><b>Qualifying Conditions:</b> User is a Frontline Employee and assigned an F3 license
<br>
<p>Applied the following Retention Policies: 
<br>
<p><b>Exchange:</b> Ruby - F3 User Remove/Delete emails that are 1 year and older
<br>
<p><b>Teams:</b> Ruby - Teams - 1:1 Chat F1  Standard Retention Policy
<br>
</p>
</p>
</p>
</p>
</p>
"


# Check Azure AD user to verify if user has an E3 or F3 license
#==============================================================
$VerifyE3License = Get-AzureADUserLicenseDetail -ObjectId $FrontEmp.UserPrincipalName | 
                   Select SkuID, SkuPartNumber |
                   Where SkuPartNumber -EQ "ENTERPRISEPACK" -Verbose

$verifyF3License = Get-AzureADUserLicenseDetail -ObjectId $FrontEmp.UserPrincipalName | 
                   Select SkuID, SkuPartNumber |
                   Where SkuPartNumber -EQ "SPE_F1" -Verbose


# Verify if user already has Retention Policy assigned
#=====================================================
$VerifyRetentionPolicy = Get-Mailbox -Identity "$($FrontEmp.SamAccountName)" |
                         Select RetentionPolicy



# If user has an F3 License Set Exchange Retention Policy
#========================================================
if ($verifyF3License -ne $null) {

                                # If User does not have a retention policy already applied
                                #=========================================================
                                if ($VerifyRetentionPolicy.RetentionPolicy -ne $MailboxRetentionPolicy) {

                                # Send message to Host Console, Apply Retention Policy, and send email notification to wizards distro
                                #====================================================================================================
                                Try {
                                     
                                     Write-Host "$($FrontEmp.SamAccountName) has an F3 License. Applying Retention Policy" -ForegroundColor Green
                                     Set-Mailbox $FrontEmp.SamAccountName -RetentionPolicy $MailboxRetentionPolicy -Verbose
                                     Send-MailMessage -SmtpServer PRLEMC01 -Body $emailBody -BodyAsHtml -From RetentionPolicyNotifs@ruby.com -To wizards@ruby.com -Subject "Retention Policy Applied for $($FrontEmp.SamAccountName)" -Priority High -Verbose

                                    }

                                    Catch {
                                    
                                           Write-Host "$($FrontEmp.SamAccountName) is already a member!" -ForegroundColor Red
                                          
                                          }
                           
                                }
                                
                                # If user is already assigned a retention policy and has an F3 License
                                #=====================================================================
                                Else {
                                      Write-Host "$($FrontEmp.SamAccountName) already is assign the Retention Policy:`r`n$($MailboxRetentionPolicy)" -ForegroundColor Cyan
                                     }
                              
                                }

# If user has an E3 License - Remove Retention Policy
#====================================================
if ($VerifyE3License -ne $null) {

                                Try {
                               
                                     Write-Host "$($FrontEmp.SamAccountName) has an E3 License. Removing Retention Policy" -ForegroundColor Yellow
                                     Set-Mailbox $FrontEmp.SamAccountName -RetentionPolicy $null -Verbose
                               
                                    }

                                    Catch {
                                   
                                          Write-Host "Policy was already removed!" -ForegroundColor Red
                                   
                                          }
    
    
    
    
    }

}


# Gets and removes Exchange PSSession
#====================================
Get-PSSession | Remove-PSSession

clear

# Security and Compliance Online PS Session initializaion
#========================================================
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ComplianceURI -Credential $LoginCredentials -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking

Write-Host "Applying the following teams retention policy: `r`n$($TeamsRetentionPolicy)`r`n" -ForegroundColor Green

foreach ($teamsFrontuser in $FrontLineEmployees) {
                                 
                                 Try {
                                      Write-Host "Applying Teams Retention Policy for $($teamsFrontuser.SamAccountName)" -ForegroundColor Magenta
                                      Set-TeamsRetentionCompliancePolicy -Identity $TeamsRetentionPolicy -AddTeamsChatLocation $teamsFrontuser.SamAccountName -RemoveTeamsChatLocation All -Verbose
                                     }
                                     
                                     catch {
                                            Write-Host "$($teamsFrontuser.SamAccountName) is already a member!" -ForegroundColor Red
                                           }
                                 
                                 }


Get-PSSession | Remove-PSSession

}

AssignRetentionPolicies
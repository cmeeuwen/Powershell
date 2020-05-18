clear

# Using the Stopwatch .NET Class to count how long this script takes to complete in minutes
#==========================================================================================
$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()


$LastScan = Get-ChildItem -Path \\prlfs01\ISWizards\SuperScript\Results\ExpiredPWDS\*.txt

            if ($LastScan -ne $null) {
                                      Write-Host "Removing Previous Results!`n" -ForegroundColor Blue
                                      Remove-Item -Path \\prlfs01\ISWizards\SuperScript\Results\ExpiredPWDS\*.txt -Force -Verbose
                                     }

                                     Else {
                                           Write-Host "No Previous Results Exist.`n" -ForegroundColor Green
                                          }




Function PassExpireWarning {

# Teams Webhook
#==============
$adminURL ="https://callruby.sharepoint.com/"
$siteURL ="https://callruby.sharepoint.com/sites/PlaceHolderSite"
$siteTitle ="PlaceHolderSite"
$connectorUri = 'https://outlook.office.com/webhook/863ef6d4-ef90-482d-832d-6f835d029a94@07b9247c-782b-40c8-b242-f2e598b2de5d/IncomingWebhook/38f26172adfc426ea44548598ecb5354/5d0b49bb-7ccd-45e0-8d41-5909742681e2'

#Connect-PnPOnline -Url $adminURL -UseWebLogin
#$site = New-PnPSite -Type CommunicationSite -Title $siteTitle -Url $siteURL -ErrorAction SilentlyContinue

# Teams Help Article
#===================
$teamsFix = "https://help.ruby.com/wb2.nhd?F=showkbentry&U=0&s=1"


# AzureAD password reset external link
#=====================================
$AzureADPasswordResetLink = "https://account.activedirectory.windowsazure.com/ChangePassword.aspx?BrandContextID=O365&ruO365="


# Font Color for template
#========================
$RedFontWarning = "color:red"



# Query Users that are members of the PSO_Standard_Ruby_User_Password Sec Group
#==============================================================================                       
$FindSamForEachMember = Get-ADGroupMember -identity "PSO_Standard_Ruby_User_Password" -Recursive | 
                        foreach{ get-aduser $_ -Properties *} | 
                        Where -Property UserPrincipalName -NE staff@callruby.com |
                        Select SamAccountName, GivenName, UserPrincipalName, PasswordLastSet, EmployeeID 
                        
                        
                        
                        
# Query Password Last Set and Sam Account Name Properties for each member (Loop Start)
#=====================================================================================
foreach($Item in $FindSamForEachMember){


# Calculates last date when password was changed
#===============================================
$PWDLastSetDate = (Get-Date $Item.PasswordLastSet).Date

# Gets todays date
#=================
$today = (get-date).Date 

# Calculates Date Difference from today and when password was last changed
#=========================================================================
$CalculateAge = $today - $PWDLastSetDate

# Displays Password Age in Console
#=================================
Write-Host "`r`nPassword Age for $($Item.SamAccountNAme) in Days: $($CalculateAge.TotalDays)" -ForegroundColor Yellow


# Calculation to determine how many days until password expiration
#=================================================================
$CalculatePWDExpire = 180 - $CalculateAge.Days


# Email Body Variables
#=====================
$emailBody7 = "
Hello $($Item.GivenName),
<br>
<p>Your password will expire in less than $($CalculatePWDExpire) days.
<br>
<p style=$($RedFontWarning)><b>Please make sure you change your password before it expires. Access to your system or other applications will be affected if you don't.</b>
<br>
<p><b>If you are on a windows system:</b>
<br>
<p>To Change your password on your PC press CTRL + ALT + Delete and choose <b>Change your Password</b>
<br>
<p><b>Note: </b>If you are working from home make sure you are connected to the VPN before trying
<br>to change your password.
<p><b>Password Requirements:</b>
<br>
<p>1. Must be a minimum of 8 characters
<br>
<p>2. Must contain at least 1 upper case
<br>
<p>3. Must contain at least 1 number
<br>
<p>4. Must contain at least 1 special character
<br>
<p>5. Cannot be your 4 previous passwords
<br>
<br>
<p><b>If you are on an inclement weather Laptop or a Mac system:</b>
<br>
<p>You will have to change your password on the Azure Password Reset Portal: <b><a href=$($AzureADPasswordResetLink)>Password Reset Link</a></b>
<br>
<br>
<p><b>If your password has already expired and you are not on a inclement weather laptop or Mac: </b>
<br>
<p><b>1.</b> Change your password on the Azure Password Reset Portal: <b><a href=$($AzureADPasswordResetLink)>Password Reset Link</a></b>
<br>
<p><b>2.</b> After the change, connect to the VPN using the new password.
<br>
<p><b>3.</b> After connecting to the VPN wait until you see a prompt from windows asking for your
<br>credentials. It will appear as a pop-up in the system tray. When you see that prompt, use
<br><b>ctrl + alt + del</b> to lock your computer, then unlock it with the new password.
<br>
<br>
<p><b>Additional Notes:</b>
<br>
<p>1. Changing your password could take 10 - 60 minutes to change across all of our systems. So it is best to change your password right before your lunch or at the end of your shift
<br>
<p>2. If you are unable to login to teams after changing your password and restarting please go through the steps in the article below
<br>
<p><b><a href=$($teamsFix)>Teams Help Article</a></b>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
"

$emailBody = "
Hello $($Item.GivenName),
<br>
<p>Your password will expire in less than 1 day.
<br>
<p style=$($RedFontWarning)><b>Please make sure you change your password before it expires. Access to your system or other applications will be affected if you don't.</b>
<br>
<p><b>If you are on a windows system:</b>
<br>
<p>To Change your password on your PC press CTRL + ALT + Delete and choose <b>Change your Password</b>
<br>
<p><b>Note: </b>If you are working from home make sure you are connected to the VPN before trying
<br>to change your password.
<p><b>Password Requirements:</b>
<br>
<p>1. Must be a minimum of 8 characters
<br>
<p>2. Must contain at least 1 upper case
<br>
<p>3. Must contain at least 1 number
<br>
<p>4. Must contain at least 1 special character
<br>
<p>5. Cannot be your 4 previous passwords
<br>
<br>
<p><b>If you are on an inclement weather Laptop or a Mac system:</b>
<br>
<p>You will have to change your password on the Azure Password Reset Portal: <b><a href=$($AzureADPasswordResetLink)>Password Reset Link</a></b>
<br>
<br>
<p><b>If your password has already expired and you are not on a inclement weather laptop or Mac: </b>
<br>
<p><b>1.</b> Change your password on the Azure Password Reset Portal: <b><a href=$($AzureADPasswordResetLink)>Password Reset Link</a></b>
<br>
<p><b>2.</b> After the change, connect to the VPN using the new password.
<br>
<p><b>3.</b> After connecting to the VPN wait until you see a prompt from windows asking for your
<br>credentials. It will appear as a pop-up in the system tray. When you see that prompt, use
<br><b>ctrl + alt + del</b> to lock your computer, then unlock it with the new password.
<br>
<br>
<p><b>Additional Notes:</b>
<br>
<p>1. Changing your password could take 10 - 60 minutes to change across all of our systems. So it is best to change your password right before your lunch or at the end of your shift
<br>
<p>2. If you are unable to login to teams after changing your password and restarting please go through the steps in the article below
<br>
<p><b><a href=$($teamsFix)>Teams Help Article</a></b>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
</p>
"


# If password is about to expire in 7 days or less but greater than 2
#====================================================================
if ($CalculatePWDExpire -le [int32]7 -and $CalculatePWDExpire -cge [int32]2){

Write-Host "Password is going to expire in $($CalculatePWDExpire) days!" -ForegroundColor Red

# HTML Body stored in Variable

# Send Email to user notifying that password will expire in less than 7 days
Send-MailMessage -Body $emailBody7 -From PasswordPolicy@ruby.com -BodyAsHtml -To "$($Item.UserPrincipalName)" -Subject "Password Expiration Warning!" -SmtpServer PRLEMC01 -Priority High -Verbose

}


#If password Expires in one day
#==============================
elseif ($CalculatePWDExpire -eq [int32]1){

Write-Host "Password is going to expire in $($CalculatePWDExpire) days!" -ForegroundColor Red



Send-MailMessage -Body $emailBody -From PasswordPolicy@ruby.com -BodyAsHtml -To "$($Item.UserPrincipalName)" -Subject "Password Expiration Warning!" -SmtpServer PRLEMC01 -Priority High -Verbose

}


# If Password is expired append results to text file
#===================================================
elseif ($CalculatePWDExpire -cle [int32]0){

Write-Host "Password is expired!" -ForegroundColor Red

"password is expired for $($Item.SamAccountNAme)`r`n" | Out-File -FilePath \\prlfs01\ISWizards\SuperScript\Results\ExpiredPWDS\ExpiredPWDS.txt -Append -Verbose


if($site){
    $body = '{"text":"Password is expired for STRA - Please Contact ASAP!"}'
    $newBody = $body.Replace("STRA",$Item.samaccountname)
}else {
    $body = '{"text":"Password is expired for STRA - Please Contact ASAP!"}'
    $newBody = $body.Replace("STRA",$Item.samaccountname)
}
Invoke-RestMethod -Uri $connectorUri -Method Post -Body $newBody -ContentType 'application/json'

}


# If Password is not close to expiring
#=====================================
elseif ($CalculatePWDExpire -gt [int32]7){

Write-Host "Password is still valid for $($CalculatePWDExpire) days!" -ForegroundColor Green

}


# Unable to determine Expiration
#===============================
Else {
      
      Write-Host "Unable to determine expiration date of password" -ForegroundColor Yellow
      
      }


}

}

PassExpireWarning

$Stopwatch.Stop()

Write-Host "`r`nThis took $($Stopwatch.Elapsed.Hours) hours $($Stopwatch.Elapsed.Minutes) minutes and $($Stopwatch.Elapsed.Seconds) seconds to run" -ForegroundColor Yellow

[System.GC]::Collect()
$error.Clear()

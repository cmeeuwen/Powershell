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


# Query Security Group Members
#=============================
$QuerySECGroupMembers = Get-ADGroup -Identity PSO_Standard_Ruby_User_Password -Properties * | 
                        Select Members -ExpandProperty Members

# Query Password Last Set and Sam Account Name Properties for each member (Loop Start)
#=====================================================================================
foreach($Member in $QuerySECGroupMembers){

$FindSamForEachMember = Get-ADUser -Filter * -Properties * | 
                        Where -Property DistinguishedName -Like "$($Member)" | 
                        Select *


# Teams Help Article
#===================
$teamsFix = "https://help.ruby.com/wb2.nhd?F=showkbentry&U=0&s=1"


# Calculates last date when password was changed
#===============================================
$PWDLastSetDate = (Get-Date $FindSamForEachMember.PasswordLastSet).Date

# Gets todays date
#=================
$today = (get-date).Date 

# Calculates Date Difference from today and when password was last changed
#=========================================================================
$CalculateAge = $today - $PWDLastSetDate

# Displays Password Age in Console
#=================================
Write-Host "Password Age for $($FindSamForEachMember.SamAccountNAme) in Days: $($CalculateAge.TotalDays)" -ForegroundColor Yellow


# Calculation to determine how many days until password expiration
#=================================================================
$CalculatePWDExpire = 180 - $CalculateAge.Days


# If password is about to expire in 7 days or less but greater than 2
#====================================================================
if ($CalculatePWDExpire -lt [int32]7 -and $CalculatePWDExpire -cge [int32]2){

Write-Host "Password is going to expire in $($CalculatePWDExpire) days!" -ForegroundColor Red

# HTML Body stored in Variable
$emailBody = "
Hello $($FindSamForEachMember.GivenName),
<br>
<p>Your password will expire in less than 7 days.
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
<p><b>Please make sure you change your password before it expires. Access to your system or other applications will be affected if you don't.</b>
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
"

# Send Email to user notifying that password will expire in less than 7 days
Send-MailMessage -Body $emailBody -From PasswordPolicy@ruby.com -BodyAsHtml -To "$($FindSamForEachMember.UserPrincipalName)" -Subject "Password Expiration Warning!" -SmtpServer PRLEMC01 -Priority High -Verbose

}


#If password Expires in one day
#==============================
elseif ($CalculatePWDExpire -eq [int32]1){

Write-Host "Password is going to expire in $($CalculatePWDExpire) days!" -ForegroundColor Red

$emailBody = "
Hello $($FindSamForEachMember.GivenName),
<br>
<p>Your password will expire in less than 1 day.
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
<p><b>Please make sure you change your password before it expires. Access to your system or other applications will be affected if you don't.</b>
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
"

Send-MailMessage -Body $emailBody -From PasswordPolicy@ruby.com -BodyAsHtml -To "$($FindSamForEachMember.UserPrincipalName)" -Subject "Password Expiration Warning!" -SmtpServer PRLEMC01 -Priority High -Verbose

}


# If Password is expired append results to text file
#===================================================
elseif ($CalculatePWDExpire -cle [int32]0){

Write-Host "Password is expired!" -ForegroundColor Red

"password is expired for $($FindSamForEachMember.SamAccountNAme)`r`n" | Out-File -FilePath \\prlfs01\ISWizards\SuperScript\Results\ExpiredPWDS\ExpiredPWDS.txt -Append -Verbose

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

Write-Host "This took $($Stopwatch.Elapsed.Hours) hours and $($Stopwatch.Elapsed.Minutes) minutes to run" -ForegroundColor Yellow

[System.GC]::Collect()
$error.Clear()
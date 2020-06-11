Function PasswordScrambler {

# Set user to password scramble
#==============================
$SamAccountName = Read-Host -Prompt "Username"

# Load System.Web .net Assembly
#==============================
add-type -AssemblyName System.Web

# Set Min / Max Length for new scrambled password
#================================================
$minLength = [int32]8 ## characters
$maxLength = [int32]20 ## characters
$length = Get-Random -Minimum $minLength -Maximum $maxLength 

# Sets the amount of non Alpha Numeric Characters in the generated password
#==========================================================================
$nonAlphaChars = [Int32]2
$password = [System.Web.Security.Membership]::GeneratePassword($length, $nonAlphaChars)

# Changes requested account password to new scrambled password and displays in console to tech
#==============================================================================================
Set-ADAccountPassword -Identity $SamAccountName -NewPassword (ConvertTo-SecureString -AsPlainText "$($password)" -Force -Verbose)
write-host "New Scambled password for $($SamAccountName) was set to: $($password)" -ForegroundColor Green

}

# Function Call
#==============
PasswordScrambler | Out-Null
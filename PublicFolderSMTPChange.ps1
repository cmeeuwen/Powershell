Function RemRepPublicFolderSMTP {

Clear

# Create Pssession for Exchange
#==============================
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

Import-PSSession $Session -DisableNameChecking

clear

Write-Host "
#==================================================#
#                                                  #
# 1. Provide smtp address for public folder for    #
#    Search Query and to remove proxyaddress from  #
#    proxyaddresses property array                 #
#                                                  #
# 2. Provide New SMTP Address to add new           #
#    proxyaddress to proxyaddresses property       #
#    Array                                         #
#                                                  #
#==================================================#
" -ForegroundColor Yellow

$SMTP = Read-Host -Prompt "Bad STMPAddress "

$FindPublicMailbox = Get-ADObject -Filter * -Properties * | 
                     Where {$_.proxyaddresses -eq "smtp:$($SMTP)" } | 
                     Select proxyaddresses -ExpandProperty proxyaddresses

$FindPublicMailboxGUID = Get-ADObject -Filter * -Properties * |
                         Where {$_.proxyAddresses -eq "smtp:$($SMTP)"} |
                         Select *

foreach($address in $FindPublicMailbox) {

if ($address -eq "smtp:$($SMTP)") {

Write-Host "$($address) is a match`r`n" -ForegroundColor Green

$newSMTP = Read-Host -Prompt "New proxyAddress "
 
Set-ADObject $FindPublicMailboxGUID.ObjectGUID -Remove @{proxyAddresses=$address} -Confirm -Verbose

Set-ADObject $FindPublicMailboxGUID.ObjectGUID -Add @{proxyAddresses="smtp:$($newSMTP)"} -Confirm -Verbose

}
else {Write-Host "$($address) is not a match!" -ForegroundColor Red}


}

Get-PSSession | Remove-PSSession

}

RemRepPublicFolderSMTP
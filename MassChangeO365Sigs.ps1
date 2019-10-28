Function MassUpdateRecepSigsVIACSV {

clear

# location of signature template
$LocationOfSignature= "\\prlfs01\PDQ_Shared_DB\Repository\Signature_Creation\OWA_Signature.htm"

# see if already connected to exchange online
$AmIConnectedToExchangeOnline= Get-PSSession | where {$_.ConfigurationName -eq "Microsoft.exchange"}


    IF ($AmIConnectedToExchangeOnline -eq $null -or $AmIConnectedToExchangeOnline.State -eq "Closed")
        {
        #Connect to Exchange Online
        
        $userCredentials= Get-Credential -message "Enter Exchange Online Email/Credentials to continue" 
        
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredentials -Authentication Basic -AllowRedirection
        
        Import-PSSession $Session -Verbose
         

        }

clear

#Import the template for the signature

$importSignatureTemplate= Get-Content $LocationOfSignature

#Import CSV for receps being updated 
$csvFile = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
           InitialDirectory = [Environment]::GetFolderPath('Desktop') 
           Filter = 'CSV (*.csv)|*.csv|SpreadSheet (*.xlsx)|*.xlsx'
           }
           $csvFile.ShowDialog()

   
            Import-Csv $csvFile.FileName | ForEach-Object {
            
            
            $OWA_Sigs_To_Update = $_.UserPrincipalName

            $GetADUserInfo= Get-ADUser -Filter * -Properties title,department,enabled,emailaddress | where {$_.title -notlike $Null -and $_.department -notlike $null} -Verbose

            $UserDetails= $GetADUserInfo | where {$_.UserPrincipalName -eq $OWA_Sigs_To_Update} -Verbose

            $UserDetails

            FOREACH ($signatureUser in $UserDetails)
    {

 #Generated Signature

 $Signature= @()


 # Find the POWERSHELLRECPNAME, and the POWERSHELLRECPTEAMNAME and replace those with the name and the team of the Receptionist.
    Foreach ($line in $importSignatureTemplate)
        {

            # See if $line contains "POWERSHELLRECPNAME and replace it with the receptionists name.
            
            IF ($line -like "* POWERSHELLRECPNAME *")
                {
                $signature+= $line.Replace("POWERSHELLRECPNAME","$($signatureUser.name)")
                }
            
            # See if $line contains "POWERSHELLRECEPTEAMNAME and replace it with the Receptionists Team Name. 
            
            IF ($line -like "* POWERSHELLRECPTEAMNAME *")
                {
                $signature+=$line.Replace("POWERSHELLRECPTEAMNAME","$($signatureUser.title)")
                }

            IF ($line -notlike "* POWERSHELLRECPTEAMNAME *" -and $line -notlike "* POWERSHELLRECPNAME *")
                {
                $signature+=$line
                }

        }
          
            Set-MailboxMessageConfiguration $signatureUser.UserPrincipalName -AutoAddSignature $true -SignatureHtml $Signature -DefaultFontColor "#352E1B"

     }
    
clear

write-host "`r`nSignatures should be created for $($_.UserPrincipalName) . Please check OWA to verify." -ForegroundColor Green


}

Remove-PSSession $Session -Verbose
} 

MassUpdateRecepSigsVIACSV
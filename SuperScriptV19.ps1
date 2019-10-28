#==========================================#
#---------------Super Script---------------#
#==========================================#
#  Written by: Chris Meeuwen   2/12/2018   #
#  Version: 19.0 - Modified    8/06/2019   #
#                                          #
#==========================================#


#=========================================Place New Functions here=======================================#
#========================================================================================================#


# ============================================= #
#              Built in functions               #
# ============================================= #

# Stop Process on any system on the domain
function StpPrc {

clear
Write-host "-------Super Process Kill!------" -BackgroundColor Blue -ForegroundColor Green 

Write-host "MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
MMMMMMMMMMMM        MMMMMMMMMMMM
MMMMMMMMMM            MMMMMMMMMM
MMMMMMMMM              MMMMMMMMM
MMMMMMMM                MMMMMMMM
MMMMMMM                 MMMMMMMM
MMMMMMM                  MMMMMMM
MMMMMMM                  MMMMMMM
MMMMMMM    MMM    MMM    MMMMMMM
MMMMMMM   MMMMM   MMMM   MMMMMMM
MMMMMMM   MMMMM   MMMM   MMMMMMM
MMMMMMMM   MMMM M MMMM  MMMMMMMM
MMMMMMMM        M        MMMMMMM
MMMMMMMM       MMM      MMMMMMMM
MMMMMMMMMMMM   MMM  MMMMMMMMMMMM
MMMMMMMMMM MM       M  MMMMMMMMM
MMMMMMMMMM  M M M M M MMMMMMMMMM
MMMM  MMMMM MMMMMMMMM MMMMM   MM
MMM    MMMM M MMMMM M MMMM    MM
MMM    MMMM   M M M  MMMMM   MMM
MMMM    MMMM         MMM      MM
MMM       MMMM     MMMM       MM
MMM         MMMMMMMM      M  MMM
MMMM  MMM      MMM      MMMMMMMM
MMMMMMMMMMM  MM       MMMMMMM  M
MMM  MMMMMMM       MMMMMMMMM   M
MM    MMM        MM            M
MM            MMMM            MM
MMM        MMMMMMMMMMMMM       M
MM      MMMMMMMMMMMMMMMMMMM    M
MMM   MMMMMMMMMMMMMMMMMMMMMM   M
MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM" -ForegroundColor green

$computer = Read-Host -Prompt 'Type in the name of the remote or local system'

Write-Host "List of active processes running on selected host
-------------------------------------------------"

Get-Process -ComputerName $computer | Out-GridView -Title $computer

#gwmi Win32_process -ComputerName wizd | Select-Object -Property Name,ProcessName,PSComputerName,Description | Out-GridView

Write-Host ""

$process = Read-Host -prompt 'Type in the process name'
$date = Get-date
$ErrorActionPreference = 'silentlycontinue'

gwmi Win32_process -ComputerName $computer | Select -Property Name,ProcessName,PSComputerName,Description | Out-GridView
(gwmi Win32_Process -ComputerName $computer | ?{ $_.ProcessName -match "$($process)"}).Terminate()

Write-host "$($process) Terminated with extreme prejudice!" -ForegroundColor Red
Write-host "`nPerformed on" + $date

pause

}

# This Function Adds assist Tracker to a users specified Workstation
function AssistTrack {

Write-Host "-----Assist Tracker Shortcut Creation Script-----" -ForegroundColor Green -BackgroundColor Blue
Write-Host ""
Write-Host ""

$ErrorActionPreference = 'silentlycontinue'
$hostname = Read-Host -prompt "What system? "
$source = ($PSScriptRoot + '\Assist Tracker\*.*')
$destination = ('\\' + $hostname + '\C$\Assist Tracker')
$testpth1 = Test-Path \\$($hostname)\C$\AssistTracker
$testpth1
$testpth2 = Test-Path ('\\' + $($hostname) + '\c$\Assist Tracker')
$testpth2 

if ($testpth1 -eq $true) {
    Write-Host "Removing $testpth1 and creating new directory \\$hostname\c$\Assist Tracker" -ForegroundColor Yellow
    Remove-Item -Path \\$hostname\C$\AssistTracker -Recurse -Verbose -Force -ErrorAction SilentlyContinue
    New-Item -Path \\$hostname\c$ -Name 'Assist Tracker' -ItemType Directory -Verbose -Force
}
if ($testpth2 -eq $true) {
    Write-Host "Removing $testpth2 and creating new directory \\$hostname\c$\Assist Tracker" -ForegroundColor Yellow
    Remove-Item -Path ('\\' + $hostname + '\C$\Assist Tracker') -Recurse -Verbose -Force -ErrorAction SilentlyContinue
    New-Item -Path \\$hostname\c$ -Name 'Assist Tracker' -ItemType Directory -Verbose -Force
}
else {
        Write-Host "Creating new directory \\$hostname\c$\Assist Tracker" -ForegroundColor Yellow
        New-Item -Path \\$hostname\c$ -Name 'Assist Tracker' -ItemType Directory -Verbose -Force
        }

Start-BitsTransfer -Source $source -Destination $destination -Description "Copy Status.." -DisplayName "Copying..."
Copy ('\\' + $hostname + '\C$\Assist Tracker\AssistTracker.lnk') ('\\' + $hostname + '\C$\Users\Public\Desktop') -Force

clear

Write-Host "Assist Tracker Shortcuts added!"  -ForegroundColor Green | pause 

}

# This Function Updates RosOS on any specified workstation on a domain
function UpdateRos {

Write-Host "-----Update Ruby OS-----" -ForegroundColor Green -BackgroundColor Blue
Write-Host ""
Write-Host ""
Write-Host "This Script will replace the 
RubyOS on any system Specified.

"

$ErrorActionPreference = 'silentlycontinue'
$hostname = Read-Host -prompt "What system? "
$stationNumb = Read-Host -Prompt "What Station? "
$source = ($PSScriptRoot + '\Applications\RubyOS\*.*')
$destination = ('\\' + $hostname + '\C$\RubyOS')

(gwmi Win32_Process -ComputerName $hostname | ?{ $_.ProcessName -match "RubyOS"}).Terminate()

rd -Path ('\\' + $hostname + '\C$\RubyOS') -Verbose -Force
md -Path ('\\' + $hostname + '\C$\RubyOS')

Start-BitsTransfer -Source $source -Destination $destination -Description "Copy Status.." -DisplayName "Copying..."

    (Get-Content \\$($hostname)\c$\RubyOS\settings.xml) -replace '%station%', $stationNumb | Set-Content \\$($hostname)\c$\RubyOS\settings.xml -Verbose

clear

Write-Host "RubyOS Updated on system $($hostname)"  -ForegroundColor Green | pause

}

# Clears Appdata Roaming Cache for Microsoft Teams
function ClearTeamsCache {

clear

$ErrorActionPreference = 'silentlycontinue'

Write-Host "
     Clear Microsoft Teams Application Cache
=================================================
| Synopsis: This script will clear the Roaming  |
| Appdata Cache for Microsoft Teams             |
|                                               |
| Once complete the user will have to           |
| log back in with thier UPN name               |
|                                               |
| EX: firstname.lastname@callruby.com           |
|                                               |
=================================================
`n" -ForegroundColor Yellow

$TeamsHost = Read-Host -Prompt "What system "
$TeamsUser = Read-Host -Prompt "What User "
$isTeamsHostOnline = Test-Connection -ComputerName $TeamsHost | Select-Object -Property IPV4Address
$TeamsProcess = ('\\' + $TeamsHost + '\c$\Users\' + $TeamsUser + '\AppData\Local\Microsoft\Teams\current\Teams.exe')

    if ($isTeamsHostOnline -ne $null){

        try {
            (gwmi Win32_Process -ComputerName $TeamsHost | ?{ $_.ProcessName -like "*Teams*"}).Terminate()
            }
            catch {
                   Write-Host "Teams Process is not running on $($TeamsHost)" -ForegroundColor Yellow
                   }

            Remove-Item -Path ('\\' + $TeamsHost + '\c$\Users\' + $TeamsUser + '\AppData\Roaming\Microsoft\Teams\' + 'Application Cache') -Recurse -Force -Verbose
            Remove-Item -Path ('\\' + $TeamsHost + '\c$\Users\' + $TeamsUser + '\AppData\Roaming\Microsoft\Teams\' + 'blob_storage') -Recurse -Force -Verbose
            Remove-Item -Path ('\\' + $TeamsHost + '\c$\Users\' + $TeamsUser + '\AppData\Roaming\Microsoft\Teams\' + 'Cache') -Recurse -Force -Verbose
            Remove-Item -Path ('\\' + $TeamsHost + '\c$\Users\' + $TeamsUser + '\AppData\Roaming\Microsoft\Teams\' + 'databases') -Recurse -Force -Verbose
            Remove-Item -Path ('\\' + $TeamsHost + '\c$\Users\' + $TeamsUser + '\AppData\Roaming\Microsoft\Teams\' + 'GPUCache') -Recurse -Force -Verbose
            Remove-Item -Path ('\\' + $TeamsHost + '\c$\Users\' + $TeamsUser + '\AppData\Roaming\Microsoft\Teams\' + 'IndexedDB') -Recurse -Force -Verbose
            Remove-Item -Path ('\\' + $TeamsHost + '\c$\Users\' + $TeamsUser + '\AppData\Roaming\Microsoft\Teams\' + 'Local Storage' ) -Recurse -Force -Verbose
            Remove-Item -Path ('\\' + $TeamsHost + '\c$\Users\' + $TeamsUser + '\AppData\Roaming\Microsoft\Teams\' + 'tmp' ) -Recurse -Force -Verbose
            Start-Process -FilePath $TeamsProcess -Verbose
       }
        else {
              Write-Host "$($TeamsHost) is either not online or not able to connect!" -ForegroundColor Red
             }

write-host "Teams Roaming Cache cleared on system $($TeamsHost)" -ForegroundColor Green
pause
clear

}

# Created new Ruby's in AD and Exchange
function NewRuby {

# Creates AD and Exchange accounts for Beaveton Office (Includes HP mailbox)
function VerNameBvt{

clear

Write-Host "
+================Warning! Read before using==================+
| This function should only be used if ED-E fails to         |
| automatically create a user.                               |
|                                                            |
| If a user exits in the Users to be filed OU please use the |
| Enable Rubys Generated from ED-E function under the new    |
| Ruby section.                                              |
|                                                            |
+============================================================+
" -ForegroundColor Red

$vfirstname = Read-Host -Prompt "First name"
$vlastname = Read-Host -Prompt "Last name"
$vusername = ($vfirstname + '.' + $vlastname)

while ($true){


     if($vusername -notmatch "^[a-z.A-Z]*$")
          {
          Clear
            Write-Host "Invalid Symbols or spaces within first or last name!" -BackgroundColor Blue -ForegroundColor Red
                
                Clear-Variable $vfirstname -Force
                Clear-Variable $vlastname -Force
                Clear-Variable $vusername -Force

                    $TryAgain = Read-Host -Prompt "Do you want to try again?  Y/N "
                if ($TryAgain -eq 'y') {VerNameBvt}
            else {NewRuby}
          }
          else {
          Clear
          Write-Host "Onboard new BVT Receptionist
==============================" -BackgroundColor Blue -ForegroundColor Green
Write-Host ""

$ErrorActionPreference = 'silentlycontinue'

# Setting Variables to create user and add them to the proper
# OU and Groups
$firstname = ($vfirstname)
$lastname = ($vlastname)
$officeloc = Read-Host -Prompt "Station Number / Office Location (example: BVT123)"
$department = 'Receptionist Services'
$manager = Read-Host -Prompt "Manager name (Account Name of manager)"
$company = 'Ruby Receptionists'
$RecepTeam = Read-Host -Prompt "What team (Novas, Racers, Etc....)"
$jobtitle = ($RecepTeam + ' Receptionist')
$fstlst = ($firstname + ' ' + $lastname)

# Creates new AD user and has you create a password for the new user
# then enables the AD user
New-ADUser -Name ($vusername) -UserPrincipalName ($vusername + '@callruby.com') -AccountPassword(Read-Host -AsSecureString "Account Password") -PassThru | Enable-ADAccount -Verbose

# Modifys Multiple attributes on newly created user
# Based on the input at the beginning of the script
# Then moves user to the proper OU
Get-ADUser $vusername | Set-ADUser -CannotChangePassword $true -Verbose
Get-ADUser $vusername | Set-ADUser -PasswordNeverExpires $true -Verbose
Get-ADUser $vusername | Set-ADUser -Surname $lastname -Verbose
Get-ADUser $vusername | Set-ADUser -GivenName $firstname -Verbose
Get-ADUser $vusername | Set-ADUser -Office $officeloc -Verbose
Get-ADUser $vusername | Set-ADUser -DisplayName $fstlst -Verbose
Get-ADUser $vusername | Set-ADUser -EmailAddress ($vusername + '@callruby.com').ToLower();
Get-ADUser $vusername | Set-ADUser -ScriptPath 'btown.bat' -Verbose
Get-ADUser $vusername | Set-ADUser -Title $jobtitle -Verbose
Get-ADUser $vusername | Set-ADUser -Manager $manager -Verbose
Get-ADUser $vusername | Set-ADUser -Company $company -Verbose
Get-ADUser $vusername | Set-ADUser -Department $department -Verbose
Get-ADUser $vusername | Move-ADObject -TargetPath 'OU=BVN Receptionists,OU=Ruby Users,DC=ruby,DC=local' -Verbose
Get-ADUser $vusername | Rename-ADObject -NewName ($fstlst) -Verbose

# Adds user to associated AD groups
Add-ADGroupMember -Identity $RecepTeam $vusername -Verbose
Add-ADGroupMember -Identity Exclaimer_Group1 $vusername -Verbose
Add-ADGroupMember -Identity beavertonTeam $vusername -Verbose
Add-ADGroupMember -Identity Receptionists $vusername -Verbose
Add-ADGroupMember -Identity TeamRuby $vusername -Verbose
Add-ADGroupMember -Identity SEC_SOPHOS_Standard_WebFiltering $vusername -Verbose
Add-ADGroupMember -Identity ALL_Receptionists_Bvt_Sec $NewRubyUser -Verbose
Add-ADGroupMember -Identity All_Receptionists_SEC $vusername -Verbose

# Enables Exchange mailbox for new on premise instance
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;
Enable-RemoteMailbox $vusername -RemoteRoutingAddress ($vusername + '@callruby.com') -Verbose

clear

Write-Host "New Receptionist created to AD! Please allow 24 hours for Office365 account to be created." -BackgroundColor Blue -ForegroundColor Green
pause

          break;
          }
          break;
}

}


# Creates AD and Exchange accounts for Pearl Office (Includes HP mailbox)
function VerNamePrl{

clear

Write-Host "
+================Warning! Read before using==================+
| This function should only be used if ED-E fails to         |
| automatically create a user.                               |
|                                                            |
| If a user exits in the Users to be filed OU please use the |
| Enable Rubys Generated from ED-E function under the new    |
| Ruby section.                                              |
|                                                            |
+============================================================+
" -ForegroundColor Red

$vfirstname = Read-Host -Prompt "First name"
$vlastname = Read-Host -Prompt "Last name"
$vusername = ($vfirstname + '.' + $vlastname)

while ($true){


     if($vusername -notmatch "^[a-z.A-Z]*$")
          {
          Clear
          Write-Host "Invalid Symbols or spaces within first or last name!" -BackgroundColor Blue -ForegroundColor Red
          
                Clear-Variable lastname -Force
                Clear-Variable firstname -Force
                Clear-Variable username -Force
          
                    $TryAgain = Read-Host -Prompt "Do you want to try again?  Y/N "
                if ($TryAgain -eq 'y') {VerNamePrl}
            else {NewRuby}
          }
          else {
          Clear
          Write-Host "Onboard new PEARL Receptionist
==============================" -BackgroundColor Blue -ForegroundColor Green
Write-Host ""

$ErrorActionPreference = 'silentlycontinue'

# Setting Variables to create user and add them to the proper
# OU and Groups
$firstname = ($vfirstname)
$lastname = ($vlastname)
$officeloc = Read-Host -Prompt "Station Number / Office Location (example: prl123)"
$department = 'Receptionist Services'
$manager = Read-Host -Prompt "Manager name (Account Name of manager)"
$company = 'Ruby Receptionists'
$RecepTeam = Read-Host -Prompt "What team (Novas, Racers, Etc....)"
$jobtitle = ($RecepTeam + ' Receptionist')
$fstlst = ($firstname + ' ' + $lastname)

# Creates new AD user and has you create a password for the new user
# then enables the AD user
New-ADUser -Name ($vusername) -UserPrincipalName ($vusername + '@callruby.com') -AccountPassword(Read-Host -AsSecureString "Account Password") -PassThru | Enable-ADAccount -Verbose

# Modifys Multiple attributes on newly created user
# Based on the input at the beginning of the script
# Then moves user to the proper OU
Get-ADUser $vusername | Set-ADUser -CannotChangePassword $true -Verbose
Get-ADUser $vusername | Set-ADUser -PasswordNeverExpires $true -Verbose
Get-ADUser $vusername | Set-ADUser -Surname $lastname -Verbose
Get-ADUser $vusername | Set-ADUser -GivenName $firstname -Verbose
Get-ADUser $vusername | Set-ADUser -Office $officeloc -Verbose
Get-ADUser $vusername | Set-ADUser -DisplayName $fstlst -Verbose
Get-ADUser $vusername | Set-ADUser -EmailAddress ($vusername + '@callruby.com').ToLower();
Get-ADUser $vusername | Set-ADUser -ScriptPath 'recep.bat' -Verbose
Get-ADUser $vusername | Set-ADUser -Title $jobtitle -Verbose
Get-ADUser $vusername | Set-ADUser -Manager $manager -Verbose
Get-ADUser $vusername | Set-ADUser -Company $company -Verbose
Get-ADUser $vusername | Set-ADUser -Department $department -Verbose
Get-ADUser $vusername | Move-ADObject -TargetPath 'OU=PDX Receptionists,OU=Ruby Users,DC=ruby,DC=local' -Verbose
Get-ADUser $vusername | Rename-ADObject -NewName ($fstlst) -Verbose

# Adds user to associated AD groups
Add-ADGroupMember -Identity $RecepTeam $vusername -Verbose
Add-ADGroupMember -Identity Exclaimer_Group1 $vusername -Verbose
Add-ADGroupMember -Identity PearlTeam $vusername -Verbose
Add-ADGroupMember -Identity Receptionists $vusername -Verbose
Add-ADGroupMember -Identity TeamRuby $vusername -Verbose
Add-ADGroupMember -Identity SEC_SOPHOS_Standard_WebFiltering $vusername -Verbose
Add-ADGroupMember -Identity ALL_Receptionists_Prl_Sec $vusername -Verbose
Add-ADGroupMember -Identity All_Receptionists_SEC $vusername -Verbose

# Enables Exchange mailbox for new on premise instance
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -Verbose
Enable-RemoteMailbox $vusername -RemoteRoutingAddress ($vusername + '@callruby.com') -Verbose

clear

Write-Host "New Receptionist created to AD! Please allow 24 hours for Office365 account to be created." -BackgroundColor Blue -ForegroundColor Green
pause

clear

          
          break;
          }
          break;
}

}


# Creates AD and Exchange accounts for Sales (Includes HP mailbox)
function OnBrdSales {

Write-Host "Onboard new Sales
==============================" -BackgroundColor Blue -ForegroundColor Green
Write-Host ""

$ErrorActionPreference = 'silentlycontinue'

# Setting Variables to create user and add them to the proper
# OU and Groups
$firstname = Read-Host -Prompt "First name"
$lastname = Read-Host -Prompt "Last name"
$officeloc = 'FXT'
$department = 'Sales'
$manager = Read-Host -Prompt "Manager name (Account Name of manager)"
$company = 'Ruby Receptionists'
$jobtitle = ('Inside Sales Representative')
$fstlst = ($firstname + ' ' + $lastname)
$hpmainname = ('HP' + ' ' + $firstname + ' ' + $lastname)
$hpAdGroup = ('HP' + $RecepTeam)
$hpUser = ('HP' + $firstname + '.' + $lastname)
$hpsur = ('HP ' + $lastname)
$hpgiv = $lastname
$hpOU = 'OU=Hurdles & Praise,OU=Non-Human Users,OU=Ruby Users - System\, Non-Human\, External,DC=ruby,DC=local'
$hpprincname = ('HP' + $firstname + '.' + $lastname + '@callruby.com')


$username = ($firstname + '.' + $lastname)

# Creates new AD user and has you create a password for the new user
# then enables the AD user
New-ADUser -Name ($username) -UserPrincipalName ($username + '@callruby.com') -AccountPassword(Read-Host -AsSecureString "Account Password") -PassThru | Enable-ADAccount -Verbose

# Modifys Multiple attributes on newly created user
# Based on the input at the beginning of the script
# Then moves user to the proper OU
Get-ADUser $username | Set-ADUser -CannotChangePassword $true -Verbose
Get-ADUser $username | Set-ADUser -PasswordNeverExpires $true -Verbose
Get-ADUser $username | Set-ADUser -Surname $lastname
Get-ADUser $username | Set-ADUser -GivenName $firstname
Get-ADUser $username | Set-ADUser -Office $officeloc
Get-ADUser $username | Set-ADUser -DisplayName $fstlst
Get-ADUser $username | Set-ADUser -EmailAddress ($username + '@callruby.com')
Get-ADUser $username | Set-ADUser -ScriptPath 'sales.bat'
Get-ADUser $username | Set-ADUser -Title $jobtitle
Get-ADUser $username | Set-ADUser -Manager $manager
Get-ADUser $username | Set-ADUser -Company $company
Get-ADUser $username | Set-ADUser -Department $department
Get-ADUser $username | Move-ADObject -TargetPath 'OU=PDX Sales,OU=Ruby Users,DC=ruby,DC=local' -Verbose
Get-ADUser $username | Rename-ADObject -NewName ($fstlst) -Verbose

# Adds user to associated AD groups
Add-ADGroupMember -Identity $RecepTeam $username -Verbose
Add-ADGroupMember -Identity non-CHfoxteam@callruby.com $username -Verbose
Add-ADGroupMember -Identity NonReceptionists $username -Verbose
Add-ADGroupMember -Identity TeamRuby $username -Verbose
Add-ADGroupMember -Identity pearl_sql_users $username -Verbose
Add-ADGroupMember -Identity SPARK_SalesMarketingPartnership $username -Verbose
Add-ADGroupMember -Identity FoxTower $username -Verbose
Add-ADGroupMember -Identity FS_TC_Audio_Agreements_RW $username -Verbose
Add-ADGroupMember -Identity GrowingLeaders $username -Verbose
Add-ADGroupMember -Identity pearl_credit_card_batch $username -Verbose
Add-ADGroupMember -Identity pearl_credit_card_process $username -Verbose
Add-ADGroupMember -Identity pearlsales $username -Verbose
Add-ADGroupMember -Identity PublicFolderAccess $username -Verbose

clear

Write-Host "New Sales created to AD and Exhange!" -BackgroundColor Blue -ForegroundColor Green
pause

clear
}

# Enable users that were created from ED-E
Function Enable_Receptionist {

function Pearly {
            
            $NewRubyUser = Read-Host "Username "
            $NewRubyTeam = Read-Host "What team "
            $NewRubyCulti = Read-Host "Manager "
            $NewRubyOfficeLoc = Read-Host "Station "

            Get-ADUser $NewRubyUser | Enable-ADAccount -Verbose

            Get-ADUser $NewRubyUser | Set-ADUser -Office $NewRubyOfficeLoc -Verbose
            Get-ADUser $NewRubyUser | Set-ADUser -EmailAddress ($NewRubyUser + '@callruby.com').ToLower();
            Get-ADUser $NewRubyUser | Set-ADUser -Manager $NewRubyCulti -Verbose
            Get-ADUser $NewRubyUser | Set-ADUser -Title "$NewRubyTeam Receptionist" -Verbose

            Get-ADUser $NewRubyUser | Move-ADObject -TargetPath 'OU=PDX Receptionists,OU=Ruby Users,DC=ruby,DC=local' -Verbose
            Add-ADGroupMember -Identity PearlTeam $NewRubyUser -Verbose
            Add-ADGroupMember -Identity $NewRubyTeam $NewRubyUser -Verbose
            Add-ADGroupMember -Identity Exclaimer_Group1 $NewRubyUser -Verbose
            Add-ADGroupMember -Identity Receptionists $NewRubyUser -Verbose
            Add-ADGroupMember -Identity TeamRuby $NewRubyUser -Verbose
            Add-ADGroupMember -Identity Ruby_All_Users $NewRubyUser -Verbose
            Add-ADGroupMember -Identity ALL_Receptionists_Prl_Sec $NewRubyUser -Verbose
            Add-ADGroupMember -Identity All_Receptionists_SEC $NewRubyUser -Verbose
            Add-ADGroupMember -Identity SEC_SOPHOS_Standard_WebFiltering $NewRubyUser -Verbose

            $StopWatch = [System.Diagnostics.StopWatch]::StartNew()

 

Function Test-Command ($Command)



{

    Try

    {

        Get-command $command -ErrorAction Stop

        Return $True

    }

    Catch [System.SystemException]

    {

        Return $False

    }

}

 

IF (Test-Command "Get-Mailbox") {Write-Host "Exchange cmdlets already present"}

Else {

    $CallEMS = ". '$env:ExchangeInstallPath\bin\RemoteExchange.ps1'; Connect-ExchangeServer PRLEMC01 -ClientApplication:ManagementShell "

    Invoke-Expression $CallEMS


$stopwatch.Stop()
$msg = "`n`nThe script took $([math]::round($($StopWatch.Elapsed.TotalSeconds),2)) seconds to execute..."
Write-Host $msg
$msg = $null
$StopWatch = $null

}



$runEnablePearl = Enable-RemoteMailbox -Identity $NewRubyUser -RemoteRoutingAddress ($NewRubyUser + '@callruby.com') -Verbose
$runEnablePearl

Write-Host "$NewRubyUser Enabled!
===========================================================================================
| Run the following command in Exchange Management Shell to verify remotemailbox user     |
| Command: Get-RemoteMailbox -Identity %Username%                                         |
===========================================================================================" -ForegroundColor Yellow
pause

clear
            

}

function Beaver {
            
            $NewRubyUser = Read-Host "Username "
            $NewRubyTeam = Read-Host "What team "
            $NewRubyCulti = Read-Host "Manager "
            $NewRubyOfficeLoc = Read-Host "Station "

            Get-ADUser $NewRubyUser | Enable-ADAccount -Verbose

            Get-ADUser $NewRubyUser | Set-ADUser -Office $NewRubyOfficeLoc -Verbose
            Get-ADUser $NewRubyUser | Set-ADUser -EmailAddress ($NewRubyUser + '@callruby.com').ToLower();
            Get-ADUser $NewRubyUser | Set-ADUser -Manager $NewRubyCulti -Verbose
            Get-ADUser $NewRubyUser | Set-ADUser -Title "$NewRubyTeam Receptionist" -Verbose

            Get-ADUser $NewRubyUser | Move-ADObject -TargetPath 'OU=BVN Receptionists,OU=Ruby Users,DC=ruby,DC=local' -Verbose
            Add-ADGroupMember -Identity beavertonteam $NewRubyUser -Verbose
            Add-ADGroupMember -Identity $NewRubyTeam $NewRubyUser -Verbose
            Add-ADGroupMember -Identity Exclaimer_Group1 $NewRubyUser -Verbose
            Add-ADGroupMember -Identity Receptionists $NewRubyUser -Verbose
            Add-ADGroupMember -Identity TeamRuby $NewRubyUser -Verbose
            Add-ADGroupMember -Identity Ruby_All_Users $NewRubyUser -Verbose
            Add-ADGroupMember -Identity ALL_Receptionists_Bvt_Sec $NewRubyUser -Verbose
            Add-ADGroupMember -Identity All_Receptionists_SEC $NewRubyUser -Verbose
            Add-ADGroupMember -Identity SEC_SOPHOS_Standard_WebFiltering $NewRubyUser -Verbose

            $StopWatch = [System.Diagnostics.StopWatch]::StartNew()

 

Function Test-Command ($Command)



{

    Try

    {

        Get-command $command -ErrorAction Stop

        Return $True

    }

    Catch [System.SystemException]

    {

        Return $False

    }

}

 

IF (Test-Command "Get-Mailbox") {Write-Host "Exchange cmdlets already present"}

Else {

    $CallEMS = ". '$env:ExchangeInstallPath\bin\RemoteExchange.ps1'; Connect-ExchangeServer PRLEMC01 -ClientApplication:ManagementShell "

    Invoke-Expression $CallEMS


$stopwatch.Stop()
$msg = "`n`nThe script took $([math]::round($($StopWatch.Elapsed.TotalSeconds),2)) seconds to execute..."
Write-Host $msg
$msg = $null
$StopWatch = $null

}

$TestEMC = Get-RemoteMailbox -Identity bob.doe
$TestEMC

if ($TestEMC -ne $null) {

Enable-RemoteMailbox -Identity $NewRubyUser -RemoteRoutingAddress ($NewRubyUser + '@callruby.com') -Verbose


Write-Host "$NewRubyUser Enabled!
===========================================================================================
| Run the following command in Exchange Management Shell to verify remotemailbox user     |
| Command: Get-RemoteMailbox -Identity %Username%                                         |
===========================================================================================" -ForegroundColor Yellow
pause
}
Else{
     Write-Host "Either Connection to EMC failed or user not found`r`nPlease verify that EMC is installed!" -ForegroundColor Red
     Pause
     }
clear
            


}

function Prochat {

$NewRubyUser = Read-Host "Username "
$KCStation = Read-Host "Office / Station Example(KC1111): "
                       
#do {
#$ColbernOrRalph = Write-Host "
#---------------------ProChat Location-------------------
#1 = Colbern
#2 = Ralph Powell
#--------------------------------------------------------" -ForegroundColor Green
#$NewRubyChoice2 = read-host -prompt "Select number & press enter "
#} until ($NewRubyChoice2 -eq "1" -or $NewRubyChoice2 -eq "2")


            Get-ADUser $NewRubyUser | Enable-ADAccount -Verbose

            Get-ADUser $NewRubyUser | Set-ADUser -Office $KCStation -Verbose
            Get-ADUser $NewRubyUser | Set-ADUser -EmailAddress ($NewRubyUser + '@callruby.com').ToLower();
            Get-ADUser $NewRubyUser | Set-ADUser -CannotChangePassword $false -Verbose

            Get-ADUser $NewRubyUser | Move-ADObject -TargetPath 'OU=ProChats,OU=Ruby Users,DC=ruby,DC=local' -Verbose
            Add-ADGroupMember -Identity ('KC' + ' ' + 'Team-1280071548') $NewRubyUser -Verbose
            Add-ADGroupMember -Identity TeamRuby $NewRubyUser -Verbose
            Add-ADGroupMember -Identity SEC_Sophos_Standard_WebFiltering $NewRubyUser -Verbose
            Add-ADGroupMember -Identity Logi_NonReceptionist -Members $NewRubyUser -Verbose
            Add-ADGroupMember -Identity APP_SSO_ProChatsOnline_Access -Members $NewRubyUser -Verbose

            #$isNonRecep = Get-ADUser -Identity $NewRubyUser | Select -Property Title


            #if ($NewRubyChoice2 -eq "1") {Add-ADGroupMember -Identity ColbernTeam-12047575227 $NewRubyUser -Verbose}

            #if ($NewRubyChoice2 -eq "2") {Add-ADGroupMember -Identity RalphPowellTeam-1-1422955080 $NewRubyUser -Verbose}

            $StopWatch = [System.Diagnostics.StopWatch]::StartNew()

 

Function Test-Command ($Command)



{

    Try

    {

        Get-command $command -ErrorAction Stop

        Return $True

    }

    Catch [System.SystemException]

    {

        Return $False

    }

}

 

IF (Test-Command "Get-Mailbox") {Write-Host "Exchange cmdlets already present"}

Else {

    $CallEMS = ". '$env:ExchangeInstallPath\bin\RemoteExchange.ps1'; Connect-ExchangeServer PRLEMC01 -ClientApplication:ManagementShell "

    Invoke-Expression $CallEMS


$stopwatch.Stop()
$msg = "`n`nThe script took $([math]::round($($StopWatch.Elapsed.TotalSeconds),2)) seconds to execute..."
Write-Host $msg
$msg = $null
$StopWatch = $null

}


Enable-RemoteMailbox -Identity $NewRubyUser -RemoteRoutingAddress ($NewRubyUser + '@callruby.com') -Verbose


Write-Host "$NewRubyUser Enabled!
===========================================================================================
| Run the following command in Exchange Management Shell to verify remotemailbox user     |
| Command: Get-RemoteMailbox -Identity %Username%                                         |
===========================================================================================" -ForegroundColor Yellow
pause
clear

            
}

function EnableSales {

                $NewRubyUser = Read-Host "Username"                

                $EnableSalesNewManager = Read-Host -Prompt "Please provide Manager's username (Example: first.lastname) "

                $EnableSalesStation = Read-Host -Prompt "Please provide station number "

                Get-ADUser $NewRubyUser | Enable-ADAccount -Verbose
                Get-ADUser $NewRubyUser | Move-ADObject -TargetPath 'OU=PDX Sales,OU=Ruby Users,DC=ruby,DC=local' -Verbose
                Get-ADUser $NewRubyUser | Set-ADUser -Manager $EnableSalesNewManager -Verbose
                Get-ADUser $NewRubyUser | Set-ADUser -Office $EnableSalesStation -Verbose

                Add-ADGroupMember -Identity FoxTower -Members $NewRubyUser -Verbose
                Add-ADGroupMember -Identity FS_TC_Audio_Agreements_RW -Members $NewRubyUser -Verbose
                Add-ADGroupMember -Identity GrowingLeaders -Members
                Add-ADGroupMember -Identity ('ISR' + ' ' + 'Team-1739626290') -Members $NewRubyUser -Verbose
                Add-ADGroupMember -Identity Logi_NonReceptionist -Members $NewRubyUser -Verbose
                Add-ADGroupMember -Identity NonReceptionists -Members $NewRubyUser -Verbose
                Add-ADGroupMember -Identity pearl_sql_users -Members $NewRubyUser -Verbose
                Add-ADGroupMember -Identity PublicFolderAccess -Members $NewRubyUser -Verbose
                Add-ADGroupMember -Identity SALES -Members $NewRubyUser -Verbose
                Add-ADGroupMember -Identity ('Sales' + ' ' + 'and' + ' ' + 'Onboarding-1-488285678') -Members $NewRubyUser -Verbose
                Add-ADGroupMember -Identity ('Sales' + ' ' + 'Team') -Members $NewRubyUser -Verbose
                Add-ADGroupMember -Identity Sales_Sales_SEC -Members $NewRubyUser -Verbose
                Add-ADGroupMember -Identity SP_SalesAssociates -Members $NewRubyUser -Verbose
                Add-ADGroupMember -Identity SEC_Sophos_Standard_WebFiltering $NewRubyUser -Verbose

                clear

                Write-Host "New Sales Ruby Enabled!" -ForegroundColor Green

                $StopWatch = [System.Diagnostics.StopWatch]::StartNew()

 

Function Test-Command ($Command)



{

    Try

    {

        Get-command $command -ErrorAction Stop

        Return $True

    }

    Catch [System.SystemException]

    {

        Return $False

    }

}

 

IF (Test-Command "Get-Mailbox") {Write-Host "Exchange cmdlets already present"}

Else {

    $CallEMS = ". '$env:ExchangeInstallPath\bin\RemoteExchange.ps1'; Connect-ExchangeServer PRLEMC01 -ClientApplication:ManagementShell "

    Invoke-Expression $CallEMS


$stopwatch.Stop()
$msg = "`n`nThe script took $([math]::round($($StopWatch.Elapsed.TotalSeconds),2)) seconds to execute..."
Write-Host $msg
$msg = $null
$StopWatch = $null

}


Enable-RemoteMailbox -Identity $NewRubyUser -RemoteRoutingAddress ($NewRubyUser + '@callruby.com') -Verbose


Write-Host "$NewRubyUser Enabled!
===========================================================================================
| Run the following command in Exchange Management Shell to verify remotemailbox user     |
| Command: Get-RemoteMailbox -Identity %Username%                                         |
===========================================================================================" -ForegroundColor Yellow
pause
clear

}

function EnableAdminGeneric {

 
                $NewRubyUser = Read-Host "Username"
                $EnableAdminNewManager = Read-Host -Prompt "Please provide Manager's username (Example: first.lastname) "
                $EnableAdminStation = Read-Host -Prompt "Please provide station number "

                Get-ADUser $NewRubyUser | Enable-ADAccount -Verbose
                Get-ADUser $NewRubyUser | Set-ADUser -Office $EnableAdminStation -Verbose
                Get-ADUser $NewRubyUser | Set-ADUser -Manager $EnableAdminNewManager -Verbose
                Add-ADGroupMember -Identity SEC_Sophos_Standard_WebFiltering $NewRubyUser -Verbose

                # .net form for a multi select box for adding new admin to multiple selected security groups or dristros
                Add-Type -AssemblyName System.Windows.Forms
                Add-Type -AssemblyName System.Drawing

                
                $form = New-Object System.Windows.Forms.Form
                $form.Text = 'Security Groups and distro memberships'
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
                $label.Text = 'Please Select Groups to add new Admin too:'
                $form.Controls.Add($label)

                $listBox = New-Object System.Windows.Forms.Listbox
                $listBox.Location = New-Object System.Drawing.Point(15,60)
                $listBox.Size = New-Object System.Drawing.Size(350,500)
                

                $listBox.SelectionMode = 'MultiSimple'


                $getgroups = Get-ADGroup -SearchBase 'OU=Ruby Groups,DC=ruby,DC=local' -Filter 'ObjectClass -eq "group"' | 
                select Name | Sort-Object -Property Name
                

                foreach ($group in $getgroups) {
                                                [void] $listBox.Items.Add($group.name)
                                                }
                

                $listBox.Height = 150
                $form.Controls.Add($listBox)
                $form.Topmost = $true

                $result = $form.ShowDialog()

                     if ($result -eq [System.Windows.Forms.DialogResult]::OK) {Write-Host "AD Groups successfully queried!" -ForegroundColor Green} else {Write-Host "Something went wrong!" -ForegroundColor Red}
                     ForEach-Object {
    
                                        foreach ($group in $listBox.SelectedItems)
                                        {
                                            try{
                                                Add-ADGroupMember -Identity $group $NewRubyUser -Verbose
                                                }
                                                catch {
                                                       Write-Host "$group does not exist! $NewRubyUser not added!" -ForegroundColor Red
                                                       }
                                        
                                        }
        
                                    }
                
                # .Net form to add user to a selected OU in AD
                Add-Type -AssemblyName System.Windows.Forms
                Add-Type -AssemblyName System.Drawing

                $form = New-Object System.Windows.Forms.Form
                $form.Text = 'AD UO'
                $form.Size = New-Object System.Drawing.Size(300,200)
                $form.StartPosition = 'CenterScreen'

                $OKButton = New-Object System.Windows.Forms.Button
                $OKButton.Location = New-Object System.Drawing.Point(340,15)
                $OKButton.Size = New-Object System.Drawing.Size(75,23)
                $OKButton.Text = 'OK'
                $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $form.AcceptButton = $OKButton
                $form.AutoSize = $true
                $form.Controls.Add($OKButton)

                $CancelButton = New-Object System.Windows.Forms.Button
                $CancelButton.Location = New-Object System.Drawing.Point(420,15)
                $CancelButton.Size = New-Object System.Drawing.Size(75,23)
                $CancelButton.Text = 'Cancel'
                $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $form.CancelButton = $CancelButton
                $form.Controls.Add($CancelButton)

                $label = New-Object System.Windows.Forms.Label
                $label.Location = New-Object System.Drawing.Point(10,20)
                $label.Size = New-Object System.Drawing.Size(280,20)
                $label.Text = 'Please Select an OU to add new Admin too:'
                $form.Controls.Add($label)

                $listBox2 = New-Object System.Windows.Forms.Listbox
                $listBox2.Location = New-Object System.Drawing.Point(10,40)
                $listBox2.Size = New-Object System.Drawing.Size(300,20)
                $listBox2.ScrollAlwaysVisible = $true

                [void] $listBox2.Items.Add('OU=Billing,OU=Ruby Users,DC=ruby,DC=local')
                [void] $listBox2.Items.Add('OU=BVN Admin,OU=Ruby Users,DC=ruby,DC=local')
                [void] $listBox2.Items.Add('OU=BVN HR,OU=Ruby Users,DC=ruby,DC=local')
                [void] $listBox2.Items.Add('OU=BVN Information Technology,OU=Ruby Users,DC=ruby,DC=local')
                [void] $listBox2.Items.Add('OU=BVN Office,OU=Ruby Users,DC=ruby,DC=local')
                [void] $listBox2.Items.Add('OU=BVN RS,OU=Ruby Users,DC=ruby,DC=local')
                [void] $listBox2.Items.Add('OU=BVN Sales,OU=Ruby Users,DC=ruby,DC=local')
                [void] $listBox2.Items.Add('OU=FXT Scheduling Ninjas,OU=Ruby Users,DC=ruby,DC=local')
                [void] $listBox2.Items.Add('OU=Contractors,OU=Ruby Users,DC=ruby,DC=local')
                [void] $listBox2.Items.Add('OU=Marketing,OU=Ruby Users,DC=ruby,DC=local')
                [void] $listBox2.Items.Add('OU=Misc Hourly Rubys,OU=Ruby Users,DC=ruby,DC=local')
                [void] $listBox2.Items.Add('OU=PDX Administrative,OU=Ruby Users,DC=ruby,DC=local')
                [void] $listBox2.Items.Add('OU=PDX Client Happiness,OU=Ruby Users,DC=ruby,DC=local')
                [void] $listBox2.Items.Add('OU=PDX HR,OU=Ruby Users,DC=ruby,DC=local')
                [void] $listBox2.Items.Add('OU=PDX Information Technology,OU=Ruby Users,DC=ruby,DC=local')
                [void] $listBox2.Items.Add('OU=PDX RS,OU=Ruby Users,DC=ruby,DC=local')
                [void] $listBox2.Items.Add('OU=PDX Sales,OU=Ruby Users,DC=ruby,DC=local')
                [void] $listBox2.Items.Add('OU=ProChats,OU=Ruby Users,DC=ruby,DC=local')
                [void] $listBox2.Items.Add('OU=Product,OU=Ruby Users,DC=ruby,DC=local')
                [void] $listBox2.Items.Add('OU=RX,OU=Ruby Users,DC=ruby,DC=local')
                [void] $listBox2.Items.Add('OU=Talent,OU=Ruby Users,DC=ruby,DC=local')
                


                $listBox2.Height = 150
                $listBox2.Width = 500
                $form.Controls.Add($listBox2)
                $form.Topmost = $true

                $result = $form.ShowDialog()

                if ($result -eq [System.Windows.Forms.DialogResult]::OK) {Write-Host "AD OU's Queried Successfully!" -ForegroundColor Green} else {Write-Host "Something went wrong!" -ForegroundColor Red}

                Write-Host "Moved $NewRubyUser to selected OU!" -ForegroundColor Green
                $listBox2.SelectedItem
                Get-ADUser $NewRubyUser | Move-ADObject -TargetPath $listBox2.SelectedItem
                Remove-Variable -Name listBox2 -Force -Verbose


Function Test-Command ($Command)



{

    Try

    {

        Get-command $command -ErrorAction Stop

        Return $True

    }

    Catch [System.SystemException]

    {

        Return $False

    }

}

 

IF (Test-Command "Get-Mailbox") {Write-Host "Exchange cmdlets already present"}

Else {

    $CallEMS = ". '$env:ExchangeInstallPath\bin\RemoteExchange.ps1'; Connect-ExchangeServer PRLEMC01 -ClientApplication:ManagementShell "

    Invoke-Expression $CallEMS


$stopwatch.Stop()
$msg = "`n`nThe script took $([math]::round($($StopWatch.Elapsed.TotalSeconds),2)) seconds to execute..."
Write-Host $msg
$msg = $null
$StopWatch = $null

}

try {
    Enable-RemoteMailbox -Identity $NewRubyUser -RemoteRoutingAddress ($NewRubyUser + '@callruby.com') -Verbose
    }
    catch {Write-Host "$NewRubyUser remote mailbox already enabled!" -ForegroundColor Yellow}

Write-Host "$NewRubyUser Enabled!
===========================================================================================
| Run the following command in Exchange Management Shell to verify remotemailbox user     |
| Command: Get-RemoteMailbox -Identity %Username%                                         |
===========================================================================================" -ForegroundColor Yellow
pause
clear

}



do {

clear

do {    
$BVTorprl = Write-Host "
-------------------What type of Ruby?-------------------
 1 = Pearl Recep
 2 = Beaverton Recep
 3 = Chat
 4 = Sales
 5 = Admin (Generic)
 6 = Exit
--------------------------------------------------------" -ForegroundColor Green
$NewRubyChoice = read-host -prompt "Select number & press enter "

clear

} until ($NewRubyChoice -eq "1" -or $NewRubyChoice -eq "2" -or $NewRubyChoice -eq "3" -or $NewRubyChoice -eq "4" -or $NewRubyChoice -eq "5" -or $NewRubyChoice -eq "6")

Switch ($NewRubyChoice) {
"1" {Pearly}
"2" {Beaver}
"3" {Prochat}
"4" {EnableSales}
"5" {EnableAdminGeneric}
}
pause

clear

} until ($NewRubyChoice = "6")



}

# Assign O365 licenses to users
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

clear

$ErrorActionPreference = 'silentlycontinue'

Do {

Do {
Write-Host "
----------------Onboard New Ruby---------------" -BackgroundColor Blue -ForegroundColor Green
Write-Host "
 1 = Enable Rubys Generated from ED-E
 2 = Assign o365 licenses
 3 = Sales
 4 = Receptionist Pearl
 5 = Receptionist Beaverton
 6 = exit
-----------------------------------------------" -ForegroundColor Green 
$choice1 = read-host -prompt "Select number & press enter "

} until ($choice1 -eq "1" -or $choice1 -eq "2" -or $choice1 -eq "3" -or $choice1 -eq "4" -or $choice1 -eq "5" -or $choice1 -eq "6")

Switch ($choice1) {
"1" {Enable_Receptionist}
"2" {Assigno365Licenses}
"3" {OnBrdSales}
"4" {VerNamePrl}
"5" {VerNameBvt}
}
pause

clear

} until ($choice1 -eq "6")



}

# Modify users for additional access
function EditRuby {

    #SubFunctions

    # Provide Extended hours access
    function ExtHours {
    
    $RubyExtUser = Read-Host -Prompt "Please provide the username (Example: first.lastname): "

    Add-ADGroupMember -Identity ('Client' + ' ' + 'Happiness') $RubyExtUser -Verbose
    Add-ADGroupMember -Identity pearl_ch_users $RubyExtUser -Verbose
    Add-ADGroupMember -Identity ('Extended' + ' ' + 'Hours') $RubyExtUser -Verbose

    $ICadminPath = "C:\Program Files (x86)\Interactive Intelligence\ServerManagerApps\IAShellU.exe"

    Start-Process -FilePath $ICadminPath

    Write-Host "
+=====================================================================================+
| Ruby added to required security groups and distro lists for Extended hours Access.  |
| Please make sure you add the user to the following Roles and Workgroups in IC Admin |
|                                                                                     |
| Workgroups: Park Queue                                                              |
|                                                                                     |
| Roles: Extended Hours, Parker                                                       |
+=====================================================================================+" -ForegroundColor Green

    
    }

    # Migrate recep AD user to sales
    function NextAdvSales {
    
    $NextAdvSalesUser = Read-Host -Prompt "Please provide the username (Example: first.lastname): "
    $NextAdvSalesPrevTeam = Read-Host -Prompt "Please provide name of previous receptionists team (Example:Novas,Blazing,Etc....): "
    $NextSalesManager = Read-Host -Prompt "Please provide new manager username (Example: first.lastname): "
    $NextSalesOfficeLoc = Read-Host -Prompt "New Station number / Office Location: "

    clear

    Get-ADUser $NextAdvSalesUser | Move-ADObject -TargetPath 'OU=PDX Sales,OU=Ruby Users,DC=ruby,DC=local' -Verbose
    Get-ADUser $NextAdvSalesUser | Set-ADUser -Manager $NextSalesManager -Verbose
    Get-ADUser $NextAdvSalesUser | Set-ADUser -Office $NextSalesOfficeLoc -Verbose

       
        
        
                Remove-ADGroupMember -Identity $NextAdvSalesPrevTeam -Members $NextAdvSalesUser -confirm:$false -Verbose
                Remove-ADGroupMember -Identity ALL_Receptionists_Prl_Sec -Members $NextAdvSalesUser -confirm:$false -Verbose
                Remove-ADGroupMember -Identity ALL_Receptionists_Bvt_Sec -Members $NextAdvSalesUser -confirm:$false -Verbose
                Remove-ADGroupMember -Identity All_Receptionists_SEC -Members $NextAdvSalesUser -confirm:$false -Verbose
                Remove-ADGroupMember -Identity PearlTeam -Members $NextAdvSalesUser -confirm:$false -Verbose
                Remove-ADGroupMember -Identity beavertonteam -Members $NextAdvSalesUser -confirm:$false -Verbose
                Remove-ADGroupMember -Identity Exclaimer_group1 -Members $NextAdvSalesUser -confirm:$false -Verbose
            
          

                Add-ADGroupMember -Identity FoxTower -Members $NextAdvSalesUser -Verbose
                Add-ADGroupMember -Identity FS_TC_Audio_Agreements_RW -Members $NextAdvSalesUser -Verbose
                Add-ADGroupMember -Identity GrowingLeaders -Members
                Add-ADGroupMember -Identity ('ISR' + ' ' + 'Team-1739626290') -Members $NextAdvSalesUser -Verbose
                Add-ADGroupMember -Identity Logi_NonReceptionist -Members $NextAdvSalesUser -Verbose
                Add-ADGroupMember -Identity NonReceptionists -Members $NextAdvSalesUser -Verbose
                Add-ADGroupMember -Identity pearl_sql_users -Members $NextAdvSalesUser -Verbose
                Add-ADGroupMember -Identity PublicFolderAccess -Members $NextAdvSalesUser -Verbose
                Add-ADGroupMember -Identity SALES -Members $NextAdvSalesUser -Verbose
                Add-ADGroupMember -Identity ('Sales' + ' ' + 'and' + ' ' + 'Onboarding-1-488285678') -Members $NextAdvSalesUser -Verbose
                Add-ADGroupMember -Identity ('Sales' + ' ' + 'Team') -Members $NextAdvSalesUser -Verbose
                Add-ADGroupMember -Identity Sales_Sales_SEC -Members $NextAdvSalesUser -Verbose
                Add-ADGroupMember -Identity SP_SalesAssociates -Members $NextAdvSalesUser -Verbose

clear
            
Write-Host "
|=========================================|
|                                         |
|  $($NextAdvSalesUser) Moved to sales!   |
|                                         |
|=========================================|
" -ForegroundColor Green
 
 pause
        
 clear     
    
    }

    # Migrate Next Adventure CH (Includes HP mailbox)
    function OnBrdCH {

Write-Host "Onboard new CH
==============================" -BackgroundColor Blue -ForegroundColor Green
Write-Host ""

$ErrorActionPreference = 'silentlycontinue'

# Setting Variables to create user and add them to the proper
# OU and Groups
$username = Read-Host -Prompt "User Name"
$officeloc = Read-Host -Prompt "Please provide the office loc / station: "
$department = 'PSHM Hourly'
$manager = Read-Host -Prompt "Manager name (Account Name of manager)"
$company = 'Ruby Receptionists'
$fstlst = ($firstname + ' ' + $lastname)
$firstinidotlastname = (($FirstName.Substring(0,1)) + ($LastName))
$hpmainname = ('HP' + ' ' + $firstname + ' ' + $lastname)
$hpAdGroup = ('HP' + $RecepTeam)
$hpUser = ('HP' + $firstname + '.' + $lastname)
$hpsur = ('HP ' + $lastname)
$hpgiv = $lastname
$hpOU = 'OU=Hurdles & Praise,OU=Non-Human Users,OU=Ruby Users - System\, Non-Human\, External,DC=ruby,DC=local'
$hpprincname = ('HP' + $firstname + '.' + $lastname + '@callruby.com')


# Modifys Multiple attributes on newly created user
# Based on the input at the beginning of the script
# Then moves user to the proper OU
Get-ADUser $username | Set-ADUser -Office $officeloc
Get-ADUser $username | Set-ADUser -ScriptPath 'ch.bat'
Get-ADUser $username | Set-ADUser -Title $jobtitle
Get-ADUser $username | Set-ADUser -Manager $manager
Get-ADUser $username | Set-ADUser -Department $department
Get-ADUser $username | Move-ADObject -TargetPath 'OU=PDX Client Happiness,OU=Ruby Users,DC=ruby,DC=local' -Verbose

# Adds user to associated AD groups
Add-ADGroupMember -Identity CH $username -confirm:$false 
Add-ADGroupMember -Identity CHommunity $username -confirm:$false 
Add-ADGroupMember -Identity CHTestUsers $username -confirm:$false 
Add-ADGroupMember -Identity Client Happiness $username -confirm:$false 
Add-ADGroupMember -Identity Fox Tower $username -confirm:$false 
Add-ADGroupMember -Identity FS_CHFolders_RW $username -confirm:$false 
Add-ADGroupMember -Identity FS_Paperless_RW $username -confirm:$false 
Add-ADGroupMember -Identity FS_PST_Archive_Hello_RW $username -confirm:$false 
Add-ADGroupMember -Identity FS_PST_Archive_Hub_RW $username -confirm:$false 
Add-ADGroupMember -Identity FS_TC_Audio_Agreements_RW $username -confirm:$false 
Add-ADGroupMember -Identity NonReceptionists $username -confirm:$false 
Add-ADGroupMember -Identity pearl_ch_users $username -confirm:$false 
Add-ADGroupMember -Identity pearl_credit_card_batch $username -confirm:$false 
Add-ADGroupMember -Identity pearl_credit_card_process $username -confirm:$false 
Add-ADGroupMember -Identity pearl_sql_users $username -confirm:$false -force:$true
Add-ADGroupMember -Identity pearl_ic_users $username -confirm:$false 
Add-ADGroupMember -Identity PearlReportingServicesUsers $username -confirm:$false  
Add-ADGroupMember -Identity SPARK_CH $username -confirm:$false 
Add-ADGroupMember -Identity TeamRuby $username -confirm:$false 
Add-ADGroupMember -Identity SP_ProblemSolversHappinessMakers $username -confirm:$false 
Add-ADGroupMember -Identity calendlee $username -confirm:$false 
 

# Remove user from associated AD groups
Remove-ADGroupMember -Identity Exclaimer_Group1 $username -confirm:$false 
Remove-ADGroupMember -Identity SPARK_BtownReceptionists $username -confirm:$false 
Remove-ADGroupMember -Identity SPARK_Receptionists $username -confirm:$false 
Remove-ADGroupMember -Identity SPARK_PearlReceptionists $username -confirm:$false 
Remove-ADGroupMember -Identity Receptionists $username -confirm:$false 
Remove-ADGroupMember -Identity SPARK_PearlReceptionists $username -confirm:$false 

clear

Write-Host "CH Next Adventure changed in AD and Exhange!" -BackgroundColor Blue -ForegroundColor Green
pause

clear
}
    
    # Add a user to multiple AD groups
    Function AddUserToMultipleGroups {

                $MassAddUser = Read-Host "What user "

                Add-Type -AssemblyName System.Windows.Forms
                Add-Type -AssemblyName System.Drawing

                
                $form = New-Object System.Windows.Forms.Form
                $form.Text = 'Security Groups and distro memberships'
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
                $label.Text = 'Please select groups to add user too:'
                $form.Controls.Add($label)

                $listBox = New-Object System.Windows.Forms.Listbox
                $listBox.Location = New-Object System.Drawing.Point(15,60)
                $listBox.Size = New-Object System.Drawing.Size(350,500)
                

                $listBox.SelectionMode = 'MultiSimple'


                $getgroups = Get-ADGroup -SearchBase 'OU=Ruby Groups,DC=ruby,DC=local' -Filter 'ObjectClass -eq "group"' | select Name | Sort-Object -Property Name
                

                foreach ($group in $getgroups) {
                                                [void] $listBox.Items.Add($group.name)
                                                }
                

                $listBox.Height = 150
                $form.Controls.Add($listBox)
                $form.Topmost = $true

                $result = $form.ShowDialog()

                     if ($result -eq [System.Windows.Forms.DialogResult]::OK) {"True"} else {"Something went wrong!"}
                     ForEach-Object {
    
                                        foreach ($group in $listBox.SelectedItems)
                                        {
                                            try{
                                                Write-Host "Adding $($MassAddUser) to $($group)" -ForegroundColor Yellow
                                                Add-ADGroupMember -Identity $group $MassAddUser -Verbose
                                                }
                                                catch {
                                                       Write-Host "$group does not exist! $MassAddUser not added!" -ForegroundColor Red
                                                       }
                                        
                                        }
        
                                    }

}


clear

$ErrorActionPreference = 'silentlycontinue'

Do {

Do {

clear

Write-Host "
+========Additional Resp / Next Adventure=======+
|                                               |
| 1 = Extended Hours Access                     |
| 2 = Next Adventure: Sales                     |
| 3 = Next Adventure: CH                        |
| 4 = Add user to multiple AD groups            |
| 5 = Rename Ruby                               |
| 6 = exit                                      |
|                                               |
+===============================================+" -ForegroundColor Green 
$choice1 = read-host -prompt "Select number & press enter "

} until ($choice1 -eq "1" -or $choice1 -eq "2" -or $choice1 -eq "3" -or $choice1 -eq 
         "4" -or $choice1 -eq "5" -or $choice1 -eq "6")

Switch ($choice1) {
"1" {ExtHours}
"2" {NextAdvSales}
"3" {OnBrdCH}
"4" {AddUserToMultipleGroups}
"5" {}
}
pause

clear

} until ($choice1 -eq "6")


}

# Departing Ruby Scripts
function DepartRuby {

# Departs Ruby's in AD and exchange via a CSV
    function DepartRubyCsv {
            
            $error.Clear()
            $ErrorActionPreference = 'silentlycontinue'
            $csvFile = "\\prlfs01\iswizards\uud\departed.csv"
            $disabledUsersOU = "OU=Disabled Users,DC=ruby,DC=local"

             Import-Csv $csvFile | ForEach-Object {
	
                        # Disable the account
	                    Disable-ADAccount -Identity $_.UserName -confirm:$false -Verbose
                        
                        # Retrieve the user object and MemberOf property
	                    $user = Get-ADUser -Identity $_.UserName -Properties MemberOf

                        # Null Office Field
                        Get-ADUser -Identity $_.UserName | Set-ADUser -Office $null -Verbose
                         
	
                                # Move user object to disabled users OU
                                try {Get-ADUser $user | Move-ADObject -TargetPath $disabledUsersOU -confirm:$false -Verbose}
                                    catch {Write-Host "Warning: user does not exist in Active Directory!" -ForegroundColor Yellow}
    
                                    # Remove all group memberships (will leave Domain Users as this is NOT in the MemberOf property returned by Get-ADUser)
	                                foreach ($group in $user | Select -ExpandProperty MemberOf)
	                                        {Remove-ADGroupMember -Identity $group -Members $_.UserName -confirm:$false -Verbose}
                                                   }
            if ($error.Count -eq 0) {
                                     Write-Host "There was "  $error.count  " errors found during runtime!" -ForegroundColor Green
                                     }
                else {
                      Write-Host $error.Count + " errors occured. Please review Superscript Logs for details!" -ForegroundColor Red
                      }
            
            }

    function SetAccountExpireDate {

clear
    
$ErrorActionPreference = 'silentlycontinue'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object Windows.Forms.Form -Property @{
    StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
    Size          = New-Object Drawing.Size 640, 480
    Text          = 'Set AD account expiration Date / Time / Timezone'
    Topmost       = $true
    Autoscale     = $true
    }
$BTFDimg = [System.Drawing.Image]::FromFile("$PSScriptRoot\Source\BTFMD3.jpg")
$form.StartPosition = "CenterScreen"
$FormIcon = New-Object System.Drawing.Icon ("$PSScriptRoot\Source\Papirus-Team-Papirus-Devices-Computer.ico")
$form.BackgroundImage = $BTFDimg
$form.Icon = $FormIcon

$CalendarLabel = New-Object System.Windows.Forms.Label
$CalendarLabel.Text = "Calendar Date"
$CalendarLabel.Font = "Georgia"
$CalendarLabel.ForeColor = "Teal"
$CalendarLabel.Location = "48,92"
$CalendarLabel.Height = 20
$CalendarLabel.Width = 80
$form.Controls.Add($CalendarLabel)

$calendar = New-Object Windows.Forms.MonthCalendar -Property @{
    Location          = New-Object Drawing.Point 50, 120
    ShowTodayCircle   = $false
    MaxSelectionCount = 1
}
$calendar.CanFocus = $true
$form.Controls.Add($calendar)

$OKButton = New-Object Windows.Forms.Button -Property @{
    Location     = New-Object Drawing.Point 380, 360
    Size         = New-Object Drawing.Size 75, 23
    Text         = 'OK'
    DialogResult = [Windows.Forms.DialogResult]::OK
}
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object Windows.Forms.Button -Property @{
    Location     = New-Object Drawing.Point 460, 360
    Size         = New-Object Drawing.Size 75, 23
    Text         = 'Cancel'
    DialogResult = [Windows.Forms.DialogResult]::Cancel
}
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$UsernameLabel = New-Object System.Windows.Forms.Label
$UsernameLabel.Text = “Username”
$UsernameLabel.Font = "Georgia"
$UsernameLabel.ForeColor = "Teal"
$UsernameLabel.Location = “360, 90”
$UsernameLabel.Height = 20
$UsernameLabel.Width = 70
$form.Controls.Add($UsernameLabel)

$UsernameTextBox = New-Object System.Windows.Forms.TextBox
$UsernameTextBox.Location = “360, 120”
$UsernameTextBox.Width = “150”
$form.Controls.Add($UsernameTextBox)

$TimePickerLabel = New-Object System.Windows.Forms.Label
$TimePickerLabel.Text = “Time”
$TimePickerLabel.Font = "Georgia"
$TimePickerLabel.ForeColor = "Teal"
$TimePickerLabel.Location = “360, 160”
$TimePickerLabel.Height = 22
$TimePickerLabel.Width = 40
$form.Controls.Add($TimePickerLabel)

$TimePicker = New-Object System.Windows.Forms.TextBox
$TimePicker.Location = “360, 195”
$TimePicker.Width = “150”
$form.Controls.Add($TimePicker)

$TimeZoneSelectLabel = New-Object System.Windows.Forms.Label
$TimeZoneSelectLabel.Text = “Time Zone”
$TimeZoneSelectLabel.Font = "Georgia"
$TimeZoneSelectLabel.ForeColor = "Teal"
$TimeZoneSelectLabel.Location = “360, 230”
$TimeZoneSelectLabel.Height = 22
$TimeZoneSelectLabel.Width = 68
$form.Controls.Add($TimeZoneSelectLabel)

$TimeZoneSelect = New-Object System.Windows.Forms.ComboBox
$TimeZoneSelect.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$TimeZoneSelect.AutoCompleteCustomSource.Add("System.Windows.Forms");
$TimeZoneSelect.AutoCompleteCustomSource.AddRange(("PST", "MST", "CST", "EST"));
$TimeZoneSelect.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend;
$TimeZoneSelect.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::CustomSource;
$TimeZoneSelect.AllowDrop = $true
$TimeZoneSelect.Location = "360,260"
$TimeZoneSelect.Size = "110,50"
$TimeZoneSelect.Items.Add("PST")
$TimeZoneSelect.Items.Add("MST")
$TimeZoneSelect.Items.Add("CST")
$TimeZoneSelect.Items.Add("EST")
$form.Controls.Add($TimeZoneSelect)

$result = $form.ShowDialog()

clear

if ($result -eq [Windows.Forms.DialogResult]::OK) {
    $date = $calendar.SelectionStart
    $time = $TimePicker.Text
    $user = $UsernameTextBox.Text
    $timezone = $TimeZoneSelect.SelectedItem
    $playExplosion = New-Object System.Media.SoundPlayer
    $playExplosion.SoundLocation = "$PSScriptRoot\Source\cat.wav"
    

    try {
    Set-ADAccountExpiration -Identity $user -DateTime "$($date.ToShortDateString()) $time" -Verbose
    $playExplosion.PlaySync()
    Write-Host "`n$user set to expire on Date selected: $($date.ToShortDateString()) at $($time.ToString()) $timezone.
Please Allow 1 hour for change to replicate." -ForegroundColor Green
    
    }
    
    Catch {
           Write-Host "Invalid Date time provided, or bad username" -ForegroundColor Red
           }
    

    }
    if($result -eq [Windows.Forms.DialogResult]::CANCEL){
      Write-Host "`nExpiration date cancelled" -ForegroundColor Yellow
      
      }
    
}

clear

$ErrorActionPreference = 'silentlycontinue'

Do {

Do {
Write-Host "
------------------Depart Ruby------------------" -BackgroundColor Blue -ForegroundColor Green
Write-Host "
 1 = Depart Ruby's via CSV
 2 = Set AD Account Expiration Date
 3 = exit
-----------------------------------------------" -ForegroundColor Green 
$choice1 = read-host -prompt "Select number & press enter "

} until ($choice1 -eq "1" -or $choice1 -eq "2" -or $choice1 -eq "3")

Switch ($choice1) {
"1" {DepartRubyCsv}
"2" {SetAccountExpireDate}
}
pause

clear

} until ($choice1 -eq "3")

}

# Creates Ruby Signature for receptionist on any system
function RubySig {
    
    Function RubySigLoc {
    
    clear

    $ErrorActionPreference = 'SilentlyContinue'

Write-Host "Email Signature Creator!
========================" -BackgroundColor Blue -ForegroundColor Green


#Creates Variables for use within modifying the signature to the receptionist
$firstName = Read-Host -Prompt "First name"
$lastName = Read-Host -Prompt "Last name"
$UserName = ($firstName + '.' + $lastName )
$firstLast = ($firstName + ' ' + $lastName)
$firstInitial = $firstName.Substring(0,1)
$firstIntLast = ($firstName[0] + $lastName);
$AdminEmail = ($firstIntLast + '@callruby.com').ToLower();

# Gathers input for the hostname, and Job title, and whether recep type sig or admin
$whatSystem = Read-Host -Prompt "What system"
$RubyRole = Read-Host -Prompt "What role / title"
$RecepOrAdmin = Write-Host "
-------------------What type of Ruby?-------------------" -BackgroundColor Blue -ForegroundColor Green
Write-Host "
 1 = Receptionist
 2 = Admin
------------------------------------------------------" -ForegroundColor Green
$RCChoice = read-host -prompt "Select number & press enter "

# Test to see if path exists / Host is powered on
$doesDIRExist = Test-Path -Path \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures
$doesDIRExist
    
    # If path exists / system is powered on will remove all folders and files within the appdata sig folder
    if ($doesDIRExist -eq $true) { 
        Remove-Item -Path \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\*.* -Recurse -Verbose -Confirm:$false
        rd -Path \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures -verbose -force -Confirm:$false
        md -Path \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures -verbose -force -Confirm:$false
        }
        # If path doesnt exist it will create the directory
        else {
              Write-Host "$doesDirExist does not exist, creating directory!" -ForegroundColor Yellow
              md -Path \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures -verbose -force -Confirm:$false
    }

    if ($RCChoice -eq "1") {
    clear

#Moves the signature to the hostname specified
Copy ($PSScriptRoot + '\Signatures\Ruby.htm') ('\\' + $whatSystem + '\C$\Users\' + $UserName + '\Appdata\Roaming\Microsoft\Signatures') -Verbose -Force
Copy ($PSScriptRoot + '\Signatures\Ruby.rtf') ('\\' + $whatSystem + '\C$\Users\' + $UserName + '\Appdata\Roaming\Microsoft\Signatures') -Verbose -Force
Copy ($PSScriptRoot + '\Signatures\Ruby.txt') ('\\' + $whatSystem + '\C$\Users\' + $UserName + '\Appdata\Roaming\Microsoft\Signatures') -Verbose -Force

#Creates a directory fold on the host specified to hold .xml files for signature and themes
md -Path \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Rubyfiles -Confirm:$false -Verbose

Copy ($PSScriptRoot + '\Signatures\Rubyfiles\*.*') ('\\' + $whatSystem + '\C$\Users\' + $UserName + '\Appdata\Roaming\Microsoft\Signatures\Rubyfiles') -Verbose -Force

Rename-Item -Path \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Rubyfiles -NewName Ruby_files -Verbose -Confirm:$false

# Sets the content for the email signature
(Get-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.htm) -replace 'xUser',$firstLast | 
 Set-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.htm -Verbose
(Get-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.htm) -replace 'Team',$RubyRole | 
 Set-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.htm -Verbose
(Get-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.htm) -replace 'email','staff@callruby.com' | 
 Set-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.htm -Verbose

(Get-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.rtf) -replace 'xUser',$firstLast | 
 Set-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.rtf -Verbose
(Get-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.rtf) -replace 'Team',$RubyRole | 
 Set-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.rtf -Verbose
(Get-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.rtf) -replace 'email','staff@callruby.com'| 
 Set-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.rtf -Verbose

(Get-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.txt) -replace 'xUser',$firstLast | 
 Set-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.txt -Verbose
(Get-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.txt) -replace 'Team',$RubyRole | 
 Set-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.txt -Verbose
(Get-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.txt) -replace 'email','staff@callruby.com' | 
 Set-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.txt -Verbose

Copy-Item  ($PSScriptRoot + '\OutlookFonts\*.reg') ( '\\' + $whatSystem + '\c$\users\Public') -Verbose -force -Confirm:$false
Copy-Item ($PSScriptRoot + '\OutlookFonts\ImportReg.txt') ('\\' + $whatSystem + '\c$\users\' + $UserName + '\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\ImportReg.bat') -Verbose -force -Confirm:$false

clear

Write-Host "Email signature configured for $($firstLast) on system $($whatSystem)" -BackgroundColor Blue -ForegroundColor Green

pause

clear
}

if ($RCChoice -eq "2") {
clear

#Moves the signature to the hostname specified
Copy ($PSScriptRoot + '\Signatures\Ruby.htm') ('\\' + $whatSystem + '\C$\Users\' + $UserName + '\Appdata\Roaming\Microsoft\Signatures') -Verbose -Force
Copy ($PSScriptRoot + '\Signatures\Ruby.rtf') ('\\' + $whatSystem + '\C$\Users\' + $UserName + '\Appdata\Roaming\Microsoft\Signatures') -Verbose -Force
Copy ($PSScriptRoot + '\Signatures\Ruby.txt') ('\\' + $whatSystem + '\C$\Users\' + $UserName + '\Appdata\Roaming\Microsoft\Signatures') -Verbose -Force

#Creates a directory fold on the host specified to hold .xml files for signature and themes
md -Path \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Rubyfiles -Confirm:$false -Verbose

Copy ($PSScriptRoot + '\Signatures\Rubyfiles\*.*') ('\\' + $whatSystem + '\C$\Users\' + $UserName + '\Appdata\Roaming\Microsoft\Signatures\Rubyfiles') -Verbose -Force

Rename-Item -Path \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Rubyfiles -NewName Ruby_files -Verbose -Confirm:$false

# Sets the content for the email signature
(Get-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.htm) -replace 'xUser',$firstLast | 
 Set-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.htm -Verbose
(Get-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.htm) -replace 'Team',$RubyRole | 
 Set-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.htm -Verbose
(Get-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.htm) -replace 'email',($UserName + '@callruby.com') | 
 Set-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.htm -Verbose

(Get-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.rtf) -replace 'xUser',$firstLast | 
 Set-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.rtf -Verbose
(Get-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.rtf) -replace 'Team',$RubyRole | 
 Set-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.rtf -Verbose
(Get-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.rtf) -replace 'email',($UserName + '@callruby.com')| 
 Set-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.rtf -Verbose

(Get-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.txt) -replace 'xUser',$firstLast | 
 Set-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.txt -Verbose
(Get-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.txt) -replace 'Team',$RubyRole| 
 Set-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.txt -Verbose
(Get-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.txt) -replace 'email',($UserName + '@callruby.com') | 
 Set-Content \\$($whatSystem)\c$\Users\$($UserName)\AppData\Roaming\Microsoft\Signatures\Ruby.txt -Verbose

clear

Write-Host "Email signature configured for $($firstLast) on system $($whatSystem)" -BackgroundColor Blue -ForegroundColor Green

pause

clear
}
}

    Function RubySigOWA {

    cls
#location of signature template
$LocationOfSignature= "\\prlfs01\PDQ_Shared_DB\Repository\Signature_Creation\OWA_Signature.htm"


# See if connected to Exchange Online

$AmIConnectedToExchangeOnline= Get-PSSession | where {$_.ConfigurationName -eq "Microsoft.exchange"}


    IF ($AmIConnectedToExchangeOnline -eq $null -or $AmIConnectedToExchangeOnline.State -eq "Closed")
        {
        #Connect to Exchange Online
        
        $userCredentials= Get-Credential -message "Enter Exchange Online Email/Credentials to continue" 
        
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredentials -Authentication Basic -AllowRedirection
        
        Import-PSSession $Session
         

        }


 #Import the template for the signature

 $importSignatureTemplate= Get-Content $LocationOfSignature

 
 # Type in the email address or UPN to the user.

    DO {

        clear

        $Stop= "No"
      $OWA_Sigs_To_Update= Read-host "Type in the UPN (usually matches the primary email address) of the employee"
      
      IF ($OWA_Sigs_To_Update -like "*@callruby.com")
            {
            $Stop= "Yes"
            }
                else 
                {
                Write-Warning "Email address does not contain @callruby.com domain...please try again."
                }

        }

        Until ($stop -eq "Yes" )

$GetADUserInfo= Get-ADUser -Filter * -Properties title,department,enabled,emailaddress | where {$_.title -notlike $Null -and $_.department -notlike $null}


$UserDetails= $GetADUserInfo | where {$_.UserPrincipalName -eq $OWA_Sigs_To_Update}

IF ($UserDetails -eq $null)
    {
    Write-Warning "UPN couldn't be found in AD.... Exiting..."
        Start-Sleep -Seconds 4
        Remove-PSSession $Session
        Exit
    }
        Else 
        {Write-host "
Found $($UserDetails.Name) in Active Directory." -ForegroundColor Green}
        
         
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
 #        $Signature | Out-File "C:\Users\chris.georgeson\Desktop\Office365 info\OutlookOnline_Recep_rollout\sigtest.htm"

        }

Remove-PSSession $Session 


write-host "
Signature should be created. Please check OWA to verify." -ForegroundColor Green
pause 

clear
    }

    clear

$ErrorActionPreference = 'silentlycontinue'

Do {

Do {
Write-Host "
----------------Email Signature Tool---------------" -BackgroundColor Blue -ForegroundColor Green
Write-Host "
 1 = 0365 Local Installation Signature Creation
 2 = Outlook Online Signature Creation
 3 = exit
---------------------------------------------------" -ForegroundColor Green 
$choice1 = read-host -prompt "Select number & press enter "

} until ($choice1 -eq "1" -or $choice1 -eq "2" -or $choice1 -eq "3")

Switch ($choice1) {
"1" {RubySigLoc}
"2" {RubySigOWA}
}
pause

clear

} until ($choice1 -eq "3")

}

# Find a windows product key on a specific workstation
function Find_productKey {

function Get-ProductKey {
     <#   
    .SYNOPSIS   
        Retrieves the product key and OS information from a local or remote system/s.
         
    .DESCRIPTION   
        Retrieves the product key and OS information from a local or remote system/s. Queries of 64bit OS from a 32bit OS will result in 
        inaccurate data being returned for the Product Key. You must query a 64bit OS from a system running a 64bit OS.
        
    .PARAMETER Computername
        Name of the local or remote system/s.
         
    .NOTES   
        Author: Boe Prox
        Version: 1.1       
            -Update of function from http://powershell.com/cs/blogs/tips/archive/2012/04/30/getting-windows-product-key.aspx
            -Added capability to query more than one system
            -Supports remote system query
            -Supports querying 64bit OSes
            -Shows OS description and Version in output object
            -Error Handling
     
    .EXAMPLE 
     Get-ProductKey -Computername Server1
     
    OSDescription                                           Computername OSVersion ProductKey                   
    -------------                                           ------------ --------- ----------                   
    Microsoft(R) Windows(R) Server 2003, Enterprise Edition Server1       5.2.3790  bcdfg-hjklm-pqrtt-vwxyy-12345     
         
        Description 
        ----------- 
        Retrieves the product key information from 'Server1'
    #>         
    [cmdletbinding()]
    Param (
        [parameter(ValueFromPipeLine=$True,ValueFromPipeLineByPropertyName=$True)]
        [Alias("CN","__Server","IPAddress","Server")]
        [string[]]$Computername = $Env:Computername
    )
    Begin {   
        $map="BCDFGHJKMPQRTVWXY2346789" 
    }
    Process {
        ForEach ($Computer in $Computername) {
            Write-Verbose ("{0}: Checking network availability" -f $Computer)
            If (Test-Connection -ComputerName $Computer -Count 1 -Quiet) {
                Try {
                    Write-Verbose ("{0}: Retrieving WMI OS information" -f $Computer)
                    $OS = gwmi -ComputerName $Computer Win32_OperatingSystem -ErrorAction Stop                
                } Catch {
                    $OS = New-Object PSObject -Property @{
                        Caption = $_.Exception.Message
                        Version = $_.Exception.Message
                    }
                }
                Try {
                    Write-Verbose ("{0}: Attempting remote registry access" -f $Computer)
                    $remoteReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$Computer)
                    If ($OS.OSArchitecture -eq '64-bit') {
                        $value = $remoteReg.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion").GetValue('DigitalProductId4')[0x34..0x42]
                    } Else {                        
                        $value = $remoteReg.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion").GetValue('DigitalProductId')[0x34..0x42]
                    }
                    $ProductKey = ""  
                    Write-Verbose ("{0}: Translating data into product key" -f $Computer)
                    for ($i = 24; $i -ge 0; $i--) { 
                      $r = 0 
                      for ($j = 14; $j -ge 0; $j--) { 
                        $r = ($r * 256) -bxor $value[$j] 
                        $value[$j] = [math]::Floor([double]($r/24)) 
                        $r = $r % 24 
                      } 
                      $ProductKey = $map[$r] + $ProductKey 
                      if (($i % 5) -eq 0 -and $i -ne 0) { 
                        $ProductKey = "-" + $ProductKey 
                      } 
                    }
                } Catch {
                    $ProductKey = $_.Exception.Message
                }        
                $object = New-Object PSObject -Property @{
                    Computername = $Computer
                    ProductKey = $ProductKey
                    OSDescription = $os.Caption
                    OSVersion = $os.Version
                } 
                $object.pstypenames.insert(0,'ProductKey.Info')
                $object
            } Else {
                $object = New-Object PSObject -Property @{
                    Computername = $Computer
                    ProductKey = 'Unreachable'
                    OSDescription = 'Unreachable'
                    OSVersion = 'Unreachable'
                }  
                $object.pstypenames.insert(0,'ProductKey.Info')
                $object                           
            }
        }
    }
} 

$whatSystem = Read-Host -Prompt "What system? "

gsv -Name RemoteRegistry -ComputerName $whatSystem  | Set-Service -StartupType Automatic -Status Running -Verbose
Get-ProductKey -Computername $whatSystem  | Out-GridView -Title "Domain Systems and assigned product keys"

}

# Clears Software Distribution folder on systems
function ClearCabs {

clear

$Machine = read-host "Type in the Computer Name"

$isWinRMactivated = gsv -ComputerName $Machine -Name WinRM | Select -Property Status
$isWinRMactivated

if ($isWinRMactivated.Status -eq 'Running') {
write-host "WinRM already configured!" -ForegroundColor Yellow
    }
if ($isWinRMactivated.Status -eq 'Stopped') {
       
    # Checks Version of windows and Enables WinRM   
    gsv -ComputerName $Machine -Name RemoteRegistry | Start-Service
    gsv -ComputerName $Machine -Name RemoteRegistry | Set-Service -StartupType Automatic
    
    $remotekeysyst = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Machine)
    $remotekeypath = $remotekeysyst.OpenSubKey("SOFTWARE\\Microsoft\Windows NT\CurrentVersion")
    $remoteKeyValue = $remotekeypath.GetValue("ProductName")

    $remotekeysyst
    $remotekeypath
    $remoteKeyValue

        if ($remoteKeyValue -eq "Windows 7 Professional") { 

            Write-Host "Windows 7 Professional detected on $Machine! Restarting system to setup WinRM" -ForegroundColor Red

Write-Output "
@Echo Off
echo y | winrm quickconfig
pause
" | Out-File ('\\' + $Machine + '\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\Winrm.bat') -Force -Verbose

            gsv -ComputerName $Machine -Name WinRM | Start-Service -Verbose | Set-Service -StartupType Automatic
            Restart-Computer -ComputerName $Machine -Wait -For Wmi -Force -Verbose
            Invoke-Command -ComputerName $Machine -ScriptBlock {Enable-PSRemoting -ErrorAction Continue -Force} -Verbose

Write-Host "
|=================================================|
| Once System Restarts WinRM should be configured |
| after the user logs in. If not please run the   |
| following command on the users system.          |
| WinRM quickconfig                               |
|=================================================|
"

}
             Else {Write-Host "$Machine is not a windows 7 system, continuing to next step..." -ForegroundColor Yellow }
             }


$windowsUpdateService = 'wuauserv'
$trustedInstallerService = 'trustedinstaller'
$RemoteRegistryService = 'WinRM'

function Set-ServiceState 
{
    [CmdletBinding()]
    param(
        [string]$ComputerName,
        [string]$ServiceName
    )
 
    Write-Verbose "Evaluating $ServiceName on $ComputerName."
 
    [string]$WaitForIt = ""
    [string]$Verb = ""
    [string]$Result = "FAILED"
 
    $svc = gsv -computername $ComputerName -name $ServiceName
    Switch ($svc.status) {
        'Stopped' {
            Write-Verbose "[$ServiceName] is currently Stopped. Starting."
            $Verb = "start"
            $WaitForIt = 'Running'
            $svc.Start()
        }
        'Running' {
            Write-Verbose "[$ServiceName] is Running. Stopping."
            $Verb = "stop"
            $WaitForIt = 'Stopped'
            $svc.Stop()
        }
        default {
            Write-Verbose "$ServiceName is $($svc.status). Taking no action."
        }
    }
 
    if ($WaitForIt -ne "") {
        Try { # For some reason, we cannot use -ErrorAction after the next statement:
            $svc.WaitForStatus($WaitForIt,'00:02:00')
        } Catch {
            Write-Warning "After waiting for 2 minutes, $ServiceName failed to $Verb."
        }
 
        $svc = (gsv -computername $ComputerName -name $ServiceName)
        if ($svc.status -eq $WaitForIt) {
            $Result = 'SUCCESS'
        }
 
        Write-Host "$Result - $ServiceName on $ComputerName is $svc.status"
        Write-Host "{0} - {1} on {2} is {4}  $Result, $ServiceName, $ComputerName, $svc.status"
    }
 
} 
 
# stop update service
Set-ServiceState -ComputerName $Machine -ServiceName $windowsUpdateService -Verbose
 
#removes temp files and renames software distribution folder
gsv -ComputerName $Machine -Name wuauserv | Stop-Service
Remove-Item \\$Machine\c$\windows\temp\*.cab -recurse -Verbose

$testSDPath = Test-Path -Path \\$Machine\c$\windows\SoftwareDistribution.old
$testSDPath

$testTempDIR = Test-Path \\$Machine\c$\Temp
$testTempDIR

    if ($testTempDIR -eq $true) {
    
    Write-Host "\\$Machine\c$\Temp already exists!" -ForegroundColor Yellow
    }
    else{
    New-Item -Path \\$Machine\c$\Temp -ItemType Directory -Force -Verbose
    }

if ($testSDPath -eq $true) {
    
    Write-Host "\\$Machine\c$\windows\SoftwareDistribution.old already exists!" -ForegroundColor Yellow
    
    Set-Location -Path \\$Machine\c$\windows\SoftwareDistribution.old

    $getDIRs = Get-ChildItem -Path ".\*.*" -Recurse | Move-Item -Destination "\\$Machine\c$\Temp" -Force -Verbose
    $getDIRs

    Write-Host "Clearing \\$Machine\c$\Temp folder!" -ForegroundColor Yellow
    Remove-Item -Path \\$Machine\c$\Temp -Recurse -Force -Verbose

    Write-Host "Removing the folder and subfolders \\$Machine\c$\windows\SoftwareDistribution.old" -ForegroundColor Yellow
    Remove-Item -Path \\$Machine\c$\windows\SoftwareDistribution.old -Recurse -Verbose -Force   
   
}
    else {
        
        write-host "Renaming SoftwareDistribution folder to SoftwareDistribution.old" -ForegroundColor Yellow
        Rename-Item \\$Machine\c$\windows\SoftwareDistribution \\$Machine\c$\windows\SoftwareDistribution.old -Force -Verbose
        Set-Location -Path \\$Machine\c$\windows\SoftwareDistribution.old
    
    $getDIRs = Get-ChildItem -Path ".\*.*" -Recurse | Move-Item -Destination "\\$Machine\c$\Temp" -Force -Verbose
    $getDIRs
    
    Write-Host "Clearing \\$Machine\c$\Temp folder!" -ForegroundColor Yellow
    Remove-Item -Path \\$Machine\c$\Temp -Recurse -Force -Verbose

    Write-Host "Removing the folder and subfolders \\$Machine\c$\windows\SoftwareDistribution.old" -ForegroundColor Yellow
    Remove-Item -Path \\$Machine\c$\windows\SoftwareDistribution.old -Recurse -Verbose -Force   
    }
 
#restarts update service
Set-ServiceState -ComputerName $Machine -ServiceName $windowsUpdateService

 
#stops trustedinstaller service
gsv -ComputerName $Machine -Name $trustedInstallerService | Stop-Service -Force -Verbose 
 

#removes cab files from trustedinstaller
remove-item \\$Machine\c$\windows\Logs\CBS\* -recurse -Verbose

 
#restarts trustedinstaller service
Set-ServiceState -ComputerName $Machine -ServiceName $trustedInstallerService


#rebuilds cab files from WSUS
Write-Output "
wuauclt.exe /detectnow
" >> \\$Machine\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\CheckUpdates.bat -Verbose

#Removes Winrm bat
$TestWinRMPath = Test-Path -Path ('\\' + $Machine + '\C$\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\winrm.bat')
$TestWinRMPath

if ($TestWinRMPath -eq $true) {
    
    Write-Host "Removing \\$Machine\C$\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\winrm.bat" -ForegroundColor Yellow
    Remove-Item -Path ('\\' + $Machine + '\C$\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\winrm.bat')

}
    else {
        
        write-host "\\$Machine\C$\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\winrm.bat does not exist" -ForegroundColor Yellow 

    }

    Set-ServiceState -ComputerName $Machine -ServiceName $windowsUpdateService

    Set-Location -Path $PSScriptRoot
    
    Write-Host "SoftwareDistribution on $Machine Cleared!" -ForegroundColor Green

    pause
    clear

}

function Errors {

$error | Out-GridView -Title Error Log

}


# ============================================= #
#                  Help Menu                    #
# ============================================= #

function Hlp {

function PrcKillH {

clear

Write-Host "

======================================================================
|	            SuperScript Verson 19.0 - Help                  |
======================================================================

Script - {Process Killer}


1. Process Killer Script: This Script allows you Remotely and locally
Kill active processes.

	- Step1: Specify the hostname you wish to target

	- Step2: Specify the process you wish to kill

	- To Exit just hit enter once its done

	or

	- Specify an additional host to target and process to kill

=====================================================================
|								    |
=====================================================================

"

}

function StartRdcH {

clear

Write-Host "

======================================================================
|	            SuperScript Verson 19.0 - Help                    |
======================================================================

Script - {Start RDC script}


1. Start RDC Script: This script start RDC

	- Step1: Once Windows remote desktop connection opens
		 Specify the hostname you wish to RDC into

	- Step2: Type the credentials that you wish to login
		 with.


=====================================================================
|								    |
=====================================================================
"

}

function StartCmdH {

clear

Write-Host "

======================================================================
|	            SuperScript Verson 19.0 - Help                    |
======================================================================

Script - {Start CMD script}


1. Start CMD Script: This script will start a windows command line
		     interphase.

=====================================================================
|								    |
=====================================================================

"

}

function AsstTrkH {

clear

Write-Host "

======================================================================
|	            SuperScript Verson 19.0 - Help                  |
======================================================================
Script - {Assist Tracker script}


1. Assist Tracker Script: This script will deploy Assist Tracker			  

	-Step1: Specify the host
	
   After this script runs it should place a shortcut on the
   Receptonists desktop

   Hint: Also make sure the receptionist has the following
         Workgroups and Roles assigned in IC admin

         1. Role: Assister
         2. Workgroups: Assist Queue, Future Assist Requests


=====================================================================
|								    |
=====================================================================

"

}

function UpdateRosH {

clear

Write-Host "

======================================================================
|	            SuperScript Verson 19.0 - Help                   |
======================================================================
Script - {Update RosOS script}


1. Update RosOS Script: 

	   This script will update / deploy RosOS to any specified 
	   system. This will also create a shortcut for the user.

	-Step1: Specify the host

    -Step2: Specify station number


=====================================================================
|								    |
=====================================================================

"

}

function OnbrdRecH {

clear

Write-Host "

======================================================================
|	            SuperScript Verson 19.0 - Help                   |
======================================================================

Script - {Create New Receptionist scripts}


1. Create New Receptionists Script: (No longer used! - ED-E now
   creates users.)

      Options 2 3 or 4 within the New Ruby Function

      These scripts will Create new users within AD.

      Script will perform the following

	* create a new user

	* enable and set the following - Password, Manager, Job 
	  Title, Email address, Company, Department, Password never
	  Expires, User cannot change password, Logon Script
	
	* Moves the new user to the proper OU

	* Adds the new user to the associated AD Groups

    * o365 user mailboxes will automatically be created 
      within one hour.
	 
	-Step1: Specify first name
	
	-Step2: Specify last name

	-Step3: Provide Managers username(Account) 
		
		Example: firstname.lastname
	
	-Step4: Specify team (Receptionist Functions Only)
		
		Example: Novas, Racers, Enchanted, etc...

	-Step5: Create a password


2. Enable Rubys Generated from ED-E

      Option 1 under the New Ruby function - Use this only!
      All users are now created via ED-E and just neet to be
      Enabled, moved to the propper ou, and added to their
      required AD groups.

     * 1 - Enable new Pearl Receptionist
            
            - Provide the new receptionists username
              Ex: firstname.lastname
            
            - Provide team name Ex: Ramblers, Novas...etc

            - Provide Managers username Ex: firstname.lastname

            - Provide Station number (Ex: prl1111, BVT1111, FXT1111)

     * 2 - Enable new Pearl Receptionist

            - Same as Enable new Pearl receptionist

     * 3 - Enable ProChat Receptionist

            - Provide the new ProChat username Ex: firstname.lastname

            - Select what office 1 Colbern or 2 Ralph powell

     * 4 - Enable New Sales user

            - Provide the new sales associates username
              Ex: firstname.lastname

            - Provide the managers username 
              Ex: firsname.lastname

            - Provide the station number 
              Ex: FXT1111, prl11111, BVT1111

     * 5 - Enable New Admin (Generic)
          
           - Provide new admin username
             Ex: firstname.lastname

           - Provide the managers username
             Ex: firstname.lastname

           - Provide the station number
             Ex: FXT1111, prl11111, BVT1111

           - Next a form will pop up asking what AD groups
             you want to add the new admin to.
             (Select each on by clicking on the names and
              then clicking on OK)

           - Next another form will pop up asking what OU
             to move the new admin to
             (Select the desired OU and click ok)


=====================================================================
|								    |
=====================================================================

"

}

function SigCreatorH {

clear

Write-Host "

======================================================================
|	            SuperScript Verson 19.0 - Help                   |
======================================================================
    
Script - Setup Outlook Signature

   Description: This script will create email signatures for both 
                local outlook installations and Outlook Online.

     * 1 - 0365 Local Installation Signature Creation

           - Provide the Receptionists first name

           - Provide the Receptionists last name

           - Provide the Hostname of the Receptionists system

           - Provide the receptionists job title

     * 2 - Outlook Online Signature Creation
      
           - Provide the receptionists full email address
             Ex: firstname.lastname@callruby.com

=====================================================================
|								    |
=====================================================================

"

}

function ClearCabsH {

clear

Write-Host "

======================================================================
|	            SuperScript Verson 19.0 - Help                   |
======================================================================
    
Script - Clear Software Distribution folder on systems

   Description: This will clear windows updateds download cache
                On remote systems within a domain environment

     * 1 - Clear Software Distribution folder on systems
          
           - Provide the name of the computer (Hostname) that
             you wish to run the script on Ex:PRL1234567D

=====================================================================
|								    |
=====================================================================

"


}

function FindPrdKeyH {

clear

Write-Host "

======================================================================
|	            SuperScript Verson 19.0 - Help                   |
======================================================================
    
Script - Get Windows Product Key (WinRM must be enabled)

   Description: This will query windows product keys used on
                  remote system

     * 1 - Get Windows Product Key (WinRM must be enabled)
          
           - Provide the name of the computer (Hostname) that
             you wish to run the script on Ex:PRL1234567D

=====================================================================
|								    |
=====================================================================

"


}

function ModRubyH {

clear

Write-Host "

======================================================================
|	            SuperScript Verson 19.0 - Help                   |
======================================================================
    
Script - Modify Ruby (Additional Responsibilities / Next Adv.)

   Description: This script is for Adding additional 
                responsibilities or for seting a 
                receptionist up for thier Next
                Adventure.

     * 1 - Extended Hours Access
          
           - Provide the Receptionists Username Ex:fistname.lastname

           Note: Make sure after you run this that you add the
                 receptionist to the following roles and workgroups
                 in IC Admin

                 a. Roles: Parker, Extended Hours
                 b. Workgroups: Park Queue

     * 2 - Next Adventure: Sales
           
           - Provide the Username Ex: firstname.lastname

           - Provide Receptionists previous team name 
             Ex: Novas, Racers, Etc

           - Provide the new managers username
             Ex: firstname.lastname

           - Provide new station number / Office Location
             Ex: prl1111, FXT1111, BVT1111

     * 3 - Next Adventure: CH
           
           - Provide the username of the receptionist going to CH
             Ex: firstname.lastname

           - Provide the new managers username
             Ex: firstname.lastname

=====================================================================
|								    |
=====================================================================

"

}


Write-Host "-----Superscript Help-----" -ForegroundColor Green -BackgroundColor Blue
Write-Host ""
Write-Host ""
Write-Host "What script are you wanting to know about?.

"

$ErrorActionPreference = 'silentlycontinue'

Do {

Do {
clear
Write-Host "
-------------------------Help Menu-------------------------" -BackgroundColor Blue -ForegroundColor Green
Write-Host "
 1 = Process Killer Script
 2 = Start Remote Desktop Connection
 3 = Windows Command Line
 4 = Deploy Assist Tracker
 5 = Update RubyOS
 6 = Create new Receptionist in AD and Exchange
 7 = Modify Ruby (Additional Responsibilities / Next Adv.)
 8 = Setup Outlook Signature
 9 = Clear Software Distribution folder on systems
10 = Get Windows Product Key (WinRM must be enabled)
11 = exit
-----------------------------------------------------------" -ForegroundColor Green 
$choice1 = read-host -prompt "Select number & press enter "

} until ($choice1 -eq "1" -or $choice1 -eq "2" -or $choice1 -eq "3" -or $choice1 -eq "4" -or $choice1 -eq "5" -or $choice1 -eq 
"6" -or $choice1 -eq "7" -or $choice1 -eq "8" -or $choice1 -eq "9" -or $choice1 -eq "10" -or $choice1 -eq "11")

Switch ($choice1) {
"1" {PrcKillH}
"2" {StartRdcH}
"3" {StartCmdH}
"4" {AsstTrkH}
"5" {UpdateRosH}
"6" {OnbrdRecH}
"7" {ModRubyH}
"8" {SigCreatorH}
"9" {ClearCabsH}
"10"{FindPrdKeyH}
}
pause

clear

} until ($choice1 -eq "11")


}


# ============================================= #
#                   Main Menu                   #
# ============================================= #

function MainMenu {
clear
Write-Host "___________________Admin SuperScript__________________" -BackgroundColor Blue -ForegroundColor Green
Write-Host ""

$ErrorActionPreference = 'silentlycontinue'

Do {

Do {
clear
Write-Host "
-------------------SuperScript Menu----------------------" -BackgroundColor Blue -ForegroundColor Green
Write-Host "
 1 = Process Killer Script
 2 = Start Remote Desktop Connection
 3 = Windows Command Line
 4 = Deploy Assist Tracker
 5 = Update RubyOS
 6 = Clear Microsoft Teams Roaming Cache
 7 = Setup Outlook(Local) or O365 Email Signatures
 8 = Find a windows product key (On specified Systems)
 9 = New Ruby
10 = Modify Ruby (Additional Responsibilities / Next Adv.)
11 = Depart Ruby
12 = Clear Software Distribution folder on systems
13 = Remove Mailbox Rules
14 = Clear SIP Dmp Logs
15 = View Error Logs (Not Working Yet)
16 = Help
17 = Exit
---------------------------------------------------------" -ForegroundColor Green 
$choice1 = read-host -prompt "Select a number & press enter "

} until ($choice1 -eq "1" -or $choice1 -eq "2" -or $choice1 -eq 
"3" -or $choice1 -eq "4" -or $choice1 -eq "5" -or $choice1 -eq 
"6" -or $choice1 -eq "7" -or $choice1 -eq "8" -or $choice1 -eq 
"9" -or $choice1 -eq "10" -or $choice1 -eq "11" -or $choice1 -eq 
"12" -or $choice1 -eq "13" -or $choice1 -eq "14" -or $choice1 -eq 
"15" -or $choice1 -eq "16" -or $choice1 -eq "17") # -or $choice1 -eq "")

Switch ($choice1) {
"1" {StpPrc}
"2" {Invoke-Item (Start mstsc)}
"3" {Invoke-Item (Start cmd)}
"4" {AssistTrack}
"5" {UpdateRos}
"6" {ClearTeamsCache}
"7" {RubySig}
"8" {Find_productKey}
"9" {NewRuby}
"10"{EditRuby}
"11"{DepartRuby}
"12"{ClearCabs}
"13"{DeleteMailboxRule}
"14"{ClearSIPDMP}
#"" {}
#"" {}
#"" {}
#"" {}
#"" {}
#"" {}
#"" {}
#"" {}
#"" {}
#"" {}
#"" {}
#"" {}
#"" {}
"15"{Errors}
"16"{Hlp}
}
clear

} until ($choice1 -eq "17")

    Write-Host "Closing SuperScript" -BackgroundColor Blue -ForegroundColor Green
    pause

}


# ==============================LOGS================================ #
# logs sessions into a txt file at the root directory of the script  #
# ================================================================== #

function Logs {



$log = ($PSScriptRoot + '\Logs\Log.txt')
Invoke-Command -ComputerName prlfs01 -ScriptBlock {Get-SmbOpenFile | Where-Object -Property Path -Match "Log.txt" | Close-SmbOpenFile -Force} -Verbose
$ErrorActionPreference = 'silentlycontinue'
$renameLog = (Get-Date).ToString("MM-dd-yyyyhhmmss")
$newlogName = $renameLog + ".txt"

if ($log -eq $true) {

$ErrorActionPreference = 'silentlycontinue'
$renameLog = (Get-Date).ToString("MM-dd-yyyyhhmmss")
$newlogName = $renameLog + ".txt"
    
    
    copy-Item -Path $PSScriptRoot\Logs\Log.txt -Destination $PSScriptRoot\Logs\Log-Archive\$newlogName -Verbose -force
    Remove-Item -Path $PSScriptRoot\Logs\Log.txt -Force -Verbose

    Start-Transcript -path "$($log)" -append -Verbose

}

else {
     copy-Item -Path $PSScriptRoot\Logs\Log.txt -Destination $PSScriptRoot\Logs\Log-Archive\$newlogName -Verbose -force
     Remove-Item -Path $PSScriptRoot\Logs\Log.txt -Force -Verbose
     Start-Transcript -Path "$($log)" -Append -Verbose 
     }

}

Logs


# ============================================= #
#             . Source functions                #
# ============================================= #


$dotsourced = Get-ChildItem -Path \\prlfs01\ISWizards\SuperScript\DotSourcedScripts\*.ps1 -Verbose
$howManyDS = $dotsourced.count
Write-Host "Importing $($howManyDS) Functions via Dot Sourcing!" -ForegroundColor Green

if ($dotsourced -ne $null) {

    foreach ($ps1 in $dotsourced) {

            Write-Host "`nImporting $($ps1)" -ForegroundColor Green
            . $ps1
  
  }

}
 
    else {
          Write-Host "`nNo functions to import via dot source!" -ForegroundColor Yellow
          } 


MainMenu

Stop-Transcript | Out-Null
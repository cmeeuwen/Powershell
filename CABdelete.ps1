$Machine = read-host "Type in the Computer Name"

$isWinRMactivated = Get-Service -ComputerName $Machine -Name WinRM | Select-Object -Property Status
$isWinRMactivated

if ($isWinRMactivated.Status -eq 'Running') {
write-host "WinRM already configured!" -ForegroundColor Yellow
    }
if ($isWinRMactivated.Status -eq 'Stopped') {
       
    # Checks Version of windows and Enables WinRM   
    Get-Service -ComputerName $Machine -Name RemoteRegistry | Start-Service
    Get-Service -ComputerName $Machine -Name RemoteRegistry | Set-Service -StartupType Automatic
    
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

            Get-Service -ComputerName $Machine -Name WinRM | Start-Service -Verbose | Set-Service -StartupType Automatic
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
 
    $svc = Get-Service -computername $ComputerName -name $ServiceName
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
 
        $svc = (get-service -computername $ComputerName -name $ServiceName)
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
Get-Service -ComputerName $Machine -Name wuauserv | Stop-Service
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
#Set-ServiceState -ComputerName $Machine -ServiceName $trustedInstallerService
 Get-Service -ComputerName $Machine -Name $trustedInstallerService | Stop-Service -Force -Verbose 
 

#removes cab files from trustedinstaller
remove-item \\$Machine\c$\windows\Logs\CBS\* -recurse -Verbose

 
#restarts trustedinstaller service
Set-ServiceState -ComputerName $Machine -ServiceName $trustedInstallerService


#rebuilds cab files from WSUS
#Invoke-Command -ComputerName $Machine -ScriptBlock {Enable-PSRemoting -Force -Verbose}
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
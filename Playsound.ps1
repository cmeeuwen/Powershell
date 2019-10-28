Function Playit {

$RemoteComputer = Read-Host -Prompt "What System"


Start-BitsTransfer -Source C:\SuperscriptProd\Source\Thriller.wav -Destination \\$($RemoteComputer)\c$\ProgramData

If (Test-Connection -ComputerName $RemoteComputer -Quiet){    
      
     
     Invoke-Command -ComputerName $RemoteComputer -ScriptBlock{
            
            $getCompName = Get-WmiObject -Class win32_computersystem -Property Name | select Name 
            $boolFileFound = (Test-Path \\$($getCompName.name)\c$\ProgramData\Thriller.wav)
            
            if ($boolFileFound) {
                $soundplayer = New-Object System.Media.SoundPlayer -Property @{
                SoundLocation = 'c:\ProgramData\Thriller.wav'
                LoadTimeout = 10
                }  
                $soundplayer.playsync();

                return "Found file - " + $boolFileFound
            } else {
                return "File doesn't exist on remote machine"
            }
        }
        
} Else {Write-Host "$($RemoteComputer) is not online!" -ForegroundColor Red}

Remove-Item \\$($RemoteComputer)\c$\ProgramData\Thriller.wav -Verbose -Force

Remove-Variable RemoteComputer
}

Playit

[System.GC]::Collect()
Write-Host "====================================
| Setting IE home page settings.   |            
| Once Run you can import settings |
| into other browsers such as      |
| Chrome or Microsoft Edge         |
====================================" -ForegroundColor Yellow

$ErrorActionPreference = 'silentlycontinue'

    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Internet Explorer\Main' -Name "Start Page" -Value "https://rubylife.callruby.com/recepmain/default.aspx" -Force -Verbose

$propValue = @( 
             "https://logiweb01.callruby.com/RubyWebParts/rdPage.aspx?rdReport=TestReport2",
             "https://app.7geese.com/login/?next=/"
             )

    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Internet Explorer\Main' -Name "Secondary Start Pages" -PropertyType MultiString -Value $propValue -Force -Verbose

    Remove-Item $PSScriptRoot\SetRecepIEhome.lnk -Force -Verbose
    Remove-item $PSScriptRoot\SetRecepIEhomeV2.ps1 -Force -Verbose
    pause
    
Function UDSelfSupportDUO {

# Stop all running dashboards
#============================
try {
Get-UDDashboard | Stop-UDDashboard
}   
   catch {"No Dashboards Running!"}




#Dot Sourced Scripts used to authenticate login and Scatch PS1 for variables
#===========================================================================
$Endpoint1 = c:\SuperscriptProd\UDDUOSupport\Endpoint1.ps1

$loginForm = c:\SuperscriptProd\UDDUOSupport\LoginBasedOnSecGroup.ps1


# Login Dashboard Authentication Page
#====================================
$loginPage = New-UDLoginPage -AuthenticationMethod $loginForm -LoginFormBackgroundColor "white" -Title "Please Login" -PageBackgroundColor "lightgray" -Logo(
                                                                                                                                                             New-UDElement -Tag Center -Content {

                                                                                                                                                             $ImageURL = "https://rubylife.callruby.com/Documents/Ruby_Logo%c2%ae_RGB_Rogers.png"

                                                                                                                                                             New-UDImage -Url "$($ImageURL)"  -Height 150 -Width 200
                                                                                                                                                             }
                                                                                                                                                            ) -WelcomeText "DUO User Support Login"


# DUO Support Main Page
#======================
$pageEndUserDuoSupport = New-UDPage -name 'DUO Multi-factor Support' -Content {
    
    # Query all Users and Devices in DUO
    #===================================
    $DuoUsers = duoGetUser
    $DuoDevices = duoGetPhone


# DUO middle column cards for enrolling users with DUO and a basic how to enroll video #
#======================================================================================#
New-UDLayout -Columns 2 -Content {
    

    New-UDCard -Size large -BackgroundColor "lightgray" -Content{
    
    New-UDCollapsible -Items{

    # Enroll a new user with new mobile device
    #=========================================
    New-UDCollapsibleItem -Icon cloud -Title "Enroll" -Content {

    New-UDInput -Title "Duo Enroll A Device" -FontColor '#379629' -Endpoint {
    
    Param($DUOEnrollusername, $DUOEnrollemailaddress)
    
    if ($DUOEnrollusername -ne '' -and $DUOEnrollemailaddress -ne ''){
          
          Try {

               duoEnrollUser -dOrg 'prod' -username "$($DUOEnrollusername)" -email "$($DUOEnrollemailaddress)"
               
               Show-UDToast -MessageSize 30 -Message "Enrollment Sent!" -MessageColor '#f8faf7' -Duration 900 -Position center -BackgroundColor '#379629'

               }
             Catch {Show-UDToast -Message "User Already exists" -MessageColor '#f8faf7' -Duration 1000 -Position center -MessageSize 20 -BackgroundColor '#b81d23'}
         
         } Else {Show-UDToast -Message "Username and Email address must be valid!" -MessageColor '#f8faf7' -Duration 1000 -Position center -MessageSize 20 -BackgroundColor '#b81d23'}
    
    }

      } -BackgroundColor 'green'

    # Reactivate 
    New-UDCollapsibleItem -Icon _lock -Title "Activate / Re-Activate" -Content {

    New-UDInput -Title "Create Duo Reactivate Code" -FontColor '#379629' -Endpoint {
    
    Param($DUOReactivateNumber)
    
    if ($DUOReactivateNumber -ne ''){
          
          Try {
               $DUOGetPhoneID = duoGetPhone -dOrg 'prod' -number $DUOReactivateNumber
               $DUOPhoneID = $DUOGetPhoneID.phone_id
               $DUOGenerateQR = duoCreateActivationCode -dOrg 'prod' -phone_id $DUOPhoneID -activation_barcode
               $QRCode = $DUOGenerateQR.activation_barcode
               
               # Modal Window Containing new DUO QR Activation Code
               #===================================================
               Show-UDModal -Content{

               $HTMLLineBreak = New-UDElement -Tag br

               New-UDElement -Tag font -Attributes @{
                                                                 color = "Teal"
                                                                 size = 5 
                                                                 } -Content {
                                                                            "QR Activation Code"
                                                                            $HTMLLineBreak
                                                                            "For $($DUOGetPhoneID.Users.Username)"
                                                                            $HTMLLineBreak
                                                                            "Device number: $($DUOGetPhoneID.number)"
                                                                            }


               New-UDElement -Tag Center -Content {

                            New-UDElement -Tag iframe -Attributes @{
                                                                    src = "$($QRCode)"
                                                                    height = 250
                                                                    width = 250
                                                                    } 
                                                  
                                                  }
               
                                    }
               
               

               }
               Catch {
                      Show-UDToast -Message "Phone does not exist!" -MessageColor '#f8faf7' -Duration 1000 -Position center -MessageSize 20 -BackgroundColor '#b81d23'
                      }
         
         } Else {
                 Show-UDToast -Message "Number must be 10 digits!" -MessageColor '#f8faf7' -Duration 1000 -Position center -MessageSize 20 -BackgroundColor '#b81d23'
                 }
    
    }

      } -BackgroundColor 'green'
     
     } -Type Expandable -BackgroundColor 'lightgray'

     New-UDElement -Tag center -Content {
     New-UDImage -Url "https://rubylife.callruby.com/Documents/Ruby_Logo%c2%ae_RGB_Rogers.png" -Height 150 -Width 200

     New-UDElement -Tag br

     


     }
     New-UDElement -Tag font -Attributes @{
                                           color = "Teal"
                                           size = 5
                                           face = "Verdana"
                                          } -Content {

                                                      $CreateActivationCodeTextArray = @(
                                                                                         "Step 1: Left Click on Activate / Re-Activate",
                                                                                         "Step 2: Type in the phone number without any spaces",
                                                                                         "Step 3: Left Click on Submit",
                                                                                         "Step 4: Open the DUO app on your device and scan the QR code!"
                                                                                        )
                                                                                        foreach ($TextString in $CreateActivationCodeTextArray) {
                                                                                                                                                $TextString
                                                                                                                                                New-UDElement -Tag br
                                                                                                                                                }

                                                     }
    }

    # DUO Enroll instructions card #
    #==============================#
    New-UDCard -Size large -BackgroundColor "lightgray" -Content{

        New-UDElement -Tag br
        New-UDElement -Tag br
        
        New-UDLayout -Columns 2 -Content {

        New-UDElement -Tag font -Attributes @{
                                              face = "Century Gothic"
                                              color = "green"
                                              size = "3"
                                              } -Content {

        New-UDElement -Tag b -Content {


        $DuoEnrollTextArray = @(
                                "",
                                "This is how to enroll with Duo!",
                                "Step 1: Fill out the username and email",
                                "Step 2: Click on Submit",
                                "Step 3: Watch this video!",
                                "",
                                "It's that Easy!!!"
                                )

                                foreach ($ArrayText in $DuoEnrollTextArray){
                                                                            New-UDElement -Tag p -Content {$ArrayText} -Attributes @{align = "left"}
                                                                            New-UDElement -Tag br
                                                                            }

        
        
        }
        }
        
                        
        New-UDElement -Tag iframe -Attributes @{
                                                width = "560" 
                                                height = "315"
                                                src = "https://www.youtube.com/embed/EMj89Ulpx6c"
                                                frameborder = "2"
                                                allow = "accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" 
                                                allowfullscreen = "True"
                                                align = "right"
                                                }
       }                                                       
                             
    }

    


}



} -Icon lock_open -Title "DUO Support"


# Main Dashboard Markup
#======================
$MainFooter = New-UDFooter -BackgroundColor "lightgray" -Links @(
                                                                 New-UDLink -FontColor black -Text "Contact ITHelp | " -Url "mailto:ithelp@ruby.com"
                                                                 New-UDLink -FontColor black -Text "Ruby Life | " -Url "https://rubylife.callruby.com/Pages/Rubywelcome.aspx" -OpenInNewWindow
                                                                 )


$DashEndUser = New-UDDashboard -Title "DUO Help" -BackgroundColor "white" -Pages @(
                                                                                   $pageEndUserDuoSupport
                                                                                   ) -EndpointInitialization $init -LoginPage $loginPage -Footer $MainFooter -NavBarColor "teal" -FontColor "Black"


Start-UDDashboard -Port 10002 -Dashboard $DashEndUser -AutoReload -Force -AllowHttpForLogin -Endpoint @($Endpoint1) -AdminMode
clear

}

# Verify if DUO Module Exists
#============================
$doesDuoModuleExist = Get-Module | select name | Where-Object -Property name -Like "*Duo*"
Unblock-File C:\SuperscriptProd\Duo-PSModule-master\*.* -Confirm:$false

if ($doesDuoModuleExist -ne $null){
                                   Write-Host "Duo Module already imported"
                                   }
   Else {
         Import-Module c:\SuperscriptProd\Duo-PSModule-master\Duo.psd1 -Force
         }


UDSelfSupportDUO
function WizardDash {

# Stop all running dashboards
#============================
try {
Get-UDDashboard | Stop-UDDashboard
}   
   catch {"No Dashboards Running!"}


# Dashboard Main Page
#============================================#
$pageMain = New-UDPage -Name 'Home' -DefaultHomePage -Content { 
    
    New-UDElement -Tag div -Id parentDiv -Content {



        New-UDElement -Tag img -Attributes @{
                                            src = "https://rubylife.callruby.com/Documents/Ruby_Logo%c2%ae_RGB_Rogers.png"
                                            height = "182"
                                            width = "320"
                                            }
        New-UDElement -Tag img -Attributes @{
                                            src = "https://rubylife.callruby.com/SiteAssets/Pages/Brand-Story-and-Guidelines-Hub/dot%20line%20break.png"
                                            alignment = "center"
                                            }

                                            New-UDElement -Tag br

        
        New-UDElement -Tag b -Content {

            New-UDElement -Tag font -Attributes @{
                                                  face="Verdana"
                                                  color = "#112457"
                                                  } -Content {
                                                              " All Business is Personal"
                                                              New-UDElement -Tag br
                                                              New-UDElement -Tag br
                                                              }

            

        }
        
    }

    New-UDLayout -Columns 3 -Content {
    
    New-UDElement -Tag center -Content {
        
    New-UDElement -Tag iframe -Attributes @{
                                            width = "560" 
                                            height = "315"
                                            src = "https://www.youtube.com/embed/yuK_pRyYEkY"
                                            frameborder = "0"
                                            allow = "accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" 
                                            allowfullscreen = "True"
                                            align = "middle"
                                            }
    }
    
    New-UDElement -Tag center -Content {

    New-UDElement -Tag iframe -Attributes @{
                                            width = "560" 
                                            height = "315"
                                            src = "https://www.youtube.com/embed/D-_yAx25hCY"
                                            frameborder = "0"
                                            allow = "accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" 
                                            allowfullscreen = "True"
                                            align = "middle"
                                            }
    }

    New-UDElement -Tag center -Content {
    
    New-UDElement -Tag iframe -Attributes @{
                                            width = "560" 
                                            height = "315"
                                            src = "https://www.youtube.com/embed/ZKvrPRD74Ck"
                                            frameborder = "0"
                                            allow = "accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" 
                                            allowfullscreen = "True"
                                            align = "middle"
                                            }        

    }


    }

    #New-UDCard -Size large -Content {

    New-UDElement -Tag center -Content {

    New-UDElement -Tag iframe -Attributes @{
                                            height = 1080
                                            width = 2020
                                            src = "https://www.ruby.com"
                                            frameborder = 1
                                            scrolling = "auto"
                                            }
    
      }
    #}

    

} -Icon home -Title "Home"


# DUO MFA Support Page
#============================================#
$pageDuoSupport = New-UDPage -name 'DUO Multi-factor Support' -Content {

    $DuoUsers = duoGetUser
    $DuoDevices = duoGetPhone
    
    
# DUO User and Device Count Cards #
#=================================#
New-UDLayout -Columns 2 -Content {
    
    New-UDCard -Id "card1" -Title 'Duo User Count' -Text ($DuoUsers.Count) -BackgroundColor 'green' -FontColor 'white'
    New-UDCard -Id "card2" -Title 'Duo Device Count' -Text ($DuoDevices.Count) -BackgroundColor 'green' -FontColor 'white' 

}


# DUO middle column cards for enrolling users with DUO and a basic how to enroll video #
#======================================================================================#
New-UDLayout -Columns 2 -Content {
    
    # 1.DUO Enroll a device for new and existing DUO users #
    # 2.Create a new device and associate with user        #
    # 3.Remove a device from a user                        #
    #======================================================#
    New-UDCard -Size large -BackgroundColor "lightgray" -Content{
    
    New-UDCollapsible -Items{

    # Enroll a new user 
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
    
    # Create a phone and associate with user
    New-UDCollapsibleItem -Icon phone -Title "Create Device" -Content {
    
    New-UDInput -Title "Create New DUO device" -FontColor '#379629' -Endpoint {
    
    Param($DuoSearchUser, $DuoNewDevice)

        if ($DuoSearchUser -ne "" -or $DuoNewDevice -ne "") {
            
            Try {
                # Creates a device in DUO
                #========================
                $DuoCreateDevice = duoCreatePhone -dOrg 'prod' -number "+1$($DuoNewDevice)" -type mobile
                
                }
                Catch {Show-UDToast -Message "Invalid Parameters Provided!" -MessageColor '#fafafa' -BackgroundColor '#de2831' -Duration 2000 -MessageSize 40 -Position center}
        
            Try {
                # Query the new devices ID string
                #================================
                $DuoQueryPhoneID = duoGetPhone -number $DuoNewDevice
                
                $NewDuoPhoneID = $DuoQueryPhoneID.phone_id
                
                }
                Catch {Show-UDToast -Message "Device does not exist!" -MessageColor '#fafafa' -BackgroundColor '#de2831' -Duration 2000 -MessageSize 40 -Position center}
    
            Try {
                # Query the desired user's ID string
                #===================================
                $DuoQueryUserID = duoGetUser -username $DuoSearchUser
                
                $DuoUserSearchID = $DuoQueryUserID.user_id
                
                }
                Catch {Show-UDToast -Message "User does not exist!" -MessageColor '#fafafa' -BackgroundColor '#de2831' -Duration 2000 -MessageSize 40 -Position center}

            Try {
                # associates new device id with existing user id
                #===============================================
                $DuoAssocUserToDevice = duoAssocUserToPhone -dOrg 'prod' -phone_id "$($NewDuoPhoneID)" -user_id "$($DuoUserSearchID)"
                
                Show-UDToast -Message "Device $($DuoNewDevice) successfully created and associated with DUO user $($DuoSearchUser)" -MessageColor '#fafafa' -BackgroundColor '#379629' -Duration 2000 -MessageSize 20 -Position center
                
                }
                Catch{Show-UDToast -Message "Unable to associate device ID with user ID!" -MessageColor '#fafafa' -BackgroundColor '#de2831' -Duration 2000 -MessageSize 40 -Position center}
        

        

        
        }
         
        Else {

              Show-UDToast -Message "Username and number required or are invalid!" -MessageColor '#fafafa' -BackgroundColor '#de2831' -Duration 800 -MessageSize 40 -Position center
              
              }

        }

      } -BackgroundColor 'green'
    
    # Delete a phone / device
    New-UDCollapsibleItem -Icon trash -Title "Remove Device" -Content {
    
    New-UDInput -Title "Remove DUO Device" -FontColor '#379629' -Endpoint {
    
    param($DUORemoveDeviceNumber)



    if ($DUORemoveDeviceNumber -ne ""){
                                      
                                      $DUOGetDeviceIDforREM = duoGetPhone -number $($DUORemoveDeviceNumber)
                                      $DUODeviceIDforREM = $DUOGetDeviceIDforREM.phone_id

                                      
                                      
                                      try {
                                           $DUORemoveDevice = duoDeletePhone -dOrg 'prod' -phone_id $($DUODeviceIDforREM)
                                           Show-UDToast -Message "Device $($DUODeviceIDforREM) has been removed from DUO!" -MessageColor '#f8faf7' -BackgroundColor '#379629' -Duration 2000 -MessageSize 40 -Position center
                                           }
                                           catch{
                                                 Show-UDToast -Message "Device does not exist!" -MessageColor '#f8faf7' -BackgroundColor '#de2831' -Duration 2000 -MessageSize 40 -Position center
                                                 } 

                                      }
                                      Else {
                                           Show-UDToast -Message "Please type in a number!" -MessageColor '#f8faf7' -BackgroundColor '#de2831' -Duration 2000 -MessageSize 40 -Position center
                                           } 
    
    
    }
    
    } -BackgroundColor 'green'
     
     } -Type Expandable -BackgroundColor 'lightgray'

     New-UDElement -Tag center -Content {
     New-UDImage -Url "https://rubylife.callruby.com/Documents/Ruby_Logo%c2%ae_RGB_Rogers.png" -Height 250 -Width 250
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


# DUO UD display grid to display enrolled users and thier devices with management elements #
#==========================================================================================#
New-UDLayout -Columns 1 -Content {

# UDGrid Headers and Property Variables
#======================================
$GridHeaders = @(
                 "Full Name", 
                 "Username", 
                 "Email", 
                 "Status", 
                 "Device", 
                 "OS", 
                 "PhoneID",
                 "Push MFA",
                 "Send SMS Activation"
                 "QR Activation Code"
                 )

$GridHeaderProperties = @(
                          "FullName", 
                          "Username", 
                          "Email", 
                          "Status", 
                          "Device", 
                          "Platform", 
                          "DeviceID"
                          "Push",
                          "SMSActivate"
                          "QRActivate"
                          )

# UDGRid for duo users list
#======================================
New-UDGrid -Id "DuoUDGrid" -Title "Duo User List" -Headers @($GridHeaders) -Properties @($GridHeaderProperties) -Endpoint{

$DuoUserlist = duoGetUser

$DuoUserlist | ForEach-Object{
          
            [PSCustomObject]@{
                              FullName = $_.realname
                              Username = $_.username
                              Email = $_.email
                              Status = $_.status
                              Device = $_.phones.model
                              Platform = $_.phones.platform
                              DeviceID = $_.phones.phone_id
                              Push = New-UDButton -Text "Push" -BackgroundColor "green" -OnClick(New-UDEndpoint -Endpoint{
                              
                              $DuoSearch = $ArgumentList[0]
                              
                              Show-UDModal -Content {
                              
                              New-UDInput -Title "Push Options" -SubmitText "Send Push"-Content {
                         
                              New-UDInputField -Name "DuoPushTitle" -Placeholder "Title" -Type "textbox" -DefaultValue "Title of MFA prompt"
                              New-UDInputField -Name "DuoPushInfo" -Placeholder "DetailInfo" -Type "textarea" -DefaultValue "MFA Info"
                              } -Endpoint {
                         
                              $DuoUser = duoGetUser -username $DuoSearch
                              $DuoUserName = $DuoUser.username
                              $DuoDeviceID = $DuoUser.phones.phone_id

                              
                              If ($DuoPushInfo -ne "") {
                              
                                 $DuoSendAuth = duoSendAuth -dOrg 'auth' -username $DuoUserName -device $DuoDeviceID -PromptTitle $DuoPushTitle -PushInfo "Message=$($DuoPushInfo)"
                              
                              } Else {Show-UDToast -Message "Message info cannot be null!" -Duration 10000 -Position center}
                         
                              
                              If($DuoSendAuth.result -eq 'allow') {
                              
                                Show-UDToast -Message "User has accepted MFA!" -Position center -MessageColor '#fafafa' -Duration 10000 -BackgroundColor '#379629'
                              
                              } Else {Show-UDToast -Message "User has failed / denied or timed out mfa!" -Position center -MessageColor '#fafafa' -Duration 10000 -BackgroundColor '#de2831'}
                            
                           Sync-UDElement -Id "DuoUDGrid" -Broadcast
                          
                        }         
                                       
                     }

                 } -ArgumentList $_.username)
                
                # Sends an activation code via SMS
                SMSActivate = New-UDButton -Text "Send" -BackgroundColor "lightgray" -FontColor "black" -OnClick (New-UDEndpoint -Endpoint {
                                                                                                                                            $DUOSendActivation = duoSendActivationCodes -dOrg 'prod' -phone_id $ArgumentList[0]
                                                                                                                                            Show-UDToast -Message "Activation code sent to $($ArgumentList[1])" -Duration 900 -MessageSize 30 -BackgroundColor "green" -Position center
                                                                                                                                            } -ArgumentList $_.phones.phone_id, $_.username)
                
                # Generates a QR Activation Code for associated phone ID in Gridview
                QRActivate = New-UDButton -Text "Generate" -BackgroundColor "lightgray" -FontColor "black" -OnClick (New-UDEndpoint -Endpoint {
                                                                                                                                               $DUOGenerateQR = duoCreateActivationCode -dOrg 'prod' -phone_id $ArgumentList[0] -activation_barcode
                                                                                                                                               $QRCode = $DUOGenerateQR.activation_barcode
                                                                                                                                               
                                                                                                                                               # Display QR Code in a Modal Pop
                                                                                                                                               Show-UDModal -Content {

                                                                                                                                               New-UDElement -Tag font -Attributes @{
                                                                                                                                                                                    face = "Verdana"
                                                                                                                                                                                    color = "#112457"
                                                                                                                                                                                    } -Content {
                                                                                                                                                                                                $linebreak = New-UDElement -Tag br

                                                                                                                                                                                                "QR Code for $($_.username)"
                                                                                                                                                                                                $linebreak
                                                                                                                                                                                                "This Barcode will be active for 60 Minutes"
                                                                                                                                                                                                $linebreak
                                                                                                                                                                                    
                                                                                                                                                                                               }

                                                                                                                                               New-UDElement -Tag center -Content {
                                                                                                                                               
                                                                                                                                               New-UDElement -Tag iframe -Attributes @{
                                                                                                                                                                                       src = "$($QRCode)"
                                                                                                                                                                                       height = 250
                                                                                                                                                                                       width = 250
                                                                                                                                                                                      }
                                                                                                                                                                                  }

                                                                                                                                               
                                                                                                                                               } -BackgroundColor "lightgray" -FontColor "Green"

                                                                                                                                               } -ArgumentList $_.phones.phone_id, $_.username)

             }

        }  | Out-UDGridData

    }

}




} -Icon lock_open -Title "DUO Multifactor Authentication Support"


# Application Management Page
#============================================#
$pageApplications = New-UDPage -Name 'Applications' -Content { 
    
    # Card with image carousel
    New-UDCard -Size large -Content {
        
        # Bottom Card Header
        New-UDHeading -Content {

        New-UDButton -Text "Ruby.com" -BackgroundColor "Teal" -FontColor "white" -OnClick {
        
        Show-UDToast -Position center -Duration 5000 -Message "Opening Ruby.com"
        start https://www.ruby.com

        }

        New-UDButton -Text "Salesforce.com" -BackgroundColor "Teal" -FontColor "white" -OnClick {
        
        Show-UDToast -Position center -Duration 5000 -Message "Opening Salesforce.com"
        start https://login.salesforce.com
        
        }

        New-UDButton -Text "IT Sharepoint" -BackgroundColor "Teal" -FontColor "white" -OnClick {
        
        Show-UDToast -Position center -Duration 5000 -Message "Opening IT Sharepoint"
        start https://callruby.sharepoint.com/sites/IT2
        
        }

        } -Color "lightgray"


        $CarouselItmes = @(
                           'https://w.wallhaven.cc/full/j8/wallhaven-j83xyw.jpg',
                           'https://w.wallhaven.cc/full/43/wallhaven-43qg9y.jpg',
                           'https://w.wallhaven.cc/full/83/wallhaven-837l2k.jpg',
                           'https://w.wallhaven.cc/full/r2/wallhaven-r2eg1w.png',
                           'https://w.wallhaven.cc/full/vg/wallhaven-vg2y5p.jpg',
                           'https://w.wallhaven.cc/full/96/wallhaven-969w3d.jpg'
                           )
        
        $IDCitem = "CarouselItem"

        $IDCitemNumber = 0

        New-UDImageCarousel -AutoCycle -Items {
        
        foreach ($imageitem in $CarouselItmes) {
        
        $IDCitemNumber++

        New-UDImageCarouselItem -Id "$($IDCitem)$($IDCitemNumber)" -BackgroundImage $imageitem -BackgroundSize 'cover' -BackgroundColor "#3f51b5" -Url $imageitem -BackgroundPosition 'center'
        
        } 

        
        } -FullWidth -FixButton -ButtonText 'Navigate to..' -ShowIndicators -Speed 10000
        
        } 

    #Header
    New-UDHeading -Text "Install Applications"

    #Page Layout for application deployments
    New-UDLayout -Columns 3 -Content {
        
        
        #Application Cards


        # Reinstall Assist tracker Card
        New-UDCollapsible -Items {
            New-UDCollapsibleItem -Title "Assist Tracker" -Icon arrow_circle_right -Content{
            New-UDCard -Title "Assist Tracker" -Size small -Content {
        

        # Assist Tracker
        New-UDInput -Endpoint {

        param($hostname)

        $source = '\\prlfs01\ISWizards\SuperScript\assist tracker\*.*'
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
        elseif ($testpth2 -eq $true) {
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


    Show-UDToast -Position center -Message "Assist tracker installed on $($hostname)" -Duration 5000


}

        } -BackgroundColor "Teal" -Horizontal
          }
        }

        # Reinstall Ruby OS Card
        New-UDCollapsible -Items {
            New-UDCollapsibleItem -Title "Ruby OS" -Icon arrow_circle_right -Content {
            New-UDCard -Title "Install Ruby OS" -Size small -Content {

        # Ruby OS
        New-UDInput -Endpoint {

        param($hostname, $station)

        if ($hostname -ne 'null' -and $station -ne 'null'){


        $ErrorActionPreference = 'silentlycontinue'
        $source = "c:\SuperscriptProd\Applications\RubyOS\*.*"
        $destination = ('\\' + $hostname + '\C$\RubyOS')

        (gwmi Win32_Process -ComputerName $hostname | ?{ $_.ProcessName -match "RubyOS"}).Terminate()

        rd -Path ('\\' + $hostname + '\C$\RubyOS') -Verbose -Force
        md -Path ('\\' + $hostname + '\C$\RubyOS')

        Start-BitsTransfer -Source $source -Destination $destination -Description "Copy Status.." -DisplayName "Copying..."

        (Get-Content \\$($hostname)\c$\RubyOS\settings.xml) -replace '%station%', $station | Set-Content \\$($hostname)\c$\RubyOS\settings.xml -Verbose

        clear

        Show-UDToast -message "RubyOS Updated on system $($hostname)" -MessageSize 30 -Position topCenter -Duration 3000

        }
        Else {Show-UDToast -Message "Please provide the hostname and station number!" -MessageSize 30 -Position topCenter -Duration 3000 }
}
        
        } -BackgroundColor "Teal" -Horizontal
          }
        }

        # Reinstall 7-Zip PDQ
        New-UDCollapsible -Items {
            
            New-UDCollapsibleItem -Title "7-Zip" -Icon arrow_circle_right -Content {
            New-UDCard -Title "Install 7-Zip" -Size small -Content {

            New-UDLayout -Columns 2 -Content {

            New-UDInput -Endpoint {
            
            
            param ([string]$hostname)

            $scriptblockcontent = {param ($hostname)  
                                   Set-Location "C:\Program Files (x86)\Admin Arsenal\PDQ Deploy\";  
                                   PDQDeploy.exe Deploy -Package "7-Zip 19.00" -Target $hostname}

            Invoke-Command -ComputerName ksc01 -ScriptBlock $scriptblockcontent -ArgumentList $hostname
            
            Show-UDToast -Message "7-Zip is being installed on $($hostname)! Check PDQ for status Deployment" -Position center -Duration 9000
            }

            New-UDCard -Text "This will push the 7-Zip Application with PDQ to a desired host."
            }
        } -BackgroundColor "Teal" -Horizontal

            }
        }

        # Reinstall Shortkeys
        New-UDCollapsible -Items {
            
            New-UDCollapsibleItem -Title "Shortkeys" -Icon arrow_circle_right -Content {
            New-UDCard -Title "Install Shortkeys" -Size small -Content {

            New-UDLayout -Columns 2 -Content {

            New-UDInput -Endpoint {
            
            
            param ([string]$hostname)

            $scriptblockcontent = {param ($hostname)  
                                   Set-Location "C:\Program Files (x86)\Admin Arsenal\PDQ Deploy\";  
                                   PDQDeploy.exe Deploy -Package "Shortkeys" -Target $hostname}

            Invoke-Command -ComputerName ksc01 -ScriptBlock $scriptblockcontent -ArgumentList $hostname
            
            Show-UDToast -Message "Shortkeys is being installed on $($hostname)! Check PDQ for status Deployment" -Position center -Duration 9000
            }

            New-UDCard -Text "This will push the Shortkeys Application with PDQ to a desired host."
            }
        } -BackgroundColor "Teal" -Horizontal

            }
        }

        # Reinstall Microsoft Teams
        New-UDCollapsible -Items {
            
            New-UDCollapsibleItem -Title "Microsoft Teams" -Icon arrow_circle_right -Content {
            New-UDCard -Title "Install Teams" -Size small -Content {

            New-UDLayout -Columns 2 -Content {

            New-UDInput -Endpoint {
            
            
            param ([string]$hostname)

            $scriptblockcontent = {param ($hostname)  
                                   Set-Location "C:\Program Files (x86)\Admin Arsenal\PDQ Deploy\";  
                                   PDQDeploy.exe Deploy -Package "Microsoft Teams x64 (With Message saying to restart NO AUTO RESTART!)" -Target $hostname}

            Invoke-Command -ComputerName ksc01 -ScriptBlock $scriptblockcontent -ArgumentList $hostname
            
            Show-UDToast -Message "Teams is being installed on $($hostname)! Check PDQ for status Deployment" -Position center -Duration 9000
            }

            New-UDCard -Text "This will push the Microsoft Teams Application with PDQ to a desired host."
            }
        } -BackgroundColor "Teal" -Horizontal

            }
        }

        # Reinstall O365 Pro Plus
        New-UDCollapsible -Items {
            
            New-UDCollapsibleItem -Title "O365 ProPlus" -Icon arrow_circle_right -Content {
            New-UDCard -Title "Install O365 ProPlus" -Size small -Content {

            New-UDLayout -Columns 2 -Content {

            New-UDInput -Endpoint {
            
            
            param ([string]$hostname)

            $scriptblockcontent = {param ($hostname)  
                                   Set-Location "C:\Program Files (x86)\Admin Arsenal\PDQ Deploy\";  
                                   PDQDeploy.exe Deploy -Package "O365 - ProPlus - Silent" -Target $hostname}

            Invoke-Command -ComputerName ksc01 -ScriptBlock $scriptblockcontent -ArgumentList $hostname
            
            
            }

            New-UDCard -Text "This will push the O365 ProPlus Application with PDQ to a desired host."
            }
        } -BackgroundColor "Teal" -Horizontal

            }
        }


    }
    

        


} -Icon folder -Title "Applications"


# Systems Management Page
#============================================#
$pageSystems = New-UDPage -Name 'Systems Management' -Content {

    New-UDLayout -Columns 2 -Content {
       

      



       #left Card Installed Applications
       New-UDCard -Content {
       
       New-UDGrid -Title "Installed Applications" -Headers @("Application Name", "Version") -Properties @("name", "version") -Endpoint {
       Get-WmiObject -Class win32_product | select name, version | Where-Object -Property name -NotLike "*Microsoft*" | Out-UDGridData
       }

       
       }

       #Right Card Running Processes
       New-UDCard -Content {
       
       
       New-UdGrid -Title "Running Processes" -Headers @("Name", "ID", "Working Set", "CPU") -Properties @("Name", "Id", "WorkingSet", "CPU") -AutoRefresh -RefreshInterval 60 -Endpoint {
       Get-Process | Select Name,ID,WorkingSet,CPU | Out-UDGridData
       }
       
       
       
       }
    
    
    }


} -Icon bolt -Title "Systems Management"

# Onboarding Page
#============================================#
$pageOnboarding = New-UDPage -Name 'Onboarding' -Content { 
 




} -Icon people_carry -Title "Onboarding"


# Departing Page
#============================================#
$pageDeparting = New-UDPage -Name 'Departing' -Content { 
 




} -Icon bomb -Title "Departing"


# Wizards Knowledgebase Page
#============================================#
$pageKB = New-UDPage -Name 'Wizards Knowledgebase' -Content {



} -Icon book -Title "Wizard Wiki"


# Dashboard Footer
#============================================#
$WizardDashFooter = New-UDFooter -Links @(
    New-UDLink -FontColor Pink -Text "Ruby.com | " -Url "https://Ruby.com" -OpenInNewWindow
    New-UDLink -FontColor Red -Text " Office 365 | " -Url "https://www.office.com/?auth=2&home=1" -OpenInNewWindow
    New-UDLink -FontColor Red -Text " IT Sharepoint | " -Url "https://callruby.sharepoint.com/sites/IT2" -OpenInNewWindow
) -BackgroundColor Gray


# Dashboard
#============================================#
$DashVar = New-UDDashboard -Title "TechX Dashboard" -NavBarColor Black -BackgroundColor white -Pages @(
                                                                                                       $pageMain,
                                                                                                       $pageDuoSupport,
                                                                                                       $pageApplications,
                                                                                                       $pageSystems,
                                                                                                       $pageOnboarding,
                                                                                                       $pageDeparting,
                                                                                                       $pageKB
                                                                                                       ) -Footer $WizardDashFooter


# Start Dashboard
#================
Start-UDDashboard -Dashboard $DashVar -Port 10001 -AutoReload -AdminMode

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

# Function Call
#==============
WizardDash

clear
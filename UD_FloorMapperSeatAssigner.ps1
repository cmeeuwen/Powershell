function SeatAssignerDash {

$Endpoint1 = C:\Scripts\Universal_Dashboard\Floormapper_Dash\SeatAssignerAuthScripts\Endpoint1.ps1

$loginForm = C:\Scripts\Universal_Dashboard\Floormapper_Dash\SeatAssignerAuthScripts\LoginBasedOnSecGroup.ps1


# Login Dashboard Authentication Page
#====================================
$loginPage = New-UDLoginPage -AuthenticationMethod $loginForm -LoginFormBackgroundColor "white" -Title "Please Login" -PageBackgroundColor "lightgray" -Logo(
                                                                                                                                                             New-UDElement -Tag Center -Content {

                                                                                                                                                             $ImageURL = "https://rubylife.callruby.com/Documents/Ruby_Logo%c2%ae_RGB_Rogers.png"

                                                                                                                                                             New-UDImage -Url "$($ImageURL)"  -Height 150 -Width 200
                                                                                                                                                             }
                                                                                                                                                            ) -WelcomeText "Floor Mapper Seat Assigner Login"


# FloorMapper Seat Assigner
#============================================#
$pageSeatAssigner = New-UDPage -Name 'Floor Mapper Seat Assignment' -Content {

New-UDLayout -Columns 1 -Content {

$OfficeGridHeaders = @(
                       "Full Name",
                       "Email Address",
                       "Job Title",
                       "Department",
                       "Office Location",
                       "Edit Office"
                       )

$OfficeGridProperties = @(
                          "DisplayName",
                          "Email",
                          "JobTitle",
                          "Department",
                          "Office",
                          "Update"
                          )

New-UDGrid -Id "UserOfficeGrid" -Title "Users and Office Locations" -Headers @($OfficeGridHeaders) -Properties @($OfficeGridProperties) -Endpoint {

$GetADUsersAndOffice = Get-ADUser -Filter * -Properties * -SearchBase 'OU=Ruby Users,DC=ruby,DC=local' | 
                       Where {$_.Enabled -EQ $true -and $_.department -ne $null -and $_.Employeeid -ne $null} |
                       Select -Property Office, SamAccountName, Title, UserPrincipalName, Department, DisplayName


                       $GetADUsersAndOffice | ForEach-Object {

                                            [PSCustomObject]@{
                                                              DisplayName = $_.DisplayName
                                                              Email = $_.UserPrincipalName
                                                              JobTitle = $_.Title
                                                              Department = $_.Department
                                                              Office = $_.Office
                                                              Update = New-UDButton -Text "Edit Office" -BackgroundColor "Teal" -FontColor "Black" -OnClick (
                                                              
                                                                New-UDEndpoint -Endpoint {
                                                                
                                                                $UserListSearch = $ArgumentList[0]
                                                                
                                                                Show-UDModal -Content {
                                                                                       
                                                                                       New-UDInput -Title "Change Office / Location for $($_.DisplayName)" -FontColor '#379629' -Endpoint {
                                                                                       
                                                                                       param($Office)

                                                                                       if ($Office) {
                                                                                             
                                                                                             

                                                                                                 $UpdateOffice = Get-ADUser $_.SamAccountName | Set-ADUser -Office "$($Office)"
                                                                                                 Show-UDToast -Message "Office / Location updated for $($_.DisplayName)!" -MessageColor '#f8faf7' -BackgroundColor '#379629' -Duration 2000 -MessageSize 40 -Position center
                                                                                                 
                                                                                                 
                                                                                             
                                                                                             }

                                                                                             Else {
                                                                                             
                                                                                                   Show-UDToast -Message "Office Cannot be null!" -MessageColor '#f8faf7' -BackgroundColor '#de2831' -Duration 2000 -MessageSize 40 -Position center
                                                                                             
                                                                                             }
                                                                                       
                                                                                       } -BackgroundColor "LightGray"
                                                                                       
                                                                                       New-UDCard -Content {
                                                                                              
                                                                                           New-UDElement -Tag font -Attributes @{
                                                                                                                                color = "Teal"
                                                                                                                                size = 3.5
                                                                                                                                face = "Verdana"
                                                                                                                                } -Content {
                                                                                                                                           
                                                                                                                                New-UDElement -Tag b -Content{

                                                                                                                                         $currentUserInfo = @(
                                                                                                                                                              "Name: $($_.DisplayName)",
                                                                                                                                                              "Email: $($_.UserPrincipalName)",
                                                                                                                                                              "Title: $($_.Title)",
                                                                                                                                                              "Department: $($_.Department)",
                                                                                                                                                              "Current Office: $($_.Office)"
                                                                                                                                                              )
                                                                                                                                                              foreach ($UserInfItem in $currentUserInfo) {
                                                                                                                                                                                                          $UserInfItem
                                                                                                                                                                                                          New-UDElement -Tag br
                                                                                                                                                                                                          }
                                                                                                                                           
                                                                                                                                           New-UDElement -Tag br
                                                                                                                                           New-UDElement -Tag p -Content  {"Note:"}
                                                                                                                                           

                                                                                                                                           }

                                                                                                                                           $UpdateNotice = @(
                                                                                                                                                            "When updating the office / location for a user, Users from other sites might not see this change reflect on this site for up to 1 hour."    
                                                                                                                                                            )
                                                                                                                                                            foreach ($updateitem in $UpdateNotice) {
                                                                                                                                                                                                   $updateitem
                                                                                                                                                                                                   }
                                                                                                                                           
                                                                                                          
                                                                                                                                           }
                                                                                       
                                                                                       
                                                                                       
                                                                                      } -Size small                                                            
                                                                                      
                                                                                      }
                                                                
                                                                
                                                                } -ArgumentList $_.SamAccountName)

                                                             }
 
    } | Out-UDGridData
    
  }

}
                          


} -Icon chair -Title "Floor Mapper Seat Assigner"


# Dashboard Footer
#============================================#
$WizardDashFooter = New-UDFooter -Links @(
    New-UDLink -FontColor Pink -Text "Ruby.com | " -Url "https://Ruby.com" -OpenInNewWindow
    New-UDLink -FontColor Red -Text " Office 365 | " -Url "https://www.office.com/?auth=2&home=1" -OpenInNewWindow
    New-UDLink -FontColor Red -Text " IT Support | " -Url "https://help.ruby.com/" -OpenInNewWindow
) -BackgroundColor Gray


# Dashboard
#============================================#
$SeatAssignerDash = New-UDDashboard -Title "FloorMapper Seat Assignment" -NavBarColor Black -BackgroundColor white -Pages @(
                                                                                                                            $pageSeatAssigner
                                                                                                                            ) -Footer $WizardDashFooter -EndpointInitialization $init -LoginPage $loginPage


# Start Dashboard
#================
Start-UDDashboard -Dashboard $SeatAssignerDash -Port 10003 -AutoReload -Force -AllowHttpForLogin -Endpoint @($Endpoint1)

}

SeatAssignerDash

read-host "stop"
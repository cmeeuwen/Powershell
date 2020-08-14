function Show-tabcontrol_psf {

	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.ServiceProcess, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$form1 = New-Object 'System.Windows.Forms.Form'
	$tabcontrol1 = New-Object 'System.Windows.Forms.TabControl'
	$tabpage1 = New-Object 'System.Windows.Forms.TabPage'
	$tabpage2 = New-Object 'System.Windows.Forms.TabPage'
	$tabpage3 = New-Object 'System.Windows.Forms.TabPage'
	$tabpage4 = New-Object 'System.Windows.Forms.TabPage'
	$tabpage5 = New-Object 'System.Windows.Forms.TabPage'
    $tabpage6 = New-Object 'System.Windows.Forms.TabPage'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------
	
	$form1_Load={
		
	}
	
	# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$form1.WindowState = $InitialFormWindowState
	}
	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$form1.remove_Load($form1_Load)
			$form1.remove_Load($Form_StateCorrection_Load)
			$form1.remove_FormClosed($Form_Cleanup_FormClosed)
            
		}
		catch { Out-Null  }
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	$form1.SuspendLayout()
	$tabcontrol1.SuspendLayout()
	
    #
	# form1
	#======================================================================================
	$form1.Controls.Add($tabcontrol1)
	$form1.FormBorderStyle = 1
	$form1.AutoScale = $True
    $form1.ClientSize = '650, 700'
	$form1.Name = 'form1'
	$form1.Text = 'Superscript GUI V2.0'
	$form1.add_Load($form1_Load)
    $form1.Opacity = 0.95
    $form1icon = New-Object System.Drawing.Icon ("$PSScriptRoot\Source\ros_app_icon.ico") 
    $form1.Icon = $form1icon
    $form1.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    


	#
	# tabcontrol1 Main Top Tabs
	#======================================================================================
	$tabcontrol1.Controls.Add($tabpage1)
	$tabcontrol1.Controls.Add($tabpage2)
	$tabcontrol1.Controls.Add($tabpage3)
	$tabcontrol1.Controls.Add($tabpage4)
	$tabcontrol1.Controls.Add($tabpage6)
    $tabcontrol1.Controls.Add($tabpage5)
	$tabcontrol1.Alignment = 'Top'
	$tabcontrol1.Location = '12, 12'
	$tabcontrol1.Multiline = $True
	$tabcontrol1.Name = 'tabcontrol1'
	$tabcontrol1.SelectedIndex = 0
	$tabcontrol1.Size = '629, 678'
	$tabcontrol1.TabIndex = 0
    $tabcontrol1.AutoSize = $True
    
	
    #================================================================#
	#                    tab page1 User Management                   #
	#================================================================#
	$tabpage1.Location = '42, 4'
	$tabpage1.Name = 'tabpage1'
	$tabpage1.Padding = '3, 3, 3, 3'
	$tabpage1.Size = '583, 442'
	$tabpage1.TabIndex = 0
	$tabpage1.Text = 'User Management'
	$tabpage1.UseVisualStyleBackColor = $True

    $LabelFonts = New-Object System.Drawing.Font("Verdana",8,[System.Drawing.FontStyle]::Bold)      

    #================================================================#
    #                      Sub tab 1 Onboarding                      #
    #================================================================#
    $subtabpageOnboard = New-Object 'System.Windows.Forms.TabPage' -Property @{
    Location = '0, 15'
    Padding = '3, 3, 3, 3'
    Size = '602, 442'
    TabIndex = 10
    Text = ' Onboarding '
    UseVisualStyleBackColor = $True
    }

    # sub tab 1 Onboarding Main Label - Done in SP
    #================================================================
    $tabpage1LabelEnableUsers = New-Object System.Windows.Forms.Label -Property @{
    Text = "Enable Rubys"
    Location = '15,20'
    Size = '120,40'
    Font = $LabelFonts
    ForeColor = "Teal"
    BackColor = "White"}
    
    # User Input Components - Done in SP
    # ===============================================================
    # Onboarding Username Label
    $onboardUsernameLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Username:"
    Size = '60,30'
    Location = '25,103'
    }

    # Onboarding Username Textbox - Done in SP
    #================================================================
    $onboardUsernameTB = New-Object System.Windows.Forms.TextBox -Property @{
    BackColor = "LightGray"
    Size = '100,30'
    Location = '85,100'
    }

    # Onboarding Team Name Label - Done in SP
    #================================================================
    $onboardingTeamNameLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Team Name:"
    Size = '69,15'
    Location = '16,145'
    }
    
    # Onboarding Team Name Textbox - Done in SP
    #================================================================
    $onboardTeamTB = New-Object System.Windows.Forms.TextBox -Property @{
    BackColor = "LightGray"
    Size = '100,30'
    Location = '85,143'
    }
    
    # Onboarding Manager Name Label - Done in SP
    #================================================================
    $onboardingManagerLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Manager:"
    Size = '52,30'
    Location = '32,186'
    }
    
    # Onboarding Manager Name Textbox - Done in SP
    #================================================================
    $onboardManagerTB = New-Object System.Windows.Forms.TextBox -Property @{
    BackColor = "LightGray"
    Size = '100,30'
    Location = '85,185'
    }
    
    # Onboarding Station Number Label - Done in SP
    #================================================================
    $onboardingstationLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Station:"
    Size = '43,30'
    Location = '40,232'
    }
    
    # Onboarding Station Number Textbox - Done in SP
    #================================================================
    $onboardStationTB = New-Object System.Windows.Forms.TextBox -Property @{
    BackColor = "LightGray"
    Size = '100,30'
    Location = '85,230'
    }

    # Onboarding Sub Tab Buttons
    # ===============================================================
    
    # New Ruby Main Label - Done in SP
    #================================================================#
    $NewRubyMainLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "New Ruby"
    ForeColor = "Teal"
    AutoSize = $True
    Font = $LabelFonts
    Location = '240,65'
    }

    # Next Adventure Main Label - Done in SP
    #================================================================#
    $NextAdventureMainLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Next Adventure"
    ForeColor = "Teal"
    Font = $LabelFonts
    AutoSize = $True
    Location = '240,210'
    }

    # Next Adventure Sales Button - Done in SP
    #================================================================#
    $NADVSales = New-Object System.Windows.Forms.Button -Property @{
    Text = ' Next Adventure Sales '
    AutoSize = $True
    ForeColor = "Teal"
    Location = '240,246'
    }
    $NADVSales.add_click({
    
    if ($onboardUsernameTB.Text -ne "" -and $onboardTeamTB.Text -ne "" -and $onboardManagerTB.Text -ne "" -and $onboardStationTB.Text -ne "") {
    
    $NextAdvSalesUser = $onboardUsernameTB.Text
    $NextAdvSalesPrevTeam = $onboardTeamTB.Text
    $NextSalesManager = $onboardManagerTB.Text
    $NextSalesOfficeLoc = $onboardStationTB.Text

    clear

    Get-ADUser $NextAdvSalesUser | Move-ADObject -TargetPath 'OU=PDX Sales,OU=Ruby Users,DC=ruby,DC=local' -Verbose
    Get-ADUser $NextAdvSalesUser | Set-ADUser -Manager $NextSalesManager -Verbose
    Get-ADUser $NextAdvSalesUser | Set-ADUser -Office $NextSalesOfficeLoc -Verbose
    Get-ADUser $NextAdvSalesUser | Set-ADUser -PasswordNeverExpires $false -Verbose

               $groupRemovalsNAS = @(
                                     "$($NextAdvSalesPrevTeam)",
                                     "ALL_Receptionists_Prl_Sec",
                                     "ALL_Receptionists_Bvt_Sec",
                                     "PearlTeam",
                                     "beavertonteam",
                                     "Exclaimer_group1"
                                    )
                                    foreach ($remGroup in $groupRemovalsNAS) {
                                          
                                        Try { Remove-ADGroupMember -Identity $remGroup -Members $NextAdvSalesUser -confirm:$false -Verbose }
                                            
                                             Catch {Write-Host "$($NextAdvSalesUser) is not a member of the group $($remGroup)" -ForegroundColor Yellow}
                                         
                                    }
                                            
                   $groupAddsNAS = @(
                                     "FoxTower",
                                     "FS_TC_Audio_Agreements_RW",
                                     "GrowingLeaders",
                                     ('ISR' + ' ' + 'Team-1739626290'),
                                     "Logi_NonReceptionist",
                                     "NonReceptionists",
                                     "pearl_sql_users",
                                     "PublicFolderAccess",
                                     "SALES",
                                     ('Sales' + ' ' + 'and' + ' ' + 'Onboarding-1-488285678'),
                                     "Sales_Sales_SEC",
                                     "SP_SalesAssociates"
                                    )
                                    foreach ($addGroup in $groupAddsNAS) {
                                    
                                        Try {Add-ADGroupMember -Identity $addGroup -Members $NextAdvSalesUser -Verbose}
                                    
                                            Catch {Write-Host " Unable to add $($NextAdvSalesUser) to the following group: $($addGroup)" -ForegroundColor Yellow}
                                    
                                    
                                    }
            

clear
            
Write-Host "
|=========================================|
|                                         |
|  $($NextAdvSalesUser) Moved to sales!   |
|                                         |
|=========================================|
" -ForegroundColor Green
    
    } 
    
    Else {
          
          $invalidInputNotifSales = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
          $invalidInputNotifSales.Popup("Username, Team Name, Manager Username, and station required!",0,"Invalid Input - Next Adventure Sales") 
         
         }
    
})

    # Next Adventure CH Button - Done in SP
    #================================================================#
    $NADVCH = New-Object System.Windows.Forms.Button -Property @{
    Text = ' Next Adventure CH '
    AutoSize = $True
    ForeColor = "Teal"
    Location = '372,246'
    }
    $NADVCH.add_click({

    if ($onboardUsernameTB.Text -ne "" -and $onboardStationTB.Text -ne "") {
    
Write-Host "Onboard new CH
==============================" -BackgroundColor Blue -ForegroundColor Green
Write-Host ""

$ErrorActionPreference = 'silentlycontinue'

# Setting Variables to create user and add them to the proper
# OU and Groups
$username = $onboardUsernameTB.Text
$officeloc = $onboardStationTB.Text
$department = 'PSHM Hourly'
$manager = $onboardManagerTB.Text
$company = 'Ruby Receptionists'
$fstlst = ($firstname + ' ' + $lastname)
$firstinidotlastname = (($FirstName.Substring(0,1)) + ($LastName))
$hpmainname = ('HP' + ' ' + $firstname + ' ' + $lastname)
$hpAdGroup = ('HP' + $RecepTeam)
$hpUser = ('HP' + $firstname + '.' + $lastname)
$hpsur = ('HP ' + $lastname)
$hpgiv = $lastname
$hpOU = 'OU=Hurdles & Praise,OU=Non-Human Users,OU=Ruby Users - System\, Non-Human\, External,DC=ruby,DC=local'
$hpprincname = ('HP' + $firstname + '.' + $lastname + '@ruby.com')


# Modifys Multiple attributes on newly created user
# Based on the input at the beginning of the script
# Then moves user to the proper OU
Get-ADUser $username | Set-ADUser -Office $officeloc -Verbose
Get-ADUser $username | Set-ADUser -ScriptPath 'ch.bat' -Verbose
Get-ADUser $username | Set-ADUser -Title $jobtitle -Verbose
Get-ADUser $username | Set-ADUser -Manager $manager -Verbose
Get-ADUser $username | Set-ADUser -Department $department -Verbose
Get-ADUser $username | Set-ADUser -PasswordNeverExpires $false -Verbose
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
Add-ADGroupMember -Identity TeamRuby $username -confirm:$false 
Add-ADGroupMember -Identity SP_ProblemSolversHappinessMakers $username -confirm:$false 
Add-ADGroupMember -Identity calendlee $username -confirm:$false
Add-ADGroupMember -Identity PSO_Standard_Ruby_User_Password $username -Confirm:$false
 

# Remove user from associated AD groups
Remove-ADGroupMember -Identity Exclaimer_Group1 $username -confirm:$false 
Remove-ADGroupMember -Identity SPARK_BtownReceptionists $username -confirm:$false 
Remove-ADGroupMember -Identity SPARK_Receptionists $username -confirm:$false 
Remove-ADGroupMember -Identity SPARK_PearlReceptionists $username -confirm:$false 
Remove-ADGroupMember -Identity Receptionists $username -confirm:$false 
Remove-ADGroupMember -Identity SPARK_PearlReceptionists $username -confirm:$false 

clear

Write-Host "CH Next Adventure changed in AD and Exhange!" -BackgroundColor Blue -ForegroundColor Green
    

}

Else {
     
      $invalidInputNACH = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
      $invalidInputNACH.Popup("Username and station required!",0,"Invalid Input - Next Adventure CH")           

     }

})
    
    # Sub Tab Onboarding PRL Recep Button - Done in SP
    #================================================================#
    $EnablePRLButton = New-Object System.Windows.Forms.Button -Property @{
    AutoSize = $True
    Text = " PRL Receptionist "
    ForeColor = "Teal"
    Location = '240,100'}
    $EnablePRLButton.add_click({

       if ($onboardUsernameTB.Text -ne "" -and $onboardTeamTB.Text -ne "" -and $onboardManagerTB.Text -ne "" -and $onboardStationTB.Text -ne "") {
    
            $NewRubyUser = $onboardUsernameTB.Text
            $NewRubyTeam = $onboardTeamTB.Text
            $NewRubyCulti = $onboardManagerTB.Text
            $NewRubyOfficeLoc = $onboardStationTB.Text

            Get-ADUser $NewRubyUser | Enable-ADAccount -Verbose

            Get-ADUser $NewRubyUser | Set-ADUser -Office $NewRubyOfficeLoc -Verbose
            Get-ADUser $NewRubyUser | Set-ADUser -EmailAddress ($NewRubyUser + '@ruby.com').ToLower();
            Get-ADUser $NewRubyUser | Set-ADUser -Manager $NewRubyCulti -Verbose
            Get-ADUser $NewRubyUser | Set-ADUser -Title "$NewRubyTeam Receptionist" -Verbose
            Get-ADUser $NewRubyUser | Set-ADUser -PasswordNeverExpires $false -Verbose
            Get-ADUser $NewRubyUser | Set-ADUser -CannotChangePassword $false -Verbose

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
            Add-ADGroupMember -Identity PSO_Standard_Ruby_User_Password $NewRubyUser -Verbose

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


     Enable-RemoteMailbox -Identity $NewRubyUser -PrimarySmtpAddress ($NewRubyUser + '@ruby.com')  -RemoteRoutingAddress ($NewRubyUser + '@callruby.mail.onmicrosoft.com') -Verbose

     Get-RemoteMailbox -Identity $NewRubyUser

Write-Host "$NewRubyUser Enabled!
===========================================================================================
| Run the following command in Exchange Management Shell to verify remotemailbox user     |
| Command: Get-RemoteMailbox -Identity %Username%                                         |
===========================================================================================" -ForegroundColor Yellow

 } 
 
 Else {
 
      $invalidInputNotifPRL = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
      $invalidInputNotifPRL.Popup("Username, Team Name, Manager Username, and station required!",0,"Invalid Input - Pearl Receptionist") 
 
      }

})

    # Sub Tab Onboarding BVT Recep Button - Done in SP
    #================================================================#
    $EnableBVTButton = New-Object System.Windows.Forms.Button -Property @{
    AutoSize = $True
    Text = " BVT Receptionist "
    ForeColor = "Teal"
    Location = '350,100'}
    $EnableBVTButton.add_click({
    
       if ($onboardUsernameTB.Text -ne "" -and $onboardTeamTB.Text -ne "" -and $onboardManagerTB.Text -ne "" -and $onboardStationTB.Text -ne "") {
     
            $NewRubyUser = $onboardUsernameTB.Text
            $NewRubyTeam = $onboardTeamTB.Text
            $NewRubyCulti = $onboardManagerTB.Text
            $NewRubyOfficeLoc = $onboardStationTB.Text

            Get-ADUser $NewRubyUser | Enable-ADAccount -Verbose

            Get-ADUser $NewRubyUser | Set-ADUser -Office $NewRubyOfficeLoc -Verbose
            Get-ADUser $NewRubyUser | Set-ADUser -EmailAddress ($NewRubyUser + '@ruby.com').ToLower();
            Get-ADUser $NewRubyUser | Set-ADUser -Manager $NewRubyCulti -Verbose
            Get-ADUser $NewRubyUser | Set-ADUser -Title "$NewRubyTeam Receptionist" -Verbose
            Get-ADUser $NewRubyUser | Set-ADUser -PasswordNeverExpires $false -Verbose
            Get-ADUser $NewRubyUser | Set-ADUser -CannotChangePassword $false -Verbose

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
            Add-ADGroupMember -Identity PSO_Standard_Ruby_User_Password $NewRubyUser -Verbose

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


Enable-RemoteMailbox -Identity $NewRubyUser -PrimarySmtpAddress ($NewRubyUser + '@ruby.com')  -RemoteRoutingAddress ($NewRubyUser + '@callruby.mail.onmicrosoft.com') -Verbose

Get-RemoteMailbox -Identity $NewRubyUser


Write-Host "$NewRubyUser Enabled!
===========================================================================================
| Run the following command in Exchange Management Shell to verify remotemailbox user     |
| Command: Get-RemoteMailbox -Identity %Username%                                         |
===========================================================================================" -ForegroundColor Yellow
    
    
 } 
 
 Else {
 
       $invalidInputNotifBVT = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
       $invalidInputNotifBVT.Popup("Username, Team Name, Manager Username, and station required!",0,"Invalid Input - Beaverton Receptionist") 
 
      }

})

    # Sub Tab Onboarding KC Recep Button - Done in SP
    #================================================================#
    $EnableKCButton = New-Object System.Windows.Forms.Button -Property @{
    AutoSize = $True
    Text = " KC Receptionist "
    ForeColor = "Teal"
    Location = '460,100'}
    $EnableKCButton.add_click({
    
        if ($onboardUsernameTB.Text -ne "") {

            $NewRubyUser = $onboardUsernameTB.Text
                       

            Get-ADUser $NewRubyUser | Enable-ADAccount -Verbose

            Get-ADUser $NewRubyUser | Set-ADUser -EmailAddress ($NewRubyUser + '@ruby.com').ToLower();
            Get-ADUser $NewRubyUser | Set-ADUser -PasswordNeverExpires $false -Verbose
            Get-ADUser $NewRubyUser | Set-ADUser -CannotChangePassword $false -Verbose

            Get-ADUser $NewRubyUser | Move-ADObject -TargetPath 'OU=ProChats,OU=Ruby Users,DC=ruby,DC=local' -Verbose
            Add-ADGroupMember -Identity ('KC' + ' ' + 'Team-1280071548') $NewRubyUser -Verbose
            Add-ADGroupMember -Identity TeamRuby $NewRubyUser -Verbose
            Add-ADGroupMember -Identity SEC_Sophos_Standard_WebFiltering $NewRubyUser -Verbose
            Add-ADGroupMember -Identity Logi_NonReceptionist -Members $NewRubyUser -Verbose
            Add-ADGroupMember -Identity APP_SSO_ProChatsOnline_Access -Members $NewRubyUser -Verbose
            Add-ADGroupMember -Identity PSO_Standard_Ruby_User_Password $NewRubyUser -Verbose

            $isNonRecep = Get-ADUser -Identity $NewRubyUser | Select -Property Title


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


Enable-RemoteMailbox -Identity $NewRubyUser -PrimarySmtpAddress ($NewRubyUser + '@ruby.com')  -RemoteRoutingAddress ($NewRubyUser + '@callruby.mail.onmicrosoft.com') -Verbose

Get-RemoteMailbox -Identity $NewRubyUser


Write-Host "$NewRubyUser Enabled!
===========================================================================================
| Run the following command in Exchange Management Shell to verify remotemailbox user     |
| Command: Get-RemoteMailbox -Identity %Username%                                         |
===========================================================================================" -ForegroundColor Yellow
    
    
 } 
 
 Else {
 
       $invalidInputNotifKC = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
       $invalidInputNotifKC.Popup("Username is required!",0,"Invalid Input - KC Receptionist") 
 
      }    
    
})

    # Sub Tab Onboarding Sales Button - Done in SP
    #================================================================#
    $EnableSalesButton = New-Object System.Windows.Forms.Button -Property @{
    AutoSize = $True
    Text = " Sales "
    ForeColor = "Teal"
    Location = '240,140'}
    $EnableSalesButton.add_click({
    
         if ($onboardUsernameTB.Text -ne "" -and $onboardManagerTB.Text -ne "" -and $onboardStationTB.Text -ne "" ) {
    
                $NewRubyUser = $onboardUsernameTB.Text                

                $EnableSalesNewManager = $onboardManagerTB.Text

                $EnableSalesStation = $onboardStationTB.Text

                Get-ADUser $NewRubyUser | Enable-ADAccount -Verbose
                Get-ADUser $NewRubyUser | Move-ADObject -TargetPath 'OU=PDX Sales,OU=Ruby Users,DC=ruby,DC=local' -Verbose
                Get-ADUser $NewRubyUser | Set-ADUser -Manager $EnableSalesNewManager -Verbose
                Get-ADUser $NewRubyUser | Set-ADUser -Office $EnableSalesStation -Verbose
                Get-ADUser $NewRubyUser | Set-ADUser -CannotChangePassword $false -Verbose
                Get-ADUser $NewRubyUser | Set-ADUser -PasswordNeverExpires $false -Verbose

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
                Add-ADGroupMember -Identity PSO_Standard_Ruby_User_Password $NewRubyUser -Verbose

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


Enable-RemoteMailbox -Identity $NewRubyUser -PrimarySmtpAddress ($NewRubyUser + '@ruby.com')  -RemoteRoutingAddress ($NewRubyUser + '@callruby.mail.onmicrosoft.com') -Verbose

Get-RemoteMailbox -Identity $NewRubyUser


Write-Host "$NewRubyUser Enabled!
===========================================================================================
| Run the following command in Exchange Management Shell to verify remotemailbox user     |
| Command: Get-RemoteMailbox -Identity %Username%                                         |
===========================================================================================" -ForegroundColor Yellow
    
    
  } 
  
  Else {
  
        $invalidInputNotifSalesNew = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
        $invalidInputNotifSalesNew.Popup("Username, Team Name, Manager Username, and station required!",0,"Invalid Input - Sales") 
       
       }  

})


    # Sub Tab Onboarding Generic Admin Button - Done in SP
    #================================================================#
    $EnableGAButton = New-Object System.Windows.Forms.Button -Property @{
    AutoSize = $True
    Text = " Generic Admin "
    ForeColor = "Teal"
    Location = '318,140'}
    $EnableGAButton.add_click({
    
       if ($onboardUsernameTB.Text -ne "" -and $onboardManagerTB.Text -ne "" -and $onboardStationTB.Text -ne "") {

                $NewRubyUser = $onboardUsernameTB.Text
                $EnableAdminNewManager = $onboardManagerTB.Text
                $EnableAdminStation = $onboardStationTB.Text

                Get-ADUser $NewRubyUser | Enable-ADAccount -Verbose
                Get-ADUser $NewRubyUser | Set-ADUser -Office $EnableAdminStation -Verbose
                Get-ADUser $NewRubyUser | Set-ADUser -Manager $EnableAdminNewManager -Verbose
                Get-ADUser $NewRubyUser | Set-ADUser -PasswordNeverExpires $false -Verbose
                Get-ADUser $NewRubyUser | Set-ADUser -CannotChangePassword $false -Verbose
                
                Add-ADGroupMember -Identity SEC_Sophos_Standard_WebFiltering $NewRubyUser -Verbose
                Add-ADGroupMember -Identity PSO_Standard_Ruby_User_Password $NewRubyUser -Verbose

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

                     if ($result -eq [System.Windows.Forms.DialogResult]::OK) {Write-Host "AD Groups successfully queried!" -ForegroundColor Green} else {Write-Host "Cancelled, No groups added to user!" -ForegroundColor Red}
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

                if ($result -eq [System.Windows.Forms.DialogResult]::OK) {Write-Host "AD OU's Queried Successfully!" -ForegroundColor Green} else {Write-Host "Cancelled! User not moved to any OU!" -ForegroundColor Red}

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
     Enable-RemoteMailbox -Identity $NewRubyUser -PrimarySmtpAddress ($NewRubyUser + '@ruby.com')  -RemoteRoutingAddress ($NewRubyUser + '@callruby.mail.onmicrosoft.com') -Verbose

     Get-RemoteMailbox -Identity $NewRubyUser
    }
    catch {Write-Host "$NewRubyUser remote mailbox already enabled!" -ForegroundColor Yellow}

Write-Host "$NewRubyUser Enabled!
===========================================================================================
| Run the following command in Exchange Management Shell to verify remotemailbox user     |
| Command: Get-RemoteMailbox -Identity %Username%                                         |
===========================================================================================" -ForegroundColor Yellow
    
    
  } 
  
  Else {
  
        $invalidInputNotifAdmin = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
        $invalidInputNotifAdmin.Popup("Username, Team Name, Manager Username, and station required!",0,"Invalid Input - Generic Admin") 
  
       }    
    
})


    # Verify AD User Button - Done in SP - Done in SP
    #================================================================#
    $EnableVerifyButton = New-Object System.Windows.Forms.Button -Property @{
    AutoSize = $True
    Text = " Verify User "
    ForeColor = "Teal"
    Location = '30,270'}
    $EnableVerifyButton.add_click({
        
        $ErrorActionPreference = 'SilentlyContinue'

        $VerifyADUser = Get-ADUser $onboardUsernameTB.Text

           If ($VerifyADUser -ne $null) {
                                         $UserDoesExist = New-Object -ComObject Wscript.Shell -ErrorAction Stop
                                         $UserDoesExist.Popup("$($onboardUsernameTB.Text) Does exist in AD!",0,"Verification Success!") 
                                         }
                                         Else { 
                                               $UserDoesntExist = New-Object -ComObject Wscript.Shell -ErrorAction Stop
                                               $UserDoesntExist.Popup("$($onboardUsernameTB.Text) does not exist in AD!",0,"Verification Failure!")
                                               }
                                   })


    # Sub Tab Onboarding Clear Textboxes Button - Done in SP
    #================================================================#
    $EnableClearButton = New-Object System.Windows.Forms.Button -Property @{
    AutoSize = $True
    Text = " Clear fields "
    ForeColor = "Teal"
    Location = '108,270'}
    $EnableClearButton.add_click({
                                  $onboardStationTB.Clear()
                                  $onboardManagerTB.Clear()
                                  $onboardTeamTB.Clear()
                                  $onboardUsernameTB.Clear()
                                  })
    
    #==========================================================#
    #
    #  Onboard Via Template Section
    #
    #==========================================================#

    
    # Onboarding Template Generator Function
    Function CreateUserTemplate {

# .Net Assembly Load
#==========================================================================#
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Main Form
#==========================================================================#
$MainForm = New-Object System.Windows.Forms.Form -Property @{
size = '1110,630'
BackColor = "lightgray"
Opacity = .975
Text = "Onboarding template generator"
StartPosition = 'CenterScreen'
}
$mainIcon = New-Object System.Drawing.Icon ("$PSScriptRoot\Source\wizard.ico")
$MainForm.Icon = $mainIcon

$LabelFont = New-Object System.Drawing.Font("Verdana",8,[System.Drawing.FontStyle]::Bold)


# Template Name Label
#==========================================================================#
$templateNameLabel = New-Object System.Windows.Forms.Label -Property @{
Location = '12,128'
Text = "Template Name:"
AutoSize = $True
Font = $LabelFont 
}

# Template Name Textbox
#==========================================================================#
$templateNameTextbox = New-Object System.Windows.Forms.TextBox -Property @{
Location = '10,150'
Size = '241,20'
BackColor = "white"
}


# Job Title Label
#==========================================================================#
$jobTitleLabel = New-Object System.Windows.Forms.Label -Property @{
Location = '12,177'
Text = "Job Title:"
AutoSize = $True
Font = $LabelFont 
}


# Job Title Textbox
#==========================================================================#
$jobTitle = New-Object System.Windows.Forms.TextBox -Property @{
BackColor = "white"
Location = '10,197'
Size = '241,20'
}


# Department Label
#==========================================================================#
$DepartmentLabel = New-Object System.Windows.Forms.Label -Property @{
Location = '12,230'
Text = "Department:"
AutoSize = $True
Font = $LabelFont
}


# Department Textbox
#==========================================================================#
$DepartmentTextBox = New-Object System.Windows.Forms.TextBox -Property @{
BackColor = "white"
Location = '10,250'
Size = '241,20'
}


# OU Dropdown Label
#==========================================================================#
$SelectOULabel = New-Object System.Windows.Forms.Label -Property @{
Location = '12,283'
Text = "Select Orginizational Unit:"
AutoSize = $True
Font = $LabelFont 
}


# Form OU Dropdown Selection
#==========================================================================# 
$OUSelectionDropdown = New-Object System.Windows.Forms.ComboBox -Property @{
AllowDrop = $True
Size = '240,50'   
Location = '10,303'
}

$OUSelectionDropdown.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend;
$OUSelectionDropdown.AutoCompleteCustomSource.AddRange(('OU=Billing,OU=Ruby Users,DC=ruby,DC=local',
           'OU=BVN Admin,OU=Ruby Users,DC=ruby,DC=local',
           'OU=BVN HR,OU=Ruby Users,DC=ruby,DC=local',
           'OU=BVN Information Technology,OU=Ruby Users,DC=ruby,DC=local',
           'OU=BVN Office,OU=Ruby Users,DC=ruby,DC=local',
           'OU=BVN Receptionists,OU=Ruby Users,DC=ruby,DC=local',
           'OU=BVN RS,OU=Ruby Users,DC=ruby,DC=local',
           'OU=BVN Sales,OU=Ruby Users,DC=ruby,DC=local',
           'OU=FXT Scheduling Ninjas,OU=Ruby Users,DC=ruby,DC=local',
           'OU=Contractors,OU=Ruby Users,DC=ruby,DC=local',
           'OU=Marketing,OU=Ruby Users,DC=ruby,DC=local',
           'OU=Misc Hourly Rubys,OU=Ruby Users,DC=ruby,DC=local',
           'OU=PDX Administrative,OU=Ruby Users,DC=ruby,DC=local',
           'OU=PDX Client Happiness,OU=Ruby Users,DC=ruby,DC=local',
           'OU=PDX Employee Engagement,OU=Ruby Users,DC=ruby,DC=local',
           'OU=PDX HR,OU=Ruby Users,DC=ruby,DC=local',
           'OU=PDX Information Technology,OU=Ruby Users,DC=ruby,DC=local',
           'OU=PDX Receptionists,OU=Ruby Users,DC=ruby,DC=local',
           'OU=PDX RS,OU=Ruby Users,DC=ruby,DC=local',
           'OU=PDX Sales,OU=Ruby Users,DC=ruby,DC=local',
           'OU=ProChats,OU=Ruby Users,DC=ruby,DC=local',
           'OU=Product,OU=Ruby Users,DC=ruby,DC=local',
           'OU=RX,OU=Ruby Users,DC=ruby,DC=local',
           'OU=Talent,OU=Ruby Users,DC=ruby,DC=local'));
$OUSelectionDropdown.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::CustomSource;

$AdOUs = @(
           'OU=Billing,OU=Ruby Users,DC=ruby,DC=local',
           'OU=BVN Admin,OU=Ruby Users,DC=ruby,DC=local',
           'OU=BVN HR,OU=Ruby Users,DC=ruby,DC=local',
           'OU=BVN Information Technology,OU=Ruby Users,DC=ruby,DC=local',
           'OU=BVN Office,OU=Ruby Users,DC=ruby,DC=local',
           'OU=BVN Receptionists,OU=Ruby Users,DC=ruby,DC=local',
           'OU=BVN RS,OU=Ruby Users,DC=ruby,DC=local',
           'OU=BVN Sales,OU=Ruby Users,DC=ruby,DC=local',
           'OU=FXT Scheduling Ninjas,OU=Ruby Users,DC=ruby,DC=local',
           'OU=Contractors,OU=Ruby Users,DC=ruby,DC=local',
           'OU=Marketing,OU=Ruby Users,DC=ruby,DC=local',
           'OU=Misc Hourly Rubys,OU=Ruby Users,DC=ruby,DC=local',
           'OU=PDX Administrative,OU=Ruby Users,DC=ruby,DC=local',
           'OU=PDX Client Happiness,OU=Ruby Users,DC=ruby,DC=local',
           'OU=PDX Employee Engagement,OU=Ruby Users,DC=ruby,DC=local',
           'OU=PDX HR,OU=Ruby Users,DC=ruby,DC=local',
           'OU=PDX Information Technology,OU=Ruby Users,DC=ruby,DC=local',
           'OU=PDX Receptionists,OU=Ruby Users,DC=ruby,DC=local',
           'OU=PDX RS,OU=Ruby Users,DC=ruby,DC=local',
           'OU=PDX Sales,OU=Ruby Users,DC=ruby,DC=local',
           'OU=ProChats,OU=Ruby Users,DC=ruby,DC=local',
           'OU=Product,OU=Ruby Users,DC=ruby,DC=local',
           'OU=RX,OU=Ruby Users,DC=ruby,DC=local',
           'OU=Talent,OU=Ruby Users,DC=ruby,DC=local'
           )
           foreach ($ou in $AdOUs){$OUSelectionDropdown.Items.Add($ou)}


# Group Selection Label
#==========================================================================#
$SelectGroupsLabel = New-Object System.Windows.Forms.Label -Property @{
Location = '12,330'
Text = "Select Groups:"
AutoSize = $True
Font = $LabelFont 
}

# AD Group Listbox
#==========================================================================#
$SelectGroupsListbox = New-Object System.Windows.Forms.Listbox -Property @{
Location = '10,350'
Size =  '240,230'
}
$SelectGroupsListbox.SelectionMode = 'MultiSimple'

$getADgroups = Get-ADGroup -SearchBase 'OU=Ruby Groups,DC=ruby,DC=local' -Filter 'ObjectClass -eq "group"' |
               select Name, SamAccountName | Sort-Object -Property Name, SamAccountName

               foreach ($SecGroup in $getADgroups){[void] $SelectGroupsListbox.Items.Add($SecGroup.Name)}
                

# Form Template display grid preview
#==========================================================================# 
$TemplatePreviewGrid = New-Object System.Windows.Forms.DataGridView -Property @{
ColumnCount = 4
ColumnHeadersVisible = $True
Size = '780,359'
Location = '290,151'
ScrollBars = "Vertical"
ReadOnly = $True
BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3d
BackgroundColor = "Teal"
}
$TemplatePreviewGrid.Columns[0].Name = "Title"
$TemplatePreviewGrid.Columns[0].Width = 160
$TemplatePreviewGrid.Columns[1].Name = "OU"
$TemplatePreviewGrid.Columns[1].Width = 260
$TemplatePreviewGrid.Columns[2].Name = "Department"
$TemplatePreviewGrid.Columns[2].Width = 160
$TemplatePreviewGrid.Columns[3].Name = "Groups"
$TemplatePreviewGrid.Columns[3].Width = 160


# Form button for template creation
#==========================================================================# 
$GenerateTemplateButton = New-Object System.Windows.Forms.Button -Property @{
Text = "Generate Template"
Location = '900,530'
AutoSize = $true
BackColor = 'white'
}
$GenerateTemplateButton.add_click({

$doesTemplateExist = Test-Path -Path "$($PSScriptRoot)\OnboardTemplates\$($templateNameTextbox.Text).csv"

if ($doesTemplateExist -eq $false) {
 
$convertGridtoTable = New-Object System.Data.DataTable -Property @{
TableName = "ConvertTable"
}

        foreach ($columns in $TemplatePreviewGrid.Columns) {$convertGridtoTable.Columns.Add($columns.Name)}

        foreach ($rows in $SelectGroupsListbox.SelectedItems) {$convertGridtoTable.Rows.Add($null,$null,$null,$rows)}

        $convertGridtoTable.Rows.Add($jobTitle.Text,$OUSelectionDropdown.SelectedItem,$DepartmentTextBox.Text,$null)

        $TemplatePreviewGrid.DataSource = $convertGridtoTable

        $convertGridtoTable | Export-Csv $PSScriptRoot\OnboardTemplates\scratch.csv -Verbose

        $Removefirstrow = get-content $PSScriptRoot\OnboardTemplates\scratch.csv
        $Removefirstrow | Select-Object -Skip 1 | Set-Content "$($PSScriptRoot)\OnboardTemplates\$($templateNameTextbox.Text).csv" -Verbose

        Remove-Item $PSScriptRoot\OnboardTemplates\scratch.csv -Force -Verbose
        
        $MainForm.Hide()

        $templateGenerateMessage = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
        $templateGenerateMessage.Popup("New Onboarding Template Generated!",0,"Success!") 
         
        $MainForm.Close() }
        
        Else {
              $templateFailGenerateMessage = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
              $templateFailGenerateMessage.Popup("$($templateNameTextbox.Text) already exists!`nPlease choose a new template name!",0,"Template Generate failure!") 
              
              Remove-Item $PSScriptRoot\OnboardTemplates\scratch.csv -Force -Verbose
             }        
})


# Form button for template review
#==========================================================================# 
$ReviewTemplateButton = New-Object System.Windows.Forms.Button -Property @{
Text = "Review Template"
Location = '790,530'
AutoSize = $true
BackColor = 'white'
}
$ReviewTemplateButton.add_click({
                                 
                                 $TemplatePreviewGrid.rows.Clear()

                                 foreach ($selectedgroup in $SelectGroupsListbox.SelectedItems)
                                         {$TemplatePreviewGrid.Rows.Add($null,$null,$null,$selectedgroup)}
                                 
                                 $TemplatePreviewGrid.Rows[0].Cells[0].Value = $jobTitle.Text
                                 $TemplatePreviewGrid.Rows[0].Cells[1].Value = $OUSelectionDropdown.SelectedItem
                                 $TemplatePreviewGrid.Rows[0].Cells[2].Value = $DepartmentTextBox.Text
                               
                                 })


# Form backgroundimage
#==========================================================================# 
$formBGimage = [System.Drawing.Image]::FromFile("$PSScriptRoot\source\Rubycom.png")
$MainFormBGimag = New-Object System.Windows.Forms.PictureBox -Property @{
Image = $formBGimage
AutoSize = $true
Location = '10,10'
}


# Form components array
#==========================================================================#
$formCompArray = @(
                   $MainFormBGimag,
                   $templateNameLabel,
                   $templateNameTextbox,
                   $jobTitleLabel,
                   $jobTitle,
                   $DepartmentLabel,
                   $DepartmentTextBox,
                   $SelectOULabel,
                   $OUSelectionDropdown,
                   $SelectGroupsLabel,
                   $SelectGroupsListbox,
                   $TemplatePreviewGrid,
                   $ReviewTemplateButton,
                   $GenerateTemplateButton
                   )
                   foreach ($ArrayComp in $formCompArray){$MainForm.controls.add($ArrayComp)}



$MainForm.ShowDialog()

}


    # Onboarding Template Section Label - Done in SP
    #================================================================#
    $onboardingTemplateMainLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Enable Ruby via generated template"
    Font = $LabelFonts
    Location = '10,345'
    AutoSize = $True
    ForeColor = "teal"
    }


    # Onboard via Template Username Textbox Label - Done in SP
    #================================================================#
    $OBTUserTextboxLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Username:"
    Location = '250,375'
    AutoSize = $True
    }


    # Onboard via Template Username Textbox - Done in SP
    #================================================================#
    $OBTUserTextbox = New-Object System.Windows.Forms.TextBox -Property @{
    BackColor = "Lightgray"
    size = '120,20'
    Location = '317,373'
    }


    # Onboard via Template Manager Username Label - Done in SP
    #================================================================#
    $OBTManagerTextboxLabel = New-Object System.Windows.Forms.Label -Property @{
    AutoSize = $True
    Text = "Manager:"
    Location = '258,410'
    }


    # Onboard via Template Manager Username Textbox - Done in SP
    #================================================================#
    $OBTManagerTextbox = New-Object System.Windows.Forms.TextBox -Property @{
    BackColor = "LightGray"
    Size = '120,20'
    Location = '317,410'
    }


    # Onboard via Template Recep Team Label - Done in SP
    #================================================================#
    $OBTTeamLabel = New-Object System.Windows.Forms.Label -Property @{
    AutoSize = $True
    Text = "Recep Team:"
    Location = '238,449'
    }


    # Onboard via Template Recep Team Textbox - Done in SP
    #================================================================#
    $OBTTeamTextbox = New-Object System.Windows.Forms.TextBox -Property @{
    BackColor = "LightGray"
    size = '120,20'
    Location = '317,447'
    }


    # Next Adventure Checkbox - Done in SP
    #================================================================#
    $OBTNACheckbox = New-Object System.Windows.Forms.CheckBox -Property @{
    Location = '450,373'
    Text = "Next Adventure?"
    AutoSize = $True
    Font = $LabelFonts
    ForeColor = 'Teal'
    }

    # Return to roots Checkbox - Done in SP
    #================================================================#
    $OBTRTRCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{
    Location = '450,410'
    Text = "Return to Roots?"
    AutoSize = $True
    Font = $LabelFonts
    ForeColor = 'Teal'
    }


    # Refresh template list button - Done in SP
    #================================================================#
    $RefreshTemplateListButton = New-Object System.Windows.Forms.Button -Property @{
    AutoSize = $True
    Text = "Refresh template list"
    ForeColor = "teal"
    Location = '457,510'
    }
    $RefreshTemplateListButton.add_click({
    
    $templateListBox.Items.clear()
    
    $RePopulateTemplateListbox = gci "$($PSScriptRoot)\OnboardTemplates\*.csv"
                                 foreach ($template in $RePopulateTemplateListbox) 
                                      {[void] $templateListBox.Items.Add($template.name)}
    
    
    })

  
    # Template listbox selection - Done in SP
    #================================================================#
    $templateListBox = New-Object System.Windows.Forms.ListBox -Property @{
    Size = '200,200'
    Location = '22,370'
    }
    $PopulateTemplateListbox = gci "$($PSScriptRoot)\OnboardTemplates\*.csv"
                               foreach ($template in $PopulateTemplateListbox) 
                                       {[void] $templateListBox.Items.Add($template.name)}
    
    # Create new template button - Done in SP
    #================================================================#
    $CreateTemplateButton = New-Object System.Windows.Forms.Button -Property @{
    AutoSize = $True
    Text = "Create template"
    ForeColor = "teal"
    Location = '250,510'
    }
    $CreateTemplateButton.add_click({
                                     CreateUserTemplate
                                     [System.Windows.Forms.Application]::DoEvents()
                                    })


    # Remove template button - Done in SP
    #================================================================#
    $RemoveTemplateButton = New-Object System.Windows.Forms.Button -Property @{
    AutoSize = $True
    Text = "Remove template"
    ForeColor = "teal"
    Location = '350,510'
    }
    $RemoveTemplateButton.add_click({
                                     
                                      if ($templateListBox.SelectedItem -ne $null) {
                                     
                                     Remove-Item -Path "$($PSScriptRoot)\OnboardTemplates\$($templateListBox.SelectedItem)" -Verbose -Force


                                     [void] $templateListBox.Items.Remove($templateListBox.SelectedItem)
                                     
                                     clear-host

                                     Write-Host "Template Removed!" -ForegroundColor Green
                                     
                                     } 
                                     
                                     Else {
                                           
                                           $OBTRemoveinvalidInput = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
                                           $OBTRemoveinvalidInput.Popup("No Template Selected to remove!",0,"Remove Error: Missing Input")
                                    
                                          }
                                    
                                     })


    # Review Selected Template Button - Done in SP
    #================================================================#
    $reviewSelectedTemplateButton = New-Object System.Windows.Forms.Button -Property @{
    Text = "Review Template"
    ForeColor = "Teal"
    AutoSize = $True
    Location = '250,545'
    }
    $reviewSelectedTemplateButton.add_click({
                                             if ($templateListBox.SelectedItem -ne $null)
                                                
                                                {
                                                 start "$($PSScriptRoot)\onboardtemplates\$($templateListBox.SelectedItem)"
                                                }
                                                 Else {
                                                       $OBTReviewinvalidInput = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
                                                       $OBTReviewinvalidInput.Popup("No Template Selected to review!",0,"Review Error: Missing Input")
                                                      }
                                             
                                             })


    # Enable Ruby via template button - Done in SP
    #================================================================#
    $EnableRubyViaTemplateButton = New-Object System.Windows.Forms.Button -Property @{
    AutoSize = $True
    Text = "Enable Ruby"
    ForeColor = "Teal"
    Location = '357,545'
    }
    $EnableRubyViaTemplateButton.add_click({
            
            if ($OBTUserTextbox.Text -ne "" -and $OBTManagerTextbox.Text -ne "" -and $OBTDeptTextbox.Text -ne ""){

            if ($OBTNACheckbox.Checked -eq $True -or $OBTRTRCheckbox.Checked -eq $True){
            
            $getGroupMembership = Get-ADUser -Identity $OBTUserTextbox.Text -Properties Memberof | select -Expand memberof
            foreach ($groupmembership in $getGroupMembership) {
                                                               Write-Host "Removing $($OBTUserTextbox.Text) from $($groupmembership)" -ForegroundColor White
                                                               Remove-ADGroupMember $groupmembership -Member $OBTUserTextbox.Text -Confirm:$false -Verbose
                                                               }
            
            }

            Get-ADUser $OBTUserTextbox.Text | Enable-ADAccount -Verbose
            Write-Host "Setting email address to $($OBTUserTextbox.text + '@ruby.com')" -ForegroundColor Green
            Get-ADUser $OBTUserTextbox.Text | Set-ADUser -EmailAddress ($OBTUserTextbox.Text + '@ruby.com').ToLower();
            Get-ADUser $OBTUserTextbox.Text | Set-ADUser -Manager $OBTManagerTextbox.Text -Verbose
            Get-ADUser $OBTUserTextbox.Text | Set-ADUser -Company "Ruby Receptionists" -Verbose
            Get-ADUser $OBTUserTextbox.Text | Set-ADUser -PasswordNeverExpires $false -Verbose
            Get-ADUser $OBTUserTextbox.Text | Set-ADUser -CannotChangePassword $false -Verbose
            
            try  {
            Add-ADGroupMember -Identity $OBTTeamTextbox.Text $OBTUserTextbox.Text -Verbose
            } catch {Write-Host "The Team $($OBTTeamTextbox.Text) does not exist!" -ForegroundColor Red}

            Import-Csv "$($PSScriptRoot)\OnboardTemplates\$($templateListBox.SelectedItem)" | ForEach-Object {
            
            if ($_.OU -ne ''){
            Write-Host "Moving to OU $($_.OU)" -ForegroundColor Green 
            Get-ADUser $OBTUserTextbox.Text | Move-ADObject -TargetPath $_.OU -Verbose
            }
            
            if ($_.Groups -ne ''){
            Write-Host "adding to group $($_.Groups)" -ForegroundColor Green           
            Add-ADGroupMember -Identity $_.Groups $OBTUserTextbox.Text -Verbose
            }

            if ($_.Title -ne '') {
            Write-Host "Setting Job title to $($OBTTeamTextbox.Text) $($_.Title)....." -ForegroundColor Green
            Get-ADUser $OBTUserTextbox.Text | Set-ADUser -Title "$($OBTTeamTextbox.Text) $($_.Title)" -Verbose
            }

            if ($_.Department -ne ''){
            Write-Host "Setting department to $($_.Department)" -ForegroundColor Green
            Get-ADUser $OBTUserTextbox.Text | Set-ADUser -Department $_.Department -Verbose
            }

            }

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


     Enable-RemoteMailbox -Identity $OBTUserTextbox.Text -PrimarySmtpAddress ($OBTUserTextbox.Text + '@ruby.com')  -RemoteRoutingAddress ($OBTUserTextbox.Text + '@callruby.mail.onmicrosoft.com') -Verbose

     Get-RemoteMailbox -Identity $OBTUserTextbox.Text

Write-Host "$($OBTUserTextbox.Text) Enabled!
===========================================================================================
| Run the following command in Exchange Management Shell to verify remotemailbox user     |
| Command: Get-RemoteMailbox -Identity %Username%                                         |
===========================================================================================" -ForegroundColor Yellow
}
Else {
      $OBTinvalidInput = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
      $OBTinvalidInput.Popup("Username, Manager Username, and department required!",0,"Enable Error: Missing Input")
     }


[System.Windows.Forms.Application]::DoEvents()
})


    #================================================================#
    #                      Sub tab 2 Departing                       # - Done in SP
    #================================================================# 
    $subtabpageDepart = New-Object 'System.Windows.Forms.TabPage' -Property @{
    Location = '0, 15'
    Padding = '3, 3, 3, 3'
    Size = '602, 442'
    TabIndex = 10
    Text = ' Departing '
    UseVisualStyleBackColor = $True}

    # Import CSV Label - Done in SP
    #=================================================================#
    $departLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = " Depart Rubys"
    Location = '30,40'
    Font = $LabelFonts
    ForeColor = "Teal"
    }
    
    # Import CSV to depart Users Open File Dialog
    #=================================================================#
    $ImportDepartCSV = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = '\\prlfs01\ISWizards\SuperScript\Departed'
    #[Environment]::GetFolderPath('Desktop') 
    Filter = 'CSV (*.csv)|*.csv|SpreadSheet (*.xlsx)|*.xlsx'
    }
    
    # Import CSV Button - Done in SP
    #=================================================================#
    $importcsvbutton = New-Object System.Windows.Forms.Button -Property @{
    Text = " Import CSV "
    ForeColor = "Teal"
    AutoSize = $True
    Location = '30,85'}
    $importcsvbutton.add_click({
                                $ImportDepartCSV.ShowDialog()
                                $showselectedCSV.Text = ''
                                $showselectedCSV.appendtext($ImportDepartCSV.FileName)
                                })

    # Show Selected CSV Label - Done in SP
    #=================================================================#
    $showselectedCSVLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Selected CSV to Depart Users"
    Location = '280,43'
    Size = '200,30'
    Font = $LabelFonts
    ForeColor = "Teal"
    }

    # Show Selected CSV Textbox - Done in SP
    #=================================================================#
    $showselectedCSV = New-Object System.Windows.Forms.TextBox -Property @{
    BackColor = "LightGray"
    ForeColor = "Green"
    Location = '280,86'
    Size = '320,10'
    ReadOnly = $True }
    
    
    # Depart Ruby's Button - Done in SP
    #==================================================================#
    $departButton = New-Object System.Windows.Forms.Button -Property @{
    Text = " Depart "
    ForeColor = "Teal"
    AutoSize = $True
    Location = '190,85'}
    $departButton.add_click({
    
    if ($ImportDepartCSV.FileName -ne $null){
            
            $error.Clear()
            $ErrorActionPreference = 'silentlycontinue'
            $disabledUsersOU = "OU=Disabled Users,DC=ruby,DC=local"

             Import-Csv $ImportDepartCSV.FileName | ForEach-Object {
	
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
            Else {Write-Host "No CSV selected to import!" -ForegroundColor Red}
    
    Write-Host "Archiving Departed CSV" -ForegroundColor Green
    Copy-Item $ImportDepartCSV.FileName -Destination $PSScriptRoot\Departed\Archive -Verbose
    Remove-Item $ImportDepartCSV.FileName -Verbose -Confirm:$false
    })

    
    # Verify DataGridview Depart - Done in SP
    #==================================================================#
    $verifyDepartGrid = New-Object System.Windows.Forms.DataGridView -Property @{
    ColumnCount = 2
    ColumnHeadersVisible = $True
    Size = '318,150'
    Location = '280,120'
    ScrollBars = "Vertical"
    ReadOnly = $True
    BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3d
    }
    $verifyDepartGrid.Columns[0].Name = "Username"
    $verifyDepartGrid.Columns[1].Name = "Office Account"
    $verifyDepartGrid.Columns[1].Width = 250
    

    # verify imported CSV button - Done in SP
    #==================================================================#
    $verifyDepartedButton = New-Object System.Windows.Forms.Button -Property @{
    Text = " Verify  "
    AutoSize = $True
    ForeColor = "Teal"
    Location = '112,85'}
    $verifyDepartedButton.add_click({

        $verifyDepartGrid.Rows.Clear()

    try {
         Import-Csv $ImportDepartCSV.FileName | foreach { $verifyDepartGrid.Rows.Add($_.Username,$_.Office_Account) }
        }
         catch {
                Write-Host "No CSV Imported to Verify" -ForegroundColor Red
           }
    })

    # Clear imported Depart CSV - Done in SP
    #==================================================================#
    $clearImportedCSVButton = New-Object System.Windows.Forms.Button -Property @{
    Text = ' Clear Import '
    AutoSize = $True
    ForeColor = "Teal"
    Location = '480,36'
    }
    $clearImportedCSVButton.add_click({
    
    Try{
        clear
        $ImportDepartCSV.Reset()
        $ImportDepartCSV.InitialDirectory = '\\prlfs01\ISWizards\SuperScript\Departed'
        $ImportDepartCSV.Filter = 'CSV (*.csv)|*.csv|SpreadSheet (*.xlsx)|*.xlsx'
        $verifyDepartGrid.Rows.Clear()
        $showselectedCSV.Clear()
        Write-Host "Imported CSV Cleared!" -ForegroundColor Green
        }
         Catch {
                Write-Host "No CSV Imported Or Verified to Clear" -ForegroundColor Red
                }
    
    })

    # Set AD Account Expiration Date Function
    #==================================================================#
    function SetAccountExpireDate {
    
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

$CalendarLabel = New-Object System.Windows.Forms.Label -Property @{
Text = "Calendar Date"
Font = "Georgia"
ForeColor = "Teal"
Location = "48,92"
Height = 20
Width = 80
}
$CalendarLabel.
$form.Controls.Add($CalendarLabel)

$calendar = New-Object Windows.Forms.MonthCalendar -Property @{
    Location          = New-Object Drawing.Point 50, 120
    ShowTodayCircle   = $true
    MaxSelectionCount = 1
}
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

$UsernameLabel = New-Object System.Windows.Forms.Label -Property @{
Text = “Username”
Font = "Georgia"
ForeColor = "Teal"
Location = “360, 90”
Height = 20
Width = 70
}
$form.Controls.Add($UsernameLabel)

$UsernameTextBox = New-Object System.Windows.Forms.TextBox -Property @{
Location = “360, 120”
Width = “150”
}
$form.Controls.Add($UsernameTextBox)

$TimePickerLabel = New-Object System.Windows.Forms.Label -Property @{
Text = “Time”
Font = "Georgia"
ForeColor = "Teal"
Location = “360, 160”
Height = 22
Width = 40
}
$form.Controls.Add($TimePickerLabel)

$TimePicker = New-Object System.Windows.Forms.TextBox -Property @{
Location = “360, 195”
Width = “150”
}
$form.Controls.Add($TimePicker)

$TimeZoneSelectLabel = New-Object System.Windows.Forms.Label -Property @{
Text = “Time Zone”
Font = "Georgia"
ForeColor = "Teal"
Location = “360, 230”
Height = 22
Width = 68
}
$form.Controls.Add($TimeZoneSelectLabel)

$TimeZoneSelect = New-Object System.Windows.Forms.ComboBox -Property @{
AllowDrop = $true
Location = "360,260"
Size = "110,50"
}
$TimeZoneSelect.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$TimeZoneSelect.AutoCompleteCustomSource.Add("System.Windows.Forms");
$TimeZoneSelect.AutoCompleteCustomSource.AddRange(("PST", "MST", "CST", "EST"));
$TimeZoneSelect.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend;
$TimeZoneSelect.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::CustomSource;
$TimeZoneSelect.Items.Add("PST")
$TimeZoneSelect.Items.Add("MST")
$TimeZoneSelect.Items.Add("CST")
$TimeZoneSelect.Items.Add("EST")
$form.Controls.Add($TimeZoneSelect)

$result = $form.ShowDialog()

if ($result -eq [Windows.Forms.DialogResult]::OK) {
    $date = $calendar.SelectionStart
    $time = $TimePicker.Text
    $user = $UsernameTextBox.Text
    $timezone = $TimeZoneSelect.SelectedItem
    $playExplosion = New-Object System.Media.SoundPlayer
    $playExplosion.SoundLocation = "$PSScriptRoot\Source\dominating.wav"
    

    try {
    Set-ADAccountExpiration -Identity $user -DateTime "$($date.ToShortDateString()) $time" -Verbose
    $playExplosion.PlaySync()
    Write-Host "`n$user set to expire on Date selected: $($date.ToShortDateString()) at $($time.ToString()) $timezone.
Please Allow 1 hour for change to replicate." -ForegroundColor Green
    
    }
    
    Catch {
           Write-Host "Invalid Date time provided, or bad username" -ForegroundColor Red
           }
    Pause

    }
    if($result -eq [Windows.Forms.DialogResult]::CANCEL){
      Write-Host "`nExpiration date cancelled" -ForegroundColor Yellow
      
      }
    
}

    # Set AD Account Expiration Date Label - Done in SP
    #==================================================================#
    $setExpireLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Set AD Account Expiration Date"
    Location = '30,300'
    Size = '220,30'
    Font = $LabelFonts
    ForeColor = "Teal"
    }

    # Block o365 Login Label - Done in SP
    #==================================================================#
    $BlockMSOLLogin = New-Object System.Windows.Forms.Label -Property @{
    Text = "Block o365 login"
    Location = '275,300'
    Size = '220,30'
    Font = $LabelFonts
    ForeColor = "Teal"
    }
    
    # Block O365 Login Textbox - Done in SP
    #==================================================================#
    $BlockMSOLoginTB = New-Object System.Windows.Forms.TextBox -Property @{
    BackColor = "LightGray"
    Size = '150,30'
    Location = '275,350'
    }

    # Block O365 Login Button - Done in SP - Need to add Event Handler
    #==================================================================#
    $BlockMSOLoginButton = New-Object System.Windows.Forms.Button -Property @{
    Text = 'Block Login'
    ForeColor = "Teal"
    Autosize = $True
    Location = '440,348'
    }
    $BlockMSOLoginButton.add_Click({
    
    

    $IsMSOLModuleThere = Get-Module -Name MSOnline

    if ($IsMSOLModuleThere -ne $false) {
        
        Connect-MsolService -Verbose

        $DoesUserExistBasedOnUPN = get-aduser -Filter * -Properties * -SearchBase 'OU=Ruby Users,DC=ruby,DC=local' |
                                   select UserPrincipalName, SamAccountName |
                                   Where -Property UserPrincipalName -EQ "$($BlockMSOLoginTB.Text)"

                                   if ($DoesUserExistBasedOnUPN -ne $null) {
                                      Get-MsolUser -UserPrincipalName "$($BlockMSOLoginTB.Text)" | Set-MsolUser -BlockCredential $True -Verbose
                                      $BlockMSOLoginSuccess = New-Object -ComObject Wscript.Shell -ErrorAction Stop
                                      $BlockMSOLoginSuccess.Popup("$($BlockMSOLoginTB.Text) o365 Login Blocked!",0,"Block Success!") 
                                      }

                                      Else {
                                           $BlockMSOLoginFail = New-Object -ComObject Wscript.Shell -ErrorAction Stop
                                           $BlockMSOLoginFail.Popup("$($BlockMSOLoginTB.Text) does not exist!",0,"Block Failure!") 
                                           }
            
        
           
           }

           Else {
             
                 Install-Module Msolonline -Verbose -Force

                 Connect-MsolService -Verbose

                 $DoesUserExistBasedOnUPN = get-aduser -Filter * -Properties * -SearchBase 'OU=Ruby Users,DC=ruby,DC=local' |
                                   select UserPrincipalName, SamAccountName |
                                   Where -Property UserPrincipalName -EQ "$($BlockMSOLoginTB.Text)"

                                   if ($DoesUserExistBasedOnUPN -ne $null) {
                                      Get-MsolUser -UserPrincipalName "$($BlockMSOLoginTB.Text)" | Set-MsolUser -BlockCredential $True -Verbose
                                      $BlockMSOLoginSuccess = New-Object -ComObject Wscript.Shell -ErrorAction Stop
                                      $BlockMSOLoginSuccess.Popup("$($BlockMSOLoginTB.Text) o365 Login Blocked!",0,"Block Success!") 
                                      }

                                      Else {
                                           $BlockMSOLoginFail = New-Object -ComObject Wscript.Shell -ErrorAction Stop
                                           $BlockMSOLoginFail.Popup("$($BlockMSOLoginTB.Text) does not exist!",0,"Block Failure!") 
                                           }
                 

                 

             }
    
    
    })

    # Set AD Account Expiration Date Button - Done in SP - Need to Add 
    # Click Event Handler to call function
    #==================================================================#
    $setExpireButton = New-Object System.Windows.Forms.Button -Property @{
    Text = ' Set Account Expiration '
    ForeColor = "Teal"
    AutoSize = $True
    Location = '30,350'}
    $setExpireButton.add_Click({SetAccountExpireDate})

    # Set AD Account Expiration Instructions Textbox - Done in SP - Need to Add AppendText Event Handler
    #==================================================================#
    $expireInstructions = New-Object System.Windows.Forms.TextBox -Property @{
    ReadOnly = $True
    Multiline = $True
    ForeColor = "Black"
    BackColor = "White"
    Size = '550,200'
    Location = '30,400'
    ScrollBars = "Vertical"
    BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3d
    Text = "Set AD user Expiration Date Instructions"}
    $expireInstructions.AppendText("`r`n==================================
    `r`n
    `r`nStep 1: Select A Date from the calendar
    `r`n
    `r`nStep 2: Type in the desired user's Username - (Example: first.last)
    `r`n
    `r`nStep 3: Type in the time that you wish the account to expire.
    `r`n
    `r`n        Note: Clock is a 24 hour clock. (Example: 6pm would be in the format of 18:00:00)
    `r`n
    `r`nStep 4: Select a timezone.")
    
    #===============================================================#
    #                   Sub tab 3 Modify Ruby                       #
    #===============================================================#
    $subtabpageModify = New-Object 'System.Windows.Forms.TabPage' -Property @{
    Location = '0, 15'
    Padding = '3, 3, 3, 3'
    Size = '602, 442'
    TabIndex = 10
    Text = ' Modify Ruby '
    UseVisualStyleBackColor = $True}

    # Modify Ruby Main Label - Done in SP
    #===============================================================#
    $ModifyRubyMainLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Modify Ruby / Additional Responsibilities"
    AutoSize = $True
    ForeColor = "Teal"
    Font = $LabelFonts
    Location = '30,50'
    }

    # Exchange / O365 Label - Done in SP
    #===============================================================#
    $ModifyRubyExchO365Label = New-Object System.Windows.Forms.Label -Property @{
    Text = "Exchange / O365"
    AutoSize = $True
    ForeColor = "Teal"
    Font = $LabelFonts
    Location = '30,175'
    }

    # Extended Hours Access Username Label For Username Textbox - Done in SP
    #===============================================================#
    $ModifyRubyUsernameLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Username:"
    Size = '60,30'
    Location = '30,85'}

    # Extended Hours Access Username Textbox - Done in SP
    #===============================================================#
    $ModifyRubyUsernameTextbox = New-Object System.Windows.Forms.TextBox -Property @{
    BackColor = "LightGray"
    Size = '100,30'
    Location = '93,83'}

    #===============================================================#
    #                   Extended Hours Access                       #
    #===============================================================#
    
    # Extended Hours Access Provision EH Access - Done in SP - Need to add Click event handler
    #===============================================================#
    $ExtHoursProvisionButton = New-Object System.Windows.Forms.Button -Property @{
    text = ' Provision EH Access '
    AutoSize = $True
    ForeColor = "Teal"
    Location = '202,81'}
    $ExtHoursProvisionButton.add_click({
    
    $RubyExtUser = $ModifyRubyUsernameTextbox.Text

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

})

    #==============================================================#
    #                 Add User to Multiple Groups                  #
    #==============================================================#


    # Add user to Multiple Groups Button - Done in SP - Need to add Click event handler
    #===============================================================#
    $addUserToMultGroupsButton = New-Object System.Windows.Forms.Button -Property @{
    Text = ' Add to AD groups '
    AutoSize = $True
    ForeColor = "Teal"
    Location = New-Object System.Drawing.Point(335,81)}
    $addUserToMultGroupsButton.add_click({
    
     $MassAddUser = $ModifyRubyUsernameTextbox.Text

                Add-Type -AssemblyName System.Windows.Forms
                Add-Type -AssemblyName System.Drawing

                
                $MAGform = New-Object System.Windows.Forms.Form 
                $MAGform.Text = 'Security Groups and distro memberships'
                $MAGform.Size = New-Object System.Drawing.Size(400,300)
                $MAGform.StartPosition = 'CenterScreen'
                

                $MAGOKButton = New-Object System.Windows.Forms.Button
                $MAGOKButton.Location = New-Object System.Drawing.Point(120,220)
                $MAGOKButton.Size = New-Object System.Drawing.Size(75,23)
                $MAGOKButton.Text = 'OK'
                $MAGOKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $MAGform.AcceptButton = $MAGOKButton
                $MAGform.AutoSize = $true
                $MAGform.Controls.Add($MAGOKButton)

                $MAGCancelButton = New-Object System.Windows.Forms.Button
                $MAGCancelButton.Location = New-Object System.Drawing.Point(200,220)
                $MAGCancelButton.Size = New-Object System.Drawing.Size(75,23)
                $MAGCancelButton.Text = 'Cancel'
                $MAGCancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $MAGform.CancelButton = $MAGCancelButton
                $MAGform.Controls.Add($MAGCancelButton)

                $MAGlabel = New-Object System.Windows.Forms.Label
                $MAGlabel.Location = New-Object System.Drawing.Point(10,20)
                $MAGlabel.Size = New-Object System.Drawing.Size(280,20)
                $MAGlabel.Text = 'Please select groups to add user too:'
                $MAGform.Controls.Add($MAGlabel)

                $MAGlistBox = New-Object System.Windows.Forms.Listbox
                $MAGlistBox.Location = New-Object System.Drawing.Point(15,60)
                $MAGlistBox.Size = New-Object System.Drawing.Size(350,500)
                

                $MAGlistBox.SelectionMode = 'MultiSimple'


                $getgroups = Get-ADGroup -SearchBase 'OU=Ruby Groups,DC=ruby,DC=local' -Filter 'ObjectClass -eq "group"' | select Name | Sort-Object -Property Name
                

                foreach ($group in $getgroups) {
                                                [void] $MAGlistBox.Items.Add($group.name)
                                                }
                

                $MAGlistBox.Height = 150
                $MAGform.Controls.Add($MAGlistBox)
                $MAGform.Topmost = $true

                $result = $MAGform.ShowDialog()

                     if ($result -eq [System.Windows.Forms.DialogResult]::OK) {"True"} else {"Something went wrong!"}
                     ForEach-Object {
    
                                        foreach ($group in $MAGlistBox.SelectedItems)
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

    
    })

    
    #==============================================================#
    #                    Assign O365 Licenses                      #
    #==============================================================#

    # Assign 0365 Licenses Username Label - Done in SP
    #===============================================================#
    $assigno365LicenseUserLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "UPN:"
    Size = '35,30'
    Location = New-Object System.Drawing.Point(30,210)}

    # Assign 0365 Licenses Textbox - Done in SP
    #===============================================================#
    $assigno365LicenseTextBox = New-Object System.Windows.Forms.TextBox -Property @{
    BackColor = "LightGray"
    Size = '160,30'
    Location = New-Object System.Drawing.Point(65,207)
    }

    # Assign 0365 Licenses Button - Done in SP - Need to add Click Event Handler
    #===============================================================#
    $assigno365LicenseButton = New-Object System.Windows.Forms.Button -Property @{
    Text = ' Assign 0365 Licenses '
    AutoSize = $True
    ForeColor = "Teal"
    Location = New-Object System.Drawing.Point(63,240)
    }
    $assigno365LicenseButton.add_click({
    
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

$o365upn = $assigno365LicenseTextBox.Text

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

                
                                    $0365form = New-Object System.Windows.Forms.Form
                                    $0365form.Text = 'o365 Licenses'
                                    $0365form.Size = New-Object System.Drawing.Size(400,300)
                                    $0365form.StartPosition = 'CenterScreen'
                

                                    $0365OKButton = New-Object System.Windows.Forms.Button
                                    $0365OKButton.Location = New-Object System.Drawing.Point(120,220)
                                    $0365OKButton.Size = New-Object System.Drawing.Size(75,23)
                                    $0365OKButton.Text = 'OK'
                                    $0365OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
                                    $0365form.AcceptButton = $0365OKButton
                                    $0365form.AutoSize = $true
                                    $0365form.Controls.Add($0365OKButton)

                                    $0365CancelButton = New-Object System.Windows.Forms.Button
                                    $0365CancelButton.Location = New-Object System.Drawing.Point(200,220)
                                    $0365CancelButton.Size = New-Object System.Drawing.Size(75,23)
                                    $0365CancelButton.Text = 'Cancel'
                                    $0365CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                                    $0365form.CancelButton = $0365CancelButton
                                    $0365form.Controls.Add($0365CancelButton)

                                    $0365label = New-Object System.Windows.Forms.Label
                                    $0365label.Location = New-Object System.Drawing.Point(10,20)
                                    $0365label.Size = New-Object System.Drawing.Size(280,20)
                                    $0365label.Text = 'Please Select licenses:'
                                    $0365form.Controls.Add($0365label)

                                    $0365listBox = New-Object System.Windows.Forms.Listbox
                                    $0365listBox.Location = New-Object System.Drawing.Point(15,60)
                                    $0365listBox.Size = New-Object System.Drawing.Size(350,500)
                

                                    $0365listBox.SelectionMode = 'MultiSimple'


                                    $geto365licenses = Get-AzureADSubscribedSku | select skupartnumber
                

                                    foreach ($license in $geto365licenses) {
                                                [void] $0365listBox.Items.Add($license.skupartnumber)
                                                }
                

                                    $0365listBox.Height = 150
                                    $0365form.Controls.Add($0365listBox)
                                    $0365form.Topmost = $true

                                    $result = $0365form.ShowDialog()

                                        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {Write-Host "Selected licenses assigned!" -ForegroundColor Green} else {Write-Host "Cancelled! No licenses assigned!" -ForegroundColor Red}
                                            ForEach-Object {

                                                                foreach ($license in $0365listBox.SelectedItems)
                                                           
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

                
                                    $0365form = New-Object System.Windows.Forms.Form
                                    $0365form.Text = 'o365 Licenses'
                                    $0365form.Size = New-Object System.Drawing.Size(400,300)
                                    $0365form.StartPosition = 'CenterScreen'
                

                                    $0365OKButton = New-Object System.Windows.Forms.Button
                                    $0365OKButton.Location = New-Object System.Drawing.Point(120,220)
                                    $0365OKButton.Size = New-Object System.Drawing.Size(75,23)
                                    $0365OKButton.Text = 'OK'
                                    $0365OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
                                    $0365form.AcceptButton = $0365OKButton
                                    $0365form.AutoSize = $true
                                    $0365form.Controls.Add($0365OKButton)

                                    $0365CancelButton = New-Object System.Windows.Forms.Button
                                    $0365CancelButton.Location = New-Object System.Drawing.Point(200,220)
                                    $0365CancelButton.Size = New-Object System.Drawing.Size(75,23)
                                    $0365CancelButton.Text = 'Cancel'
                                    $0365CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                                    $0365form.CancelButton = $0365CancelButton
                                    $0365form.Controls.Add($0365CancelButton)

                                    $0365label = New-Object System.Windows.Forms.Label
                                    $0365label.Location = New-Object System.Drawing.Point(10,20)
                                    $0365label.Size = New-Object System.Drawing.Size(280,20)
                                    $0365label.Text = 'Please Select licenses:'
                                    $0365form.Controls.Add($0365label)

                                    $0365listBox = New-Object System.Windows.Forms.Listbox
                                    $0365listBox.Location = New-Object System.Drawing.Point(15,60)
                                    $0365listBox.Size = New-Object System.Drawing.Size(350,500)
                

                                    $0365listBox.SelectionMode = 'MultiSimple'


                                    $geto365licenses = Get-AzureADSubscribedSku | select skupartnumber
                

                                    foreach ($license in $geto365licenses) {
                                                [void] $0365listBox.Items.Add($license.skupartnumber)
                                                }
                

                                    $0365listBox.Height = 150
                                    $0365form.Controls.Add($0365listBox)
                                    $0365form.Topmost = $true

                                    $result = $0365form.ShowDialog()

                                        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {Write-Host "Selected licenses assigned!" -ForegroundColor Green} else {Write-Host "Cancelled! No licenses assigned!" -ForegroundColor Red}
                                            ForEach-Object {

                                                                foreach ($license in $0365listBox.SelectedItems)
                                                           
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
    
    })


    #==============================================================#
    #                   Setup O365 Email Signature                 #
    #==============================================================#

    $Setupo365EmailSig = New-Object System.Windows.Forms.Button -Property @{
    Text = ' Setup 0365 Email Signature '
    AutoSize = $True
    ForeColor = "Teal"
    Location = New-Object System.Drawing.Point(200,240)
    }
    $Setupo365EmailSig.add_click({
    
    if ($assigno365LicenseTextBox.Text -ne "") {

    cls
    
    #location of signature template
    $LocationOfSignature = "\\prlfs01\PDQ_Shared_DB\Repository\Signature_Creation\OWA_SignatureRebrand.htm"


    # See if connected to Exchange Online
 
    $AmIConnectedToExchangeOnline= Get-PSSession | 
    where {$_.ConfigurationName -eq "Microsoft.exchange"}


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
        $OWA_Sigs_To_Update= $assigno365LicenseTextBox.Text
      
      IF ($OWA_Sigs_To_Update -like "*@ruby.com")
            {
            $Stop= "Yes"
            }
                else 
                {
                Write-Warning "Email address does not contain @ruby.com domain...please try again."
                }

        }

        Until ($stop -eq "Yes" )

        $GetADUserInfo= Get-ADUser -Filter * -Properties title,department,enabled,emailaddress | 
        where {$_.title -notlike $Null -and $_.department -notlike $null}

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
    #$Signature | Out-File "C:\Users\chris.georgeson\Desktop\Office365 info\OutlookOnline_Recep_rollout\sigtest.htm"

        }

    Remove-PSSession $Session 


    write-host "
Signature should be created. Please check OWA to verify." -ForegroundColor Green
    
    } 
    Else {
          $invalidUPN = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
          $invalidUPN.Popup("UPN Required for setting up o365 Email Signature!",0,"Invalid Input - UPN")
         }
    })


    #==============================================================#
    #                     Remove Mailbox Rules                     #
    #==============================================================#

    $removeMailboxRules = New-Object System.Windows.Forms.Button -Property @{
    Text = ' Remove Mailbox Rules '
    AutoSize = $True
    ForeColor = "Teal"
    Location = New-Object System.Drawing.Point(367,240)
    }
    $removeMailboxRules.add_click({
    
    clear

$ErrorActionPreference = "SilentlyContinue"

$userCred = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $userCred -Authentication Basic -AllowRedirection
Import-PSSession $Session

clear

$whatUser = $assigno365LicenseTextBox.Text

$DoesUserExist = Get-Mailbox $whatUser | select name

if ($DoesUserExist -ne $null) {

$getrules = Get-InboxRule -Mailbox $whatUser | select -Property name


                Add-Type -AssemblyName System.Windows.Forms
                Add-Type -AssemblyName System.Drawing

                
                $DMRform = New-Object System.Windows.Forms.Form -Property @{
                Text = "Mailbox Rules for $whatUser"
                Size = New-Object System.Drawing.Size(400,300)
                StartPosition = 'CenterScreen'}
                

                $DMROKButton = New-Object System.Windows.Forms.Button -Property @{
                Location = New-Object System.Drawing.Point(120,220)
                Size = New-Object System.Drawing.Size(75,23)
                Text = 'OK'
                DialogResult = [System.Windows.Forms.DialogResult]::OK
                AutoSize = $true}
                $DMRform.AcceptButton = $DMROKButton
                $DMRform.Controls.Add($DMROKButton)

                $DMRCancelButton = New-Object System.Windows.Forms.Button -Property @{
                Location = New-Object System.Drawing.Point(200,220)
                Size = New-Object System.Drawing.Size(75,23)
                Text = 'Cancel'
                DialogResult = [System.Windows.Forms.DialogResult]::Cancel}
                $DMRform.CancelButton = $DMRCancelButton
                $DMRform.Controls.Add($DMRCancelButton)

                $DMRlabel = New-Object System.Windows.Forms.Label -Property @{
                Location = New-Object System.Drawing.Point(10,20)
                Size = New-Object System.Drawing.Size(280,20)
                Text = 'Please Select rules to remove:'}
                $DMRform.Controls.Add($DMRlabel)

                $DMRlistBox = New-Object System.Windows.Forms.Listbox -Property @{
                Location = New-Object System.Drawing.Point(15,60)
                Size = New-Object System.Drawing.Size(350,500)
                SelectionMode = 'MultiSimple'}
                
                
                Get-InboxRule -Mailbox $whatUser | 
                select name, from, FromAddressContainsWords, 
                SubjectContainsWords, BodyContainsWords, movetofolder, deletemessage,
                errortype | 
                Out-GridView -Title "Mailbox Rules for $whatUser"
                
                $getrules = Get-InboxRule -Mailbox $whatUser | select -Property name

                foreach ($rule in $getrules) {
                                              [void] $DMRlistBox.Items.Add($rule.name)
                                              }
                

                $DMRlistBox.Height = 150
                $DMRform.Controls.Add($DMRlistBox)
                $DMRform.Topmost = $true

                $result = $DMRform.ShowDialog()

                     if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                                                                               Write-Host "`nRules successfully queried!" -ForegroundColor Green
                                                                               } 
                        
                        else {
                              Write-Host "`nCancelled! No rules removed!" -ForegroundColor Red
                              }
                     
                     ForEach-Object {
    
                                        foreach ($rule in $DMRlistBox.SelectedItems)
                                        {
                                            try{
                                                Write-Host "`nRemoving the rule $rule" -ForegroundColor Yellow
                                                Remove-InboxRule "$rule" -Mailbox $whatUser -Force -Verbose
                                                }
                                                catch {
                                                       Write-Host "`n$rule does not exist! $whatUser not added!" -ForegroundColor Red
                                                       }
                                        
                                        }
        
                                    }

                                    }

    else {
          Write-Host "$whatUser does not exist!" -ForegroundColor Red
          }
    
    })


    #==============================================================#
    #   Mass Modify Group Members / Additional Responsibilities    #
    #==============================================================#

    # Mass Modify Function
    #==============================================================#
    Function MassUserEditor {

	#----------------------------------------------
	#region Import the Assemblies and resources
	#----------------------------------------------
	[void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.ServiceProcess, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')

    $LabelFonts = New-Object System.Drawing.Font("Verdana",8,[System.Drawing.FontStyle]::Bold)
    $FormIcon = New-Object System.Drawing.Icon ("$PSScriptRoot\Source\ros_app_icon.ico")

# Main Form Build
#=============================================================================================================#
$mainEditorForm = New-Object System.Windows.Forms.Form -Property @{
Size = '800,420'
Opacity = 0.95
StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
Icon = $FormIcon
Name = "Ruby User Modifier"
Text = "Ruby User Modifier"
}


# User List Label
#=============================================================================================================#
$UserListLabel = New-Object System.Windows.Forms.Label -Property @{
Autosize = $true
Text = "User List"
Font = $LabelFonts
Forecolor = "Teal"
location = '15,20'
}


# AD user listbox
#=============================================================================================================#
$UserList = New-Object System.Windows.Forms.ListBox -Property @{
size = '350,320'
Location = '15,40'
SelectionMode = 'MultiSimple'
Font = $LabelFonts
}

$GetActiveAccounts = get-aduser -Filter * -Properties Name, SamAccountName, Title, Department, Manager, EmployeeID, Enabled, UserPrincipalName |
                     Where {$_.title -notlike $NUll -and $_.department -notlike "$null" -and $_.manager -ne $null -and $_.employeeid -ne $Null -and $_.Enabled -eq $true} |
                     Sort -Property SamAccountName |
                     foreach {$UserList.Items.Add($_.samaccountname)}



# Additional Responsibilities Label
#=============================================================================================================#
$AdditionalResponsibilitiesLabel = New-Object System.Windows.Forms.Label -Property @{
Text = "Additional Responsibilities"
Autosize = $true
Font = $LabelFonts
Forecolor = "Teal"
Location = '400,20'
}


# Extended Hours Button
#=============================================================================================================#
$ExtHoursButton = New-Object System.Windows.Forms.Button -Property @{
Text = "Extended Hours"
Autosize = $true
Location = '400,40'
Forecolor = "Teal"
}
$ExtHoursButton.add_click({

if($UserList.SelectedItems -ne '') {

  foreach($User in $UserList.SelectedItems) {
       
         Write-Host "`r`nProvisioning Extended hours access for $($User)" -ForegroundColor Magenta
       
         Add-ADGroupMember -Identity ('Client' + ' ' + 'Happiness') $User -Verbose
         Add-ADGroupMember -Identity pearl_ch_users $User -Verbose
         Add-ADGroupMember -Identity ('Extended' + ' ' + 'Hours') $User -Verbose

  }
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

} Else {
        
        $NoUsersSelected = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
        $NoUsersSelected.Popup("No Users Selected!",0,"Invalid Input - Extended Hours") 

       }

}) # End of event handler scriptblock



# Group Drop Down Label
#=============================================================================================================#
$GroupsDropdownLabel = New-Object System.Windows.Forms.Label -Property @{
Text = "Group"
Autosize = $true
Font = $LabelFonts
Forecolor = "Teal"
Location = '400,131'
}



# AD group drop down list
#=============================================================================================================#
$AddToGroupList = New-Object System.Windows.Forms.ComboBox -Property @{
DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
Location = '400,150'
Size = '220,50'
DropDownHeight = '300'
}
$AddToGroupList.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend;
$AddToGroupList.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::CustomSource;


$getADgroups = Get-ADGroup -SearchBase 'OU=Ruby Groups,DC=ruby,DC=local' -Filter 'ObjectClass -eq "group"' |
               select Name, SamAccountName | Sort-Object -Property Name, SamAccountName

               foreach($GroupItem in $getADgroups) {
                                                    # Add Items to dropdown / ComboBox
                                                    $AddToGroupList.Items.Add($GroupItem.SamAccountName)
                                                    
                                                    # Add Items to Auto Complete Source
                                                    $AddToGroupList.AutoCompleteCustomSource.Add($GroupItem.SamAccountName)
                                                   }
                                                   clear


# Add to Group Button
#=============================================================================================================#
$MassAddtoGroupButton = New-Object System.Windows.Forms.Button -Property @{
Text = "Add Users"
Autosize = $true
Forecolor = "Teal"
Location = '400,180'
}                      
$MassAddtoGroupButton.add_click({

if(($UserList.SelectedItem -ne $null -and $AddToGroupList.SelectedItem -ne $null) -or ($UserList.SelectedItems -ne $null -and $AddToGroupList.SelectedItem -ne $null)){

    foreach($userItem in $UserList.SelectedItems) {
                                                  
                                                  Write-Host "Adding $($userItem) to the group $($AddToGroupList.SelectedItem)" -ForegroundColor Green
                                                  Add-ADGroupMember -Identity $AddToGroupList.SelectedItem $userItem -Verbose

                                                  }


} Else {
        
        $NoUsersToAdd = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
        $NoUsersToAdd.Popup("No Users or Group Selected!",0,"Invalid Selection - Add Users Error") 


       }

       #$SuccessUsersToAdd = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
       #$SuccessUsersToAdd.Popup("Selected User / Users Removed!",0,"Success - Removed Users from group $($AddToGroupList.SelectedItem)") 

})



# Remove from Group Button
#=============================================================================================================#
$MassRemoveFromGroupButton = New-Object System.Windows.Forms.Button -Property @{
Text = "Remove Users"
Autosize = $true
Forecolor = "Teal"
Location = '480,180'
}                      
$MassRemoveFromGroupButton.add_click({

if(($UserList.SelectedItem -ne $null -and $AddToGroupList.SelectedItem -ne $null) -or ($UserList.SelectedItems -ne $null -and $AddToGroupList.SelectedItem -ne $null)){

    foreach($userItemR in $UserList.SelectedItems) {
                                                  
                                                   Write-Host "Removing $($userItemR) from the selected group $($AddToGroupList.SelectedItem)" -ForegroundColor Magenta
                                                  
                                                   Remove-ADGroupMember $AddToGroupList.SelectedItem -Member $userItemR -Confirm:$false -Verbose

                                                  
                                                   }


} Else {
        
        $NoUsersToRemove = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
        $NoUsersToRemove.Popup("No Users or Group Selected!",0,"Invalid Selection - Remove Users Error") 


       }

       #$SuccessUsersToRemove = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
       #$SuccessUsersToRemove.Popup("Selected User / Users Removed!",0,"Success - Remove Users from group $($AddToGroupList.SelectedItem)") 

})



# Exchange Mass edit Section Label
#=============================================================================================================#
$ExchangeMassEditLabel = New-Object System.Windows.Forms.Label -Property @{
Text = "Exchange and O365"
Autosize = $true
Font = $LabelFonts
Forecolor = "Teal"
Location = '400,270'
}




# Setup Outlook Online Email Signature Button
#=============================================================================================================#
$MassSetupEmailSigButton = New-Object System.Windows.Forms.Button -Property @{
Text = "Setup o365 Email Signature"
Autosize = $true
Forecolor = "Teal"
Location = '400,290'
}
$MassSetupEmailSigButton.add_click({

if($UserList.SelectedItem -ne $null -or $UserList.SelectedItems -ne $null) {

#clear

# location of signature template
$LocationOfSignature= "\\prlfs01\PDQ_Shared_DB\Repository\Signature_Creation\OWA_SignatureRebrand.html"

# see if already connected to exchange online
$AmIConnectedToExchangeOnline= Get-PSSession | where {$_.ConfigurationName -eq "Microsoft.exchange"}


    IF ($AmIConnectedToExchangeOnline -eq $null -or $AmIConnectedToExchangeOnline.State -eq "Closed")
        {
        #Connect to Exchange Online
        
        $userCredentials= Get-Credential -message "Enter Exchange Online Email/Credentials to continue" 
        
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredentials -Authentication Basic -AllowRedirection
        
        Import-PSSession $Session -Verbose
         

        }

#clear

#Import the template for the signature

$importSignatureTemplate= Get-Content $LocationOfSignature

   
            foreach($SigUser in $UserList.SelectedItems) {
            
            
            $OWA_Sigs_To_Update = "$($SigUser)@ruby.com"

            $GetADUserInfo= Get-ADUser -Filter * -Properties title,department,enabled,emailaddress | where {$_.title -notlike $Null -and $_.department -notlike $null} -Verbose

            $UserDetails= $GetADUserInfo | where {$_.UserPrincipalName -eq $OWA_Sigs_To_Update} -Verbose

            $UserDetails

            Write-host "Found $($SigUser) in Active Directory." -ForegroundColor Magenta

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
    
#clear

write-host "`r`nSignatures should be created for $($SigUser) . Please check OWA to verify." -ForegroundColor Green


}

Remove-PSSession $Session




} Else {
        
        $NoUsersToSetupSig = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
        $NoUsersToSetupSig.Popup("No User / Users Selected!",0,"Invalid Selection - Selection Error") 

       }




})






# View Group Button
#=============================================================================================================#
$ViewGroupButton = New-Object System.Windows.Forms.Button -Property @{
Text = "View Group"
Autosize = $true
Forecolor = "Teal"
Location = '575,180'
}                      
$ViewGroupButton.add_click({

clear


if($AddToGroupList.SelectedItem -ne $null) {

# View Group info Form
#=======================================================================#
$GroupForm = New-Object System.Windows.Forms.Form -Property @{
Opacity = 0.95
StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
Icon = $FormIcon
Name = "Ruby User Modifier"
Size = '400,500'
Text = "Group info"
}

$ListGroupMembers = Get-ADGroupMember -identity "$($AddToGroupList.SelectedItem)" -Recursive | 
                    foreach{ get-aduser $_ -Properties SamAccountName } | 
                    Select SamAccountName


# Group Name Label
#=======================================================================#
$groupviewLabel = New-Object System.Windows.Forms.Label -Property @{
Text = "Group Name: $($AddToGroupList.SelectedItem)"
AutoSize = $true
Font = $LabelFonts
Forecolor = "Teal"
Location = '10,10'
}




# Member List Label
#=======================================================================#
$GroupMembersLabel = New-Object System.Windows.Forms.Label -Property @{
Text = "Current Members:"
Font = $LabelFonts
Forecolor = "Teal"
Autosize = $true
Location = '10,100'
}



# Goup Members List Textbox
#=======================================================================#
$groupMemberListTextBox = New-Object System.Windows.Forms.RichTextBox -Property @{
Size = '200,200'
Location = '10,120'
BackColor = "LightGray"
ReadOnly = $True
Multiline = $True
Font = $LabelFonts
}

foreach($groupmember in $ListGroupMembers){
                                           $groupMemberListTextBox.AppendText("`n$($groupmember.SamAccountName)")
                                          }



# Export Group Members to list Button
#=======================================================================#
$exportgroupmemberlistbutton = New-Object System.Windows.Forms.Button -Property @{
Text = "Export Member list"
Autosize = $true
Forecolor = "Teal"
Location = '243,118'
}
$exportgroupmemberlistbutton.add_click({

# Save CSV Save File Dialog
#==========================#
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog -Property @{
initialDirectory = "c:\"
filter = 'CSV (*.csv)|*.csv|SpreadSheet (*.xlsx)|*.xlsx'
FileName = "$($AddToGroupList.SelectedItem)"
}
$saveFile = $SaveFileDialog.ShowDialog()

if ($saveFile -eq [System.Windows.Forms.DialogResult]::OK) {

   foreach ($member in $ListGroupMembers) {"$($member.SamAccountName)" | Out-File "$($SaveFileDialog.FileName)" -Append -Verbose}

   clear

   $successfulExportNotif = New-Object -ComObject Wscript.Shell -ErrorAction Stop

   $successfulExportNotif.Popup("Exported Members list for $($AddToGroupList.SelectedItem)!",0,"File Name: $($SaveFileDialog.FileName)") 

} Elseif ($saveFile -eq [System.Windows.Forms.DialogResult]::Cancel) {
        
        $CancelExportNotif = New-Object -ComObject Wscript.Shell -ErrorAction Stop

        $CancelExportNotif.Popup("Cancelled export for $($AddToGroupList.SelectedItem)!",0,"Export Cancelled: $($SaveFileDialog.FileName)") 

       }

})



$groupinfoCompArray = @(
                       $groupviewLabel,
                       $GroupMembersLabel,
                       $groupMemberListTextBox,
                       $exportgroupmemberlistbutton
                       )
                       foreach($infoComp in $groupinfoCompArray){$GroupForm.Controls.Add($infoComp)}





$showGroupForm = $GroupForm.ShowDialog()

} Else {
        
        $NoGroupSelected = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
        $NoGroupSelected.Popup("No Group Selected!",0,"Invalid Selection - View Group") 


       }




})
                      








# Form Components array / build
#==============================#

$compntArray = @(
                 $UserListLabel,
                 $UserList,
                 $GroupsDropdownLabel,
                 $AddToGroupList,
                 $AdditionalResponsibilitiesLabel,
                 $ExtHoursButton,
                 $MassAddtoGroupButton,
                 $MassRemoveFromGroupButton,
                 $ViewGroupButton,
                 $ExchangeMassEditLabel,
                 $MassSetupEmailSigButton
                )
                foreach ($compnt in $compntArray) {
                                                   $mainEditorForm.Controls.Add($compnt)
                                                  }








$result = $mainEditorForm.ShowDialog()

[System.GC]::Collect()
$error.Clear()


}

    
    
    # Mass Modify Label
    #==============================================================#
    $MassModifyLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Mass Modify Group Memberships / Provision Additional Responsibilities"
    AutoSize = $True
    Font = $LabelFonts
    Forecolor = "Teal"
    Location = '30,350'
    }


    # Mass Modify Button
    #==============================================================#
    $MassModifyButton = New-Object System.Windows.Forms.Button -Property @{
    Text = "Launch"
    Autosize = $True
    Forecolor = "Teal"
    Location = '30,380'
    }
    $MassModifyButton.add_click({MassUserEditor})

    #==============================================================#
    #             Mass Assign Yammer Licenses VIA CSV              #
    #==============================================================#

    # Mass Assign Yammer Feature Licenses Label
    #==============================================================#
    $MassAssignYammerLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Mass Assign Yammer Licenses"
    Font = $LabelFonts
    Forecolor = "Teal"
    Autosize = $True
    Location = '30,430'
    }

    # Mass Assign Yammer License Function
    #==============================================================#
    Function AssignYammerFeatureLicense {

$ErrorActionPreference = 'SilentlyContinue'


[void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
[void][reflection.assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
[void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][reflection.assembly]::Load('System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][reflection.assembly]::Load('System.ServiceProcess, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')



# Connect to MSOnline - Check if module exists, Installs Module if missing
#=========================================================================
$DoesMSOlineModuleExist = Get-Module -Name MSOnline

if ($DoesMSOlineModuleExist -ne $null) {
                                        Connect-MsolService
                                       }

   Else {
         Write-Host "MSOline Module missing, Installing......" -ForegroundColor Yellow
         Install-Module -Name MSOline -Verbose -Force -ErrorAction SilentlyContinue
         Connect-MsolService
        }

#Import CSV for receps being updated
#=================================== 
$importcsvFile = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
           InitialDirectory = '\\prlfs01\ISWizards\SuperScript\Tools\MassAssignYammerLicenses' 
           Filter = 'CSV (*.csv)|*.csv|SpreadSheet (*.xlsx)|*.xlsx'
           }
           $importcsvFile.ShowDialog()

Import-Csv $importcsvFile.FileName | ForEach-Object {


$UPN = $_.UserPrincipalName
$LicenseDetails = (Get-MsolUser -UserPrincipalName $UPN).Licenses

ForEach ($License in $LicenseDetails) {
                                       $DisabledOptions = @()
                                       $License.ServiceStatus | ForEach {
     
     If ($_.ProvisioningStatus -eq "Disabled" -and $_.ServicePlan.ServiceName -notlike "*YAMMER*") { 
                                                                                                    
                                                                                                    $DisabledOptions += "$($_.ServicePlan.ServiceName)" 
                                                                                                    
                                                                                                    } 
                                                                        }
   
   $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $License.AccountSkuId -DisabledPlans $DisabledOptions
   
   Set-MsolUserLicense -UserPrincipalName $UPN -LicenseOptions $LicenseOptions

   

}
Write-Host "Yammer Feature License Assigned for $($UPN)!" -ForegroundColor Green
}

}

    # Mass Assign Yammer Licenses via CSV button
    #==============================================================#
    $MassAssignYammerButton = New-Object System.Windows.Forms.Button -Property @{
    Autosize = $True
    Location = '30,460'
    ForeColor = "Teal"
    Text = 'Mass Assign'
    }
    $MassAssignYammerButton.add_click({AssignYammerFeatureLicense})


    #======================================================================#
    #   Verify / Refresh Frontline Exchange and Teams Retention Policies   #
    #======================================================================#

    # Verify / Refresh Retention Policy Function
    #==============================================================#
    Function AssignRetentionPolicies {

$ErrorActionPreference = 'silentlycontinue'

clear

Write-Host """
+====================================================+
|                                                    |
|       Retention Policy Assignment Script           |
|                                                    |
| Steps:                                             |
|                                                    |
|      1. Provide UPN credentials at prompt          |
|                                                    |
| Notes: This script will apply retention policies   |
|        for Exchange and Teams. This script should  |
|        be run every 30 days.                       |
|                                                    |
|                                                    |
+====================================================+
""" -ForegroundColor Green


# Provide Admin Credentials
#==========================
$LoginCredentials = Get-Credential


# Connection URI variables
#=========================
$EXOnlineURI = 'https://outlook.office365.com/powershell-liveid/'
$ComplianceURI = 'https://ps.compliance.protection.outlook.com/powershell-liveid/'


# Retention Policy Variables
#===========================
$MailboxRetentionPolicy = "Ruby - F3 User Remove/Delete emails that are 1 year and older"
$TeamsRetentionPolicy = "Ruby - Teams - 1:1 Chat F1  Standard Retention Policy"



# Query AD users based on Criteria
#=================================
$QueryUsers = get-aduser -Filter * -Properties title,department,enabled,employeeid,emailaddress,manager,PasswordLastSet,PasswordExpired,DistinguishedName,Created,UserPrincipalName,SamAccountName |
              Where {$_.title -notlike $NUll -and $_.department -notlike "$null" -and $_.manager -ne $null -and $_.employeeid -ne $Null -and $_.Enabled -eq $true}

$FrontLineEmployees = ($QueryUsers | where {$_.title -like "Chat Specialists (*)" -or $_.title -like "* Receptionist" })


# Verify AzureAD Module exists and connect
#=======================================================
$DoesAzureADModuleExist = Get-Module -Name AzureAD

if ($DoesAzureADModuleExist -ne $null) {
                                        Connect-AzureAD
                                       }

   Else {
         Write-Host "AzureAD Module missing, Installing......" -ForegroundColor Yellow
         Install-Module -Name AzureAD -Verbose -Force
         Connect-AzureAD
        }



# Exchange Online PS Session initializaion
#=========================================
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $EXOnlineURI -Credential $LoginCredentials -Authentication Basic -AllowRedirection
Import-PSSession $Session

$ErrorActionPreference = 'silentlycontinue'

foreach ($FrontEmp in $FrontLineEmployees){

# Font Color for template
#========================
$RedFontWarning = "color:red"

# Email HTMLBody
#===============
$emailBody = "
<p style=$($RedFontWarning)>Retention policy applied for  $($FrontEmp.SamAccountName)
<br>
<p><b>Qualifying Conditions:</b> User is a Frontline Employee and assigned an F3 license
<br>
<p>Applied the following Retention Policies: 
<br>
<p><b>Exchange:</b> Ruby - F3 User Remove/Delete emails that are 1 year and older
<br>
<p><b>Teams:</b> Ruby - Teams - 1:1 Chat F1  Standard Retention Policy
<br>
</p>
</p>
</p>
</p>
</p>
"


# Check Azure AD user to verify if user has an E3 or F3 license
#==============================================================
$VerifyE3License = Get-AzureADUserLicenseDetail -ObjectId $FrontEmp.UserPrincipalName | 
                   Select SkuID, SkuPartNumber |
                   Where SkuPartNumber -EQ "ENTERPRISEPACK" -Verbose

$verifyF3License = Get-AzureADUserLicenseDetail -ObjectId $FrontEmp.UserPrincipalName | 
                   Select SkuID, SkuPartNumber |
                   Where SkuPartNumber -EQ "SPE_F1" -Verbose


# Verify if user already has Retention Policy assigned
#=====================================================
$VerifyRetentionPolicy = Get-Mailbox -Identity "$($FrontEmp.SamAccountName)" |
                         Select RetentionPolicy



# If user has an F3 License Set Exchange Retention Policy
#========================================================
if ($verifyF3License -ne $null) {

                                # If User does not have a retention policy already applied
                                #=========================================================
                                if ($VerifyRetentionPolicy.RetentionPolicy -ne $MailboxRetentionPolicy) {

                                # Send message to Host Console, Apply Retention Policy, and send email notification to wizards distro
                                #====================================================================================================
                                Try {
                                     
                                     Write-Host "$($FrontEmp.SamAccountName) has an F3 License. Applying Retention Policy" -ForegroundColor Green
                                     Set-Mailbox $FrontEmp.SamAccountName -RetentionPolicy $MailboxRetentionPolicy -Verbose
                                     Send-MailMessage -SmtpServer PRLEMC01 -Body $emailBody -BodyAsHtml -From RetentionPolicyNotifs@ruby.com -To wizards@ruby.com -Subject "Retention Policy Applied for $($FrontEmp.SamAccountName)" -Priority High -Verbose

                                    }

                                    Catch {
                                    
                                           Write-Host "$($FrontEmp.SamAccountName) is already a member!" -ForegroundColor Red
                                          
                                          }
                           
                                }
                                
                                # If user is already assigned a retention policy and has an F3 License
                                #=====================================================================
                                Else {
                                      Write-Host "$($FrontEmp.SamAccountName) already is assign the Retention Policy:`r`n$($MailboxRetentionPolicy)" -ForegroundColor Cyan
                                     }
                              
                                }

# If user has an E3 License - Remove Retention Policy
#====================================================
if ($VerifyE3License -ne $null) {

                                Try {
                               
                                     Write-Host "$($FrontEmp.SamAccountName) has an E3 License. Removing Retention Policy" -ForegroundColor Yellow
                                     Set-Mailbox $FrontEmp.SamAccountName -RetentionPolicy $null -Verbose
                               
                                    }

                                    Catch {
                                   
                                          Write-Host "Policy was already removed!" -ForegroundColor Red
                                   
                                          }
    
    
    
    
    }

}


# Gets and removes Exchange PSSession
#====================================
Get-PSSession | Remove-PSSession

clear

# Security and Compliance Online PS Session initializaion
#========================================================
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ComplianceURI -Credential $LoginCredentials -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking

Write-Host "Applying the following teams retention policy: `r`n$($TeamsRetentionPolicy)`r`n" -ForegroundColor Green

foreach ($teamsFrontuser in $FrontLineEmployees) {
                                 
                                 Try {
                                      Write-Host "Applying Teams Retention Policy for $($teamsFrontuser.SamAccountName)" -ForegroundColor Magenta
                                      Set-TeamsRetentionCompliancePolicy -Identity $TeamsRetentionPolicy -AddTeamsChatLocation $teamsFrontuser.SamAccountName -RemoveTeamsChatLocation All -Verbose
                                     }
                                     
                                     catch {
                                            Write-Host "$($teamsFrontuser.SamAccountName) is already a member!" -ForegroundColor Red
                                           }
                                 
                                 }


Get-PSSession | Remove-PSSession

}

    # Verify / Refresh Frontline Retention Policies Label
    #==============================================================#
    $RefreshRetentionPolicyLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Refresh Frontline Ruby Exchange and Teams Retention Policies"
    Font = $LabelFonts
    Forecolor = "Teal"
    Autosize = $True
    Location = '30,510'
    }

    # Verify / Refresh Frontline Retention Policies Button
    #==============================================================#
    $RefreshRetentionPolicyButton = New-Object System.Windows.Forms.Button -Property @{
    Autosize = $True
    Location = '30,540'
    ForeColor = "Teal"
    Text = "Refresh Policies"
    }
    $RefreshRetentionPolicyButton.add_click({AssignRetentionPolicies})
                                      
    #================================================================#
    # Sub tab controller for User management tab                     #
    #================================================================#
    
    $tabcontrol2 = New-Object 'System.Windows.Forms.TabControl' -Property @{
    Size = '622, 650'
    AutoSize = $True
    BackColor = "Gray"
    Alignment = "Top"}
    $tabpage1.Controls.Add($tabcontrol2)


    # Add Sub Tab 1 Onboarding Components
    #================================================================
    $tabcontrol2.Controls.Add($subtabpageOnboard)

    $onboardingComponentsArray = @($tabpage1LabelEnableUsers,
                                   $NewRubyMainLabel,
                                   $NextAdventureMainLabel,
                                   $NADVSales,
                                   $NADVCH, 
                                   $EnablePRLButton,
                                   $EnableBVTButton,
                                   $EnableKCButton,
                                   $EnableSalesButton,
                                   $EnableGAButton,
                                   $onboardUsernameLabel,
                                   $onboardUsernameTB,
                                   $onboardingTeamNameLabel,
                                   $onboardTeamTB,
                                   $onboardingManagerLabel,
                                   $onboardManagerTB,
                                   $onboardingstationLabel,
                                   $onboardStationTB,
                                   $EnableVerifyButton,
                                   $EnableClearButton,
                                   $onboardingTemplateMainLabel,
                                   $OBTUserTextboxLabel,
                                   $OBTUserTextbox,
                                   $OBTNACheckbox,
                                   $OBTRTRCheckbox,
                                   $OBTManagerTextboxLabel,
                                   $OBTManagerTextbox,
                                   $OBTTeamLabel,
                                   $OBTTeamTextbox,
                                   $templateListBox,
                                   $RefreshTemplateListButton,
                                   $CreateTemplateButton,
                                   $RemoveTemplateButton,
                                   $reviewSelectedTemplateButton,
                                   $EnableRubyViaTemplateButton
                                   )

      foreach ($ArrayOnboardComp in $onboardingComponentsArray) 
              {
               $subtabpageOnboard.Controls.Add($ArrayOnboardComp)
              }

    
    # Add Sub Tab 2 Departing Components
    #================================================================
    $tabcontrol2.Controls.Add($subtabpageDepart)

    $departingComponentsArray = @($departLabel,
                                  $importcsvbutton,
                                  $showselectedCSVLabel,
                                  $showselectedCSV,
                                  $verifyDepartGrid,
                                  $verifyDepartedButton,
                                  $departButton,
                                  $clearImportedCSVButton,
                                  $setExpireLabel,
                                  $BlockMSOLLogin,
                                  $BlockMSOLoginTB,
                                  $BlockMSOLoginButton,
                                  $setExpireButton,
                                  $expireInstructions
                                  )

        foreach ($ArrayDepartComp in $departingComponentsArray)
                {
                $subtabpageDepart.Controls.Add($ArrayDepartComp)
                }
    

    # Add Sub Tab 3 Modify Ruby Components
    #================================================================ 
    $tabcontrol2.Controls.Add($subtabpageModify)

    $ModifyRubyComponentsArray = @(
                                   $ModifyRubyMainLabel,
                                   $ModifyRubyExchO365Label,
                                   $ModifyRubyUsernameLabel,
                                   $ModifyRubyUsernameTextbox,
                                   $ExtHoursProvisionButton,
                                   $addUserToMultGroupsButton,
                                   $assigno365LicenseUserLabel,
                                   $assigno365LicenseTextBox,
                                   $assigno365LicenseButton,
                                   $Setupo365EmailSig,
                                   $removeMailboxRules,
                                   $MassModifyLabel,
                                   $MassModifyButton,
                                   $MassAssignYammerLabel,
                                   $MassAssignYammerButton,
                                   $RefreshRetentionPolicyLabel,
                                   $RefreshRetentionPolicyButton
                                   )
          
          foreach ($modComp in $ModifyRubyComponentsArray)
                  {
                  $subtabpageModify.Controls.Add($modComp)
                  }
    
    
	#===============================================================================#
	#                              tabpage2 Applications                            #
	#===============================================================================#
	$tabpage2.Location = '23, 4'
	$tabpage2.Name = 'tabpage2'
	$tabpage2.Padding = '3, 3, 3, 3'
	$tabpage2.Size = '602, 442'
	$tabpage2.TabIndex = 1
	$tabpage2.Text = 'Applications'
	$tabpage2.UseVisualStyleBackColor = $True
	
    
    #===============================================================================#
    #                     Assist Tracker Sub tab for application                    #
    #===============================================================================#
    $subtabpageAssistTrack = New-Object 'System.Windows.Forms.TabPage' -Property @{
    Location = '0, 15'
    Padding = '3, 3, 3, 3'
    Size = '602, 442'
    TabIndex = 9
    Text = ' Assist Tracker '
    UseVisualStyleBackColor = $True}


    # Assist Tracker Main Label
    #===============================================================================#
    $InstallAssistTrackerLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Deploy Assist Tracker"
    ForeColor = "Teal"
    AutoSize = $True
    Font = $LabelFonts
    Location = New-Object System.Drawing.Point(30,30)
    }


    # Assist Tracker Hostname Label
    #===============================================================================#
    $InstallAssistTrackerHostnameLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Hostname:"
    AutoSize = $True
    Location = New-Object System.Drawing.Point(30,100)
    }


    # Assist Tracker Hostname Textbox
    #===============================================================================#
    $AssistTrackerHostnameTextbox = New-Object System.Windows.Forms.TextBox -Property @{
    BackColor = "LightGray"
    Size = '100,30'
    Location = New-Object System.Drawing.Point(90,98)
    }


    # Deploy Assist Tracker Button
    #===============================================================================#
    $DeployAssistTrackerButton = New-Object System.Windows.Forms.Button -Property @{
    Text = ' Deploy Assist Tracker '
    ForeColor = "Teal"
    AutoSize = $True
    Location = New-Object System.Drawing.Point(197,96)
    } 
    $DeployAssistTrackerButton.add_click({
    
        if ($AssistTrackerHostnameTextbox.Text -ne "") {

        clear
        
        Write-Host "-----Assist Tracker Shortcut Creation Script-----" -ForegroundColor Green -BackgroundColor Blue


            $ErrorActionPreference = 'silentlycontinue'
            $hostname = $AssistTrackerHostnameTextbox.Text
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

Write-Host "Assist Tracker Shortcuts added!"  -ForegroundColor Green
}

Else {
      clear
      Write-Host "Hostname is required!" -ForegroundColor Red
      $NoHostAssist = New-Object -ComObject Wscript.Shell -ErrorAction Stop
      $NoHostAssist.Popup("Hostname is Required!",0,"Input Error!")
      }
    
    })


    # Add Components via Array Variable to Assist Tracker Sub Tab
    #===============================================================================#
    
    $AssistTrackArray = @(
                          $InstallAssistTrackerLabel,
                          $InstallAssistTrackerHostnameLabel,
                          $AssistTrackerHostnameTextbox,
                          $DeployAssistTrackerButton
                          )

           foreach ($AssistComp in $AssistTrackArray)
                   {
                   $subtabpageAssistTrack.Controls.Add($AssistComp)
                   }



    #===============================================================================#
    #                          ROS Sub tab for applications                         #
    #===============================================================================#
    $subtabpageRubyOS = New-Object 'System.Windows.Forms.TabPage' -Property @{
    Location = '0, 15'
    Padding = '3, 3, 3, 3'
    Size = '602, 442'
    TabIndex = 9
    Text = ' Ruby OS '
    UseVisualStyleBackColor = $True}

    # Ros Picture
    #===============================================================================#
    $ROSIMG = [System.Drawing.Image]::FromFile("$PSScriptRoot\Source\ros_app_icon.jpg")
    $ReinstallRosPicture = New-Object System.Windows.Forms.PictureBox -Property @{
    BackgroundImage = $ROSIMG
    size = '96,97'
    location = New-Object System.Drawing.Point(420,45)
    }

    # Reinstall Ros Main Label
    #===============================================================================#
    $ReinstallRosMainLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Re-Install Ruby OS"
    ForeColor = "Teal"
    AutoSize = $True
    Font = $LabelFonts
    Location = New-Object System.Drawing.Point(30,30)
    }

    # Reinstall Ros Hostname Label
    #===============================================================================#
    $ReinstallRosHostnameLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Hostname:"
    AutoSize = $True
    Location = New-Object System.Drawing.Point(20,70)
    }

    # Reinstal Ros Hostname Textbox
    #===============================================================================#
    $ReinstallRosHostnameTextbox = New-Object System.Windows.Forms.TextBox -Property @{
    BackColor = "LightGray"
    size = '150,30'
    Location = New-Object System.Drawing.Point(80,69)
    }

    # Reinstall Ros Station Number Label
    #===============================================================================#
    $ReinstallRosStationLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Station:"
    AutoSize = $True
    Location = New-Object System.Drawing.Point(35,110)
    }

    # Reinstall Ros Station Number Textbox
    #===============================================================================#
    $ReinstallRosStationTextbox = New-Object System.Windows.Forms.TextBox -Property @{
    BackColor = "LightGray"
    Size = '100,30'
    Location = New-Object System.Drawing.Point(80,108)
    }

    # Reinstall Ros Button
    #===============================================================================#
    $ReinstallRosButton = New-Object System.Windows.Forms.Button -Property @{
    Text = ' Reinstall '
    ForeColor = "Teal"
    AutoSize = $True
    Location = New-Object System.Drawing.Point(250,67)
    }
    $ReinstallRosButton.add_click({

    if ($ReinstallRosHostnameTextbox.Text -ne "" -and $ReinstallRosStationTextbox.Text -ne "") {
    
    clear

    Write-Host "-----Update Ruby OS-----" -ForegroundColor Green -BackgroundColor Blue
    Write-Host ""
    Write-Host ""
    Write-Host "This Script will replace the RubyOS on any system Specified."

$ErrorActionPreference = 'silentlycontinue'
$hostname = $ReinstallRosHostnameTextbox.Text
$stationNumb = $ReinstallRosStationTextbox.Text
$source = ($PSScriptRoot + '\Applications\RubyOS\*.*')
$destination = ('\\' + $hostname + '\C$\RubyOS')

(gwmi Win32_Process -ComputerName $hostname | ?{ $_.ProcessName -match "RubyOS"}).Terminate()

rd -Path ('\\' + $hostname + '\C$\RubyOS') -Verbose -Recurse -Force 
md -Path ('\\' + $hostname + '\C$\RubyOS')

Start-BitsTransfer -Source $source -Destination $destination -Description "Copy Status.." -DisplayName "Copying..."

    (Get-Content \\$($hostname)\c$\RubyOS\settings.xml) -replace '%station%', $stationNumb | Set-Content \\$($hostname)\c$\RubyOS\settings.xml -Verbose

clear

Write-Host "RubyOS Updated on system $($hostname)"  -ForegroundColor Green
}
Else {
      clear
      Write-Host "Hostname and station required!" -ForegroundColor Red
      $NoHostOrStation = New-Object -ComObject Wscript.Shell -ErrorAction Stop
      $NoHostOrStation.Popup("Hostname and Station Required!",0,"Input Error!")
      }

    })

    # Reinstall Ros Instuctions Textbox
    #===============================================================================#
    $ReinstallRosUnstructions = New-Object System.Windows.Forms.TextBox -Property @{
    ReadOnly = $True
    Multiline = $True
    ForeColor = "Black"
    BackColor = "White"
    Size = New-Object System.Drawing.Point(550,200)
    Location = New-Object System.Drawing.Point(30,210)
    ScrollBars = "Vertical"
    BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3d
    Text = "Reinstall Ruby OS Instructions"}
    $ReinstallRosUnstructions.AppendText("`r`n===============================
    `r`n
    `r`nStep 1: Type in the desired hostname
    `r`n
    `r`nStep 2: Type in the desired station number Example: 89000
    `r`n
    `r`nStep 3: Left Click Reinstall
    `r`n
    `r`n        Note: This will Reinstall Ros and set the station
    `r`n
    `r`n")


    # Add Components to RubyOS Subtab
    #===============================================================================#
    
    $RubyOSArray = @(
                     $ReinstallRosPicture,
                     $ReinstallRosMainLabel,
                     $ReinstallRosHostnameLabel,
                     $ReinstallRosHostnameTextbox,
                     $ReinstallRosStationLabel,
                     $ReinstallRosStationTextbox,
                     $ReinstallRosButton,
                     $ReinstallRosUnstructions
                     )
         
         foreach ($RubyOSCcomp in $RubyOSArray)
                 {
                 $subtabpageRubyOS.Controls.Add($RubyOSCcomp)
                 }
                 
    
    
    
    #===============================================================================#
    #                   Microsoft Teams Sub tab for applications                    #
    #===============================================================================#
    $subtabpageTeams = New-Object 'System.Windows.Forms.TabPage' -Property @{
    Location = '0, 15'
    Padding = '3, 3, 3, 3'
    Size = '602, 442'
    TabIndex = 9
    Text = ' Microsoft Teams '
    UseVisualStyleBackColor = $True}

    # Teams Picture
    #===============================================================================#
    $ClearTeamsCacheIcon = [System.Drawing.Image]::FromFile("$PSScriptRoot\Source\TeamsLogo.jpg")
    $ClearTeamsCachePicture = New-Object System.Windows.Forms.PictureBox -Property @{
    BackgroundImage = $ClearTeamsCacheIcon
    size = '100,100'
    location = New-Object System.Drawing.Point(420,75)
    }

    # Clear Teams Cache Main Label
    #===============================================================================#
    $ClearTeamsCacheMainLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Clear Microsoft Teams cache"
    AutoSize = $True
    Font = $LabelFonts
    ForeColor = "Teal"
    Location = New-Object System.Drawing.Point(30,30)
    }

    # Hostname Label
    #===============================================================================#
    $ClearTeamsCacheHostnameLable = New-Object System.Windows.Forms.Label -Property @{
    Text = "Hostname:"
    AutoSize = $True
    Location = New-Object System.Drawing.Point(30,100)
    }
    
    # Hostname Textbox
    #===============================================================================#
    $ClearTeamsCacheHostnameTextbox = New-Object System.Windows.Forms.TextBox -Property @{
    BackColor = "LightGray"
    Size = '100,30'
    Location = New-Object System.Drawing.Point(90,98)
    }

    # Username Label
    #===============================================================================#
    $ClearTeamsCacheUserNameLable = New-Object System.Windows.Forms.Label -Property @{
    Text = "Username:"
    AutoSize = $True
    Location = New-Object System.Drawing.Point(30,150)
    }

    # Username Textbox
    #===============================================================================#
    $ClearTeamsCacheUsernameTextbox = New-Object System.Windows.Forms.TextBox -Property @{
    BackColor = "LightGray"
    Size = '120,30'
    Location = New-Object System.Drawing.Point(90,147)
    }

    # Clear Teams Cache Button
    #===============================================================================#
    $ClearTeamsCacheButton = New-Object System.Windows.Forms.Button -Property @{
    Text = ' Clear Cache '
    ForeColor = "Teal"
    AutoSize = $True
    Location = New-Object System.Drawing.Point(220,145)
    }
    $ClearTeamsCacheButton.add_click({
    
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
| EX: firstname.lastname@ruby.com           |
|                                               |
=================================================
`n" -ForegroundColor Yellow

$TeamsHost = $ClearTeamsCacheHostnameTextbox.Text
$TeamsUser = $ClearTeamsCacheUsernameTextbox.Text
$isTeamsHostOnline = Test-Connection -ComputerName $TeamsHost | Select-Object -Property IPV4Address
$TeamsProcess = ('\\' + $TeamsHost + '\c$\Users\' + $TeamsUser + '\AppData\Local\Microsoft\Teams\current\Teams.exe')

    if ($isTeamsHostOnline -ne $null){
            
            $ErrorActionPreference = 'silentlycontinue' 

        try {
            (gwmi Win32_Process -ComputerName $TeamsHost | ?{ $_.ProcessName -like "*Teams*"}).Terminate()
            }
            catch {
                   Write-Host "Teams Process is not running on $($TeamsHost)" -ForegroundColor Yellow
                   }
            Try {
                 Remove-Item -Path ('\\' + $TeamsHost + '\c$\Users\' + $TeamsUser + '\AppData\Roaming\Microsoft\Teams\' + 'Application Cache') -Recurse -Force -Verbose
                 } Catch {Write-Host "Path does not exist!" -ForegroundColor Yellow}
            Try {
                 Remove-Item -Path ('\\' + $TeamsHost + '\c$\Users\' + $TeamsUser + '\AppData\Roaming\Microsoft\Teams\' + 'blob_storage') -Recurse -Force -Verbose
                 } Catch {Write-Host "Path does not exist!" -ForegroundColor Yellow}
            Try {
                 Remove-Item -Path ('\\' + $TeamsHost + '\c$\Users\' + $TeamsUser + '\AppData\Roaming\Microsoft\Teams\' + 'Cache') -Recurse -Force -Verbose
                 } Catch {Write-Host "Path does not exist!" -ForegroundColor Yellow}
            Try {
                 Remove-Item -Path ('\\' + $TeamsHost + '\c$\Users\' + $TeamsUser + '\AppData\Roaming\Microsoft\Teams\' + 'databases') -Recurse -Force -Verbose
                 } Catch {Write-Host "Path does not exist!" -ForegroundColor Yellow}
            Try {
                 Remove-Item -Path ('\\' + $TeamsHost + '\c$\Users\' + $TeamsUser + '\AppData\Roaming\Microsoft\Teams\' + 'GPUCache') -Recurse -Force -Verbose
                 } Catch {Write-Host "Path does not exist!" -ForegroundColor Yellow}
            Try {
                 Remove-Item -Path ('\\' + $TeamsHost + '\c$\Users\' + $TeamsUser + '\AppData\Roaming\Microsoft\Teams\' + 'IndexedDB') -Recurse -Force -Verbose
                 } Catch {Write-Host "Path does not exist!" -ForegroundColor Yellow}
            Try {
                 Remove-Item -Path ('\\' + $TeamsHost + '\c$\Users\' + $TeamsUser + '\AppData\Roaming\Microsoft\Teams\' + 'Local Storage' ) -Recurse -Force -Verbose
                 } Catch {Write-Host "Path does not exist!" -ForegroundColor Yellow}
            Try {
                 Remove-Item -Path ('\\' + $TeamsHost + '\c$\Users\' + $TeamsUser + '\AppData\Roaming\Microsoft\Teams\' + 'tmp' ) -Recurse -Force -Verbose
                 } Catch {Write-Host "Path does not exist!" -ForegroundColor Yellow}
            Start-Process -FilePath $TeamsProcess -Verbose
       }
        else {
              Write-Host "$($TeamsHost) is either not online or not able to connect!" -ForegroundColor Red
             }

write-host "Teams Roaming Cache cleared on system $($TeamsHost)" -ForegroundColor Green
    
    })

    # Clear Teams Cache instructions
    #===============================================================================#

    # Add Components to Microsoft Teams Subtab
    #===============================================================================#

    $TeamsArray = @(
                    $ClearTeamsCacheMainLabel,
                    $ClearTeamsCachePicture,
                    $ClearTeamsCacheHostnameLable,
                    $ClearTeamsCacheHostnameTextbox,
                    $ClearTeamsCacheUserNameLable,
                    $ClearTeamsCacheUsernameTextbox,
                    $ClearTeamsCacheButton
                    )
          foreach ($TeamsComp in $TeamsArray)
                  {
                  $subtabpageTeams.Controls.Add($TeamsComp)
                  }

    # Application Tab Controller
    #==============================================================================#
    $tabcontrol3 = New-Object 'System.Windows.Forms.TabControl'
    $tabcontrol3.Size = '622, 650'
    $tabcontrol3.AutoSize = $True
    $tabpage2.Controls.Add($tabcontrol3)
    $tabcontrol3.Controls.Add($subtabpageAssistTrack)
    $tabcontrol3.Controls.Add($subtabpageTeams)
    $tabcontrol3.Controls.Add($subtabpageRubyOS)

    #===============================================================================#
	#                                 tabpage3 OS Tools                             #
	#===============================================================================#
	$tabpage3.Location = '23, 4'
	$tabpage3.Name = 'tabpage3'
	$tabpage3.Padding = '3, 3, 3, 3'
	$tabpage3.Size = '602, 442'
	$tabpage3.TabIndex = 2
	$tabpage3.Text = 'OS / System Tools'
	$tabpage3.UseVisualStyleBackColor = $True
    
    # System Information Main Label
    #===============================================================================#
    $SystemInfoMainLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "System / Network Information"
    ForeColor = "Teal"
    Font = $LabelFonts
    AutoSize = $True
    Location = New-Object System.Drawing.Point(10,280)
    }

    # System Info Hostname Label
    #===============================================================================#
    $SystemInfoHostnameLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Hostname:"
    AutoSize = $True
    Location = New-Object System.Drawing.Point(10,320)
    }

    # System Information Hostname Textbox
    #===============================================================================#
    $SystemInfoHostnameTextbox = New-Object System.Windows.Forms.TextBox -Property @{
    BackColor = "LightGray"
    Size = '120,30'
    Location = New-Object System.Drawing.Point(70,317)
    }

    # System Information Get Info Button
    #===============================================================================#
    $SystemInfoGetButton = New-Object System.Windows.Forms.Button -Property @{
    Text = ' Get Info '
    ForeColor = "Teal"
    AutoSize = $True
    Location = New-Object System.Drawing.Point(200,315)
    }
    $SystemInfoGetButton.add_click({
    
    if ($SystemInfoHostnameTextbox.Text -ne $null) {
        
        $ErrorActionPreference = 'silentlycontinue'
            
        $SystemInfoDisplayGrid.Columns[0].Name = "$($SystemInfoHostnameTextbox.Text).ruby.local"

             $SystemInfoDisplayGrid.rows.clear()
             $NetworkInfoDisplayGrid.rows.clear()
             
             Try {
             
             $OSinfo = Get-WmiObject -ComputerName $SystemInfoHostnameTextbox.Text -Class Win32_operatingsystem -Property * | 
             select -Property * 

             $BIOSinfo = Get-WmiObject -ComputerName $SystemInfoHostnameTextbox.Text -Class win32_bios -Property * |
             select -Property * 

             $Systeminfo = Get-WmiObject -ComputerName $SystemInfoHostnameTextbox.Text -Class win32_ComputerSystem -Property * |
             select -Property * 

             $DiskInfo = Get-WmiObject -ComputerName $SystemInfoHostnameTextbox.Text -Class win32_diskdrive -Property * | 
             select -Property *
             
             $NetworkWirelessInfo = Get-WmiObject -ComputerName $SystemInfoHostnameTextbox.Text -Class win32_NetworkAdapterConfiguration -Property * |
                                     Select * |
                                     where -Property Description -Like "*Wireless*"

             $NetworkEthernetInfo = Get-WmiObject -ComputerName $SystemInfoHostnameTextbox.Text -Class win32_NetworkAdapterConfiguration -Property * |
                                     Select * |
                                     where -Property Description -Like "*Ethernet*"


             } Catch {
                      clear
                      Write-Host "$($SystemInfoHostnameTextbox.Text) is not online!" -ForegroundColor Red
                      }


             # Rows Value input for System info DisplayGrid
             #=============================================
             $SystemInfoDisplayGrid.rows.add("$($OSinfo.Caption) Version: $($OSinfo.Version)")
             $SystemInfoDisplayGrid.rows.add("$($OSinfo.CSName)")
             $SystemInfoDisplayGrid.rows.add("$($Systeminfo.Domain)")
             $SystemInfoDisplayGrid.rows.add("$($BIOSinfo.Manufacturer)")
             $SystemInfoDisplayGrid.rows.add("$($Systeminfo.SystemFamily)")
             $SystemInfoDisplayGrid.rows.add("$($Systeminfo.ChassisSKUNumber)")
             $SystemInfoDisplayGrid.rows.add("$($Systeminfo.Model)")
             $SystemInfoDisplayGrid.rows.add("$($BIOSinfo.SerialNumber)")
             $SystemInfoDisplayGrid.rows.add("$($DiskInfo.Caption)")
             $SystemInfoDisplayGrid.rows.add("$($DiskInfo.Size)")
             $SystemInfoDisplayGrid.rows.add("$($DiskInfo.Status)")
             $SystemInfoDisplayGrid.rows.add("$($Systeminfo.Username)")
             $SystemInfoDisplayGrid.rows[0].HeaderCell.Value = "OS:"
             $SystemInfoDisplayGrid.rows[0].DefaultCellStyle.BackColor = "LightGray"
             $SystemInfoDisplayGrid.rows[1].HeaderCell.Value = "Hostname:"
             $SystemInfoDisplayGrid.rows[1].DefaultCellStyle.BackColor = "LightGray"
             $SystemInfoDisplayGrid.rows[2].HeaderCell.Value = "Domain:"
             $SystemInfoDisplayGrid.rows[2].DefaultCellStyle.BackColor = "LightGray"
             $SystemInfoDisplayGrid.rows[3].HeaderCell.Value = "Brand:"
             $SystemInfoDisplayGrid.rows[3].DefaultCellStyle.BackColor = "LightGray"
             $SystemInfoDisplayGrid.rows[4].HeaderCell.Value = "Series:"
             $SystemInfoDisplayGrid.rows[4].DefaultCellStyle.BackColor = "LightGray"
             $SystemInfoDisplayGrid.rows[5].HeaderCell.Value = "Type:"
             $SystemInfoDisplayGrid.rows[5].DefaultCellStyle.BackColor = "LightGray"
             $SystemInfoDisplayGrid.rows[6].HeaderCell.Value = "Model:"
             $SystemInfoDisplayGrid.rows[6].DefaultCellStyle.BackColor = "LightGray"
             $SystemInfoDisplayGrid.rows[7].HeaderCell.Value = "Service Tag / SN:"
             $SystemInfoDisplayGrid.rows[7].DefaultCellStyle.BackColor = "LightGray"
             $SystemInfoDisplayGrid.rows[8].HeaderCell.Value = "Hard Drive:"
             $SystemInfoDisplayGrid.rows[8].DefaultCellStyle.BackColor = "LightGray"
             $SystemInfoDisplayGrid.rows[9].HeaderCell.Value = "Storage Size:"
             $SystemInfoDisplayGrid.rows[9].DefaultCellStyle.BackColor = "LightGray"
             $SystemInfoDisplayGrid.rows[10].HeaderCell.Value = "S.M.A.R.T Status:"
             $SystemInfoDisplayGrid.rows[10].DefaultCellStyle.BackColor = "LightGray"
             $SystemInfoDisplayGrid.rows[11].HeaderCell.Value = "Logged in User:"
             $SystemInfoDisplayGrid.rows[11].DefaultCellStyle.BackColor = "LightGray"

             # Row Values input for Network info DisplayGRid
             #==============================================


             if ($NetworkWirelessInfo.IPAddress -ne $null){
             $NetworkInfoDisplayGrid.Columns[0].Name = "Network Adapter: $($NetworkWirelessInfo.Description)"
             $NetworkInfoDisplayGrid.rows.add("$($NetworkWirelessInfo.IPAddress)")
             $NetworkInfoDisplayGrid.rows[0].HeaderCell.Value = "IPV4 Address:"
             $NetworkInfoDisplayGrid.rows[0].DefaultCellStyle.BackColor = "LightGray"
             $NetworkInfoDisplayGrid.rows.add("$($NetworkWirelessInfo.IPSubnet)")
             $NetworkInfoDisplayGrid.rows[1].HeaderCell.Value = "Subnet:"
             $NetworkInfoDisplayGrid.rows[1].DefaultCellStyle.BackColor = "LightGray"
             $NetworkInfoDisplayGrid.rows.add("$($NetworkWirelessInfo.MACAddress)")
             $NetworkInfoDisplayGrid.rows[2].HeaderCell.Value = "Mac Address:"
             $NetworkInfoDisplayGrid.rows[2].DefaultCellStyle.BackColor = "LightGray"
             $NetworkInfoDisplayGrid.rows.add("$($NetworkWirelessInfo.DNSDomain)")
             $NetworkInfoDisplayGrid.rows[3].HeaderCell.Value = "DNS Domain:"
             $NetworkInfoDisplayGrid.rows[3].DefaultCellStyle.BackColor = "LightGray"
             $NetworkInfoDisplayGrid.rows.add("$($NetworkWirelessInfo.DHCPEnabled)")
             $NetworkInfoDisplayGrid.rows[4].HeaderCell.Value = "DHCP Enabled:"
             $NetworkInfoDisplayGrid.rows[4].DefaultCellStyle.BackColor = "LightGray"
             
             }
             else {
                   clear
                   Write-Host "System connected via Ethernet" -ForegroundColor green
                   }

        }

         Else {
               clear
               Write-Host "Hostname Cannont be Null!" -ForegroundColor Red
               }


             if ($NetworkEthernetInfo.IPAddress -ne $null){

             $NetworkInfoDisplayGrid.Columns[0].Name = "Network Adapter: $($NetworkEthernetInfo.Description)"
             $NetworkInfoDisplayGrid.rows.add("$($NetworkEthernetInfo.IPAddress)")
             $NetworkInfoDisplayGrid.rows[0].HeaderCell.Value = "IPV4 Address:"
             $NetworkInfoDisplayGrid.rows[0].DefaultCellStyle.BackColor = "LightGray"
             $NetworkInfoDisplayGrid.rows.add("$($NetworkEthernetInfo.IPSubnet)")
             $NetworkInfoDisplayGrid.rows[1].HeaderCell.Value = "Subnet:"
             $NetworkInfoDisplayGrid.rows[1].DefaultCellStyle.BackColor = "LightGray"
             $NetworkInfoDisplayGrid.rows.add("$($NetworkEthernetInfo.MACAddress)")
             $NetworkInfoDisplayGrid.rows[2].HeaderCell.Value = "Mac Address:"
             $NetworkInfoDisplayGrid.rows[2].DefaultCellStyle.BackColor = "LightGray"
             $NetworkInfoDisplayGrid.rows.add("$($NetworkEthernetInfo.DNSDomain)")
             $NetworkInfoDisplayGrid.rows[3].HeaderCell.Value = "DNS Domain:"
             $NetworkInfoDisplayGrid.rows[3].DefaultCellStyle.BackColor = "LightGray"
             $NetworkInfoDisplayGrid.rows.add("$($NetworkEthernetInfo.DHCPEnabled)")
             $NetworkInfoDisplayGrid.rows[4].HeaderCell.Value = "DHCP Enabled:"
             $NetworkInfoDisplayGrid.rows[4].DefaultCellStyle.BackColor = "LightGray"

             }
             Else {
                   clear
                   Write-Host "System connected via Wireless" -ForegroundColor green
                   }
        

    })


    # Sytem Information Clear Button
    #===============================================================================#
    $SystemInfoClearButton = New-Object System.Windows.Forms.Button -Property @{
    Text = ' Clear Info '
    ForeColor = "Teal"
    AutoSize = $True
    Location = New-Object System.Drawing.Point(280,315)
    }
    $SystemInfoClearButton.add_click({
                                      
                                      $SystemInfoDisplayGrid.Columns[0].Name = ""
                                      $NetworkInfoDisplayGrid.Columns[0].Name = ""
                                      $SystemInfoHostnameTextbox.Clear()
                                      $SystemInfoDisplayGrid.rows.clear()
                                      $NetworkInfoDisplayGrid.rows.clear()
                                      
                                      })

    # OS Info Subtab for System Information Tab Controller
    #===============================================================================#
    $SystemInfoSubTab = New-Object System.Windows.Forms.TabPage -Property @{
    Name = "System Info"
    Text = "System Info"
    Padding = '3, 3, 3, 3'
    Size = '200,200'
    TabIndex = 10
    Location = '10,10'
    }

    # Display Gridview for OS System Info Tab
    #===============================================================================#
    $SystemInfoDisplayGrid = New-Object System.Windows.Forms.DataGridView -Property @{
    ColumnCount = 1
    ColumnHeadersVisible = $True
    BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3d
    ScrollBars = "Vertical"
    Size = '595,267'
    ReadOnly = $True
    Location = '10,10'
    Font = $LabelFonts
    AllowUserToResizeRows = $fals;
    AllowUserToDeleteRows = $false
    AllowUserToResizeColumns = $false
    ShowCellErrors = $false
    RowHeadersVisible = $True
    RowHeadersWidthSizeMode = 2
    BackgroundColor = "Teal"
    }
    $SystemInfoDisplayGrid.Columns[0].Width = 568
    
    
    
    # Network Info Subtab for System Info Tab Controller
    #===============================================================================#
    $NetworkInfoSubTab = New-Object System.Windows.Forms.TabPage -Property @{
    Name = "Network Info"
    Text = "Network Info"
    Padding = '3, 3, 3, 3'
    Size = '200,200'
    TabIndex = 10
    Location = '10,10'
    }

    # Display Gridview for Network Info Tab
    #===============================================================================#
    $NetworkInfoDisplayGrid = New-Object System.Windows.Forms.DataGridView -Property @{
    ColumnCount = 1
    ColumnHeadersVisible = $True
    BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3d
    ScrollBars = "Vertical"
    Size = '595,267'
    ReadOnly = $True
    Location = '10,10'
    Font = $LabelFonts
    AllowUserToResizeRows = $fals;
    AllowUserToDeleteRows = $false
    AllowUserToResizeColumns = $false
    ShowCellErrors = $false
    RowHeadersVisible = $True
    RowHeadersWidthSizeMode = 2
    BackgroundColor = "Teal"
    }
    $NetworkInfoDisplayGrid.Columns[0].Width = 568

    # System Information Tab Contoller
    #===============================================================================#
    $tabcontrolNetwork = New-Object 'System.Windows.Forms.TabControl' -Property @{
    Size = '620, 300'
    AutoSize = $True
    BackColor = "Gray"
    Alignment = "Top"
    Location = New-Object System.Drawing.Point(0,350)
    }
    $tabpage3.Controls.Add($tabcontrolNetwork)

              $tabcontrolNetwork.Controls.Add($SystemInfoSubTab)
              $tabcontrolNetwork.Controls.Add($NetworkInfoSubTab)

                                 $SystemInfoSubTab.Controls.Add($SystemInfoDisplayGrid)
                                 $NetworkInfoSubTab.Controls.Add($NetworkInfoDisplayGrid)
    

    # OS Tools Main Label
    #===============================================================================#
    $OSToolsMainLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "OS / System Tools"
    ForeColor = "Teal"
    AutoSize = $True
    Font = $LabelFonts
    Location = New-Object System.Drawing.Point(10,30)
    }

    # Windows Command Line
    #===============================================================================#
    $CMD = New-Object System.Windows.Forms.Button -Property @{
    Text = ' Windows Command Line '
    ForeColor = "Teal"
    AutoSize = $True
    Location = New-Object System.Drawing.Point(30,100)
    }
    $CMD.add_click({Start cmd})

    # Remote Desktop
    #===============================================================================#
    $RDP = New-Object System.Windows.Forms.Button -Property @{
    Text = ' Remote Desktop '
    ForeColor = "Teal"
    AutoSize = $True
    Location = New-Object System.Drawing.Point(360,315)
    }
	$RDP.add_click({Start mstsc /v:"$($SystemInfoHostnameTextbox.Text)"})

    # Powershell Console
    #===============================================================================#
    $PowerShell = New-Object System.Windows.Forms.Button -Property @{
    Text = ' Poweshell '
    ForeColor = "Teal"
    AutoSize = $True
    Location = New-Object System.Drawing.Point(30,65)
    }
    $PowerShell.add_click({start powershell})

    # Restart Remote system
    #===============================================================================#
    $RestartRemoteSystem = New-Object System.Windows.Forms.Button -Property @{
    Text = ' Restart System '
    ForeColor = "Teal"
    AutoSize = $True
    Location = New-Object System.Drawing.Point(470,315)
    }
    $RestartRemoteSystem.add_click({Restart-Computer -ComputerName $SystemInfoHostnameTextbox.Text -Wait -For Wmi -Timeout 300 -Force})

    # Add Components to OS Tools Tab
    #===============================================================================#
    
    $OSToolsArray = @(
                      $SystemInfoMainLabel,
                      $PowerShell,
                      $CMD,
                      $SystemInfoHostnameLabel,
                      $SystemInfoHostnameTextbox,
                      $SystemInfoGetButton,
                      $SystemInfoClearButton,
                      $RDP,
                      $RestartRemoteSystem,
                      $OSToolsMainLabel
                      )
            foreach ($OSToolComp in $OSToolsArray)
                    {
                    $tabpage3.Controls.Add($OSToolComp)
                    }
    #
	# tabpage4 Admin Consoles
	#=============================================================================================
	$tabpage4.Location = '23, 4'
	$tabpage4.Name = 'tabpage4'
	$tabpage4.Padding = '3, 3, 3, 3'
	$tabpage4.Size = '602, 442'
	$tabpage4.TabIndex = 3
	$tabpage4.Text = ' Admin Consoles '
	$tabpage4.UseVisualStyleBackColor = $True

    # Admin Console Mail Label
    $adminConsoleMainLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Admin Consoles"
    Font = $LabelFonts
    AutoSize = $True
    ForeColor = "Teal"
    Location = New-Object System.Drawing.Point(20,20)
    }
    
    # Exchange Admin Portal Button
    $ExchangeAdminCenterButton = New-Object System.Windows.Forms.Button -Property @{
    Text = ' Exchange Admin Center '
    AutoSize = $True
    ForeColor = "Teal"
    Location = New-Object System.Drawing.Point(20,50)}
    $ExchangeAdminCenterButton.add_click({start chrome.exe https://mail.callruby.com/ecp})

    # Landesk Admin Portal Button
    $LandeskAdminButton = New-Object System.Windows.Forms.Button -Property @{
    Text = ' Landesk Admin Portal '
    AutoSize = $True
    ForeColor = "Teal"
    Location = New-Object System.Drawing.Point(170,50)}
    $LandeskAdminButton.add_click({start chrome.exe http://roxy.ruby.local:81/Default.aspx})

    # O365 Admin Portal Button
    $O365AdminButton = New-Object System.Windows.Forms.Button -Property @{
    Text = ' O365 Admin Portal '
    AutoSize = $True
    ForeColor = "Teal"
    Location = New-Object System.Drawing.Point(307,50)}
    $O365AdminButton.add_click({start chrome.exe https://portal.office.com/adminportal/})

    # Sophos Admin Console Button
    $SophosAdminButton = New-Object System.Windows.Forms.Button -Property @{
    Text = ' Sophos Admin Portal '
    AutoSize = $True
    ForeColor = "Teal"
    Location = New-Object system.drawing.point(430,50)}
    $SophosAdminButton.add_click({start chrome.exe https://central.sophos.com/manage/login})

    # 1 Password Portal Button
    $1PasswordAdminButton = New-Object System.Windows.Forms.Button -Property @{
    Text = ' 1 Password Portal '
    AutoSize = $True
    ForeColor = "Teal"
    Location = New-Object System.Drawing.Point(20,90)}
    $1PasswordAdminButton.add_click({start chrome.exe https://start.1password.com/signin?l=en})


    #=============================================================================#
    #                              Printer Consoles                               #
    #=============================================================================#

    # Print Admin Consoles Label
    $PrinterAdminConsoleLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Printer Admin Consoles"
    Font = $LabelFonts
    ForeColor = "Teal"
    AutoSize = $True
    Location = New-Object System.Drawing.Point(20,180)
    }
    
    # Beaverton Printer Admin Console Button
    $BeavertonAdminConsoleButton = New-Object System.Windows.Forms.Button -Property @{
    Text = "Beaverton Ricoh"
    AutoSize = $True
    ForeColor = "Teal"
    Location = New-Object System.Drawing.Point(20,210)
    }
    $BeavertonAdminConsoleButton.add_click({start chrome.exe https://10.2.1.20/})

    # Pearl Printer Admin Console Button
    $PearlAdminConsoleButton = New-Object System.Windows.Forms.Button -Property @{
    Text = "Pearl Ricoh"
    AutoSize = $True
    ForeColor = "Teal"
    Location = New-Object System.Drawing.Point(125,210)
    }
    $PearlAdminConsoleButton.add_click({start chrome.exe https://10.1.2.250/})

    # Fox 9 Printer Admin Console Button
    $Fox9AdminConsoleButton = New-Object System.Windows.Forms.Button -Property @{
    Text = "Fox 9 Ricoh"
    AutoSize = $True
    ForeColor = "Teal"
    Location = New-Object System.Drawing.Point(207,210)
    }
    $Fox9AdminConsoleButton.add_click({start chrome.exe https://10.3.1.20/})

    # Fox 6 Printer Admin Console Button
    $Fox6AdminConsoleButton = New-Object System.Windows.Forms.Button -Property @{
    Text = "Fox 6 Ricoh"
    AutoSize = $True
    ForeColor = "Teal"
    Location = New-Object System.Drawing.Point(289,210)
    }
    $Fox6AdminConsoleButton.add_click({start chrome.exe https://10.3.2.9/})

    # Kansas City Printer Admin Console Button
    $KCAdminConsoleButton = New-Object System.Windows.Forms.Button -Property @{
    Text = "KC Ricoh"
    AutoSize = $True
    ForeColor = "Teal"
    Location = '370,210'
    }
    $KCAdminConsoleButton.add_click({start chrome.exe https://10.3.2.9/})


    #=============================================================================#
    #                        Superscript Logs Viewer                              #
    #=============================================================================#

    # Superscript Log Explorer Section Label
    $ssLogViewerLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "SuperScript Log Viewer"
    Font = $LabelFonts
    ForeColor = "Teal"
    AutoSize = [boolean]$true
    Location = '20,330'
    }

    # Log Explorer Function
    Function SSLogExplorer {

    [void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.ServiceProcess, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $LogExicon = New-Object System.Drawing.Icon ("$PSScriptRoot\Source\ros_app_icon.ico")

    # Main Form
    #===============================
    $LogExmainForm = New-Object System.Windows.Forms.Form -Property @{
    Size = '1600,600'
    Icon = $LogExicon
    Text = [String]"SuperScript Log Explorer"
    }
    
    
    # Query Log Files in Log Archive
    #===============================
    $QueryLogs = gci -Path "\\prlfs01\iswizards\superscript\Logs\Log-Archive" -Recurse |
    Where {$_.Extension -eq ".txt"} | 
    Select -Property Directory, Name, CreationTime |
    Sort-Object -Property CreationTime -Descending

    
    # Listbox .net object
    #============================== 
    $LogListBox = New-Object System.Windows.Forms.ListBox -Property @{
    Location = '16,15'
    Size = '450,500'
    }
    $LogCount = [Int32]0
    foreach ($Log in $QueryLogs) {
                                  $LogCount++
                                  [void] $LogListBox.Items.Add("$($Log.Directory)\$($Log.Name)")
                                  }

    
    # Log Viewer Textbox .net object
    #==============================
    $LogViewTB = New-Object System.Windows.Forms.TextBox -Property @{
    ReadOnly = [boolean]$true
    BackColor = 'black'
    Location = '480,15'
    Size = '1092,500'
    Multiline = [boolean]$true
    ForeColor = 'lightgreen'
    Wordwrap = [boolean]$true
    Scrollbars = 'Vertical'
    }
    
    
    # Open Log .net Button
    #==============================
    $OpenLogButton = New-Object System.Windows.Forms.Button -Property @{
    Text = "Open Log"
    Autosize = [boolean]$true
    Location = '310,520'
    }
    $OpenLogButton.add_click({
                              $LogViewTB.Clear()

                              $SelectToString = "$($LogListBox.SelectedItem)"

                              foreach($line in Get-Content -Path "$($SelectToString)"){
                              
                                     if($line -match $regex){$LogViewTB.AppendText("`r`n$($line)")}
                              
                              
                              }

                              

                              
                              })


    # Remove Log .net Button
    #==============================
    $RemoveLogButton = New-Object System.Windows.Forms.Button -Property @{
    Text = "Remove Log"
    Autosize = [boolean]$true
    Location = '388,520'
    }
    $RemoveLogButton.add_click({
                                Remove-Item -Path $LogListBox.SelectedItem -Verbose -Force
                                $LogListBox.Items.Clear()

                                $reQueryLogs = gci -Path "\\prlfs01\iswizards\superscript\Logs\Log-Archive" -Recurse |
                                               Where {$_.Extension -eq ".txt"} | 
                                               Select -Property Directory, Name, CreationTime |
                                               Sort-Object -Property CreationTime -Descending

                                               foreach ($ReQ in $reQueryLogs) {
                                                                              [void] $LogListBox.Items.Add("$($ReQ.Directory)\$($ReQ.Name)")
                                                                              }

                                $LogRemovNotif = New-Object -ComObject Wscript.Shell -ErrorAction Stop
                                $LogRemovNotif.Popup("$($ReQ.Directory)\$($ReQ.Name) was removed!",0,"Removal Success!")
                                $LogViewTB.Clear()
                                })

    $FormAssemblyBuild = @(
                           $LogListBox,
                           $LogViewTB,
                           $OpenLogButton,
                           $RemoveLogButton
                           )
                           foreach ($buildItem in $FormAssemblyBuild) {
                           
                                                                      $LogExmainForm.Controls.Add($buildItem)
                           
                                                                      }
    
    
    $LogExmainForm.ShowDialog()

}

    
    # Superscript Log Explorer Button
    $ssLogViewerButton = New-Object System.Windows.Forms.Button -Property @{
    Text = [System.String]"View Logs"
    AutoSize = [boolean]$True
    ForeColor = "Teal"
    Location = '20,360'
    }
    $ssLogViewerButton.add_click({SSLogExplorer})


    # Add Components to Admin Consoles Tab
    #======================================
    
    $adminconsolearray = @(
                           $adminConsoleMainLabel,
                           $PrinterAdminConsoleLabel,
                           $ExchangeAdminCenterButton,
                           $SophosAdminButton,
                           $O365AdminButton,
                           $1PasswordAdminButton,
                           $LandeskAdminButton,
                           $BeavertonAdminConsoleButton,
                           $PearlAdminConsoleButton,
                           $Fox9AdminConsoleButton,
                           $Fox6AdminConsoleButton,
                           $KCAdminConsoleButton,
                           $ssLogViewerLabel,
                           $ssLogViewerButton
                           )

        foreach ($consolelink in $adminconsolearray)
                 {$tabpage4.Controls.Add($consolelink)}
	



    #==========================================================================#
	#                              tabpage5 About                              #
	#==========================================================================#
	$tabpage5.Location = '42, 4'
	$tabpage5.Name = 'tabpage5'
	$tabpage5.Padding = '3, 3, 3, 3'
	$tabpage5.Size = '583, 442'
	$tabpage5.TabIndex = 7
	$tabpage5.Text = 'About'
	$tabpage5.UseVisualStyleBackColor = $True


    # About: Applicaton Info
    #=====================================================================


    # Contact us Link
    #=====================================================================
    $AboutInfoMailLink = New-Object System.Windows.Forms.LinkLabel -Property @{
    ActiveLinkColor = [System.Drawing.Color]::DarkViolet
    Name = "Contact Developer"
    Text = "Contact Us"
    AutoSize = $True
    LinkColor = [System.Drawing.Color]::Teal
    BackColor = "LightGray"
    Location = '370,229'
    Font = $LabelFonts
    }
    $AboutInfoMailLink.add_click({[system.Diagnostics.Process]::start("https://help.ruby.com/")})

    # Main About Textbox
    #=====================================================================
    $AboutInfoTextbox = New-Object System.Windows.Forms.RichTextBox -Property @{
    BackColor = "LightGray"
    ReadOnly = $True
    Multiline = $True
    Size = '500,500'
    Font = $LabelFonts 
    Location = '55,70'
    }
    $AboutInfoTextbox.AppendText("`r`n
    Superscript GUI Version 2.0
    ====================
    `r`n
    Written by: Chris Meeuwen
    `r`n
    Company: Ruby Receptionists
    `r`n
    For script / application update suggestions:
    "
    )
    
	

    # Add Controls to about tab
    #=====================================================================
    $AboutCompsArray = @(
                         $AboutInfoTextbox,
                         $AboutInfoMailLink
                         )
                         foreach ($aboutComps in $AboutCompsArray) {
                                                                    $tabpage5.Controls.Add($aboutComps)
                                                                    }
    
    
    $tabpage5.Controls.Add($AboutInfoTextbox)
    
    
    #================================================================#
	#                    tab page6 HotFixes                          #
	#================================================================#
	$tabpage6.Location = '42, 4'
	$tabpage6.Name = 'tabpage6'
	$tabpage6.Padding = '3, 3, 3, 3'
	$tabpage6.Size = '583, 442'
	$tabpage6.TabIndex = 6
	$tabpage6.Text = 'Hotfixes'
	$tabpage6.UseVisualStyleBackColor = $True


    #================================================================#
    # GTA Unattended Mouse Cursor Fix                                #
    #================================================================#
    
    
    # GTA Unattend Mouse Cursor Fix Label
    #================================================================#	
    $GTAUMCFL = New-Object System.Windows.Forms.Label -Property @{
    Text = "GTA Unttended Mouse Cursor Fix"
    Font = $LabelFonts
    ForeColor = "Teal"
    AutoSize = $True
    Location = '20,20'
    }


    # GTA Unattend Mouse Cursor Fix Hostname Label
    #================================================================#
    $GTAUMCFHL = New-Object System.Windows.Forms.Label -Property @{
    Text = "Hostname:"
    AutoSize = $True
    Location = '20,50'
    }


    # GTA Unattend Mouse Cursor Fix Hostname Textbox
    #================================================================#
    $GTAUMCFTB = New-Object System.Windows.Forms.TextBox -Property @{
    BackColor = "LightGray"
    Size = '100,30'
    Location = '80,47'
    }



    # GTA Unattend Mouse Cursor Fix Function
    #================================================================#
    Function BlastGTAUnattend {

clear



Invoke-Command -ComputerName $GTAUMCFTB.Text -ScriptBlock {

                                                    Get-Process | where -Property Name -Like "LogmeinRemoteSupport*" | Stop-Process -Force -Verbose
                                                    Get-Process | where -Property Name -Like "GoToAssist*" | Stop-Process -Force -Verbose

                                                   }



$getService = Get-Service -ComputerName $GTAUMCFTB.Text | where -Property name -Like "LMIR*"

if ($getService -ne $null) { 

cmd.exe /c sc \\"$($GTAUMCFTB.Text)" delete "$($getService.Name)" /w

}

Else {

      Write-Host "Service Not Running! Starting Removal." -ForegroundColor Magenta

     }

$DIRPATHS = @(
              "\\$($GTAUMCFTB.Text)\c$\Program Files (x86)\LogMeIn Remote Support Unattended",
              "\\$($GTAUMCFTB.Text)\c$\Program Files (x86)\LogMeIn RescueAssist Unattended",
              "\\$($GTAUMCFTB.Text)\c$\Program Files (x86)\GoToAssist Remote Support Unattended"
              )


Foreach ($dir in $DIRPATHS) {
                                           
                             $DoesFolderExist = Test-Path -Path $dir
                                                
                             if ($DoesFolderExist -ne $false) {
                                                                                  
                                                               Remove-Item -Path $dir -Recurse -Force -Verbose
                                                
                                                              }

                                                              Else {
                                                                                 
                                                                    Write-Host "$($dir) does not exist on $($GTAUMCFTB.Text)" -ForegroundColor Red
                                                                                 
                                                                   }
                                                
                                                
                                              
              
                             }



Write-Host "Unattended Support Obliterated!" -ForegroundColor Green


}


    # GTA Unattend Mouse Cursor Fix Button
    #================================================================#
    $GTAUMCFBT = New-Object System.Windows.Forms.Button -Property @{
    Text = "Run"
    Autosize = $True
    Location = '185,45'
    }
    $GTAUMCFBT.add_click({BlastGTAUnattend})



    #================================================================#
    # Public Folder SMTP array change - Finish Me!                   #
    #================================================================#

    

    # Public Folder STMP Address Updater Label
    #================================================================#
    $PFSMTPChangeLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Public Folder SMTP Proxy Address change Utility"
    AutoSize = $True
    Font = $LabelFonts
    Forecolor = "Teal"
    Location = '20,100'
    }

    
    # Old Public Folder SMTP Proxy Label
    #================================================================# 
    $PFSMTPCPATBLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Public Folder Proxy Address lookup:"
    Autosize = $True
    Location = '20,130'
    }
    
    
    # Old Public Folder SMTP Proxy Address Textbox
    #================================================================#
    $PFSMTPCPATB = New-Object System.Windows.Forms.TextBox -Property @{
    BackColor = "lightgray"
    Size = '150,30'
    Location = '210,127'
    }
   

    # Verify Current SMTP Proxy addresses for current Public Folder Button
    #================================================================#
    $PFSMTPCPABTVerify = New-Object System.Windows.Forms.Button -Property @{
    Text = "Verify"
    Autosize = $True
    Location = '365,125'
    }
    $PFSMTPCPABTVerify.add_click({
    
    if ($PFSMTPCPATB.Text -ne '') {
    
    Clear

    $PFSMTPCPAListbox.Items.Clear()


    $FindPublicMailbox = Get-ADObject -Filter * -Properties proxyaddresses | 
                         Where {$_.proxyaddresses -eq "smtp:$($PFSMTPCPATB.Text)"} | 
                         Select proxyaddresses -ExpandProperty proxyaddresses

    $FindPublicMailboxGUID = Get-ADObject -Filter * -Properties ObjectGUID, proxyaddresses, ObjectClass, Name |
                             Where {$_.proxyAddresses -eq "smtp:$($PFSMTPCPATB.Text)"} |
                             Select ObjectGUID, proxyaddresses, ObjectClass, Name

   

    Foreach ($smtpAddy in $FindPublicMailbox) {$PFSMTPCPAListbox.Items.Add($smtpAddy)}
    
    
    $PFSMTPCPAListboxLabel.Text = " SMTP Array for: $($FindPublicMailboxGUID.Name)"
    
    
    
    
    
    } Else {
            
            $NoProxyInputError = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
            $NoProxyInputError.Popup("No SMTP address provided!",0,"Invalid Input - Verify SMTP proxy address") 
    
           
           }
    
    })


    
    # Clear Public folder SMTP proxy address lookup
    #================================================================#
    $PFSMTPCPABTClear = New-Object System.Windows.Forms.Button -Property @{
    Text = "Clear"
    Autosize = $True
    Location = '445,125'
    }
    $PFSMTPCPABTClear.add_click({
    
    $PFSMTPCPAListbox.Items.Clear()
    
    $PFSMTPCPAListboxLabel.Text = " SMTP Array for: "
    
    $PFSMTPCPATB.Clear()
    
    })


    
    # Current SMTP Proxy addresses Listbox Label
    #================================================================#
    $PFSMTPCPAListboxLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = " SMTP Array for: "
    Font = $LabelFonts
    Forecolor = "Teal"
    Size = '200,16'
    Location = '20,174'
    }



    # Current SMTP Proxy addresses Listbox
    #================================================================#
    $PFSMTPCPAListbox = New-Object System.Windows.Forms.ListBox -Property @{
    size = '200,100'
    Location = '20,190'
    Font = $LabelFonts
    BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3d
    }
    


    # New Proxy Address to replace in array label
    #================================================================#
    $PFSMTPNewSTMPLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = " New SMTP Proxy Address "
    Size = '170,17'
    Font = $LabelFonts
    Forecolor = "Teal"
    Location = '230,174'
    }


    
    # New Proxy Address to replace in array Textbox
    #================================================================#
    $PFSMTPNewSTMPTB = New-Object System.Windows.Forms.TextBox -Property @{
    Backcolor = "LightGray"
    Size = '170,30'
    Location = '230,191'
    BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3d
    }



    # Replace Proxy address in array Button
    #================================================================#
    $PFSMTPReplaceBT = New-Object System.Windows.Forms.Button -Property @{
    Text = " Replace SMTP Proxy Address "
    Autosize = $True
    Location = '230,222'
    }
    $PFSMTPReplaceBT.add_click({

    if ($PFSMTPNewSTMPTB.Text -ne '') {

    $FindPublicMailboxGUID = Get-ADObject -Filter * -Properties ObjectGUID, proxyaddresses, ObjectClass, Name |
                             Where {$_.proxyAddresses -eq "smtp:$($PFSMTPCPATB.Text)"} |
                             Select ObjectGUID, proxyaddresses, ObjectClass, Name
    
    Write-Host "Removing proxyAddress from property array: $($PFSMTPCPAListbox.SelectedItem)" -ForegroundColor Yellow 
    Set-ADObject $FindPublicMailboxGUID.ObjectGUID -Remove @{proxyAddresses="$($PFSMTPCPAListbox.SelectedItem)"} -Confirm:$false -Verbose
    

    Write-Host "Adding proxyAddress from property array: smtp:$($PFSMTPNewSTMPTB.Text)" -ForegroundColor Yellow 
    Set-ADObject $FindPublicMailboxGUID.ObjectGUID -Add @{proxyAddresses="smtp:$($PFSMTPNewSTMPTB.Text)"} -Confirm:$false -Verbose
    

    clear

    $ProxyReplaceSuccess = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
    $ProxyReplaceSuccess.Popup("SMTP address successfully replaced!",0,"Success - Replace SMTP proxy address") 

    
    } Else {
            
            $NoProxyReplaceError = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
            $NoProxyReplaceError.Popup("No SMTP address provided!",0,"Invalid Input - Replace SMTP proxy address") 
    
    
           }


    })



    
    # Add Controls to Hotfix tab
    #================================================================#
    $HotFixArray = @(
                    $GTAUMCFL,
                    $GTAUMCFHL,
                    $GTAUMCFTB,
                    $GTAUMCFBT,
                    $PFSMTPChangeLabel,
                    $PFSMTPCPATBLabel,
                    $PFSMTPCPATB,
                    $PFSMTPCPABTVerify,
                    $PFSMTPCPABTClear,
                    $PFSMTPCPAListboxLabel,
                    $PFSMTPCPAListbox,
                    $PFSMTPNewSTMPLabel,
                    $PFSMTPNewSTMPTB,
                    $PFSMTPReplaceBT
                    )
                    foreach ($HFComponent in $HotFixArray){
                                                           $tabpage6.Controls.Add($HFComponent)
                                                          }



	$tabcontrol1.ResumeLayout()
	$form1.ResumeLayout()
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $form1.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$form1.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$form1.add_FormClosed($Form_Cleanup_FormClosed)
	#Show the Form
	return $form1.ShowDialog()

} #End Function

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

clear

#Call the form
Show-tabcontrol_psf | Out-Null

Stop-Transcript | Out-Null
[System.GC]::Collect()
$error.Clear()
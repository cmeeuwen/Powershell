Function CreateUserTemplate {

# .Net Assembly Load
#==========================================================================#
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Main Form
#==========================================================================#
$MainForm = New-Object System.Windows.Forms.Form -Property @{
size = '945,620'
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


# OU Dropdown Label
#==========================================================================#
$SelectOULabel = New-Object System.Windows.Forms.Label -Property @{
Location = '12,223'
Text = "Select Orginizational Unit:"
AutoSize = $True
Font = $LabelFont 
}


# Form OU Dropdown Selection
#==========================================================================# 
$OUSelectionDropdown = New-Object System.Windows.Forms.ComboBox -Property @{
AllowDrop = $True
Size = '240,50'   
Location = '10,243'
}

$OUSelectionDropdown.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend;
$OUSelectionDropdown.AutoCompleteCustomSource.AddRange(('OU=Billing,OU=Ruby Users,DC=ruby,DC=local',
           'OU=BVN Admin,OU=Ruby Users,DC=ruby,DC=local',
           'OU=BVN HR,OU=Ruby Users,DC=ruby,DC=local',
           'OU=BVN Information Technology,OU=Ruby Users,DC=ruby,DC=local',
           'OU=BVN Office,OU=Ruby Users,DC=ruby,DC=local',
           'OU=BVN RS,OU=Ruby Users,DC=ruby,DC=local',
           'OU=BVN Sales,OU=Ruby Users,DC=ruby,DC=local',
           'OU=FXT Scheduling Ninjas,OU=Ruby Users,DC=ruby,DC=local',
           'OU=Contractors,OU=Ruby Users,DC=ruby,DC=local',
           'OU=Marketing,OU=Ruby Users,DC=ruby,DC=local',
           'OU=Misc Hourly Rubys,OU=Ruby Users,DC=ruby,DC=local',
           'OU=PDX Administrative,OU=Ruby Users,DC=ruby,DC=local',
           'OU=PDX Client Happiness,OU=Ruby Users,DC=ruby,DC=local',
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
           'OU=BVN RS,OU=Ruby Users,DC=ruby,DC=local',
           'OU=BVN Sales,OU=Ruby Users,DC=ruby,DC=local',
           'OU=FXT Scheduling Ninjas,OU=Ruby Users,DC=ruby,DC=local',
           'OU=Contractors,OU=Ruby Users,DC=ruby,DC=local',
           'OU=Marketing,OU=Ruby Users,DC=ruby,DC=local',
           'OU=Misc Hourly Rubys,OU=Ruby Users,DC=ruby,DC=local',
           'OU=PDX Administrative,OU=Ruby Users,DC=ruby,DC=local',
           'OU=PDX Client Happiness,OU=Ruby Users,DC=ruby,DC=local',
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
Location = '12,271'
Text = "Select Groups:"
AutoSize = $True
Font = $LabelFont 
}

# AD Group Listbox
#==========================================================================#
$SelectGroupsListbox = New-Object System.Windows.Forms.Listbox -Property @{
Location = '10,289'
Size =  '240,230'
}
$SelectGroupsListbox.SelectionMode = 'MultiSimple'

$getADgroups = Get-ADGroup -SearchBase 'OU=Ruby Groups,DC=ruby,DC=local' -Filter 'ObjectClass -eq "group"' |
               select Name | Sort-Object -Property Name

               foreach ($SecGroup in $getADgroups){[void] $SelectGroupsListbox.Items.Add($SecGroup.name)}
                

# Form Template display grid preview
#==========================================================================# 
$TemplatePreviewGrid = New-Object System.Windows.Forms.DataGridView -Property @{
ColumnCount = 3
ColumnHeadersVisible = $True
Size = '620,350'
Location = '290,165'
ScrollBars = "Vertical"
ReadOnly = $True
BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3d
BackgroundColor = "Teal"
}
$TemplatePreviewGrid.Columns[0].Name = "Title"
$TemplatePreviewGrid.Columns[0].Width = 160
$TemplatePreviewGrid.Columns[1].Name = "OU"
$TemplatePreviewGrid.Columns[1].Width = 260
$TemplatePreviewGrid.Columns[2].Name = "Groups"
$TemplatePreviewGrid.Columns[2].Width = 160


# Form button for template creation
#==========================================================================# 
$GenerateTemplateButton = New-Object System.Windows.Forms.Button -Property @{
Text = "Generate Template"
Location = '780,530'
AutoSize = $true
BackColor = 'white'
}
$GenerateTemplateButton.add_click({

$doesTemplateExist = Test-Path -Path "$($PSScriptRoot)\$($templateNameTextbox.Text).csv"

if ($doesTemplateExist -eq $false) {
 
$convertGridtoTable = New-Object System.Data.DataTable -Property @{
TableName = "ConvertTable"
}

        foreach ($columns in $TemplatePreviewGrid.Columns) {$convertGridtoTable.Columns.Add($columns.Name)}

        foreach ($rows in $SelectGroupsListbox.SelectedItems) {$convertGridtoTable.Rows.Add($null,$null,$rows)}

        $convertGridtoTable.Rows.Add($jobTitle.Text,$OUSelectionDropdown.SelectedItem,$null)

        $TemplatePreviewGrid.DataSource = $convertGridtoTable

        $convertGridtoTable | Export-Csv $PSScriptRoot\scratch.csv -Verbose

        $Removefirstrow = get-content $PSScriptRoot\scratch.csv
        $Removefirstrow | Select-Object -Skip 1 | Set-Content "$($PSScriptRoot)\$($templateNameTextbox.Text).csv" -Verbose

        Remove-Item $PSScriptRoot\scratch.csv -Force -Verbose
        
        $MainForm.Hide()

        $templateGenerateMessage = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
        $templateGenerateMessage.Popup("New Onboarding Template Generated!",0,"Success!") 
         
        $MainForm.Close() }
        
        Else {
              $templateFailGenerateMessage = New-Object -ComObject Wscript.Shell -ErrorAction Stop
          
              $templateFailGenerateMessage.Popup("$($templateNameTextbox.Text) already exists!`nPlease choose a new template name!",0,"Template Generate failure!") 
              
              Remove-Item $PSScriptRoot\scratch.csv -Force -Verbose
             }        
})


# Form button for template review
#==========================================================================# 
$ReviewTemplateButton = New-Object System.Windows.Forms.Button -Property @{
Text = "Review Template"
Location = '670,530'
AutoSize = $true
BackColor = 'white'
}
$ReviewTemplateButton.add_click({
                                 
                                 $TemplatePreviewGrid.rows.Clear()

                                 foreach ($selectedgroup in $SelectGroupsListbox.SelectedItems)
                                         {$TemplatePreviewGrid.Rows.Add($null,$null,$selectedgroup)}
                                 
                                 $TemplatePreviewGrid.Rows[0].Cells[0].Value = $jobTitle.Text
                                 $TemplatePreviewGrid.Rows[0].Cells[1].Value = $OUSelectionDropdown.SelectedItem
                               
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

CreateUserTemplate | Out-Null

[System.GC]::Collect()

$Error.Clear()
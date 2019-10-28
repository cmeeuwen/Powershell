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
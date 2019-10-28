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

$TimePickerLabel = New-Object System.Windows.Forms.Label
$TimePickerLabel.Text = “Time”
$TimePickerLabel.Font = "Georgia"
$TimePickerLabel.ForeColor = "Teal"
$TimePickerLabel.Location = “360, 90”
$TimePickerLabel.Height = 22
$TimePickerLabel.Width = 40
$form.Controls.Add($TimePickerLabel)

$TimePicker = New-Object System.Windows.Forms.DateTimePicker 
$TimePicker.Location = “360, 120”
$TimePicker.Width = “150”
$TimePicker.Format = [windows.forms.datetimepickerFormat]::custom
$TimePicker.CustomFormat = “HH:mm:ss”
$TimePicker.ShowUpDown = $TRUE
$form.Controls.Add($TimePicker)

$result = $form.ShowDialog()

if ($result -eq [Windows.Forms.DialogResult]::OK) {
    $date = $calendar.SelectionStart
    Write-Host "Date selected: $($date.ToShortDateString())"
}
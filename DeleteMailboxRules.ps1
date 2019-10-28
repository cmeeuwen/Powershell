Function DeletMailboxRule {

$ErrorActionPreference = "SilentlyContinue"

$userCred = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $userCred -Authentication Basic -AllowRedirection
Import-PSSession $Session

clear

$whatUser = Read-Host -Prompt "Provide users UPN ( EX:firstname.lastname@callruby.com ) "

$getrules = Get-InboxRule -Mailbox $whatUser | select -Property name


                Add-Type -AssemblyName System.Windows.Forms
                Add-Type -AssemblyName System.Drawing

                
                $form = New-Object System.Windows.Forms.Form
                $form.Text = 'Mailbox Rules for $whatUser'
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
                $label.Text = 'Please Select rules to remove:'
                $form.Controls.Add($label)

                $listBox = New-Object System.Windows.Forms.Listbox
                $listBox.Location = New-Object System.Drawing.Point(15,60)
                $listBox.Size = New-Object System.Drawing.Size(350,500)
                

                $listBox.SelectionMode = 'MultiSimple'
                
                $getrules = Get-InboxRule -Mailbox $whatUser | select -Property name

                foreach ($rule in $getrules) {
                                              [void] $listBox.Items.Add($rule.name)
                                              }
                

                $listBox.Height = 150
                $form.Controls.Add($listBox)
                $form.Topmost = $true

                $result = $form.ShowDialog()

                     if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                                                                               Write-Host "`nRules successfully queried!" -ForegroundColor Green
                                                                               } 
                        
                        else {
                              Write-Host "`nCancelled! No rules removed!" -ForegroundColor Red
                              }
                     
                     ForEach-Object {
    
                                        foreach ($rule in $listBox.SelectedItems)
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
DeletMailboxRule
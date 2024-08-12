[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
import-module activedirectory
$button_click=
{
if($objListBox.SelectedItem -eq "View All users")
{
$data=get-aduser -filter *
foreach($line in $data)
{
$rich.Appendtext(($line | Out-String))
}
}
if($objListBox.SelectedItem -eq "View All FSMO roles")
{
$data=netdom query fsmo 
foreach($line in $data)
{
$rich.Appendtext(($line | Out-String ))
}
}
if($objListBox.SelectedItem -eq "Check AD health")
{
$data=dcdiag /test:dcpromo /dnsdomain:soft.com
foreach($line in $data)
{
$rich.Appendtext(($line | Out-String))
}
}
if($objListBox.SelectedItem -eq "Check DNS")
{
$data=dcdiag /dnsall
foreach($line in $data)
{
$rich.Appendtext(($line | Out-String))
}
}
if($objListBox.SelectedItem -eq "View All Services")
{
$data=get-service 
foreach($line in $data)
{
$rich.Appendtext(($line | Out-String))
}
}
if($objListBox.SelectedItem -eq "View All processes")
{
$data=ps
foreach($line in $data)
{
$rich.Appendtext(($line | Out-String))
}
}

}

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "View Status"
$objForm.Size = New-Object System.Drawing.Size(300,200) 
$objForm.StartPosition = "CenterScreen"

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(75,120)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.Add_Click($button_click)
$objForm.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(150,120)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Clear"
$CancelButton.Add_Click({$rich.clear()})
$objForm.Controls.Add($CancelButton)

$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,20) 
$objLabel.Size = New-Object System.Drawing.Size(280,20) 
$objLabel.Text = "Please select an option:"
$objForm.Controls.Add($objLabel) 

$objListBox = New-Object System.Windows.Forms.ListBox 
$objListBox.Location = New-Object System.Drawing.Size(10,40) 
$objListBox.Size = New-Object System.Drawing.Size(260,20) 
$objListBox.Height = 80

[void] $objListBox.Items.Add("View All users")
[void] $objListBox.Items.Add("View All FSMO roles")
[void] $objListBox.Items.Add("Check AD health")
[void] $objListBox.Items.Add("Check DNS")
[void] $objListBox.Items.Add("View all Services")
[void] $objListBox.Items.Add("View all processes")

$objForm.Controls.Add($objListBox) 

$rich=New-Object System.Windows.Forms.RichTextBox
$rich.location=New-Object System.Drawing.Size(10,150)
$rich.size=New-Object System.Drawing.Size(400,350)
$rich.tabindex=2
$rich.font="Arial"
$objForm.Controls.Add($rich)

[void] $objForm.ShowDialog()

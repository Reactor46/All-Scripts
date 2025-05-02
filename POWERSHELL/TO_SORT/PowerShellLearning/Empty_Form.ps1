#Adding the windows forms assembly to make use of windows forms objects
Add-Type -AssemblyName System.Windows.Forms

#Building the Windows Forms object, beginning with the main one and sizing it appropriately
$form = New-Object System.Windows.Forms.Form
$form.Text = "My form"
$form.Width = 500
$form.Height = 500

#Show the main form
$form.ShowDialog()
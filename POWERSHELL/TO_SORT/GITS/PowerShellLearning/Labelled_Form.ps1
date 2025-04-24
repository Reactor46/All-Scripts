#Importing the windows form assembly to make use of form objects
Add-Type -AssemblyName System.Windows.Forms

#Creating the form, storing it as a variable and specifying its dimensions
$form = New-Object System.Windows.Forms.Form
$form.Text = "My form"
$form.Width = 500
$form.Height = 500

#Creating a label to display 'This is a label'
$label = New-Object System.Windows.Forms.Label
$label.Text = 'This is a label'
$label.AutoSize = $true
$label.Location = New-Object System.Drawing.Point (10,10)

#Add the label to the form and then open the form
$form.Controls.Add($label)
$form.ShowDialog()



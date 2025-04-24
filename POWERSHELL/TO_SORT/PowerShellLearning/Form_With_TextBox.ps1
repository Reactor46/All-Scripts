#Importing the assembly of objects required for the forum
Add-Type -AssemblyName System.Windows.Forms

#Creating the form 
$form = New-Object System.Windows.Forms.Form
$form.Text = 'My Form'
$form.Width = 500
$form.Height = 500

#Creating the label
$label = New-Object System.Windows.Forms.Label
$label.Text = 'Full Name'
$label.AutoSize = $true
$label.Location = New-Object System.Drawing.Point(10,10)

#Creating the text box
$TextBox1 = New-Object System.Windows.Forms.TextBox
$TextBox1.Multiline = $false
$TextBox1.Width = 116
$TextBox1.Height = 20
$TextBox1.Location = New-Object System.Drawing.Point(90, 10)
$TextBox1.Font = New-Object System.Drawing.Font('Microsoft Sans Serif',12)

#Adding the label and text box to the script and opening up the dialogue box
$form.Controls.AddRange(@($label,$TextBox1))
$form.ShowDialog()
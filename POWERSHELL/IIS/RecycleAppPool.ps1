<# 
.NAME
    RecycleAppPool
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(400,400)
$Form.text                       = "Form"
$Form.TopMost                    = $false



#region Logic 
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


$form = New-Object System.Windows.Forms.Form
$form.Text = "Recycle App Pool"
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = "CenterScreen"

$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(10,20)
$label1.Size = New-Object System.Drawing.Size(280,20)
$label1.Text = "Remote Server:"

$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox1.Location = New-Object System.Drawing.Point(10,40)
$textBox1.Size = New-Object System.Drawing.Size(260,20)

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(10,70)
$label2.Size = New-Object System.Drawing.Size(280,20)
$label2.Text = "App Pool Name:"

$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(10,90)
$textBox2.Size = New-Object System.Drawing.Size(260,20)

$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(10,120)
$button.Size = New-Object System.Drawing.Size(260,20)
$button.Text = "Recycle App Pool"

$form.Controls.Add($label1)
$form.Controls.Add($textBox1)
$form.Controls.Add($label2)
$form.Controls.Add($textBox2)
$form.Controls.Add($button)

$form.Add_Shown({$form.Activate()})


$button.Add_Click({
    $remoteServer = $textBox1.Text
    $appPoolName = $textBox2.Text
    $scriptBlock = {
        param($appPoolName)
        Import-Module WebAdministration
        Restart-WebAppPool -Name $appPoolName
    }

    # Using Invoke-Command for the action
    try {
        Invoke-Command -ComputerName $remoteServer -ScriptBlock $scriptBlock -ArgumentList $appPoolName
        [System.Windows.Forms.MessageBox]::Show("App Pool Recycled Successfully", "Success")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to Recycle App Pool: $_", "Error")
    }
})

$form.ShowDialog()

#endregion

$env:TITLE = " FINDVM GUI v1.2"

Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

function Hide-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 0)
}

Hide-Console

function accessVC{
Param($credTIT,$credText,$credUSR,$credPSW)
$Host.ui.PromptForCredential("$credTIT","$credText","$credUSR","$credPSW")
}

function VCConnection{
Param($VC)
Connect-VIServer $VC -Credential $cred -ErrorAction SilentlyContinue
}

do{
$testVCconn = "INSERT THE FIRST VIRTUAL CENTER"
$cred = accessVC -credTIT "NEED CREDENTIAL" -credText "Please type your userid" -credUSR "INSERT YOU DOMAIN CREDENTIAL" -credPSW "NetBiosUserName"
$conn = VCConnection "$testVCconn"

if ($conn.Name -ne $testVCconn){
$wshell = New-Object -ComObject Wscript.Shell
$ans = $wshell.Popup("Cannot complete login due to an incorrect user name or password. 
Do you want to retry?",0," CONNECTION ERROR",4)
if ($ans -eq "7"){exit}}

} Until ($conn.Name -eq $testVCconn)

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region begin GUI{ 

$FINDVMGUI                       = New-Object system.Windows.Forms.Form
$FINDVMGUI.ClientSize            = '520,350'
$FINDVMGUI.MaximizeBox           = $false
$FINDVMGUI.FormBorderStyle       = "Fixed3d"
$FINDVMGUI.text                  = $env:TITLE
$FINDVMGUI.TopMost               = $false

$Groupbox1                       = New-Object system.Windows.Forms.Groupbox
$Groupbox1.height                = 60
$Groupbox1.width                 = 190
$Groupbox1.text                  = "CRITICAL VM OPERATION"
$Groupbox1.location              = New-Object System.Drawing.Point(10,277)

$Groupbox2                       = New-Object system.Windows.Forms.Groupbox
$Groupbox2.height                = 60
$Groupbox2.width                 = 170
$Groupbox2.text                  = "TEST CONNECTION"
$Groupbox2.location              = New-Object System.Drawing.Point(210,277)

$VMNAME                          = New-Object system.Windows.Forms.TextBox
$VMNAME.multiline                = $false
$VMNAME.text                     = "VM"
$VMNAME.width                    = 259
$VMNAME.height                   = 20
$VMNAME.location                 = New-Object System.Drawing.Point(10,32)
$VMNAME.Font                     = 'Microsoft Sans Serif,10'
$FINDVMGUI.Controls.Add($VMNAME)

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "TYPE YOUR VM"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(10,14)
$Label1.Font                     = 'Microsoft Sans Serif,10'

$Label2                          = New-Object System.Windows.Forms.LinkLabel
$Label2.Location                 = New-Object System.Drawing.Point(401,327)
$Label2.AutoSize                 = $true
$Label2.LinkColor                = "#0074A2"
$Label2.ActiveLinkColor          = "#114C7F"
$Label2.Text                     = "-Andrea Springolo-"

$outputBox                        = New-Object System.Windows.Forms.richTextBox 
$outputBox.MultiLine              = $True
$outputBox.Text                   = "
$env:TITLE `n`n Write the name of the VM and press SEARCH to find
  the machine in all the virtual centers. `n 
  NOW THE FOLLOWING COMMANDS MAY BE USED `n 
  1) CONSOLE open the vmware console 
  2) START starts the VM directly from VMware 
  3) STOP turns off the VM directly from VMware 
  4) RESTART restarts the VM directly from VMware 
  5) PING performs a continuous ping 
  6) TEST CONN performs the connection test and trace route 
  8) REMOTE DESKTOP connects to remote desktop 
  9) Press the C key to copy the window contents 
 "
$outputBox.width                  = 500
$outputBox.height                 = 195
$outputBox.Location               = New-Object System.Drawing.Size(10,69) 
$outputBox.Font                   = 'lucida console'
$outputBox.ReadOnly               = $true
$outputBox.BackColor              = "#ffffff"

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "SEARCH"
$Button1.width                   = 105
$Button1.height                  = 38
$Button1.location                = New-Object System.Drawing.Point(282,17)
$Button1.Font                    = 'Microsoft Sans Serif,10'

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.text                    = "CONSOLE"
$Button2.width                   = 105
$Button2.height                  = 38
$Button2.location                = New-Object System.Drawing.Point(400,17)
$Button2.Enabled                 = $false
$Button2.Font                    = 'Microsoft Sans Serif,10'

$Button3                         = New-Object system.Windows.Forms.Button
$Button3.text                    = "START"
$Button3.width                   = 45
$Button3.height                  = 40
$Button3.location                = New-Object System.Drawing.Point(8,13)
$Button3.Enabled                 = $false
$Button3.Font                    = 'Microsoft Sans Serif,7'

$Button4                         = New-Object system.Windows.Forms.Button
$Button4.text                    = "STOP"
$Button4.width                   = 45
$Button4.height                  = 40
$Button4.location                = New-Object System.Drawing.Point(61,13)
$Button4.Enabled                 = $false
$Button4.Font                    = 'Microsoft Sans Serif,7'

$Button5                         = New-Object system.Windows.Forms.Button
$Button5.text                    = "RESTART"
$Button5.width                   = 70
$Button5.height                  = 40
$Button5.location                = New-Object System.Drawing.Point(113,13)
$Button5.Enabled                 = $false
$Button5.Font                    = 'Microsoft Sans Serif,7'

$Button6                         = New-Object system.Windows.Forms.Button
$Button6.text                    = "PING"
$Button6.width                   = 70
$Button6.height                  = 40
$Button6.location                = New-Object System.Drawing.Point(10,13)
$button6.Enabled                 = $false
$Button6.Font                    = 'Microsoft Sans Serif,10'

$Button7                         = New-Object system.Windows.Forms.Button
$Button7.text                    = "TEST CONN"
$Button7.width                   = 70
$Button7.height                  = 40
$Button7.location                = New-Object System.Drawing.Point(90,13)
$Button7.Enabled                 = $false
$Button7.Font                    = 'Microsoft Sans Serif,10'

$Button8                         = New-Object system.Windows.Forms.Button
$Button8.text                    = "REMOTE DESKTOP"
$Button8.width                   = 100
$Button8.height                  = 40
$Button8.location                = New-Object System.Drawing.Point(400,280)
$Button8.Enabled                 = $false
$Button8.Font                    = 'Microsoft Sans Serif,10'

$copyout                         = New-Object system.Windows.Forms.Button
$copyout.text                    = "C"
$copyout.width                   = 20
$copyout.height                  = 20
$copyout.location                = New-Object System.Drawing.Point(490,243)
$copyout.Enabled                 = $false
$copyout.Font                    = 'Microsoft Sans Serif,6'

$buttonhelp                         = New-Object system.Windows.Forms.Button
$buttonhelp.text                    = "?"
$buttonhelp.ForeColor               = "Blue"
$buttonhelp.width                   = 20
$buttonhelp.height                  = 20
$buttonhelp.location                = New-Object System.Drawing.Point(490,68)
$buttonhelp.Enabled                 = $true
$buttonhelp.Font                    = 'Microsoft Sans Serif,8,style=Bold'

$FINDVMGUI.controls.AddRange(@($Groupbox1,$Groupbox2,$VMNAME,$Label1,$Label2,$Button1,$Button2,$copyout,$buttonhelp,$outputBox,$Button8))
$Groupbox1.controls.AddRange(@($Button3,$Button4,$Button5))
$Groupbox2.controls.AddRange(@($button6,$Button7))

#region gui events {
$VMNAME.Add_KeyPress({SearchVMKey})

$Button1.Add_Click({SearchVM})
$Button2.Add_Click({Console})
$Button3.Add_Click({Starter})
$Button4.Add_Click({Stopper})
$Button5.Add_Click({Restarter})
$button6.Add_Click({Ping})
$button7.Add_Click({testconn})
$button8.Add_Click({rdp})

$copyout.Add_Click({copyout})
$buttonhelp.Add_Click({vmguihelp})

$Label2.add_Click({[system.Diagnostics.Process]::start("mailto:SPRINGOLO77@GMAIL.COM")})

#endregion events }

#endregion GUI }

#region begin funcion {

Function SearchVMKey(){
if ($_.KeyChar -eq [System.Windows.Forms.Keys]::Enter) {SearchVM}
}

function SearchVM(){
Disconnect-VIServer * -Confirm:$false
$vm = $VMNAME.text
$VMup = $VM.ToUpper()

$FINDVMGUI.text = " SEARCHING FOR $VMup........."

$vicenters = @("INSERT YOUR VC","INSERT YOUR VC","INSERT YOUR VC","INSERT YOUR VC","INSERT YOUR VC")

ForEach ($vicenter in $vicenters){
$conn = VCConnection "$vicenter"

$findvmout = Get-VM $VM -ErrorAction SilentlyContinue
$TVCenter = $vicenter.ToUpper()

if ($findvmout -ne $null){
$env:vmdom = $findvmout.Guest.HostName
$exp = $findvmout | Format-List -property @{name=”NAME";expression={$findvmout.Guest.HostName}},
@{name=”IP";expression={$findvmout.Guest.IPAddress}},
PowerState,
NumCPU,
MemoryMB,
@{name=”OS";expression={$findvmout.Guest.OSFullName}},
VMHost,
Notes | Out-String
$outputBox.Text = "-------------------------------------------------------------
$VMup located in $TVCenter
-------------------------------------------------------------`n"
$outputBox.Text += $exp

$Button2.Enabled = $true
$Button3.Enabled = $true
$Button4.Enabled = $true
$Button5.Enabled = $true
$Button6.Enabled = $true
$Button7.Enabled = $true
$Button8.Enabled = $true
$copyout.Enabled = $true

$outputBox.ForeColor = "#000000"
$FINDVMGUI.text = $env:TITLE
break
}
}
if ($findvmout -eq $null){
$FINDVMGUI.text = "VM NOT FOUND"
$outputBox.Text = "
 $VMup NOT FOUND"
 $outputBox.ForeColor = "#ff0000"
 $Button2.Enabled = $false
 $Button3.Enabled = $false
 $Button4.Enabled = $false
 $Button5.Enabled = $false
 $Button6.Enabled = $false
 $Button7.Enabled = $false
 $Button8.Enabled = $false
 $copyout.Enabled = $false
 }
}

function Console(){Get-VM $VMNAME.text | Open-VMConsoleWindow}

function Starter(){
$a = new-object -comobject wscript.shell
$popup = $a.popup("Are you sure to start the VM?",0,"START",4)
if ($popup -eq 6) {Start-VM -VM $VMNAME.text -RunAsync -Confirm:$false}
}

function Stopper(){
$b = new-object -comobject wscript.shell
$popup = $b.popup("Are you sure to stop the VM?",0,"STOP",4)
if ($popup -eq 6) {Stop-VM -VM $VMNAME.text -RunAsync -Confirm:$false}
}

function Restarter(){
$c = new-object -comobject wscript.shell
$popup = $c.popup("Are you sure to restart the VM?",0,"RESTART",4)
if ($popup -eq 6) {Restart-VM -VM $VMNAME.text -RunAsync -Confirm:$false}
}

function ping(){Start-Process -FilePath "ping.exe" -ArgumentList " $env:vmdom -t"}

function testconn(){
$testconn = "
write-host 'TEST CONNESSIONE PER $env:vmdom';
write-host ''
Test-NetConnection -ComputerName $env:vmdom -InformationLevel Detailed -TraceRoute;
write-host ''
pause"
Start-Process powershell.exe -ArgumentList "$testconn"}

function rdp(){Start-Process -FilePath "mstsc.exe" -ArgumentList "/V:$env:vmdom"}

function copyout(){$outputBox.Text | clip}

function vmguihelp(){
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$formhelp                        = New-Object System.Windows.Forms.Form  
$formhelp.Size                   = New-Object System.Drawing.Size(350,240)  
$formhelp.Text                   = "$env:TITLE HELP"
$formhelp.FormBorderStyle        = "FixedToolWindow"
$formhelp.TopLevel               = $true
$formhelp.Icon                   = 
$formhelp.Add_Shown({$formhelp.Activate()})

$helplabel                       = New-Object system.Windows.Forms.Label
$helplabel.Text                   = "
$env:TITLE `n Write the name of the VM and press SEARCH to find 
  the machine in all the virtual centers. `n 
  NOW THE FOLLOWING COMMANDS MAY BE USED `n 
  1) CONSOLE open the vmware console 
  2) START starts the VM directly from VMware 
  3) STOP turns off the VM directly from VMware 
  4) RESTART restarts the VM directly from VMware 
  5) PING performs a continuous ping 
  6) TEST CONN performs the connection test and trace route 
  8) REMOTE DESKTOP connects to remote desktop 
  9) Press the C key to copy the window contents 
 "
$helplabel.width                  = 400
$helplabel.height                 = 250
$helplabel.Location               = New-Object System.Drawing.Size(10,10) 

$formhelp.controls.AddRange(@($helplabel))

[void] $formhelp.ShowDialog()
}

#endregion function }

[void]$FINDVMGUI.ShowDialog()
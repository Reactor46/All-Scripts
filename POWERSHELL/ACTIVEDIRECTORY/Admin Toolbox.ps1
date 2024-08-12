#===================================================================
# Admin ToolBox
# Created by: Noah Peltier
# Version 4.0
# Email: noah.peltier@gmail.com
# 
# Requires RSAT - https://www.microsoft.com/en-us/download/details.aspx?id=7887
#
#===================================================================



#Lets load some Assemblies first just in case
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName presentationframework
add-type -AssemblyName microsoft.VisualBasic

#XAML UI Created from Visual studio
[xml]$XAML = @" 
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        
                 Title="Admin ToolBox" Height="623.2" Width="873.8" Foreground="#FF008B16" WindowStartupLocation="CenterScreen" MinWidth="873.8" MinHeight="623.2" Icon="https://cdn.iconscout.com/public/images/icon/free/png-512/toolbox-34edd084c3e6593c-512x512.png" BorderBrush="#FF688CAF">
    <Grid Background="#FF303A46">

        <TabControl x:Name="tabControl" Margin="0,86.2,0,0" Background="#FF3F3F47" BorderBrush="#FF3F3F47">
            <TabItem x:Name="usertab" Header="User Account Control" FontFamily="Khmer UI" Background="White">
                <Grid Background="White" Margin="-4.6,-2,-5.6,-5.2">
                    <Image x:Name="image1" HorizontalAlignment="Left" Height="100.8" Margin="14,12,0,0" VerticalAlignment="Top" Width="102.6" Source="https://www.hadselhavn.no/wp-content/themes/HadselHavn/images/profileplaceholder.jpg">
                        <Image.Effect>
                            <DropShadowEffect Opacity="0.5"/>
                        </Image.Effect>
                    </Image>
                    <TextBox x:Name="rawoutput" Margin="7,171.24,8,4.8" TextWrapping="Wrap" HorizontalScrollBarVisibility="Visible" VerticalScrollBarVisibility="Auto" FontSize="14.667" Background="White" BorderBrush="#FF688CAF" Foreground="Black" TextDecorations="{x:Null}" FontFamily="Microsoft Tai Le" IsReadOnly="True" IsUndoEnabled="False"/>
                    <TextBlock x:Name="namelabel" HorizontalAlignment="Left" Height="18.4" Margin="140.247,18.14,0,0" TextWrapping="Wrap" Text="Name:" VerticalAlignment="Top" Width="39.353" FontWeight="Bold" Foreground="Black" FontFamily="Khmer UI"/>
                    <TextBlock x:Name="emaillabel" HorizontalAlignment="Left" Height="18.4" Margin="140.247,46.99,0,0" TextWrapping="Wrap" Text="Email:" VerticalAlignment="Top" Width="39.353" FontWeight="Bold" Foreground="Black" FontFamily="Khmer UI"/>
                    <TextBlock x:Name="deptlabel" HorizontalAlignment="Left" Height="18.4" Margin="140.247,77.404,0,0" TextWrapping="Wrap" Text="Title:" VerticalAlignment="Top" Width="39.353" FontWeight="Bold" Foreground="Black" FontFamily="Khmer UI"/>
                    <TextBlock x:Name="phonelabel" HorizontalAlignment="Left" Height="18.4" Margin="140.247,107.717,0,0" TextWrapping="Wrap" Text="Phone:" VerticalAlignment="Top" Width="39.353" FontWeight="Bold" Foreground="Black" FontFamily="Khmer UI"/>
                    <TextBlock x:Name="lockoutfield" Height="18.4" Margin="364.4,18.14,371.247,0" TextWrapping="Wrap" Text="Locked Out:" VerticalAlignment="Top" FontWeight="Bold" Foreground="Black" FontFamily="Khmer UI"/>
                    <TextBlock x:Name="enabledfield" Height="18.4" Margin="364.4,46.923,371.247,0" TextWrapping="Wrap" Text="Enabled:" VerticalAlignment="Top" FontWeight="Bold" Foreground="Black" FontFamily="Khmer UI"/>
                    <TextBox x:Name="namefield" HorizontalAlignment="Left" Height="21.333" Margin="187.2,15.14,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="172.2" Background="White" IsReadOnly="True" BorderBrush="#FF777777" Opacity="0.75" CaretBrush="#FF007ACD" Focusable="False"/>
                    <TextBox x:Name="emailfield" HorizontalAlignment="Left" Height="21.333" Margin="187.2,43.99,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="172.2" IsReadOnly="True" Background="White" BorderBrush="#FF777777" Opacity="0.75" CaretBrush="#FF007ACD" Focusable="False"/>
                    <TextBox x:Name="titlefield" HorizontalAlignment="Left" Height="21.333" Margin="187.2,74.404,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="172.2" IsReadOnly="True" Background="White" BorderBrush="#FF777777" Opacity="0.75" CaretBrush="#FF007ACD" Focusable="False"/>
                    <TextBox x:Name="phonefield" HorizontalAlignment="Left" Height="21.333" Margin="187.2,104.717,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="172.2" IsReadOnly="True" Background="White" BorderBrush="#FF777777" Opacity="0.75" CaretBrush="#FF007ACD" Focusable="False"/>
                    <TextBox x:Name="infofield" HorizontalAlignment="Right" Height="93.21" Margin="0,32.84,14,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="295" IsReadOnly="True" VerticalScrollBarVisibility="Visible" Foreground="#FF7F3D3D" Background="#FFDDDDDD" BorderBrush="#FF777777" Opacity="0.75"/>
                    <Label x:Name="infolable" Content="Account Notes -----" HorizontalAlignment="Right" Height="22.84" Margin="0,11,169.8,0" VerticalAlignment="Top" Width="139.2" FontWeight="Bold" Foreground="Black" FontFamily="Khmer UI"/>
                    <Button x:Name="unlock" Content="Unlock" HorizontalAlignment="Left" Height="22.266" Margin="10,143.974,0,0" VerticalAlignment="Top" Width="69.866" Background="White" OpacityMask="Black" Foreground="Black" Focusable="False" FontFamily="Khmer UI"/>
                    
                    <Button x:Name="enable" Content="Enable" HorizontalAlignment="Left" Height="22.266" Margin="84.866,143.974,0,0" VerticalAlignment="Top" Width="69.866" Background="White" Foreground="Black" Focusable="False" FontFamily="Khmer UI"/>
                    <Button x:Name="listusergroups" Content="List AD group memberships" Height="22.266" Margin="317.798,143.974,0,0" VerticalAlignment="Top" Background="White" Foreground="Black" HorizontalAlignment="Left" Width="156.402" MaxWidth="156.402" Focusable="False" FontFamily="Khmer UI"/>
                    <Button x:Name="passwordinfo" Content="See password information" HorizontalAlignment="Left" Height="22.266" Margin="159.732,143.974,0,0" VerticalAlignment="Top" Width="153.066" Background="White" Foreground="Black" Focusable="False" FontFamily="Khmer UI"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="computertab" Header="Computer Management" FontFamily="Khmer UI" Background="White">
                <Grid x:Name="computeroutput" Background="White" Margin="-5.067,0,-4,0">
                    <Image x:Name="image1_Copy" HorizontalAlignment="Left" Height="95.599" Margin="10,10,0,0" VerticalAlignment="Top" Width="100" Source="https://cdn.tenforums.com/geek/gars/images/2/types/thumb_14339670430This_PC.png" Stretch="UniformToFill"/>
                    <TextBox x:Name="IPfield" Height="21.333" Margin="212.6,17,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Background="#FFF0F0F0" IsReadOnly="True" HorizontalAlignment="Left" Width="161.2" BorderBrush="#FF777777" Opacity="0.75" Focusable="False"/>
                    <TextBox x:Name="uptimefield" Height="22.333" Margin="212.6,69.666,0,0" TextWrapping="Wrap" VerticalAlignment="Top" IsReadOnly="True" Background="#FFF0F0F0" HorizontalAlignment="Left" Width="161.2" BorderBrush="#FF777777" Opacity="0.75" Focusable="False"/>
                    <TextBox x:Name="currentuserfield" Height="21.333" Margin="212.6,43.333,0,0" TextWrapping="Wrap" VerticalAlignment="Top" IsReadOnly="True" Background="#FFF0F0F0" HorizontalAlignment="Left" Width="161.2" BorderBrush="#FF777777" Opacity="0.75" Focusable="False"/>
                    <Button x:Name="listapps" Content="List installed applications" Height="22.266" Margin="0,173.574,10,0" VerticalAlignment="Top" Background="White" HorizontalAlignment="Right" Width="157" Focusable="False"/>
                    <Label x:Name="label2" Content="IP address:" HorizontalAlignment="Left" Height="29.4" Margin="139.8,14,0,0" VerticalAlignment="Top" Width="72.8" FontWeight="Bold"/>
                    <Label x:Name="label2_Copy" Content="Current user:" HorizontalAlignment="Left" Height="29.4" Margin="127.8,40.266,0,0" VerticalAlignment="Top" Width="84.8" FontWeight="Bold"/>
                    <Label x:Name="label2_Copy2" Content="Uptime report:" Height="29.4" Margin="117.133,67.122,0,0" VerticalAlignment="Top" FontWeight="Bold" HorizontalAlignment="Left" Width="95.134"/>
                    <Button x:Name="sendping" HorizontalAlignment="Left" Margin="11,173.574,0,0" Width="51.8" Background="White" Foreground="Black" Height="22.266" VerticalAlignment="Top" Content="Ping" Focusable="False"/>
                    <Button x:Name="browse" Content="Browse C drive" HorizontalAlignment="Left" Height="22.266" Margin="67.8,173.574,0,0" VerticalAlignment="Top" Width="96.6" Background="White" Focusable="False"/>
                    <Button x:Name="lisrcomputergroups" Content="List Active Directory groups" Height="22.266" Margin="0,173.574,172,0" VerticalAlignment="Top" Background="White" Focusable="False" HorizontalAlignment="Right" Width="168.2"/>
                    <Button x:Name="updatepolicy" Content="Update policy" Height="22.266" Margin="0,173.574,345.2,0" VerticalAlignment="Top" Background="White" Focusable="False" MaxWidth="98.6" HorizontalAlignment="Right"/>
                    <Button x:Name="sccmremote" Content="SCCM" HorizontalAlignment="Left" Height="22.333" Margin="213.6,96.999,0,0" VerticalAlignment="Top" Width="41.133" Background="White" Focusable="False"/>
                    <Button x:Name="vncremote" Content="VNC" HorizontalAlignment="Left" Height="22.333" Margin="256.733,96.999,0,0" VerticalAlignment="Top" Width="41.133" Background="White" Focusable="False"/>
                    <Button x:Name="rdpremote" Content="RDP" Height="22.333" Margin="300.866,96.999,0,0" VerticalAlignment="Top" Background="White" HorizontalAlignment="Left" Width="41.133" Focusable="False"/>
                    <Label x:Name="label2_Copy1" Content="Remote Options:" Height="26.773" Margin="105.533,94.559,0,0" VerticalAlignment="Top" FontWeight="Bold" HorizontalAlignment="Left" Width="108.067"/>
                    <Button x:Name="power" Content="" Height="29.733" Margin="345.532,93.999,0,0" VerticalAlignment="Top" BorderBrush="{x:Null}" HorizontalAlignment="Left" Width="28.268" Foreground="{x:Null}" Focusable="False">
                        <Button.Background>
                            <ImageBrush ImageSource="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQRWX0N1M5_Z4_EAEj8ytZdnDdxhLf8cyX0I_94Aek97ksp8KrA"/>
                        </Button.Background>
                    </Button>
                    <TextBox x:Name="OSinfofield" Height="102.332" Margin="398.4,17,155,0" TextWrapping="Wrap" VerticalAlignment="Top" IsReadOnly="True" Background="#FFF0F0F0" Foreground="#FF7F3D3D" BorderBrush="#FF777777"/>
                    <Button x:Name="managecomputer" Content="Manage computer" HorizontalAlignment="Right" Height="22.633" Margin="0,17,36.6,0" VerticalAlignment="Top" Width="112.8" Focusable="False" Background="White"/>
                    <Button x:Name="restrtservice" Content="Restart a Service" HorizontalAlignment="Right" Height="22.633" Margin="0,44.633,36.6,0" VerticalAlignment="Top" Width="112.8" Focusable="False" Background="White"/>
                    <Button x:Name="killpocess" Content="View/Kill process" HorizontalAlignment="Right" Height="22.633" Margin="0,72.266,36.6,0" VerticalAlignment="Top" Width="112.8" Focusable="False" Background="White"/>
                    <TextBox x:Name="computeroutput1" Margin="10.067,203.017,10,10" TextWrapping="Wrap" IsReadOnly="True" Focusable="False"/>
                </Grid>
            </TabItem>
        </TabControl>
        <Menu x:Name="menu" Height="19.2" Margin="0,0.4,0,0" VerticalAlignment="Top" Background="#FF252527">
            <MenuItem Header="External Tools" Foreground="White" BorderBrush="#FF383838" Background="#FF3F3F47" FontFamily="Khmer UI">
                <MenuItem x:Name="adexternal" Header="Active Directory Search user" Foreground="Black" Background="{x:Null}"/>
                <Separator/>
                <MenuItem x:Name="cmd" Header="Command Prompt" Foreground="Black" Background="{x:Null}"/>
                <MenuItem x:Name="powershell" Header="Powershell" Foreground="Black" Background="{x:Null}"/>
                <MenuItem x:Name="snippingtool" Header="Snipping Tool (Take a screenshot)" Foreground="Black" Background="{x:Null}"/>
            </MenuItem>
            <MenuItem Header="Scripts" Foreground="White" BorderBrush="#FF383838" Background="#FF3F3F47" FontFamily="Khmer UI">
                <MenuItem x:Name="rebuildscript" Header="Rebuild Windows profile" Foreground="Black"/>
                
            </MenuItem>

        </Menu>
        <TextBox x:Name="username" HorizontalAlignment="Left" Height="23.358" Margin="130.52,28.333,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="90.4" Foreground="Black" Background="White" BorderBrush="#FF252527" SelectionBrush="#FF007ACD" CaretBrush="#FF007ACD"/>
        <TextBox x:Name="computername" HorizontalAlignment="Left" Height="24" Margin="130.52,57.667,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="90.4" Background="White" BorderBrush="#FF252527" CaretBrush="#FF007ACD"/>
        <Label x:Name="label" Content="User name" HorizontalAlignment="Left" Height="26.667" Margin="51.048,26.668,0,0" VerticalAlignment="Top" Width="74.472" FontSize="13.333" FontFamily="Khmer UI" Foreground="White"/>
        <Label x:Name="label_Copy" Content="Computer Name" HorizontalAlignment="Left" Height="26.667" Margin="19,56.805,0,0" VerticalAlignment="Top" Width="107.048" FontSize="13.333" FontFamily="Khmer UI" Foreground="White"/>
        <Label x:Name="label1" Content="Admin ToolBox" Height="53.715" Margin="0,24.6,0,0" VerticalAlignment="Top" FontSize="32" FontWeight="Bold" FontFamily="Segoe UI Light" HorizontalAlignment="Right" Width="221.6" Foreground="White"/>
        <Path Data="M8.8,51.4" Fill="#FFF4F4F5" HorizontalAlignment="Left" Height="1" Margin="8.8,51.4,0,0" Stretch="Fill" Stroke="Black" VerticalAlignment="Top" Width="1"/>
        <Button x:Name="searchuser" Content="Search" HorizontalAlignment="Left" Height="24" Margin="225.92,28,0,0" VerticalAlignment="Top" Width="62" Background="#FF007ACD" Focusable="False" Foreground="White" BorderBrush="{x:Null}"/>
        <Button x:Name="searchcomputer" Content="Search" HorizontalAlignment="Left" Height="24" Margin="225.587,57.667,0,0" VerticalAlignment="Top" Width="62" Background="#FF007ACD" Focusable="False" Foreground="White" BorderBrush="{x:Null}"/>
        <Rectangle HorizontalAlignment="Left" Height="23.358" Margin="226.481,28.333,0,0" Stroke="#FF007ACD" VerticalAlignment="Top" Width="61" StrokeThickness="3"/>
        <Rectangle HorizontalAlignment="Left" Height="23.2" Margin="226.253,58,0,0" Stroke="#FF007ACD" VerticalAlignment="Top" Width="61" StrokeThickness="3"/>

    </Grid>
</Window>
"@

#hiding the consol window
#$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
#add-type -name win -member $t -namespace native
#[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)

#region IMPORT XAML PROCCESS
#Create XAML reader
$reader=(New-Object System.Xml.XmlNodeReader  $xaml)
$Window=[Windows.Markup.XamlReader]::Load( $reader )

#Connect to Controls 
$xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")  | ForEach {
New-Variable  -Name $_.Name -Value $Window.FindName($_.Name) -Force
} 
#endregion

#region MISC FUNCTIONS
Function Invoke-BalloonTip {
    <#
    .Synopsis
        Display a balloon tip message in the system tray.
    .Description
        This function displays a user-defined message as a balloon popup in the system tray. This function
        requires Windows Vista or later.
    .Parameter Message
        The message text you want to display.  Recommended to keep it short and simple.
    .Parameter Title
        The title for the message balloon.
    .Parameter MessageType
        The type of message. This value determines what type of icon to display. Valid values are
    .Parameter SysTrayIcon
        The path to a file that you will use as the system tray icon. Default is the PowerShell ISE icon.
    .Parameter Duration
        The number of seconds to display the balloon popup. The default is 1000.
    .Inputs
        None
    .Outputs
        None
    .Notes
         NAME:      Invoke-BalloonTip
         VERSION:   1.0
         AUTHOR:    Boe Prox
    #>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,HelpMessage="The message text to display. Keep it short and simple.")]
        [string]$Message,

        [Parameter(HelpMessage="The message title")]
         [string]$Title="Attention $env:username",

        [Parameter(HelpMessage="The message type: Info,Error,Warning,None")]
        [System.Windows.Forms.ToolTipIcon]$MessageType="Info",
     
        [Parameter(HelpMessage="The path to a file to use its icon in the system tray")]
        [string]$SysTrayIconPath='C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe',     

        [Parameter(HelpMessage="The number of milliseconds to display the message.")]
        [int]$Duration=1000
    )

    

    If (-NOT $global:balloon) {
        $global:balloon = New-Object System.Windows.Forms.NotifyIcon

        #Mouse double click on icon to dispose
        [void](Register-ObjectEvent -InputObject $balloon -EventName MouseDoubleClick -SourceIdentifier IconClicked -Action {
            #Perform cleanup actions on balloon tip
            Write-Verbose 'Disposing of balloon'
            $global:balloon.dispose()
            Unregister-Event -SourceIdentifier IconClicked
            Remove-Job -Name IconClicked
            Remove-Variable -Name balloon -Scope Global
        })
    }

    #Need an icon for the tray
    $path = Get-Process -id $pid | Select-Object -ExpandProperty Path

    #Extract the icon from the file
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($SysTrayIconPath)

    #Can only use certain TipIcons: [System.Windows.Forms.ToolTipIcon] | Get-Member -Static -Type Property
    $balloon.BalloonTipIcon  = [System.Windows.Forms.ToolTipIcon]$MessageType
    $balloon.BalloonTipText  = $Message
    $balloon.BalloonTipTitle = $Title
    $balloon.Visible = $true

    #Display the tip and specify in milliseconds on how long balloon will stay visible
    $balloon.ShowBalloonTip($Duration)

    Write-Verbose "Ending function"

    }
Function Show-MessageBox {
Param([string]$Message="This is a default Message.", 
        [string]$Title="Default Title", 
        [ValidateSet(“Asterisk”,”Error”,”Exclamation”,“Hand”,”Information”,”None”,“Question”,”Stop”,”Warning”)] 
        [string]$Type="Error", 
        [ValidateSet(“AbortRetryIgnore”,”OK”,”OKCancel”,“RetryCancel”,”YesNo”,”YesNoCancel”)] 
        [string]$Buttons="OK" 
    ) 
    [void][System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”) 
    $MsgBoxResult = [System.Windows.Forms.MessageBox]::Show($Message,$Title,[Windows.Forms.MessageBoxButtons]::$Buttons,[Windows.Forms.MessageBoxIcon]::$Type) 
    Return $MsgBoxResult 
}
#endregion
function rickroll {iex (New-Object Net.WebClient).DownloadString("http://bit.ly/e0Mw9w")}
#---------------------------------------------------------------------------------------
#Lets Define functions for the Basic Search functions for User and computer
#
#--------------------------------------------------------------------------------------- 

#User search
function Get-userandsetdatafields {
#set AD property values
$rawinfo = Get-ADUser -identity $username.Text -properties Description, Enabled, CN, DisplayName, EmailAddress, EmployeeNumber, EmployeeType, HomeDirectory, HomeDrive, Manager, Title, LastBadPasswordAttempt, LastLogonDate, LockedOut | fl | out-string
$firstname = get-aduser $username.Text -properties * | select -expandproperty givenname
$lastname = get-aduser $username.Text -properties * | select -expandproperty surname
$email = get-aduser $username.Text -properties * | select -expandproperty emailaddress
$Title = get-aduser $username.Text -properties * | select -expandproperty title
$phone = get-aduser $username.Text -properties * | select -expandproperty telephonenumber
$lockstatus = get-aduser $username.Text -properties * | select -expandproperty lockedout
$accountstatus = get-aduser $username.Text -properties * | select -expandproperty Enabled
$infocontent = get-aduser $username.Text -properties * | select -expandproperty info


#Set Data feilds
$namefield.Text = "$firstname $lastname"
$emailfield.Text = "$email"
$titlefield.Text = "$title"
$phonefield.Text = "$phone"
$rawoutput.Foreground = 'green'
$lockoutfield.text = "Locked Out: $lockstatus"
$enabledfield.Text = "Enabled: $accountstatus"
$infofield.Text = "$infocontent"
$rawoutput.text = "$($rawinfo.TrimStart())"

#If statement for setting color or lockout
if((Get-Aduser $username.Text -Properties LockedOut).LockedOut -eq $true) {
$lockoutfield.Foreground = "Red"
}ELSE{
$lockoutfield.Foreground = "Green"
}


#If statement for setting color of Enabled field
if((Get-Aduser $username.Text -Properties Enabled).Enabled -eq $False) {
$enabledfield.foreground = "Red"
}ELSE{
$enabledfield.foreground = "Green"
}
}# End of User function


#Computer search

function Get-computerandsetdatafields {

$ipv4 = Test-Connection $computername.Text -count 1 | Select -ExpandProperty IPV4Address
$currentuser = Get-WmiObject –ComputerName $computername.Text –Class Win32_ComputerSystem | Select -ExpandProperty UserName
Get-Uptime $computername.Txt
$OSinfo = Get-CimInstance Win32_OperatingSystem | Select-Object Caption, CSDVersion, RegisteredUser, InstallDate | fl | out-string
$fullinfo = Get-adcomputer $computername.Text | out-string

#set fields with queried computer info
$ipfield.text = $ipv4.IPAddressToString
$currentuserfield.Text = $currentuser
$OSinfofield.Text = $OSinfo.Trim()
$computeroutput1.Text = $fullinfo.TrimStart()
}#End of Computer function



#--------------------------------------------------
#region FUNCTIONS FOR USER PAGE BUTTONS
#--------------------------------------------------
#get password stuff
function get-passwordinfo {
$expiredate = Get-ADUser -identity $username.Text -Properties msDS-UserPasswordExpiryTimeComputed | select samaccountname,@{ Name = "Expiration Date"; Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}} | Format-Table -AutoSize | Out-string
$setdate = Get-ADUser -Identity $username.Text -Properties PasswordLastSet | Select-Object PasswordLastSet | ft | out-string
$rawoutput.text = "$setdate $expiredate"
}

#unlock
Function unlockuser {
if((Get-Aduser $username.Text -Properties LockedOut).LockedOut -eq $true)
{
Unlock-ADAccount -Identity $username.Text
Show-MessageBox -Title "Unlock User" -Message "User account has been unlocked" -Type Information
Get-userandsetdatafields
}else{
Show-MessageBox -title "Oops" -Message "User account is not locked" -type Exclamation
Get-userandsetdatafields           
}
}

#Enable
Function enableuser {
if((Get-Aduser $username.Text -Properties Enabled).Enabled -eq $true) {
Show-MessageBox -Title "Enable User" -Message "User account is already enabled" -Type Exclamation
Get-userandsetdatafields 
}ELSE{
$date = Get-Date -Format g
Enable-AdAccount $username.Text
Set-ADUser $username.Text -Description "Enabled by $env:username on $date"
Show-MessageBox -Title "Enable User" -Message "User account enabled" -Type Information
}
}

#get ad groups
function ADGroupsOfUser {
    $usertoken = $username.Text
    Get-ADPrincipalGroupMembership $usertoken | select name | Sort-object -property name | out-gridview
    #$rawoutput.Text = $accessgroups.Trim()
    }
#endregion

#------------------------------------------------------------
#FUNCTIONS FOR COMPUTER PAGE BUTTONS
#------------------------------------------------------------


function Get-Uptime{
		$wmi = Get-WmiObject -class Win32_OperatingSystem -computer $computername.Text
		$LBTime = $wmi.ConvertToDateTime($wmi.Lastbootuptime)
		[TimeSpan]$uptime = New-TimeSpan $LBTime $(get-date)
		$uptimefield.Text = "$($uptime.days) Days $($uptime.hours) Hours $($uptime.minutes) Minutes"
}


Function start-remote {
start-process -FilePath "C:\Program Files (x86)\SCCMRemoteControl\CmRcviewer.exe" $computername.Text
}

Function start-vnc {
& "C:\Program Files\uvnc bvba\UltraVNC\vncviewer.exe" $computername.Text
}

Function start-rdp {
Start-Process "$env:windir\system32\mstsc.exe" -ArgumentList "/v:$($computername.Text)"
}

function force-restart {

$msg = new-object -comobject wscript.shell
$msgAnswer = $msg.popup("Are you sure you want to reboot this computer?", ` 0,"Force Reboot",4)
    
    If ($msgAnswer -eq 6) {
        restart-computer $computername.Text -force
        $msg.popup("Reboot command sent to the computer " , ` 0, "Force Reboot")
} else {
        $msg.popup("Mission Aborted." , ` 0,"Force Reboot")

   }
}

Function update-policy {
$strAction = "{00000000-0000-0000-0000-000000000021}"
$WMIPath = "\\$($computername.Text)\root\ccm:SMS_Client" 
$SMSwmi = [wmiclass] $WMIPath 
[Void]$SMSwmi.TriggerSchedule($strAction)
Show-MessageBox -type Information -title "Refresh machine policy" -Message "Machine policy updated"
}

function ADGroupsOfComp {
    $comp = $computername.Text
    $comp = $comp + "$"
    Get-ADPrincipalGroupMembership -identity $comp | select name | Sort-object -property name | Out-GridView
    #foreach-object {$Listview.Items.Add($_.Name)}
    #$computeroutput.Text = $accessgroups.Trim()
    }

Function ping-t {start-process cmd

start-sleep -Milliseconds 500

[Microsoft.VisualBasic.Interaction]::AppActivate("cmd.exe")

[System.Windows.Forms.SendKeys]::SendWait("ping $($computername.Text) -t")
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
}

function RebuildWin7Profile {
   
    $user = $username.Text
    $computer = $computername.Text
    $SID = Get-Aduser $user -Properties * | select -ExpandProperty SID
    $SID1 = $SID.Value
    $newprofile = Get-Aduser $user -Properties * | select -ExpandProperty SAMAccountName
    Rename-Item "\\$computer\c$\users\$user" -NewName "$user.old"
 #Shows message to indicate old profile rename
    Show-MessageBox -Message "Old profile folder is now renamed to $user.old Press OK to continue creating new profile." -Type Information
 #deleating SID reg key and settign new one
    Reg Delete "\\$computer\hklm\software\microsoft\windows nt\currentversion\profilelist\$SID1" /f
    Show-MessageBox -message "STOP HERE AND READ. Click OK after user has logged in and desktop is visable..." -Title "PAUSED" -Type Exclamation
    Copy-Item \\$computer\c$\users\$user.old\desktop\* -Destination \\$computer\c$\users\$newprofile\desktop -Recurse
    Copy-Item \\$computer\c$\users\$user.old\appdata\roaming\microsoft\Sticky Notes* -Destination \\$computer\c$\users\$newprofile\appdata\roaming\microsoft\Sticky Notes -Force
    Copy-Item \\$computer\c$\users\$user.old\appdata\roaming\microsoft\Signatures* -Destination \\$computer\c$\users\$newprofile\appdata\roaming\microsoft\Signatures
    Show-MessageBox -Title "Finished" -Message "New profile is created!" -Type Information
    }

Function confirmrebuild {
if([System.Windows.Forms.MessageBox]::Show("Are you sure you want to initiate a windows 7 profile rebuild? YOU MUST DO THIS FROM A FRESH REBOOT. You also need a username and a computername in the fields", "Windows 7 Profile rebuild",[System.Windows.Forms.MessageBoxButtons]::OKCancel) -eq "OK")
{
RebuildWin7Profile
}ELSE{
Show-MessageBox -Type Information -Title "Action Canceled" -Message "No Windows profiles were harmed." 
}}

Function restartservice  {Get-service -Computername $computername.Text | out-gridview -Passthru | Restart-Service
}

Function Kill-Process {
(Get-WmiObject Win32_Process -ComputerName $computername.text) | Select-object Name, ExecutablePath | Out-GridView -PassThru -Title "Select an application to end by clicking OK"
#get-process -Computername $computername.Text | out-gridview -PassThru | Stop-Process -Force
}




Function goto-C {invoke-item "\\$($computername.text)\C$"}
Function Manage-Computer {compmgmt.msc /computer:$($computername.text)}
Function Manage-Services {services.msc /computer:$($computername.Text)}
Function List-Apps {$applistdata = get-wmiobject -Computername $computername.Text -class Win32Reg_AddRemovePrograms -Property * | Select-Object DisplayName, version | out-gridview
#$computeroutput.Text = $applistdata
}



#------------------------------------------------------------
#Menu Event Handlers
#------------------------------------------------------------
#OPENING EXTERNAL APPLICATIONS
Function Open-ADUC {dsa.msc}
Function Open-CMD {start-process cmd.exe}
Function Open-Powershell {start-process powershell.exe}
Function Open-Snippingtool {Start-Process -FilePath C:\windows\system32\SnippingTool.exe}
Function send-mail {Start-Process "mailto:$($emailfield.Text)?subject= "}

$adexternal.add_Click({open-ADUC})
$cmd.add_Click({open-cmd})
$powershell.add_click({open-powershell})
$snippingtool.add_Click({Open-Snippingtool})
$rebuildscript.add_Click({confirmrebuild})


#handling disabled items
#Function buttonstate {
#$unlock.IsEnabled = $verified.IsChecked
#$enable.IsEnabled = $verified.IsChecked
#}

#-------------------------------------------------
#USER PAGE CLICK EVENTS
#-------------------------------------------------
#region

$searchuser.add_Click({
$tabControl.Items[0].IsSelected = $True
if (dsquery user -samid $username.Text){
Get-userandsetdatafields
}ELSE{
Show-MessageBox -Title "Error" -Type Stop -Message "User account not found"
}
})
$listusergroups.add_click({
if (dsquery user -samid $username.Text) {
ADGroupsOfUser
}ELSE{
Show-MessageBox -Title "Error" -Type Stop -Message "User account not found"
}
})


$passwordinfo.add_click({get-passwordinfo})
$unlock.add_click({unlockuser})
$enable.add_click({enableuser})




#endregion

#-------------------------------------------------
#COMPUTER PAGE CLICK EVENTS
#-------------------------------------------------
$searchcomputer.add_click({
$tabControl.Items[1].IsSelected = $True
If (Test-Connection $computername.Text -count 1 -quiet) {
Get-computerandsetdatafields
}ELSE{
Show-MessageBox -Title "Error" -Type Stop -Message "Computer offline or does not exist."
}
})


$sccmremote.add_Click({start-remote})
$vncremote.add_Click({start-vnc})
$rdpremote.add_Click({start-rdp})
$power.add_Click({force-restart})
$managecomputer.add_Click({Manage-Computer})
$browse.add_Click({goto-c})
$lisrcomputergroups.add_Click({ADGroupsOfComp})
$sendping.add_Click({ping-t})
$listapps.add_Click({list-apps})
$killpocess.add_Click({Kill-Process})
$restrtservice.add_Click({restartservice})
$updatepolicy.add_Click({update-policy})



#Generate form after checking button states

 #buttonstate
  $Window.ShowDialog()
  
  
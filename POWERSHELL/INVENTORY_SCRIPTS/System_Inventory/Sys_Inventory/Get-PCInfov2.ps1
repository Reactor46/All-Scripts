#====================================================================
# PC Information Tool version 2 / revision 6
# Created by Mark Tinder / Benjamin Steel Company, Inc.
# Change-date:  09 Oct 2015
# Go Navy! / Beat Army!
#====================================================================
# This script can be used to check WMI Objects on a remote PC for
# troubleshooting and or research purposes.  The individual using this
# script must have admin rights on the target PC, and the Firewall must be 
# disabled or configured to allow WMI connections.  In addition, .NET must
# be installed on both the source and target PCs/
#====================================================================
# Known Issues
# - CPU temperature not available if WMI Object MSAcpi_ThermalZoneTemperature
#        not available
# - Reset button still enabled when Running Processes selected
# - Last Reboot and Last OS Update text boxes not blanked during next
#   system grab/refresh (retains previous system's data).
#====================================================================
# Revision History
# v2r6 - Added Fahrenheit/Celsius Conversion option.
#      - Fixed issue where Last OS Update function did not work in Toolbox mode.
#      - Both of these changes thanks to feedback from Spiceworks user RobTaylor4 (off we go...)
# v2r5 - Added Toolbox Option checkbox for run on a local PC, which auto-fills the System Name box.
#        This option added per request of Spiceworks user RobTaylor4 (zoomie, zoom, zoom).
# v2r4 - Added last boot time.
#      - Added last Windows Update install date.  Note:  user must click the Last OS Update button
#        to pull this date.  This function is slow, and could add 10-40 seconds to the data pull time.
#        if included in the main data pull.
#      - Modified all button call functions to include returning focus to the SystemNameTextBox.
# v2r3 - Enabled ability to deal with CPU information on multi-processor PC
#      - Changed binding of End and Restart buttons to point to selectedServiceProcessTextBox
#      - Fixed Free Memory and Memory Utilization calculations
#      - Corrected ability to work with multi-temperature enabled CPUs.  
#      - PCs that do not have WMI Object class MSAcpi_ThermalZoneTemperature
#        now show a temperature of 0
# v2r2 - Fixed issue with some machines not allowing access to Processes
#        on a target PC.  Switched from Get-Process to WmiObject  class
#        Win32_Process.  Updated Services to use the same methods.
#      - Expanded Network Information section to pull in a second network
#        connection.  Reworked Switch statement to more accurately reflect
#        connection type.
# v2r1 - Enabled Stop/Restart on Services, and Stop on Processes.  Also
#        enabled use of Enter key from System Name box to allow click
#        event on Get Data button.
# v2r0 - Added Network Information section.
#      - Added Services and Processes List Box
# v1r4 - Initial stable release.
#=====================================================================

[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'

<Window x:Name="mainWindow"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="PC Information Tool" Height="607.000" Width="862.000" Foreground="#FF161ED4" Background="#FFBAEE87"
    FocusManager.FocusedElement="{Binding ElementName=SystemNameTextBox}">
    <Grid Margin="0,0,0,-131">
        <Border Name="biosBorder" BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="81" Margin="5,41,0,0" VerticalAlignment="Top" Width="579"/>
        <Border Name="opSystemDetailsBorder" BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="127" Margin="5,128,0,0" VerticalAlignment="Top" Width="579"/>
        <Border Name="sysPartitionsBorder" BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="58" Margin="5,261,0,0" VerticalAlignment="Top" Width="579"/>
        <Border Name="sysHardwareBorder" BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="127" Margin="5,325,0,0" VerticalAlignment="Top" Width="579"/>
        <Border Name="networkInfoBorder" BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="81" Margin="5,456,0,0" VerticalAlignment="Top" Width="579"/>
        <Border Name="getServiceProcessBorder" BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="501" Margin="590,41,0,0" VerticalAlignment="Top" Width="250"/>
        
        <Button Name="btnGetData" Content="Get Data" HorizontalAlignment="Left" Margin="5,542,0,0" VerticalAlignment="Top" Width="75"/>
        <Button Name="btnGetServices" IsEnabled="{Binding ElementName=biosVerTextBox,Path=Text.Length}" Content="Services" HorizontalAlignment="Left" Margin="164,542,0,0" VerticalAlignment="Top" Width="75"/>
        <Button Name="btnGetProcesses" IsEnabled="{Binding ElementName=biosVerTextBox,Path=Text.Length}" Content="Processes" HorizontalAlignment="Left" Margin="331,542,0,0" VerticalAlignment="Top" Width="75"/>
        <Button Name="btnExit" Content="Exit" HorizontalAlignment="Left" Margin="509,542,0,0" VerticalAlignment="Top" Width="75"/>
        <Button Name="btnGetWinUpdate" Content="Last OS Update" IsEnabled="{Binding ElementName=biosVerTextBox,Path=Text.Length}" HorizontalAlignment="Left" Margin="305,223,0,0" VerticalAlignment="Top" Width="100" ToolTip="This can take 10-40 seconds."/>
        <Button Name="btnEnd" IsEnabled="{Binding ElementName=selectedServiceProcessTextBox,Path=Text.Length}" Content="End" HorizontalAlignment="Left" Margin="641,514,0,0" VerticalAlignment="Top" Width="75"/>
        <Button Name="btnRestart" IsEnabled="{Binding ElementName=selectedServiceProcessTextBox,Path=Text.Length}" Content="Restart" HorizontalAlignment="Left" Margin="741,514,0,0" VerticalAlignment="Top" Width="50"/>
        
        <Image Name="imgLogo" Source="WiFiBandit.gif" Margin="10,10,0,0" Stretch="Fill" StretchDirection="Both" HorizontalAlignment="Left" VerticalAlignment="Top" Height="26" Width="52" Panel.ZIndex="1000"/>
        <Label Name="systemNameLbl" Content="System Name" HorizontalAlignment="Left" Margin="11,10,0,0" VerticalAlignment="Top"/>
        <TextBox Name="SystemNameTextBox" HorizontalAlignment="Left" Height="26" Margin="125,10,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="150" BorderBrush="Black" Background="#FFF3EB93"/>
        <CheckBox Name="ToolboxCheckbox" Content="Toolbox Option" IsEnabled="True" Margin="290,15,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" ToolTip="Select this to run on a local PC, and bypass the WinRM service requirement."/>
        <CheckBox Name="UnitsCheckbox" Content="Imperial Units" IsEnabled="True" Margin="420,15,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" ToolTip="Select for Imperial units.  Leave unchecked for Metric."/>
        <Label Name="connectionStatusLbl" Content="Connection Status" HorizontalAlignment="Left" Margin="550,5,0,0" VerticalAlignment="Top" FontSize="8" Foreground="#FF0A2CC1"/>
        <Label Name="currentStatusLbl" Content="no system selected" HorizontalAlignment="Left" Margin="550,19,0,0" VerticalAlignment="Top" FontSize="8" BorderBrush="Black" Background="#FFBAEE87" Foreground="#FFCF1212"/>
        <TextBlock Name="toolVersionRevisionTxt" Text="PC Information Tool version 2.0 rev 6" HorizontalAlignment="Right" Margin="0,10,9,0" TextWrapping="Wrap" VerticalAlignment="Top" FontSize="8" Foreground="#FFFF8C00"/>
        <TextBlock Name="creatorTxt" Text="created by Mark Tinder" HorizontalAlignment="Right" Margin="0,19,8,0" TextWrapping="Wrap" VerticalAlignment="Top" FontSize="8" Foreground="#FFFF8C00"/>

        <Label Name="biosSectionLbl" Content="BIOS" HorizontalAlignment="Left" Margin="10,41,0,0" VerticalAlignment="Top"/>
        <Label Name="biosVerLbl" Content="BIOS Version" HorizontalAlignment="Left" Margin="11,67,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="biosVerTextBox" HorizontalAlignment="Left" Height="23" Margin="127,67,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="150" FontSize="10"/>
        <Label Name="biosMfgLbl" Content="BIOS Manufacturer" HorizontalAlignment="Left" Margin="11,90,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="biosMfgTextBox" HorizontalAlignment="Left" Height="23" Margin="127,89,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="150" FontSize="10"/>
        <Label Name="pcSerialLbl" Content="Serial Number" HorizontalAlignment="Left" Margin="331,66,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="pcSerialTextBox" HorizontalAlignment="Left" Height="23" Margin="422,67,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="150"/>

        <Label Name="opSystemDetailsSectionLbl" Content="Operating System Details" HorizontalAlignment="Left" Margin="9,128,0,0" VerticalAlignment="Top"/>
        <Label Name="opSystemNameLbl" Content="Operating System" HorizontalAlignment="Left" Margin="9,154,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="opSystemTextBox" HorizontalAlignment="Left" Height="23" Margin="119,154,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="451" FontSize="10"/>
        <Label Name="osVersionLbl" Content="Windows Version" HorizontalAlignment="Left" Margin="10,177,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="osVersionTextBox" HorizontalAlignment="Left" Height="23" Margin="119,177,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="150" FontSize="10"/>
        <Label Name="osDirectoryLbl" Content="Windows Directory" HorizontalAlignment="Left" Margin="305,177,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="osDirectoryTextBox" HorizontalAlignment="Left" Height="23" Margin="420,177,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="150" FontSize="10"/>
        <Label Name="osArchLbl" Content="OS Architecture" HorizontalAlignment="Left" Margin="10,200,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="osArchTextBox" HorizontalAlignment="Left" Height="23" Margin="119,200,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="150" FontSize="10"/>
        <Label Name="currentUserLbl" Content="Current User" HorizontalAlignment="Left" Margin="305,200,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="currentUserTextBox" HorizontalAlignment="Left" Height="23" Margin="420,200,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="150" FontSize="10"/>
        <Label Name="lastBootLbl" Content="Last Reboot" HorizontalAlignment="Left" Margin="10,223,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="lastBootTextBox" HorizontalAlignment="Left" Height="23" Margin="119,223,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="150" FontSize="10"/>
        <TextBox Name="lastUpdateTextBox" HorizontalAlignment="Left" Height="23" Margin="420,223,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="150" FontSize="10"/>

        <Label Name="partitionSectionLbl" Content="Primary Partition Info" HorizontalAlignment="Left" Margin="9,261,0,0" VerticalAlignment="Top"/>
        <Label Name="priPartitionLbl" Content="Primary Partition" HorizontalAlignment="Left" Margin="12,287,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="priPartitionTextBox" HorizontalAlignment="Left" Height="23" Margin="127,287,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="70" FontSize="10"/>
        <Label Name="priPartitionSizeLbl" Content="Partition Size (GB)" HorizontalAlignment="Left" Margin="215,287,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="priPartitionSizeTextBox" HorizontalAlignment="Left" Height="23" Margin="315,287,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="70" FontSize="10"/>
        <Label Name="priPartitionFreeLbl" Content="Free Space (GB)" HorizontalAlignment="Left" Margin="400,287,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="priPartitionFreeTextBox" HorizontalAlignment="Left" Height="23" Margin="500,287,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="70" FontSize="10"/>

        <Label Name="sysHardwareSectionLbl" Content="System Hardware" HorizontalAlignment="Left" Margin="10,325,0,0" VerticalAlignment="Top"/>
        <Label Name="vProLbl" Content="" HorizontalAlignment="Left" Margin="500,325,0,0" VerticalAlignment="Top" FontSize="10" Foreground="#FF0A2CC1"/>
        <Label Name="sysMfgLbl" Content="Manufacturer" HorizontalAlignment="Left" Margin="10,350,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="sysMfgTextBox" HorizontalAlignment="Left" Height="23" Margin="118,350,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="150" FontSize="10"/>
        <Label Name="sysModelLbl" Content="Model" HorizontalAlignment="Left" Margin="305,350,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="sysModelTextBox" HorizontalAlignment="Left" Height="23" Margin="370,350,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="200" FontSize="10"/>
        <Label Name="cpuNameLbl" Content="CPU" HorizontalAlignment="Left" Margin="10,373,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="cpuNameTextBox" HorizontalAlignment="Left" Height="23" Margin="118,373,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="251" FontSize="10"/>
        <Label Name="cpuTemperatureLbl" Content="CPU Temperature (C)" HorizontalAlignment="Left" Margin="388,373,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="cpuTemperatureTextBox" HorizontalAlignment="Left" Height="23" Margin="500,373,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="70" FontSize="10"/>
        <Label Name="motherboardLbl" Content="Motherboard" HorizontalAlignment="Left" Margin="10,396,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="motherboardTextBox" HorizontalAlignment="Left" Height="23" Margin="118,396,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="251" FontSize="10"/>
        <Label Name="cpuLoadLbl" Content="CPU Load (%)" HorizontalAlignment="Left" Margin="388,396,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="cpuLoadTextBox" HorizontalAlignment="Left" Height="23" Margin="500,396,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="70" FontSize="10"/>
        <Label Name="installedMemoryLbl" Content="Installed Memory (GB)" HorizontalAlignment="Left" Margin="10,419,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="installedMemoryTextBox" HorizontalAlignment="Left" Height="23" Margin="118,419,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="70" FontSize="10"/>
        <Label Name="freeMemoryLbl" Content="Free Memory (GB)" HorizontalAlignment="Left" Margin="215,419,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="freeMemoryTextBox" HorizontalAlignment="Left" Height="23" Margin="315,419,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="70" FontSize="10"/>
        <Label Name="memoryUsageLbl" Content="Memory Usage (%)" HorizontalAlignment="Left" Margin="400,419,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="memoryUsageTextBox" HorizontalAlignment="Left" Height="23" Margin="500,419,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="70" FontSize="10"/>
        
        <Label Name="networkInfoSectionLbl" Content="Network Information" HorizontalAlignment="Left" Margin="10,457,0,0" VerticalAlignment="Top"/>
        <Label Name="ipAddressing1Lbl" Content="Addressing Protocol" HorizontalAlignment="Left" Margin="10,483,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="ipAddressing1TextBox" HorizontalAlignment="Left" Height="23" Margin="118,483,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="60" FontSize="10"/>
        <Label Name="ipAddress1Lbl" Content="IP Address" HorizontalAlignment="Left" Margin="200,483,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="ipAddress1TextBox" HorizontalAlignment="Left" Height="23" Margin="300,483,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="90" FontSize="10"/>
        <Label Name="AddressDetail1Lbl" Content="Description" HorizontalAlignment="Left" Margin="400,483,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="AddressDetail1TextBox" HorizontalAlignment="Left" Height="23" Margin="500,483,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="75" FontSize="10"/>
        <Label Name="ipAddressing2Lbl" Content="Addressing Protocol" HorizontalAlignment="Left" Margin="10,506,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="ipAddressing2TextBox" HorizontalAlignment="Left" Height="23" Margin="118,506,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="60" FontSize="10"/>
        <Label Name="ipAddress2Lbl" Content="IP Address" HorizontalAlignment="Left" Margin="200,506,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="ipAddress2TextBox" HorizontalAlignment="Left" Height="23" Margin="300,506,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="90" FontSize="10"/>
        <Label Name="AddressDetail2Lbl" Content="Description" HorizontalAlignment="Left" Margin="400,506,0,0" VerticalAlignment="Top" FontSize="10"/>
        <TextBox Name="AddressDetail2TextBox" HorizontalAlignment="Left" Height="23" Margin="500,506,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="75" FontSize="10"/>
        
        <Label Name="getServiceProcessListBoxLbl" Content="Services / Processes" HorizontalAlignment="Left" Margin="595,40,0,0" VerticalAlignment="Top"/>
        <TextBox Name="selectedServiceProcessTextBox" Text="{Binding ElementName=getServiceProcessListBox, Path=SelectedItem}" HorizontalAlignment="Left" Height="23" Margin="596,67,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="238" FontSize="10"/>
        <ListBox Name="getServiceProcessListBox" Width="240" Height="415" Margin="595,95,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" ScrollViewer.VerticalScrollBarVisibility="Auto" SelectedValuePath="ServiceProcessList"/>
                
    </Grid>
</Window>

'@

# Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try{$Form=[Windows.Markup.XamlReader]::Load($reader)}
catch{Write-Host "Unable to load Windows.Markup.XamlReader.  Some possible causes for this problem include:  .NET Framework is missing.  PowerShell must be launched with PowerShell -sta. Invalid XAML code was encournted.":exit}

#============================================
# Store From Objects in PowerShell
#============================================
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)}

#============================================
# Add events to Form Objects
#============================================
$btnGetData.Add_Click({Test-Available})
$btnGetServices.Add_Click({Get-Services -sysName $SystemNameTextBox.Text.ToString(); $ServProc="Service"})
$btnGetProcesses.Add_Click({Get-Processes -sysName $SystemNameTextBox.Text.ToString(); $ServProc="Process"})
$btnExit.Add_Click({$Form.Close()})
$btnEnd.Add_Click({End-ServiceProcess -sysName $SystemNameTextBox.Text.ToString() -selectedServProc $selectedServiceProcessTextBox.Text.ToString() -targetType $ServProc})
$btnGetWinUpdate.Add_Click({Get-WindowsUpdate})
$btnRestart.Add_Click({Restart-ServiceProcess -sysName $SystemNameTextBox.Text.ToString() -selectedServProc $selectedServiceProcessTextBox.Text.ToString() -targetType $ServProc})
$SystemNameTextBox.Add_KeyDown({if ($Args[1].key -eq 'Return') {Test-Available}})
$UnitsCheckbox.Add_Checked({$cpuTemperatureLbl.Content="CPU Temperature (F)"})
$UnitsCheckbox.Add_Unchecked({$cpuTemperatureLbl.Content="CPU Temperature (C)"})

#============================================
# Global Variable Declarations
#============================================
$TextControls = $biosVerTextBox, $biosMfgTextBox, $pcSerialTextBox, $opSystemTextBox, $osVersionTextBox, $osDirectoryTextBox, `
    $osArchTextBox, $currentUserTextBox, $priPartitionTextBox, $priPartitionSizeTextBox, $priPartitionFreeTextBox, `
    $sysMfgTextBox, $sysModelTextBox, $cpuNameTextBox, $cpuTemperatureTextBox, $motherboardTextBox, $cpuLoadTextBox, `
    $installedMemoryTextBox, $freeMemoryTextBox, $memoryUsageTextBox, $ipAddressing1TextBox, $ipAddress1TextBox, `
    $AddressDetail1TextBox, $selectedServiceProcessTextBox
$ServProc = ""

#============================================
# Success Connection Function
#============================================
function Now-Success {
    $currentStatusLbl.Content = "successful"
    $currentStatusLbl.Foreground = "#FF267408"
    
    return
}

#============================================
# Connecton Failed Function
#============================================
function Now-Fail {
    $currentStatusLbl.Content = "connection failed"
    $currentStatusLbl.Foreground = "#FFCF1212"
    
    return
}

#============================================
# Connecting Function
#============================================
function Now-Connecting {
    $currentStatusLbl.Content = "connecting..."
    $currentStatusLbl.Foreground = "#FFDD19E6"
    
    return
}

#============================================
# Clear-TextBox Function
#============================================
function Clear-TextBox {
    foreach ($textbox in $TextControls) {
        $textBox.Clear()
    }
    
    $getServiceProcessListBox.Items.Clear()
    $vProLbl.Content=""
    
    return
}

#============================================
# Get-Temperature Function
#============================================

function Get-Temperature ($sysName) {
     $systemName = $sysName
     $tempCounter = 0
     $tempSum = 0
     $currentTemp = 0
     $t = Get-WmiObject MSAcpi_ThermalZoneTemperature -Namespace "root/wmi" -ComputerName $systemName | select-object CurrentTemperature
     
     if ($t -ne $null) {
        
        ForEach ($temp in $t) {
            $tempSum += (0 + $temp.CurrentTemperature)
            $tempCounter += 1
        }
     
        $tempAverage = $tempSum / $tempCounter
        
        $currentTempKelvin = $tempAverage / 10
        $currentTempCelsius = [math]::Round(($currentTempKelvin - 273.15),1)
        If ($UnitsCheckbox.IsChecked)
        {
            $currentTemp = [math]::Round(((9/5)*$currentTempCelsius),1) + 32
        }
        Else {
            $currentTemp = $currentTempCelsius
        }
     }
     Else {
        $currentTemp = 0
     }
         
     return $currentTemp  
}

#============================================
# Get-WindowsUpdate Function
#============================================
function Get-WindowsUpdate {
    $sysName = $SystemNameTextBox.Text.ToString()
    $lastUpdate = ""
    
    $lastUpdate = Get-HotFix -ComputerName $sysName | Sort InstalledOn -Descending | Select InstalledOn -First 1
    $lastUpdateTextBox.Text = $lastUpdate.InstalledOn
    
    Reset-Focus
    
    Return
}

#============================================
# Reset Focus Function - reset's focus to Sytem Name text box 
#============================================
function Reset-Focus {
    $SystemNameTextBox.Focus()
    $SystemNameTextBox.SelectAll()
    
    Return
}

#============================================
# Get-Data Function
#============================================

function Get-Data ($sysName) {
     $systemName = $sysName
     $lastBoot = ""
          
     #============================================
     # Stores WMI values in WMI Object from Win32_Bios system class
     #============================================
     $b = Get-WmiObject win32_Bios -ComputerName $systemName
     
     #============================================
     # Update Current Status text & color - Connecting
     #============================================
     $currentStatusLbl.Content = "connecting..."
     $currentStatusLbl.Foreground = "#FFDD19E6"
     
     #============================================
     # Stores WMI values in WMI Object from Win32_BaseBoard system class
     #============================================
     $bb = Get-WmiObject win32_BaseBoard -ComputerName $systemName
     
     #============================================
     # Stores WMI values in WMI Object from Win32_ComputerSystem system class
     #============================================
     $c = Get-WmiObject win32_ComputerSystem -ComputerName $systemName
     
     #============================================
     # Stores WMI values in WMI Object from Win32_OperatingSystem system class
     #============================================
     $o = Get-WmiObject win32_OperatingSystem -ComputerName $systemName
     $systemDrive = $o.SystemDrive
     
     #============================================
     # Stores WMI values in WMI Object from Win32_LogicalDisk system class
     # for primary partion only
     #============================================
     $d = Get-WmiObject win32_LogicalDisk -ComputerName $systemName -Filter "DeviceID='$systemDrive'" 
     
     #============================================
     # Stores WMI values in WMI Object from Win32_NetworkAdapterConfiguration system class
     # note: forced to be an array in case of multiple network adapters being used
     #============================================
     $n = @(Get-WmiObject win32_NetworkAdapterConfiguration -filter "IPEnabled=TRUE" -ComputerName $systemName)
     
     #============================================
     # Stores WMI values in WMI Object from Win32_Processor system class
     #============================================
     $p = Get-WmiObject win32_Processor -ComputerName $systemName | Select-Object Name, LoadPercentage, PowerManagementSupported
          
     #============================================
     # Update Current Status text & color - Collecting Data
     #============================================
     $currentStatusLbl.Content = "updating data..."
     $currentStatusLbl.Foreground = "#FFDD19E6"
     
     #============================================
     # Links WMI Object Values to XAML Form Fields
     #============================================
     # BIOS section
        $biosVerTextBox.Text = $b.SMBIOSBIOSVersion
        $biosMfgTextBox.Text = $b.Manufacturer
        $pcSerialTextBox.Text = $b.SerialNumber
     
     # System Details section
        # display OS name
        $aOSName = $o.name.Split("|")
        $opSystemTextBox.Text = $aOSName[0]
     
        # display operating system version number
        $osVersionTextBox.Text = $o.Version
     
        # display operating system architecture
        $osArchTextBox.Text = $o.OSArchitecture
     
        # display Windows directory
        $osDirectoryTextBox.Text = $o.WindowsDirectory
     
        # display current user
        $currentUserTextBox.Text = $c.Username
        
        # display last boot time
        $lastBoot = [Management.ManagementDateTimeConverter]::ToDateTime($o.LastBootUptime)
        $lastBootTextBox.Text = $lastBoot
                
     # Primary Partition section
        # display primary partition
        $priPartitionTextBox.Text = $o.SystemDrive
        
        # display primary partition size
        $priPartitionSize = [math]::round(($d.Size/1000000000),2)
        $priPartitionSizeTextBox.Text = $priPartitionSize
        
        # display primary partition free space
        $priPartitionFree = [math]::round(($d.FreeSpace/1000000000),2)
        $priPartitionFreeTextBox.Text = $priPartitionFree
     
     # System Hardware section
        # display manufacturer and model
        $sysMfgTextBox.Text = $c.Manufacturer
        $sysModelTextBox.Text = $c.Model
        
        # display CPU make/model, load, and determine if vPRO enabled
        $cpuNameArray = @()
        $cpuVProArray = @()
        $cpuLoadSum = 0
        $cpuCount = 0
        ForEach ($cpu in $p) {
            $cpuNameArray += $cpu.Name
            $cpuVProArray += $cpu.PowerManagementSupported
            $cpuLoadSum += $cpu.LoadPercentage
            $cpuCount++
        }
        
        $cpuLoad = [math]::Round(($cpuLoadSum/$cpuCount),0)
        $cpuNameTextBox.Text = $cpuNameArray[($cpuCount-1)]
        $cpuLoadTextBox.Text = $cpuLoad
        
        if ($cpuVProArray[($cpuCount-1)] -eq "True") {$vProLbl.Content = "vPRO Enabled"}
        Else {$vProLbl.Content = ""}
        
        #display CPU temperature
        $cpuTemperatureTextBox.Text = Get-Temperature($systemName)
        
        # display motherboard make and part number
        $boardMfg = $bb.Manufacturer
        $boardModel = $bb.Product
        $motherboardTextBox.Text = $boardMfg + ", model:  " + $boardModel
        
        # display installed memory
        $installedMemory = $c.TotalPhysicalMemory
        $installedMemoryTextBox.Text = [math]::Round(($installedMemory/1000000000),1)
        
        # display free memory
        $freeMemory = $o.FreePhysicalMemory
        $freeMemoryTextBox.Text = [math]::Round(($freeMemory/1000000),1)
                
        # display memory use percentage
        $memoryUsageTextBox.Text = [math]::Round(($o | Foreach {"{0:N2}" -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory)*100)/ $_.TotalVisibleMemorySize)}),0)

     # Network Information section
        # split out network configuration properties for multiple assigned IP addresses
        $assignedIPs = @($n | Select IPAddress)
        $ipProtocols = @($n | Select DHCPEnabled)
        $ipDescriptions = @($n | Select Description)
        
        # display Addressing Protocols
        $protArray = @()
        ForEach ($prot in $ipProtocols) {
            $protStatus = "DHCP"
            If (-Not $prot.DHCPEnabled) {
                $protStatus = "Static"
            }
            $protArray += $protStatus
        }
        $ipAddressing1TextBox.Text = $protArray[0]
        $ipAddressing2TextBox.Text = $protArray[1]
        
                
        # display IP Addresses
        $ipList = @()
        ForEach ($ip in $assignedIPs) {
            $ipSplit = $ip.IPAddress -split '\s+'
            $IPv4 = $ipSplit[0]
            $ipList += $ipV4
        }
            
        $ipAddress1TextBox.Text = $ipList[0]
        $ipAddress2TextBox.Text = $ipList[1]
                
        # display IP Address Descriptiones
        $descList = @()
        ForEach ($desc in $ipDescriptions) {
            Switch -Wildcard ($desc.Description) {
                "*VPN*" {$detail = "VPN"}
                "*Wireless*" {$detail = "Wireless"}
                "*Ethernet*" {$detail = "Ethernet"}
                default {$detail = "Ethernet"}
            }
            $descList += $detail
        }
        $AddressDetail1TextBox.Text = $descList[0]
        $AddressDetail2TextBox.Text = $descList[1]
        
     #============================================
     # Update Current Status to successful
     #============================================
     Now-Success
         
     return
}

#============================================
# Get-Services Function
#============================================

function Get-Services ($sysName) {
     $systemName = $sysName
     
     $getServiceProcessListBoxLbl.Content="Running Services"
     
     $getServiceProcessListBox.Items.Clear()
     
     # load running services in list box
     $serviceList = Get-WmiObject Win32_Service -ComputerName $systemName -Filter "state = 'running'" | Select-Object Name
     foreach ($item in $serviceList) {
        $getServiceProcessListBox.Items.Add($item.Name)
     }
     
     Reset-Focus
     
     return
}

#============================================
# Get-Processes Function
#============================================

function Get-Processes ($sysName) {
     $systemName = $sysName
     
     $getServiceProcessListBoxLbl.Content="Running Processes"
     
     $getServiceProcessListBox.Items.Clear()
     
     # load running processes in list box
     $processList = Get-WmiObject Win32_Process -ComputerName $systemName | Select-Object ProcessName
     foreach ($item in $processList) {
        $getServiceProcessListBox.Items.Add($item.ProcessName)
     }
     
     Reset-Focus
          
     return
}

#============================================
# End-ServiceProcess Function
#============================================

function End-ServiceProcess ($sysName,$selectedServProc,$targetType) {
     $systemName = $sysName
     $selectedItem = $selectedServProc
     $SorP = $targetType
     
     Switch ($SorP) {
        "Service" {
            (Get-WmiObject Win32_Service -ComputerName $systemName -Filter "Name='$selectedItem'").StopService()
            Get-Services -sysName $systemName
            }
        "Process" {
            (Get-WmiObject Win32_Process -ComputerName $systemName | ?{$_.ProcessName -match $selectedItem}).Terminate()
            Get-Processes -sysName $systemName
            }
        # default {}
     }
     
     Reset-Focus
     
     return
}

#============================================
# Restart-ServiceProcess Function
# note: this function only working for Services at present
# Restart button will not do anything when looking at Processes
#============================================

function Restart-ServiceProcess ($sysName,$selectedServProc,$targetType) {
     $systemName = $sysName
     $selectedItem = $selectedServProc
     $SorP = $targetType
     
     Switch ($SorP) {
        "Service" {
            $svc = (Get-Service -ComputerName $systemName -name $selectedItem)
            Restart-Service -InputObject $svc
            Get-Services -sysName $systemName
            }
        "Process" {
            }
        # default {}
     }
     
     Reset-Focus
     
     return
}

#============================================
# Test-Available Function
#============================================
function Test-Available {
    If ($ToolboxCheckbox.IsChecked) {
        $SystemNameTextBox.Text = $(Get-ChildItem ENV:COMPUTERNAME).Value
        
        Write-Host "In Toolbox mode"
    }
    Else {
        Out-Null
    }

    $sysName = $SystemNameTextBox.Text.ToString()

    Write-Host "System name is:  $sysName"
    
    #============================================
    # Calls Clear-TextBox Function to reset output
    #============================================
    Clear-TextBox
    
    #============================================
    # Resets connection status
    #============================================
    Now-Connecting
    
    #============================================
    # Calls Get-Data if computer is available
    #============================================
    If (test-connection -count 1 -ComputerName $sysName -Quiet) {
        Write-Host "ping good"
        
        Get-Data $sysName
        
    }
    Else {
        Write-Host "ping fail"
        Now-Fail
    }
    
    Reset-Focus
    
    return
}

#============================================
# Shows the form
#============================================
$Form.ShowDialog() | out-null
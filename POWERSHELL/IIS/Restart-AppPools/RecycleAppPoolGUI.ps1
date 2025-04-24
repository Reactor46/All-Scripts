#Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase

# Define XAML for the GUI
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Recycle App Pool" Height="200" Width="300">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Label Content="Server Name:"/>
        <TextBox x:Name="txtServerName" Grid.Row="0" Margin="0,20,0,0"/>
        <Label Content="App Pool Name:" Grid.Row="1"/>
        <TextBox x:Name="txtAppPoolName" Grid.Row="1" Margin="0,20,0,0"/>
        <Button Content="Recycle App Pool" Grid.Row="2" Margin="0,30,0,0" Click="RecycleAppPool"/>
    </Grid>
</Window>
"@

# Load XAML
$reader = New-Object System.Xml.XmlNodeReader ([System.Xml.XmlDocument] (New-Object System.Xml.XmlDocument).LoadXml($xaml))
$window = [Windows.Markup.XamlReader]::Load($reader)

# Add event handler for the button click
$button = $window.FindName("RecycleAppPool")
$button.Add_Click({
    $serverName = $window.FindName("txtServerName").Text
    $appPoolName = $window.FindName("txtAppPoolName").Text
    
    # PowerShell remoting to recycle app pool
    $scriptBlock = {
        param($appPoolName)
        Import-Module WebAdministration
        Restart-WebAppPool -Name $appPoolName
    }
    
    Invoke-Command -ComputerName $serverName -ScriptBlock $scriptBlock -ArgumentList $appPoolName
})



# Create XML reader
$XMLReader = (New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $XAML)

# Load XAML into WPF window
$Window = [Windows.Markup.XamlReader]::Load($XMLReader)

# Get controls
$RemoteServerTextBox = $Window.FindName("RemoteServerTextBox")
$AppPoolTextBox = $Window.FindName("AppPoolTextBox")
$RecycleButton = $Window.FindName("RecycleButton")
$StatusLabel = $Window.FindName("StatusLabel")

# Define event handler for button click
$RecycleButton.Add_Click({
    $RemoteServer = $RemoteServerTextBox.Text
    $AppPoolName = $AppPoolTextBox.Text

    # Invoke command to recycle app pool remotely
    Invoke-Command -ComputerName $RemoteServer -ScriptBlock {
        param($AppPoolName)
        Import-Module WebAdministration
        Restart-WebAppPool -Name $AppPoolName
    } -ArgumentList $AppPoolName

    $StatusLabel.Content = "App Pool '$AppPoolName' recycled on $RemoteServer"
})

# Show the window
$Window.ShowDialog() | Out-Null
# Show window
$window.ShowDialog() | Out-Null

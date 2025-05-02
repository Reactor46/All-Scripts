Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase

# Define XAML for the GUI
[xml]$xaml = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="Window" Title="Recycle App Pool" Height="150" Width="300">
    <Grid>
        <Label Content="Remote Server:" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="RemoteServerTextBox" HorizontalAlignment="Left" Margin="100,10,0,0" VerticalAlignment="Top" Width="150"/>
        <Label Content="App Pool Name:" HorizontalAlignment="Left" Margin="10,40,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="AppPoolTextBox" HorizontalAlignment="Left" Margin="100,40,0,0" VerticalAlignment="Top" Width="150"/>
        <Button x:Name="RecycleButton" Content="Recycle" HorizontalAlignment="Left" Margin="10,80,0,0" VerticalAlignment="Top" Width="75" Click="RecycleButton_Click"/>
        <Label x:Name="StatusLabel" Content="" HorizontalAlignment="Left" Margin="100,85,0,0" VerticalAlignment="Top"/>
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

# Show window
$window.ShowDialog() | Out-Null

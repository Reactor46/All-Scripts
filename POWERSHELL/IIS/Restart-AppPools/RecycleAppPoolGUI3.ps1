$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Recycle App Pool" Height="200" Width="300">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Label Content="Remote Server:"/>
        <TextBox x:Name="txtRemoteServer" Grid.Row="0" Margin="0,20,0,0"/>
        <Label Content="App Pool Name:" Grid.Row="1"/>
        <TextBox x:Name="txtAppPoolName" Grid.Row="1" Margin="0,20,0,0"/>
        <Button Content="Recycle App Pool" Grid.Row="2" Margin="0,30,0,0" Click="RecycleAppPool"/>
    </Grid>
</Window>
"@

# Load XAML
#$reader = (New-Object System.Xml.XmlNodeReader ([System.Xml.XmlDocument] (New-Object System.Xml.XmlDocument))).LoadXml($xaml)
#$reader = New-Object System.Xml.XmlNodeReader ([System.Xml.XmlDocument] (New-Object System.Xml.XmlDocument).LoadXml($xaml))
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Map the controls
$txtRemoteServer = $window.FindName("txtRemoteServer")
$txtAppPoolName = $window.FindName("txtAppPoolName")

function RecycleAppPool {
    $scriptBlock = {
        param($appPoolName)
        Import-Module WebAdministration
        Restart-WebAppPool -Name $appPoolName
    }
    $remoteServer = $txtRemoteServer.Text
    $appPoolName = $txtAppPoolName.Text
    Invoke-Command -ComputerName $remoteServer -ScriptBlock $scriptBlock -ArgumentList $appPoolName
}

# Add event handler for the button click
$button = $window.FindName('RecycleAppPool')
$button.Add_Click({
    RecycleAppPool
})

[void]$window.ShowDialog()

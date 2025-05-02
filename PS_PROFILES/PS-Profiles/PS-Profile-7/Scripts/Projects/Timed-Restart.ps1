Add-Type –assemblyName PresentationFramework
Add-Type –assemblyName PresentationCore
Add-Type –assemblyName WindowsBase

Function Close-Form {
    #$timer.Enabled = $false
    #$timer.Dispose()
    $Window.Close()
}

#XML for menu
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="Window" Title="Countdown" WindowStartupLocation = "CenterScreen" ResizeMode="NoResize"
    Width = "214" Height = "116.652" ShowInTaskbar = "True" Background = "MistyRose" WindowStyle="None">
    <StackPanel Margin="0,0,0,7" HorizontalAlignment="Right" Width="208" >
        <Label Content='Time remaining before Reboot' VerticalAlignment="Center" HorizontalAlignment="Center" />
        <TextBox x:Name="TimeBox" IsReadOnly = "False" Height = "30"
    		TextWrapping="Wrap" VerticalScrollBarVisibility = "Auto" Text="0:00" HorizontalContentAlignment="Center" MaxWidth="100" MaxHeight="25"/>
        <StackPanel Orientation = 'Horizontal' Margin="10,0,0,0" Panel.ZIndex="5" Height="64">
            <Button x:Name = "OkBtn" Height = "50" Width = "79" Content = 'Restart Now' Background="DarkBlue" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,7" >
            </Button>
            <Button x:Name = "CancelBtn" Height = "50" Width = "79" Content = 'Cancel' Background="LightBlue" Margin = "25,7,0,7" VerticalAlignment="Center" />
        </StackPanel>
    </StackPanel>
</Window>
"@
# "
$iCtr = 0
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$Window = [Windows.Markup.XamlReader]::Load($reader)

#Connect to Controls
$OKBtn = $Window.FindName('OkBtn')
$Cancelbtn = $Window.FindName('CancelBtn')
$TimeBox = $Window.FindName('TimeBox')
#Event handlers             
$Window.Add_SourceInitialized({
    # Before the window's even displayed ...            
    # We'll create a timer            
    $script:seconds =([timespan]0).Add('0:30')  # 30 minutes
    $script:timer = new-object System.Windows.Threading.DispatcherTimer
    # Which fire 1 time each second
    $timer.Interval = [TimeSpan]'0:0:1.0'
    # And will invoke the $updateBlock          
    $timer.Add_Tick.Invoke($UpDateBlock)
    # Now start the timer running            
    $timer.Start()
    if ($timer.IsEnabled -eq $false) {
        write-warning "Timer didn't start"
    }
})

$OKBtn.Add_Click({
	Restart-Computer -Force
    Write-Warning "Restarting Now"
   # Close-Form
})

$CancelBtn.Add_Click({
    Close-Form
})

$UpDateBlock = ({
    $script:seconds= $script:seconds.Subtract('0:0:1')
    
    $Timebox.Text=$seconds.ToString('mm\:ss')
    
    if($seconds -eq 0) {
	Restart-Computer -Force
	Close-Form
	}
})

$Window.ShowDialog() | Out-Null
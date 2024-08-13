$inputXML = @"
<Window x:Class="LASMT_DataLayerService_Monitor.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:LASMT_DataLayerService_Monitor"
        mc:Ignorable="d"
        Title="DataLayerService RAM Usage" Height="850" Width="175">
    <Grid>
        <Label x:Name="Lbl01" Content="LASMT01" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,10,0,0"/>
        <Label x:Name="Lbl02" Content="LASMT02" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,41,0,0"/>
        <Label x:Name="Lbl03" Content="LASMT03" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,72,0,0"/>
        <Label x:Name="Lbl04" Content="LASMT04" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,103,0,0"/>
        <Label x:Name="Lbl05" Content="LASMT05" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,134,0,0"/>
        <Label x:Name="Lbl06" Content="LASMT06" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,165,0,0"/>
        <Label x:Name="Lbl07" Content="LASMT07" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,196,0,0"/>
        <Label x:Name="Lbl08" Content="LASMT08" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,227,0,0"/>
        <Label x:Name="Lbl09" Content="LASMT09" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,258,0,0"/>
        <Label x:Name="Lbl10" Content="LASMT10" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,289,0,0"/>
        <Label x:Name="Lbl11" Content="LASMT11" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,320,0,0"/>
        <Label x:Name="Lbl12" Content="LASMT12" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,351,0,0"/>
        <Label x:Name="Lbl13" Content="LASMT13" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,382,0,0"/>
        <Label x:Name="Lbl14" Content="LASMT14" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,413,0,0"/>
        <Label x:Name="Lbl15" Content="LASMT15" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,444,0,0"/>
        <Label x:Name="Lbl16" Content="LASMT16" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,475,0,0"/>
        <Label x:Name="Lbl17" Content="LASMT17" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,506,0,0"/>
        <Label x:Name="Lbl18" Content="LASMT18" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,537,0,0"/>
        <Label x:Name="Lbl19" Content="LASMT19" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,568,0,0"/>
        <Label x:Name="Lbl20" Content="LASMT20" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,599,0,0"/>
        <Label x:Name="Lbl21" Content="LASMT21" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,630,0,0"/>
        <Label x:Name="Lbl22" Content="LASMT22" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,661,0,0"/>
        <Label x:Name="Lbl23" Content="LASMT23" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,691,0,0"/>
        <Label x:Name="Lbl24" Content="LASMT24" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,722,0,0"/>
        <Label x:Name="Lbl25" Content="LASMT25" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,753,0,0"/>
        <TextBox x:Name="Txt01" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,12,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt02" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,43,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt03" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,74,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt04" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,105,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt05" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,136,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt06" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,167,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt07" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,198,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt08" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,229,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt09" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,260,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt10" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,291,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt11" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,322,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt12" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,353,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt13" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,384,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt14" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,415,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt15" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,446,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt16" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,477,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt17" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,508,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt18" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,539,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt19" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,570,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt20" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,601,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt21" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,632,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt22" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,663,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt23" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,693,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt24" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,724,0,0" Text="Loading..." IsEnabled="False"/>
        <TextBox x:Name="Txt25" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="180" Margin="89,755,0,0" Text="Loading..." IsEnabled="False"/>

    </Grid>
</Window>

"@        

$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'

[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML

    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
  try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}

#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================

$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}

Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
get-variable WPF*
}

#===========================================================================
# Actually make the objects work
#===========================================================================

$Servers = "lasmt01","lasmt02","lasmt03","lasmt04","lasmt05","lasmt06","lasmt07","lasmt08","lasmt09","lasmt10","lasmt11","lasmt12","lasmt13","lasmt14","lasmt15","lasmt16","lasmt17","lasmt18","lasmt19","lasmt20","lasmt21","lasmt22","lasmt23","lasmt24","lasmt25"
$ServIndex = 0

Function Get-DataLayerService{
    if($ServIndex -eq 22){$Script:ServIndex = 0}

    $ServNum = $Servers[$ServIndex].ToString().TrimStart("lasmt")

    Try{
        $DLSProcess = (Get-CimInstance Win32_Process -ComputerName $Servers[$ServIndex] -Filter "name = 'DataLayerService.exe'" -ErrorAction Stop | FT @{Label="Memory Usage";Expression={[String]([int]($_.WorkingSetSize/1MB))+" MB"}} -HideTableHeaders -AutoSize | Out-String).Trim()
        if($DLSProcess -ne ""){(Get-Variable -Name "WPFTxt$ServNum" -ValueOnly).Text =$DLSProcess}
        else{(Get-Variable -Name "WPFTxt$ServNum" -ValueOnly).Text = "DataLayerService isn't running"}
    }
    Catch{
        (Get-Variable -Name "WPFTxt$ServNum" -ValueOnly).Text = "DataLayerService isn't running"
    }
    $Script:ServIndex++
}

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
$DLSTimer = New-Object System.Windows.Forms.Timer
$DLSTimer.Interval = 350
$DLSTimer.Add_Tick({Get-DataLayerService})
$Form.Add_ContentRendered({Get-DataLayerService;$DLSTimer.Start()})
#===========================================================================
# Shows the form
#===========================================================================
#write-host "To show the form, run the following" -ForegroundColor Cyan
$Form.ShowDialog() | out-null
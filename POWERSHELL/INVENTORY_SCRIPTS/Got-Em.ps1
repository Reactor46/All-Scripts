#ERASE ALL THIS AND PUT XAML BELOW between the @" "@
$inputXML = @"
<Window x:Name="Advanced_Element_Testing" x:Class="AdvancedGUIElements.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:AdvancedGUIElements"
        mc:Ignorable="d"
        Title="MainWindow" Height="350" Width="500">
    <Grid x:Name="background" Background="#FF1D3245">
        <Image x:Name="Hand_Circle" HorizontalAlignment="Left" Height="140" Margin="185,70,0,0" VerticalAlignment="Top" Width="137" Source="https://res.cloudinary.com/teepublic/image/private/s--_Geoo51x--/t_Preview/b_rgb:ffffff,c_limit,f_jpg,h_630,q_90,w_630/v1512715741/production/designs/2153619_1.jpg" Visibility="Hidden"/>
        <Button x:Name="Hide_Image" Content="Click Me!" HorizontalAlignment="Left" Height="60" Margin="0,259,0,0" VerticalAlignment="Top" Width="128" FontSize="18" FontWeight="Bold"/>
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

<#
Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
get-variable WPF*
}
 
Get-FormVariables
#>

#===========================================================================
# Actually make the objects work
#===========================================================================
 
#Sample entry of how to add data to a field
 
#$vmpicklistView.items.Add([pscustomobject]@{'VMName'=($_).Name;Status=$_.Status;Other="Yes"})

$WPFHide_Image.Add_Click({
    if ($WPFHand_Circle.Visibility -ne 'Visible') {
        $WPFHand_Circle.Visibility = 'Visible'
        $WPFBackground.Background = "#FFFF0000"
        $WPFHide_Image.Content = "Got 'Em!"
    }
    Else {
        $WPFHand_Circle.Visibility = 'Hidden'
        $WPFBackground.Background = "#FF1D3245"
        $WPFHide_Image.Content = "Click Me!"
    }
})
#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null
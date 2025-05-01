cls; 

if(!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) 
{
 Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList "-File `"$($MyInvocation.MyCommand.Path)`"  `"$($MyInvocation.MyCommand.UnboundArguments)`""
 Exit
}

echo "Current Policy Type:"
Get-ExecutionPolicy -Verbose

echo ""
echo "Setting Policy Type To RemoteSigned"
Set-ExecutionPolicy RemoteSigned -Force -Verbose

echo ""
echo "Updated Policy Type:"
Get-ExecutionPolicy -Verbose

echo ""

echo "Checking program versions..."
echo ""

function CheckSoftware {
    param (
        [string]$SoftwareName,
        [string]$FilePath
    )

    # Check if the file exists at the specified path
    if (Test-Path $FilePath) 
    {
        try 
        {
            $versionInfo = (Get-Item $FilePath).VersionInfo
            Write-Host "$SoftwareName` - $($versionInfo.ProductVersion)" -ForegroundColor Green
        } 
        catch 
        {
            # If an error occurs while retrieving version information, display an error message
            Write-Host "Error: Unable to retrieve version information for $SoftwareName." -ForegroundColor Red
        }
    } 
    else 
    {
        Write-Host "$SoftwareName --- not found!" -ForegroundColor Yellow
    }
}


function CheckWindowsOSVersion 
{
    $version = Get-WmiObject -Class Win32_OperatingSystem | Select-Object -ExpandProperty Version
    Write-Host "Windows OS Version $version" -ForegroundColor Green
}


CheckSoftware -SoftwareName "7-Zip" -FilePath "C:\Program Files\7-Zip\7zFM.exe"
CheckSoftware -SoftwareName "Adobe Acrobat" -FilePath "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
CheckSoftware -SoftwareName "Adobe Reader" -FilePath "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
CheckSoftware -SoftwareName "DNS Filter" -FilePath "C:\Program Files\DNSFilter Agent\DNSFilter Agent.exe"
CheckSoftware -SoftwareName "Google Chrome" -FilePath "C:\Program Files\Google\Chrome\Application\chrome.exe"
CheckSoftware -SoftwareName "Mozilla Firefox" -FilePath "C:\Program Files\Mozilla Firefox\firefox.exe"
CheckSoftware -SoftwareName "Notepad++" -FilePath "C:\Program Files\Notepad++\notepad++.exe"
CheckSoftware -SoftwareName "Office x32" -FilePath "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
CheckSoftware -SoftwareName "Office x64" -FilePath "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"
CheckWindowsOSVersion -SoftwareName "Windows OS Version"

echo ""
pause

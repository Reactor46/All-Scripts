Write-Host "Restoring Data from N Drive" -ForegroundColor Yellow
robocopy  \\misfs1\wreeves\Documents C:\Users\wreeves\Documents /MIR /FFT /Z /XA:H /W:5

Write-Host "Installing Chocolatey" -ForegroundColor Yellow
iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))

Write-Host "Installing Packages" -ForegroundColor Yellow
$Packages = Get-Content $home\Documents\ChocoPackages.txt
Foreach ( $Package in $Packages )
	{
	Write-Host "Installing $Package" -ForegroundColor Yellow
	cinst $Package -y
	}

Write-Host "Setting ConEmu at Startup" -ForegroundColor Yellow
cp "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\ConEmu\ConEmu (x64).lnk" "C:\Users\wreeves\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\"

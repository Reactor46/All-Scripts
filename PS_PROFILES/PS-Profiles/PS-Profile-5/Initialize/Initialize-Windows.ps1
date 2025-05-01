Push-Location HKCU:\Software\Microsoft\Windows\CurrentVersion\

Set-ItemProperty -Path Explorer\Advanced -Name TaskbarSmallIcons -Value 1

Pop-Location

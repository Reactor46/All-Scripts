#MTSetup.ps1

Write-Host "Checking Script path..."
 if(!(Test-Path 'C:\Scripts\Functions')){
    #If the path C:\Scripts\Functions  is not there create it
    if(!(Test-Path C:\Scripts)){
        New-Item -ItemType Directory -path ' C:\Scripts' -Force
        New-Item -ItemType Directory -Path ' C:\Scripts\Functions' -Force
     } Else { 
        New-Item -ItemType Directory -Path 'C:\Scripts\Functions' -Force
     }

 }

#Copy the Scrip over to the Funtion folder
    Write-Host 'Copying Files...'
    Copy-Item '\\lasfs02\winsys$\Scripts\MTRecycle\mtrecycle.ps1'  -Destination 'C:\Scripts\Functions'

#Check to see if PowerCLI is already Installed. If it is not then Install it
Write-Host ''
      if (!(Get-Module -ListAvailable -Name 'VMWare.PowerCLI')) {
            Write-Host 'Installing PowerCLI Modual'
            Install-Module -Name VMware.PowerCLI 
        }

        Write-Host 'Adding Modual to Local Profile.'
#Import the Function at Powershell Startup
$ProfPath = Get-Variable -Name profile
$ProfPath = $ProfPath.Value
$ProfPath
$CommandTxt = 'Import-Module C:\Scripts\Functions\MTRecycle.ps1 -force'

Add-Content -Path $ProfPath -Value $CommandTxt

        
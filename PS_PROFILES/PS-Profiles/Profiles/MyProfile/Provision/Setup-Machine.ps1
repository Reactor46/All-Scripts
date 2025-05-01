#Requires -RunAsAdministrator

Set-ExecutionPolicy Bypass -Scope Process -Force
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))

# Add windows defender exclusion folder
# C:\ProgramData\chocolatey\lib, C:\src as well

# Tools instalation
choco install -s=chocolatey powershell-core googlechrome firacode less microsoft-windows-terminal vscode git gitextensions kdiff3 windirstat beyondcompare gimp ilspy sysinternals mp3tag far mls-software-openssh -y

# linqpad5 CRC broken, find alternatives, for now we are taking it from OneDrive and MyProfile
#choco install linqpad5 -y

# Visual studio
choco install -s=chocolatey visualstudio2019enterprise visualstudio2019-workload-azure visualstudio2019-workload-manageddesktop visualstudio2019-workload-nativedesktop visualstudio2019-workload-netcoretools -y


# MS specific
if( $MsSpecific )
{
    choco install -s=chocolatey microsoft-teams -y
}

# Home specific
if( $HomeSpecific )
{
    choco install -s=chocolatey -y steam geforce-game-ready-driver
    # OneDrive setup - unlink and relink to change folder
}

# Console setup
git clone https://github.com/FallenGameR/MyProfile.git (Split-Path $profile)

$modulesFolder = $env:PSModulePath -split ";" | select -f 1
mkdir $modulesFolder -ea Ignore
cd $modulesFolder
git clone https://github.com/microsoft/PSToolset

# Mouse setup
& "$env:OneDriveConsumer\Distrib\Hardware\Deft Pro Trackball\mouse_driver_ma5111000.exe"

# Git setup - username, diff, merge via GitExtensions

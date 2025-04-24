#Enable-WindowsOptionalFeature -Online -FeatureName containers –All
#Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V –All
# Run Windows Containers/Docker Images on Windows Server 2019
Install-Module -Name DockerMsftProvider -Repository PSGallery -Force
Install-Package -Name docker -ProviderName DockerMsftProvider -Force

# Restart the local server
# Restart-Computer -Force

# Verify the Docker Package is installed
Get-Package -Name Docker -ProviderName DockerMsftProvider
Invoke-WebRequest "https://github.com/docker/compose/releases/download/v2.18.1/docker-compose-windows-x86_64.exe" -UseBasicParsing -OutFile $Env:ProgramFiles\Docker\docker-compose.exe

# Check docker version
docker version

# Start and check the Docker Service
Start-Service -Name Docker
Get-Service -Name Docker

# Pull image from Microsoft Docker Hub **Note: 1809 is the only version compatible with Server 2019 for this test
# docker pull mcr.microsoft.com/dotnet/samples:dotnetapp-nanoserver-1809
# Run image to test
# docker run mcr.microsoft.com/dotnet/samples:dotnetapp-nanoserver-1809

# Pull ltsc2019 ServerCore
# docker pull mcr.microsoft.com/windows/servercore:ltsc2019

# Setup Protainer
docker volume create portainer_data
# Portainer Business Edtion
# docker run -d -p 8000:8000 -p 9443:9443 --name portainer --restart always -v \\.\pipe\docker_engine:\\.\pipe\docker_engine -v portainer_data:C:\data portainer/portainer-ee:latest
# Portainer Community Edtion
docker run -d -p 8000:8000 -p 9000:9000 --name portainer --restart always -v \\.\pipe\docker_engine:\\.\pipe\docker_engine -v portainer_data:C:\data portainer/portainer-ce:latest


# Open Portainer Admin  
# Comment Out the browser(s) you do not wish to use
# MSEdge Browser
[system.Diagnostics.Process]::Start("msedge","http://localhost:9000")
# admin | ZfTwsE9fO8V8RyR
# Chrome Browser
# [system.Diagnostics.Process]::Start("chrome","https://localhost:9443")
# Firefox Browser
# [system.Diagnostics.Process]::Start("firefox","https://localhost:9443")


# Start the containers
# docker-compose.exe up --detach
# Stop all containers
# docker container stop $(docker container list -q)

 

function Exit-WACCOContainer {
<#

.SYNOPSIS
End the process of a container

.DESCRIPTION
End the process of a container

.ROLE
Administrators

#>
Param([string]$containerId)
$command = "docker container kill $containerId"

PowerShell -NoProfile -Windowstyle Hidden {
  Param(
    [string]$command
  )
  Invoke-Expression $command
}-args $command

}
## [END] Exit-WACCOContainer ##
function Get-WACCOAksDeploymentConfig {
<#

.SYNOPSIS
Retreives deployment details from a yaml file.

.DESCRIPTION
Retreives deployment details from a yaml file.

.ROLE
Administrators

#>

param(
  [string]$yamlFile
)

$yamlContent = Get-Content -Path $yamlFile
$encodedContent = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($yamlContent))

$encodedContent

}
## [END] Get-WACCOAksDeploymentConfig ##
function Get-WACCOAksDeployments {
<#

.SYNOPSIS
Retreives deployment details in a Kubernetes cluster.

.DESCRIPTION
Retreives deployment details in a Kubernetes cluster. Uses a temporary config.

.ROLE
Administrators

#>

param(
  [string]$kubeconfig
)

PowerShell -NoProfile -Windowstyle Hidden {
  param(
    [string]$kubeconfig
  )
  $basePath = Join-Path $env:TEMP ".kube";

  if (!(Test-Path $basePath)) {
    New-Item $basePath -ItemType Directory | Out-Null
  }

  $configPath = Join-Path $basePath "config";
  Set-Content -Path $configPath -Value $kubeconfig
  kubectl get deployments -o json --kubeconfig $configPath | ConvertFrom-Json

  Remove-Item $configPath -Force
} -args $kubeconfig

}
## [END] Get-WACCOAksDeployments ##
function Get-WACCOAksHciK8sCluster {
<#

.SYNOPSIS
Get The AksHci kuberbetes clusters

.DESCRIPTION
Get The AksHci kuberbetes clusters

.ROLE
Readers

#>
PowerShell -NoProfile -Windowstyle Hidden {
  Import-Module AksHci -ErrorAction SilentlyContinue

  $AksClusters = @();

  $clusters = Get-AksHciCluster

  Foreach ($cluster in $clusters) {
    $clusterInfo = @{
      name              = $cluster.name;
      version           = $cluster.kubernetesversion;
      controlPlanes     = $cluster.controlplanenodecount;
      linuxWorkers      = $cluster.linuxnodecount;
      windowsWorkers    = $cluster.windowsnodecount;
      provisioningState = $cluster.provisioningState;
    }
    $managementCluster = (Get-AksHciConfig).Kva.KvaName;
    if ($clusterInfo.name -ne $managementCluster) {
      $AksClusters += $clusterInfo;
    }
  }

  $AksClusters
}

}
## [END] Get-WACCOAksHciK8sCluster ##
function Get-WACCOAksHciK8sCredential {
<#

.SYNOPSIS
Get The AksHci kubeconfig credentials

.DESCRIPTION
Get The AksHci kubeconfig credentials

.ROLE
Readers

#>
param (
  [string]$clusterName
)

PowerShell -NoProfile -Windowstyle Hidden {

  param (
    [string]$clusterName
  )

  $basePath = Join-Path $env:TEMP ".kube"

  if (!(Test-Path $basePath)) {
    New-Item $basePath -ItemType Directory | Out-Null
  }

  $configPath = Join-Path $basePath "$clusterName-kubeconfig"
  Get-AksHciCredential -name $clusterName -configPath $basePath -Confirm:$false
  $EncodedKubeconfigData = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($configPath))
  Remove-Item $configPath -Force

  $EncodedKubeconfigData
} -args $clusterName

}
## [END] Get-WACCOAksHciK8sCredential ##
function Get-WACCOAuthenticatedRegistries {
<#

.SYNOPSIS
Get Authenticated Registries

.DESCRIPTION
Get Authenticated Registries

.ROLE
Readers

#>
function getAuthenticatedRegistries() {
  $registries = New-Object System.Collections.Generic.List[System.Object]
  $file_path = $home + "\.docker\config.json"
  $content = Get-Content -Path $file_path -ErrorAction SilentlyContinue | out-string
  if ($content.Contains('auths')) {
    $jsonContent = $content | ConvertFrom-Json
    foreach ($c in $jsonContent.auths.PSObject.Properties ) {
      $registries.Add($c.Name)
    }
  }
  return $registries
}
getAuthenticatedRegistries

}
## [END] Get-WACCOAuthenticatedRegistries ##
function Get-WACCOAzureRegistryToken {
<#

.SYNOPSIS
Get The Azure Container Registry Token

.DESCRIPTION
Get The Azure Container Registry Token

.ROLE
Administrators

#>
param (
  [string]$registryUrl,
  [string]$body,
  [string]$contentType
)

$processedBody = $body | ConvertFrom-StringData

$response = Invoke-WebRequest -Uri $registryUrl -body $processedBody -ContentType $contentType -Method 'POST' -UseBasicParsing

if ($response -and ($response.StatusCode -eq 200)) {
  $jsonContent = $response.Content | ConvertFrom-Json
  $azureRegistryToken = $jsonContent.access_token
  Write-Output $azureRegistryToken
}

}
## [END] Get-WACCOAzureRegistryToken ##
function Get-WACCOCimWin32OperatingSystem {
<#

.SYNOPSIS
Gets Win32_OperatingSystem object.

.DESCRIPTION
Gets Win32_OperatingSystem object.

.ROLE
Readers

#>

##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_OperatingSystem

}
## [END] Get-WACCOCimWin32OperatingSystem ##
function Get-WACCOContainerLogs {
<#

.SYNOPSIS
Get logs of a container

.DESCRIPTION
Get logs of a container

.ROLE
Readers

#>
Param([string]$containerId)

PowerShell -NoProfile -Windowstyle Hidden {
  Param([string]$containerId)
  docker logs -t $containerId
} -args $containerId

}
## [END] Get-WACCOContainerLogs ##
function Get-WACCOContainers {
<#

.SYNOPSIS
Get containers

.DESCRIPTION
Get containers

.ROLE
Readers

#>

PowerShell -NoProfile -Windowstyle Hidden {
  $containerList = docker ps -a -s --no-trunc --format='{{.Names}}#{{.Image}}#{{.ID}}#{{.Status}}#{{.Ports}}#{{.CreatedAt}}#{{.Size}}'
  $containers = @()
  $format = "{{.State.Pid}}#{{.Driver}}#{{.Image}}"`
    + "#{{range `$key, `$value := .NetworkSettings.Networks}}{{`$key}}({{`$value.IPAddress}}) {{end}}"`
    + "#{{range `$value := .Mounts}}{{`$value.Name}}({{`$value.Type}}) {{end}}#{{.HostConfig.Isolation}}"
  foreach ($containerItem in $containerList) {
    $container = @{
      name        = ''
      created     = ''
      status      = ''
      image       = ''
      command     = ''
      ports       = ''
      containerId = ''
      imageFullId = ''
      pid         = ''
      driver      = ''
      networks    = ''
      volumes     = ''
      isolation   = ''
      size        = ''
    }

    $containerItems = $containerItem -split '#'
    $container.name = $containerItems[0]
    $container.image = $containerItems[1]
    $container.containerId = $containerItems[2]
    $container.status = $containerItems[3]
    $container.ports = $containerItems[4]
    $container.created = $containerItems[5]
    $container.size = $containerItems[6]

    $containerDetail = docker inspect $container.containerId --format=$format
    $containerDetails = $containerDetail -split '#'
    $container.pid = $containerDetails[0]
    $container.driver = $containerDetails[1]
    $container.imageFullId = $containerDetails[2]
    $container.networks = $containerDetails[3]
    $container.volumes = $containerDetails[4]
    $container.isolation = $containerDetails[5]

    # for those could contain delimiters
    $container.command = docker inspect $container.containerId --format="{{.Config.Cmd}}"

    $containers += $container
  }
  $containers
}

}
## [END] Get-WACCOContainers ##
function Get-WACCOContainersToolAvailability {
<#

.SYNOPSIS
Get if container tool is available

.DESCRIPTION
Get if container tool is available

.ROLE
Readers

#>
$dockerInstalled = $false
$dockerServiceRunning = $false
$kubectlInstalled = $false

# No need to check Containers feature since installing docker will automatically enable it
$cmd = Get-Command 'docker' -ErrorAction SilentlyContinue
$dockerInstalled = $cmd -and $cmd.Name -eq 'docker.exe'
if ($dockerInstalled) {
    $service = Get-Service 'docker' -ErrorAction SilentlyContinue
    $dockerServiceRunning = $service -and $service.Status -eq 'Running'
}

$kubeCtlCmd = Get-Command 'kubectl' -ErrorAction SilentlyContinue
$kubectlInstalled = $kubeCtlCmd -and $kubeCtlCmd.Name -eq 'kubectl.exe'

return @{
    DockerInstalled = $dockerInstalled
    DockerServiceRunning = $dockerServiceRunning
    KubectlInstalled = $kubectlInstalled
}

}
## [END] Get-WACCOContainersToolAvailability ##
function Get-WACCODockerfileText {
<#

.SYNOPSIS
Get Dockerfile text.

.DESCRIPTION
Get the source code as text from the Dockerfile in the parent_dir

.ROLE
Readers

#>
Param(
  [string]$parent_dir
)

function loadDockerfileContent($parent_dir) {
  [string] $file_path = (Join-Path $parent_dir 'Dockerfile')
  return Get-Content -Path $file_path
}

loadDockerfileContent $parent_dir

}
## [END] Get-WACCODockerfileText ##
function Get-WACCOGridStats {
<#

.SYNOPSIS
Get stats displayed in grid

.DESCRIPTION
Get stats displayed in grid

.ROLE
Readers

#>
PowerShell -NoProfile -Windowstyle Hidden {
  $stats = @{}
  $statsList = docker stats -a --no-stream --no-trunc --format='{{.ID}}#{{.CPUPerc}}${{.MemUsage}}'
  foreach ($statsItem in $statsList) {
    $statsItems = $statsItem -split '#'
    $stats.add($statsItems[0], $statsItems[1])
  }
  $stats
}

}
## [END] Get-WACCOGridStats ##
function Get-WACCOImages {
<#

.SYNOPSIS
Get images

.DESCRIPTION
Get images

.ROLE
Readers

#>
PowerShell -NoProfile -Windowstyle Hidden {
  $imageList = docker image ls --no-trunc --format='{{.Repository}}#{{.Tag}}#{{.Size}}#{{.ID}}#{{.CreatedAt}}'
  $images = @()
  foreach ($imageItem in $imageList) {
    $image = @{
      repo          = ''
      tag           = ''
      size          = ''
      id            = ''
      created       = ''
      command       = ''
      entryPoint    = ''
      dockerVersion = ''
      os            = ''
      architecture  = ''
      parent        = ''
    }

    $imageItems = $imageItem -split '#'
    $image.repo = $imageItems[0]
    $image.tag = $imageItems[1]
    $image.size = $imageItems[2]
    $image.id = $imageItems[3]
    $image.created = $imageItems[4]

    $imageDetail = docker inspect $image.id --format='{{.DockerVersion}}#{{.Os}}#{{.Architecture}}#{{.Parent}}'
    $imageDetails = $imageDetail -split '#'
    $image.dockerVersion = $imageDetails[0]
    $image.os = $imageDetails[1]
    $image.architecture = $imageDetails[2]
    $image.parent = $imageDetails[3]

    # for those could contain delimiters
    $image.command = docker inspect $image.id --format='{{.Config.Cmd}}'
    $image.entryPoint = docker inspect $image.id --format='{{.Config.Entrypoint}}'

    $images += $image
  }
  $images
}

}
## [END] Get-WACCOImages ##
function Get-WACCOKubeletProcess {
<#

.SYNOPSIS
Checks if kubelet.exe is present

.DESCRIPTION
Checks if kubelet.exe is present

.ROLE
Administrators

#>

$k8Service = Get-Service kubelet -ErrorAction SilentlyContinue

$kubeletPresent  = $null -ne $k8Service

$kubeletPresent

}
## [END] Get-WACCOKubeletProcess ##
function Get-WACCONetworks {
<#

.SYNOPSIS
Get Networks

.DESCRIPTION
Get Networks

.ROLE
Readers

#>

PowerShell -NoProfile -Windowstyle Hidden {
  $IDs = docker network ls --no-trunc --format='{{.ID}}'
  $networks = @()
  $format = "{{.Driver}}#{{.Scope}}#{{range `$value := .IPAM.Config}}{{`$value.Subnet}} {{end}}#{{range `$value := .IPAM.Config}}{{`$value.Gateway}} {{end}}"
  foreach ($id in $IDs) {
    $network = @{
      name    = ''
      id      = $id
      driver  = ''
      scope   = ''
      subnet  = ''
      gateway = ''
    }

    $networkItem = docker inspect $id --format=$format
    $networkItems = $networkItem -split '#'
    $network.driver = $networkItems[0]
    $network.scope = $networkItems[1]
    $network.subnet = $networkItems[2]
    $network.gateway = $networkItems[3]

    # for those could contain delimiters
    $network.name = docker inspect $id --format='{{.Name}}'

    $networks += $network
  }
  $networks
}

}
## [END] Get-WACCONetworks ##
function Get-WACCORefreshToken {
<#

.SYNOPSIS
Get a Refresh Token For Azure Container Registry

.DESCRIPTION
Get a Refresh Token For Azure Container Registry

.ROLE
Administrators

#>
param (
  [string]$registryUrl,
  [string]$body,
  [string]$contentType
)

$processedBody = $body | ConvertFrom-StringData

$response = Invoke-WebRequest -Uri $registryUrl -body $processedBody -ContentType $contentType -Method 'POST' -UseBasicParsing
if ($response -and ($response.StatusCode -eq 200)) {
  $jsonContent = $response.Content | ConvertFrom-Json
  $refresh_token = $jsonContent.refresh_token
  Write-Output $refresh_token
}

}
## [END] Get-WACCORefreshToken ##
function Get-WACCORegistryLogin {
<#

.SYNOPSIS
Login to a registry

.DESCRIPTION
Login to a registry

.ROLE
Administrators

#>
param (
  [string]$registryUrl,
  [string]$registryUsername,
  [string]$registryPswd
)

PowerShell -NoProfile -Windowstyle Hidden {

  param (
    [string]$registryUrl,
    [string]$registryUsername,
    [string]$registryPswd
  )
  function login($registryUrl, $registryUsername, $registryPswd) {
    Write-Output $registryPswd | docker login $registryUrl --username $registryUsername --password-stdin
  }

  login $registryUrl $registryUsername $registryPswd
} -args $registryUrl $registryUsername $registryPswd

}
## [END] Get-WACCORegistryLogin ##
function Get-WACCORegistryRequest {
<#

.SYNOPSIS
Get The Azure Container Registry HTTP Requests

.DESCRIPTION
Get The Azure Container Registry HTTP Requests

.ROLE
Administrators

#>

param (
  [Parameter(Mandatory = $true)]
  [string]$url,
  [Parameter(Mandatory = $true)]
  [string]$token,
  [Parameter(Mandatory = $true)]
  [string]$method
)

function registryRequest($url, $token, $method) {

  $response = Invoke-WebRequest -Uri $url -Headers @{ 'Authorization' = 'Bearer ' + $token } -Method $method -UseBasicParsing
  $jsonContent = $response.Content | ConvertFrom-Json

  return $jsonContent
}

registryRequest $url $token $method

}
## [END] Get-WACCORegistryRequest ##
function Get-WACCOStats {
<#

.SYNOPSIS
Get stats for charts

.DESCRIPTION
Get stats for charts

.ROLE
Readers

#>
Param([string]$containerId)

PowerShell -NoProfile -Windowstyle Hidden {
  Param([string]$containerId)
  docker stats --no-stream --format='{{.CPUPerc}}${{.MemUsage}}${{.NetIO}}' $containerId
} -args $containerId

}
## [END] Get-WACCOStats ##
function Get-WACCOSummary {
<#

.SYNOPSIS
Get summary

.DESCRIPTION
Get summary

.ROLE
Readers

#>

PowerShell -NoProfile -Windowstyle Hidden {
  $containerList = docker container list -a --no-trunc --format='{{.Status}}'
  $totalContainersCount = 0
  $runningContainersCount = 0
  $stoppedContainersCount = 0
  foreach ($container in $containerList) {
    if ($container -like '*Up*') {
      $runningContainersCount++
    }
    else {
      $stoppedContainersCount++
    }
    $totalContainersCount++
  }

  $imageList = docker image ls --no-trunc --format='{{.ID}}#{{.Tag}}'
  $totalImagesCount = 0
  $latestImagesCount = 0
  $totalImagesSize = 0
  $imageMap = @{}
  foreach ($image in $imageList) {
    $imageItems = $image -split '#'
    $id = $imageItems[0]
    if (-not ($imageMap[$id] -eq 1)) {
      $size = docker inspect $id --format='{{.Size}}'
      $size /= 1000000000
      $size = [Math]::Round($size, 2)
      $totalImagesSize += $size
      $imageMap.add($id, 1)
    }
    $tag = $imageItems[1]
    if ($tag -eq 'latest') {
      $latestImagesCount++
    }
    $totalImagesCount++
  }
  $totalImagesSize = [string]$totalImagesSize + "GB"

  $localNetworksCount = 0
  $otherNetworksCount = 0
  $totalNetworksCount = 0
  $networkList = docker network ls --no-trunc --format='{{.Scope}}'
  foreach ($network in $networkList) {
    if ($network -eq "local") {
      $localNetworksCount++
    }
    else {
      $otherNetworksCount++
    }
    $totalNetworksCount++
  }

  $localVolumesCount = 0
  $otherVolumesCount = 0
  $totalVolumesCount = 0
  $volumeList = docker volume ls --format='{{.Scope}}'
  foreach ($volume in $volumeList) {
    if ($volume -eq "local") {
      $localVolumesCount++
    }
    else {
      $otherVolumesCount++
    }
    $totalVolumesCount++
  }

  $summary = @{
    totalContainersCount   = $totalContainersCount
    runningContainersCount = $runningContainersCount
    stoppedContainersCount = $stoppedContainersCount
    totalImagesCount       = $totalImagesCount
    latestImagesCount      = $latestImagesCount
    totalImagesSize        = $totalImagesSize
    totalNetworksCount     = $totalNetworksCount
    localNetworksCount     = $localNetworksCount
    otherNetworksCount     = $otherNetworksCount
    totalVolumesCount      = $totalVolumesCount
    localVolumesCount      = $localVolumesCount
    otherVolumesCount      = $otherVolumesCount
  }

  $summary
}

}
## [END] Get-WACCOSummary ##
function Get-WACCOVolumes {
<#

.SYNOPSIS
Get volumes

.DESCRIPTION
Get volumes

.ROLE
Readers

#>

PowerShell -NoProfile -Windowstyle Hidden {
  $volumeList = docker volume ls --format='{{.Name}}#{{.Driver}}#{{.Scope}}'
  $volumes = @()
  foreach ($volumeItem in $volumeList) {
    $volume = @{
      name       = ''
      driver     = ''
      scope      = ''
      mountPoint = ''
    }

    $volumeItems = $volumeItem -split '#'
    $volume.name = $volumeItems[0]
    $volume.driver = $volumeItems[1]
    $volume.scope = $volumeItems[2]

    # for those could contain delimiters
    $volume.mountPoint = docker inspect $volume.name --format='{{.Mountpoint}}'

    $volumes += $volume
  }
  $volumes
}

}
## [END] Get-WACCOVolumes ##
function Get-WACCOVsSolutionProperties {
<#

.SYNOPSIS
Get Visual Studio Solution properties
.

.DESCRIPTION
Get text in a Visual Studio Solution file to extract properties

.ROLE
Readers

#>

Param(
  [string]$parent_dir
)

function loadVsSolutionFile($parent_dir){

  [string] $file_path = (Join-Path $parent_dir '*.sln')
  return Get-Content -Path $file_path

}

loadVsSolutionFile $parent_dir


}
## [END] Get-WACCOVsSolutionProperties ##
function Install-WACCODocker {
<#

.SYNOPSIS
Installs Docker on Host

.DESCRIPTION
Installs Docker on Host

.ROLE
Administrators

#>

# Avoid restart if an error occured with installation commands
$ErrorActionPreference = "Stop"

function installFeatures() {
  # Nuget provider is required before installing docker
  Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force

  Install-Module -Name DockerMsftProvider -Repository PSGallery -Force

  Install-Package -Name docker -ProviderName DockerMsftProvider -Force
}
function main() {
  try {
    installFeatures
    Restart-Computer -Force
  }
  catch {
    if ($error[0].ErrorDetails) {
      throw $error[0].ErrorDetails
    }
    throw $error[0]
  }
}

main

}
## [END] Install-WACCODocker ##
function Install-WACCOKubectl {
<#

.SYNOPSIS
Installs Kubectl on Host

.DESCRIPTION
Installs Kubectl on Host

.ROLE
Administrators

#>


Param(
    [string]
    $Downloadlocation = $env:TEMP
)

function main () {
  $InitialProtocol = [Net.SecurityProtocolType]::[Net.ServicePointManager]::SecurityProtocol
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

  $DownloadLocation = $ExecutionContext.InvokeCommand.ExpandString($DownloadLocation)
  $Webresponse = Invoke-WebRequest "https://storage.googleapis.com/kubernetes-release/release/stable.txt" -UseBasicParsing
  $Source = "https://storage.googleapis.com/kubernetes-release/release/$($Webresponse.content)/bin/windows/amd64/kubectl.exe"

  try {
    Start-BitsTransfer -Source $Source -Destination $Downloadlocation
    $Downloadfile = Join-Path $Downloadlocation "kubectl.exe"
    Unblock-File $Downloadfile
    .$Downloadfile version
    $Kube_Local = New-Item -ItemType directory "$($HOME)/.kube" -force

    # clean up
    [Net.ServicePointManager]::SecurityProtocol = $InitialProtocol
    Remove-Item $Downloadfile -Force
  }
  catch {
    if ($error[0].ErrorDetails) {
      throw $error[0].ErrorDetails
    }
    throw $error[0]
  }
}

main

}
## [END] Install-WACCOKubectl ##
function Invoke-WACCODockerOauthLogin {
<#

.SYNOPSIS
Login to a registry

.DESCRIPTION
Login to a registry

.ROLE
Administrators

#>

param (
  [string]$accessToken,
  [string]$service,
  [string]$tenant,
  [string]$contentType,
  [string]$registryUrl,
  [string]$grantType)

PowerShell -NoProfile -Windowstyle Hidden {

  param (
    [string]$accessToken,
    [string]$service,
    [string]$tenant,
    [string]$contentType,
    [string]$registryUrl,
    [string]$grantType)
  function invokeDockerOauthLogin($accessToken, $service, $tenant, $contentType, $registryUrl, $grantType) {
    $body = @{
      grant_type   = $grantType
      service      = $service
      tenant       = $tenant
      access_token = $accessToken
    }

    $result = Invoke-WebRequest -Uri $registryUrl -body $body -ContentType $contentType -Method 'POST'
    if ($result -and ($result.StatusCode -eq 200)) {
      $jsonContent = $result.Content | ConvertFrom-Json
      $refresh_token = $jsonContent.refresh_token
      $username = '00000000-0000-0000-0000-000000000000'
      dockerLogin $service $username $refresh_token
    }
  }
  function dockerLogin($registry, $username, $passwd) {
    Write-Output $passwd | docker login $registry --username $username --password-stdin
  }

  invokeDockerOauthLogin $accessToken $service $tenant $contentType $registryUrl $grantType

} -args $accessToken $service $tenant $contentType $registryUrl $grantType

}
## [END] Invoke-WACCODockerOauthLogin ##
function Invoke-WACCOKubernetesDeployment {
<#

.SYNOPSIS
Create a deployment in a Kubernetes cluster.

.DESCRIPTION
Create a deployment in a Kubernetes cluster. Uses a temporary config.

.ROLE
Administrators

#>

param(
  [Parameter(Mandatory = $true)]
  [string]$kubeconfig,
  [Parameter(Mandatory = $true)]
  [string]$yamlContent,
  [Parameter(Mandatory = $false)]
  [string]$token,
  [Parameter(Mandatory = $false)]
  [string]$secret,
  [Parameter(Mandatory = $false)]
  [string]$registry
)

$basePath = Join-Path $env:TEMP ".kube";
$username = "00000000-0000-0000-0000-000000000000"

if (!(Test-Path $basePath)) {
  New-Item $basePath -ItemType Directory | Out-Null
}

$configPath = Join-Path $basePath "config";
Set-Content -Path $configPath -Value $kubeconfig

$yamlPath = Join-Path $basePath "deployment.yaml";
Set-Content -Path $yamlPath -Value $yamlContent

$command = "kubectl create secret docker-registry $secret --docker-server=$registry --docker-username=$username --docker-password=$token --kubeconfig $configPath"
Invoke-Expression $command 2>&1

kubectl apply -f $yamlPath  --kubeconfig $configPath
Remove-Item "$basePath\*"  -Force

}
## [END] Invoke-WACCOKubernetesDeployment ##
function Invoke-WACCOPullImages {
<#

.SYNOPSIS
Pull images.

.DESCRIPTION
Pull images from a registry

.ROLE
Readers

#>
Param(
  [string]$imageName,
  [string]$imageTag,
  [boolean]$allTaggedImages
)
function pullImages($imageName, $imageTag, $allTaggedImages) {
  $command = '';
  if ($allTaggedImages) {
    $command = "docker pull $imageName --all-tags"
  }
  else {
    $fullName = $imageName
    if ($imageTag) {
      $fullName += ":" + $imageTag
    }
    $command = "docker pull $fullName"
  }

  PowerShell -NoProfile -Windowstyle Hidden {
    Param(
      [string]$command
    )
    Invoke-Expression $command
  }-args $command
}

pullImages $imageName $imageTag $allTaggedImages

}
## [END] Invoke-WACCOPullImages ##
function Invoke-WACCOPushImage {
<#

.SYNOPSIS
Push images.

.DESCRIPTION
Push images from a registry

.ROLE
Administrators

#>
Param(
  [string]$imageRegistry,
  [string]$imageName,
  [string]$imageTag,
  [string]$selectedImageName
)


PowerShell -NoProfile -Windowstyle Hidden {
  Param(
    [string]$imageRegistry,
    [string]$imageName,
    [string]$imageTag,
    [string]$selectedImageName
  )
  function pushImage($imageRegistry, $imageName, $imageTag, $selectedImageName) {
    $newImage = $imageRegistry + "/" + $imageName + ":" + $imageTag
    docker tag $selectedImageName $newImage
    docker push $newImage
  }

  pushImage $imageRegistry $imageName $imageTag $selectedImageName

} -args $imageRegistry $imageName $imageTag $selectedImageName

}
## [END] Invoke-WACCOPushImage ##
function Invoke-WACCORunImage {
<#

.SYNOPSIS
Run an image

.DESCRIPTION
Runs an image using docker run

.ROLE
Administrators

#>
Param(
  [Parameter(Mandatory = $true)]
  [string]$imageName,
  [string]$tag,
  [string]$portMapping,
  [string]$publishAll,
  [string]$containerName,
  [string]$memory,
  [string]$cpus,
  [string]$hypervIsolation,
  [string[]]$additionalParams
)
$fullName = $imageName
$parameterList = @()
if ($tag) {
  $fullName += ":" + $tag
}

if ($containerName) {
  $parameterList += '--name'
  $parameterList += $containerName
}

if ($publishAll -eq 'True') {
  $parameterList += '-P'
}

if ($hypervIsolation -eq 'True') {
  $parameterList += '--isolation'
  $parameterList += 'hyperv'
}

if ($memory) {
  $parameterList += '--memory'
  $parameterList += $memory
}

if ($cpus) {
  $parameterList += '--cpus'
  $parameterList += $cpus
}

if ($portMapping) {
  $parameterList += "--publish"
  $parameterList += $portMapping
}

$parameterList += $additionalParams;

$parameterList += $fullName

$command = "docker run -d $parameterList"

PowerShell -NoProfile -Windowstyle Hidden {
  Param(
    [string]$command
  )
  Invoke-Expression $command
} -args $command

}
## [END] Invoke-WACCORunImage ##
function New-WACCOImage {
<#

.SYNOPSIS
Create image.

.DESCRIPTION
Create docker image from the generated Dockerfile text

.ROLE
Readers

#>
Param(
  [string]$imageName,
  [string]$artifactPath,
  [bool]$buildOnly,
  [string]$sourceCode,
  [string]$scriptPath
)

PowerShell -NoProfile -Windowstyle Hidden {
  Param(
    [string]$imageName,
    [string]$artifactPath,
    [bool]$buildOnly,
    [string]$sourceCode,
    [string]$scriptPath
  )

  function buildDockerImage($imageName, $artifactPath, $buildOnly, $dockerText, $scriptPath) {
    if (($null -ne $scriptPath) -and ($scriptPath -ne '')) {
      if ((Test-Path $scriptPath) -and (Test-Path $artifactPath)) {
        Copy-Item -Path $scriptPath -Destination $artifactPath
      }
    }

    if (!$buildOnly) {
      [string] $file_path = (Join-Path $artifactPath 'Dockerfile')
      if (!(Test-Path $file_path)) {
        New-Item -ItemType File -Path $file_path
      }

      Set-Content -Path $file_path -Value $dockerText
    }

    Set-Location $artifactPath
    docker build -t "$imageName" .
  }

  buildDockerImage $imageName $artifactPath $buildOnly $sourceCode $scriptPath
} -args $imageName, $artifactPath, $buildOnly, $sourceCode, $scriptPath

}
## [END] New-WACCOImage ##
function Remove-WACCOContainer {
<#

.SYNOPSIS
Delete a container

.DESCRIPTION
Delete a container

.ROLE
Administrators

#>
Param([string]$containerId)

PowerShell -NoProfile -Windowstyle Hidden {
  Param([string]$containerId)
  docker rm $containerId
} -args $containerId

}
## [END] Remove-WACCOContainer ##
function Remove-WACCOImage {
<#

.SYNOPSIS
Delete an image

.DESCRIPTION
Delete an image

.ROLE
Administrators

#>
Param([string]$id)

PowerShell -NoProfile -Windowstyle Hidden {
  Param([string]$id)
  docker rmi -f $id
} -args $id

}
## [END] Remove-WACCOImage ##
function Start-WACCODocker {
<#

.SYNOPSIS
Starts the Docker Service

.DESCRIPTION
Starts the Docker Service

.ROLE
Administrators

#>

# net start docker works for both servers and clients
net start docker

}
## [END] Start-WACCODocker ##
function Add-WACCOFolderShare {
<#

.SYNOPSIS
Gets a new share name for the folder.

.DESCRIPTION
Gets a new share name for the folder. It starts with the folder name. Then it keeps appending "2" to the name
until the name is free. Finally return the name.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Path
    String -- The path to the folder to be shared.

.PARAMETER Name
    String -- The suggested name to be shared (the folder name).

.PARAMETER Force
    boolean -- override any confirmations

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path,    

    [Parameter(Mandatory = $true)]
    [String]
    $Name
)

Set-StrictMode -Version 5.0

while([bool](Get-SMBShare -Name $Name -ea 0)){
    $Name = $Name + '2';
}

New-SmbShare -Name "$Name" -Path "$Path"
@{ shareName = $Name }

}
## [END] Add-WACCOFolderShare ##
function Add-WACCOFolderShareNameUser {
<#

.SYNOPSIS
Adds a user to the folder share.

.DESCRIPTION
Adds a user to the folder share.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Name
    String -- Name of the share.

.PARAMETER AccountName
    String -- The user identification (AD / Local user).

.PARAMETER AccessRight
    String -- Access rights of the user.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Name,

    [Parameter(Mandatory = $true)]
    [String]
    $AccountName,

    [Parameter(Mandatory = $true)]
    [String]
    $AccessRight
)

Set-StrictMode -Version 5.0

Grant-SmbShareAccess -Name "$Name" -AccountName "$AccountName" -AccessRight "$AccessRight" -Force


}
## [END] Add-WACCOFolderShareNameUser ##
function Add-WACCOFolderShareUser {
<#

.SYNOPSIS
Adds a user access to the folder.

.DESCRIPTION
Adds a user access to the folder.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Path
    String -- The path to the folder.

.PARAMETER Identity
    String -- The user identification (AD / Local user).

.PARAMETER FileSystemRights
    String -- File system rights of the user.

.PARAMETER AccessControlType
    String -- Access control type of the user.    

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path,

    [Parameter(Mandatory = $true)]
    [String]
    $Identity,

    [Parameter(Mandatory = $true)]
    [String]
    $FileSystemRights,

    [ValidateSet('Deny','Allow')]
    [Parameter(Mandatory = $true)]
    [String]
    $AccessControlType
)

Set-StrictMode -Version 5.0

function Remove-UserPermission
{
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $Path,
    
        [Parameter(Mandatory = $true)]
        [String]
        $Identity,
        
        [ValidateSet('Deny','Allow')]
        [Parameter(Mandatory = $true)]
        [String]
        $ACT
    )

    $Acl = Get-Acl $Path
    $AccessRule = New-Object system.security.accesscontrol.filesystemaccessrule($Identity, 'ReadAndExecute','ContainerInherit, ObjectInherit', 'None', $ACT)
    $Acl.RemoveAccessRuleAll($AccessRule)
    Set-Acl $Path $Acl
}

If ($AccessControlType -eq 'Deny') {
    $FileSystemRights = 'FullControl'
    Remove-UserPermission $Path $Identity 'Allow'
} else {
    Remove-UserPermission $Path $Identity 'Deny'
}

$Acl = Get-Acl $Path
$AccessRule = New-Object system.security.accesscontrol.filesystemaccessrule($Identity, $FileSystemRights,'ContainerInherit, ObjectInherit', 'None', $AccessControlType)
$Acl.AddAccessRule($AccessRule)
Set-Acl $Path $Acl

}
## [END] Add-WACCOFolderShareUser ##
function Clear-WACCOEventLogChannel {
<#

.SYNOPSIS
Clear the event log channel specified.

.DESCRIPTION
Clear the event log channel specified.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>
 
Param(
    [string]$channel
)

[System.Diagnostics.Eventing.Reader.EventLogSession]::GlobalSession.ClearLog("$channel") 
}
## [END] Clear-WACCOEventLogChannel ##
function Clear-WACCOEventLogChannelAfterExport {
<#

.SYNOPSIS
Clear the event log channel after export the event log channel file (.evtx).

.DESCRIPTION
Clear the event log channel after export the event log channel file (.evtx).
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>

Param(
    [string]$channel
)

$segments = $channel.Split("-")
$name = $segments[-1]

$randomString = [GUID]::NewGuid().ToString()
$ResultFile = $env:temp + "\" + $name + "_" + $randomString + ".evtx"
$ResultFile = $ResultFile -replace "/", "-"

wevtutil epl "$channel" "$ResultFile" /ow:true

[System.Diagnostics.Eventing.Reader.EventLogSession]::GlobalSession.ClearLog("$channel") 

return $ResultFile

}
## [END] Clear-WACCOEventLogChannelAfterExport ##
function Compress-WACCOArchiveFileSystemEntity {
<#

.SYNOPSIS
Compresses the specified file system entity (files, folders) of the system.

.DESCRIPTION
Compresses the specified file system entity (files, folders) of the system on this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER pathSource
    String -- The path to compress.

.PARAMETER PathDestination
    String -- The destination path to compress into.

.PARAMETER Force
    boolean -- override any confirmations

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $PathSource,    

    [Parameter(Mandatory = $true)]
    [String]
    $PathDestination,

    [Parameter(Mandatory = $false)]
    [boolean]
    $Force
)

Set-StrictMode -Version 5.0

if ($Force) {
    Compress-Archive -Path $PathSource -Force -DestinationPath $PathDestination
} else {
    Compress-Archive -Path $PathSource -DestinationPath $PathDestination
}
if ($error) {
    $code = $error[0].Exception.HResult
    @{ status = "error"; code = $code; message = $error }
} else {
    @{ status = "ok"; }
}

}
## [END] Compress-WACCOArchiveFileSystemEntity ##
function Disable-WACCOKdcProxy {
<#
.SYNOPSIS
Disables kdc proxy on the server

.DESCRIPTION
Disables kdc proxy on the server

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [string]
    $KdcPort
)

$urlLeft = "https://+:"
$urlRight = "/KdcProxy/"
$url = $urlLeft + $KdcPort + $urlRight
$deleteOutput = netsh http delete urlacl url=$url
if ($LASTEXITCODE -ne 0) {
    throw $deleteOutput
}
Remove-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "HttpsClientAuth"
Remove-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "DisallowUnprotectedPasswordAuth"
Stop-Service -Name kpssvc
Set-Service -Name kpssvc -StartupType Disabled
$firewallString = "KDC Proxy Server service (KPS) for SMB over QUIC"
Remove-NetFirewallRule -DisplayName $firewallString -ErrorAction SilentlyContinue
}
## [END] Disable-WACCOKdcProxy ##
function Disable-WACCOSmbOverQuic {
<#

.SYNOPSIS
Disables smb over QUIC on the server.

.DESCRIPTION
Disables smb over QUIC on the server.

.ROLE
Administrators

#>

Set-SmbServerConfiguration -EnableSMBQUIC $false -Force
}
## [END] Disable-WACCOSmbOverQuic ##
function Edit-WACCOFolderShareInheritanceFlag {
<#

.SYNOPSIS
Modifies all users' IsInherited flag to false

.DESCRIPTION
Modifies all users' IsInherited flag to false
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Path
    String -- The path to the folder.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path
)

Set-StrictMode -Version 5.0

$Acl = Get-Acl $Path
$Acl.SetAccessRuleProtection($True, $True)
Set-Acl -Path $Path -AclObject $Acl

}
## [END] Edit-WACCOFolderShareInheritanceFlag ##
function Edit-WACCOFolderShareUser {
<#

.SYNOPSIS
Edits a user access to the folder.

.DESCRIPTION
Edits a user access to the folder.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Path
    String -- The path to the folder.

.PARAMETER Identity
    String -- The user identification (AD / Local user).

.PARAMETER FileSystemRights
    String -- File system rights of the user.

.PARAMETER AccessControlType
    String -- Access control type of the user.    

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path,

    [Parameter(Mandatory = $true)]
    [String]
    $Identity,

    [Parameter(Mandatory = $true)]
    [String]
    $FileSystemRights,

    [ValidateSet('Deny','Allow')]
    [Parameter(Mandatory = $true)]
    [String]
    $AccessControlType
)

Set-StrictMode -Version 5.0

function Remove-UserPermission
{
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $Path,
    
        [Parameter(Mandatory = $true)]
        [String]
        $Identity,
        
        [ValidateSet('Deny','Allow')]
        [Parameter(Mandatory = $true)]
        [String]
        $ACT
    )

    $Acl = Get-Acl $Path
    $AccessRule = New-Object system.security.accesscontrol.filesystemaccessrule($Identity, 'ReadAndExecute','ContainerInherit, ObjectInherit', 'None', $ACT)
    $Acl.RemoveAccessRuleAll($AccessRule)
    Set-Acl $Path $Acl
}

If ($AccessControlType -eq 'Deny') {
    $FileSystemRights = 'FullControl'
    Remove-UserPermission $Path $Identity 'Allow'
} else {
    Remove-UserPermission $Path $Identity 'Deny'
}

$Acl = Get-Acl $Path
$AccessRule = New-Object system.security.accesscontrol.filesystemaccessrule($Identity, $FileSystemRights,'ContainerInherit, ObjectInherit', 'None', $AccessControlType)
$Acl.SetAccessRule($AccessRule)
Set-Acl $Path $Acl




}
## [END] Edit-WACCOFolderShareUser ##
function Edit-WACCOSmbFileShare {
<#

.SYNOPSIS
Edits the smb file share details on the server.

.DESCRIPTION
Edits the smb file share details on the server.

.ROLE
Administrators

#>
param (
    [Parameter(Mandatory = $true)]
    [String]
    $name,

    [Parameter(Mandatory = $false)]
    [String[]]
    $noAccess,

    [Parameter(Mandatory = $false)]
    [String[]]
    $fullAccess,

    [Parameter(Mandatory = $false)]
    [String[]]
    $changeAccess,

    [Parameter(Mandatory = $false)]
    [String[]]
    $readAccess,
    
    [Parameter(Mandatory = $false)]
    [String[]]
    $unblockAccess,

    [Parameter(Mandatory = $false)]
    [Int]
    $cachingMode,

    [Parameter(Mandatory = $false)]
    [boolean]
    $encryptData,

    # TODO: 
    # [Parameter(Mandatory = $false)]
    # [Int]
    # $folderEnumerationMode

    [Parameter(Mandatory = $false)]
    [boolean]
    $compressData,

    [Parameter(Mandatory = $false)]
    [boolean]
    $isCompressDataEnabled
)

if($fullAccess.count -gt 0){
    Grant-SmbShareAccess -Name "$name" -AccountName $fullAccess -AccessRight Full -SmbInstance Default -Force
}
if($changeAccess.count -gt 0){
    Grant-SmbShareAccess -Name "$name" -AccountName $changeAccess -AccessRight Change -SmbInstance Default -Force
}
if($readAccess.count -gt 0){
    Grant-SmbShareAccess -Name "$name" -AccountName $readAccess -AccessRight Read -SmbInstance Default -Force
}
if($noAccess.count -gt 0){
    Revoke-SmbShareAccess -Name "$name" -AccountName $noAccess -SmbInstance Default  -Force
    Block-SmbShareAccess -Name "$name" -AccountName $noAccess -SmbInstance Default -Force
}
if($unblockAccess.count -gt 0){
    Unblock-SmbShareAccess -Name "$name" -AccountName $unblockAccess -SmbInstance Default  -Force
}
if($isCompressDataEnabled){
    Set-SmbShare -Name "$name" -CompressData $compressData -Force
}

Set-SmbShare -Name "$name" -CachingMode "$cachingMode" -EncryptData  $encryptData -Force




}
## [END] Edit-WACCOSmbFileShare ##
function Edit-WACCOSmbServerCertificateMapping {
<#
.SYNOPSIS
Edit SMB Server Certificate Mapping

.DESCRIPTION
Edits smb over QUIC certificate.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER thumbprint
    String -- The thumbprint of the certifiacte selected.

.PARAMETER newSelectedDnsNames
    String[] -- The addresses newly added to the certificate mapping.

.PARAMETER unSelectedDnsNames
    String[] -- To addresses to be removed from the certificate mapping.

#>
param (
    [Parameter(Mandatory = $true)]
    [String]
    $Thumbprint,

    [Parameter(Mandatory = $false)]
    [String[]]
    $NewSelectedDnsNames,

    [Parameter(Mandatory = $false)]
    [String[]]
    $UnSelectedDnsNames,

    [Parameter(Mandatory = $true)]
    [boolean]
    $IsKdcProxyEnabled,

    [Parameter(Mandatory = $true)]
    [String]
    $KdcProxyOptionSelected,

    [Parameter(Mandatory = $true)]
    [boolean]
    $IsKdcProxyMappedForSmbOverQuic,

    [Parameter(Mandatory = $false)]
    [String]
    $KdcPort,

    [Parameter(Mandatory = $false)]
    [String]
    $CurrentkdcPort,

    [Parameter(Mandatory = $false)]
    [boolean]
    $IsSameCertificate
)

Import-Module -Name Microsoft.PowerShell.Management -ErrorAction SilentlyContinue
Set-Variable -Name LogName -Option Constant -Value "Microsoft-ServerManagementExperience" -ErrorAction SilentlyContinue
Set-Variable -Name LogSource -Option Constant -Value "SmeScripts-ConfigureKdcProxy" -ErrorAction SilentlyContinue

function writeInfoLog($logMessage) {
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Information `
        -Message $logMessage -ErrorAction SilentlyContinue
}

Set-Location Cert:\LocalMachine\My

$port = "0.0.0.0:"
$urlLeft = "https://+:"
$urlRight = "/KdcProxy"

if ($UnSelectedDnsNames.count -gt 0) {
    foreach ($unSelectedDnsName in $UnSelectedDnsNames) {
        Remove-SmbServerCertificateMapping -Name $unSelectedDnsName -Force
    }
}

if ($NewSelectedDnsNames.count -gt 0) {
    foreach ($newSelectedDnsName in $NewSelectedDnsNames) {
        New-SmbServerCertificateMapping -Name $newSelectedDnsName -Thumbprint $Thumbprint -StoreName My -Force
    }
}

function Delete-KdcSSLCert([string]$deletePort) {
    $ipport = $port+$deletePort
    $deleteCertKdc = netsh http delete sslcert ipport=$ipport
    if ($LASTEXITCODE -ne 0) {
        throw $deleteCertKdc
    }
    $message = 'Completed deleting  ssl certificate port'
    writeInfoLog $message
    return;
}

function Enable-KdcProxy {

    $ipport = $port+$KdcPort
    $ComputerName = (Get-CimInstance Win32_ComputerSystem).DNSHostName + "." + (Get-CimInstance Win32_ComputerSystem).Domain
    try {
        if(!$IsSameCertificate -or ($KdcPort -ne $CurrentkdcPort) -or (!$IsKdcProxyMappedForSmbOverQuic) ) {
            $guid = [Guid]::NewGuid()
            Add-NetIPHttpsCertBinding -ipport $ipport -CertificateHash $Thumbprint -CertificateStoreName "my" -ApplicationId "{$guid}" -NullEncryption $false
            $message = 'Completed adding ssl certificate port'
            writeInfoLog $message
        }
        if ($NewSelectedDnsNames.count -gt 0) {
            foreach ($newSelectedDnsName in $NewSelectedDnsNames) {
                if($ComputerName.trim() -ne $newSelectedDnsName.trim()){
                    $output = Echo 'Y' | netdom computername $ComputerName /add $newSelectedDnsName
                    if ($LASTEXITCODE -ne 0) {
                        throw $output
                    }
                }
                $message = 'Completed adding alternate names for the computer'
                writeInfoLog $message
            }
        }
        if (!$IsKdcProxyEnabled) {
            $url = $urlLeft + $KdcPort + $urlRight
            $netshOutput = netsh http add urlacl url=$url user="NT authority\Network Service"
            if ($LASTEXITCODE -ne 0) {
                throw $netshOutput
            }
            $message = 'Completed adding urlacl'
            writeInfoLog $message
            New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings"  -force
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "HttpsClientAuth" -Value 0x0 -type DWORD
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "DisallowUnprotectedPasswordAuth" -Value 0x0 -type DWORD
            Set-Service -Name kpssvc -StartupType Automatic
            Start-Service -Name kpssvc
        }
        $message = 'Returning from call Enable-KdcProxy'
        writeInfoLog $message
        return $true;
    }
    catch {
        throw $_
    }
}

if($IsKdcProxyEnabled -and $KdcProxyOptionSelected -eq "enabled" -and ($KdcPort -ne $CurrentkdcPort)) {
    $url = $urlLeft + $CurrentkdcPort + $urlRight
    $deleteOutput = netsh http delete urlacl url=$url
    if ($LASTEXITCODE -ne 0) {
        throw $deleteOutput
    }
    $message = 'Completed deleting urlacl'
    writeInfoLog $message
    $newUrl = $urlLeft + $KdcPort + $urlRight
    $netshOutput = netsh http add urlacl url=$newUrl user="NT authority\Network Service"
    if ($LASTEXITCODE -ne 0) {
        throw $netshOutput
    }
    $message = 'Completed adding urlacl'
    writeInfoLog $message
}

if ($KdcProxyOptionSelected -eq "enabled" -and $KdcPort -ne $null) {
    if($IsKdcProxyMappedForSmbOverQuic -and (!$IsSameCertificate -or ($KdcPort -ne $CurrentkdcPort))) {
        Delete-KdcSSLCert $CurrentkdcPort
    }
    $result = Enable-KdcProxy
    if ($result) {
        $firewallString = "KDC Proxy Server service (KPS) for SMB over QUIC"
        $firewallDesc = "The KDC Proxy Server service runs on edge servers to proxy Kerberos protocol messages to domain controllers on the corporate network. Default port is TCP/443."
        New-NetFirewallRule -DisplayName $firewallString -Description $firewallDesc -Protocol TCP -LocalPort $KdcPort -Direction Inbound -Action Allow
    }
}

if ($IsKdcProxyMappedForSmbOverQuic -and $KdcProxyOptionSelected -ne "enabled" ) {
    Delete-KdcSSLCert $CurrentKdcPort
    $firewallString = "KDC Proxy Server service (KPS) for SMB over QUIC"
    Remove-NetFirewallRule -DisplayName $firewallString
}



}
## [END] Edit-WACCOSmbServerCertificateMapping ##
function Enable-WACCOSmbOverQuic {
<#

.SYNOPSIS
Disables smb over QUIC on the server.

.DESCRIPTION
Disables smb over QUIC on the server.

.ROLE
Administrators

#>

Set-SmbServerConfiguration -EnableSMBQUIC $true -Force
}
## [END] Enable-WACCOSmbOverQuic ##
function Expand-WACCOArchiveFileSystemEntity {
<#

.SYNOPSIS
Expands the specified file system entity (files, folders) of the system.

.DESCRIPTION
Expands the specified file system entity (files, folders) of the system on this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER pathSource
    String -- The path to expand.

.PARAMETER PathDestination
    String -- The destination path to expand into.

.PARAMETER Force
    boolean -- override any confirmations

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $PathSource,    

    [Parameter(Mandatory = $true)]
    [String]
    $PathDestination,

    [Parameter(Mandatory = $false)]
    [boolean]
    $Force
)

Set-StrictMode -Version 5.0

if ($Force) {
    Expand-Archive -Path $PathSource -Force -DestinationPath $PathDestination
} else {
    Expand-Archive -Path $PathSource -DestinationPath $PathDestination
}

if ($error) {
    $code = $error[0].Exception.HResult
    @{ status = "error"; code = $code; message = $error }
} else {
    @{ status = "ok"; }
}

}
## [END] Expand-WACCOArchiveFileSystemEntity ##
function Export-WACCOEventLogChannel {
<#

.SYNOPSIS
Export the event log channel file (.evtx) with filter XML.

.DESCRIPTION
Export the event log channel file (.evtx) with filter XML.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>

Param(
    [string]$channel,
    [string]$filterXml
)

$segments = $channel.Split("-")
$name = $segments[-1]

$randomString = [GUID]::NewGuid().ToString()
$ResultFile = $env:temp + "\" + $name + "_" + $randomString + ".evtx"
$ResultFile = $ResultFile -replace "/", "-"

wevtutil epl "$channel" "$ResultFile" /q:"$filterXml" /ow:true

return $ResultFile

}
## [END] Export-WACCOEventLogChannel ##
function Get-WACCOBestHostNode {
<#

.SYNOPSIS
Returns the list of available cluster node names, and the best node name to host a new virtual machine.

.DESCRIPTION
Use the cluster CIM provider (MSCluster) to ask the cluster which node is the best to host a new virtual machine.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

Import-Module CimCmdlets -ErrorAction SilentlyContinue
Import-Module FailoverClusters -ErrorAction SilentlyContinue

Import-LocalizedData -BindingVariable strings -FileName strings.psd1 -ErrorAction SilentlyContinue


<#

.SYNOPSIS
Setup the script environment.

.DESCRIPTION
Setup the script environment.  Create read only (constant) variables
that add context to the said constants.

#>

function setupScriptEnv() {
    Set-Variable -Name LogName -Option ReadOnly -Value "Microsoft-ServerManagementExperience" -Scope Script
    Set-Variable -Name LogSource -Option ReadOnly -Value "SMEScripts" -Scope Script
    Set-Variable -Name clusterCimNameSpace -Option ReadOnly -Value "root/MSCluster" -Scope Script
    Set-Variable -Name ScriptName -Option ReadOnly -Value "Get-BestHostNode.ps1" -Scope Script
    Set-Variable -Name BestNodePropertyName -Option ReadOnly -Value "BestNode" -Scope Script
    Set-Variable -Name StateUp -Option ReadOnly -Value "0" -Scope Script
}

<#

.SYNOPSIS
Cleanup the script environment.

.DESCRIPTION
Cleanup the script environment.

#>

function cleanupScriptEnv() {
    Remove-Variable -Name LogName -Scope Script -Force
    Remove-Variable -Name LogSource -Scope Script -Force
    Remove-Variable -Name clusterCimNameSpace -Scope Script -Force
    Remove-Variable -Name ScriptName -Scope Script -Force
    Remove-Variable -Name BestNodePropertyName -Scope Script -Force
    Remove-Variable -Name StateUp -Scope Script -Force
}
    
<#

.SYNOPSIS
Get the fully qualified domain name for the passed in server name from DNS.

.DESCRIPTION
Get the fully qualified domain name for the passed in server name from DNS.

#>

function GetServerFqdn([string]$netBIOSName) {
    try {
        $fqdn = [System.Net.DNS]::GetHostByName($netBIOSName).HostName

        return $fqdn.ToLower()
    } catch {
        $errMessage = $_.Exception.Message

        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message "[$ScriptName]: There was an error looking up the FQDN for server $netBIOSName.  Error: $errMessage"  -ErrorAction SilentlyContinue

        return $netBIOSName
    }    
}

<#

.SYNOPSIS
Are the cluster PowerShell cmdlets installed on this server?

.DESCRIPTION
Are the cluster PowerShell cmdlets installed on this server?

#>

function getIsClusterCmdletsAvailable() {
    $cmdlet = Get-Command "Get-Cluster" -ErrorAction SilentlyContinue

    return !!$cmdlet
}

<#

.SYNOPSIS
is the cluster CIM (WMI) provider installed on this server?

.DESCRIPTION
Returns true when the cluster CIM provider is installed on this server.

#>

function isClusterCimProviderAvailable() {
    $namespace = Get-CimInstance -Namespace $clusterCimNamespace -ClassName __NAMESPACE -ErrorAction SilentlyContinue

    return !!$namespace
}

<#

.SYNOPSIS
Get the MSCluster Cluster Service CIM instance from this server.

.DESCRIPTION
Get the MSCluster Cluster Service CIM instance from this server.

#>

function getClusterServiceCimInstance() {
    return Get-CimInstance -Namespace $clusterCimNamespace MSCluster_ClusterService -ErrorAction SilentlyContinue
}

<#

.SYNOPSIS
Get the list of the cluster nodes that are running.

.DESCRIPTION
Returns a list of cluster node names that are running using PowerShell.

#>

function getAllUpClusterNodeNames() {
    # Constants
    Set-Variable -Name stateUp -Option Readonly -Value "up" -Scope Local

    try {
        return Get-ClusterNode | Where-Object { $_.State -eq $stateUp } | ForEach-Object { (GetServerFqdn $_.Name) }
    } finally {
        Remove-Variable -Name stateUp -Scope Local -Force
    }
}

<#

.SYNOPSIS
Get the list of the cluster nodes that are running.

.DESCRIPTION
Returns a list of cluster node names that are running using CIM.

#>

function getAllUpClusterCimNodeNames() {
##SkipCheck=true##
    $query = "select name, state from MSCluster_Node Where state = '{0}'" -f $StateUp
##SkipCheck=false##
    return Get-CimInstance -Namespace $clusterCimNamespace -Query $query | ForEach-Object { (GetServerFqdn $_.Name) }
}

<#

.SYNOPSIS
Create a new instance of the "results" PS object.

.DESCRIPTION
Create a new PS object and set the passed in nodeNames to the appropriate property.

#>

function newResult([string []] $nodeNames) {
    $result = new-object PSObject
    $result | Add-Member -Type NoteProperty -Name Nodes -Value $nodeNames

    return $result;
}

<#

.SYNOPSIS
Remove any old lingering reservation for our typical VM.

.DESCRIPTION
Remove the reservation from the passed in id.

#>

function removeReservation($clusterService, [string] $rsvId) {
    Set-Variable removeReservationMethodName -Option Constant -Value "RemoveVmReservation"

    Invoke-CimMethod -CimInstance $clusterService -MethodName $removeReservationMethodName -Arguments @{ReservationId = $rsvId} -ErrorVariable +err | Out-Null    
}

<#

.SYNOPSIS
Create a reservation for our typical VM.

.DESCRIPTION
Create a reservation for the passed in id.

#>

function createReservation($clusterService, [string] $rsvId) {
    Set-Variable -Name createReservationMethodName -Option ReadOnly -Value "CreateVmReservation" -Scope Local
    Set-Variable -Name reserveSettings -Option ReadOnly -Value @{VmMemory = 2048; VmVirtualCoreCount = 2; VmCpuReservation = 0; VmFlags = 0; TimeSpan = 2000; ReservationId = $rsvId; LocalDiskSize = 0; Version = 0} -Scope Local

    try {
        $vmReserve = Invoke-CimMethod -CimInstance $clusterService -MethodName $createReservationMethodName -ErrorAction SilentlyContinue -ErrorVariable va -Arguments $reserveSettings

        if (!!$vmReserve -and $vmReserve.ReturnValue -eq 0 -and !!$vmReserve.NodeId) {
            return $vmReserve.NodeId
        }

        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message "[$ScriptName]: Could not create a reservation for a virtual machine. Output from $createReservationMethodName is $vmReserve"  -ErrorAction SilentlyContinue

        return $null
    } finally {
        Remove-Variable -Name createReservationMethodName -Scope Local -Force
        Remove-Variable -Name reserveSettings -Scope Local -Force
    }
}

<#

.SYNOPSIS
Use the Cluster CIM provider to find the best host name for a typical VM.

.DESCRIPTION
Returns the best host node name, or null when none are found.

#>

function askClusterServiceForBestHostNode() {
    # API parameters
    Set-Variable -Name rsvId -Option ReadOnly -Value "TempVmId1" -Scope Local
    
    try {
        # If the class exist, using api to get optimal host
        $clusterService = getClusterServiceCimInstance
        if (!!$clusterService) {
            $nodeNames = @(getAllUpClusterCimNodeNames)
            $result = newResult $nodeNames
        
            # remove old reserveration if there is any
            removeReservation $clusterService $rsvId

            $id = createReservation $clusterService $rsvId

            if (!!$id) {
    ##SkipCheck=true##            
                $query = "select name, id from MSCluster_Node where id = '{0}'" -f $id
    ##SkipCheck=false##
                $bestNode = Get-CimInstance -Namespace $clusterCimNamespace -Query $query -ErrorAction SilentlyContinue

                if ($bestNode) {
                    $result | Add-Member -Type NoteProperty -Name $BestNodePropertyName -Value (GetServerFqdn $bestNode.Name)

                    return $result
                }
            } 
        }

        return $null
    } finally {
        Remove-Variable -Name rsvId -Scope Local -Force
    }
}

<#

.SYNOPSIS
Get the name of the cluster node that has the least number of VMs running on it.

.DESCRIPTION
Return the name of the cluster node that has the least number of VMs running on it.

#>

function getLeastLoadedNode() {
    # Constants
    Set-Variable -Name vmResourceTypeName -Option ReadOnly -Value "Virtual Machine" -Scope Local
    Set-Variable -Name OwnerNodePropertyName -Option ReadOnly -Value "OwnerNode" -Scope Local

    try {
        $nodeNames = @(getAllUpClusterNodeNames)
        $bestNodeName = $null;

        $result = newResult $nodeNames

        $virtualMachinesPerNode = @{}

        # initial counts as 0
        $nodeNames | ForEach-Object { $virtualMachinesPerNode[$_] = 0 }

        $ownerNodes = Get-ClusterResource | Where-Object { $_.ResourceType -eq $vmResourceTypeName } | Microsoft.PowerShell.Utility\Select-Object $OwnerNodePropertyName
        $ownerNodes | ForEach-Object { $virtualMachinesPerNode[$_.OwnerNode.Name]++ }

        # find node with minimum count
        $bestNodeName = $nodeNames[0]
        $min = $virtualMachinesPerNode[$bestNodeName]

        $nodeNames | ForEach-Object { 
            if ($virtualMachinesPerNode[$_] -lt $min) {
                $bestNodeName = $_
                $min = $virtualMachinesPerNode[$_]
            }
        }

        $result | Add-Member -Type NoteProperty -Name $BestNodePropertyName -Value (GetServerFqdn $bestNodeName)

        return $result
    } finally {
        Remove-Variable -Name vmResourceTypeName -Scope Local -Force
        Remove-Variable -Name OwnerNodePropertyName -Scope Local -Force
    }
}

<#

.SYNOPSIS
Main

.DESCRIPTION
Use the various mechanism available to determine the best host node.

#>

function main() {
    if (isClusterCimProviderAvailable) {
        $bestNode = askClusterServiceForBestHostNode
        if (!!$bestNode) {
            return $bestNode
        }
    }

    if (getIsClusterCmdletsAvailable) {
        return getLeastLoadedNode
    } else {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Warning `
            -Message "[$ScriptName]: The required PowerShell module (FailoverClusters) was not found."  -ErrorAction SilentlyContinue

        Write-Warning $strings.FailoverClustersModuleRequired
    }

    return $null    
}

###############################################################################
# Script execution begins here.
###############################################################################

if (-not ($env:pester)) {
    setupScriptEnv

    try {
        Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

        $result = main
        if (!!$result) {
            return $result
        }

        # If neither cluster CIM provider or PowerShell cmdlets are available then simply
        # return this computer's name as the best host node...
        $nodeName = GetServerFqdn $env:COMPUTERNAME

        $result = newResult @($nodeName)
        $result | Add-Member -Type NoteProperty -Name $BestNodePropertyName -Value $nodeName

        return $result
    } finally {
        cleanupScriptEnv
    }
}

}
## [END] Get-WACCOBestHostNode ##
function Get-WACCOCertificates {
<#

.SYNOPSIS
Get the certificates stored in my\store

.DESCRIPTION
Get the certificates stored in my\store

.ROLE
Readers

#>

$nearlyExpiredThresholdInDays = 60

$dnsNameList = @{}

<#
.Synopsis
    Name: Compute-ExpirationStatus
    Description: Computes expiration status based on notAfter date.
.Parameters
    $notAfter: A date object refering to certificate expiry date.

.Returns
    Enum values "Expired", "NearlyExpired" and "Healthy"
#>
function Compute-ExpirationStatus {
    param (
        [Parameter(Mandatory = $true)]
        [DateTime]$notAfter
    )

    if ([DateTime]::Now -gt $notAfter) {
        $expirationStatus = "Expired"
    }
    else {
        $nearlyExpired = [DateTime]::Now.AddDays($nearlyExpiredThresholdInDays);

        if ($nearlyExpired -ge $notAfter) {
            $expirationStatus = "NearlyExpired"
        }
        else {
            $expirationStatus = "Healthy"
        }
    }

    $expirationStatus
}

<# main - script starts here #>

Set-Location Cert:\LocalMachine\My

$certificates = Get-ChildItem -Recurse | Microsoft.PowerShell.Utility\Select-Object Subject, FriendlyName, NotBefore, NotAfter,
 Thumbprint, Issuer, @{n="DnsNameList";e={$_.DnsNameList}}, @{n="SignatureAlgorithm";e={$_.SignatureAlgorithm.FriendlyName}} |
ForEach-Object {
    return @{
        CertificateName = $_.Subject;
        FriendlyName = $_.FriendlyName;
        NotBefore = $_.NotBefore;
        NotAfter = $_.NotAfter;
        Thumbprint = $_.Thumbprint;
        Issuer = $_.Issuer;
        DnsNameList = $_.DnsNameList;
        Status = $(Compute-ExpirationStatus $_.NotAfter);
        SignatureAlgorithm  = $_.SignatureAlgorithm;
    }
}

return $certificates;

}
## [END] Get-WACCOCertificates ##
function Get-WACCOCimEventLogRecords {
<#

.SYNOPSIS
Get Log records of event channel by using Server Manager CIM provider.

.DESCRIPTION
Get Log records of event channel by using Server Manager CIM provider.

.ROLE
Readers

#>

Param(
    [string]$FilterXml,
    [bool]$ReverseDirection
)

import-module CimCmdlets

$machineName = [System.Net.DNS]::GetHostByName('').HostName
Invoke-CimMethod -Namespace root/Microsoft/Windows/ServerManager -ClassName MSFT_ServerManagerTasks -MethodName GetServerEventDetailEx -Arguments @{FilterXml = $FilterXml; ReverseDirection = $ReverseDirection; } |
    ForEach-Object {
        $result = $_
        if ($result.PSObject.Properties.Match('ItemValue').Count) {
            foreach ($item in $result.ItemValue) {
                @{
                    ItemValue = 
                    @{
                        Description  = $item.description
                        Id           = $item.id
                        Level        = $item.level
                        Log          = $item.log
                        Source       = $item.source
                        Timestamp    = $item.timestamp
                        __ServerName = $machineName
                    }
                }
            }
        }
    }

}
## [END] Get-WACCOCimEventLogRecords ##
function Get-WACCOCimWin32LogicalDisk {
<#

.SYNOPSIS
Gets Win32_LogicalDisk object.

.DESCRIPTION
Gets Win32_LogicalDisk object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_LogicalDisk

}
## [END] Get-WACCOCimWin32LogicalDisk ##
function Get-WACCOCimWin32NetworkAdapter {
<#

.SYNOPSIS
Gets Win32_NetworkAdapter object.

.DESCRIPTION
Gets Win32_NetworkAdapter object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_NetworkAdapter

}
## [END] Get-WACCOCimWin32NetworkAdapter ##
function Get-WACCOCimWin32PhysicalMemory {
<#

.SYNOPSIS
Gets Win32_PhysicalMemory object.

.DESCRIPTION
Gets Win32_PhysicalMemory object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_PhysicalMemory

}
## [END] Get-WACCOCimWin32PhysicalMemory ##
function Get-WACCOCimWin32Processor {
<#

.SYNOPSIS
Gets Win32_Processor object.

.DESCRIPTION
Gets Win32_Processor object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_Processor

}
## [END] Get-WACCOCimWin32Processor ##
function Get-WACCOClusterEvents {
<#
.SYNOPSIS
Gets CIM instance

.DESCRIPTION
Gets CIM instance

.ROLE
Readers

#>

param (
		[Parameter(Mandatory = $true)]
		[string]
    $namespace,

    [Parameter(Mandatory = $true)]
		[string]
    $className

)
Import-Module CimCmdlets
Get-CimInstance -Namespace  $namespace -ClassName $className

}
## [END] Get-WACCOClusterEvents ##
function Get-WACCOClusterInventory {
<#

.SYNOPSIS
Retrieves the inventory data for a cluster.

.DESCRIPTION
Retrieves the inventory data for a cluster.

.ROLE
Readers

#>

Import-Module CimCmdlets -ErrorAction SilentlyContinue

# JEA code requires to pre-import the module (this is slow on failover cluster environment.)
Import-Module FailoverClusters -ErrorAction SilentlyContinue

Import-Module Storage -ErrorAction SilentlyContinue
<#

.SYNOPSIS
Get the name of this computer.

.DESCRIPTION
Get the best available name for this computer.  The FQDN is preferred, but when not avaialble
the NetBIOS name will be used instead.

#>

function getComputerName() {
    $computerSystem = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object Name, DNSHostName

    if ($computerSystem) {
        $computerName = $computerSystem.DNSHostName

        if ($null -eq $computerName) {
            $computerName = $computerSystem.Name
        }

        return $computerName
    }

    return $null
}

<#

.SYNOPSIS
Are the cluster PowerShell cmdlets installed on this server?

.DESCRIPTION
Are the cluster PowerShell cmdlets installed on this server?

#>

function getIsClusterCmdletAvailable() {
    $cmdlet = Get-Command "Get-Cluster" -ErrorAction SilentlyContinue

    return !!$cmdlet
}

<#

.SYNOPSIS
Get the MSCluster Cluster CIM instance from this server.

.DESCRIPTION
Get the MSCluster Cluster CIM instance from this server.

#>
function getClusterCimInstance() {
    $namespace = Get-CimInstance -Namespace root/MSCluster -ClassName __NAMESPACE -ErrorAction SilentlyContinue

    if ($namespace) {
        return Get-CimInstance -Namespace root/mscluster MSCluster_Cluster -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object fqdn, S2DEnabled
    }

    return $null
}


<#

.SYNOPSIS
Determines if the current cluster supports Failover Clusters Time Series Database.

.DESCRIPTION
Use the existance of the path value of cmdlet Get-StorageHealthSetting to determine if TSDB
is supported or not.

#>
function getClusterPerformanceHistoryPath() {
    return $null -ne (Get-StorageSubSystem clus* | Get-StorageHealthSetting -Name "System.PerformanceHistory.Path")
}

<#

.SYNOPSIS
Get some basic information about the cluster from the cluster.

.DESCRIPTION
Get the needed cluster properties from the cluster.

#>
function getClusterInfo() {
    $returnValues = @{}

    $returnValues.Fqdn = $null
    $returnValues.isS2DEnabled = $false
    $returnValues.isTsdbEnabled = $false

    $cluster = getClusterCimInstance
    if ($cluster) {
        $returnValues.Fqdn = $cluster.fqdn
        $isS2dEnabled = !!(Get-Member -InputObject $cluster -Name "S2DEnabled") -and ($cluster.S2DEnabled -eq 1)
        $returnValues.isS2DEnabled = $isS2dEnabled

        if ($isS2DEnabled) {
            $returnValues.isTsdbEnabled = getClusterPerformanceHistoryPath
        } else {
            $returnValues.isTsdbEnabled = $false
        }
    }

    return $returnValues
}

<#

.SYNOPSIS
Are the cluster PowerShell Health cmdlets installed on this server?

.DESCRIPTION
Are the cluster PowerShell Health cmdlets installed on this server?

s#>
function getisClusterHealthCmdletAvailable() {
    $cmdlet = Get-Command -Name "Get-HealthFault" -ErrorAction SilentlyContinue

    return !!$cmdlet
}
<#

.SYNOPSIS
Are the Britannica (sddc management resources) available on the cluster?

.DESCRIPTION
Are the Britannica (sddc management resources) available on the cluster?

#>
function getIsBritannicaEnabled() {
    return $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_Cluster -ErrorAction SilentlyContinue)
}

<#

.SYNOPSIS
Are the Britannica (sddc management resources) virtual machine available on the cluster?

.DESCRIPTION
Are the Britannica (sddc management resources) virtual machine available on the cluster?

#>
function getIsBritannicaVirtualMachineEnabled() {
    return $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_VirtualMachine -ErrorAction SilentlyContinue)
}

<#

.SYNOPSIS
Are the Britannica (sddc management resources) virtual switch available on the cluster?

.DESCRIPTION
Are the Britannica (sddc management resources) virtual switch available on the cluster?

#>
function getIsBritannicaVirtualSwitchEnabled() {
    return $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_VirtualSwitch -ErrorAction SilentlyContinue)
}

###########################################################################
# main()
###########################################################################

$clusterInfo = getClusterInfo

$result = New-Object PSObject

$result | Add-Member -MemberType NoteProperty -Name 'Fqdn' -Value $clusterInfo.Fqdn
$result | Add-Member -MemberType NoteProperty -Name 'IsS2DEnabled' -Value $clusterInfo.isS2DEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsTsdbEnabled' -Value $clusterInfo.isTsdbEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsClusterHealthCmdletAvailable' -Value (getIsClusterHealthCmdletAvailable)
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaEnabled' -Value (getIsBritannicaEnabled)
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaVirtualMachineEnabled' -Value (getIsBritannicaVirtualMachineEnabled)
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaVirtualSwitchEnabled' -Value (getIsBritannicaVirtualSwitchEnabled)
$result | Add-Member -MemberType NoteProperty -Name 'IsClusterCmdletAvailable' -Value (getIsClusterCmdletAvailable)
$result | Add-Member -MemberType NoteProperty -Name 'CurrentClusterNode' -Value (getComputerName)

$result

}
## [END] Get-WACCOClusterInventory ##
function Get-WACCOClusterNodes {
<#

.SYNOPSIS
Retrieves the inventory data for cluster nodes in a particular cluster.

.DESCRIPTION
Retrieves the inventory data for cluster nodes in a particular cluster.

.ROLE
Readers

#>

import-module CimCmdlets

# JEA code requires to pre-import the module (this is slow on failover cluster environment.)
import-module FailoverClusters -ErrorAction SilentlyContinue

###############################################################################
# Constants
###############################################################################

Set-Variable -Name LogName -Option Constant -Value "Microsoft-ServerManagementExperience" -ErrorAction SilentlyContinue
Set-Variable -Name LogSource -Option Constant -Value "SMEScripts" -ErrorAction SilentlyContinue
Set-Variable -Name ScriptName -Option Constant -Value $MyInvocation.ScriptName -ErrorAction SilentlyContinue

<#

.SYNOPSIS
Are the cluster PowerShell cmdlets installed?

.DESCRIPTION
Use the Get-Command cmdlet to quickly test if the cluster PowerShell cmdlets
are installed on this server.

#>

function getClusterPowerShellSupport() {
    $cmdletInfo = Get-Command 'Get-ClusterNode' -ErrorAction SilentlyContinue

    return $cmdletInfo -and $cmdletInfo.Name -eq "Get-ClusterNode"
}

<#

.SYNOPSIS
Get the cluster nodes using the cluster CIM provider.

.DESCRIPTION
When the cluster PowerShell cmdlets are not available fallback to using
the cluster CIM provider to get the needed information.

#>

function getClusterNodeCimInstances() {
    # Change the WMI property NodeDrainStatus to DrainStatus to match the PS cmdlet output.
    return Get-CimInstance -Namespace root/mscluster MSCluster_Node -ErrorAction SilentlyContinue | `
        Microsoft.PowerShell.Utility\Select-Object @{Name="DrainStatus"; Expression={$_.NodeDrainStatus}}, DynamicWeight, Name, NodeWeight, FaultDomain, State
}

<#

.SYNOPSIS
Get the cluster nodes using the cluster PowerShell cmdlets.

.DESCRIPTION
When the cluster PowerShell cmdlets are available use this preferred function.

#>

function getClusterNodePsInstances() {
    return Get-ClusterNode -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object DrainStatus, DynamicWeight, Name, NodeWeight, FaultDomain, State
}

<#

.SYNOPSIS
Use DNS services to get the FQDN of the cluster NetBIOS name.

.DESCRIPTION
Use DNS services to get the FQDN of the cluster NetBIOS name.

.Notes
It is encouraged that the caller add their approprate -ErrorAction when
calling this function.

#>

function getClusterNodeFqdn([string]$clusterNodeName) {
    return ([System.Net.Dns]::GetHostEntry($clusterNodeName)).HostName
}

<#

.SYNOPSIS
Writes message to event log as warning.

.DESCRIPTION
Writes message to event log as warning.

#>

function writeToEventLog([string]$message) {
    Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Warning `
        -Message $message  -ErrorAction SilentlyContinue
}

<#

.SYNOPSIS
Get the cluster nodes.

.DESCRIPTION
When the cluster PowerShell cmdlets are available get the information about the cluster nodes
using PowerShell.  When the cmdlets are not available use the Cluster CIM provider.

#>

function getClusterNodes() {
    $isClusterCmdletAvailable = getClusterPowerShellSupport

    if ($isClusterCmdletAvailable) {
        $clusterNodes = getClusterNodePsInstances
    } else {
        $clusterNodes = getClusterNodeCimInstances
    }

    $clusterNodeMap = @{}

    foreach ($clusterNode in $clusterNodes) {
        $clusterNodeName = $clusterNode.Name.ToLower()
        try 
        {
            $clusterNodeFqdn = getClusterNodeFqdn $clusterNodeName -ErrorAction SilentlyContinue
        }
        catch 
        {
            $clusterNodeFqdn = $clusterNodeName
            writeToEventLog "[$ScriptName]: The fqdn for node '$clusterNodeName' could not be obtained. Defaulting to machine name '$clusterNodeName'"
        }

        $clusterNodeResult = New-Object PSObject

        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'FullyQualifiedDomainName' -Value $clusterNodeFqdn
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'Name' -Value $clusterNodeName
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'DynamicWeight' -Value $clusterNode.DynamicWeight
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'NodeWeight' -Value $clusterNode.NodeWeight
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'FaultDomain' -Value $clusterNode.FaultDomain
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'State' -Value $clusterNode.State
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'DrainStatus' -Value $clusterNode.DrainStatus

        $clusterNodeMap.Add($clusterNodeName, $clusterNodeResult)
    }

    return $clusterNodeMap
}

###########################################################################
# main()
###########################################################################

getClusterNodes

}
## [END] Get-WACCOClusterNodes ##
function Get-WACCOComputerName {
<#

.SYNOPSIS
Gets the computer name.

.DESCRIPTION
Gets the compuiter name.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

$ComputerName = $env:COMPUTERNAME
@{ computerName = $ComputerName }

}
## [END] Get-WACCOComputerName ##
function Get-WACCODecryptedDataFromNode {
<#

.SYNOPSIS
Gets data after decrypting it on a node.

.DESCRIPTION
Decrypts data on node using a cached RSAProvider used during encryption within 3 minutes of encryption and returns the decrypted data.
This script should be imported or copied directly to other scripts, do not send the returned data as an argument to other scripts.

.PARAMETER encryptedData
Encrypted data to be decrypted (String).

.ROLE
Readers

#>
param (
  [Parameter(Mandatory = $true)]
  [String]
  $encryptedData
)

Set-StrictMode -Version 5.0

New-Variable -Name rsaProviderInstanceName -Value "RSA" -Option Constant

function DecryptDataWithJWKOnNode {
  if(Get-Variable -Scope Global -Name $rsaProviderInstanceName -EA SilentlyContinue) {
    $rsaProvider = (Get-Variable -Scope Global -Name $rsaProviderInstanceName).Value
    $decryptedBytes = $rsaProvider.Decrypt([Convert]::FromBase64String($encryptedData), [System.Security.Cryptography.RSAEncryptionPadding]::OaepSHA1)
    return [System.Text.Encoding]::UTF8.GetString($decryptedBytes)
  }
  # If you copy this script directly to another, you can get rid of the throw statement and add custom error handling logic such as "Write-Error"
  throw [System.InvalidOperationException] "Password decryption failed. RSACryptoServiceProvider Instance not found"
}

}
## [END] Get-WACCODecryptedDataFromNode ##
function Get-WACCOEflowSshParameters {
<#
.SYNOPSIS
Gets the SSH parameters required to establish an SSH connection to an EFLOW device

.DESCRIPTION
This function retrieves the eflow user name, vm IP address and SSH private key from the EFLOW PS module

.ROLE
Administrators
#>

if (-Not (Get-Module -ListAvailable -Name AzureEFLOW)) {
    @{
        Status = "EFLOWModuleNotFound";
    }
    exit
}

if ((Get-ExecutionPolicy) -eq 'Restricted')
{
    $env:PSExecutionPolicyPreference = "AllSigned"
}

$userName = $(Start-Job { Get-EflowVmUserName } | Receive-Job -Wait -AutoRemoveJob)

if ([string]::IsNullOrEmpty($userName)) {
    @{
        Status = "EFLOWUserNameNotFound";
    }
    exit
}

$addrInfo = $(Start-Job { Get-EflowVmAddr } | Receive-Job -Wait -AutoRemoveJob)

if (($addrInfo.length -ne 2) -Or [string]::IsNullOrEmpty($addrInfo[1])) {
    @{
        Status = "EFLOWIPAddressNotFound";
    }
    exit
}

$sshKey = $(Start-Job { Get-EflowVmSshKey } | Receive-Job -Wait -AutoRemoveJob)

if ([string]::IsNullOrEmpty($sshKey)) {
    @{
        Status = "EFLOWSSHKeyNotFound";
    }
    exit
}

@{
    Status = "OK";
    UserName = $userName;
    AddressInfomation = $addrInfo[1];
    SshKey = $sshKey;
}
}
## [END] Get-WACCOEflowSshParameters ##
function Get-WACCOEncryptionJWKOnNode {
<#

.SYNOPSIS
Gets encrytion JSON web key from node.

.DESCRIPTION
Gets encrytion JSON web key from node.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

New-Variable -Name rsaProviderInstanceName -Value "RSA" -Option Constant

function Get-RSAProvider
{
    if(Get-Variable -Scope Global -Name $rsaProviderInstanceName -EA SilentlyContinue)
    {
        return (Get-Variable -Scope Global -Name $rsaProviderInstanceName).Value
    }

    $Global:RSA = New-Object System.Security.Cryptography.RSACryptoServiceProvider -ArgumentList 4096
    return $RSA
}

function Get-JsonWebKey
{
    $rsaProvider = Get-RSAProvider
    $parameters = $rsaProvider.ExportParameters($false)
    return [PSCustomObject]@{
        kty = 'RSA'
        alg = 'RSA-OAEP'
        e = [Convert]::ToBase64String($parameters.Exponent)
        n = [Convert]::ToBase64String($parameters.Modulus).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    }
}

$jwk = Get-JsonWebKey
ConvertTo-Json $jwk -Compress

}
## [END] Get-WACCOEncryptionJWKOnNode ##
function Get-WACCOEventLogDisplayName {
<#

.SYNOPSIS
Get the EventLog log name and display name by using Get-EventLog cmdlet.

.DESCRIPTION
Get the EventLog log name and display name by using Get-EventLog cmdlet.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>


return (Get-EventLog -LogName * | Microsoft.PowerShell.Utility\Select-Object Log,LogDisplayName)
}
## [END] Get-WACCOEventLogDisplayName ##
function Get-WACCOEventLogFilteredCount {
<#

.SYNOPSIS
Get the total amout of events that meet the filters selected by using Get-WinEvent cmdlet.

.DESCRIPTION
Get the total amout of events that meet the filters selected by using Get-WinEvent cmdlet.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>

Param(
    [string]$filterXml
)

return (Get-WinEvent -FilterXml "$filterXml" -ErrorAction 'SilentlyContinue').count
}
## [END] Get-WACCOEventLogFilteredCount ##
function Get-WACCOEventLogRecords {
<#

.SYNOPSIS
Get Log records of event channel by using Get-WinEvent cmdlet.

.DESCRIPTION
Get Log records of event channel by using Get-WinEvent cmdlet.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers
#>

Param(
    [string]
    $filterXml,
    [bool]
    $reverseDirection
)

$ErrorActionPreference = 'SilentlyContinue'
Import-Module Microsoft.PowerShell.Diagnostics;

#
# Prepare parameters for command Get-WinEvent
#
$winEventscmdParams = @{
    FilterXml = $filterXml;
    Oldest    = !$reverseDirection;
}

Get-WinEvent  @winEventscmdParams -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object recordId,
id, 
@{Name = "Log"; Expression = {$_."logname"}}, 
level, 
timeCreated, 
machineName, 
@{Name = "Source"; Expression = {$_."ProviderName"}}, 
@{Name = "Description"; Expression = {$_."Message"}}



}
## [END] Get-WACCOEventLogRecords ##
function Get-WACCOEventLogSummary {
<#

.SYNOPSIS
Get the log summary (Name, Total) for the channel selected by using Get-WinEvent cmdlet.

.DESCRIPTION
Get the log summary (Name, Total) for the channel selected by using Get-WinEvent cmdlet.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>

Param(
    [string]$channel
)

Import-Module Microsoft.PowerShell.Diagnostics

$channelList = $channel.split(",")

Get-WinEvent -ListLog $channelList -Force -ErrorAction SilentlyContinue |`
    Microsoft.PowerShell.Utility\Select-Object LogName, IsEnabled, RecordCount, IsClassicLog, LogType, OwningProviderName
}
## [END] Get-WACCOEventLogSummary ##
function Get-WACCOFileNamesInPath {
<#

.SYNOPSIS
Enumerates all of the file system entities (files, folders, volumes) of the system.

.DESCRIPTION
Enumerates all of the file system entities (files, folders, volumes) of the system on this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

.PARAMETER Path
    String -- The path to enumerate.

.PARAMETER OnlyFolders
    switch -- 

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path,

    [Parameter(Mandatory = $false)]
    [switch]
    $OnlyFolders
)

Set-StrictMode -Version 5.0

function isFolder($item) {
    return $item.Attributes -match "Directory"
}

function getName($item) {
    $slash = '';

    if (isFolder $item) {
        $slash = '\';
    }

    return "$($_.Name)$slash"
}

if ($onlyFolders) {
    return (Get-ChildItem -Path $Path | Where-Object {isFolder $_}) | ForEach-Object { return "$($_.Name)\"} | Sort-Object
}

return (Get-ChildItem -Path $Path) | ForEach-Object { return getName($_)} | Sort-Object

}
## [END] Get-WACCOFileNamesInPath ##
function Get-WACCOFileSystemEntities {
<#

.SYNOPSIS
Enumerates all of the file system entities (files, folders, volumes) of the system.

.DESCRIPTION
Enumerates all of the file system entities (files, folders, volumes) of the system on this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

.PARAMETER Path
    String -- The path to enumerate.

.PARAMETER OnlyFiles
    switch --

.PARAMETER OnlyFolders
    switch --

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path,

    [Parameter(Mandatory = $false)]
    [Switch]
    $OnlyFiles,

    [Parameter(Mandatory = $false)]
    [Switch]
    $OnlyFolders
)

Set-StrictMode -Version 5.0

<#
.Synopsis
    Name: Get-FileSystemEntities
    Description: Gets all the local file system entities of the machine.

.Parameter Path
    String -- The path to enumerate.

.Returns
    The local file system entities.
#>
function Get-FileSystemEntities {
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $Path
    )

    $folderShares = Get-CimInstance -Class Win32_Share;

    if ($Path -match '\[' -or $Path -match '\]') {
        return Get-ChildItem -LiteralPath $Path -Force |
        Microsoft.PowerShell.Utility\Select-Object @{Name = "Caption"; Expression = { $_.FullName } },
        @{Name = "CreationDate"; Expression = { $_.CreationTimeUtc } },
        Extension,
        @{Name = "IsHidden"; Expression = { $_.Attributes -match "Hidden" } },
        @{Name = "IsShared"; Expression = { [bool]($folderShares | Where-Object Path -eq $_.FullName) } },
        Name,
        @{Name = "Type"; Expression = { Get-FileSystemEntityType -Attributes $_.Attributes } },
        @{Name = "LastModifiedDate"; Expression = { $_.LastWriteTimeUtc } },
        @{Name = "Size"; Expression = { if ($_.PSIsContainer) { $null } else { $_.Length } } };
    }


    return Get-ChildItem -Path $Path -Force |
    Microsoft.PowerShell.Utility\Select-Object @{Name = "Caption"; Expression = { $_.FullName } },
    @{Name = "CreationDate"; Expression = { $_.CreationTimeUtc } },
    Extension,
    @{Name = "IsHidden"; Expression = { $_.Attributes -match "Hidden" } },
    @{Name = "IsShared"; Expression = { [bool]($folderShares | Where-Object Path -eq $_.FullName) } },
    Name,
    @{Name = "Type"; Expression = { Get-FileSystemEntityType -Attributes $_.Attributes } },
    @{Name = "LastModifiedDate"; Expression = { $_.LastWriteTimeUtc } },
    @{Name = "Size"; Expression = { if ($_.PSIsContainer) { $null } else { $_.Length } } };
}

<#
.Synopsis
    Name: Get-FileSystemEntityType
    Description: Gets the type of a local file system entity.

.Parameter Attributes
    The System.IO.FileAttributes of the FileSystemEntity.

.Returns
    The type of the local file system entity.
#>
function Get-FileSystemEntityType {
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.FileAttributes]
        $Attributes
    )

    if ($Attributes -match "Directory") {
        return "Folder";
    }
    else {
        return "File";
    }
}

$entities = Get-FileSystemEntities -Path $Path;
if ($OnlyFiles -and $OnlyFolders) {
    return $entities;
}

if ($OnlyFiles) {
    return $entities | Where-Object { $_.Type -eq "File" };
}

if ($OnlyFolders) {
    return $entities | Where-Object { $_.Type -eq "Folder" };
}

return $entities;

}
## [END] Get-WACCOFileSystemEntities ##
function Get-WACCOFileSystemRoot {
<#

.SYNOPSIS
Enumerates the root of the file system (volumes and related entities) of the system.

.DESCRIPTION
Enumerates the root of the file system (volumes and related entities) of the system on this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0
import-module CimCmdlets

<#
.Synopsis
    Name: Get-FileSystemRoot
    Description: Gets the local file system root entities of the machine.

.Returns
    The local file system root entities.
#>
function Get-FileSystemRoot
{
    $volumes = Enumerate-Volumes;

    return $volumes |
        Microsoft.PowerShell.Utility\Select-Object @{Name="Caption"; Expression={$_.DriveLetter +":\"}},
                      @{Name="CreationDate"; Expression={$null}},
                      @{Name="Extension"; Expression={$null}},
                      @{Name="IsHidden"; Expression={$false}},
                      @{Name="Name"; Expression={if ($_.FileSystemLabel) { $_.FileSystemLabel + " (" + $_.DriveLetter + ":)"} else { "(" + $_.DriveLetter + ":)" }}},
                      @{Name="Type"; Expression={"Volume"}},
                      @{Name="LastModifiedDate"; Expression={$null}},
                      @{Name="Size"; Expression={$_.Size}},
                      @{Name="SizeRemaining"; Expression={$_.SizeRemaining}}
}

<#
.Synopsis
    Name: Get-Volumes
    Description: Gets the local volumes of the machine.

.Returns
    The local volumes.
#>
function Enumerate-Volumes
{
    Remove-Module Storage -ErrorAction Ignore; # Remove the Storage module to prevent it from automatically localizing

    $isDownlevel = [Environment]::OSVersion.Version.Major -lt 10;
    if ($isDownlevel)
    {
        $disks = Get-CimInstance -ClassName MSFT_Disk -Namespace root/Microsoft/Windows/Storage | Where-Object { !$_.IsClustered };
        $partitions = @($disks | Get-CimAssociatedInstance -ResultClassName MSFT_Partition)
        if ($partitions.Length -eq 0) {
            $volumes = Get-CimInstance -ClassName MSFT_Volume -Namespace root/Microsoft/Windows/Storage;
        } else {
            $volumes = $partitions | Get-CimAssociatedInstance -ResultClassName MSFT_Volume;
        }
    }
    else
    {
        $subsystem = Get-CimInstance -ClassName MSFT_StorageSubSystem -Namespace root/Microsoft/Windows/Storage| Where-Object { $_.FriendlyName -like "*Win*" };
        $volumes = $subsystem | Get-CimAssociatedInstance -ResultClassName MSFT_Volume;
    }

    return $volumes | Where-Object {
        try {
            [byte]$_.DriveLetter -ne 0 -and $_.DriveLetter -ne $null -and $_.Size -gt 0
        } catch {
            $false
        }
    };
}

Get-FileSystemRoot;

}
## [END] Get-WACCOFileSystemRoot ##
function Get-WACCOFolderItemCount {
<#

.SYNOPSIS
Gets the count of elements in the folder

.DESCRIPTION
Gets the count of elements in the folder
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

.PARAMETER Path
    String -- The path to the folder

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path
)

Set-StrictMode -Version 5.0

$directoryInfo = Get-ChildItem $Path | Microsoft.PowerShell.Utility\Measure-Object
$directoryInfo.count

}
## [END] Get-WACCOFolderItemCount ##
function Get-WACCOFolderOwner {
<#

.SYNOPSIS
Gets the owner of a folder.

.DESCRIPTION
Gets the owner of a folder.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path
)

Set-StrictMode -Version 5.0

$Owner = (Get-Acl $Path).Owner
@{ owner = $Owner; }

}
## [END] Get-WACCOFolderOwner ##
function Get-WACCOFolderShareNames {
<#

.SYNOPSIS
Gets the existing share names of a shared folder

.DESCRIPTION
Gets the existing share names of a shared folder
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

.PARAMETER Path
    String -- The path to the folder.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path
)

Set-StrictMode -Version 5.0

Get-CimInstance -Class Win32_Share -Filter Path="'$Path'" | Microsoft.PowerShell.Utility\Select-Object Name

}
## [END] Get-WACCOFolderShareNames ##
function Get-WACCOFolderSharePath {
<#

.SYNOPSIS
Gets the existing share names of a shared folder

.DESCRIPTION
Gets the existing share names of a shared folder
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

.PARAMETER Name
    String -- The share name to the shared folder.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Name
)

Set-StrictMode -Version 5.0

Get-SmbShare -Includehidden | Where-Object { $_.Name -eq $Name } | Microsoft.PowerShell.Utility\Select-Object Path

}
## [END] Get-WACCOFolderSharePath ##
function Get-WACCOFolderShareStatus {
<#

.SYNOPSIS
Checks if a folder is shared

.DESCRIPTION
Checks if a folder is shared
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

.PARAMETER Path
    String -- the path to the folder.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path
)

Set-StrictMode -Version 5.0

$Shared = [bool](Get-CimInstance -Class Win32_Share -Filter Path="'$Path'")
@{ isShared = $Shared }

}
## [END] Get-WACCOFolderShareStatus ##
function Get-WACCOFolderShareUsers {
<#

.SYNOPSIS
Gets the user access rights of a folder

.DESCRIPTION
Gets the user access rights of a folder
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

.PARAMETER Path
    String -- The path to the folder.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path
)

Set-StrictMode -Version 5.0

Get-Acl $Path |  Microsoft.PowerShell.Utility\Select-Object -ExpandProperty Access | Microsoft.PowerShell.Utility\Select-Object IdentityReference, FileSystemRights, AccessControlType

}
## [END] Get-WACCOFolderShareUsers ##
function Get-WACCOIsAzureTurbineServer {
<#
.SYNOPSIS
Checks if the current server is Azure Turbine edition.

.DESCRIPTION
Returns true if the current server is Azure Turbine which supports smb over QUIC.

.ROLE
Readers

#>

$result = Get-WmiObject -Class Win32_OperatingSystem | Microsoft.PowerShell.Utility\Select-Object OperatingSystemSKU, Version

$result.OperatingSystemSKU -eq "407" -and [version]$result.version -ge [version]"10.0.20348"

}
## [END] Get-WACCOIsAzureTurbineServer ##
function Get-WACCOItemProperties {
<#

.SYNOPSIS
Get item's properties.

.DESCRIPTION
Get item's properties on this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

.PARAMETER Path
    String -- the path to the item whose properites are requested.

.PARAMETER ItemType
    String -- What kind of item?

#>

param (
    [Parameter(Mandatory = $true)]
    [String[]]
    $Path,

    [Parameter(Mandatory = $true)]
    [String]
    $ItemType
)

Set-StrictMode -Version 5.0

switch ($ItemType) {
    0 {
        Get-Volume $Path | Microsoft.PowerShell.Utility\Select-Object -Property *
    }
    default {
        Get-ItemProperty -Path $Path | Microsoft.PowerShell.Utility\Select-Object -Property *
    }
}

}
## [END] Get-WACCOItemProperties ##
function Get-WACCOItemType {
<#

.SYNOPSIS
Enumerates all of the file system entities (files, folders, volumes) of the system.

.DESCRIPTION
Enumerates all of the file system entities (files, folders, volumes) of the system on this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

.PARAMETER Path
    String -- the path to the folder where enumeration should start.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path
)

Set-StrictMode -Version 5.0

<#
.Synopsis
    Name: Get-FileSystemEntityType
    Description: Gets the type of a local file system entity.

.Parameter Attributes
    The System.IO.FileAttributes of the FileSystemEntity.

.Returns
    The type of the local file system entity.
#>
function Get-FileSystemEntityType
{
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.FileAttributes]
        $Attributes
    )

    if ($Attributes -match "Directory")
    {
        return "Folder";
    }
    else
    {
        return "File";
    }
}

if (Test-Path -LiteralPath $Path) {
    return Get-FileSystemEntityType -Attributes (Get-Item $Path).Attributes
} else {
    return ''
}

}
## [END] Get-WACCOItemType ##
function Get-WACCOLocalGroups {
<#

.SYNOPSIS
Gets the local groups.

.DESCRIPTION
Gets the local groups. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>

Import-Module Microsoft.PowerShell.LocalAccounts -ErrorAction SilentlyContinue

$isWinServer2016OrNewer = [Environment]::OSVersion.Version.Major -ge 10;
# ADSI does NOT support 2016 Nano, meanwhile New-LocalUser, Get-LocalUser, Set-LocalUser do NOT support downlevel

    if ($isWinServer2016OrNewer)
    {
       return  Get-LocalGroup | Sort-Object -Property Name | Microsoft.PowerShell.Utility\Select-Object Name
                                
    }
    else
    {
       return  Get-WmiObject -Class Win32_Group -Filter "LocalAccount='True'" | Sort-Object -Property Name | Microsoft.PowerShell.Utility\Select-Object Name
    }


}
## [END] Get-WACCOLocalGroups ##
function Get-WACCOLocalUsers {
<#

.SYNOPSIS
Gets the local users.

.DESCRIPTION
Gets the local users. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>


$isWinServer2016OrNewer = [Environment]::OSVersion.Version.Major -ge 10;

if ($isWinServer2016OrNewer){

	return Get-LocalUser | Sort-Object -Property Name | Microsoft.PowerShell.Utility\Select-Object Name

}
else{
    return Get-WmiObject -Class Win32_UserAccount -Filter "LocalAccount='True'" | Sort-Object -Property Name | Microsoft.PowerShell.Utility\Select-Object  Name;
}
}
## [END] Get-WACCOLocalUsers ##
function Get-WACCOOSVersion {
<#

.SYNOPSIS
Get OS Version

.DESCRIPTION
Get OS Version

.ROLE
Readers

#>

$result = Get-WmiObject -Class Win32_OperatingSystem | Microsoft.PowerShell.Utility\Select-Object Version

return [version]$result.version -ge [version]"10.0.20348"

}
## [END] Get-WACCOOSVersion ##
function Get-WACCOServerInventory {
<#

.SYNOPSIS
Retrieves the inventory data for a server.

.DESCRIPTION
Retrieves the inventory data for a server.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

Import-Module CimCmdlets

Import-Module Storage -ErrorAction SilentlyContinue

<#

.SYNOPSIS
Converts an arbitrary version string into just 'Major.Minor'

.DESCRIPTION
To make OS version comparisons we only want to compare the major and
minor version.  Build number and/os CSD are not interesting.

#>

function convertOsVersion([string]$osVersion) {
  [Ref]$parsedVersion = $null
  if (![Version]::TryParse($osVersion, $parsedVersion)) {
    return $null
  }

  $version = [Version]$parsedVersion.Value
  return New-Object Version -ArgumentList $version.Major, $version.Minor
}

<#

.SYNOPSIS
Determines if CredSSP is enabled for the current server or client.

.DESCRIPTION
Check the registry value for the CredSSP enabled state.

#>

function isCredSSPEnabled() {
  Set-Variable credSSPServicePath -Option Constant -Value "WSMan:\localhost\Service\Auth\CredSSP"
  Set-Variable credSSPClientPath -Option Constant -Value "WSMan:\localhost\Client\Auth\CredSSP"

  $credSSPServerEnabled = $false;
  $credSSPClientEnabled = $false;

  $credSSPServerService = Get-Item $credSSPServicePath -ErrorAction SilentlyContinue
  if ($credSSPServerService) {
    $credSSPServerEnabled = [System.Convert]::ToBoolean($credSSPServerService.Value)
  }

  $credSSPClientService = Get-Item $credSSPClientPath -ErrorAction SilentlyContinue
  if ($credSSPClientService) {
    $credSSPClientEnabled = [System.Convert]::ToBoolean($credSSPClientService.Value)
  }

  return ($credSSPServerEnabled -or $credSSPClientEnabled)
}

<#

.SYNOPSIS
Determines if the Hyper-V role is installed for the current server or client.

.DESCRIPTION
The Hyper-V role is installed when the VMMS service is available.  This is much
faster then checking Get-WindowsFeature and works on Windows Client SKUs.

#>

function isHyperVRoleInstalled() {
  $vmmsService = Get-Service -Name "VMMS" -ErrorAction SilentlyContinue

  return $vmmsService -and $vmmsService.Name -eq "VMMS"
}

<#

.SYNOPSIS
Determines if the Hyper-V PowerShell support module is installed for the current server or client.

.DESCRIPTION
The Hyper-V PowerShell support module is installed when the modules cmdlets are available.  This is much
faster then checking Get-WindowsFeature and works on Windows Client SKUs.

#>
function isHyperVPowerShellSupportInstalled() {
  # quicker way to find the module existence. it doesn't load the module.
  return !!(Get-Module -ListAvailable Hyper-V -ErrorAction SilentlyContinue)
}

<#

.SYNOPSIS
Determines if Windows Management Framework (WMF) 5.0, or higher, is installed for the current server or client.

.DESCRIPTION
Windows Admin Center requires WMF 5 so check the registey for WMF version on Windows versions that are less than
Windows Server 2016.

#>
function isWMF5Installed([string] $operatingSystemVersion) {
  Set-Variable Server2016 -Option Constant -Value (New-Object Version '10.0')   # And Windows 10 client SKUs
  Set-Variable Server2012 -Option Constant -Value (New-Object Version '6.2')

  $version = convertOsVersion $operatingSystemVersion
  if (-not $version) {
    # Since the OS version string is not properly formatted we cannot know the true installed state.
    return $false
  }

  if ($version -ge $Server2016) {
    # It's okay to assume that 2016 and up comes with WMF 5 or higher installed
    return $true
  }
  else {
    if ($version -ge $Server2012) {
      # Windows 2012/2012R2 are supported as long as WMF 5 or higher is installed
      $registryKey = 'HKLM:\SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine'
      $registryKeyValue = Get-ItemProperty -Path $registryKey -Name PowerShellVersion -ErrorAction SilentlyContinue

      if ($registryKeyValue -and ($registryKeyValue.PowerShellVersion.Length -ne 0)) {
        $installedWmfVersion = [Version]$registryKeyValue.PowerShellVersion

        if ($installedWmfVersion -ge [Version]'5.0') {
          return $true
        }
      }
    }
  }

  return $false
}

<#

.SYNOPSIS
Determines if the current usser is a system administrator of the current server or client.

.DESCRIPTION
Determines if the current usser is a system administrator of the current server or client.

#>
function isUserAnAdministrator() {
  return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}

<#

.SYNOPSIS
Get some basic information about the Failover Cluster that is running on this server.

.DESCRIPTION
Create a basic inventory of the Failover Cluster that may be running in this server.

#>
function getClusterInformation() {
  $returnValues = @{ }

  $returnValues.IsS2dEnabled = $false
  $returnValues.IsCluster = $false
  $returnValues.ClusterFqdn = $null
  $returnValues.IsBritannicaEnabled = $false

  $namespace = Get-CimInstance -Namespace root/MSCluster -ClassName __NAMESPACE -ErrorAction SilentlyContinue
  if ($namespace) {
    $cluster = Get-CimInstance -Namespace root/MSCluster -ClassName MSCluster_Cluster -ErrorAction SilentlyContinue
    if ($cluster) {
      $returnValues.IsCluster = $true
      $returnValues.ClusterFqdn = $cluster.Fqdn
      $returnValues.IsS2dEnabled = !!(Get-Member -InputObject $cluster -Name "S2DEnabled") -and ($cluster.S2DEnabled -gt 0)
      $returnValues.IsBritannicaEnabled = $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_Cluster -ErrorAction SilentlyContinue)
    }
  }

  return $returnValues
}

<#

.SYNOPSIS
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the passed in computer name.

.DESCRIPTION
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the passed in computer name.

#>
function getComputerFqdnAndAddress($computerName) {
  $hostEntry = [System.Net.Dns]::GetHostEntry($computerName)
  $addressList = @()
  foreach ($item in $hostEntry.AddressList) {
    $address = New-Object PSObject
    $address | Add-Member -MemberType NoteProperty -Name 'IpAddress' -Value $item.ToString()
    $address | Add-Member -MemberType NoteProperty -Name 'AddressFamily' -Value $item.AddressFamily.ToString()
    $addressList += $address
  }

  $result = New-Object PSObject
  $result | Add-Member -MemberType NoteProperty -Name 'Fqdn' -Value $hostEntry.HostName
  $result | Add-Member -MemberType NoteProperty -Name 'AddressList' -Value $addressList
  return $result
}

<#

.SYNOPSIS
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the current server or client.

.DESCRIPTION
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the current server or client.

#>
function getHostFqdnAndAddress($computerSystem) {
  $computerName = $computerSystem.DNSHostName
  if (!$computerName) {
    $computerName = $computerSystem.Name
  }

  return getComputerFqdnAndAddress $computerName
}

<#

.SYNOPSIS
Are the needed management CIM interfaces available on the current server or client.

.DESCRIPTION
Check for the presence of the required server management CIM interfaces.

#>
function getManagementToolsSupportInformation() {
  $returnValues = @{ }

  $returnValues.ManagementToolsAvailable = $false
  $returnValues.ServerManagerAvailable = $false

  $namespaces = Get-CimInstance -Namespace root/microsoft/windows -ClassName __NAMESPACE -ErrorAction SilentlyContinue

  if ($namespaces) {
    $returnValues.ManagementToolsAvailable = !!($namespaces | Where-Object { $_.Name -ieq "ManagementTools" })
    $returnValues.ServerManagerAvailable = !!($namespaces | Where-Object { $_.Name -ieq "ServerManager" })
  }

  return $returnValues
}

<#

.SYNOPSIS
Check the remote app enabled or not.

.DESCRIPTION
Check the remote app enabled or not.

#>
function isRemoteAppEnabled() {
  Set-Variable key -Option Constant -Value "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Terminal Server\\TSAppAllowList"

  $registryKeyValue = Get-ItemProperty -Path $key -Name fDisabledAllowList -ErrorAction SilentlyContinue

  if (-not $registryKeyValue) {
    return $false
  }
  return $registryKeyValue.fDisabledAllowList -eq 1
}

<#

.SYNOPSIS
Check the remote app enabled or not.

.DESCRIPTION
Check the remote app enabled or not.

#>

<#
c
.SYNOPSIS
Get the Win32_OperatingSystem information as well as current version information from the registry

.DESCRIPTION
Get the Win32_OperatingSystem instance and filter the results to just the required properties.
This filtering will make the response payload much smaller. Included in the results are current version
information from the registry

#>
function getOperatingSystemInfo() {
  $operatingSystemInfo = Get-CimInstance Win32_OperatingSystem | Microsoft.PowerShell.Utility\Select-Object csName, Caption, OperatingSystemSKU, Version, ProductType, OSType, LastBootUpTime
  $currentVersion = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" | Microsoft.PowerShell.Utility\Select-Object CurrentBuild, UBR, DisplayVersion

  $operatingSystemInfo | Add-Member -MemberType NoteProperty -Name CurrentBuild -Value $currentVersion.CurrentBuild
  $operatingSystemInfo | Add-Member -MemberType NoteProperty -Name UpdateBuildRevision -Value $currentVersion.UBR
  $operatingSystemInfo | Add-Member -MemberType NoteProperty -Name DisplayVersion -Value $currentVersion.DisplayVersion

  return $operatingSystemInfo
}

<#

.SYNOPSIS
Get the Win32_ComputerSystem information

.DESCRIPTION
Get the Win32_ComputerSystem instance and filter the results to just the required properties.
This filtering will make the response payload much smaller.

#>
function getComputerSystemInfo() {
  return Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue | `
    Microsoft.PowerShell.Utility\Select-Object TotalPhysicalMemory, DomainRole, Manufacturer, Model, NumberOfLogicalProcessors, Domain, Workgroup, DNSHostName, Name, PartOfDomain, SystemFamily, SystemSKUNumber
}

<#

.SYNOPSIS
Get SMBIOS locally from the passed in machineName


.DESCRIPTION
Get SMBIOS locally from the passed in machine name

#>
function getSmbiosData($computerSystem) {
  <#
    Array of chassis types.
    The following list of ChassisTypes is copied from the latest DMTF SMBIOS specification.
    REF: https://www.dmtf.org/sites/default/files/standards/documents/DSP0134_3.1.1.pdf
  #>
  $ChassisTypes =
  @{
    1  = 'Other'
    2  = 'Unknown'
    3  = 'Desktop'
    4  = 'Low Profile Desktop'
    5  = 'Pizza Box'
    6  = 'Mini Tower'
    7  = 'Tower'
    8  = 'Portable'
    9  = 'Laptop'
    10 = 'Notebook'
    11 = 'Hand Held'
    12 = 'Docking Station'
    13 = 'All in One'
    14 = 'Sub Notebook'
    15 = 'Space-Saving'
    16 = 'Lunch Box'
    17 = 'Main System Chassis'
    18 = 'Expansion Chassis'
    19 = 'SubChassis'
    20 = 'Bus Expansion Chassis'
    21 = 'Peripheral Chassis'
    22 = 'Storage Chassis'
    23 = 'Rack Mount Chassis'
    24 = 'Sealed-Case PC'
    25 = 'Multi-system chassis'
    26 = 'Compact PCI'
    27 = 'Advanced TCA'
    28 = 'Blade'
    29 = 'Blade Enclosure'
    30 = 'Tablet'
    31 = 'Convertible'
    32 = 'Detachable'
    33 = 'IoT Gateway'
    34 = 'Embedded PC'
    35 = 'Mini PC'
    36 = 'Stick PC'
  }

  $list = New-Object System.Collections.ArrayList
  $win32_Bios = Get-CimInstance -class Win32_Bios
  $obj = New-Object -Type PSObject | Microsoft.PowerShell.Utility\Select-Object SerialNumber, Manufacturer, UUID, BaseBoardProduct, ChassisTypes, Chassis, SystemFamily, SystemSKUNumber, SMBIOSAssetTag
  $obj.SerialNumber = $win32_Bios.SerialNumber
  $obj.Manufacturer = $win32_Bios.Manufacturer
  $computerSystemProduct = Get-CimInstance Win32_ComputerSystemProduct
  if ($null -ne $computerSystemProduct) {
    $obj.UUID = $computerSystemProduct.UUID
  }
  $baseboard = Get-CimInstance Win32_BaseBoard
  if ($null -ne $baseboard) {
    $obj.BaseBoardProduct = $baseboard.Product
  }
  $systemEnclosure = Get-CimInstance Win32_SystemEnclosure
  if ($null -ne $systemEnclosure) {
    $obj.SMBIOSAssetTag = $systemEnclosure.SMBIOSAssetTag
  }
  $obj.ChassisTypes = Get-CimInstance Win32_SystemEnclosure | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty ChassisTypes
  $obj.Chassis = New-Object -TypeName 'System.Collections.ArrayList'
  $obj.ChassisTypes | ForEach-Object -Process {
    $obj.Chassis.Add($ChassisTypes[[int]$_])
  }
  $obj.SystemFamily = $computerSystem.SystemFamily
  $obj.SystemSKUNumber = $computerSystem.SystemSKUNumber
  $list.Add($obj) | Out-Null

  return $list

}

<#

.SYNOPSIS
Checks if a system lockdown policy is enforced on this computer.

.DESCRIPTION
By checking the system lockdown policy, we can infer that PowerShell is in ConstrainedLanguage mode (WDAC).
Note: $ExecutionContext.SessionState.LanguageMode should not be used within a script context as in the case of allowed scripts (by the WDAC policy),
being executed locally, the language mode will always be FullLanguage and does NOT reflect the default system lockdown policy/language mode.

#>
function isSystemLockdownPolicyEnforced() {
  return [System.Management.Automation.Security.SystemPolicy]::GetSystemLockdownPolicy() -eq [System.Management.Automation.Security.SystemEnforcementMode]::Enforce
}

###########################################################################
# main()
###########################################################################

$operatingSystem = getOperatingSystemInfo
$computerSystem = getComputerSystemInfo
$isAdministrator = isUserAnAdministrator
$fqdnAndAddress = getHostFqdnAndAddress $computerSystem
$hostname = [Environment]::MachineName
$netbios = $env:ComputerName
$managementToolsInformation = getManagementToolsSupportInformation
$isWmfInstalled = isWMF5Installed $operatingSystem.Version
$clusterInformation = getClusterInformation -ErrorAction SilentlyContinue
$isHyperVPowershellInstalled = isHyperVPowerShellSupportInstalled
$isHyperVRoleInstalled = isHyperVRoleInstalled
$isCredSSPEnabled = isCredSSPEnabled
$isRemoteAppEnabled = isRemoteAppEnabled
$smbiosData = getSmbiosData $computerSystem
$isSystemLockdownPolicyEnforced = isSystemLockdownPolicyEnforced

$result = New-Object PSObject
$result | Add-Member -MemberType NoteProperty -Name 'IsAdministrator' -Value $isAdministrator
$result | Add-Member -MemberType NoteProperty -Name 'OperatingSystem' -Value $operatingSystem
$result | Add-Member -MemberType NoteProperty -Name 'ComputerSystem' -Value $computerSystem
$result | Add-Member -MemberType NoteProperty -Name 'Fqdn' -Value $fqdnAndAddress.Fqdn
$result | Add-Member -MemberType NoteProperty -Name 'AddressList' -Value $fqdnAndAddress.AddressList
$result | Add-Member -MemberType NoteProperty -Name 'Hostname' -Value $hostname
$result | Add-Member -MemberType NoteProperty -Name 'NetBios' -Value $netbios
$result | Add-Member -MemberType NoteProperty -Name 'IsManagementToolsAvailable' -Value $managementToolsInformation.ManagementToolsAvailable
$result | Add-Member -MemberType NoteProperty -Name 'IsServerManagerAvailable' -Value $managementToolsInformation.ServerManagerAvailable
$result | Add-Member -MemberType NoteProperty -Name 'IsWmfInstalled' -Value $isWmfInstalled
$result | Add-Member -MemberType NoteProperty -Name 'IsCluster' -Value $clusterInformation.IsCluster
$result | Add-Member -MemberType NoteProperty -Name 'ClusterFqdn' -Value $clusterInformation.ClusterFqdn
$result | Add-Member -MemberType NoteProperty -Name 'IsS2dEnabled' -Value $clusterInformation.IsS2dEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaEnabled' -Value $clusterInformation.IsBritannicaEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsHyperVRoleInstalled' -Value $isHyperVRoleInstalled
$result | Add-Member -MemberType NoteProperty -Name 'IsHyperVPowershellInstalled' -Value $isHyperVPowershellInstalled
$result | Add-Member -MemberType NoteProperty -Name 'IsCredSSPEnabled' -Value $isCredSSPEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsRemoteAppEnabled' -Value $isRemoteAppEnabled
$result | Add-Member -MemberType NoteProperty -Name 'SmbiosData' -Value $smbiosData
$result | Add-Member -MemberType NoteProperty -Name 'IsSystemLockdownPolicyEnforced' -Value $isSystemLockdownPolicyEnforced

$result

}
## [END] Get-WACCOServerInventory ##
function Get-WACCOShareEntities {
<#

.SYNOPSIS
Enumerates all of the file system entities (files, folders, volumes) of the system.

.DESCRIPTION
Enumerates all of the file system entities (files, folders, volumes) of the system on this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

.PARAMETER Path
    String -- The path to enumerate.

.PARAMETER OnlyFiles
    switch --

.PARAMETER OnlyFolders
    switch --

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $ComputerName
)

Set-StrictMode -Version 5.0

<#
.Synopsis
    Name: Get-FileSystemEntities
    Description: Gets all the local file system entities of the machine.

.Parameter Path
    String -- The path to enumerate.

.Returns
    The local file system entities.
#>
function Get-FileSystemEntities {
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $ComputerName
    )

    return Invoke-Command -ComputerName $ComputerName -ScriptBlock { get-smbshare | Where-Object { -not ($_.name.EndsWith('$')) } } |
    Microsoft.PowerShell.Utility\Select-Object @{Name = "Caption"; Expression = { "\\" + $_.PSComputerName + "\" + $_.Name } },
    @{Name = "CreationDate"; Expression = { $_.CreationTimeUtc } },
    Extension,
    @{Name = "IsHidden"; Expression = { [bool]$false } },
    @{Name = "IsShared"; Expression = { [bool]$true } },
    Name,
    @{Name = "Type"; Expression = { "FileShare" } },
    @{Name = "LastModifiedDate"; Expression = { $_.LastWriteTimeUtc } },
    @{Name = "Size"; Expression = { if ($_.PSIsContainer) { $null } else { $_.Length } } };
}

$entities = Get-FileSystemEntities -ComputerName $ComputerName;

return $entities;

}
## [END] Get-WACCOShareEntities ##
function Get-WACCOSmb1InstallationStatus {
<#

.SYNOPSIS
Get SMB1 installation status.

.DESCRIPTION
Get SMB1 installation status.

.ROLE
Readers

#>

Import-Module DISM

$Enabled = [bool]( Get-WindowsOptionalFeature -online -featurename SMB1Protocol | Where-Object State -eq "Enabled")
@{ isEnabled = $Enabled }

}
## [END] Get-WACCOSmb1InstallationStatus ##
function Get-WACCOSmbFileShareDetails {
<#

.SYNOPSIS
Enumerates all of the smb local file shares of the system.

.DESCRIPTION
Enumerates all of the smb local file shares of the system.

.ROLE
Readers

#>

<#
.Synopsis
    Name: Get-SmbFileShareDetails
    Description: Retrieves the SMB shares on the computer.
.Returns
    The local smb file share(s).
#>

$shares = Get-SmbShare -includehidden | Where-Object {-not ($_.Name -eq "IPC$")} | Microsoft.PowerShell.Utility\Select-Object Name, Path, CachingMode, EncryptData, CurrentUsers, Special, LeasingMode, FolderEnumerationMode, CompressData

$uncPath = (Get-CimInstance Win32_ComputerSystem).DNSHostName + "." + (Get-CimInstance Win32_ComputerSystem).Domain

return @{
    shares = $shares;
    uncPath = $uncPath
}
   
}
## [END] Get-WACCOSmbFileShareDetails ##
function Get-WACCOSmbOverQuicSettings {
<#

.SYNOPSIS
Retrieves smb over QUIC settings from the server.

.DESCRIPTION
Returns smb over QUIC settings and server dns name

.ROLE
Readers

#>

Import-Module SmbShare

$serverConfigurationSettings = Get-SmbServerConfiguration | Microsoft.PowerShell.Utility\Select-Object DisableSmbEncryptionOnSecureConnection, RestrictNamedpipeAccessViaQuic

return @{
    serverConfigurationSettings = $serverConfigurationSettings
}

}
## [END] Get-WACCOSmbOverQuicSettings ##
function Get-WACCOSmbServerCertificateHealth {
<#

.SYNOPSIS
Retrieves health of the current certificate for smb over QUIC.

.DESCRIPTION
Retrieves health of the current certificate for smb over QUIC based on if the certificate is self signed or not.

.ROLE
Readers

#>

param (
  [Parameter(Mandatory = $true)]
  [String]
  $thumbprint,
  [Parameter(Mandatory = $true)]
  [boolean]
  $isSelfSigned,
  [Parameter(Mandatory = $true)]
  [boolean]
  $fromTaskScheduler,
  [Parameter(Mandatory = $false)]
  [String]
  $ResultFile,
  [Parameter(Mandatory = $false)]
  [String]
  $WarningsFile,
  [Parameter(Mandatory = $false)]
  [String]
  $ErrorFile
)

Set-StrictMode -Version 5.0



function getSmbServerCertificateHealth() {
  param (
    [String]
    $thumbprint,
    [boolean]
    $isSelfSigned,
    [String]
    $ResultFile,
    [String]
    $WarningsFile,
    [String]
    $ErrorFile
  )

  # create local runspace
  $ps = [PowerShell]::Create()
  # define input data but make it completed
  $inData = New-Object -Typename  System.Management.Automation.PSDataCollection[PSObject]
  $inData.Complete()
  # define output data to receive output
  $outData = New-Object -Typename  System.Management.Automation.PSDataCollection[PSObject]
  # register the script
  if ($isSelfSigned) {
    $ps.Commands.AddScript("Get-Item -Path " + $thumbprint + "| Test-Certificate -AllowUntrustedRoot") | Out-Null
  }
  else {
    $ps.Commands.AddScript("Get-Item -Path " + $thumbprint + "| Test-Certificate") | Out-Null
  }
  # execute async way.
  $async = $ps.BeginInvoke($inData, $outData)
  # wait for completion (callback will be called if any)
  $ps.EndInvoke($async)
  Start-Sleep -MilliSeconds 10
  # read output
  if ($outData.Count -gt 0) {
    @{ Output = $outData[0]; } | ConvertTo-Json | Out-File -FilePath $ResultFile
  }
  # read warnings
  if ($ps.Streams.Warning.Count -gt 0) {
    $ps.Streams.Warning | % { $_.ToString() } | Out-File -FilePath $WarningsFile
  }
  # read errors
  if ($ps.HadErrors) {
    $ps.Streams.Error | % { $_.ToString() } | Out-File -FilePath $ErrorFile
  }
}

#---- Script execution starts here ----
$isWdacEnabled = $false;
Import-Module CimCmdlets -ErrorAction SilentlyContinue
$dg = Get-CimInstance -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard
$psLanguageMode = $ExecutionContext.SessionState.LanguageMode
if (($psLanguageMode -eq 'ConstrainedLanguage') -and ($dg.CodeIntegrityPolicyEnforcementStatus -eq 2) -and ($dg.UsermodeCodeIntegrityPolicyEnforcementStatus -eq 2)) {
  $isWdacEnabled = $true;
}

#In WDAC environment script file will already be available on the machine
#In WDAC mode the same script is executed - once normally and once through task Scheduler
if ($isWdacEnabled) {
  if ($fromTaskScheduler) {
    getSmbServerCertificateHealth $thumbprint $isSelfSigned $ResultFile $WarningsFile $ErrorFile;
    return;
  }
}
else {
  #In non-WDAC environment script file will not be available on the machine
  #Hence, a dynamic script is created which is executed through the task Scheduler
  $ScriptFile = $env:temp + "\smbOverQuic-certificateHealth.ps1"
}

$thumbprint = Join-Path "Cert:\LocalMachine\My" $thumbprint

# Pass parameters tpt and generate script file in temp folder
$ResultFile = $env:temp + "\smbOverQuic-certificateHealth_result.txt"
$WarningsFile = $env:temp + "\smbOverQuic-certificateHealth_warnings.txt"
$ErrorFile = $env:temp + "\smbOverQuic-certificateHealth_error.txt"
if (Test-Path $ErrorFile) {
  Remove-Item $ErrorFile
}

if (Test-Path $ResultFile) {
  Remove-Item $ResultFile
}

if (Test-Path $WarningsFile) {
  Remove-Item $WarningsFile
}
$isSelfSignedtemp = if ($isSelfSigned) { "`$true" } else { "`$false" }

if ($isWdacEnabled) {
  $arg = "-NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -command ""&{Import-Module Microsoft.SME.FileExplorer; Get-WACFESmbServerCertificateHealth -fromTaskScheduler `$true -thumbprint $thumbprint -isSelfSigned $isSelfSignedtemp -ResultFile $ResultFile -WarningsFile $WarningsFile -ErrorFile $ErrorFile }"""
}
else {
  (Get-Command getSmbServerCertificateHealth).ScriptBlock | Set-Content -path $ScriptFile
  $arg = "-NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -command ""&{Set-Location -Path $env:temp; .\smbOverQuic-certificateHealth.ps1 -thumbprint $thumbprint -isSelfSigned $isSelfSignedtemp -ResultFile $ResultFile -WarningsFile $WarningsFile -ErrorFile $ErrorFile }"""
}

# Create a scheduled task
$TaskName = "SMESmbOverQuicCertificate"
$User = [Security.Principal.WindowsIdentity]::GetCurrent()
$Role = (New-Object Security.Principal.WindowsPrincipal $User).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if (!$Role) {
  Write-Warning "To perform some operations you must run an elevated Windows PowerShell console."
}

$Scheduler = New-Object -ComObject Schedule.Service
#Try to connect to schedule service 3 time since it may fail the first time
for ($i = 1; $i -le 3; $i++) {
  Try {
    $Scheduler.Connect()
    Break
  }
  Catch {
    if ($i -ge 3) {
      writeErrorLog "Can't connect to Schedule service"
      throw "Can't connect to Schedule service"
    }
    else {
      Start-Sleep -s 1
    }
  }
}

$RootFolder = $Scheduler.GetFolder("\")

#Delete existing task
if ($RootFolder.GetTasks(0) | Where-Object { $_.Name -eq $TaskName }) {
  Write-Debug("Deleting existing task" + $TaskName)
  try {
    $RootFolder.DeleteTask($TaskName, 0)
  }
  catch {

  }
}

$Task = $Scheduler.NewTask(0)
$RegistrationInfo = $Task.RegistrationInfo
$RegistrationInfo.Description = $TaskName
$RegistrationInfo.Author = $User.Name

$Triggers = $Task.Triggers
$Trigger = $Triggers.Create(7) #TASK_TRIGGER_REGISTRATION: Starts the task when the task is registered.
$Trigger.Enabled = $true

$Settings = $Task.Settings
$Settings.Enabled = $True
$Settings.StartWhenAvailable = $True
$Settings.Hidden = $False

$Action = $Task.Actions.Create(0)
$Action.Path = "powershell"
$Action.Arguments = $arg

#Tasks will be run with the highest privileges
$Task.Principal.RunLevel = 1

#### example Start the task with user specified invoke username and password
####$Task.Principal.LogonType = 1
####$RootFolder.RegisterTaskDefinition($TaskName, $Task, 6, $invokeUserName, $invokePassword, 1) | Out-Null

#### Start the task with SYSTEM creds
$RootFolder.RegisterTaskDefinition($TaskName, $Task, 6, "SYSTEM", $Null, 1) | Out-Null
#Wait for running task finished
$RootFolder.GetTask($TaskName).Run(0) | Out-Null
while ($RootFolder.GetTask($TaskName).State -ne 4) {
  Start-Sleep -MilliSeconds 10
}
while ($Scheduler.GetRunningTasks(0) | Where-Object { $_.Name -eq $TaskName }) {
  Start-Sleep -Seconds 1
}

#Clean up
try {
  $RootFolder.DeleteTask($TaskName, 0)
}
catch {

}
if (!$isWdacEnabled) {
  Remove-Item $ScriptFile
}

#Return result
if (Test-Path $ResultFile) {
  $result = Get-Content -Raw -Path $ResultFile | ConvertFrom-Json
  Remove-Item $ResultFile
  if ($result.Output) {
    return $result.Output;
  }
  else {
    if (Test-Path $WarningsFile) {
      $result = Get-Content -Path $WarningsFile
      Remove-Item $WarningsFile
    }
    if (Test-Path $ErrorFile) {
      Remove-Item $ErrorFile
    }
  }
  return $result;
}
else {
  if (Test-Path $ErrorFile) {
    $result = Get-Content -Path $ErrorFile
    Remove-Item $ErrorFile
    throw $result
  }
}

}
## [END] Get-WACCOSmbServerCertificateHealth ##
function Get-WACCOSmbServerCertificateMapping {
<#

.SYNOPSIS
Retrieves the current certifcate installed for smb over QUIC and smboverQuic status on the server.

.DESCRIPTION
Retrieves the current certifcate installed for smb over QUIC and smboverQuic status on the server.

.ROLE
Readers

#>

Import-Module SmbShare

$certHash = $null;
$kdcPort = $null;

function Retrieve-WACPort {
    $details = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\ServerManagementGateway -ErrorAction SilentlyContinue
    if($details){
        $smePort =  $details | Microsoft.PowerShell.Utility\Select-Object SmePort;
        return $smePort.smePort -eq 443
    } else {
        return $false;
    }
}

# retrieving smbcertificate mappings, if any
$smbCertificateMapping = @(Get-SmbServerCertificateMapping | Microsoft.PowerShell.Utility\Select-Object Thumbprint, Name)

# determining if WAC is installed on port 443
$isWacOnPort443 = Retrieve-WACPort;

# retrieving if smbOverQuic is enable on the server
$isSmbOverQuicEnabled = Get-SmbServerConfiguration | Microsoft.PowerShell.Utility\Select-Object EnableSMBQUIC

try {
    # retrieving kdc Proxy status on the server
    $kdcUrl = netsh http show urlacl | findstr /i "KdcProxy"
    if($kdcUrl) {
        $pos = $kdcUrl.IndexOf("+")
        $rightPart = $kdcUrl.Substring($pos+1)
        $pos1 = $rightPart.IndexOf("/")
        $kdcPort = $rightPart.SubString(1, $pos1-1)
    }

    [array]$path = Get-Item  -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -ErrorAction SilentlyContinue

    [array]$clientAuthproperty = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "HttpsClientAuth" -ErrorAction SilentlyContinue

    [array]$passwordAuthproperty = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "DisallowUnprotectedPasswordAuth" -ErrorAction SilentlyContinue

    $status = Get-Service -Name kpssvc | Microsoft.PowerShell.Utility\Select-Object Status

    if($null -ne $kdcPort){
        $port = "0.0.0.0:"
        $ipport = $port+$kdcPort
        $certBinding = @(netsh http show sslcert ipport=$ipport | findstr "Hash")
        if($null -ne $certBinding -and $certBinding.count -gt 0) {
            $index = $certBinding[0].IndexOf(":")
            $certHash = $certBinding[0].Substring($index+1).trim()
        }
    }
}
catch {
    throw $_
}

return @{
    smbCertificateMapping = $smbCertificateMapping
    isSmbOverQuicEnabled = $isSmbOverQuicEnabled
    isKdcProxyMappedForSmbOverQuic = if ($null -ne $certHash -and $null -ne $smbCertificateMapping -and $smbCertificateMapping.count -eq 1) { $smbCertificateMapping.thumbprint -eq $certHash } else { $false }
    kdcPort = $kdcPort
    isKdcProxyEnabled = $kdcPort -and ($null -ne $path -and $path.count -gt 0) -and ($null -ne $clientAuthproperty -and $clientAuthproperty.count -gt 0) -and ($null -ne $passwordAuthproperty -and $passwordAuthproperty.count -gt 0) -and $status.status -eq "Running"
    isWacOnPort443 = $isWacOnPort443;

}

}
## [END] Get-WACCOSmbServerCertificateMapping ##
function Get-WACCOSmbServerCertificateValues {
<#

.SYNOPSIS
Retrieves other values based on the installed certifcate for smb over QUIC.

.DESCRIPTION
Retrieves other values based on the installed certifcate for smb over QUIC.

.ROLE 
Readers

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $thumbprint
)

Set-Location Cert:\LocalMachine\My

[array] $smbServerDnsNames = Get-SmbServerCertificateMapping | Where-Object { $_.Thumbprint -eq $thumbprint } | Microsoft.PowerShell.Utility\Select-Object Name

$smbCertificateValues = Get-ChildItem -Recurse | Where-Object { $_.Thumbprint -eq $thumbprint } | Microsoft.PowerShell.Utility\Select-Object Subject, Thumbprint, Issuer, NotBefore, NotAfter

return @{ 
    smbServerDnsNames = $smbServerDnsNames
    smbCertificateValues = $smbCertificateValues
}

}
## [END] Get-WACCOSmbServerCertificateValues ##
function Get-WACCOSmbServerSettings {

<#

.SYNOPSIS
Enumerates the SMB server configuration settings on the computer.

.DESCRIPTION
Enumerates  the SMB server configuration settings on the computer.

.ROLE
Readers

#>

<#
.Synopsis
    Name: Get-SmbServerConfiguration
    Description: Retrieves the SMB server configuration settings on the computer.
.Returns
    SMB server configuration settings
#>

Import-Module SmbShare

$settings = Get-SmbServerConfiguration | Microsoft.PowerShell.Utility\Select-Object RequireSecuritySignature,RejectUnencryptedAccess,AuditSmb1Access,EncryptData
$compressionSettings = Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\LanManServer\parameters" -Name "DisableCompression" -ErrorAction SilentlyContinue

@{ settings = $settings
    compressionSettings = $compressionSettings }

}
## [END] Get-WACCOSmbServerSettings ##
function Get-WACCOSmbShareAccess {
<#

.SYNOPSIS
Enumerates the SMB server access rights and details on the server.

.DESCRIPTION
Enumerates the SMB server access rights and details on the server.

.ROLE
Readers

#>

<#
.Synopsis
    Name: Get-SmbShareAccess
    Description: Retrieves the SMB server access rights and details on the server.
.Returns
    Retrieves the SMB server access rights and details on the server.
#>
param (
    [Parameter(Mandatory = $true)]
    [String]
    $Name
)

[array]$shareAccess = Get-SmbShareAccess -Name "$Name" | Microsoft.PowerShell.Utility\Select-Object AccountName, AccessRight, AccessControlType

$details = Get-SmbShare -Name "$Name" | Microsoft.PowerShell.Utility\Select-Object CachingMode, EncryptData, FolderEnumerationMode, CompressData

return @{ 
    details = $details
    shareAccess = $shareAccess
  }   
}
## [END] Get-WACCOSmbShareAccess ##
function Get-WACCOStorageFileShare {
<#

.SYNOPSIS
Enumerates all of the local file shares of the system.

.DESCRIPTION
Enumerates all of the local file shares of the system.

.ROLE
Readers

.PARAMETER FileShareId
    The file share ID.
#>
param (
    [Parameter(Mandatory = $false)]
    [String]
    $FileShareId
)

Import-Module CimCmdlets

<#
.Synopsis
    Name: Get-FileShares-Internal
    Description: Gets all the local file shares of the machine.

.Parameters
    $FileShareId: The unique identifier of the file share desired (Optional - for cases where only one file share is desired).

.Returns
    The local file share(s).
#>
function Get-FileSharesInternal
{
    param (
        [Parameter(Mandatory = $false)]
        [String]
        $FileShareId
    )

    Remove-Module Storage -ErrorAction Ignore; # Remove the Storage module to prevent it from automatically localizing

    $isDownlevel = [Environment]::OSVersion.Version.Major -lt 10;
    if ($isDownlevel)
    {
        # Map downlevel status to array of [health status, operational status, share state] uplevel equivalent
        $statusMap = @{
            "OK" =         @(0, 2, 1);
            "Error" =      @(2, 6, 2);
            "Degraded" =   @(1, 3, 2);
            "Unknown" =    @(5, 0, 0);
            "Pred Fail" =  @(1, 5, 2);
            "Starting" =   @(1, 8, 0);
            "Stopping" =   @(1, 9, 0);
            "Service" =    @(1, 11, 1);
            "Stressed" =   @(1, 4, 1);
            "NonRecover" = @(2, 7, 2);
            "No Contact" = @(2, 12, 2);
            "Lost Comm" =  @(2, 13, 2);
        };
        
        $shares = Get-CimInstance -ClassName Win32_Share |
            ForEach-Object {
                return @{
                    ContinuouslyAvailable = $false;
                    Description = $_.Description;
                    EncryptData = $false;
                    FileSharingProtocol = 3;
                    HealthStatus = $statusMap[$_.Status][0];
                    IsHidden = $_.Name.EndsWith("`$");
                    Name = $_.Name;
                    OperationalStatus = ,@($statusMap[$_.Status][1]);
                    ShareState = $statusMap[$_.Status][2];
                    UniqueId = "smb|" + (Get-CimInstance Win32_ComputerSystem).DNSHostName + "." + (Get-CimInstance Win32_ComputerSystem).Domain + "\" + $_.Name;
                    VolumePath = $_.Path;
                }
            }
    }
    else
    {        
        $shares = Get-CimInstance -ClassName MSFT_FileShare -Namespace Root\Microsoft\Windows/Storage |
            ForEach-Object {
                return @{
                    IsHidden = $_.Name.EndsWith("`$");
                    VolumePath = $_.VolumeRelativePath;
                    ContinuouslyAvailable = $_.ContinuouslyAvailable;
                    Description = $_.Description;
                    EncryptData = $_.EncryptData;
                    FileSharingProtocol = $_.FileSharingProtocol;
                    HealthStatus = $_.HealthStatus;
                    Name = $_.Name;
                    OperationalStatus = $_.OperationalStatus;
                    UniqueId = $_.UniqueId;
                    ShareState = $_.ShareState;
                }
            }
    }

    if ($FileShareId)
    {
        $shares = $shares | Where-Object { $_.UniqueId -eq $FileShareId };
    }

    return $shares;
}

if ($FileShareId)
{
    Get-FileSharesInternal -FileShareId $FileShareId;
}
else
{
    Get-FileSharesInternal;
}

}
## [END] Get-WACCOStorageFileShare ##
function Get-WACCOTempFolderPath {
<#

.SYNOPSIS
Gets the temporary folder (%temp%) for the user.

.DESCRIPTION
Gets the temporary folder (%temp%) for the user.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>

Set-StrictMode -Version 5.0

return $env:TEMP

}
## [END] Get-WACCOTempFolderPath ##
function Move-WACCOFile {
<#

.SYNOPSIS
Moves or Copies a file or folder

.DESCRIPTION
Moves or Copies a file or folder from the source location to the destination location
Folders will be copied recursively

.ROLE
Administrators

.PARAMETER Path
    String -- the path to the source file/folder to copy

.PARAMETER Destination
    String -- the path to the new location

.PARAMETER Copy
    boolean -- Determine action to be performed

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $path,

    [Parameter(Mandatory = $true)]
    [String]
    $destination,

    [Parameter(Mandatory = $true)]
    [boolean]
    $copy,

    [Parameter(Mandatory = $true)]
    [string]
    $entityType,

    [Parameter(Mandatory = $false)]
    [string]
    $existingPath
)

Set-StrictMode -Version 5.0

if($copy){
  $result = Copy-Item -Path $path -Destination $destination -Recurse -Force -PassThru -ErrorAction SilentlyContinue
  if(!$result){
    return $Error[0].Exception.Message
  }
} else {
  if ($entityType -eq "File" -Or  !$existingPath) {
    $result = Move-Item -Path $path -Destination $destination -Force -PassThru -ErrorAction SilentlyContinue
    if(!$result){
      return $Error[0].Exception.Message
    }
  }
  else {
    # Move-Item -Force doesn't work when replacing folders, remove destination folder before replacing
    Remove-Item -Path $existingPath -Confirm:$false -Force -Recurse
    $forceResult = Move-Item -Path $path -Destination $destination -Force -PassThru -ErrorAction SilentlyContinue
    if (!$forceResult) {
      return $Error[0].Exception.Message
    }
  }
}

}
## [END] Move-WACCOFile ##
function New-WACCOFile {
<#

.SYNOPSIS
Create a new file.

.DESCRIPTION
Create a new file on this server. If the file already exists, it will be overwritten.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Path
    String -- the path to the parent of the new file.

.PARAMETER NewName
    String -- the file name.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path,

    [Parameter(Mandatory = $true)]
    [String]
    $NewName
)

Set-StrictMode -Version 5.0

$newItem = New-Item -ItemType File -Path (Join-Path -Path $Path -ChildPath $NewName) -Force

return $newItem |
    Microsoft.PowerShell.Utility\Select-Object @{Name = "Caption"; Expression = { $_.FullName } },
                  @{Name = "CreationDate"; Expression = { $_.CreationTimeUtc } },
                  Extension,
                  @{Name = "IsHidden"; Expression = { $_.Attributes -match "Hidden" } },
                  @{Name = "IsShared"; Expression = { [bool]($folderShares | Where-Object Path -eq $_.FullName) } },
                  Name,
                  @{Name = "Type"; Expression = { Get-FileSystemEntityType -Attributes $_.Attributes } },
                  @{Name = "LastModifiedDate"; Expression = { $_.LastWriteTimeUtc } },
                  @{Name = "Size"; Expression = { if ($_.PSIsContainer) { $null } else { $_.Length } } };

}
## [END] New-WACCOFile ##
function New-WACCOFolder {
<#

.SYNOPSIS
Create a new folder.

.DESCRIPTION
Create a new folder on this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Path
    String -- the path to the parent of the new folder.

.PARAMETER NewName
    String -- the folder name.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path,

    [Parameter(Mandatory = $true)]
    [String]
    $NewName
)

Set-StrictMode -Version 5.0

$pathSeparator = [System.IO.Path]::DirectorySeparatorChar;
$newItem = New-Item -ItemType Directory -Path ($Path.TrimEnd($pathSeparator) + $pathSeparator + $NewName)

return $newItem |
    Microsoft.PowerShell.Utility\Select-Object @{Name="Caption"; Expression={$_.FullName}},
                  @{Name="CreationDate"; Expression={$_.CreationTimeUtc}},
                  Extension,
                  @{Name="IsHidden"; Expression={$_.Attributes -match "Hidden"}},
                  Name,
                  @{Name="Type"; Expression={Get-FileSystemEntityType -Attributes $_.Attributes}},
                  @{Name="LastModifiedDate"; Expression={$_.LastWriteTimeUtc}},
                  @{Name="Size"; Expression={if ($_.PSIsContainer) { $null } else { $_.Length }}};

}
## [END] New-WACCOFolder ##
function New-WACCOSmbFileShare {
<#

.SYNOPSIS
Gets the SMB file share  details on the server.

.DESCRIPTION
Gets the SMB file share  details on the server.

.ROLE
Administrators

#>

<#
.Synopsis
    Name: New-SmbFileShare
    Description: Gets the SMB file share  details on the server.
.Returns
    Retrieves all the SMB file share  details on the server.
#>


param (
    [Parameter(Mandatory = $true)]
    [String]
    $path,    

    [Parameter(Mandatory = $true)]
    [String]
    $name,

    [Parameter(Mandatory = $false)]
    [String[]]
    $fullAccess,

    [Parameter(Mandatory = $false)]
    [String[]]
    $changeAccess,

    [Parameter(Mandatory = $false)]
    [String[]]
    $readAccess,

    [Parameter(Mandatory = $false)]
    [String[]]
    $noAccess,

    [Parameter(Mandatory = $false)]
    [Int]
    $cachingMode,

    [Parameter(Mandatory = $false)]
    [boolean]
    $encryptData,

    # TODO:
    # [Parameter(Mandatory = $false)]
    # [Int]
    # $FolderEnumerationMode

    [Parameter(Mandatory = $false)]
    [boolean]
    $compressData,

    [Parameter(Mandatory = $false)]
    [boolean]
    $isCompressDataEnabled
)

$HashArguments = @{
  Name = "$name"
}

if($fullAccess.count -gt 0){
    $HashArguments.Add("FullAccess", $fullAccess)
}
if($changeAccess.count -gt 0){
    $HashArguments.Add("ChangeAccess", $changeAccess)
}
if($readAccess.count -gt 0){
    $HashArguments.Add("ReadAccess", $readAccess)
}
if($noAccess.count -gt 0){
    $HashArguments.Add("NoAccess", $noAccess)
}
if($cachingMode){
     $HashArguments.Add("CachingMode", "$cachingMode")
}
if($encryptData -ne $null){
    $HashArguments.Add("EncryptData", $encryptData)
}
# TODO: if($FolderEnumerationMode -eq 0){
#     $HashArguments.Add("FolderEnumerationMode", "AccessBased")
# } else {
#     $HashArguments.Add("FolderEnumerationMode", "Unrestricted")
# }
if($isCompressDataEnabled){
    $HashArguments.Add("CompressData", $compressData)
}

New-SmbShare -Path "$path" @HashArguments
@{ shareName = $name } 

}
## [END] New-WACCOSmbFileShare ##
function Remove-WACCOAllShareNames {
<#

.SYNOPSIS
Removes all shares of a folder.

.DESCRIPTION
Removes all shares of a folder.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Path
    String -- The path to the folder.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path    
)

Set-StrictMode -Version 5.0

$CimInstance = Get-CimInstance -Class Win32_Share -Filter Path="'$Path'"
$RemoveShareCommand = ''
if ($CimInstance.name -And $CimInstance.name.GetType().name -ne 'String') { $RemoveShareCommand = $CimInstance.ForEach{ 'Remove-SmbShare -Name "' + $_.name + '" -Force'} } 
Else { $RemoveShareCommand = 'Remove-SmbShare -Name "' + $CimInstance.Name + '" -Force'}
if($RemoveShareCommand) { $RemoveShareCommand.ForEach{ Invoke-Expression $_ } }


}
## [END] Remove-WACCOAllShareNames ##
function Remove-WACCOFileSystemEntity {
<#

.SYNOPSIS
Remove the passed in file or path.

.DESCRIPTION
Remove the passed in file or path from this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER path
    String -- the file or path to remove.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path
)

Set-StrictMode -Version 5.0

Remove-Item -Path $Path -Confirm:$false -Force -Recurse

}
## [END] Remove-WACCOFileSystemEntity ##
function Remove-WACCOFolderShareUser {
<#

.SYNOPSIS
Removes a user from the folder access.

.DESCRIPTION
Removes a user from the folder access.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Path
    String -- The path to the folder.

.PARAMETER Identity
    String -- The user identification (AD / Local user).

.PARAMETER FileSystemRights
    String -- File system rights of the user.

.PARAMETER AccessControlType
    String -- Access control type of the user.    

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path,

    [Parameter(Mandatory = $true)]
    [String]
    $Identity,

    [Parameter(Mandatory = $true)]
    [String]
    $FileSystemRights,

    [ValidateSet('Deny','Allow')]
    [Parameter(Mandatory = $true)]
    [String]
    $AccessControlType
)

Set-StrictMode -Version 5.0

function Remove-UserPermission
{
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $Path,
    
        [Parameter(Mandatory = $true)]
        [String]
        $Identity,
        
        [ValidateSet('Deny','Allow')]
        [Parameter(Mandatory = $true)]
        [String]
        $ACT
    )

    $Acl = Get-Acl $Path
    $AccessRule = New-Object system.security.accesscontrol.filesystemaccessrule($Identity, 'ReadAndExecute','ContainerInherit, ObjectInherit', 'None', $ACT)
    $Acl.RemoveAccessRuleAll($AccessRule)
    Set-Acl $Path $Acl
}

Remove-UserPermission $Path $Identity 'Allow'
Remove-UserPermission $Path $Identity 'Deny'
}
## [END] Remove-WACCOFolderShareUser ##
function Remove-WACCOSmbServerCertificateMapping {
<#
.SYNOPSIS
Removes the currently installed certificate for smb over QUIC on the server.

.DESCRIPTION
Removes the currently installed certificate for smb over QUIC on the server and also sets the status of smbOverQuic to enabled on the server.

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $false)]
    [String[]]
    $ServerDNSNames,

    [Parameter(Mandatory = $false)]
    [String]
    $KdcPort,

    [Parameter(Mandatory = $true)]
    [boolean]
    $IsKdcProxyMappedForSmbOverQuic
)

Set-Location Cert:\LocalMachine\My


if($ServerDNSNames.count -gt 0){
    foreach($serverDNSName in $ServerDNSNames){
        Remove-SmbServerCertificateMapping -Name $serverDNSName -Force
    }
}

if($IsKdcProxyMappedForSmbOverQuic -and $KdcPort -ne $null){
    $port = "0.0.0.0:"
    $ipport = $port+$KdcPort
    $deleteCertKdc = netsh http delete sslcert ipport=$ipport
    if ($LASTEXITCODE -ne 0) {
        throw $deleteCertKdc
    }
    $firewallString = "KDC Proxy Server service (KPS) for SMB over QUIC"
    Remove-NetFirewallRule -DisplayName $firewallString
}

Set-SmbServerConfiguration -EnableSMBQUIC $true -Force
}
## [END] Remove-WACCOSmbServerCertificateMapping ##
function Remove-WACCOSmbShare {
<#

.SYNOPSIS
Removes shares of a folder.

.DESCRIPTION
Removes selected shares of a folder.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER name
    String -- The name of the folder.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Name    
)

Remove-SmbShare -Name $Name -Force


}
## [END] Remove-WACCOSmbShare ##
function Rename-WACCOFileSystemEntity {
<#

.SYNOPSIS
Rename a folder.

.DESCRIPTION
Rename a folder on this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Path
    String -- the path to the folder.

.PARAMETER NewName
    String -- the new folder name.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path,

    [Parameter(Mandatory = $true)]
    [String]
    $NewName
)

Set-StrictMode -Version 5.0

<#
.Synopsis
    Name: Get-FileSystemEntityType
    Description: Gets the type of a local file system entity.

.Parameters
    $Attributes: The System.IO.FileAttributes of the FileSystemEntity.

.Returns
    The type of the local file system entity.
#>
function Get-FileSystemEntityType
{
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.FileAttributes]
        $Attributes
    )

    if ($Attributes -match "Directory")
    {
        return "Folder";
    }
    else
    {
        return "File";
    }
}

Rename-Item -Path $Path -NewName $NewName -PassThru |
    Microsoft.PowerShell.Utility\Select-Object @{Name="Caption"; Expression={$_.FullName}},
                @{Name="CreationDate"; Expression={$_.CreationTimeUtc}},
                Extension,
                @{Name="IsHidden"; Expression={$_.Attributes -match "Hidden"}},
                Name,
                @{Name="Type"; Expression={Get-FileSystemEntityType -Attributes $_.Attributes}},
                @{Name="LastModifiedDate"; Expression={$_.LastWriteTimeUtc}},
                @{Name="Size"; Expression={if ($_.PSIsContainer) { $null } else { $_.Length }}};

}
## [END] Rename-WACCOFileSystemEntity ##
function Restore-WACCOConfigureSmbServerCertificateMapping {
<#
.SYNOPSIS
Rolls back to the previous state of certificate maping if configure/edit action failed.

.DESCRIPTION
Rolls back to the previous state of certificate maping if configure/edit action failed.

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [boolean]
    $IsKdcProxyEnabled,

    [Parameter(Mandatory = $false)]
    [string]
    $KdcPort,

    [Parameter(Mandatory = $true)]
    [string]
    $KdcProxyOptionSelected
)

[array]$smbCertificateMappings = Get-SmbServerCertificateMapping | Microsoft.PowerShell.Utility\Select-Object Name
[array]$mappingNames = $smbCertificateMappings.name

if ($mappingNames.count -eq 0) {
    return;
}
if ($mappingNames.count -gt 0) {
    foreach ($mappingName in $mappingNames) {
        Remove-SmbServerCertificateMapping -Name $mappingName -Force
    }
}

if (!$IsKdcProxyEnabled -and $KdcProxyOptionSelected -eq 'enabled' -and $KdcPort -ne $null) {
    $urlLeft = "https://+:"
    $urlRight = "/KdcProxy/"
    $url = $urlLeft + $KdcPort + $urlRight
    $output = @(netsh http show urlacl | findstr /i "KdcProxy")
    if ($LASTEXITCODE -ne 0) {
        throw $output
    }

    if ($null -ne $output -and $output.count -gt 0) {
        $deleteOutput = netsh http delete urlacl url=$url
        if ($LASTEXITCODE -ne 0) {
            throw $deleteOutput
        }
    }
    [array]$clientAuthproperty = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "HttpsClientAuth" -ErrorAction SilentlyContinue
    if ($null -ne $clientAuthproperty -and $clientAuthproperty.count -gt 0) {
        Remove-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "HttpsClientAuth"
    }
    [array]$passwordAuthproperty = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "DisallowUnprotectedPasswordAuth" -ErrorAction SilentlyContinue
    if ($null -ne $passwordAuthproperty -and $passwordAuthproperty.count -gt 0) {
        Remove-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "DisallowUnprotectedPasswordAuth"
    }
    [array]$path = Get-Item  -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -ErrorAction SilentlyContinue
    if ($null -ne $path -and $path.count -gt 0) {
        Remove-Item  -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings"
    }
    Stop-Service -Name kpssvc
    $firewallString = "KDC Proxy Server service (KPS) for SMB over QUIC"
    $rule = @(Get-NetFirewallRule -DisplayName $firewallString -ErrorAction SilentlyContinue)
    if($null -ne $rule -and $rule.count -gt 0) {
        Remove-NetFirewallRule -DisplayName $firewallString
    }

    $port = "0.0.0.0:"
    $ipport = $port+$KdcPort
    $certBinding =  @(netsh http show sslcert ipport=$ipport | findstr "Hash")
    if ($LASTEXITCODE -ne 0) {
        throw $certBinding
    }
    if($null -ne $certBinding -and $certBinding.count -gt 0) {
        $deleteCertKdc = netsh http delete sslcert ipport=$ipport
        if ($LASTEXITCODE -ne 0) {
            throw $deleteCertKdc
        } 
    }
}


}
## [END] Restore-WACCOConfigureSmbServerCertificateMapping ##
function Set-WACCOEventLogChannelStatus {
 <#

.SYNOPSIS
 Change the current status (Enabled/Disabled) for the selected channel.

.DESCRIPTION
Change the current status (Enabled/Disabled) for the selected channel.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>

Param(
    [string]$channel,
    [boolean]$status
)

$ch = Get-WinEvent -ListLog $channel
$ch.set_IsEnabled($status)
$ch.SaveChanges()
}
## [END] Set-WACCOEventLogChannelStatus ##
function Set-WACCOSmbOverQuicServerSettings {
<#

.SYNOPSIS
Sets smb server settings for QUIC.

.DESCRIPTION
Sets smb server settings for QUIC.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER disableSmbEncryptionOnSecureConnection
To enable or disable smbEncryption on the server.

.PARAMETER restrictNamedPipeAccessViaQuic
To enable or diable namedPipeAccess on the server.

#>

param (
[Parameter(Mandatory = $true)]
[boolean]
$disableSmbEncryptionOnSecureConnection,

[Parameter(Mandatory = $true)]
[boolean]
$restrictNamedPipeAccessViaQuic
)


Set-SmbServerConfiguration -DisableSmbEncryptionOnSecureConnection $disableSmbEncryptionOnSecureConnection -RestrictNamedPipeAccessViaQuic $restrictNamedPipeAccessViaQuic -Force;

}
## [END] Set-WACCOSmbOverQuicServerSettings ##
function Set-WACCOSmbServerCertificateMapping {
<#
.SYNOPSIS
Set Smb Server Certificate Mapping

.DESCRIPTION
Configures smb over QUIC.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Thumbprint
    String -- The thumbprint of the certifiacte selected.

.PARAMETER ServerDNSNames
    String[] -- The addresses of the server for certificate mapping.
#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Thumbprint,

    [Parameter(Mandatory = $true)]
    [String[]]
    $ServerDNSNames,

    [Parameter(Mandatory = $true)]
    [boolean]
    $IsKdcProxyEnabled,

    [Parameter(Mandatory = $true)]
    [String]
    $KdcProxyOptionSelected,

    [Parameter(Mandatory = $false)]
    [String]
    $KdcPort
)

Import-Module -Name Microsoft.PowerShell.Management -ErrorAction SilentlyContinue
Set-Variable -Name LogName -Option Constant -Value "Microsoft-ServerManagementExperience" -ErrorAction SilentlyContinue
Set-Variable -Name LogSource -Option Constant -Value "SmeScripts-ConfigureKdcProxy" -ErrorAction SilentlyContinue

function writeInfoLog($logMessage) {
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Information `
        -Message $logMessage -ErrorAction SilentlyContinue
}

Set-Location Cert:\LocalMachine\My

function Enable-KdcProxy {

    $urlLeft = "https://+:"
    $urlRight = "/KdcProxy"
    $url = $urlLeft + $KdcPort + $urlRight

    $port = "0.0.0.0:"
    $ipport = $port+$KdcPort

    $ComputerName = (Get-CimInstance Win32_ComputerSystem).DNSHostName + "." + (Get-CimInstance Win32_ComputerSystem).Domain

    try {
        $certBinding = @(netsh http show sslcert ipport=$ipport | findstr "Hash")
        if($null -ne $certBinding -and $certBinding.count -gt 0) {
            $deleteCertKdc = netsh http delete sslcert ipport=$ipport
            if ($LASTEXITCODE -ne 0) {
                throw $deleteCertKdc
            }
        }
        $guid = [Guid]::NewGuid()
        Add-NetIPHttpsCertBinding -ipport $ipport -CertificateHash $Thumbprint -CertificateStoreName "my" -ApplicationId "{$guid}" -NullEncryption $false
        $message = 'Completed adding ssl certificate port'
        writeInfoLog $message
        if ($ServerDNSNames.count -gt 0) {
            foreach ($serverDnsName in $ServerDnsNames) {
                if($ComputerName.trim() -ne $serverDnsName.trim()){
                    $output = Echo 'Y' | netdom computername $ComputerName /add $serverDnsName
                    if ($LASTEXITCODE -ne 0) {
                        throw $output
                    }
                }
            }
            $message = 'Completed adding alternative names'
            writeInfoLog $message
        }
        if(!$IsKdcProxyEnabled) {
            $netshOutput = netsh http add urlacl url=$url user="NT authority\Network Service"
            if ($LASTEXITCODE -ne 0) {
                throw $netshOutput
            }
            $message = 'Completed adding urlacl'
            writeInfoLog $message
            New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings"  -force
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "HttpsClientAuth" -Value 0x0 -type DWORD
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "DisallowUnprotectedPasswordAuth" -Value 0x0 -type DWORD
            Set-Service -Name kpssvc -StartupType Automatic
            Start-Service -Name kpssvc
        }
        $message = 'Returning method Enable-KdcProxy'
        writeInfoLog $message
        return $true;
    }
    catch {
        throw $_
    }
}
function New-SmbCertificateMapping {
    if ($ServerDNSNames.count -gt 0) {
        $message = 'Setting server dns names'
        writeInfoLog $message
        foreach ($serverDNSName in $ServerDNSNames) {
            New-SmbServerCertificateMapping -Name $serverDNSName -Thumbprint $Thumbprint -StoreName My -Force
        }
        if ($KdcProxyOptionSelected -eq "enabled" -and $KdcPort -ne $null){
            $message = 'Enabling Kdc Proxy'
            writeInfoLog $message
            $result = Enable-KdcProxy
            if($result) {
                $firewallString = "KDC Proxy Server service (KPS) for SMB over QUIC"
                $firewallDesc = "The KDC Proxy Server service runs on edge servers to proxy Kerberos protocol messages to domain controllers on the corporate network. Default port is TCP/443."
                New-NetFirewallRule -DisplayName $firewallString -Description $firewallDesc -Protocol TCP -LocalPort $KdcPort -Direction Inbound -Action Allow
            }
        }
        return $true;
    }
    $message = 'Exiting method smb certificate mapping '
    writeInfoLog $message
    return $true;
}

return New-SmbCertificateMapping;

}
## [END] Set-WACCOSmbServerCertificateMapping ##
function Set-WACCOSmbServerSettings {
<#
.SYNOPSIS
Updates the server configuration settings on the server.

.DESCRIPTION
Updates the server configuration settings on the server.

.ROLE
Administrators

#>

<#
.Synopsis
    Name: Set-SmbServerSettings
    Description: Updates the server configuration settings on the server.
#>


param (

    [Parameter(Mandatory = $true)]
    [boolean]
    $AuditSmb1Access,

    [Parameter(Mandatory = $true)]
    [boolean]
    $RequireSecuritySignature,

    [Parameter(Mandatory = $false)]
    [boolean]
    $RejectUnencryptedAccess,

    [Parameter(Mandatory = $true)]
    [boolean]
    $EncryptData,

    [Parameter(Mandatory = $true)]
    [boolean]
    $CompressionSettingsClicked
)

$HashArguments = @{
}

if($CompressionSettingsClicked) {
    Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\LanManServer\parameters" -Name "DisableCompression" -Value 0x1 -type DWORD
} else {
    Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\LanManServer\parameters" -Name "DisableCompression" -Value 0x0 -type DWORD
}

if($RejectUnencryptedAccess -ne $null){
    $HashArguments.Add("RejectUnencryptedAccess", $RejectUnencryptedAccess)
}

Set-SmbServerConfiguration -AuditSmb1Access $AuditSmb1Access -RequireSecuritySignature $RequireSecuritySignature  -EncryptData $EncryptData @HashArguments -Force
}
## [END] Set-WACCOSmbServerSettings ##
function Test-WACCOFileSystemEntity {
<#

.SYNOPSIS
Checks if a file or folder exists

.DESCRIPTION
Checks if a file or folder exists
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Path
    String -- The path to check if it exists

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path    
)

Set-StrictMode -Version 5.0

Test-Path -path $Path

}
## [END] Test-WACCOFileSystemEntity ##
function Uninstall-WACCOSmb1 {
<#
.SYNOPSIS
Disables SMB1 on the server.

.DESCRIPTION
Disables SMB1 on the server.

.ROLE
Administrators

#>

<#
.Synopsis
    Name: UninstallSmb1
    Description: Disables SMB1 on the server.
#>

Disable-WindowsOptionalFeature -Online -FeatureName smb1protocol -NoRestart
}
## [END] Uninstall-WACCOSmb1 ##

# SIG # Begin signature block
# MIIn0QYJKoZIhvcNAQcCoIInwjCCJ74CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBbdI4N8v5NPKQa
# dL76M8Slkt88Jpj0VCeZUyv/viHRL6CCDYUwggYDMIID66ADAgECAhMzAAACzfNk
# v/jUTF1RAAAAAALNMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjIwNTEyMjA0NjAyWhcNMjMwNTExMjA0NjAyWjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDrIzsY62MmKrzergm7Ucnu+DuSHdgzRZVCIGi9CalFrhwtiK+3FIDzlOYbs/zz
# HwuLC3hir55wVgHoaC4liQwQ60wVyR17EZPa4BQ28C5ARlxqftdp3H8RrXWbVyvQ
# aUnBQVZM73XDyGV1oUPZGHGWtgdqtBUd60VjnFPICSf8pnFiit6hvSxH5IVWI0iO
# nfqdXYoPWUtVUMmVqW1yBX0NtbQlSHIU6hlPvo9/uqKvkjFUFA2LbC9AWQbJmH+1
# uM0l4nDSKfCqccvdI5l3zjEk9yUSUmh1IQhDFn+5SL2JmnCF0jZEZ4f5HE7ykDP+
# oiA3Q+fhKCseg+0aEHi+DRPZAgMBAAGjggGCMIIBfjAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQU0WymH4CP7s1+yQktEwbcLQuR9Zww
# VAYDVR0RBE0wS6RJMEcxLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJh
# dGlvbnMgTGltaXRlZDEWMBQGA1UEBRMNMjMwMDEyKzQ3MDUzMDAfBgNVHSMEGDAW
# gBRIbmTlUAXTgqoXNzcitW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8v
# d3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIw
# MTEtMDctMDguY3JsMGEGCCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDEx
# XzIwMTEtMDctMDguY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIB
# AE7LSuuNObCBWYuttxJAgilXJ92GpyV/fTiyXHZ/9LbzXs/MfKnPwRydlmA2ak0r
# GWLDFh89zAWHFI8t9JLwpd/VRoVE3+WyzTIskdbBnHbf1yjo/+0tpHlnroFJdcDS
# MIsH+T7z3ClY+6WnjSTetpg1Y/pLOLXZpZjYeXQiFwo9G5lzUcSd8YVQNPQAGICl
# 2JRSaCNlzAdIFCF5PNKoXbJtEqDcPZ8oDrM9KdO7TqUE5VqeBe6DggY1sZYnQD+/
# LWlz5D0wCriNgGQ/TWWexMwwnEqlIwfkIcNFxo0QND/6Ya9DTAUykk2SKGSPt0kL
# tHxNEn2GJvcNtfohVY/b0tuyF05eXE3cdtYZbeGoU1xQixPZAlTdtLmeFNly82uB
# VbybAZ4Ut18F//UrugVQ9UUdK1uYmc+2SdRQQCccKwXGOuYgZ1ULW2u5PyfWxzo4
# BR++53OB/tZXQpz4OkgBZeqs9YaYLFfKRlQHVtmQghFHzB5v/WFonxDVlvPxy2go
# a0u9Z+ZlIpvooZRvm6OtXxdAjMBcWBAsnBRr/Oj5s356EDdf2l/sLwLFYE61t+ME
# iNYdy0pXL6gN3DxTVf2qjJxXFkFfjjTisndudHsguEMk8mEtnvwo9fOSKT6oRHhM
# 9sZ4HTg/TTMjUljmN3mBYWAWI5ExdC1inuog0xrKmOWVMIIHejCCBWKgAwIBAgIK
# YQ6Q0gAAAAAAAzANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlm
# aWNhdGUgQXV0aG9yaXR5IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEw
# OTA5WjB+MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYD
# VQQDEx9NaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+la
# UKq4BjgaBEm6f8MMHt03a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc
# 6Whe0t+bU7IKLMOv2akrrnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4D
# dato88tt8zpcoRb0RrrgOGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+
# lD3v++MrWhAfTVYoonpy4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nk
# kDstrjNYxbc+/jLTswM9sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6
# A4aN91/w0FK/jJSHvMAhdCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmd
# X4jiJV3TIUs+UsS1Vz8kA/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL
# 5zmhD+kjSbwYuER8ReTBw3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zd
# sGbiwZeBe+3W7UvnSSmnEyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3
# T8HhhUSJxAlMxdSlQy90lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS
# 4NaIjAsCAwEAAaOCAe0wggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRI
# bmTlUAXTgqoXNzcitW2oynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTAL
# BgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBD
# uRQFTuHqp8cx0SOJNDBaBgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jv
# c29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFf
# MDNfMjIuY3JsMF4GCCsGAQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFf
# MDNfMjIuY3J0MIGfBgNVHSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEF
# BQcCARYzaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1h
# cnljcHMuaHRtMEAGCCsGAQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkA
# YwB5AF8AcwB0AGEAdABlAG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn
# 8oalmOBUeRou09h0ZyKbC5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7
# v0epo/Np22O/IjWll11lhJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0b
# pdS1HXeUOeLpZMlEPXh6I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/
# KmtYSWMfCWluWpiW5IP0wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvy
# CInWH8MyGOLwxS3OW560STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBp
# mLJZiWhub6e3dMNABQamASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJi
# hsMdYzaXht/a8/jyFqGaJ+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYb
# BL7fQccOKO7eZS/sl/ahXJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbS
# oqKfenoi+kiVH6v7RyOA9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sL
# gOppO6/8MO0ETI7f33VtY5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtX
# cVZOSEXAQsmbdlsKgEhr/Xmfwb1tbWrJUnMTDXpQzTGCGaIwghmeAgEBMIGVMH4x
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01p
# Y3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTECEzMAAALN82S/+NRMXVEAAAAA
# As0wDQYJYIZIAWUDBAIBBQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIE4F
# 8dLQ6Kbp32iExxOE/YEOkIBL6ba3P6uFdpzzzqI1MEIGCisGAQQBgjcCAQwxNDAy
# oBSAEgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20wDQYJKoZIhvcNAQEBBQAEggEA2DXivbYxYtDdYkAKzC36yJINV6AAc52WLkl+
# CmqV035I9ZKTq0KC0eKzp3ZGuwv7cMA2GGmUPgut58WAcbTF7zuE/1nOJiWixV3y
# jYMDLcsmmqSx398Ga2pMkg/NRSSOVU/Ci46DqXlQkrBPAzeL+8F4pUaAl0Obp7Qv
# jnHDkjxcJyNQ1yuRiWjx1WgFK/pLeTqcAGDET82GOsTstMgDcpwt4A+p+97jwx1x
# CrQ00YagyqFDn0kub0rSkAMwljspVkKoPHGXajVCzeWcmGsjhCkfe1ifOTyxW5z5
# 4gKF55O1m+TESAxnz2T0hOvT4rjjxMtFQfgH7JeN11RVnMA7Z6GCFywwghcoBgor
# BgEEAYI3AwMBMYIXGDCCFxQGCSqGSIb3DQEHAqCCFwUwghcBAgEDMQ8wDQYJYIZI
# AWUDBAIBBQAwggFZBgsqhkiG9w0BCRABBKCCAUgEggFEMIIBQAIBAQYKKwYBBAGE
# WQoDATAxMA0GCWCGSAFlAwQCAQUABCDSNDJbKgbWIlhZnEZLY/kwpaqEUnHnAya/
# t0NzXiXRzwIGZD/Sir5CGBMyMDIzMDQyMTAwMjIxMC40NTFaMASAAgH0oIHYpIHV
# MIHSMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQL
# EyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsT
# HVRoYWxlcyBUU1MgRVNOOjNCRDQtNEI4MC02OUMzMSUwIwYDVQQDExxNaWNyb3Nv
# ZnQgVGltZS1TdGFtcCBTZXJ2aWNloIIRezCCBycwggUPoAMCAQICEzMAAAG0+4AI
# RAXSLfoAAQAAAbQwDQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAg
# UENBIDIwMTAwHhcNMjIwOTIwMjAyMjA5WhcNMjMxMjE0MjAyMjA5WjCB0jELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEtMCsGA1UECxMkTWljcm9z
# b2Z0IElyZWxhbmQgT3BlcmF0aW9ucyBMaW1pdGVkMSYwJAYDVQQLEx1UaGFsZXMg
# VFNTIEVTTjozQkQ0LTRCODAtNjlDMzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgU2VydmljZTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBALRH
# pp5lBzJCH7zortuyvOmW8FoZLBsFe9g5dbhnaq9qSpvpn86E/mJ4JKvWixH/lw7Q
# A8gPtiiGVNIjvFhu/XiY889vX5WaQSmyoPMZdj9zvXa5XrkMN05zXzTePkCIIzF6
# RN7cTxezOyESymTIjrdxX5BVlZolyQAOxNziMCYKYYNPbYd0786fDE/PhzrRt23a
# 0Xf8trvFa0LEEy2YlcE2eqg2CjU/D0GZe8Ra0kjt0M12vdS4qWZ2Dpd7IhiQwnnt
# QWu19Ytd3UBR8SpeRX+Ccw3bjgWfOXtla6chctWt2shlMwayMOfY4TG4yMPWFXEL
# fZFFp7cgpjZNeVsmwkvoV6RAwy1Y9V+VvbJ5qFtartN/rp6a0I1kGlbjuwX3L0HT
# VXcikqgHistXk9h3HOZ9WgFXlxZurG1SZmcz0BEEdya+1vGHE45KguYU9qq2LiHG
# Bjn9z4+DqnV5tUKobsLbJMb4r+8st2fj8SacSsftnusxkWqEJiJS34P2uNlzVR03
# +ls6+ZO0NcO79LgP7BbIMipiOx8yh19PMQw0piaKFwOW7Q+gdJcfy6rOkG+CrYZw
# OzdiBHSebIzCIch2cAa+38w7JFP/koKdlJ36qzdVXWv4G/qZpWycIvDKYbxJWM40
# +z2Stg5uHqK3I8e09kFXtxCHpS7hm8c8m25WaEU5AgMBAAGjggFJMIIBRTAdBgNV
# HQ4EFgQUy0SF5fGUuDqcuxIot07eOMwy2X4wHwYDVR0jBBgwFoAUn6cVXQBeYl2D
# 9OXSZacbUzUZ6XIwXwYDVR0fBFgwVjBUoFKgUIZOaHR0cDovL3d3dy5taWNyb3Nv
# ZnQuY29tL3BraW9wcy9jcmwvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUy
# MDIwMTAoMSkuY3JsMGwGCCsGAQUFBwEBBGAwXjBcBggrBgEFBQcwAoZQaHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNyb3NvZnQlMjBUaW1l
# LVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcnQwDAYDVR0TAQH/BAIwADAWBgNVHSUB
# Af8EDDAKBggrBgEFBQcDCDAOBgNVHQ8BAf8EBAMCB4AwDQYJKoZIhvcNAQELBQAD
# ggIBABLRDwWMKbeCYqEqtI6Bs8KmF+kqDR+2G6qYAK3ZZ63bert7pCkRJbihFakt
# l2o18cdFJFxnOF4vXadm0sabskJ05KviEMJIO6dXSq8AGtr3Zmjc895q0mnlBLuN
# Mgk4R8KrkJMHqBuHqkUWXtfTrVUpgwzQt2UOiINKs+/b4r14MuXRVpOJ6cQOS8Uh
# keMAWl2iLlYaBGtOr3f/f9mLEPfWwoke0sSUbdV60OZCRh1ItBYYM9efKr14H5qu
# 6jan6n00prEEa7W3uGb/1/qj6P5emnvkqy5HI0X69DjVdLxVbjSsegm/dA+S4DaX
# PcfFf6iBxK/iV21l1upgEVVajUApl5VR40wY4XF8EpmnUdTqLXDf7CqdhDjPST2K
# /OjvWPyQGQvc7oPapYyk66GU32AOyyHXJj6+vbtRUg/+ory+h0R2Xf5NhC+xbWcM
# zXEUXRRf1YKZDsRyH6r412pm8KDKE/r7Rk7aoKK7oYUpNGzNRf6QaYv5z2bVTSxk
# zWivFrepLHGwvRun9PYM/8AQSTgZr0yzzjk/97WghkqCaAwAVpyvg3uaYnuCl/Ac
# cSkGyb8c+70bFSeUephsfgb2r+QI7Mb2WcOnkJpCNLz0XJMS/UwlQn1ktLsiCpsq
# Ok3aLJ2wTv6LK3u69I0vQB/LKRKlZYRXKUDXzoPwr3UtsTVTMIIHcTCCBVmgAwIB
# AgITMwAAABXF52ueAptJmQAAAAAAFTANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UE
# BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAc
# BgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0
# IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IDIwMTAwHhcNMjEwOTMwMTgyMjI1
# WhcNMzAwOTMwMTgzMjI1WjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGlu
# Z3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
# cmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDCC
# AiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAOThpkzntHIhC3miy9ckeb0O
# 1YLT/e6cBwfSqWxOdcjKNVf2AX9sSuDivbk+F2Az/1xPx2b3lVNxWuJ+Slr+uDZn
# hUYjDLWNE893MsAQGOhgfWpSg0S3po5GawcU88V29YZQ3MFEyHFcUTE3oAo4bo3t
# 1w/YJlN8OWECesSq/XJprx2rrPY2vjUmZNqYO7oaezOtgFt+jBAcnVL+tuhiJdxq
# D89d9P6OU8/W7IVWTe/dvI2k45GPsjksUZzpcGkNyjYtcI4xyDUoveO0hyTD4MmP
# frVUj9z6BVWYbWg7mka97aSueik3rMvrg0XnRm7KMtXAhjBcTyziYrLNueKNiOSW
# rAFKu75xqRdbZ2De+JKRHh09/SDPc31BmkZ1zcRfNN0Sidb9pSB9fvzZnkXftnIv
# 231fgLrbqn427DZM9ituqBJR6L8FA6PRc6ZNN3SUHDSCD/AQ8rdHGO2n6Jl8P0zb
# r17C89XYcz1DTsEzOUyOArxCaC4Q6oRRRuLRvWoYWmEBc8pnol7XKHYC4jMYcten
# IPDC+hIK12NvDMk2ZItboKaDIV1fMHSRlJTYuVD5C4lh8zYGNRiER9vcG9H9stQc
# xWv2XFJRXRLbJbqvUAV6bMURHXLvjflSxIUXk8A8FdsaN8cIFRg/eKtFtvUeh17a
# j54WcmnGrnu3tz5q4i6tAgMBAAGjggHdMIIB2TASBgkrBgEEAYI3FQEEBQIDAQAB
# MCMGCSsGAQQBgjcVAgQWBBQqp1L+ZMSavoKRPEY1Kc8Q/y8E7jAdBgNVHQ4EFgQU
# n6cVXQBeYl2D9OXSZacbUzUZ6XIwXAYDVR0gBFUwUzBRBgwrBgEEAYI3TIN9AQEw
# QTA/BggrBgEFBQcCARYzaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9E
# b2NzL1JlcG9zaXRvcnkuaHRtMBMGA1UdJQQMMAoGCCsGAQUFBwMIMBkGCSsGAQQB
# gjcUAgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/
# MB8GA1UdIwQYMBaAFNX2VsuP6KJcYmjRPZSQW9fOmhjEMFYGA1UdHwRPME0wS6BJ
# oEeGRWh0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01p
# Y1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYB
# BQUHMAKGPmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljUm9v
# Q2VyQXV0XzIwMTAtMDYtMjMuY3J0MA0GCSqGSIb3DQEBCwUAA4ICAQCdVX38Kq3h
# LB9nATEkW+Geckv8qW/qXBS2Pk5HZHixBpOXPTEztTnXwnE2P9pkbHzQdTltuw8x
# 5MKP+2zRoZQYIu7pZmc6U03dmLq2HnjYNi6cqYJWAAOwBb6J6Gngugnue99qb74p
# y27YP0h1AdkY3m2CDPVtI1TkeFN1JFe53Z/zjj3G82jfZfakVqr3lbYoVSfQJL1A
# oL8ZthISEV09J+BAljis9/kpicO8F7BUhUKz/AyeixmJ5/ALaoHCgRlCGVJ1ijbC
# HcNhcy4sa3tuPywJeBTpkbKpW99Jo3QMvOyRgNI95ko+ZjtPu4b6MhrZlvSP9pEB
# 9s7GdP32THJvEKt1MMU0sHrYUP4KWN1APMdUbZ1jdEgssU5HLcEUBHG/ZPkkvnNt
# yo4JvbMBV0lUZNlz138eW0QBjloZkWsNn6Qo3GcZKCS6OEuabvshVGtqRRFHqfG3
# rsjoiV5PndLQTHa1V1QJsWkBRH58oWFsc/4Ku+xBZj1p/cvBQUl+fpO+y/g75LcV
# v7TOPqUxUYS8vwLBgqJ7Fx0ViY1w/ue10CgaiQuPNtq6TPmb/wrpNPgkNWcr4A24
# 5oyZ1uEi6vAnQj0llOZ0dFtq0Z4+7X6gMTN9vMvpe784cETRkPHIqzqKOghif9lw
# Y1NNje6CbaUFEMFxBmoQtB1VM1izoXBm8qGCAtcwggJAAgEBMIIBAKGB2KSB1TCB
# 0jELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1Jl
# ZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEtMCsGA1UECxMk
# TWljcm9zb2Z0IElyZWxhbmQgT3BlcmF0aW9ucyBMaW1pdGVkMSYwJAYDVQQLEx1U
# aGFsZXMgVFNTIEVTTjozQkQ0LTRCODAtNjlDMzElMCMGA1UEAxMcTWljcm9zb2Z0
# IFRpbWUtU3RhbXAgU2VydmljZaIjCgEBMAcGBSsOAwIaAxUAZZzYkPObl/ZzeCkS
# bf4B5CceCQiggYMwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGlu
# Z3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
# cmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAN
# BgkqhkiG9w0BAQUFAAIFAOfsSz4wIhgPMjAyMzA0MjEwNzM3MzRaGA8yMDIzMDQy
# MjA3MzczNFowdzA9BgorBgEEAYRZCgQBMS8wLTAKAgUA5+xLPgIBADAKAgEAAgIp
# 1AIB/zAHAgEAAgISCDAKAgUA5+2cvgIBADA2BgorBgEEAYRZCgQCMSgwJjAMBgor
# BgEEAYRZCgMCoAowCAIBAAIDB6EgoQowCAIBAAIDAYagMA0GCSqGSIb3DQEBBQUA
# A4GBADIevngeMReQeuYLwr04UMfuC6K7jfy2xunvVUyf9cnGE7k4mzOiNcq4ELAe
# 2+wslETxCJNx2+Ebr4kZJBxT7f/V1uqS/6mWYUX6daq26x3bV2o7TZGnjNam4euq
# fSBmhU1NPet5quHAFkh/eIGvhkkUg1F4ESxcpHivKw1tNfoGMYIEDTCCBAkCAQEw
# gZMwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcT
# B1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UE
# AxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAG0+4AIRAXSLfoA
# AQAAAbQwDQYJYIZIAWUDBAIBBQCgggFKMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0B
# CRABBDAvBgkqhkiG9w0BCQQxIgQgkS5+UVGrmzURH5zhZ9pRSyY24X7s/59HTzJq
# RzP+LQYwgfoGCyqGSIb3DQEJEAIvMYHqMIHnMIHkMIG9BCDTyPd75qMwcAZRcb36
# /6xJa3hT0eLse71ysdp4twH3BjCBmDCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQSAyMDEwAhMzAAABtPuACEQF0i36AAEAAAG0MCIEICjZsOmMMdL/dp7IFR9h
# RnfmB7O4AeqEP4YONseBD3JnMA0GCSqGSIb3DQEBCwUABIICAKlMPY0b4BicOlNM
# iwP6qMI8cGfuQt0dgT8n79qtr4aMyXE+gS+FAiMlDNHTUyhFQ7DlBRoldFvnWBko
# wsCWYWpDMinQ593uOnMaDFv5K7ofJvd5XNKOH+GRxaYK+n0zN+4jbrk+LVKijgWh
# VOj3vTPM+M+6E1uepndksvXFr8V46s7r1DKiuGAPLQUt/9SEM3oYivsV3gV16Os8
# eDERChNW/jV7AuHRRw0pa6MJQlCd2pms7CBaugOxj/6SAQ6n6VK1f0veUR0YWAwK
# WLg7QIPCHFlEU4vNfzYcvPgIGjkPpghbSOYJox6NBFYKHtSGWpggYeWzJa7GyhNs
# nTsun3gX6aE0aVGitVIRtdCHKIZLVaewDGI+3p01KOit0Ph8jp7bh0shAKKvw3vn
# DaCXPkZbmtyV8O7kHWUHYzix8VcV+EZdxsSyuHGRw/1s7YX3LfCgm4tO8xH6sJDN
# YThAT+RCTf4xRmyKuvJwEaAhW7YwtgjenmYbsIQ+4HpfgSFnehDQMzfUXI8/tbVB
# YDni77+uqsBQHfAYpln1fWrIo4vOuqNjyL6Plj/hpy94ZoRs9T+0u1lrkSXfNXqL
# 2497F6ZtJj/glqROv+rKNj6GzCBh5+Jqi1hlB+ifFUy1r6bHWRzbgFKhAUZvljhV
# hb6aUpjlktCattSMiXOSa19OhGwg
# SIG # End signature block

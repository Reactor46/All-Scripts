[CmdletBinding()]
param(
    [string]$server,
    [string]$webSiteName,
    [string]$webFolder,
    [string]$anonymousAuthentication,
    [string]$basicAuthentication,
    [string]$windowsAuthentication
)

Import-Module "Microsoft.TeamFoundation.DistributedTask.Task.Internal"
Import-Module "Microsoft.TeamFoundation.DistributedTask.Task.Common"

Write-Verbose "Entering script $($MyInvocation.MyCommand.Name)"
Write-Verbose "Parameter Values"
$PSBoundParameters.Keys | % { Write-HOST "  ${_} = $($PSBoundParameters[$_])" }

$command = {  
    param(
        [string]$webSiteName,
        [string]$path,
        [boolean]$anonymousAuthentication,
        [boolean]$basicAuthentication,
        [boolean]$windowsAuthentication
    )
    Import-Module WebAdministration; 
    if ([string]::IsNullOrWhiteSpace($path)){
        Write-HOST "Setting authentication on $webSiteName"
        Set-WebConfigurationProperty -filter /system.webServer/security/authentication/anonymousAuthentication -name enabled -value $anonymousAuthentication -PSPath IIS:\ -location $webSiteName;  
        Set-WebConfigurationProperty -filter /system.webServer/security/authentication/basicAuthentication -name enabled -value $basicAuthentication -PSPath IIS:\ -location $webSiteName;  
        Set-WebConfigurationProperty -filter /system.webServer/security/authentication/windowsAuthentication -name enabled -value $windowsAuthentication -PSPath IIS:\ -location $webSiteName;  
    }else{
        $location = "$webSiteName/$path"
        Write-HOST "Setting authentication on $location"
        Set-WebConfigurationProperty -filter /system.webServer/security/authentication/anonymousAuthentication -name enabled -value $anonymousAuthentication  -location $location;  
        Set-WebConfigurationProperty -filter /system.webServer/security/authentication/basicAuthentication -name enabled -value $basicAuthentication -location $location;  
        Set-WebConfigurationProperty -filter /system.webServer/security/authentication/windowsAuthentication -name enabled -value $windowsAuthentication -location $location; 
    }
    Write-Verbose "Authentication Set"
}
[bool]$anonymousAuthenticationBool= Convert-String $anonymousAuthentication Boolean
[bool]$basicAuthenticationBool= Convert-String $basicAuthentication Boolean
[bool]$windowsAuthenticationBool= Convert-String $windowsAuthentication Boolean

Invoke-Command -ComputerName $server $command -ArgumentList $webSiteName,$webFolder,$anonymousAuthenticationBool,$basicAuthenticationBool,$windowsAuthenticationBool
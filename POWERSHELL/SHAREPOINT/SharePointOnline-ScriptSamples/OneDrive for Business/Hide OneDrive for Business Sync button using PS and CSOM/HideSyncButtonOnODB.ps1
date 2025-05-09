﻿function Set-SPOTenant{
	param (
		[Parameter(Mandatory=$true,Position=1)]
		[string]$Username,
		[Parameter(Mandatory=$true,Position=2)]
		[string]$Url,
		[Parameter(Mandatory=$true,Position=3)]
		$password,
		[Parameter(Mandatory=$true,Position=4)]
		[Bool]$HideSyncButtonOnODB
	)

	$ctx=New-Object Microsoft.SharePoint.Client.ClientContext($url)
	$ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username, $password)
	$ctx.Load($ctx.Web)
	$ctx.ExecuteQuery()

	$spoTenant= New-Object Microsoft.Online.SharePoint.TenantAdministration.Tenant($Ctx) 
	$ctx.Load($spoTenant)
	$ctx.ExecuteQuery()
	$spoTenant.HideSyncButtonOnODB=$HideSyncButtonOnODB
	$ctx.Load($spoTenant)
	$ctx.ExecuteQuery()
	Write-Output $spoTenant  
}

#Paths to SDK
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\SharePoint Client Components\16.0\Assemblies\Microsoft.Online.SharePoint.Client.Tenant.dll"  

#Enter the data
$AdminPassword=Read-Host -Prompt "Enter password" -AsSecureString
$username="arleta@TENANT.onmicrosoft.com"
$Url="https://TENANT-admin.sharepoint.com"
$HideSyncButtonOnODB=$true   #Whether the sync button should be hidden

Set-SPOTenant -Username $username -Url $Url -password $AdminPassword -HideSyncButtonOnODB $HideSyncButtonOnODB

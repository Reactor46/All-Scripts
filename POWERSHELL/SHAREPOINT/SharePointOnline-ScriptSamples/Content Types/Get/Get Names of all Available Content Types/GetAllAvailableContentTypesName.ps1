﻿

function Get-SPOContentType{
	param (
		[Parameter(Mandatory=$true,Position=1)]
		[string]$Username,
		[Parameter(Mandatory=$true,Position=2)]
		$AdminPassword,
		[Parameter(Mandatory=$true,Position=3)]
		[string]$Url
	)

	  $ctx=New-Object Microsoft.SharePoint.Client.ClientContext($Url)
	  $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username, $AdminPassword)
	  $ctx.ExecuteQuery() 

	  $ctx.Load($ctx.Web)
	  $ctx.Load($ctx.Web.AvailableContentTypes)
	  $ctx.ExecuteQuery()
	  Write-Host 
	#  Write-Host $ctx.Url -BackgroundColor White -ForegroundColor DarkGreen

	foreach( $cc in $ctx.Web.AvailableContentTypes){  
	   Write-Output $cc.Name 
	}   
}


  # Paths to SDK. Please verify location on your computer.
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll" 
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll" 

# Insert the credentials and the name of the admin site
$Username="admin@tenant.onmicrosoft.com"
$AdminPassword=Read-Host -Prompt "Password" -AsSecureString
$AdminUrl="https://tenant.sharepoint.com/sites/teamsitewithlibraries/sub"



Get-SPOContentType -Username $Username -AdminPassword $AdminPassword -Url $AdminUrl

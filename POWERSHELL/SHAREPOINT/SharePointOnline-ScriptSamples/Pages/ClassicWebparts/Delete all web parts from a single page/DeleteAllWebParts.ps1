﻿ 
 function Remove-SPOWebParts{
	param (
        	[Parameter(Mandatory=$true,Position=1)]
		[string]$Username,
		[Parameter(Mandatory=$true,Position=2)]
		[string]$Url,
        	[Parameter(Mandatory=$true,Position=3)]
		$password,
		[Parameter(Mandatory=$true,Position=4)]
		[string]$pageUrl
	)


	$ctx=New-Object Microsoft.SharePoint.Client.ClientContext($Url)
	$ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username, $password)
	$ctx.Load($ctx.Web)
	$ctx.ExecuteQuery()

  	$page = $ctx.Web.GetFileByServerRelativeUrl($pageUrl)
  	$wpm = $page.GetLimitedWebPartManager("Shared")
        $ctx.Load($wpm);
        $ctx.Load($wpm.WebParts);
        $ctx.ExecuteQuery()

        foreach($webbie in $wpm.WebParts){
             Write-Host "Deleting web part id: " $webbie.Id
             $webbie.DeleteWebPart()

             $ctx.ExecuteQuery()
        }
}
 
 
 
  # Paths to SDK. Please verify location on your computer.
Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.dll" 
Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.Runtime.dll"


$AdminPassword=Read-Host -Prompt "Enter password" -AsSecureString
$username="t@trial567.onmicrosoft.com"
$Url="https://trial567.sharepoint.com/sites/powie64"


Remove-SPOWebParts -Username $username -Url $Url -password $AdminPassword -pageUrl "/sites/powie64/SitePages/pgie.aspx"

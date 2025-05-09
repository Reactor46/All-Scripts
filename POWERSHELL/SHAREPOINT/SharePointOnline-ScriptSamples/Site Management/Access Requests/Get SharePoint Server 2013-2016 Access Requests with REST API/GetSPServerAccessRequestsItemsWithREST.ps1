﻿function Get-AccessRequests{
	param (
		[Parameter(Mandatory=$true,Position=1)]
		[string]$Username,
		[Parameter(Mandatory=$true,Position=2)]
		$password,
		[Parameter(Mandatory=$true,Position=3)]
		[string] $url
	)


	$Credentials = New-Object System.Net.NetworkCredential($Username, $password)
	# $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username, $password)
	$RestUrl=$url+"/_api/web/lists/getbytitle('Access Requests')/items"

	$request = [System.Net.WebRequest]::Create($RESTUrl) 
	$request.Credentials = $Credentials 
	$request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f") 
	$request.Accept = "application/json;odata=verbose" 
	[Microsoft.PowerShell.Commands.WebRequestMethod]$Method = [Microsoft.PowerShell.Commands.WebRequestMethod]::Get
	$request.Method=$Method 
	$response = $request.GetResponse() 
	$requestStream = $response.GetResponseStream() 
	$read = New-Object System.IO.StreamReader $requestStream 
	$data=$read.ReadToEnd() 
	# Because the string contains duplicate 'ID' keys.
	$results = $data.ToString().Replace("ID", "_ID") | ConvertFrom-Json

	return ($results.d.results)
}

#Enter the data
$AdminPassword=Read-Host -Prompt "Enter password" -AsSecureString
$username="administrator"
$Url="http://dva"

Get-AccessRequests -Username $username -Url $Url -password $AdminPassword 

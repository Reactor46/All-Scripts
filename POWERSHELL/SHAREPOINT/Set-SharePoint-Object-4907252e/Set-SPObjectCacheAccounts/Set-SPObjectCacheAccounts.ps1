Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

<#
.SYNOPSIS
Helper function to set the the SharePoint object cache accounts.	

.DESCRIPTION
This function sets the SharePoint web application super user/reader object cache accounts to the accounts provided.

.EXAMPLE
Set-SPObjectCacheAccounts -WebAppURL http://www.example.com -SuperUser "DOMAIN\superuser" -SuperReader "DOMAIN\superreader"

This command sets the object cache accounts for the http://www.example.com web application to DOMAIN\superuser and DOMAIN\superreader, respectively.

.EXAMPLE
Set-SPObjectCacheAccounts -WebAppURL http://www.example.com

This command sets the object cache accounts for the http://www.example.com web application back to SharePoint's defaults: NT Authority\System and NT Authority\Local Service, respectively.

.EXAMPLE
Set-SPObjectCacheAccounts

This command sets the object cache accounts for all the web applications back to SharePoint's defaults: NT Authority\System and NT Authority\Local Service, respectively.

.EXAMPLE
Set-SPObjectCacheAccounts -SuperUser "DOMAIN\superuser" -SuperReader "DOMAIN\superreader"

This command sets the object cache accounts for all the web applications to DOMAIN\superuser and DOMAIN\superreader, respectively.

.NOTES
NAME: Set-SPObjectCacheAccounts
AUTHOR: Wes Kroesbergen

.LINK
http://www.kroesbergens.com
#>

Function Set-SPObjectCacheAccounts 
{
	Param([Parameter(Position=0,Mandatory=$false)][string] $WebAppURL,
	[Parameter(Position=1,Mandatory=$false)][string] $SuperUser,
	[Parameter(Position=2,Mandatory=$false)][string] $SuperReader)

    # Check if there is a URL specified
	if ($WebAppURL -notlike "")
	{
		$WebApp = Get-SPWebApplication $WebAppURL
		$WebApps = @($WebApp)
	}

    # If no URL specified, let's ask user if we should update all web apps
	else
	{
        $Response = ""

        # Validate a y/n response
        While ($Response -notmatch '^(Y|N)$')
        {
            $Response = Read-Host "No webapp specified, update all of them? (y/n)"
        }

        # If user selects Y, then update all
        if ($Response -like "Y")
        {
            Write-Host "Updating all web applications!"  -fore yellow 
            $WebApps = Get-SPWebApplication
        }

        # If user selects N, cancel
        else
        {
            Write-Host "Update cancelled."
            $WebApps = @()
        }
	}
	
	foreach ($WebApp in $WebApps)
	{

        # Check if web app is claims
	    $IsClaims = $WebApp.UseClaimsAuthentication

        # Get descriptive URL for user feedback
	    $URL = $WebApp.URL

        # Ensure user has not already set claims format of username
	    $SuperUser = $SuperUser.Replace("i:0#.w|","")
	    $SuperReader = $SuperReader.Replace("i:0#.w|","")

	    if ($SuperUser -like "")
	    {
		    Write-Host "Super User wasn't defined, resetting to SharePoint defaults..."
		    $SuperUser = "NT Authority\System"
	    }

	    if ($SuperReader -like "")
	    {
		    Write-Host "Super Reader wasn't defined, resetting to SharePoint defaults..."
		    $SuperReader = "NT Authority\Local Service"
	    }

        # If web app was claims, update format of username accordingly, as long as they are not the default accounts
	    if ($IsClaims -eq $true -and $SuperUser -notlike "*System" -and $SuperReader -notlike "*Service")
	    {
		    Write-Host "Web application uses claims, ensuring user identities have correct prefix..."
		    $SuperUser = ("i:0#.w|" + $SuperUser)
		    $SuperReader = ("i:0#.w|" + $SuperReader)
	    }

	    try 
	    { 
            # Update web app User Policy
		    $SUpolicy = $WebApp.Policies.Add($SuperUser,"Super User") 
		    $SUpolicy.PolicyRoleBindings.Add($WebApp.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::FullControl)) 
		    $SRpolicy = $WebApp.Policies.Add($SuperReader,"Super Reader") 
		    $SRpolicy.PolicyRoleBindings.Add($WebApp.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::FullRead)) 

            # Update the actual web app object cache accounts
		    $WebApp.Properties["portalsuperuseraccount"] = [string]$SuperUser 
		    $WebApp.Properties["portalsuperreaderaccount"] = [string]$SuperReader

            # Commit our changes
		    $WebApp.Update()

		    Write-Host "Object cache accounts for $URL have been updated." 
	    } 
	    catch 
	    { 
		    Write-Host "Error adding the object cache accounts for $URL!" -fore red 
	    } 
	}
} 
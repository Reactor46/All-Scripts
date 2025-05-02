#region Information
	#Description: This script removes and updates the search navigation of all subsites in a site collection or in a web application.
	#Author: Kampan 14.01.2015
#endregion

#region Parameters
param(
	[ValidateSet('site','webapplication')]
	[string]$Scope,
	
	[string]$URL,
	
	[bool]$RemoveExisting = $false
	)
if(-not($Scope)) { Throw "You must supply a value for -Scope. Examples -Scope site or -Scope webapplication" }
if(-not($URL)) { Throw "You must supply a value for -URL. Examples -URL http://mysharepoint.com" }
#endregion

#region Fuctions

#Save the search navigation nodes
$masterweb = Get-SPWeb $URL
$masternodes = $masterweb.Navigation.SearchNav

function Remove-SearchNav([string]$SiteURL)
{
	$site = Get-SPSite $SiteURL
	
	 foreach ($web in $site.AllWebs) 
	{
		$SearchNav = $web.Navigation.SearchNav
		
		IF ($SearchNav -ne $NULL -and $web.url -ne $site.Rootweb.Url)
		{
			Write-Host -ForegroundColor Green "Removeing Search Navigation for" $web.url;
			foreach($node in $web.Navigation.SearchNav)
			{
				$web.Navigation.SearchNav.delete($node)
			}
		}
		$web.Dispose()
	}
	 
	$site.Dispose()
	Write-Host -ForegroundColor Red "============================================="
}


function Update-SearchNav([string]$SiteURL)
{
 Write-Host -ForegroundColor Red "============================================="
 Write-Host -ForegroundColor Green "Updating Search Navigation at URL " -NoNewline;
 Write-Host -ForegroundColor Green $SiteURL

 remove-varable 
 $site = Get-SPSite $SiteURL

 foreach ($web in $site.AllWebs) 
 { 
  Write-Host -ForegroundColor Red "============================================="
  Write-Host -ForegroundColor Green "Updating Search Navigation at URL " -NoNewline;
  Write-Host -ForegroundColor Green $web.Url
  
  $SearchNav = $web.Navigation.SearchNav
  
  IF ($SearchNav -ne $NULL)
  {
   Write-Host -ForegroundColor Red "This Site Search Navigation Already containing values";
  }
  ELSE
  {
   foreach ($node in $masternodes)
   {
	Write-Host -ForegroundColor Green "Adding Search Navigation" $node.Title;
	$web.Navigation.SearchNav.AddAsLast($node)
   }
  }
   $web.Dispose()
 }
 $site.Dispose()
 Write-Host -ForegroundColor Red "============================================="
}

If ($Scope -eq "webApplication")
{
	$WebApplication = Get-SPWebapplication $URL
	
	if($RemoveExisting -eq $true)
	{
		Foreach ($Sites in $WebApplication.Sites)
		{
			Remove-SearchNav($Sites.URL)
		}
		Write-Host "Cleared the existing search navigation"
	}
	$site.Dispose();
	Foreach ($Sites in $WebApplication.Sites)
	{	
		Update-SearchNav($Sites.url.trim())
	}
}
elseif ($Scope -eq "site")
{
	if($RemoveExisting -eq $true)
	{
		Remove-SearchNav($URL)
		Write-Host "Cleared the existing search navigation"
	}
	Update-SearchNav($URL)
}

Write-Host "Press any key to continue ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

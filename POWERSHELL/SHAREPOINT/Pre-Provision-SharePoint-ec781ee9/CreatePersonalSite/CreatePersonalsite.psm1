#--------------------------------------------------------------------------------- 
#The sample scripts are not supported under any Microsoft standard support 
#program or service. The sample scripts are provided AS IS without warranty  
#of any kind. Microsoft further disclaims all implied warranties including,  
#without limitation, any implied warranties of merchantability or of fitness for 
#a particular purpose. The entire risk arising out of the use or performance of  
#the sample scripts and documentation remains with you. In no event shall 
#Microsoft, its authors, or anyone else involved in the creation, production, or 
#delivery of the scripts be liable for any damages whatsoever (including, 
#without limitation, damages for loss of business profits, business interruption, 
#loss of business information, or other pecuniary loss) arising out of the use 
#of or inability to use the sample scripts or documentation, even if Microsoft 
#has been advised of the possibility of such damages 
#--------------------------------------------------------------------------------- 

Function New-OSCPersonlSite
{
<#
 	.SYNOPSIS
        New-OSCPersonlSite is an advanced function which can be used to create personal site for each user in a SharePoint site.
    .DESCRIPTION
        New-OSCPersonlSite is an advanced function which can be used to create personal site for each user in a SharePoint site.
    .PARAMETER  SiteUrl
		The specified site URL.
    .EXAMPLE
        C:\PS>  New-OSCPersonlSite -SiteURL "http://sp2010:8888/sites/TopSite2"
		
		This command shows how to create personal site for each user in site "http://sp2010:8888/sites/TopSite2".
#>
	[CmdletBinding()]
	Param
	(
	    [Parameter(Mandatory = $True,Position=0)]
		[String]$SiteURL
	)
	#Add "Microsoft.SharePoint.PowerShell" Snapin
	if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null) 
	{
	    Add-PSSnapin "Microsoft.SharePoint.PowerShell"
	}
	#Load "Microsoft.Office.Server" Assembly
	[Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Server") | Out-Null 
	#Get SharePoint site
	$Site = Get-SPSite -Identity $SiteURL
	#Get service context
	$context = Get-SPServiceContext -Site $site
	$upm =  New-Object Microsoft.Office.Server.UserProfiles.UserProfileManager($context)
	$AllProfiles = $upm.GetEnumerator()
	#Create personal site for each user
	foreach($profile in $AllProfiles)
	{
		$AccountName = $profile[[Microsoft.Office.Server.UserProfiles.PropertyConstants]::AccountName].Value    
		Try
		{
			if($profile.PersonalSite -eq $Null)
			{
				write-host "Creating personel site for  $AccountName" 
				$profile.CreatePersonalSite()       
				write-host "Personal Site Admin has assigned" 
			}
			else
			{
				Write-Warning "$AccountName already has personel site"
			}
		}
		Catch
		{
			Write-Error "Failed to create personal site for '$AccountName'"
		}
	} 
	$Site.Dispose();
}



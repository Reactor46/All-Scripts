$snapin = Get-PSSnapin | Where-Object {$_.Name -eq 'Microsoft.SharePoint.Powershell'} 
if ($snapin -eq $null) 
{    
	Write-Host "Loading SharePoint Powershell Snapin"    
	Add-PSSnapin "Microsoft.SharePoint.Powershell" 
}
#Name of MMS Proxy
$mmsServiceName = "Metadata Service Application Proxy"
#Name of Group
$grpName = "CustomGroup"
 
[xml]$xmlinput =  (get-content .\Config.xml)

$item = $xmlinput.Config
$WebApplications = $item.WebApplications
foreach($WebApplication in $WebApplications.WebApplication)
{

	#iterates though all the site collections 
	$Mywebapplication = Get-SPWebApplication $WebApplication.url
	#$Mywebapplication.FormDigestSettings.Enabled = $false
	foreach($site in $Mywebapplication.Sites)
	{
		#Declaring the Name of the Term Set and Metadata Columns, it is the same in both cases.
		$customContentTypesArray = @("Country", "City", "County", "Province")
		Write-Host "In the Site Collection" $site.url
		foreach ($taxonomy in $customContentTypesArray) 
		{
			
			$spweb = $site.RootWeb		
			$session = new-object Microsoft.SharePoint.Taxonomy.TaxonomySession($site)
			$termStore = $Session.TermStores[$mmsServiceName];
			$group=$termStore.Groups[$grpName]
			$termSet = $group.TermSets[$taxonomy]					
			$targetField = [Microsoft.SharePoint.Taxonomy.TaxonomyField]$spweb.Fields[$taxonomy]
			if($targetField -ne $null)
			{
				 Write-Host "Connecting with Term set!" $termSet.id " Namely"  $termSet.Name 
				$targetField.sspid = $termstore.id
				$targetField.termsetid = $termSet.id
				$targetField.Update($true)								
			}	
			$spweb.Dispose()
								
		}		
    }	
	$site.Dispose()	
}
       


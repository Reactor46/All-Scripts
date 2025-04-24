Param([string]$rootSiteURL)

Function FixPage() {
	Param (
		[string]$siteCollectionPath,
		[string]$webUrl,
		[string]$pageName
	)

	Write-Host $webUrl
	$spWeb = Get-SPWeb($webUrl)
	$filePath = $webUrl+"/Pages/"+$pageName
	$spFile = $spWeb.GetFile($filePath)
	$spFile.UndoCheckOut()
	$spFile.CheckOut("Online",$null)
	$spFile.Properties["PublishingPageLayout"] = $siteCollectionPath+"/_catalogs/masterpage/DetailPage.aspx, Detail Page Layout"
	$spFile.Update()
	$spFile.CheckIn("Update page layout via PowerShell",[Microsoft.SharePoint.SPCheckinType]::MajorCheckIn)
	#$spFile.Approve("Approved page layout via PowerShell")
	$spWeb.Dispose()

}#end function

#run through the sites and fix any pages that are checked out
Function SearchAndFix() {

	$rootSite = Get-SPsite $rootSiteURL

	#Create Site Array
	#$RootSite = $env:PulseSite
	$SiteArray = "/strategic-planning/ehs/"

	foreach ($subSite in $SiteArray) {

		#get the full url
		$fullURL = [string]::Concat($rootSiteUrl,$subSite)
		$web = Get-SPWeb($fullURL)
		$pagesLibrary = $web.GetFolder("Pages")
		$pagesLibrary.EnableModeration = $true

		Write-Host "Searching $($web.url)" 
		Write-Host "------------------------------------------------"

		$pagesLibrary.Files | Where { $_.CheckOutStatus -ne "None" } | ForEach {

			Write-Host "Updating $($_.Name)"
			FixPage $web.Name $web.url $_.Name
			$_.Approve("Approved page layout via PowerShell")

		}#end ForEach

		Write-Host `n
		
		$web.Dispose()

	}#end foreach

	$rootSite.Dispose()

}#end function

#main
Function Main() {
	SearchAndFix 
}#end function

Main
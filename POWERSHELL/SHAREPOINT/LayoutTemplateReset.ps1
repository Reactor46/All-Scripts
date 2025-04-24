
Param(
	[string]$RootUrl = $(throw "Root Site URL required.")
)#end params

##### Global Variables
##### ----------------------------------------------------------------------------------------------------------------------------------------
#$RootSite = $env:PulseSite
$LogsDirectory = "D:\PowerShell_Logs"
$Logfile = "$($LogsDirectory)\LogFile $(get-date -Format MM-dd-yyyy_hh-mm-ss).log"
$pulseBrandingComponentsFeatureGuid = "48b42642-6cf8-463c-bc34-c25cb3792842"
$pulseTermStoreSetupFeatureGuid = "2237983c-9758-4869-830b-ffb665631fba"
$rootsiteMasterpageUrl = "/_catalogs/masterpage/RootSite.master"
$subsiteMasterpageUrl ="/_catalogs/masterpage/SubSite.master"

$SiteArray = “/strategic-planning/ehs”

Function CreateLogsDirectory() {
	
    Write-Host `n
	Write-Host "Checking if exists: $($LogsDirectory)"
    Write-Host "---------------------------"

	if( !(Test-Path -Path $LogsDirectory) ) {

		New-Item -ItemType directory -Path $LogsDirectory
		Write-Host "Created Directory: $($LogsDirectory)"

	}#end if

	Write-Host -fore Green "$($LogsDirectory) Exists"

}#end function

Function LogComment( $comment ) {
   Add-content $Logfile -value $comment
}#end function

#add page layouts
Function AddPageLayoutsFor( $parentWeb ) {

    $publishingSite = New-Object Microsoft.SharePoint.Publishing.PublishingSite( $parentWeb.Site )				
    $publishingWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb( $parentWeb )

    #get current web available page layouts collection
    $layoutsToSet = @()

    #get collection of all the page layouts in Site collection
    $pageLayouts = $publishingSite.PageLayouts

    $requiredLayouts = "HomePage.aspx", "DetailPage.aspx"

    #loop thru layouts wanting to be added
    foreach( $layout in $requiredLayouts ) {

        #pull the required layouts out of the page layouts collection
        #add to the layouts that you want to add to the layouts to set
        $layoutsToSet += $pageLayouts | ? { $_.Name -eq $layout }

    }#end foreach
	
    $publishingWeb.SetAvailablePageLayouts( $layoutsToSet, $false )
    $publishingWeb.Update()

    foreach( $item in $layoutsArray ) {

        $layout = $layouts | ? { $_.Name -eq $item }

        if( $layout -ne $null ) {

            LogComment "$($layout.Name) layout added to available page layouts for: $($parentWeb.Url)"
	        Write-Host -fore Green "$($layout.Name) layout added to available page layouts for: $($parentWeb.Url)"

        } else {

            LogComment "$($layout.Name) layout NOT added to available page layouts for: $($parentWeb.Url)"
	        Write-Host -fore Green "$($layout.Name) layout NOT added to available page layouts for: $($parentWeb.Url)"  
      
        }#end if

    }#end foreach

}#end function

#set default page layout
Function SetDefaultPageLayoutFor( $parentWeb ) {

    $publishingSite = New-Object Microsoft.SharePoint.Publishing.PublishingSite( $parentWeb.Site )				
    $publishingWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb( $parentWeb )
    $pageLayouts = $publishingweb.GetAvailablePageLayouts()

    $layout = $null
    if( $parentWeb.ServerRelativeUrl -eq "/" ) {
        $layout = $pageLayouts | ? { $_.Name -eq "HomePage.aspx" }
    } else {
        $layout = $pageLayouts | ? { $_.Name -eq "DetailPage.aspx" }
    }#end if

    if( $layout -ne $null ) {

        $publishingWeb.SetDefaultPageLayout( $layout, $false )
        $publishingWeb.Update()

        LogComment "Default Page layout has been set to: $($layout.Name)"
        Write-Host -fore Green "Default Page layout has been set to: $($layout.Name)"
            
    } else {

        LogComment "Default Layout Not Set"
        Write-Host -fore Red "Default Layout Not Set" 
           
    }#end if
    
}#end function

#iterate through all sites
Function IterateSubsites() {

    Write-Host `n
    Write-Host "Iterating Sites"

	foreach( $site in $SiteArray ) {
        
        #get the full site url
	    $siteURL = [string]::Concat( $RootUrl, $site )

	    #get the site object
	    $site = Get-SPsite $siteURL

        foreach( $parentWeb in $site.AllWebs ) {

            $spWeb = Get-SPWeb -identity $parentWeb.Url

            #Add new page layouts
            LogComment "Adding New Page Layouts for: '$($spWeb.Url)'"
		    Write-Host "Adding New Page Layouts for: '$($spWeb.Url)'"
            Write-Host "---------------------------"
            AddPageLayoutsFor $spWeb
		    Write-Host `n

            #Setting Default Page Layout
            LogComment "Setting Default Page Layout for: '$($spWeb.Url)'"
		    Write-Host "Setting Default Page Layout for: '$($spWeb.Url)'"
            Write-Host "---------------------------"
            SetDefaultPageLayoutFor $spWeb
		    Write-Host `n

        }#end foreach

    }#end foreach

}#end function

##### Init Point
##### ----------------------------------------------------------------------------------------------------------------------------------------
Function Main() {
	
	#create directory if not exists
	CreateLogsDirectory
	
	#iterate through all the subsites
	IterateSubsites

}#end function

##### Entry Point
##### ----------------------------------------------------------------------------------------------------------------------------------------
[Microsoft.SharePoint.SPSecurity]::RunWithElevatedPrivileges({ Main })
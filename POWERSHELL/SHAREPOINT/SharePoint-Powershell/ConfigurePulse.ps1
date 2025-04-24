#
#Purpose: 
#			Update the Pulse Intranet sitecollection doing various things 
#				1. Setting masterpage and pagelayouts.
#				2. Creating and setting managed metadata

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

$SiteArray = “/”, “/business-development/Marketing”, “/business-development/occupational-medicine”, “/business-development/wellness”, “/corporate-finance/cash-accounting”, 
“/corporate-finance/home”, “/corporate-finance/payroll”, “/corporate-finance/purchasing”, “/HCF/cbo”, “/hcf/HCC-Coding”, “/hcf/home”, “/HCF/HPPS”, “/hps/home”, 
“/hr/home”, “/hr/payroll”, “/IT/EMR”, “/it/helpdesk”, “/it/home”, “/it/infosec”, “/it/support-information”, “/IT/Telecom”, “/IT/Training”, “/managed-care/home”, 
“/Managed-Care/KCA”, “/Operations/AHNT”, “/operations/ASC”, “/Operations/cancer-center”, “/Operations/Cancer-Services”, “/operations/clinical-education”, 
“/operations/contact-center”, “/operations/home”, “/operations/labcorp”, “/operations/labcorponsite”, “/operations/lab-services”, “/operations/medical-imaging”, 
“/operations/patient-education”, “/Operations/Pharmacy”, “/Operations/Sleep-Center”, “/patient-experience/home”, “/physicians-lunchroom/home”, 
“/strategic-planning/coding”, “/strategic-planning/ehs”, “/strategic-planning/facilities”, “/strategic-planning/forms”, “/strategic-planning/medical-records”, 
“/strategic-planning/qi”, “/strategic-planning/risk-management”

#$SiteArray = "/", "/hr/home", "/strategic-planning/ehs"

##### Functions
##### ----------------------------------------------------------------------------------------------------------------------------------------

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

Function IsFeatureActivated() {

    param(
		[string]$GuidId = $(throw "-GuidId parameter is required!"),
		[Microsoft.SharePoint.SPFeatureScope]$Scope = $(throw "-Scope parameter is required!"),
		[string]$Url
	)#end params

    if( $Scope -ne "Farm" -and [string]::IsNullOrEmpty($Url) ) {
        throw "-Url parameter is required for scopes WebApplication, Site, and Web"
    }#end if

    $feature = $null

    switch( $Scope ) {

        "Farm" { $feature = Get-SPFeature $GuidId -Farm -erroraction SilentlyContinue }
        "WebApplication" { $feature = Get-SPFeature $GuidId -WebApplication $Url -erroraction SilentlyContinue }
        "Site" { $feature = Get-SPFeature $GuidId -Site $Url -erroraction SilentlyContinue }
        "Web" { $feature = Get-SPFeature $GuidId -Web $Url -erroraction SilentlyContinue }

    }#end switch

    #return feature found or not (activated at scope) in the pipeline
    return $feature -ne $null

}#end function

Function EnableSpFeature() {
	
	Param(
		[string]$siteURL = $(throw "Site URL Required to check or enable feature"),
		[string]$scope,
		[string]$featureGiudId = $(throw "Feature Guid Id Required to check or enable feature")
	)#end params

	if( IsFeatureActivated $featureGiudId $scope $siteURL ) {

		LogComment "Feature $featureGiudId is already active"
		Write-Host -fore Green "Feature $featureGiudId is already active"
		return

	}#end if

    LogComment "Activating Feature: $featureGiudId"
	Write-Host -fore Yellow "Activating Feature: $featureGiudId"

	enable-spfeature -identity $featureGiudId -url $siteURL #-erroraction SilentlyContinue

	if(IsFeatureActivated $featureGiudId $scope $siteURL) {

	    LogComment "Feature: $featureGiudId Activated"
	    Write-Host -fore Green "Feature: $featureGiudId Activated"

	}#end if

}#end function

Function EnableModerationFor( $parentWeb ) {

    $pagesLibrary = $parentWeb.Lists["Pages"]
    $pagesLibrary.EnableModeration = "True"
    $pagesLibrary.Update()
    LogComment "Complete"
	Write-Host -fore Green "Complete"

}#end function

Function AllowAllWebTemplatesFor( $parentWeb ) {

	$parentWeb.AllowAllWebTemplates()
	$parentWeb.Update()
    LogComment "Complete"
	Write-Host -fore Green "Complete"

}#end function

Function UpdateMasterPageFor( $parentWeb ) {
			
    if( $parentWeb.ServerRelativeUrl -eq "/" ) {

	    LogComment "Updating the Site Master Page for: '$($parentWeb.Url)'"
	    Write-Host -fore Yellow "Updating the Site Master Page for: '$($parentWeb.Url)'"

	    $parentWeb.CustomMasterUrl = $rootsiteMasterpageUrl
        $parentWeb.Update()

	    LogComment "Updated the Site Master Page for: '$($parentWeb.CustomMasterUrl)'"
	    Write-Host -fore Green "Updated the Site Master Page to: '$($parentWeb.CustomMasterUrl)'" 

        foreach($childWeb in $parentWeb.Webs) {

            $childWeb.SetProperty("_InheritsCustomMasterUrl", "False")
            $childWeb.CustomMasterUrl = $subsiteMasterpageUrl
            $childWeb.Update()
                    
            LogComment "Updated the Child Web: '$childWeb' Master Page to: '$($childWeb.CustomMasterUrl)'"
		    Write-Host -fore Green "Updated the Child Web '$childWeb' Master Page to: '$($childWeb.CustomMasterUrl)'" 

        }#end foreach

    } else { 

	    LogComment "Updating the Subsite Master Page for: '$($parentWeb.Url)'"
	    Write-Host -fore Yellow "Updating the Subsite Master Page for: '$($parentWeb.Url)'"

        if( $siteUrl -eq "$($RootUrl)/" ) {

            $parentWeb.CustomMasterUrl = "$subsiteMasterpageUrl"                
            $parentWeb.Update()

        } else {

            $parentWeb.CustomMasterUrl = "$($site.ServerRelativeUrl)$subsiteMasterpageUrl"
		    $parentWeb.Update()

        }#end if

        LogComment "Updated the Master Page to: '$($parentWeb.CustomMasterUrl)'"
	    Write-Host -fore Green "Updated the Master to: '$($parentWeb.CustomMasterUrl)'"   

        foreach($childWeb in $parentWeb.Webs) {

            $childWeb.CustomMasterUrl = "$($parentWeb.ServerRelativeUrl)$subsiteMasterpageUrl"
            $childWeb.Update()

            LogComment "Updated the Child Web '$childWeb' Master Page to: '$($childWeb.CustomMasterUrl)'"
		    Write-Host -fore Green "Updated the Child Web '$childWeb' Master Page to: '$($childWeb.CustomMasterUrl)'"

        }#end foreach

    }#end if

}#end function

Function AddPageLayoutsFor( $parentWeb ) {

    $publishingSite = New-Object Microsoft.SharePoint.Publishing.PublishingSite($parentWeb.Site)				
    $publishingWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($parentWeb)

    #get current web available page layouts collection
    $layouts = $publishingweb.GetAvailablePageLayouts()

    #get collection of all the page layouts in Site collection
    $currentLayouts = $publishingSite.PageLayouts

    $layoutsArray = "HomePage.aspx", "DetailPage.aspx"

    #loop thru layouts wanting to be added
    foreach($item in $layoutsArray) {

        #add to the layouts that you want to add to the layouts
        $layouts += $currentLayouts | ? { $_.Name -eq $item }

    }#end foreach
	
    $publishingWeb.SetAvailablePageLayouts($layouts, $false)
    $publishingWeb.Update()

    foreach($item in $layoutsArray) {

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

Function SetDefaultPageLayoutFor( $parentWeb ) {

    $publishingSite = New-Object Microsoft.SharePoint.Publishing.PublishingSite($parentWeb.Site)				
    $publishingWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($parentWeb)
    $pageLayouts = $publishingweb.GetAvailablePageLayouts()

    $layout = $null
    if( $parentWeb.ServerRelativeUrl -eq "/" ) {
        $layout = $pageLayouts | ? { $_.Name -eq "HomePage.aspx" }
    } else {
        $layout = $pageLayouts | ? { $_.Name -eq "DetailPage.aspx" }
    }#end if

    if($layout -ne $null) {

        $publishingWeb.SetDefaultPageLayout($layout, $false)
        $publishingWeb.Update()

        LogComment "Default Page layout has been set to: $($layout.Name)"
        Write-Host -fore Green "Default Page layout has been set to: $($layout.Name)"
            
    } else {

        LogComment "Default Layout Not Set"
        Write-Host -fore Red "Default Layout Not Set" 
           
    }#end if
    
}#end function

Function AddContentTypesFor( $parentWeb ) {

    $publishingSite = New-Object Microsoft.SharePoint.Publishing.PublishingSite($parentWeb.Site)				
    $publishingWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($parentWeb)
    $pagesLibrary = $parentWeb.Lists["Pages"]
    
    #add content type(s)
    if( $pagesLibrary -ne $null ) {

        $pagesLibrary.ContentTypesEnabled = $true
        $pagesLibrary.Update()

        $homePageContentType = $parentWeb.ContentTypes | ? { $_.Name -eq "Home Page Layout" }

        if($parentWeb.IsRootWeb){
            $detailPageContentType = $parentWeb.ContentTypes | ? { $_.Name -eq "Detail Page Layout" }
        } else {
            $detailPageContentType = $parentWeb.ParentWeb.ContentTypes | ? { $_.Name -eq "Detail Page Layout" }
        }#end if

        if( $parentWeb.ServerRelativeUrl -eq "/" ) {
            
            if( ($pagesLibrary.ContentTypes | ? { $_.Name -eq "Home Page Layout" }) -eq $null ) {

                $pagesLibrary.ContentTypes.Add($homePageContentType)
                $pagesLibrary.Update()

                LogComment "$($homePageContentType.Name) ContentType Added for: $($parentWeb.Url)"
		        Write-Host -fore Green "$($homePageContentType.Name) ContentType Added for: $($parentWeb.Url)" 

            } else {

                LogComment "$($homePageContentType.Name) ContentType already added for: $($parentWeb.Url)"
		        Write-Host -fore Green "$($homePageContentType.Name) ContentType already added for: $($parentWeb.Url)"

            }#end if

            if( ($pagesLibrary.ContentTypes | ? { $_.Name -eq "Detail Page Layout" }) -eq $null) {

                $pagesLibrary.ContentTypes.Add($detailPageContentType)
                $pagesLibrary.Update()

                LogComment "$($detailPageContentType.Name) ContentType Added for: $($parentWeb.Url)"
		        Write-Host -fore Green "$($detailPageContentType.Name) ContentType Added for: $($parentWeb.Url)"  

            } else {

                LogComment "$($detailPageContentType.Name) ContentType already added for: $($parentWeb.Url)"
		        Write-Host -fore Green "$($detailPageContentType.Name) ContentType already added for: $($parentWeb.Url)"

            }#end if
            
        } else {
            
            $childPagesLibrary = $parentWeb.Lists["Pages"]
            $childPagesLibrary.ContentTypesEnabled = $true
            $childPagesLibrary.Update()

            $childContentType = $null

            if($parentWeb.IsRootWeb) {
                $childContentType = $parentWeb.ContentTypes | ? { $_.Name -eq "Detail Page Layout" }
            } else {
                $childContentType = $parentWeb.ParentWeb.ContentTypes | ? { $_.Name -eq "Detail Page Layout" }
            }#end if

            if( $childContentType -ne $null ) {
                
                if( ($childPagesLibrary.ContentTypes | ? { $_.Name -eq "Detail Page Layout" }) -ne $null) {

                    LogComment "$($childContentType.Name) ContentType already added for: $($parentWeb.Url)"
		            Write-Host -fore Green "$($childContentType.Name) ContentType already added for: $($parentWeb.Url)"  

                } else {

                    $childPagesLibrary.ContentTypes.Add($childContentType)
                    $childPagesLibrary.Update()

                    LogComment "$($childContentType.Name) ContentType Added for: $($parentWeb.Url)"
		            Write-Host -fore Green "$($childContentType.Name) ContentType Added for: $($parentWeb.Url)"
               
                }#end if

            } else {

                LogComment "ContentType Null for: $($parentWeb.Url)"
		        Write-Host -fore red "ContentType Null for: $($parentWeb.Url)"

            }#end if
                
        }#end if

    } else {

        LogComment "Pages Library is Null"
		Write-Host -fore Red "Pages Library is Null"    

    }#end if

}#end function

Function UpdatePagesLibraryLayoutsFor( $parentWeb ) {

    $publishingSite = New-Object Microsoft.SharePoint.Publishing.PublishingSite($parentWeb.Site)				
    $publishingWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($parentWeb)
    $pageLayouts = $publishingSite.GetPageLayouts($false)
    $pagesLibrary = $parentWeb.Lists["Pages"]

    #$query = New-Object Microsoft.SharePoint.SPQuery
	#$query.ViewAttributes = "Scope='RecursiveAll'"
	#$query.Query = "<Where><Eq><FieldRef Name='ContentType' /><Value Type='Computed'>Page</Value></Eq></Where>"
    #$Pages = $publishingWeb.GetPublishingPages($query)
    
    $Pages = $publishingWeb.GetPublishingPages()

    foreach( $page in $Pages ) {

	    if( $page.ListItem.File.Level -eq "Checkout" -or $page.ListItem.File.CheckOutStatus -ne "None" ) {
		    $page.ListItem.File.UndoCheckOut()
		    LogComment "UndoCheckOut performed on $($page.Name)"
		    Write-Host -fore Yellow "UndoCheckOut performed on $($page.Name)"
	    }#end if

	    $page.CheckOut()
				   
        $layout = $null
        $contentType = $null
        if( ($parentWeb.ServerRelativeUrl -eq "/") -and ($page.Name -eq "default.aspx") ) {

            $contentType = $parentWeb.ContentTypes | ? { $_.Name -eq "Home Page Layout" }
            $layout = $pageLayouts | ? { $_.Name -eq "HomePage.aspx" }

        } else {

            if($parentWeb.IsRootWeb) {
                $contentType = $parentWeb.ContentTypes | ? { $_.Name -eq "Detail Page Layout" }
            } else {
                $contentType = $parentWeb.ParentWeb.ContentTypes | ? { $_.Name -eq "Detail Page Layout" }
            }#end if
            
            $layout = $pageLayouts | ? { $_.Name -eq "DetailPage.aspx" }

        }#end if

        LogComment "Updating '$($page.Name)' Page to: $($layout.Name) Page Layout [Current Level: $($page.ListItem.File.Level)]"
	    Write-Host -fore Yellow "Updating: '$($page.Name)' Page to: $($layout.Name) Layout [Current Level: $($page.ListItem.File.Level)]"

        $page.ListItem["ContentTypeId"] = $contentType.Id
        $page.ListItem.Update();
        
        $page.Layout = $layout
	    $page.ListItem.Update();

	    $page.CheckIn("Updated PublishingPageLayout field for: '$($page.Title)'")
	    $page.ListItem.File.Approve("Updated PublishingPageLayout field for: '$($page.Title)'")

	    LogComment "Update Complete"
	    Write-Host -fore Green "Update Complete"
        Write-Host "--"

    }#end foreach

}#end function

Function CreateTermSetItemsFor( $parentWeb ) {

    $session = New-Object Microsoft.SharePoint.Taxonomy.TaxonomySession($parentWeb.Url)

    if ($session.TermStores.Count -lt 1) {
        throw "No term stores found. The Taxonomy Service is offline or missing"
    }#end if

    [Microsoft.SharePoint.Taxonomy.TermStore]$store = $session.TermStores[0]

    #get the termset for the particular web
    [string]$providerName = [Microsoft.SharePoint.Publishing.Navigation.StandardNavigationProviderNames]::CurrentNavigationTaxonomyProvider
    [Microsoft.SharePoint.Publishing.Navigation.NavigationTermSet]$termSet = [Microsoft.SharePoint.Publishing.Navigation.TaxonomyNavigation]::GetTermSetForWeb($parentWeb, $providerName, $true)
    
    if($termSet -ne $null) {
    
        [Microsoft.SharePoint.Publishing.Navigation.NavigationTermSet]$editableTermSet = $termSet.GetAsEditable($session)
        $editableTermSet.IsNavigationTermSet = $true

        #$editableTermSet.Terms | ForEach-Object { Write-Host $_ }

        #create the term for each page
        foreach( $page in $parentWeb.Lists["Pages"].Items ) {

            #Write-Host $page.Title
            
            $pageCount = $($parentWeb.Lists["Pages"].Items.Count)
            if($pageCount -le 1) {

                LogComment "$($parentWeb) only includes $($pageCount) page(s). No terms needing creating"
                Write-Host -fore Green "$($parentWeb) only includes $($pageCount) page(s). No terms needing creating"
                break
                                             
            }#end if

		    if( $page.DisplayName -eq "default" ) { continue }
		    if( $page.DisplayName -eq "PageNotFoundError" ) { continue }

            $termValue = $editableTermSet.Terms | Where-Object { $_.Title -eq $page.DisplayName }

            if( $termValue -eq $null ) {

                [string]$linkType = [Microsoft.SharePoint.Publishing.Navigation.NavigationLinkType]::SimpleLink
                [Microsoft.SharePoint.Publishing.Navigation.NavigationTerm]$term = $editableTermSet.CreateTerm($page.Title, $linkType, [System.Guid]::NewGuid())
                $term.SimpleLinkUrl = [Microsoft.SharePoint.Publishing.PublishingPage]::GetPublishingPage($page).Uri.AbsoluteUri

                LogComment "Term Created: $($term)"
                Write-Host -fore Green "Term Created: $($term)"

            } else {

                LogComment "Term Created: $($termValue)"
                Write-Host -fore Green "Term Created: $($termValue)" 
                           
            }#end if

        }#end foreach    
    
    } else {

        LogComment "No termset found for: $($webUrl)"
        Write-Host -fore Red "No termset found for: $($webUrl)"

    }#end if

    $store.CommitAll()

}#end function

Function CopyXlsFileFor( $parentweb ) {

    $list = $parentweb.Lists["Style Library"]
    if($list -ne $null) {

        LogComment "Copying XLS file for: '$($parentweb.Url)'"
	    Write-Host "Copying XLS file for: '$($parentweb.Url)'"
        Write-Host "---------------------------"

        #get the file to copy
        $xlsSite = Get-SPSite $RootUrl
        $xlsWeb = $xlsSite.OpenWeb()
        $styleLibrary = $xlsWeb.Lists["Style Library"]
        $xlsFolder = $styleLibrary.ParentWeb.GetFolder("$($xlsWeb.Url)/$($styleLibrary.RootFolder.Url)/xsl style sheets")
        $xlsFile = $xlsFolder.Files["pulse_custom_itemstyle.xsl"]
        
        if($xlsFile.CheckOutType -ne "Online") {
            $xlsFile.CheckOut()
            LogComment $xlsFile.Name "Root File Checked Out"
            Write-Host -fore Yellow $xlsFile.Name "Root File Checked Out"
        }#end if       

        #get the byte[] of that file
        $stream = $xlsFile.OpenBinary()

        #get the folder to copy to
        $folder = $list.ParentWeb.GetFolder("$($parentweb.Url)/$($styleLibrary.RootFolder.Url)/xsl style sheets")
        
        #check out the file so it can be overwritten
        $file = $folder.Files["pulse_custom_itemstyle.xsl"]
        if($file.CheckOutType -ne "Online") {
            $file.CheckOut()
            LogComment $file.Name "File Checked Out"
            Write-Host -fore Yellow $file.Name "File Checked Out"
        }#end if

        #add the root file to the folder
        $done = $folder.Files.Add("$($xlsFile.Name)", $stream, $true)

        if($xlsFile.CheckOutType -eq "Online") {
            $xlsFile.CheckIn("Checked In")
            $xlsFile.Publish("Published")
        }#end if

        if($parentweb.ServerRelativeUrl -ne "/") {
            if($file.CheckOutType -eq "Online") {
                $file.CheckIn("Checked In")
                $file.Publish("Published")
            }#end if        
        }#end if

        #update the folder
        $folder.Update()

        #dispose the sites
        $xlsWeb.Dispose()
        $xlsSite.Dispose()

        LogComment "$($xlsFile.Name) copied to: '$($folder.Name)' Folder"
        Write-Host -fore Green "$($xlsFile.Name) copied to: '$($folder.Name)' Folder"
        
    }#end if
    
}#end Function

Function IterateSubsites() {
	
    Write-Host `n
    Write-Host "Iterating Sites"

	foreach( $site in $SiteArray ) {
        
        #$Host.UI.RawUI.ReadKey("NoEcho, IncludeKeyDown")     		

        #get the full site url
	    $siteURL = [string]::Concat($RootUrl,$site)

	    #get the site object
	    $site = Get-SPsite $siteURL

        Write-Host `n
        Write-Host "***************************"
        Write-Host "Updating: '$($siteUrl)'"
        Write-Host "***************************"
        Write-Host `n

        #enable site feature
        LogComment "Enabling Site Feature: Pulse Branding Components ($($pulseBrandingComponentsFeatureGuid)) in '$($siteUrl)'"
	    Write-Host "Enabling Site Feature: Pulse Branding Components ($($pulseBrandingComponentsFeatureGuid)) in '$($siteUrl)'"
        Write-Host "---------------------------"
	    EnableSpFeature $siteUrl "Site" $pulseBrandingComponentsFeatureGuid #-erroraction SilentlyContinue
        Write-Host `n

		foreach( $parentWeb in $site.AllWebs ) {
            
            #Write-Host `n
            #Write-Host "Press any Key to continue"
            #$Host.UI.RawUI.ReadKey("NoEcho, IncludeKeyDown")

            $parentWeb = Get-SPWeb -identity $parentWeb.Url

            #set content approval
            LogComment "Setting Content Approval for: '$($parentWeb.Url)'"
		    Write-Host "Setting Content Approval for: '$($parentWeb.Url)'"
            Write-Host "---------------------------"
            EnableModerationFor $parentWeb
		    Write-Host `n

            #enable site feature
            LogComment "Enabling Site Feature: Pulse Branding Components ($($pulseBrandingComponentsFeatureGuid)) in '$($siteUrl)'"
	        Write-Host "Enabling Site Feature: Pulse Branding Components ($($pulseBrandingComponentsFeatureGuid)) in '$($siteUrl)'"
            Write-Host "---------------------------"
	        EnableSpFeature $siteURL "Site" $pulseBrandingComponentsFeatureGuid #-erroraction SilentlyContinue
	        Write-Host `n  
				
            #enable web feature
			LogComment "Enabling Web Feature: Pulse Term Store Setup ($($pulseTermStoreSetupFeatureGuid)) in '$($parentWeb.Url)'"
			Write-Host "Enabling Web Feature: Pulse Term Store Setup ($($pulseTermStoreSetupFeatureGuid)) in '$($parentWeb.Url)'"
            Write-Host "---------------------------"
			EnableSpFeature $parentWeb.Url "Web" $pulseTermStoreSetupFeatureGuid #-erroraction SilentlyContinue
			Write-Host `n

			#allow web updates
			LogComment "Allowing all web templates"
			Write-Host "Allowing all web templates"
            Write-Host "---------------------------"
            AllowAllWebTemplatesFor $parentWeb
            Write-Host `n

            #add content type to parentweb
            LogComment "Adding ContentTypes to: '$($parentWeb.Url)'"
            Write-Host "Adding ContentTypes to: '$($parentWeb.Url)'"
            Write-Host "---------------------------"
            AddContentTypesFor $parentweb
            Write-Host `n				

            #set the web masterpage
            LogComment "Setting '$($parentWeb.Url)' Master Page"
            Write-Host "Setting '$($parentWeb.Url)' Master Page"
            Write-Host "---------------------------"
            UpdateMasterPageFor $parentweb
            Write-Host `n

			#add the page layouts for web
			LogComment "Adding Page Layouts for: '$($parentWeb.Url)'"
			Write-Host "Adding Page Layouts for: '$($parentWeb.Url)'"
            Write-Host "---------------------------"
			AddPageLayoutsFor $parentWeb
			Write-Host `n

			#set default page layout
            LogComment "Setting Default Page Layout for: '$($parentWeb.Url)'"
            Write-Host "Setting Default Page Layout for: '$($parentWeb.Url)'"
            Write-Host "---------------------------"
            SetDefaultPageLayoutFor $parentWeb			
            Write-Host `n

            #update page layouts
			LogComment "Updating Page Layouts [$($parentWeb.Lists["Pages"].Items.Count) Pages]"
			Write-Host "Updating Page Layouts [$($parentWeb.Lists["Pages"].Items.Count) Pages]"
            Write-Host "---------------------------"
            UpdatePagesLibraryLayoutsFor $parentWeb
            Write-Host `n

            #create term set items
			LogComment "Creating Term Set Items for: '$($parentWeb.Url)'"
			Write-Host "Creating Term Set Items for: '$($parentWeb.Url)'"
            Write-Host "---------------------------"
			CreateTermSetItemsFor $parentWeb
            Write-Host `n

            #copy xls file
			#header written in function
            CopyXlsFileFor $parentWeb
            Write-Host `n

            #dispose the web
            $parentWeb.Update()
			$parentWeb.Dispose()
			
            Write-Host -fore Green "***************************"
            Write-Host -fore Green "'$($parentWeb.Url)' Update Complete"
            Write-Host -fore Green "***************************"
            Write-Host `n

		}#end foreach
        
        #dispose the site
		$site.Dispose()

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
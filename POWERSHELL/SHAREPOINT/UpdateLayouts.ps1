#Created by David Harris on 03/26/2020
#Note- in order to add variables to string you need to add the variable as $($VariableName) within the string

#Create new Directory on D to store logs only if the directory does not already exist
$TARGETDIR = 'D:\PowerShell_Logs'
if(!(Test-Path -Path $TARGETDIR )){
    New-Item -ItemType directory -Path $TARGETDIR
}

#Create Log File with date and Time
$Logfile = "D:\PowerShell_Logs\LogFile $(get-date -Format MM-dd-yyyy_hh-mm-ss).log"

#Create function to log Comments into log file
Function LogComment
{
   Param ([string]$logstring)

   Add-content $Logfile -value $logstring
}

#Function Created by Steven Chin
Function CreateTermSetItems() {
    
	Param([string]$siteURL)

    if($siteURL -eq "") {
        throw "No Parameter 'Site URL' found. Please run .\TermStoreManagement 'Site URL'"
    }#end if

    $web = Get-SPWeb -Identity $siteURL

    $session = New-Object Microsoft.SharePoint.Taxonomy.TaxonomySession($siteURL)
    #$session = Get-SPTaxonomySession -Site $web.Site

    if ($session.TermStores.Count -lt 1) {
        throw "No term stores found. The Taxonomy Service is offline or missing"
    }#end if

    [Microsoft.SharePoint.Taxonomy.TermStore]$store = $session.TermStores[0]

    Write-Host `n
    Write-Host "Processing $web subsite"
    Write-Host "________________________________"

    #get the termset for the particular web
    [string]$providerName = [Microsoft.SharePoint.Publishing.Navigation.StandardNavigationProviderNames]::CurrentNavigationTaxonomyProvider
    [Microsoft.SharePoint.Publishing.Navigation.NavigationTermSet]$termSet = [Microsoft.SharePoint.Publishing.Navigation.TaxonomyNavigation]::GetTermSetForWeb($web, $providerName, $true)
    [Microsoft.SharePoint.Publishing.Navigation.NavigationTermSet]$editableTermSet = $termSet.GetAsEditable($session)

    #create the term for each page
    if($editableTermSet -ne $null) {
            
        foreach($page in $web.Lists["Pages"].Items) {
            
			#this section currently not working
			#working on a fix
            foreach($t in $termSet.Terms) {
                Write-Host $page.DisplayName - $t.Title.DefaultValue
                #if($page.DisplayName -eq $t.Title.DefaultValue) {continue}
            }#end if

			if($page.DisplayName -eq "default") { continue }
			if($page.DisplayName -eq "PageNotFoundError") { continue }

            [string]$linkType = [Microsoft.SharePoint.Publishing.Navigation.NavigationLinkType]::SimpleLink
            [Microsoft.SharePoint.Publishing.Navigation.NavigationTerm]$term = $editableTermSet.CreateTerm($page.Title, $linkType, [System.Guid]::NewGuid())
            $term.SimpleLinkUrl = [Microsoft.SharePoint.Publishing.PublishingPage]::GetPublishingPage($page).Uri.AbsoluteUri

            Write-Host $page.Title

        }#end foreach

        $editableTermSet.IsNavigationTermSet = $true
        $store.CommitAll()

    }#end if

    Write-Host `n
    Write-Host "Process Complete."
    Write-Host `n

}#end function

#Create funtion to check if Feature is Activated
function IsFeatureActivated
{
    param([string]$Id=$(throw "-Id parameter is required!"),
            [Microsoft.SharePoint.SPFeatureScope]$Scope=$(throw "-Scope parameter is required!"),
            [string]$Url)  
    if($Scope -ne "Farm" -and [string]::IsNullOrEmpty($Url))
    {
        throw "-Url parameter is required for scopes WebApplication,Site and Web"
    }
    $feature=$null

    switch($Scope)
    {
        "Farm" { $feature=Get-SPFeature $Id -Farm }
        "WebApplication" { $feature=Get-SPFeature $Id -WebApplication $Url }
        "Site" { $feature=Get-SPFeature $Id -Site $Url }
        "Web" { $feature=Get-SPFeature $Id -Web $Url }
    }
    #return if feature found or not (activated at scope) in the pipeline
    return ($feature -ne $null)
}

function SetAvailablePageLayout($webUrl, $pageLayoutName)
{  
    #web instance  
    $web = Get-SPweb $webURL  
      
    #Publishing web instance  
    $pweb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($web);  
     
    #getting current web available page layouts collection  
    $currentLayOuts = $pWeb.GetAvailablePageLayouts()  
     
    #site instance  
    $site = $web.Site  
    #publishing site instance  
    [Microsoft.Sharepoint.Publishing.PublishingSite]$publishingSite = New-Object Microsoft.SharePoint.Publishing.PublishingSite($site)  
     
    #getting collection of all the page layouts in a Site collection  
    $allPageLayouts = $publishingSite.PageLayouts  
 
    #looping through all the page layouts  
    foreach($pageLayout in $allPageLayouts)  
       {  
        
        #checking for the page layout which we want to make available
		  
		if($pageLayout.Name -eq $pageLayoutName)  
			{  
				#adding the new page layout to current webs available page layout collection  
                $currentLayOuts+=$pageLayout;  
				break;  
			}
		}  
     
    #updating available page layout collection of current site  
    $pweb.SetAvailablePageLayouts($currentLayOuts,$false)  
    $pWeb.Update()  
      
}

#Create Site Array
$RootSite = $env:PulseSite
$SiteArray = "/", "/hr/home", "/business-development/Marketing", "/business-development/occupational-medicine", "/business-development/wellness", 
"/corporate-finance/cash-accounting", "/corporate-finance/home",  "/corporate-finance/payroll", "/corporate-finance/purchasing", "/HCF/cbo", 
"/hcf/home", "/HCF/HPPS", "/IT/EMR", "/it/helpdesk", "/it/home", "/it/infosec", "/it/support-information", "/IT/Telecom", "/IT/Training", 
"/managed-care/home", "/Managed-Care/KCA", "/Operations/AHNT", "/operations/ASC", "/Operations/cancer-center", "/Operations/Cancer-Services",
"/operations/clinical-education", "/operations/contact-center", "/operations/home", "/operations/labcorp", "/operations/labcorponsite"


#Loop Trough Array
foreach ($Site in $SiteArray)
{
	$fullURL = [string]::Concat($RootSite,$Site)
	Write-Host -fore Yellow "Check to see if The Pulse Branding feature is active in $($fullURL)" 
	LogComment "Check to see if The Pulse Master Page feature is active in $($fullURL)"
	Try
	{
		if(IsFeatureActivated -Id 48b42642-6cf8-463c-bc34-c25cb3792842 -Scope "Site" -Url $fullURL)
		{
			Write-Host "This feature is already active"
			LogComment "This feature is already active"
		}
		else
		{
			#Pulse Master Page
			enable-spfeature -id 48b42642-6cf8-463c-bc34-c25cb3792842 -url $fullURL
			Write-Host "Activating Feature Pulse Master Page"
			LogComment "Activating Feature Pulse Master Page"
		}
	}
	Catch
	{
		Write-Host -fore Red $_.Exception.Message
		LogComment $_.Exception.Message
	}

	###### Activate Managed Navigation

	$newSite = get-spsite $fullURL

    LogComment "Site is " $newSite
 		
	foreach ($web in $newSite.AllWebs)
 	{
		$fullWebUrl = $web.Url

		Try
		{
			if(IsFeatureActivated -Id 2237983c-9758-4869-830b-ffb665631fba -Scope "Web" -Url $fullWebUrl)
			{
				Write-Host "This feature is already active"
				LogComment "This feature is already active"
			}
			else
			{
				#Pulse Master Page
				enable-spfeature -id 2237983c-9758-4869-830b-ffb665631fba -url $fullWebUrl
				Write-Host "Activating Pulse Term Store Management"
				LogComment "Activating Pulse Term Store Management"
                                CreateTermSetItems($fullWebUrl)

			}
		}
		Catch
		{
			Write-Host -fore Red $_.Exception.Message
			LogComment $_.Exception.Message
		}

		
		Write-Host -fore Yellow "Update the Sites Master Page to Internal Pulse Branding"
		LogComment "Update the Sites Master Page to Internal Pulse Branding"
		Try
		{
			if ($Site -eq "/")
			{
				$web.CustomMasterUrl = "/_catalogs/masterpage/RootSite.master"
			}
			else
			{
				$web.CustomMasterUrl = "$($Site)/_catalogs/masterpage/SubSite.master"
			}
 			
 			$web.update()
			Write-Host "The Master Page has been Updated"
			LogComment "The Master Page has been Updated"
		}
		Catch
		{
			Write-Host -fore Red $_.Exception.Message
			LogComment $_.Exception.Message
		}

		$web.AllowAllWebTemplates()
		$web.Update()

		Write-Host -fore Yellow "Adding New Pulse Layouts"
		LogComment "Adding New Pulse Layouts"

		Try
		{
			SetAvailablePageLayout -WebURL $fullWebUrl  -pageLayoutName "DetailPage.aspx"
			Write-Host "Pulse Detail Added"
			LogComment "Pulse Detail Added"
		}
		Catch
		{
			Write-Host -fore Red $_.Exception.Message
			LogComment $_.Exception.Message
		}

		Write-Host -fore Yellow "Setting Default Page Layout"
		LogComment "Setting Default Page Layout"

        Try
		{
			SetAvailablePageLayout -WebURL $fullWebUrl  -pageLayoutName "HomePage.aspx"
			Write-Host "Pulse Home Added"
			LogComment "Pulse Home Added"
		}
		Catch
		{
			Write-Host -fore Red $_.Exception.Message
			LogComment $_.Exception.Message
		}

		Write-Host -fore Yellow "Setting Default Page Layout"
		LogComment "Setting Default Page Layout"

		Try
		{
			Add-PSSnapin Microsoft.SharePoint.PowerShell -erroraction SilentlyContinue
			$pweb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($web)
			$layout = $pweb.GetAvailablePageLayouts() | ? {$_.Name -eq "DetailPage.aspx"}
			$pweb.SetDefaultPageLayout($layout, $false)
			$pweb.Update()
			Write-Host "Default Page layout has been set to PulseDetail"
			LogComment "Default Page layout has been set to PulseDetail"
		}
		Catch
		{
			Write-Host -fore Red $_.Exception.Message
			LogComment $_.Exception.Message
		}

		Write-Host -fore Yellow "Query Pages in $($fullWebUrl)"
		LogComment "Query Pages in $($fullWebUrl)"
		$query = New-Object Microsoft.SharePoint.SPQuery
		$query.ViewAttributes = "Scope='RecursiveAll'"
		$query.Query = "<Where><Eq><FieldRef Name='ContentType' /><Value Type='Computed'>Page</Value></Eq></Where>"
		$SourcePages = $pweb.GetPublishingPages($query)

		Write-Host -fore Yellow "loop through pages and update the content type"
		LogComment "loop through pages and update the content type"
		foreach ($SourcePage in $SourcePages)
		{
			Try
			{
				$SourcePageListItem = $SourcePage.ListItem
				Write-Host "Updating " $SourcePage.Name
				LogComment "Updating $($SourcePage.Name)" 
				#$SourcePageListItem.File.UndoCheckOut()
				$SourcePageListItem.File.CheckOut()
				$SourcePageListItem["PublishingPageLayout"] = "$($Site)/_catalogs/masterpage/DetailPage.aspx, Detail Page Layout"			
				$SourcePageListItem.Update();
				$SourcePageListItem.File.CheckIn("Updated PublishingPageLayout field")
				$SourcePageListItem.File.Approve("Updated PublishingPageLayout field")
			}
			Catch 
			{
				Write-Host -fore Red $_.Exception.Message
				LogComment $_.Exception.Message
			}
		}
 	}
	$newSite.Dispose()
}

$spWeb = Get-SPWeb("https://uat-pulse.kscpulse.com")
$spFile = $spWeb.GetFile("https://uat-pulse.kscpulse.com/Pages/default.aspx")
$spFile.CheckOut("Online",$null)
$spFile.Properties["PublishingPageLayout"] = "/_catalogs/masterpage/HomePage.aspx, Home Page Layout"
$spFile.Update()
$spFile.CheckIn("Update page layout via PowerShell",[Microsoft.SharePoint.SPCheckinType]::MajorCheckIn)
$spWeb.Dispose()
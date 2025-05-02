[void] [Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Publishing")

#Integration
[Environment]::SetEnvironmentVariable("KSCWebsite", "http://int.www.kelsey-seybold.com")
[Environment]::SetEnvironmentVariable("KSCNewWebsite", "http://int.beta.www.kelsey-seybold.com")

############################# BEGIN BIOS SCRIPT #####################################################################

$SourceSite = Get-SPSite $env:KSCWebsite
$SourceWeb = $SourceSite.OpenWeb("/Find-a-Doctor")

$DestinationSite = Get-SPSite $env:KSCNewWebsite
$DestinationWeb = $DestinationSite.OpenWeb("/Find-a-Houston-Doctor")

$SourcepubWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($SourceWeb);
$DestinationpubWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($DestinationWeb);

#Get pages in source web
$query = New-Object Microsoft.SharePoint.SPQuery
$query.ViewAttributes = "Scope='RecursiveAll'"
$query.Query = "<Where><Eq><FieldRef Name='ContentType' /><Value Type='Computed'>Physician Bio</Value></Eq></Where>"
$SourcePages = $SourcepubWeb.GetPublishingPages($query)

#loop through pages in source to create page in destination
foreach ($SourcePage in $SourcePages)
{
    if ($SourcePage.ContentType.Name -eq "Folder")
    {
        continue
    }

	$NewPage = $DestinationpubWeb.GetPublishingPages().Add($SourcePage.Name, $DestinationWeb.DefaultPageLayout)
    $NewPageListItem = $NewPage.ListItem
	$SourcePageListItem = $SourcePage.ListItem

    $NewPageListItem["PageTitle"] = $SourcePageListItem["PageTitle"]
    $NewPageListItem["MetaDescription"] = $SourcePageListItem["MetaDescription"]
    $NewPageListItem["PhysicianFirstName"] = $SourcePageListItem["PhysicianFirstName"]
    $NewPageListItem["PhysicianLastName"] = $SourcePageListItem["PhysicianLastName"]
    $NewPageListItem["PhysicianMiddleInitial"] = $SourcePageListItem["PhysicianMiddleInitial"]
    $NewPageListItem["YearHired"] = $SourcePageListItem["YearHired"]
    $NewPageListItem["SpecialInterest"] = $SourcePageListItem["SpecialInterest"]
    $NewPageListItem["AcademicAppointments"] = $SourcePageListItem["AcademicAppointments"]
    $NewPageListItem["MedicalSchool"] = $SourcePageListItem["MedicalSchool"]
    $NewPageListItem["InternshipResidency"] = $SourcePageListItem["InternshipResidency"]
    $NewPageListItem["Sex"] = $SourcePageListItem["Sex"]
    $NewPageListItem["HospitalAffiliations"] = $SourcePageListItem["HospitalAffiliations"]
    $NewPageListItem["TaxonomyFieldTypeMultiTaxHTField0"] = $SourcePageListItem["TaxonomyFieldTypeMultiTaxHTField0"]
    $NewPageListItem["PrimaryHospitalAffiliations"] = $SourcePageListItem["PrimaryHospitalAffiliations"]
    $NewPageListItem["PrimaryHospitalAffiliationsTaxHTField0"] = $SourcePageListItem["PrimaryHospitalAffiliationsTaxHTField0"]
    $NewPageListItem["PhysicianMedicalSpecialties"] = $SourcePageListItem["PhysicianMedicalSpecialties"]
    $NewPageListItem["PhysicianMedicalSpecialtiesTaxHTField0"] = $SourcePageListItem["PhysicianMedicalSpecialtiesTaxHTField0"]
    $NewPageListItem["PhysicianLocations"] = $SourcePageListItem["PhysicianLocations"]
    $NewPageListItem["PhysicianLocationsTaxHTField0"] = $SourcePageListItem["PhysicianLocationsTaxHTField0"]
    $NewPageListItem["PhysicianLanguages"] = $SourcePageListItem["PhysicianLanguages"]
    $NewPageListItem["PhysicianLanguagesTaxHTField0"] = $SourcePageListItem["PhysicianLanguagesTaxHTField0"]
    $NewPageListItem["PhysicianSpecialtyBoards"] = $SourcePageListItem["PhysicianSpecialtyBoards"]
    $NewPageListItem["PhysicianSpecialtyBoardsTaxHTField0"] = $SourcePageListItem["PhysicianSpecialtyBoardsTaxHTField0"]
    $NewPageListItem["DepartmentChief"] = $SourcePageListItem["DepartmentChief"]
    $NewPageListItem["DepartmentChiefTaxHTField0"] = $SourcePageListItem["DepartmentChiefTaxHTField0"]
    $NewPageListItem["PhysiciansFaxNumbers"] = $SourcePageListItem["PhysiciansFaxNumbers"]
    $NewPageListItem["AppointmentsNowEnabled"] = $SourcePageListItem["AppointmentsNowEnabled"]
	$NewPageListItem["YoutubeURL"] = $SourcePageListItem["YoutubeURL"]
	
	Write-Host "Migrating " $SourcePage.Name
    $NewPageListItem.Update();

    $NewPageListItem.File.CheckIn("Initial Data Migration - Checked in via script")
    $NewPageListItem.File.Approve("Initial Data Migration - Approved via script")
}
$DestinationWeb.Dispose()
$DestinationSite.Dispose()
$SourceWeb.Dispose()
$SourceSite.Dispose()

############################# END BIOS SCRIPT #####################################################################

############################# BEGIN LOCATIONS SCRIPT #####################################################################

$SourceSite = Get-SPSite $env:KSCWebsite
$SourceWeb = $SourceSite.OpenWeb("/Locations")

$DestinationSite = Get-SPSite $env:KSCNewWebsite
$DestinationWeb = $DestinationSite.OpenWeb("/Find-a-Location")

$SourcepubWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($SourceWeb);
$DestinationpubWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($DestinationWeb);

#Get pages in source web
$query = New-Object Microsoft.SharePoint.SPQuery
$query.ViewAttributes = "Scope='RecursiveAll'"
$query.Query = "<Where><Eq><FieldRef Name='ContentType' /><Value Type='Computed'>Locations</Value></Eq></Where>"
$SourcePages = $SourcepubWeb.GetPublishingPages($query)

#loop through pages in source to create page in destination
foreach ($SourcePage in $SourcePages)
{
    if ($SourcePage.ContentType.Name -eq "Folder")
    {
        continue
    }

	$NewPage = $DestinationpubWeb.GetPublishingPages().Add($SourcePage.Name, $DestinationWeb.DefaultPageLayout)
    $NewPageListItem = $NewPage.ListItem
	$SourcePageListItem = $SourcePage.ListItem

    $NewPageListItem["PageTitle"] = $SourcePageListItem["PageTitle"]
    $NewPageListItem["MetaDescription"] = $SourcePageListItem["MetaDescription"]
    $NewPageListItem["Address"] = $SourcePageListItem["Address"]
    $NewPageListItem["City"] = $SourcePageListItem["City"]
    $NewPageListItem["State"] = $SourcePageListItem["State"]
    $NewPageListItem["Zip"] = $SourcePageListItem["Zip"]
    $NewPageListItem["Phone"] = $SourcePageListItem["Phone"]
    $NewPageListItem["Fax"] = $SourcePageListItem["Fax"]
    $NewPageListItem["Latitude"] = $SourcePageListItem["Latitude"]
    $NewPageListItem["Longitude"] = $SourcePageListItem["Longitude"]
    $NewPageListItem["ClinicName"] = $SourcePageListItem["ClinicName"]
    $NewPageListItem["ClinicNameTaxHTField0"] = $SourcePageListItem["ClinicNameTaxHTField0"]
    $NewPageListItem["HealthCareServices"] = $SourcePageListItem["HealthCareServices"]
    $NewPageListItem["HealthCareServicesTaxHTField0"] = $SourcePageListItem["HealthCareServicesTaxHTField0"]
    $NewPageListItem["LocationId"] = $SourcePageListItem["LocationId"]
    $NewPageListItem["EpicLocationName"] = $SourcePageListItem["EpicLocationName"]
	$NewPageListItem["LocationImage"] = $SourcePageListItem["LocationImage"]
    $NewPageListItem["AppointmentsNowLocationName"] = $SourcePageListItem["AppointmentsNowLocationName"]
	
	Write-Host "Migrating " $SourcePage.Name
    $NewPageListItem.Update();

    $NewPageListItem.File.CheckIn("Initial Data Migration - Checked in via script")
    $NewPageListItem.File.Approve("Initial Data Migration - Approved via script")
}
$DestinationWeb.Dispose()
$DestinationSite.Dispose()
$SourceWeb.Dispose()
$SourceSite.Dispose()


############################# END LOCATIONS SCRIPT #####################################################################

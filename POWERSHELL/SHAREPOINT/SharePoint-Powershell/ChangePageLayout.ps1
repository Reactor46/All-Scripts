
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
 
#Variables for Web and Page URLs
$WebURL="https://pulse.kscpulse.com/managed-care/home/"
$PageURL="https://pulse.kscpulse.com/managed-care/home/Pages/Tools.aspx"
$PageLayout="https://pulse.kscpulse.com/managed-care/home/_catalogs/masterpage/DetailPage.aspx"
 
#Get the web and page
$Web = Get-SPWeb $WebURL
$File = $Web.GetFile($PageURL)
 
#change page layout sharepoint 2013 powershell
$File.CheckOut("Online",$null)
$File.Properties["PublishingPageLayout"] = $PageLayout
$File.Update()
$File.CheckIn("Page layout updated via PowerShell",[Microsoft.SharePoint.SPCheckinType]::MajorCheckIn)
 
$Web.Dispose()


#Read more: https://www.sharepointdiary.com/2014/10/change-page-layout-in-sharepoint-2013-with-powershell.html#ixzz81KIsYX00
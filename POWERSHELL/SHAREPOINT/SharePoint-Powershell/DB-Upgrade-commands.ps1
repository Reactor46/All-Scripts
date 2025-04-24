(Get-SPDatabase | ?{$_.type -eq "App Management Database"}).Provision()                                                             
(Get-SPDatabase | ?{$_.type -eq "Configuration Database"}).Provision()                                                              
(Get-SPDatabase | ?{$_.type -eq "Content Database"}).Provision()                                                                    
(Get-SPDatabase | ?{$_.type -eq "Microsoft SharePoint Foundation Subscription Settings Database"}).Provision()                      
(Get-SPDatabase | ?{$_.type -eq "Microsoft.Office.SecureStoreService.Server.SecureStoreServiceDatabase"}).Provision()               
(Get-SPDatabase | ?{$_.type -eq "Microsoft.Office.Server.Administration.StateDatabase"}).Provision()                                
(Get-SPDatabase | ?{$_.type -eq "Microsoft.Office.Server.Search.Administration.SearchAdminDatabase"}).Provision()                   
(Get-SPDatabase | ?{$_.type -eq "Microsoft.Office.Server.Search.Administration.SearchAnalyticsReportingDatabase"}).Provision()      
(Get-SPDatabase | ?{$_.type -eq "Microsoft.Office.Server.Search.Administration.SearchGathererDatabase"}).Provision()                
(Get-SPDatabase | ?{$_.type -eq "Microsoft.Office.Server.Search.Administration.SearchLinksDatabase"}).Provision()                   
(Get-SPDatabase | ?{$_.type -eq "Microsoft.Office.TranslationServices.QueueDatabase"}).Provision()                                  
(Get-SPDatabase | ?{$_.type -eq "Microsoft.Office.Word.Server.Service.QueueDatabase"}).Provision()                                  
(Get-SPDatabase | ?{$_.type -eq "Microsoft.SharePoint.Administration.SPUsageDatabase"}).Provision()                                 
(Get-SPDatabase | ?{$_.type -eq "Microsoft.SharePoint.BusinessData.SharedService.BdcServiceDatabase"}).Provision()                  
(Get-SPDatabase | ?{$_.type -eq "Microsoft.SharePoint.Taxonomy.MetadataWebServiceDatabase"}).Provision()      
(Get-SPDatabase | ?{$_.type -eq "Microsoft.SharePoint.BusinessData.SharedService.BdcServiceDatabase"}).Provision()     


Get-SPContentDatabase -Identity SP2013_INT_KSC_Intranet_Depts_Marketing | Upgrade-SPContentDatabase

Get-SPContentDatabase | ?{$_.NeedsUpgrade -eq $true} | Upgrade-SPContentDatabase
# Get-SPDatabase | ?{$_.NeedsUpgrade -eq $true} | Upgrade-SPDatabase
           

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
 
#Set this variable value accordingly
$DatabaseName = "content_database_name_here"
 
#Get the Database ID
$DatabaseID = (Get-SPContentDatabase -identity $DatabaseName).ID
 
#Update Content Database
Upgrade-SPContentDatabase -id $DatabaseID -Confirm:$false


#Read more: https://www.sharepointdiary.com/2014/05/database-is-up-to-date-but-some-sites-are-not-completely-upgraded.html#ixzz7pXQZ3YnV


$WebAppURL= "https://intranet.crescent.com"
 
#Get all content databases of the particular web application
$ContentDBColl = (Get-SPWebApplication -Identity $WebAppURL).ContentDatabases
 
foreach ($contentDB in $ContentDBColl)
{
   #Updade each content database
   Upgrade-SPContentDatabase -id $contentDB.Id -Confirm:$false
}


#Read more: https://www.sharepointdiary.com/2014/05/database-is-up-to-date-but-some-sites-are-not-completely-upgraded.html#ixzz7pXQgqChz
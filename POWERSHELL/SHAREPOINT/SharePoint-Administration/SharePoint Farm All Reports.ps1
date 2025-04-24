#######################################################
#Add Add-PSSnapin Microsoft.SharePoint.PowerShell
Set-ExecutionPolicy "Unrestricted"
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
#######################################################
function SPWebAppScanReport()
  { 

		Try
		{
			Write-Host "SharePoint Web Application Scan Report" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"
            $ex = "Exceed the Supported Limit"
            $wa = "Within the Limit"
            #SharePoint Web Application
            $SPWebApp = Get-SPWebApplication 
            $SPWebAppcount = $SPWebApp.count
            #SharePoint Web Application Pool
            $SPWebAppPool = Get-SPWebApplication | select ApplicationPool -Unique
            #SharePoint Service Application Pool
            $SPSrvWebAppPool = Get-SPServiceApplicationPool  | Select -Unique
            $SPAppPoolCount = $SPSrvWebAppPool.count + $SPWebAppPool.count
            #SharePoint Service Applications
            $SPSrvApp = Get-SPServiceApplication | select id,name
            $SPSrvAppcount= $SPSrvApp.count
            
            
            ##################################################################################
            #How Many Web Applications in farm 
            Write-Host "SharePoint Web Applications per Farm" -ForegroundColor cyan
            Write-Host "The supported limit for the SharePoint web application per farm is 20 web applications."
            switch($SPWebAppcount)
            {
            {$_ -ge 20} {Write-Host "Total Number of SharePoint Web Application:" $_ "|"$ex -ForegroundColor red }
            {$_ -lt 20} {Write-Host "Total Number of SharePoint Web Application:" $_ "|"$wa -ForegroundColor yellow}
            }
            $SPWebApp
            Write-Host "For more details, please check https://spgeeks.devoworx.com/sharepoint-2019-limitations" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"


            ##################################################################################
            #How Many SharePoint Web Application Pools for the web server Per farm 
            Write-Host "SharePoint Web Application Pools for the web server Per farm" -ForegroundColor cyan
            Write-Host "SharePoint Application Pool Limits for Web Server per Farm is 10 Application Pools (This limit depends mainly on the Server Hardware specifications)."
            switch($SPAppPoolCount)
            {
            {$_ -ge 10} {Write-Host "Total Number of SharePoint Application Pool:" $_ "|"$ex -ForegroundColor red }
            {$_ -lt 10} {Write-Host "Total Number of SharePoint Application Pool:" $_ "|"$wa -ForegroundColor yellow}
            }
            $SPWebAppPool
            $SPSrvWebAppPool | select name
            Write-Host "For more details, please check https://spgeeks.devoworx.com/sharepoint-2019-service-accounts-best-practice/" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"

            
            ##################################################################################
            #How Many SharePoint Service Applications Running on farm
            Write-Host "Running SharePoint Service Applications" -ForegroundColor cyan 
            Write-Host "Total Number of Running SharePoint Service Applications:" $SPSrvApp.count -ForegroundColor Green
            $SPSrvApp 
            Write-Host "--------------------------------------------------------------------"


            ##################################################################################
            #Web Application Report Summary 
            Write-Host "Web Application Report Summary" -ForegroundColor Green
            Write-Host "Total Number of SharePoint Web Application Per Farm:" $SPWebAppcount
            Write-Host "Total Number of SharePoint Application Pool Per Farm:" $SPAppPoolCount
            Write-Host "Total Number of SharePoint Service Application Per Farm:" $SPSrvAppcount
            Write-Host "Check the details at: https://spgeeks.devoworx.com/get-all-web-applications-per-farm/" -ForegroundColor cyan
            Write-Host "Check also the SharePoint Farm Scan Report at https://spgeeks.devoworx.com/sharepoint-farm-scan-report/" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"
		}
		Catch
		{
			Write-Host $_.Exception.Message -ForegroundColor Red
		}
  }


function SPSiteScanReport()
  { 

		Try
		{
			Write-Host "SharePoint Site Collection Scan Report" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"
            $ex = "Exceed the Supported Limit"
            $wa = "Within the Limit"
            #SharePoint Web Application
            $SPWebApp = Get-SPWebApplication 
            $SPWebAppcount = $SPWebApp.count
            # Site Collections
            $SiteCollections = Get-SPSite | select url,contentdatabase,webapplication,@{Name="Site Collection Size (MB)";Expression={[math]::Round($_.usage.storage/(1MB),2)}},@{Name="Within the Limit"; Expression={if([math]::Round($_.usage.storage/1MB,2) -ge 100*1024){ "No"} else {"Yes"}}} | Format-List
            $SiteCollectionscount = (Get-SPSite).count
            # SubSites
            $Subsites = Get-SPSite | Get-SPWeb -Limit All
            $Subsitescount= $Subsites.count
            
            
            ##################################################################################
            #How Many SharePoint Site Collections Per farm 
            Write-Host "SharePoint Site Collections Per farm" -ForegroundColor cyan
            Write-Host "The supported limit for SharePoint Site Collections per farm is 750,000 site collections."
            switch($SiteCollectionscount)
            {
            {$_ -ge 750000} {Write-Host "Total Number of SharePoint Site Collections:" $_ "|"$ex -ForegroundColor red }
            {$_ -lt 750000} {Write-Host "Total Number of SharePoint Site Collections:" $_ "|"$wa -ForegroundColor yellow}
            }
            $SiteCollections
            Write-Host "For more details, please check https://spgeeks.devoworx.com/sharepoint-2019-limitations" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"


            ##################################################################################
            #How Many SharePoint Site Collection Per Web Application
            Write-Host "SharePoint Site Collection Per Web Application" -ForegroundColor cyan
            foreach($WebApp in $SPWebApp){
            Write-Host "The Total Number of site collection per Web Application" $WebApp.Url "is" (Get-SPSite -WebApplication $WebApp).count -ForegroundColor green
            Get-SPSite -WebApplication $WebApp | select url,contentdatabase,@{Name="Site Collection Size (MB)";Expression={[math]::Round($_.usage.storage/(1MB),2)}},@{Name="Within the Limit"; Expression={if([math]::Round($_.usage.storage/1MB,2) -ge 100*1024){ "No"} else {"Yes"}}} | format-list
            Write-Host "---------------------------------"
            }

            ##################################################################################
            #How Many SharePoint Site Collection Per Content Database
            Write-Host "SharePoint Site Collection Per Content Database" -ForegroundColor cyan
            foreach($cDB in Get-SPContentDatabase){
            Write-Host "The Total Number of site collection per content database" $cDB.name "is" (Get-SPSite -ContentDatabase $cDB).count -ForegroundColor green
            Get-SPSite -ContentDatabase $cDB | select url,@{Name="Site Collection Size (MB)";Expression={[math]::Round($_.usage.storage/(1MB),2)}},@{Name="Within the Limit"; Expression={if([math]::Round($_.usage.storage/1MB,2) -ge 100*1024){ "No"} else {"Yes"}}} | format-list | format-list
            Write-Host "---------------------------------"
            }


            ##################################################################################
            #How Many SharePoint SubSites Per farm 
            Write-Host "SharePoint SubSites Per farm" -ForegroundColor cyan
            Write-Host "The supported limit for subsites per farm is 250,000 subsites"
            switch($Subsitescount)
            {
            {$_ -ge 250000} {Write-Host "Total Number of SharePoint SubSites:" $_ "|"$ex -ForegroundColor red }
            {$_ -lt 250000} {Write-Host "Total Number of SharePoint SubSites:" $_ "|"$wa -ForegroundColor yellow}
            }
            $Subsites
            Write-Host "For more details, please check https://spgeeks.devoworx.com/sharepoint-2019-limitations" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"

            ##################################################################################
            #How Many SharePoint SubSites per Site Collection
            Write-Host "SharePoint SubSites per Site Collection" -ForegroundColor cyan
            foreach($SC in Get-SPSite){
            Write-Host "The Total Number of SubSites per Site Collection" $SC.Url "is" (Get-SPWeb -Site $SC -Limit All).count -ForegroundColor green
            Get-SPWeb -Site $SC -Limit All | select url | format-list
            Write-Host "---------------------------------"
            }

            ##################################################################################
            #Site Collection Scan Report Summary 
            Write-Host "Site Collection Scan Report Summary" -ForegroundColor Green
            Write-Host "Total Number of SharePoint Web Application Per Farm:" $SPWebAppcount
            Write-Host "Total Number of SharePoint Site Collection Per Farm:" $SiteCollectionscount
            Write-Host "Total Number of SharePoint SubSites Per Farm:" $Subsitescount
            Write-Host "Check the details at: https://spgeeks.devoworx.com/all-site-collections-and-subsites-per-farm" -ForegroundColor cyan
            Write-Host "For more details, please check https://spgeeks.devoworx.com/sharepoint-farm-scan-report/" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"
		}
		Catch
		{
			Write-Host $_.Exception.Message -ForegroundColor Red
		}
  }



function SPFramScanReport()
  { 

		Try
		{
			Write-Host "SharePoint Farm Scan Report" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"
            $ex = "Exceed the Supported Limit"
            $wa = "Within the Limit"
            #SharePoint Web Application
            $SPWebApp = Get-SPWebApplication 
            $SPWebAppcount = $SPWebApp.count
            #SharePoint Web Application Pool
            $SPWebAppPool = Get-SPWebApplication | select ApplicationPool -Unique
            #SharePoint Service Application Pool
            $SPSrvWebAppPool = Get-SPServiceApplicationPool  | Select -Unique
            $SPAppPoolCount = $SPSrvWebAppPool.count + $SPWebAppPool.count
            #SharePoint Service Applications
            $SPSrvApp = Get-SPServiceApplication | select id,name
            $SPSrvAppcount= $SPSrvApp.count
            # Content Databases
            $ContentDB = Get-SPContentDatabase | select name,WebApplication,@{Name="Contenet Database Size (GB)"; Expression={[math]::Round($_.disksizerequired/1024MB,2)}},@{Name="Within the Limit"; Expression={if([math]::Round($_.disksizerequired/1024MB,2) -ge 200){ "No"} else {"Yes"}}}
            $ContentDBcount = $ContentDB.count
            # Site Collections
            $SiteCollections = Get-SPSite -Limit All | select url,contentdatabase,webapplication,@{Name="Site Collection Size (MB)";Expression={[math]::Round($_.usage.storage/(1MB),2)}},@{Name="Within the Limit"; Expression={if([math]::Round($_.usage.storage/1MB,2) -ge 100*1024){ "No"} else {"Yes"}}} | Format-List
            $SiteCollectionscount = (Get-SPSite).count
            # SubSites
            $Subsites = Get-SPSite -Limit All | Get-SPWeb -Limit All
            $Subsitescount= $Subsites.count
            
            ##################################################################################
            #How Many Web Applications in farm 
            Write-Host "SharePoint Web Applications per Farm" -ForegroundColor cyan
            Write-Host "The supported limit for the SharePoint web application per farm is 20 web applications."
            switch($SPWebAppcount)
            {
            {$_ -ge 20} {Write-Host "Total Number of SharePoint Web Application:" $_ "|"$ex -ForegroundColor red }
            {$_ -lt 20} {Write-Host "Total Number of SharePoint Web Application:" $_ "|"$wa -ForegroundColor yellow}
            }
            $SPWebApp
            Write-Host "For more details, please check https://spgeeks.devoworx.com/sharepoint-2019-limitations" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"


            ##################################################################################
            #How Many SharePoint Web Application Pools for the web server Per farm 
            Write-Host "SharePoint Web Application Pools for the web server Per farm" -ForegroundColor cyan
            Write-Host "SharePoint Application Pool Limits for Web Server per Farm is 10 Application Pools (This limit depends mainly on the Server Hardware specifications)."
            switch($SPAppPoolCount)
            {
            {$_ -ge 10} {Write-Host "Total Number of SharePoint Application Pool:" $_ "|"$ex -ForegroundColor red }
            {$_ -lt 10} {Write-Host "Total Number of SharePoint Application Pool:" $_ "|"$wa -ForegroundColor yellow}
            }
            $SPWebAppPool
            $SPSrvWebAppPool | select name
            Write-Host "For more details, please check https://spgeeks.devoworx.com/sharepoint-2019-service-accounts-best-practice/" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"

            
            ##################################################################################
            #How Many SharePoint Service Applications Running on farm
            Write-Host "Running SharePoint Service Applications" -ForegroundColor cyan 
            Write-Host "Total Number of Running SharePoint Service Applications:" $SPSrvApp.count -ForegroundColor Green
            $SPSrvApp 
            Write-Host "--------------------------------------------------------------------"

            
            ##################################################################################
            #How Many SharePoint Content Database Per farm 
            Write-Host "SharePoint Content Database Per farm" -ForegroundColor cyan
            Write-Host "The supported limit for SharePoint Content Databases per farm is 500 content databases."
            switch($ContentDBcount)
            {
            {$_ -ge 500} {Write-Host "Total Number of SharePoint Content Database:" $_ "|"$ex -ForegroundColor red }
            {$_ -lt 500} {Write-Host "Total Number of SharePoint Content Database:" $_ "|"$wa -ForegroundColor yellow}
            }
            $ContentDB
            Write-Host "For more details, please check https://spgeeks.devoworx.com/sql-server-best-practices-sharepoint/" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"


            ##################################################################################
            #How Many SharePoint Content Database Per Web Application
            Write-Host "SharePoint Content Database Per Web Application" -ForegroundColor cyan
            foreach($WebApp in $SPWebApp){
            Write-Host "The Total Number of Content Database per Web Application" $WebApp.Url "is" (Get-SPContentDatabase -WebApplication $WebApp).count -ForegroundColor green
            Get-SPContentDatabase -WebApplication $WebApp | select name,@{Name="Contenet Database Size (GB)"; Expression={[math]::Round($_.disksizerequired/1024MB,2)}},@{Name="Within the Limit"; Expression={if([math]::Round($_.disksizerequired/1024MB,2) -ge 200){ "No"} else {"Yes"}}} | format-list
            Write-Host "---------------------------------"
            }

            ##################################################################################
            #How Many SharePoint Site Collections Per farm 
            Write-Host "SharePoint Site Collections Per farm" -ForegroundColor cyan
            Write-Host "The supported limit for SharePoint Site Collections per farm is 750,000 site collections."
            switch($SiteCollectionscount)
            {
            {$_ -ge 750000} {Write-Host "Total Number of SharePoint Site Collections:" $_ "|"$ex -ForegroundColor red }
            {$_ -lt 750000} {Write-Host "Total Number of SharePoint Site Collections:" $_ "|"$wa -ForegroundColor yellow}
            }
            $SiteCollections
            Write-Host "For more details, please check https://spgeeks.devoworx.com/sharepoint-2019-limitations" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"


            ##################################################################################
            #How Many SharePoint Site Collection Per Web Application
            Write-Host "SharePoint Site Collection Per Web Application" -ForegroundColor cyan
            foreach($WebApp in $SPWebApp){
            Write-Host "The Total Number of site collection per Web Application" $WebApp.Url "is" (Get-SPSite -WebApplication $WebApp).count -ForegroundColor green
            Get-SPSite -WebApplication $WebApp | select url,contentdatabase,@{Name="Site Collection Size (MB)";Expression={[math]::Round($_.usage.storage/(1MB),2)}},@{Name="Within the Limit"; Expression={if([math]::Round($_.usage.storage/1MB,2) -ge 100*1024){ "No"} else {"Yes"}}} | format-list
            Write-Host "---------------------------------"
            }

            ##################################################################################
            #How Many SharePoint Site Collection Per Content Database
            Write-Host "SharePoint Site Collection Per Content Database" -ForegroundColor cyan
            foreach($cDB in Get-SPContentDatabase){
            Write-Host "The Total Number of site collection per content database" $cDB.name "is" (Get-SPSite -ContentDatabase $cDB).count -ForegroundColor green
            Get-SPSite -ContentDatabase $cDB | select url,@{Name="Site Collection Size (MB)";Expression={[math]::Round($_.usage.storage/(1MB),2)}},@{Name="Within the Limit"; Expression={if([math]::Round($_.usage.storage/1MB,2) -ge 100*1024){ "No"} else {"Yes"}}} | format-list | format-list
            Write-Host "---------------------------------"
            }


            ##################################################################################
            #How Many SharePoint SubSites Per farm 
            Write-Host "SharePoint SubSites Per farm" -ForegroundColor cyan
            Write-Host "The supported limit for subsites per farm is 250,000 subsites"
            switch($Subsitescount)
            {
            {$_ -ge 250000} {Write-Host "Total Number of SharePoint SubSites:" $_ "|"$ex -ForegroundColor red }
            {$_ -lt 250000} {Write-Host "Total Number of SharePoint SubSites:" $_ "|"$wa -ForegroundColor yellow}
            }
            $Subsites
            Write-Host "For more details, please check https://spgeeks.devoworx.com/sharepoint-2019-limitations" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"

            ##################################################################################
            #How Many SharePoint SubSites per Site Collection
            Write-Host "SharePoint SubSites per Site Collection" -ForegroundColor cyan
            foreach($SC in Get-SPSite){
            Write-Host "The Total Number of SubSites per Site Collection" $SC.Url "is" (Get-SPWeb -Site $SC -Limit All).count -ForegroundColor green
            Get-SPWeb -Site $SC -Limit All | select url | format-list
            Write-Host "---------------------------------"
            }

            ##################################################################################
            #Farm Summary 
            Write-Host "Farm Report Summary" -ForegroundColor Green
            Write-Host "Total Number of SharePoint Web Application Per Farm:" $SPWebAppcount
            Write-Host "Total Number of SharePoint Application Pool Per Farm:" $SPAppPoolCount
            Write-Host "Total Number of SharePoint Service Application Per Farm:" $SPSrvAppcount
            Write-Host "Total Number of SharePoint Content Database Per Farm:" $ContentDBcount
            Write-Host "Total Number of SharePoint Site Collection Per Farm:" $SiteCollectionscount
            Write-Host "Total Number of SharePoint SubSites Per Farm:" $Subsitescount
            Write-Host "For more details, please check https://spgeeks.devoworx.com/sharepoint-farm-scan-report/" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"
		}
		Catch
		{
			Write-Host $_.Exception.Message -ForegroundColor Red
		}
  }



function SPCDBScanReport()
  { 

		Try
		{
			Write-Host "SharePoint Content Database Scan Report" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"
            $ex = "Exceed the Supported Limit"
            $wa = "Within the Limit"
            #SharePoint Web Application
            $SPWebApp = Get-SPWebApplication 
            # Content Databases
            $ContentDB = Get-SPContentDatabase | select name,WebApplication,@{Name="Contenet Database Size (GB)"; Expression={[math]::Round($_.disksizerequired/1024MB,2)}},@{Name="Within the Limit"; Expression={if([math]::Round($_.disksizerequired/1024MB,2) -ge 200){ "No"} else {"Yes"}}}
            $ContentDBcount = $ContentDB.count
            
            
            ##################################################################################
            #How Many SharePoint Content Database Per farm 
            Write-Host "SharePoint Content Database Per farm" -ForegroundColor cyan
            Write-Host "The supported limit for SharePoint Content Databases per farm is 500 content databases."
            switch($ContentDBcount)
            {
            {$_ -ge 500} {Write-Host "Total Number of SharePoint Content Database:" $_ "|"$ex -ForegroundColor red }
            {$_ -lt 500} {Write-Host "Total Number of SharePoint Content Database:" $_ "|"$wa -ForegroundColor yellow}
            }
            $ContentDB  | Format-List
            Write-Host "For more details, please check https://spgeeks.devoworx.com/sql-server-best-practices-sharepoint/" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"


            ##################################################################################
            #How Many SharePoint Content Database Per Web Application
            Write-Host "SharePoint Content Database Per Web Application" -ForegroundColor cyan
            foreach($WebApp in $SPWebApp){
            Write-Host "The Total Number of Content Database per Web Application" $WebApp.Url "is" (Get-SPContentDatabase -WebApplication $WebApp).count -ForegroundColor green
            Get-SPContentDatabase -WebApplication $WebApp | select name,@{Name="Contenet Database Size (GB)"; Expression={[math]::Round($_.disksizerequired/1024MB,2)}},@{Name="Within the Limit"; Expression={if([math]::Round($_.disksizerequired/1024MB,2) -ge 200){ "No"} else {"Yes"}}} | format-list
            Write-Host "---------------------------------"
            }         
           

            ##################################################################################
            #Content Database Scan Report Summary 
            Write-Host "Content Database Scan Report Summary" -ForegroundColor Green
            Write-Host "Total Number of SharePoint Content Database Per Farm:" $ContentDBcount
            Write-Host "Check the details at: https://spgeeks.devoworx.com/get-all-content-databases-per-farm/" -ForegroundColor cyan
            Write-Host "Check also the SharePoint Farm Scan Report at https://spgeeks.devoworx.com/sharepoint-farm-scan-report/" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"
		}
		Catch
		{
			Write-Host $_.Exception.Message -ForegroundColor Red
		}
  }



function Get-SharePointEdition()
 {
   Write-Host "-----------------------------------------------------------------------" -ForegroundColor yellow
   Write-Host "The Installed SharePoint" -ForegroundColor yellow
   Write-Host "-----------------------------------------------------------------------" -ForegroundColor yellow
   $SharePointEditionGuid = (Get-SPFarm).Products 
   switch ($SharePointEditionGuid) 
      { 
             #SharePoint 2016 Editions
             5DB351B8-C548-4C3C-BFD1-82308C9A519B {"SharePoint 2016 Trail."}
             4F593424-7178-467A-B612-D02D85C56940 {"SharePoint 2016 Standard."} 
             716578D2-2029-4FF2-8053-637391A7E683 {"SharePoint 2016 Enterprise."} 
             #SharePoint 2013 Editions
             9FF54EBC-8C12-47D7-854F-3865D4BE8118 {"SharePoint Foundation 2013."} 
             35466B1A-B17B-4DFB-A703-F74E2A1F5F5E {"SharePoint Server 2013 Enterprise plus Project Server 2013.";break}
			 BC7BAF08-4D97-462C-8411-341052402E71 {"SharePoint Server 2013 Enterprise plus Project Server 2013 Trail.";break}
             B7D84C2B-0754-49E4-B7BE-7EE321DCE0A9 {"SharePoint Server 2013 Enterprise."} 
			 298A586A-E3C1-42F0-AFE0-4BCFDC2E7CD0 {"SharePoint Server 2013 Enterprise Trail."} 
             C5D855EE-F32B-4A1C-97A8-F0A28CE02F9C {"SharePoint Server 2013."}
			 CBF97833-C73A-4BAF-9ED3-D47B3CFF51BE {"SharePoint Server 2013 Trail."}
             #SharePoint 2010 Editions
             BEED1F75-C398-4447-AEF1-E66E1F0DF91E {"SharePoint Foundation 2010."} 
             B2C0B444-3914-4ACB-A0B8-7CF50A8F7AA0 {"SharePoint Server 2010 Standard Trial."}
             3FDFBCC8-B3E4-4482-91FA-122C6432805C {"SharePoint Server 2010 Standard."} 
             88BED06D-8C6B-4E62-AB01-546D6005FE97 {"SharePoint Server 2010 Enterprise Trial."} 
             D5595F62-449B-4061-B0B2-0CBAD410BB51 {"SharePoint Server 2010 Enterprise."} 
             84902853-59F6-4B20-BC7C-DE4F419FEFAD {"Project Server 2010 Trial."} 
             ED21638F-97FF-4A65-AD9B-6889B93065E2 {"Project Server 2010."} 
             default {"The SharePoint edition can't be determined"}
      }
       
        Write-Host "-----------------------------------------------------------------------" -ForegroundColor yellow
        Write-Host "The Biuld Version" -ForegroundColor yellow
        Write-Host "-----------------------------------------------------------------------" -ForegroundColor yellow
        Write-Host  (Get-SPFarm).buildversion
      
 }

#Get SharePoint Edition 
Get-SharePointEdition
#Run SharePoint Web APplication Scan Report
SPWebAppScanReport
#Run SharePoint Content Database Scan Report
SPCDBScanReport
#Run the SharePoint Farm Scan Report
SPFramScanReport
#Run SharePoint Site Collection Scan Report
SPSiteScanReport
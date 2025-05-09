
# Get Size of all Sub-sites in a Site Collection
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") > $null

#Region MOSS2007-CmdLets

Function Get-SPWebApplication()
{   
  Param( [Parameter(Mandatory=$true)] [string]$WebAppURL )
  return [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup($WebAppURL)
}

Function global:Get-SPSite()
{
  Param( [Parameter(Mandatory=$true)] [string]$SiteCollURL )

   if($SiteCollURL -ne '')
    {
  		return new-Object Microsoft.SharePoint.SPSite($SiteCollURL)
  	}
}
 
Function global:Get-SPWeb()
{
 Param( [Parameter(Mandatory=$true)] [string]$SiteURL )
  $site = Get-SPSite($SiteURL)
        if($site -ne $null)
            {
               $web=$site.OpenWeb();
            }
    return $web
}
#EndRegion

 Function GenerateVersionSizeReport()
 {  
    #Define 'Web Application URL' as Mandatory Parameter
    Param( [Parameter(Mandatory=$true)] [string]$WebAppURL )
	
	#Get the Web Application
    $WebApp=Get-SPWebApplication($WebAppURL)

    #Write the CSV Header - Tab Separated
	"Site Name`t Library `t File Name `t File URL `t File Type `t Last Modified `t No. of Versions `t Latest Version Size(MB) `t Versions Size(MB) `t Total File Size(MB)" | out-file VersionSizeReport.csv

	#Loop through each site collection
	 foreach($Site in $WebApp.Sites)
	  {
	   #Loop through each site in the site collection
	    foreach($Web in $Site.AllWebs)
		 {
            #Loop through  each List
            foreach ($List in $Web.Lists)
            {
                #Get only Document Libraries & Exclude Hidden System libraries
                if ( ($List.BaseType -eq "DocumentLibrary") -and ($List.Hidden -eq $false) )
                {
                    foreach ($ListItem  in $List.Items)
                    {
 					    #Consider items with 5+ versions
                        if ($ListItem.Versions.Count -gt 1)
                        {
						    $versionSize=0

                            #Get the versioning details
                            foreach ($FileVersion in $ListItem.File.Versions)
                            {
                                $versionSize = $versionSize + $FileVersion.Size;
                            }
							#To Calculate Total Size(MB)
							$ToalFileSize= [Math]::Round(((($ListItem.File.Length + $versionSize)/1024)/1024),2)
							
                            #Convert Size to MB
                            $VersionSize= [Math]::Round((($versionSize/1024)/1024),2)
							
							#Get the Size of the current version
							$CurrentVersionSize= [Math]::Round((($ListItem.File.Length/1024)/1024),2)
							
                            #Get Site Name
                            if ($Web.IsRootWeb -eq $true)
                            {
                                $siteName = $Web.Title +" - Root";
                            }
                            else
                            {
                                $siteName= $Site.RootWeb.Title + " - " + $Web.Title;
                            }

                            #Log the data to a CSV file where versioning size > 0MB!
                            if ($versionSize -gt 0)
                            {
                                "$($siteName) `t $($List.Title) `t $($ListItem.Name) `t $($Web.Url)/$($ListItem.Url) `t $($ListItem['File Type'].ToString()) `t $($ListItem['Modified'].ToString())`t $($ListItem.Versions.Count) `t $CurrentVersionSize `t $($versionSize) `t $($ToalFileSize)" | Out-File VersionSizeReport.csv -Append
                            }
                        }
                    }
                }
            }
		$Web.Dispose()          
        }
	$Site.Dispose()          
    }
 
    #Send message to console
    write-host "Versioning Report Generated Successfully!"
}

#Call the Function to Generate Version History Report
GenerateVersionSizeReport "http://sharepoint.YOURCompany.com"
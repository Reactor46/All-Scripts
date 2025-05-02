# Load the SharePoint PowerShell module (if not already loaded)
if ((Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null) {
    Add-PSSnapin Microsoft.SharePoint.PowerShell
}

$SiteCollURL="https://prod-pulse.kscpulse.com/"
$ReportOutput="D:\SPScripts\MMS-Terms-Usage.csv"
 
$WebsColl = (Get-SPSite $SiteCollURL).AllWebs
 
#Array to hold Results
$ResultColl = @()
 
#Loop through each site, List and Fields
Foreach($web in $WebsColl)
    {
        Write-host "Scanning Web:"$Web.URL
        Foreach($list in $web.Lists)           
        {
            #Get all Managed metadata fields
            Foreach($Field in $list.Fields | where {$_.GetType().Name -eq "TaxonomyField"})
                {
                    Foreach ($item in $list.Items)
                    {
                        #Get All values of MMS field
                        $MMSFieldValueColl = $item[$Field.Title] #-as [Microsoft.SharePoint.Taxonomy.TaxonomyFieldValueCollection]
                         
                        #concatenate each term in the collection
                        $MMSFieldTerms=[string]::Empty
                        Foreach ($MMSFieldValue in $MMSFieldValueColl)
                        {
                            if($MMSFieldValue.label -ne $null)
                            {
                                $MMSFieldTerms+=$MMSFieldValue.label+";"
                            }
                        }
                         
                        #Collect the result
                        if($MMSFieldTerms -ne '')
                        {
                            $Result = New-Object PSObject
                            $Result | Add-Member -type NoteProperty -name "MMs Column Name" -value $Field.Title
                            $Result | Add-Member -type NoteProperty -name "MMS Column Value" -value $MMSFieldTerms
                            #Get the URL of the Item
                            $ItemURL= $Item.ParentList.ParentWeb.Site.MakeFullUrl($item.ParentList.DefaultDisplayFormUrl)
                            $ItemURL=$ItemURL+"?ID=$($Item.ID)"
                            $Result | Add-Member -type NoteProperty -name "Item URL" -value $ItemURL
                            $Result | Add-Member -type NoteProperty -name "List Name" -value $List.Title
                            $Result | Add-Member -type NoteProperty -name "List URL" -value "$($List.ParentWeb.Url)/$($List.RootFolder.Url)"
        
                            $ResultColl += $Result
                        }
                    }
                }
        }
    }
 
#Export Results to a CSV File
$ResultColl | Export-csv $ReportOutput -notypeinformation
Write-Host "Managed Metadata columns usage Report has been Generated!" -f Green


#Read more: https://www.sharepointdiary.com/2015/12/managed-metadata-columns-usage-report-using-powershell.html#ixzz8F6zvevkA
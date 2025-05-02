# Load the SharePoint PowerShell module (if not already loaded)
if ((Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null) {
    Add-PSSnapin Microsoft.SharePoint.PowerShell
}
$SiteCollURL="https://prod-pulse.kscpulse.com/"
$ReportOutput="D:\SPScripts\MMS-Columns.csv"
 
$WebsColl = (Get-SPSite $SiteCollURL).AllWebs
 
#Array to hold Results
$ResultColl = @()
 
#Loop through each site, List and Fields
Foreach($web in $WebsColl)
    {
        Write-host "Scanning Web:"$Web.URL
        Foreach($list in $web.Lists)           
        {
            Foreach($Field in $list.Fields)           
            {           
                if($field.GetType().Name -eq "TaxonomyField")
                {
                    $Result = New-Object PSObject
                    $Result | Add-Member -type NoteProperty -name "List Name" -value $List.Title
                    $Result | Add-Member -type NoteProperty -name "URL" -value "$($List.ParentWeb.Url)/$($List.RootFolder.Url)"
                    $Result | Add-Member -type NoteProperty -name "Field Name" -value $Field.Title
        
                    $ResultColl += $Result
                }
            }
        }
    }
#Export Results to a CSV File
$ResultColl | Export-csv $ReportOutput -notypeinformation
Write-Host "Managed Metadata columns usage Report has been Generated!" -f Green
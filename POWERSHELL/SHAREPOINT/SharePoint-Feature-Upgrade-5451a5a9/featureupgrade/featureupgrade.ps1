if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null) {
    Add-PSSnapin "Microsoft.SharePoint.PowerShell"
}
cls

 ##Provide the Web Application URL
 $webApp = Get-SPWebApplication  ##Web Application URL 
 $featureCollection = $webApp.QueryFeatures([Guid]"a18e2918-ee9d-4e62-9cf7-5fea788c27ed",[bool]$true) ##Provide your Feature ID
 foreach ($feature in $featureCollection) 
 { 
     Write-Host -ForegroundColor Green Site feature is activated on : $feature.Parent.Url
     Write-Host -ForegroundColor Gray Current Feature :- $feature.Definition.DisplayName : $feature.Version   
     $feature.Upgrade([bool]$true);
     Write-Host -ForegroundColor Gray Updated Feature :- $feature.Definition.DisplayName : $feature.Version     
 }

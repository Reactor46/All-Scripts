[cmdletbinding()]
Param()

Write-Host "Getting installed modules" -ForegroundColor Yellow
$modules = Get-Module -ListAvailable

#group to identify modules with multiple versions installed
$g = $modules | Group-Object name -NoElement | Where-Object count -GT 1

Write-Host "Filter to modules from the PSGallery" -ForegroundColor Yellow
$gallery = $modules.where({$_.repositorysourcelocation})

Write-Host "Comparing to online versions" -ForegroundColor Yellow
foreach ($module in $gallery) {

     #find the current version in the gallery
     Try {
        $online = Find-Module -Name $module.name -Repository PSGallery -ErrorAction Stop
     }
     Catch {
        Write-Warning "Module $($module.name) was not found in the PSGallery"
     }

     #compare versions
     if ($online.version -gt $module.version) {
        $UpdateAvailable = $True
     }
     else {
        $UpdateAvailable = $False
     }

     #write a custom object to the pipeline
     [pscustomobject]@{
        Name = $module.name
        MultipleVersions = ($g.name -contains $module.name)
        InstalledVersion = $module.version
        OnlineVersion = $online.version
        Update = $UpdateAvailable
        Path = $module.modulebase
     }
 
} #foreach


Write-Host "Check complete" -ForegroundColor Green
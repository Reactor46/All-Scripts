##############################################
# Author: Megren Dot Net                     # 
# Download SharePoint User Solutions Script  # 
# Version 1.0    Published on 2012-AUG-30    # 
##############################################
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop; 
# Get SharePoint Site
$SiteURL=Read-Host "Insert Site Collection URL";
try
{$Site = Get-SPWeb -Identity $SiteURL -ErrorAction Stop;}
catch
{Write-Host "Error Occured!! On getting Site URL make sure you insert correct URL e.g http://contoso " -ForegroundColor Red
Read-Host  "Press enter to exit, then run the script again"
exit}
# Get SharePoint Solution Catalog URL
$SolutionCatalog=$Site.GetCatalog("SolutionCatalog")
$SolutionCatalogPath=$SolutionCatalog.RootFolder.ServerRelativeUrl
$folder = $Site.GetFolder($SolutionCatalogPath)
$Site.Dispose()
# Make sure the target location path is correct
$Save_Target= Read-Host "Insert Path of Target Location e.g 'C:\' "
if ((Test-Path -path $Save_Target)) 
{# Download files
foreach ($file in $folder.Files) 
{$binary = $file.OpenBinary()
$stream = New-Object System.IO.FileStream($Save_Target + “/” + $file.Name), Create
$writer = New-Object System.IO.BinaryWriter($stream)
$writer.write($binary)
$writer.Close()
$stream.Dispose()
# Write result
Write-Host
Write-Host "Target Folder Path: " $Save_Target -ForegroundColor Yellow
Write-Host "Total of User Solutions: " $folder.Files.Count -ForegroundColor Yellow
Write-Host 
Read-Host  "Press enter to exit"
}}
elseif (!(Test-Path -path $Save_Target)) 
{Write-Host "The Target is incorrect! Make sure you insert correct Path exists on your Server/PC" -ForegroundColor Red
Read-Host  "Press enter to exit, then run the script again"
exit}
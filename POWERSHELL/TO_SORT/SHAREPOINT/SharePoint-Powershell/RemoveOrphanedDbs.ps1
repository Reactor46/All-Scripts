##
## CheckIn file or undo CheckOut with PowerShell
##

# Loading the SharePoint PowerShell Snapin if not already loaded
if ((Get-PsSnapin | ?{$_.Name -eq "Microsoft.SharePoint.PowerShell"})-eq $null)
{
  $PSSnapin = Add-PsSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
}


$dbs = Get-SPDatabase            
$dbs | ?{$_.Exists -eq $false}  | %{Write-Host "DB"$_.Name"does not exist. Deleting and cleaning up references." -f red; $_.Delete();$_.Unprovision()}
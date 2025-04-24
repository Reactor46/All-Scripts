$Sites = (Get-SPSite -Limit All).Url
Foreach($Site in $Sites){
(Get-SPSite $Site).RecycleBin.DeleteAll()
}
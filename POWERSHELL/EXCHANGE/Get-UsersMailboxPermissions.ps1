$MBXFolders = @()
$MBXFoldersCorr = New-Object System.Collections.ArrayList
$Permissions = @()
#$MBX_tocheck = "partler"
#$MBX_tocheck = "stranquillo"
$MBX_tocheck = "jbattista"
$MBXFolders = Get-MailboxFolderStatistics $MBX_tocheck | select folderpath
foreach ($item in $MBXFolders) {
 $temp = $item.FolderPath
 $temp = $Temp.Replace("/","\")
 $MBXFoldersCorr.Add($temp) | out-null
}
foreach ($item in $MBXFoldersCorr) {
Try {
 $MailboxFolder = $MBX_tocheck + ":" + $item
 $FolderPermission = $(Get-MailboxFolderPermission $MailboxFolder -ErrorAction Stop | Select-Object FolderName,User,AccessRights)
 Foreach ($perm in $FolderPermission){
 $ReturnedObj1 = New-Object PSObject
 $ReturnedObj1 | Add-Member NoteProperty -Name "FolderName" -Value $($Perm.FolderName)
 $ReturnedObj1 | Add-Member NoteProperty -Name "FolderPath" -Value $MailboxFolder
 $ReturnedObj1 | Add-Member NoteProperty -Name "User" -Value $($Perm.User)
 $ReturnedObj1 | Add-Member NoteProperty -Name "AccessRights" -Value $($Perm.AccessRights)
 $Permissions += $ReturnedObj1
 }
 }
Catch {
 $ReturnedObj = New-Object PSObject
 $ReturnedObj | Add-Member NoteProperty -Name "FolderName" -Value $item
 $ReturnedObj | Add-Member NoteProperty -Name "FolderPath" -Value $MailboxFolder
 $ReturnedObj | Add-Member NoteProperty -Name "User" -Value "*Not Applicable*"
 $ReturnedObj | Add-Member NoteProperty -Name "AccessRights" -Value "*Not Applicable*"
 $Permissions += $ReturnedObj
 Continue
 }
}
$Permissions | Sort-Object FolderName,User | FT #Export-Csv -Path .\Jbattista-Permissions.csv -Append -NoTypeInformation

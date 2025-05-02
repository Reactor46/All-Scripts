$OldOU = ‘OU=Las_Vegas,DC=contoso,DC=com’
$NewOU = ‘OU=Las_Vegas - Testing,DC=contoso,DC=com’
$oucsv = 'C:\LazyWinAdmin\ActiveDirectory\AD-BACKUP\Las-Vegas-OU.csv'
$success = 0
$failed = 0
$oulist = Import-Csv $oucsv
$oulist | foreach {
$outemp = $_.Distinguishedname -replace $OldOU,$NewOU
#need to split ouTemp and lose the first item
$ousplit = $outemp -split ‘,’,2
$outemp
Try {
$newOU = New-ADOrganizationalUnit -name $_.Name -path $ousplit[1] -EA stop
Write-Host “Successfully created OU: $_.Name”
$success++
}
Catch {
Write-host “ERROR creating OU: $outemp” #$error[0].exception.message”
$failed++
}
Finally {
echo “”
}
}
Write-host “Created $success OUs with $failed errors”
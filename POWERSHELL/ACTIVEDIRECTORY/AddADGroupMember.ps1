$user = import-csv .\footprints.csv
$user
pause
foreach($u in $user){Add-ADGroupMember -identity "GS-Footprints-Dev-Specialty Apps" -Member $u.users}
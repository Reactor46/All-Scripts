# Import active directory module for running AD cmdlets
Import-module ActiveDirectory

#Store the data from UserList.csv in the $List variable
$List = Import-CSV "C:\scripts\AD group scripts\infile\New-Members.csv"

#Loop through user in the CSV
ForEach ($User in $List)
{

#Add the user to the specified group in AD
Add-ADGroupMember -Identity 'USON FILES' -Member $User.username
}
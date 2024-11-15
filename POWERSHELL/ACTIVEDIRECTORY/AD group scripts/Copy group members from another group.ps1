#Input Parameters. Change these as per your requirement
#group1 is the group you are copying membership too
#group2 is the target group with current members
$group1 = "faxmakertest"
$group2 = "nvcs staff"

$membersInGroup1 = Get-ADGroupMember $group1
$membersInGroup2 = Get-ADGroupMember $group2

if($membersInGroup1 -eq $null)
{
    Add-ADGroupMember -Identity $group1 -Members $membersInGroup2
}
elseif($membersInGroup2 -ne $null)
{
  $separateMembers = diff $membersInGroup1 $membersInGroup2

  if($separateMembers -ne $null)
  {
    foreach($member in $separateMembers)
    {
      $currentUserToAdd = Get-ADUser -Identity $member.InputObject
      Add-ADGroupMember -Identity $group1 -Members $currentUserToAdd
      }
  }
}
# Circular.ps1
# List all AD groups that contain one or more circular nested groups
# outputs the parent group's DN and a list of the nested groups.

# Limitations
# The script works by scanning through every group, so any circular relationships
# will be listed twice, once for the parent group and once for the child.

# ADSI has a limit of 1500 items for a multi-valued attribute so 
# groups containing more than 1500 members may return the error:
#  "Get-ADGroupMember : The size limit for this request was exceeded.."

# If a circular relationship exists between a large (>1500) group and small one,
# the script may fail enumerating the large group, but still detect the circular
# relationship against the small one.

# This script only checks direct members: A>B>A
# inherited circular memberships A>B>C>A will not be detected.

cls
Import-Module Activedirectory

write-host "Circular.ps1  Search for nested groups - getting Groups"
# Retrieve all top/parent level AD groups.
$Parents = get-adgroup -ResultPageSize 1000 -filter 'ObjectClass -eq "group"'

# Loop through each parent group
ForEach ($Parent in $Parents)
{
   [int]$Len = 0
   # Create an array of the group members, limited to sub-groups (not users)
   $Children = @(Get-ADGroupMember $Parent | where {$_.objectClass -eq "group"} )
   $Len = @($Children).Count

   if ($Len -eq 1) {"$Parent contains 1 group"}
   elseif ($Len -gt 0) {"$Parent contains $Len groups"}
   
   if ($Len -gt 0)
   {
   "--check nesting"
      ForEach ($Child in $Children)
      {
          # Now find any member of $Child which is also the childs $Parent
          $nested = @(Get-ADGroupMember $Child | where {$_.objectClass -eq "group" -and "$_" -eq "$Parent"} )
          $NestCount = @($nested).Count
          if ($NestCount -gt 0)
          {
            " "
            "   Found a circular nested group: "
            "   $nested is both a parent and a member of:"
            "   $Child"
            "   ========================================"
          }
      }
   "--done"
   }
}
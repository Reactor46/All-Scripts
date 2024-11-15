function Get-AdGroupNestedGroupMembership {
##########################################################################################################
<#
.SYNOPSIS
    Get's an Active Directory group's nested group memberships.

.DESCRIPTION
    Displays nested group membership details for all the groups that an AD group is a member of. Searches 
    up through the group membership hiererachy using a slight modification of the script founde here:

    https://blogs.msdn.microsoft.com/b/adpowershell/archive/2009/09/05/token-bloat-troubleshooting-by-analyzing-group-nesting-in-ad.aspx

    Produces a custom object for each group and has an option -TreeView switch for a hierarchical
    display.

.EXAMPLE
    Get-AdGroupNestedGroupMembership -Group "CN=Dreamers,OU=Groups,DC=halo,DC=net" -Domain "halo.net"

    Shows the nested group membership for the group 'CN=Dreamers,OU=Groups,DC=halo,DC=net', from the 
    'halo.net' domain.

    For example:

    BaseGroup                  : CN=Dreamers,OU=Groups,DC=halo,DC=net
    BaseGroupAdminCount        : 1
    MaxNestingLevel            : 2
    NestedGroupMembershipCount : 3
    DistinguishedName          : CN=Super Admins,OU=Groups,DC=corp,DC=halo,DC=net
    GroupCategory              : Security
    GroupScope                 : Universal
    Name                       : Super Admins
    ObjectClass                : group
    ObjectGUID                 : a96c6da9-b45b-4820-b30f-8e7f8a256ca8
    SamAccountName             : Super Admins
    SID                        : S-1-5-21-1741080060-3640901959-2469866113-1106


.EXAMPLE
    Get-AdGroupNestedGroupMembership -Group "Dreamers" -Domain "halo.net" -TreeView 

    Shows the nested group membership for the group 'Dreamers', from the 'halo.net' domain, 
    with a hierarchical tree view.

    For example:

    Super Admins
    └─Enterprise Admins
      ├─Denied RODC Password Replication Group
      └─Administrators

    BaseGroup                  : CN=Dreamers,OU=Groups,DC=halo,DC=net
    BaseGroupAdminCount        : 1
    MaxNestingLevel            : 2
    NestedGroupMembershipCount : 3
    DistinguishedName          : CN=Super Admins,OU=Groups,DC=corp,DC=halo,DC=net
    GroupCategory              : Security
    GroupScope                 : Universal
    Name                       : Super Admins
    ObjectClass                : group
    ObjectGUID                 : a96c6da9-b45b-4820-b30f-8e7f8a256ca8
    SamAccountName             : Super Admins
    SID                        : S-1-5-21-1741080060-3640901959-2469866113-1106

.EXAMPLE
    Get-AdGroup -Identity 'Dreamers' | Get-AdGroupNestedGroupMembership -TreeView

    Gets an object for the AD group 'Dreamers' and then pipes it into the Get-AdGroupNestedGroupMembership
    function. Provides a hierarchical tree view. Uses the current domain.

    For example:

    Super Admins
    └─Enterprise Admins
      ├─Denied RODC Password Replication Group
      └─Administrators

    BaseGroup                  : CN=Dreamers,OU=Groups,DC=halo,DC=net
    BaseGroupAdminCount        : 1
    MaxNestingLevel            : 2
    NestedGroupMembershipCount : 3
    DistinguishedName          : CN=Super Admins,OU=Groups,DC=corp,DC=halo,DC=net
    GroupCategory              : Security
    GroupScope                 : Universal
    Name                       : Super Admins
    ObjectClass                : group
    ObjectGUID                 : a96c6da9-b45b-4820-b30f-8e7f8a256ca8
    SamAccountName             : Super Admins
    SID                        : S-1-5-21-1741080060-3640901959-2469866113-1106

.EXAMPLE
    Get-AdGroup -Filter * -SearchBase "OU=Groups,DC=halo,DC=net" -Server 'halo.net' |
    Get-AdGroupNestedGroupMembership | 
    Export-CSV -Path d:\users\timh\nestings.csv

    Gets all of the groups from the 'Groups' OU in the 'halo.net' domain and each AD object into
    the Get-AdGroupNestedGroupMembership function. Objects from the Get-AdGroupGroupMembership function
    are then exported to a CSV file named d:\users\timh\nestings.csv

    For example:

    #TYPE Microsoft.ActiveDirectory.Management.AdGroup
    "BaseGroup","BaseGroupAdminCount","MaxNestingLevel","NestedGroupMembershipCount","DistinguishedName"...
    "CN=Dreamers,OU=Groups,DC=halo,DC=net","1","2","3","CN=Super Admins,OU=Groups,DC=corp,DC=halo,DC=net"...

.NOTES
    THIS CODE-SAMPLE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED 
    OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR 
    FITNESS FOR A PARTICULAR PURPOSE.

    This sample is not supported under any Microsoft standard support program or service. 
    The script is provided AS IS without warranty of any kind. Microsoft further disclaims all
    implied warranties including, without limitation, any implied warranties of merchantability
    or of fitness for a particular purpose. The entire risk arising out of the use or performance
    of the sample and documentation remains with you. In no event shall Microsoft, its authors,
    or anyone else involved in the creation, production, or delivery of the script be liable for 
    any damages whatsoever (including, without limitation, damages for loss of business profits, 
    business interruption, loss of business information, or other pecuniary loss) arising out of 
    the use of or inability to use the sample or documentation, even if Microsoft has been advised 
    of the possibility of such damages, rising out of the use of or inability to use the sample script, 
    even if Microsoft has been advised of the possibility of such damages. 
#>
##########################################################################################################

##################################
## Script Options and Parameters
##################################

#Requires -version 3

#Define and validate parameters
[CmdletBinding()]
Param(
      #The target group
      [parameter(Mandatory,Position=1,ValueFromPipeline=$True)]
      [ValidateScript({Get-AdGroup -Identity $_})] 
      [String]$Group,

      #The target domain
      [parameter(Position=2)]
      [ValidateScript({Get-ADDomain -Identity $_})] 
      [String]$Domain = $(Get-ADDomain).Name,

      #Whether to produce the tree view
      [Switch]$TreeView
      )

#Set strict mode to identify typographical errors (uncomment whilst editing script)
#Set-StrictMode -version Latest

#Version : 1.0


##########################################################################################################


    ########
    ## Main
    ########

    Begin {

        #Load Get-ADNestedGroups function...

        ##########################################################################################################

        #################################
        ## Function - Get-ADNestedGroups
        #################################

        function Get-ADNestedGroups {

        <#
        Code adapted from:
        https://blogs.msdn.microsoft.com/b/adpowershell/archive/2009/09/05/token-bloat-troubleshooting-by-analyzing-group-nesting-in-ad.aspx    
        #>

            Param ( 
                [Parameter(Mandatory=$true, 
                    Position=0, 
                    ValueFromPipeline=$true, 
                    HelpMessage="DN or ObjectGUID of the AD Group." 
                )] 
                [string]$groupIdentity, 
                [string]$groupDn,
                [int]$groupAdmin = 0,
                [switch]$showTree 
                ) 

            $global:numberOfRecursiveGroupMemberships = 0 
            $lastGroupAtALevelFlags = @() 

            function Get-GroupNesting ([string] $identity, [int] $level, [hashtable] $groupsVisitedBeforeThisOne, [bool] $lastGroupOfTheLevel) 
            { 
                $group = $null 
                $group = Get-AdGroup -Identity $identity -Properties "memberOf"
                if($lastGroupAtALevelFlags.Count -le $level) 
                { 
                    $lastGroupAtALevelFlags = $lastGroupAtALevelFlags + 0 
                } 
                if($group -ne $null) 
                { 
                    if($showTree) 
                    { 
                        for($i = 0; $i -lt $level – 1; $i++) 
                        { 
                            if($lastGroupAtALevelFlags[$i] -ne 0) 
                            { 
                                Write-Host -ForegroundColor Yellow -NoNewline "  " 
                            } 
                            else 
                            { 
                                Write-Host -ForegroundColor Yellow -NoNewline "│ " 
                            } 
                        } 
                        if($level -ne 0) 
                        { 
                            if($lastGroupOfTheLevel) 
                            { 
                                Write-Host -ForegroundColor Yellow -NoNewline "└─" 
                            } 
                            else 
                            { 
                                Write-Host -ForegroundColor Yellow -NoNewline "├─" 
                            } 
                        } 
                        Write-Host -ForegroundColor Yellow $group.Name 
                    } 
                    $groupsVisitedBeforeThisOne.Add($group.distinguishedName,$null) 
                    $global:numberOfRecursiveGroupMemberships ++ 
                    $groupMemberShipCount = $group.memberOf.Count 
                    if ($groupMemberShipCount -gt 0) 
                    { 
                        $maxMemberGroupLevel = 0 
                        $count = 0 
                        foreach($groupDN in $group.memberOf) 
                        { 
                            $count++ 
                            $lastGroupOfThisLevel = $false 
                            if($count -eq $groupMemberShipCount){$lastGroupOfThisLevel = $true; $lastGroupAtALevelFlags[$level] = 1} 
                            if(-not $groupsVisitedBeforeThisOne.Contains($groupDN)) #prevent cyclic dependancies 
                            { 
                                $memberGroupLevel = Get-GroupNesting -Identity $groupDN -Level $($level+1) -GroupsVisitedBeforeThisOne $groupsVisitedBeforeThisOne -lastGroupOfTheLevel $lastGroupOfThisLevel 
                                if ($memberGroupLevel -gt $maxMemberGroupLevel){$maxMemberGroupLevel = $memberGroupLevel} 
                            } 
                        } 
                        $level = $maxMemberGroupLevel 
                    } 
                    else #we’ve reached the top level group, return it’s height 
                    { 
                        return $level 
                    } 
                    return $level 
                } 
            } 
            $global:numberOfRecursiveGroupMemberships = 0 
            $groupObj = $null 
            $groupObj = Get-AdGroup -Identity $groupIdentity
            if($groupObj) 
            { 
                [int]$maxNestingLevel = Get-GroupNesting -Identity $groupIdentity -Level 0 -GroupsVisitedBeforeThisOne @{} -lastGroupOfTheLevel $false 
                Add-Member -InputObject $groupObj -MemberType NoteProperty  -Name BaseGroup -Value $groupDn -Force
                Add-Member -InputObject $groupObj -MemberType NoteProperty  -Name BaseGroupAdminCount -Value $groupAdmin -Force
                Add-Member -InputObject $groupObj -MemberType NoteProperty  -Name MaxNestingLevel -Value $maxNestingLevel -Force 
                Add-Member -InputObject $groupObj -MemberType NoteProperty  -Name NestedGroupMembershipCount -Value $($global:numberOfRecursiveGroupMemberships – 1) -Force 
                $groupObj 
            }

        }   #end of function Get-ADNestedGroups


        ##########################################################################################################


        #Connect to a Global Catalogue
        $GC = New-PSDrive -PSProvider ActiveDirectory -Server $Domain -Root "" –GlobalCatalog –Name GC

        #Error checking
        if ($GC) {

            #Set location to GC drive
            Set-Location -Path GC:

        }   #end of if ($GC)
        else {

            #Error and exit
            Write-Error -Message "Failed to create GC drive. Exiting function..."

        }   #end of else ($GC)


    }   #end of Begin


    Process {

        #Now get a list of the group's group memberships
        $AdGroup = Get-AdGroup -Identity $Group -Server $Domain -Properties MemberOf,AdminCount -ErrorAction SilentlyContinue
        $Groups = ($AdGroup).MemberOf

        #Error checking
        if ($Groups) {

            #Loop through each of the groups found
            foreach ($Group in $Groups) {

                #Run group query with or without -TreeView
                if ($TreeView) {

                    #Call Get-ADNestedGroups function with -showTree
                    Get-ADNestedGroups -groupIdentity $Group -groupDN ($AdGroup).DistinguishedName -groupAdmin ($AdGroup).AdminCount -showTree 


                }   #end of if ($TreeView)
                else {

                    #Call Get-ADNestedGroups function without -showTree
                    Get-ADNestedGroups -groupIdentity $Group -groupDN ($AdGroup).DistinguishedName -groupAdmin ($AdGroup).AdminCount

                }   #end of else $TreeView


            }   #end of foreach ($Group in $Groups)

        }   #end of if ($Groups)
        else {

            Write-Warning -Message "No group memberships returned for group - $group"

        }   #end of else ($Groups)

    }   #end of Process

    End {

        #Exit the GC PS drive and remove
        if ((Get-Location).Drive.Name -eq "GC") {

            #Move to C: drive
            C:

        }   #end of if ((Get-Location).Drive.Name -eq "GC")

    }   #end of End

}   #end of function Get-AdGroupNestedGroupMembership
# @TODO Bug about NT AUTHORITY
#
#: Processing role IIS_IUSRS
# : Found  1  group(s) or member(s) in  IIS_IUSRS
#: Role IIS_IUSRS Found IUSR (Group) ads=WinNT://NT AUTHORITY/IUSR
#: Role IIS_IUSRS Group detected. listing users in group IUSR (Group)
#: Listing member in Group "IUSR" in NT AUTHORITY Directory

# @TODO  size limit
#Get-ADGroupMember : The size limit for this request was exceeded
    #At D:\GitHub\ServerAccessFromAD\develop2.ps1:64 char:56
#+                      $MemberInGroup = Get-ADGroupMember <<<<  -identity $groupName -Recursive -server $company | select name
#    + CategoryInfo          : NotSpecified: (Domain Users:ADGroup) [Get-ADGroupMember], ADException
#    + FullyQualifiedErrorId : The size limit for this request was exceeded,Microsoft.ActiveDirectory.Management.Commands.GetADGroupMember
    
# @ TODO  Caching
# Beware of cache the result of query

# @ Todo query user company on varible $AccountCompany @discarded

# @Todo  Add environment support for the computer list. by detect format 
# xyzdev|computername1 
# xyzprod|computername2

Import-Module ActiveDirectory
function get-main {
    param(
    [Parameter(Mandatory=$true,valuefrompipeline=$true)]
    [string]$computername
    )
    begin {}
    Process {
        ## looping for one computer machine.
            Write-Host "Start processing " $computername
        $roles =  @() # create a new array
        $MachineADSI = [ADSI]("WinNT://" + $computername + ",computer")
        #$AccountLevel = "Administrators"
        $roles = $MachineADSI.psbase.Children | Where {$_.psbase.schemaClassName -eq "group"}
        ForEach ($Role In $roles)
        {
            $AccountLevel = $Role.name
            Write-Host "$computername : Processing role $AccountLevel"
            #-------------------------------------
            $RoleAD = $MachineADSI.psbase.children.find($AccountLevel)
            $Rolemembers = @($RoleAD.psbase.invoke("Members"))
            Write-Host "$computername : Found " $rolemembers.count " group(s) or member(s) in " $AccountLevel
            if ($rolemembers.count -eq 0) {
                printtoexcel $computername $AccountLevel "No" "No" "No" "No"
            }
            foreach ($admin in $Rolemembers) {
                $Class = $admin.GetType().InvokeMember("Class", 'GetProperty', $Null, $admin, $Null)
                $accountOrGroupName = $admin.GetType().InvokeMember("Name", 'GetProperty', $Null, $admin, $Null)
                $adspath = $admin.GetType().InvokeMember("adspath", 'GetProperty', $Null, $admin, $Null)
                Write-Host "$computername : Role $AccountLevel Found $accountOrGroupName ($Class) ads=$adspath"
                if ($Class -like "group" ) {
                    Write-Host "$computername : Role $AccountLevel Group detected. listing users in group $accountOrGroupName ($Class)"
                    $groupName = $accountOrGroupName
                    $GroupDirectoryServer = $adspath.substring(8) ## substring WinNT:// out
                    $GroupDirectoryServer = $GroupDirectoryServer.substring(0,$GroupDirectoryServer.indexOf("/"))
                    if ($GroupDirectoryServer -eq 'NT AUTHORITY') {
                        # do anything need to query an NT Authority group.
                        printtoexcel $computername $AccountLevel "NT AUTHORITY" "GROUP" $accountOrGroupName $AccountCompany $path
                    }
                    Elseif ($GroupDirectoryServer -eq 'NT SERVICE') {
                        printtoexcel $computername $AccountLevel "NT SERVICE" "GROUP" $accountOrGroupName $AccountCompany $path
                        # printtoExcel parameter =  [String]$Computername, [String]$AccountLevel, [String]$GroupDirectoryServer,  [String]$UserAccount,  [String]$AccountCompany,  [String]$Accesspath)
                    }
                    Else {
                        ListPeopleInGroup $computername $AccountLevel $GroupName $GroupDirectoryServer
                    }
                } else {
                    $path =  $AccountLevel + "->" + $accountOrGroupName
                    printtoexcel $computername $AccountLevel "Direct" "Direct" $accountOrGroupName $AccountCompany $path
                }
            }
            #------------------------------------------------------
        }
        Write-Host "$computername : Finished processing $computername"
        
    }
    end {}
    
}


function ListPeopleInGroup(
[String]$Computername,
[String]$AccountLevel,
[String]$groupName,
[String]$Company) {
    $AccountCompany = "nextversion"
    Write-Host "$computername : Listing member in Group `"$groupName`" in $company Directory"
    $MemberInGroup = @()
    Try
    {
        $MemberInGroup = Get-ADGroupMember -identity $groupName -Recursive -server $company | select name
    }
    Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
    {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
    }
    Catch [Microsoft.ActiveDirectory.Management.Commands.GetADGroupMember]
    {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
    }
    foreach($member in $MemberInGroup ) {
        $path = $AccountLevel + "->" + $groupName + "->" + $member.Name
        printtoexcel $computername $AccountLevel $company $groupName $member.Name $AccountCompany $path
    }
    Write-Host "$computername : Found " $MemberInGroup.Count " member(s) in $company directory checking next group/user"
}

function printtoexcel(
[String]$Computername,
[String]$AccountLevel,
[String]$GroupDirectoryServer,
[String]$UserAccount,
[String]$AccountCompany,
[String]$Accesspath)
{
    $out = New-Object psobject
    $AccountCompany = "Next version"
    $out | Add-Member noteproperty ComputerName $computername
    $out | Add-Member noteproperty AccountLevel $AccountLevel
    $out | Add-Member noteproperty Company $Company
    $out | Add-Member noteproperty GroupDirectoryServer $GroupDirectoryServer
    $out | Add-Member noteproperty UserAccount $UserAccount
    $out | Add-Member noteproperty AccountCompany $AccountCompany
    $out | Add-Member noteproperty AccessGrantedPath $AccessPath
    Write-Output $out
}
Get-Content computers.txt | get-main | Export-Csv export.csv -notype
Import-Module ActiveDirectory

#'====================================='
#'(1) corp.vegas.com'
#'(2) svc.prod.vegas.com'
#'(3) test.vegas.com'
#'(4) res.vegas.com'
#'(5) Exit'
#" "
$domain = $env:UserDomain

switch ($domain)
{
    CORP { $domain = "corp.vegas.com" }
    SVC { $domain = "svc.prod.vegas.com" }
    TEST { $domain = "test.vegas.com" }
    RES { $domain = "res.vegas.com" }
    PROD { $domain = "prod.vegas.com" }
    6 { exit }
}
" "
'====================================='

echo "Domain is $domain"
" "
'====================================='
'(1)Add Unix Attributes to a User'
'(2)Display User Properties'
'(3)Display Group Properties'
'(4)Add Unix Attributes to a group'
'(5)Add a user to a Unix group'
'(6)Exit'
" "
$result = Read-Host -Prompt 'Selection'

switch ($result)
    {
        1 { #add attributes to user
            $userquery = Read-Host -Prompt 'Enter the username to query'
           #get user properties            
           "======================================================================================================"
           "Current Unix User Attributes" 
           Get-ADUser -server $domain $userquery -Properties * | Select SamAccountName, msSFU30NisDomain,uidNumber, unixHomeDirectory, loginShell, gidnumber, @{Label='PrimaryGroupDN';Expression={(Get-ADGroup -Filter {GIDNUMBER -eq $_.gidnumber}).SamAccountName}}
           "______________________________________________________________________________________________________" 
           #get unix group properties            
           $gidnumber2 = Get-ADUser -server $domain $userquery -Properties * | Select -expand gidnumber
           #"gidnumber $gidnumber2"
           
           #"PrimaryGroupDN $groupname"
           "Current Unix Group Attributes"
           #check if Unix group is assigned
           if([string]::IsNullOrEmpty($gidnumber2)) {            
              Write-Host "No Unix group assigned"            
            } else {            
                $groupname = (Get-ADGroup -server $domain -Filter {GIDNUMBER -eq $gidnumber2}).SamAccountName
                Get-ADGroup -server $domain $groupname -Properties *| format-list msSFU30Name, msSFU30NisDomain, gidnumber
                "Members:"
                Get-ADGroup -server $domain $groupname -Properties msSFU30PosixMember | select-object -expandproperty msSFU30PosixMember        
           
           }
           "======================================================================================================"
           " "


            $nisdomain = Read-Host -Prompt 'Please select a NIS domain; corp, svc, test, res (lowercase only)'
            switch ($nisdomain)
            {
                corp { " $nisdomain selected" }
                svc { " $nisdomain selected" }
                test { " $nisdomain selected" }
                res { " $nisdomain selected" }
                default { " $nisdomain is invalid"
                         exit }
            }

            #check if user already has an uid
            $useridnumber = Get-ADUser -server $domain $userquery -Properties * | select -expand  uidNumber

            if ([string]::IsNullOrEmpty($useridnumber))
            {
                #get next uidnumber
                $useridnumber = get-aduser -server $domain -filter 'uidnumber -ge 2000 -and uidnumber -lt 3000' -properties * | select -expand uidnumber | sort-object | select-object -last 1            
                if ([string]::IsNullOrEmpty($useridnumber)) { $useridnumber = 2000 }

                         
                "last item $useridnumber "
                $useridnumber++
                "next id $useridnumber "
            }

            $userhome = Read-Host -Prompt 'Please enter the Unix home directory path'
            $usershell = Read-Host -Prompt 'Please enter the login shell path'

            #list Unix groups
            get-adgroup -server $domain -filter 'gidnumber -ge 0' -properties * | select samaccountname,gidnumber | ft
            $usergroup = Read-Host -Prompt 'Please enter the Unix group number'

            $unixgroupname = (Get-ADGroup -server $domain -Filter {GIDNUMBER -eq $usergroup}).SamAccountName
            " "
            "=====Values Entered===================================================================================="
            
            "SamAccountName    : $userquery "
            "msSFU30NisDomain  : $nisdomain "
            "uidNumber         : $useridnumber "
            "unixHomeDirectory : $userhome "
            "loginShell        : $usershell "
            "gidnumber         : $usergroup "
            "PrimaryGroupDN    : $unixgroupname "
            " "
            $continue = Read-Host -Prompt 'Would you like to write these values to the user; y or n'

            switch ($continue)
            {
                n { exit }
                y { 
                    Set-ADUser -server $domain $userquery -Replace @{uid=$userquery;unixUserPassword='ABCD!efgh12345$67890';uidNumber=$useridnumber;gidNumber=$usergroup;unixHomeDirectory=$userhome;loginShell=$usershell;msSFU30NisDomain=$nisdomain;msSFU30Name=$userquery} 
                    $userdistinguishedname = get-aduser -server $domain $userquery | Select -expand DistinguishedName
                    Add-ADGroupMember -server $domain -Identity $unixgroupname -Members $userquery
                    Set-ADGroup -server $domain $unixgroupname -Add @{msSFU30PosixMember="$userdistinguishedname";memberUid="$userquery"}
                    #Get-ADGroup -server $domain $unixgroupname -Properties *| format-list msSFU30Name, msSFU30NisDomain, gidnumber
                    #Start-Sleep -s 5
                    #"Members:"
                    #Get-ADGroup -server $domain $unixgroupname -Properties msSFU30PosixMember | select-object -expandproperty msSFU30PosixMember
                    
                  }

            }


          }
        2 { #get user Unix attributes
            $userquery = Read-Host -Prompt 'Enter the username to query'
                      
           "======================================================================================================"
           "Unix User Attributes" 
           Get-ADUser -server $domain $userquery -Properties * | Select SamAccountName, msSFU30NisDomain,uidNumber, unixHomeDirectory, loginShell, gidnumber, @{Label='PrimaryGroupDN';Expression={(Get-ADGroup -Filter {GIDNUMBER -eq $_.gidnumber}).SamAccountName}}
           "______________________________________________________________________________________________________" 
           #get unix group properties            
           $gidnumber2 = Get-ADUser -server $domain $userquery -Properties * | Select -expand gidnumber
           #"gidnumber $gidnumber2"
           $groupname = (Get-ADGroup -server $domain -Filter {GIDNUMBER -eq $gidnumber2}).SamAccountName
           #"PrimaryGroupDN $groupname"
           "Unix Group Attributes"
           Get-ADGroup -server $domain $groupname -Properties *| format-list msSFU30Name, msSFU30NisDomain, gidnumber
           "Members:"
           Get-ADGroup -server $domain $groupname -Properties msSFU30PosixMember | select-object -expandproperty msSFU30PosixMember        
           "======================================================================================================"
          } 
        3 { #Show Group Unix attributes
            $groupquery = Read-Host -Prompt 'Enter the group name to query'
            #get unix group properties            
            "======================================================================================================"
            "Unix Group Attributes"
            Get-ADGroup -server $domain $groupquery -Properties * | format-list msSFU30Name, msSFU30NisDomain, gidnumber
            "Members:"
            Get-ADGroup -server $domain $groupquery -Properties msSFU30PosixMember | select-object -expandproperty msSFU30PosixMember
            "======================================================================================================"
          }
        4 { #add attributes to group
            $groupname = Read-Host -Prompt 'Please enter a group name'
            Get-ADGroup -server $domain $groupname -Properties *| format-list msSFU30Name, msSFU30NisDomain, gidnumber
            
            $continue = Read-Host -Prompt 'Would you like to modify the attributes of this group (y or n)?'
            switch ($continue)
            {
              y {
                $nisdomain = Read-Host -Prompt 'Please select a NIS domain; corp, svc, test, res (lowercase only)'
                switch ($nisdomain)
                {
                    corp { " $nisdomain selected" }
                    svc { " $nisdomain selected" }
                    test { " $nisdomain selected" }
                    res { " $nisdomain selected" }
                    default { " $nisdomain is invalid"
                            exit }
                }

                #get next group id number
                $groupidarray = get-adgroup -server $domain -filter 'gidnumber -ge 0' -properties * | select -expand gidnumber | sort-object | select-object -last 1
                "last item $groupidarray "
                $groupidarray++
                "next id $groupidarray "

                Set-ADGroup -server $domain $groupname -Replace @{msSFU30Name=$groupname;gidNumber=$groupidarray;msSFU30NisDomain=$nisdomain}
                Start-Sleep -s 5
                Get-ADGroup -server $domain $groupname -Properties *| format-list msSFU30Name, msSFU30NisDomain, gidnumber
            

                #"$groupname $nisdomain $groupmember"
                }
              n { exit }
            }
            
          }
        5 {  #add a user to a unix group       
            
            $userquery = Read-Host -Prompt "Please enter the username to add "
            $unixgroupname = Read-Host -Prompt "What group would you like to add $userquery to"
            $userdistinguishedname = get-aduser -server $domain $userquery | Select -expand DistinguishedName
            Add-ADGroupMember -server $domain -Identity $unixgroupname -Members $userquery
            Set-ADGroup -server $domain $unixgroupname -Add @{msSFU30PosixMember="$userdistinguishedname";memberUid="$userquery"}
            Get-ADGroup -server $domain $unixgroupname -Properties *| format-list msSFU30Name, msSFU30NisDomain, gidnumber
            #Start-Sleep -s 5
            #"Members:"
            #Get-ADGroup $groupname -server $domain -Properties msSFU30PosixMember | select-object -expandproperty msSFU30PosixMember

          }
        6 { exit }
                  
        default { " $result is invalid" }
            
        
         
    }
    exit


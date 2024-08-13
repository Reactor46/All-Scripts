function Get-ADACL {
    Param($DNPath,[switch]$SDDL,[switch]$help,[switch]$verbose)
    function HelpMe{
        Write-Host
        Write-Host " Get-ADACL.ps1:" -fore Green
        Write-Host "   Gets ACL object or SDDL for AD Object"
        Write-Host
        Write-Host " Parameters:" -fore Green
        Write-Host "   -DNPath                : Parameter: DN of Object"
        Write-Host "   -sddl                  : [SWITCH]:  Output SDDL instead of ACL Object"
        Write-Host "   -Verbose               : [SWITCH]:  Enables Verbose Output"
        Write-Host "   -Help                  : [SWITCH]:  Displays This"
        Write-Host
        Write-Host " Examples:" -fore Green
        Write-Host "   Get ACL for ‘cn=users,dc=corp,dc=lab’" -fore White
        Write-Host "     .\Get-ADACL.ps1 ‘cn=users,dc=corp,dc=lab’" -fore Yellow
        Write-Host "   Get SDDL for ‘cn=users,dc=corp,dc=lab’" -fore White
        Write-Host "     .\Get-ADACL.ps1 ‘cn=users,dc=corp,dc=lab’ -sddl " -fore Yellow
        Write-Host
    }
    
    if(!$DNPath -or $help){HelpMe;return}
    
    Write-Host
    if($verbose){$verbosepreference="continue"}
    
    Write-Verbose " + Processing Object [$DNPath]"
    $DE = [ADSI]"LDAP://$DNPath"
    
    Write-Verbose "   – Getting ACL"
    $acl = $DE.psbase.ObjectSecurity
    if($SDDL)
    {
        Write-Verbose "   – Returning SDDL"
        $acl.GetSecurityDescriptorSddlForm([System.Security.AccessControl.AccessControlSections]::All)
    }
    else
    {
        Write-Verbose "   – Returning ACL Object [System.DirectoryServices.ActiveDirectoryAccessRule]"
        $acl.GetAccessRules($true,$true,[System.Security.Principal.SecurityIdentifier])
    }
}
function Set-ADACL {
    Param($DNPath,$acl,$sddl,[switch]$verbose,[switch]$help)
    function HelpMe{
        Write-Host
        Write-Host " Set-ADACL.ps1:" -fore Green
        Write-Host "   Sets the AD Object ACL to ‘ACL Object’ or ‘SDDL’ String"
        Write-Host
        Write-Host " Parameters:" -fore Green
        Write-Host "   -DNPath                : Parameter: DN of Object"
        Write-Host "   -ACL                   : Parameter: ACL Object"
        Write-Host "   -sddl                  : Parameter: SDDL String"
        Write-Host "   -Verbose               : [SWITCH]:  Enables Verbose Output"
        Write-Host "   -Help                  : [SWITCH]:  Displays This"
        Write-Host
        Write-Host " Examples:" -fore Green
        Write-Host "   Set ACL on ‘cn=users,dc=corp,dc=lab’ using ACL Object" -fore White
        Write-Host "     .\Set-ADACL.ps1 ‘cn=users,dc=corp,dc=lab’ -ACL $acl" -fore Yellow
        Write-Host "   Set ACL on ‘cn=users,dc=corp,dc=lab’ using SDDL" -fore White
        Write-Host "     .\Set-ADACL.ps1 ‘cn=users,dc=corp,dc=lab’ -sddl `$mysddl" -fore Yellow
        Write-Host
    }
    
    if(!$DNPath -or (!$acl -and !$sddl) -or $help){HelpMe;Return}
    
    Write-Host
    if($verbose){$verbosepreference="continue"}
    Write-Verbose " + Processing Object [$DNPath]"
    
    $DE = [ADSI]"LDAP://$DNPath"
    if($sddl)
    {
        Write-Verbose "   – Setting ACL using SDDL [$sddl]"
        $DE.psbase.ObjectSecurity.SetSecurityDescriptorSddlForm($sddl)
    }
    else
    {
        foreach($ace in $acl)
        {
            Write-Verbose "   – Adding Permission [$($ace.ActiveDirectoryRights)] to [$($ace.IdentityReference)]"
            $DE.psbase.ObjectSecurity.SetAccessRule($ace)
        }
    }
    $DE.psbase.commitchanges()
    Write-Host
}
function New-ADACE {
    Param([System.Security.Principal.IdentityReference]$identity,
        [System.DirectoryServices.ActiveDirectoryRights]$adRights,
        [System.Security.AccessControl.AccessControlType]$type,
        $Guid)
    
    $help = @"
    $identity
        System.Security.Principal.IdentityReference
        http://msdn.microsoft.com/en-us/library/system.security.principal.ntaccount.aspx
    
    $adRights
        System.DirectoryServices.ActiveDirectoryRights
        http://msdn.microsoft.com/en-us/library/system.directoryservices.activedirectoryrights.aspx
    
    $type
        System.Security.AccessControl.AccessControlType
        http://msdn.microsoft.com/en-us/library/w4ds5h86(VS.80).aspx 
    
    $Guid
        Object Type of the property
        The schema GUID of the object to which the access rule applies.
"@
    $ACE = New-Object System.DirectoryServices.ActiveDirectoryAccessRule($identity,$adRights,$type,$guid)
    $ACE
}
function ConvertTo-Sid($UserName,$domain = $env:Computername) {
   $ID = New-Object System.Security.Principal.NTAccount($domain,$UserName)
   $SID = $ID.Translate([System.Security.Principal.SecurityIdentifier])
   $SID.Value
}
function ConvertTo-Name($sid) {
   $ID = New-Object System.Security.Principal.SecurityIdentifier($sid)
   $User = $ID.Translate( [System.Security.Principal.NTAccount])
   $User.Value
}
function Get-Self{
   ([Security.Principal.WindowsIdentity]::GetCurrent())
}
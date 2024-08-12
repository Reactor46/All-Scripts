# FindUserBadPwdAttempts.ps1
# PowerShell Version 1 script to assist in troubleshooting accounts
# experiencing bad password attempts. It can also be used to investigate
# how accounts get locked out in Active Directory. The script finds
# the values of the sAMAccountName, pwdLastSet, lockoutTime, lastLogon,
# logonCount, badPwdCount, and badPasswordTime attributes for a specified
# user. The last 4 attributes are not replicated, so a different value is
# saved on every domain controller in the domain. A separate line of output
# is generated for each domain controller. The script prompts for either
# the distingished name or the sAMAccountName of an account.

# Author: Richard L. Mueller
# Version 1.0 - October 6, 2015

# Check for parameter identifying the object.
If ($Args.Count -eq 1)
{
    $UserName = $Args[0]
}
Else
{
    $UserName = Read-Host "Enter user sAMAccountName or distinguished name"
}

# Setup the DirectorySearcher object.
$Searcher = New-Object System.DirectoryServices.DirectorySearcher
$Searcher.PageSize = 200
$Searcher.SearchScope = "subtree"
$Domain = New-Object System.DirectoryServices.DirectoryEntry
$BaseDN = $Domain.distinguishedName

If ($UserName -Like "*,*")
{
    $Searcher.Filter = "(distinguishedName=$UserName)"
}
Else
{
    $Searcher.Filter = "(sAMAccountName=$UserName)"
}

# Check for existence of the user object on any DC.
$Searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$BaseDN"
$Check = $Searcher.FindOne()
If (-Not $Check)
{
    Write-Host "Error: User $UserName not found" -foregroundcolor red
    Break
}

# Specify the attribute values to retrieve.
$Searcher.PropertiesToLoad.Add("distinguishedName") > $Null
$Searcher.PropertiesToLoad.Add("sAMAccountName") > $Null
$Searcher.PropertiesToLoad.Add("badPwdCount") > $Null
$Searcher.PropertiesToLoad.Add("badPasswordTime") > $Null
$Searcher.PropertiesToLoad.Add("lockoutTime") > $Null
$Searcher.PropertiesToLoad.Add("logonCount") > $Null
$Searcher.PropertiesToLoad.Add("lastLogon") > $Null
$Searcher.PropertiesToLoad.Add("pwdLastSet") > $Null

$D = [system.directoryservices.activedirectory.Domain]::GetCurrentDomain()
$PDC= $D.PdcRoleOwner

# Output a heading line.
"sAMAccountName, pwdLastSet, lockoutTime, DC, lastLogon, logonCount, badPwdCount, badPasswordTime"

# Query every domain controller in the domain.
ForEach ($DC In $D.DomainControllers)
{
    $Server = $DC.Name
    $Result = $Null
    # Identify the DC with the PDC Emulator role.
    If ($Server -eq $PDC) {$ServerName = "$Server (PDCe)"}
    Else {$ServerName = $Server}
    # Specify the DC and domain in the Base of the query.
    $Base = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$Server/$BaseDN"
    $Searcher.SearchRoot = $Base
    $Result = $Searcher.FindOne()
    If ($Result)
    {
        # Retrieve the values.
        $DN = $Result.Properties.Item("distinguishedName")
        $NTName = $Result.Properties.Item("sAMAccountName")
        $BadCount = $Result.Properties.Item("badPwdCount")
        $LogonCount = $Result.Properties.Item("logonCount")
        $Time = $Result.Properties.Item("badPasswordTime")
        $Last = $Result.Properties.Item("lastLogon")
        $Pwd = $Result.Properties.Item("pwdLastSet")
        $Lock = $Result.Properties.Item("lockoutTime")
        # Convert LargeInteger values into datetime values in the local time zone.
        If ($Time.Count -eq 0) {$BadTime = "Never"}
        Else
        {
            If ($Time.Item(0) -eq 0) {$BadTime = "Never"}
            Else {$BadTime = ([DateTime]$Time.Item(0)).AddYears(1600).ToLocalTime()}
        }
        If ($Last.Count -eq 0) {$LastLogon = "Never"}
        Else
        {
            If ($Last.Item(0) -eq 0) {$LastLogon = "Never"}
            Else {$LastLogon = ([DateTime]$Last.Item(0)).AddYears(1600).ToLocalTime()}
        }
        If ($Pwd.Count -eq 0) {$PwdLastSet = "Never"}
        Else
        {
            If ($Pwd.Item(0) -eq 0) {$PwdLastSet = "Never"}
            Else {$PwdLastSet = ([DateTime]$Pwd.Item(0)).AddYears(1600).ToLocalTime()}
        }
        If ($Lock.Count -eq 0) {$Lockout = "Never"}
        Else
        {
            If ($Lock.Item(0) -eq 0) {$Lockout = "Never"}
            Else {$Lockout = ([DateTime]$Lock.Item(0)).AddYears(1600).ToLocalTime()}
        }
        # Output in comma delimited format.
        "$NTName,$PwdLastSet,$Lockout,$ServerName,$LastLogon,$LogonCount,$BadCount,$BadTime"
    }
    Else
    {
        Write-Host "ERROR: Failed to connect to DC $Server" -foregroundcolor red -backgroundcolor black
        ",,,$ServerName,,,,"
    }
}

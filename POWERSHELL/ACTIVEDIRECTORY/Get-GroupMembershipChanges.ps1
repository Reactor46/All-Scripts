<#
.SYNOPSIS
  Name: Get-GroupMembershipChanges.ps1
  The purpose of this script is to monitor and inform you for membership changes on groups

.DESCRIPTION
  This is a simple script that checks for changes of the members of groups. The script can run as one off
  or it can be configured to run a scheduled basis to monitor and inform you for group membership changes.

.RELATED LINKS
  https://www.sconstantinou.com

.PARAMETER OrganizationalUnit
  This is the only parameter for the script. It is used to define the Organizational Unit in Active Directory
  that you want to run the script. You need to provide the DistinguishedName of the Organizational Unit.
  The default value is "".

.NOTES
    Release Date: 10-08-2018

  Author: Stephanos Constantinou

.EXAMPLE
  Run the Get-GroupMembershipChanges.ps1 script without any parameter to run it for the entire domain.
  Get-GroupMembershipChanges.ps1

.EXAMPLE
  Run the Get-GroupMembershipChanges.ps1 script with Organizational Unit parameter to run it on specific
  Organizational Unit in Active Directory.
  Get-GroupMembershipChanges.ps1 -OrganizationalUnit "OU=Groups,DC=Domain,DC=com"
#>

param (
    [string]$OrganizationalUnit = ""
    )


Import-Module ActiveDirectory

#$PasswordFile = "C:\Scripts\Password.txt"
#$Key = (1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32)
#$EmailUser = "Script-User@domain.com"
#$Password = Get-Content $PasswordFile | ConvertTo-SecureString -Key $Key
#$EmailCredentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $EmailUser,$Password
$To = 'john.battista@creditone.com'#,'User2@domain.com'
$From = 'GroupMembershipChanges@creditone.com'
$Path = "C:\LazyWinAdmin\ActiveDirectory\Logs"
$Date = Get-Date -format dd-MM-yyyy-HH-mm-ss
$LogFile = "$Path\log-$Date .log"
$EmailResult = ""

$EmailUp = @"
<style>

body { font-family:Segoe, "Segoe UI", "DejaVu Sans", "Trebuchet MS", Verdana, sans-serif !important; color:#434242;}
TABLE { font-family:Segoe, "Segoe UI", "DejaVu Sans", "Trebuchet MS", Verdana, sans-serif !important; border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TR {border-width: 1px;padding: 10px;border-style: solid;border-color: white; }
TD {font-family:Segoe, "Segoe UI", "DejaVu Sans", "Trebuchet MS", Verdana, sans-serif !important; border-width: 1px;padding: 10px;border-style: solid;border-color: white; background-color:#C3DDDB;}
.colorm {background-color:#58A09E; color:white;}
.colort{background-color:#58A09E; padding:20px; color:white; font-weight:bold;}
.colorn{background-color:transparent;}
</style>
<body>

<h3>Script has been completed successfully</h3>

<h4>Changes applied:</h4>

<table>
    <tr>
   	    <td class="colort">User</td>
        <td class="colort">Action</td>
        <td class="colort">Group</td>
    </tr>
"@

$EmailDown = @"
</table>
</body>
"@

Function Write-Log{
    Param ([string]$LogDetails)

    Add-content $Logfile -value $LogDetails
}

if ($OrganizationalUnit -eq ""){
    $Groups = (Get-ADGroup -Filter *).Name
}
else{
    try{$Groups = (Get-ADGroup -Filter * -SearchBase "$OrganizationalUnit").Name}
    catch{
        $TimeStamp = Get-Date
        $LogDetails = "$TimeStamp" + " " + "$_"
        Write-Log $LogDetails
        Exit}
}

Foreach ($Group in $Groups){

    $oldmembers = $null
    $newmembers = $null
    $file = "$Path" + "$Group" + ".csv"

    try{$oldmembers = (Import-Csv $file).SamAccountName}
    catch{
        $TimeStamp = Get-Date
        $LogDetails = "$TimeStamp" + " " + "$_"
        Write-Log $LogDetails}

    try{Get-ADGroupMember -Identity $Group |
            Select-Object SamAccountName |
                Export-Csv $file -NoTypeInformation}
    catch{
        $TimeStamp = Get-Date
        $LogDetails = "$TimeStamp" + " " + "$_"
        Write-Log $LogDetails}

    $newmembers = (Import-Csv $file).SamAccountName

    switch -Regex ($oldmembers){
        {($oldmembers -ne $null) -and ($newmembers -ne $null)}{

            $Difference = Compare-Object $oldmembers $newmembers

            if ($Difference -ne ""){

                Foreach ($DifferenceValue in $Difference){

                    $DifferenceValueIndicator = $DifferenceValue.SideIndicator

                    switch -Regex ($DifferenceValueIndicator){
                        {$_ -eq "<="}{
                            $GroupMember = $DifferenceValue.InputObject
                            $Action = "Removed"
                            $EmailTemp = @"
    <tr>
   	    <td class="colorm">$GroupMember</td>
        <td>$Action</td>
        <td>$Group</td>
    </tr>
"@
                            $EmailResult = $EmailResult + $EmailTemp
                            $TimeStamp = Get-Date
                            $LogDetails = "$TimeStamp $GroupMember has been $Action from $Group"
                            Write-Log $LogDetails
                        }
                        {$_ -eq "=>"}{
                            $GroupMember = $DifferenceValue.InputObject
                            $Action = "Added"
                            $EmailTemp = @"
    <tr>
   	    <td class="colorm">$GroupMember</td>
        <td>$Action</td>
        <td>$Group</td>
    </tr>
"@
                            $EmailResult = $EmailResult + $EmailTemp
                            $TimeStamp = Get-Date
                            $LogDetails = "$TimeStamp $GroupMember has been $Action to $Group"
                            Write-Log $LogDetails
                        }
                    }
                }
            }
            Break
        }
        {($oldmembers -eq $null) -and ($newmembers -ne $null)}{

            Foreach ($newmember in $newmembers){
                $Action = "Added"
                $EmailTemp = @"
    <tr>
   	    <td class="colorm">$newmember</td>
        <td>$Action</td>
        <td>$Group</td>
    </tr>
"@
                $EmailResult = $EmailResult + $EmailTemp
                $TimeStamp = Get-Date
                $LogDetails = "$TimeStamp $newmember has been $Action to $Group"
                Write-Log $LogDetails
            }
            Break
        }
        {($oldmembers -ne $null) -and ($newmembers -eq $null)}{

            Foreach ($oldmember in $oldmembers){
                $Action = "Removed"
                $EmailTemp = @"
    <tr>
   	    <td class="colorm">$oldmember</td>
        <td>$Action</td>
        <td>$Group</td>
    </tr>
"@
                $EmailResult = $EmailResult + $EmailTemp
                $TimeStamp = Get-Date
                $LogDetails = "$TimeStamp $oldmember has been $Action from $Group"
                Write-Log $LogDetails
            }
            Break
        }
    }
}

$Email = $EmailUp + $EmailResult + $EmailDown

if ($EmailResult -ne ""){

    $EmailParameters = @{
        To = $To
        Subject = "Group Membership Changes Report $(Get-Date -format dd/MM/yyyy)"
        Body = $Email
        BodyAsHtml = $True
        UseSsl = $False
        Port = "2525"
        SmtpServer = "mailgateway.Contoso.corp"
        From = $From}

    send-mailmessage @EmailParameters
}
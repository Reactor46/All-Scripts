<#
    .SYNOPSIS
    Creates a HTML Report showing Distribution and Dynamic Distribution Group Summary and Members 
   
   	Serkan Varoglu
	
	http:\\Get-Mailbox.org
	http:\\Mshowto.org
	@SRKNVRGL
	
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
	Version 1.0, 24 March 2012
	
    .DESCRIPTION
	
    This script creates a HTML report showing the following information:
	
	* Distribution Group Summary:
		o Distribution Group Count
		o Hidden From Address List
		o Moderation Status
		o Sender Authentication Required
		o Join
			* Open to Join
			* Approval to Join
			* Closed to Join
		o Depart
			* Open to Depart
			* Closed to Depart
		o Is Valid False
		o No Manager
		o Empty Groups
    
	Group Membership
	* Distribution Group Name
		o If Group is hidden from Address Lists this cell will have a DASHED BLACK BORDER
	* Join (Column J)
		o If Group is Open to join this cell will be highlighted with color GREEN
		o If Group is ApprovalRequired to join this cell will be highlighted with color ORANGE
	* Depart (Column D)
		o If Group is Open to depart this cell will have a DOTTED RED BORDER
	* Recipient Type
		o If the group is not valid this cell will be highligted with color YELLOW
	* Primary SMTP Address
		o If Sender Authentication Required this cell will have a SOLID RED BORDER
	* Member of
	
	Empty Groups
	This table will show Groups with no members
	
	Duplicate Recipient
	This table will show the recipients who are duplicate as a member of this group. It will also show if you have Distribution Groups for Duplicate membership.
	
	Recipient List
	This Table will show all unique recipients. Distribution Groups will NOT be shown.

	Also please be aware that if have Dynamic Distribution Groups they will be discarded in this report.

	If you migrated from a previous version of Exchange Server some groups might have invalid alias. These will be listed during the process as warning and also in the report you can find these groups as “Is valid false” highlighted with light yellow background.
	
	.PARAMETER Name 
	Distribution Group Name that you want to report.
		
	.PARAMETER ReportName
    Filename to write HTML Report to.
	
	!! If you do not use ReportName parameter the report will be created in the directory that you ran this powershell script.
	    
	.EXAMPLE
    Generate the HTML report 
    .\Report-DistributionGroupMember.ps1 "IT ADMINS GROUP"
	
#>
param ( [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true,HelpMessage='Distribution Group Name')][string]$Name,[Parameter(Position=1,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Report Name')][string]$ReportName)
$ErrorActionPreference = "silentlycontinue"
$Watch = [System.Diagnostics.Stopwatch]::StartNew()
#Initialize
$DistGroup = $ADistGroup = Get-DistributionGroup $Name -ErrorAction "silentlycontinue"
if(!$DistGroup)
{
"Could not retrieve any information from your input. Please make sure Distribution Group Name is correct."
exit
}
$DistGroupMembers = Get-DistributionGroupMember $DistGroup -Resultsize Unlimited
$NestedGroups = @{}
$NestedGroups.Add($DistGroup.Name,"PARENT GROUP")
$MemberCount = @{}
$DuplicateMemberGroup=$Null
function _Progress
{
    param($PercentComplete,$Status)
    Write-Progress -id 1 -activity "Report for Distribution Group Members" -status $Status -percentComplete ($PercentComplete)
}
_Progress (20) "Collecting Distribution Group Information"

#Function to Collect Member Distribution Groups
function collect($DistGroupMembers)
{
$NestedGroup = @{}
if ($DistGroupMembers)
        {    
            foreach ($member in $DistGroupMembers)
            {
                if(($member.RecipientTypeDetails -like "*Group*") -and ($member.RecipientTypeDetails -notlike "*Dynamic*"))
                {
                    $name = $member.SamAccountName.ToString()
                    if ($NestedGroup.ContainsKey($name) -eq $false)
                    {
                        $NestedGroup.Add($name,$member.DisplayName.ToString())
						$NestedGroups.Add($name,$DistGroup.Name.ToString())
                    }
                }    
            }
        }
		if($NestedGroup.Values.Count -gt 0)
		{
			foreach ($NestedGroup in $NestedGroup.values)
            {
				$DistGroup = Get-DistributionGroup $NestedGroup
				$Nest=Get-DistributionGroupMember $NestedGroup -Resultsize Unlimited
				collect $Nest
            }
			
		}
$global:NestedGroups=$NestedGroups
}
$AllDistGroups=@()
$EmptyGroup=@()

#Run Function to Collect Member Distribution Groups
collect $DistGroupMembers
_Progress (40) "Collecting Member Information for Distribution Group"
#Collect Information for Each Group and Members
foreach ($DistGroup in $NestedGroups.keys)
{
$DistGroup = Get-DistributionGroup $DistGroup
$AllDistGroups += $DistGroup
$srknvrgl=Get-DistributionGroupMember $DistGroup -Resultsize Unlimited
if ($srknvrgl)
{
	foreach ($srkn in $srknvrgl)
	{
	$srkn | Add-Member -type NoteProperty -name GroupName -value $DistGroup.SamAccountName
	$srkn | foreach {$MemberCount[$_.SamaccountName] += 1}
	}
}
else
{
$EmptyGroup+=$DistGroup
}
$AllMembers+=$srknvrgl
}

#Get Duplicate Member Group Membership Information
foreach ($DuplicateMember in $MemberCount.keys)
{
	if ($MemberCount.$DuplicateMember -gt 1)
	{
	$DuplicateMemberGroup+=$AllMembers | ?{$_.SamAccountName -like $DuplicateMember}
	}
}
_Progress (80) "Compiling Report"
#Start HTML Output

$Output=$Null
$Output="<body><font size=""1"" face=""Arial,sans-serif""><h3 align=""center"">$($ADistGroup) Distribution Group Membership Report</h3><h4 align=""center""><a name=""top"">Generated $((Get-Date).ToString())</a></h4><br><table cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"" width=""100%"">"

#Add Index
if ($AllDistGroups)
{
$Output+="<tr><th><a href=""#GroupSummary"">Distribution Group Configuration Summary</a></th></tr>"
}
if($NestedGroups)
{
$Output+="<tr><th><a href=""#GroupMembership"">Group Membership</a></th></tr>"
}
if($EmptyGroup)
{
$Output+="<tr><th><a href=""#EmptyGroups"">Empty Groups</a></th></tr>"
}
if ($DuplicateMemberGroup)
{
$Output+="<tr><th><a href=""#DuplicateRecipient"">Duplicate Recipient</a></th></tr>"
}
if($AllMembers)
{
$Output+="<tr><th><a href=""#RecipientList""> Recipient List</a></th></tr>"
}
$Output+="</table><br>"

#Add Group Configuration Summary
if ($AllDistGroups)
{
	$hiddengroups=($AllDistGroups | ?{$_.HiddenFromAddressListsEnabled -like "True"} | measure-object).count
	$moderatedgroups=($AllDistGroups | ?{$_.ModerationEnabled -like "True"} | measure-object).count
	$authgroups=($AllDistGroups | ?{$_.RequireSenderAuthenticationEnabled -like "True"} | measure-object).count
	$invalidgroups=($AllDistGroups | ?{$_.isvalid -like "False"} | measure-object).count
	$opengroups=($AllDistGroups | ?{$_.MemberJoinRestriction -like "Open"} | measure-object).count
	$approvalgroups=($AllDistGroups | ?{$_.MemberJoinRestriction -like "Approval*"} | measure-object).count
	$closedgroups=($AllDistGroups | ?{$_.MemberJoinRestriction -like "Closed"} | measure-object).count
	$departopengroups=($AllDistGroups | ?{$_.MemberDepartRestriction -like "Open"} | measure-object).count
	$departclosedgroups=($AllDistGroups | ?{$_.MemberDepartRestriction -like "Closed"} | measure-object).count
	$groupswithoutManager=($AllDistGroups | ?{!$_.managedby} | measure-object).count
	$Output+="<table border=""1"" bordercolor=""#1F497B"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"" width=""100%"">
	<tr bgcolor=""#1F497B"" align=""""center"""">
	<th nowrap=""nowrap"" rowspan=""2""><font color=""#FFFFFF""><a name=""GroupSummary"">Distribution Group Count</a></th>
	<th nowrap=""nowrap"" rowspan=""2""><font color=""#FFFFFF"">Hidden From Address List</th>
	<th nowrap=""nowrap"" rowspan=""2""><font color=""#FFFFFF"">Moderated</th>
	<th nowrap=""nowrap"" rowspan=""2""><font color=""#FFFFFF"">Sender Authentication Required</th>
	<th nowrap=""nowrap"" colspan=""3""><font color=""#FFFFFF"">Join</th>
	<th nowrap=""nowrap"" colspan=""2""><font color=""#FFFFFF"">Depart</th>
	<th nowrap=""nowrap"" rowspan=""2""><font color=""#FFFFFF"">Is Valid False</th>
	<th nowrap=""nowrap"" rowspan=""2""><font color=""#FFFFFF"">No Manager</th>
	<th nowrap=""nowrap"" rowspan=""2""><font color=""#FFFFFF"">Empty Groups</th>
	</tr>
	<tr bgcolor=""#1F497B"" align=""""center"""">
	<th nowrap=""nowrap""><font color=""#FFFFFF"">Open to Join</th>
	<th nowrap=""nowrap""><font color=""#FFFFFF"">Approval to Join</th>
	<th nowrap=""nowrap""><font color=""#FFFFFF"">Closed to Join</th>
	<th nowrap=""nowrap""><font color=""#FFFFFF"">Open to Depart</th>
	<th nowrap=""nowrap""><font color=""#FFFFFF"">Closed to Depart</th>
	</tr>
	<tr>
	<td align=""center"">$(($AllDistGroups).count)</td>
	<td style=""border: 2px dashed black"" align=""center"">$hiddengroups</td>
	<td bgcolor=""#FFF200"" align=""center"">$moderatedgroups</td>
	<td style=""border: 1px solid red"" align=""center"">$authgroups</td>
	<td bgcolor=""#A4FFA4"" align=""center"">$opengroups</td>
	<td bgcolor=""#FFB366"" align=""center"">$approvalgroups</td>
	<td align=""center"">$closedgroups</td>
	<td style=""border: 2px dotted red"" align=""center"">$departopengroups</td>
	<td align=""center"">$departclosedgroups</td>
	<td bgcolor=""#FFFFB3"" align=""center"">$invalidgroups</td>
	<td align=""center"">$groupswithoutManager</td>
	<td align=""center"">$(($EmptyGroup).count)</td>
	</tr></table><br><a href=""#top"">&#9650;</a>"
}

#Add Nested Group Information
if($NestedGroups)
{
	$Output+="<table border=""1"" bordercolor=""#1F497B"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"" width=""100%""><tr bgcolor=""#1F497B"" align=""center""><th colspan=""6""><font color=""#FFFFFF""><a name=""GroupMembership"">Group Membership</a></font></th></tr><tr bgcolor=""#1F497B""><th><font color=""#FFFFFF"">Group Name</font></th><th><font color=""#FFFFFF"">J</font></th><th><font color=""#FFFFFF"">D</font></th><th><font color=""#FFFFFF"">SMTP Address</font></th><th><font color=""#FFFFFF"">Recipient Type</font></th><th><font color=""#FFFFFF"">Member Of</font></th></tr>"
	foreach ($group in $NestedGroups.keys)
	{
	$groupdata=Get-DistributionGroup $group
	if ($groupdata.ModerationEnabled -like "True")
	{$bgcolor="bgcolor=""#FFF200"""}
	else
	{$bgcolor=""}
	if ($groupdata.HiddenFromAddressListsEnabled -eq $True)
	{$border="style=""border: 2px dashed black"""}
	else
	{$border=""}
	$Output+="<tr><th $($bgcolor) $($border)>$($groupdata.DisplayName)</th>"
	if ($groupdata.MemberJoinRestriction -like "Open")
	{$bgcolor="bgcolor=""#A4FFA4"""}
	elseif ($groupdata.MemberJoinRestriction -like "Approval*")
	{$bgcolor="bgcolor=""#FFB366"""}
	else
	{$bgcolor=""}
	if ($groupdata.MemberDepartRestriction -like "Open")
	{$border="style=""border: 2px dotted red"""}
	else
	{$border=""}
	$Output+="<th $($bgcolor)>&nbsp;</th><th $($border)>&nbsp;</th>"
	if ($groupdata.RequireSenderAuthenticationEnabled -like "True")
	{$border="style=""border: 1px solid red"""}
	else
	{$border=""}
	$Output+="<td $($border)>$($groupdata.PrimarySmtpAddress)</td>"
	if ($groupdata.isvalid -notlike "True")
	{
	$Output+="<td bgcolor=""#FFFFB3"">$($groupdata.RecipientType)</td>"
	}
	else
	{
	$Output+="<td>$($groupdata.RecipientType)</td>"
	}
	$Output+="<th>$($NestedGroups.$group)</th></tr>"
	}
	$Output+="</table><br><a href=""#top"">&#9650;</a>"
}

#Add Empty Group Information
if($EmptyGroup)
{
	$Output+="<table border=""1"" bordercolor=""#1F497B"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"" width=""100%""><tr bgcolor=""#1F497B"" align=""center""><th colspan=""2""><font color=""#FFFFFF""><a name=""EmptyGroups"">Empty Distribution Groups</a></font></th></tr>"
	foreach($Empty in $EmptyGroup)
	{
		$Output+="<tr bgcolor=""#1F497B""><th><font color=""#FFFFFF"">Distribution Group Name</font></th><th><font color=""#FFFFFF"">Recipient Type</font></th></tr><tr><th>$($Empty.Name)</th><th>$($Empty.RecipientType)</th></tr>"
	}
	$Output+="</table><br><a href=""#top"">&#9650;</a>"
}

#Add Duplicate User Table
if ($DuplicateMemberGroup)
{
	$Output+="<table border=""1"" bordercolor=""#1F497B"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"" width=""100%""><tr bgcolor=""#1F497B"" align=""center""><th colspan=""3""><font color=""#FFFFFF""><a name=""DuplicateRecipient"">Duplicate Recipient</a></font></th></tr><tr bgcolor=""#1F497B""><th><font color=""#FFFFFF"">Name</font></th><th><font color=""#FFFFFF"">Recipient Type</font></th><th><font color=""#FFFFFF"">Member Of</font></th></tr>"
	foreach($DuplicateMember in $DuplicateMemberGroup)
	{
		$Output+="<tr><th>$($DuplicateMember.DisplayName)</th><th>$($DuplicateMember.RecipientType)</th><th>$($DuplicateMember.GroupName)</th></tr>"
	}
	$Output+="</table><br><a href=""#top"">&#9650;</a>"
}

#Add  User Table
if ($AllMembers)
{
	$Users=$AllMembers | ?{$_.RecipientType -notlike "*Group*"} | Sort -uniq
	$Output+="<table border=""1"" bordercolor=""#1F497B"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"" width=""100%""><tr bgcolor=""#1F497B"" align=""center""><th colspan=""6""><font color=""#FFFFFF""><a name=""RecipientList""> Recipient List</a></font></th></tr><tr bgcolor=""#1F497B"" align=""center""><th><font color=""#FFFFFF"">Name</font></th><th><font color=""#FFFFFF"">SMTP Address</font></th><th><font color=""#FFFFFF"">Member of</font></th><th><font color=""#FFFFFF"">Recipient Type</font></th><th><font color=""#FFFFFF"">Valid</font></th></tr>"
	foreach($User in $Users)
	{
	$Output+="<tr><th>$($User.DisplayName)</th>"
	if ($User.RecipientType -like "*contact*")
	{
	$Contact=Get-Contact $User.Name | Select-Object WindowsEmailAddress
	$Output+="<th>$($Contact.windowsemailaddress)</th>"
	}
	else
	{
	$Output+="<th>$($User.PrimarySmtpAddress)</th>"
	}
	$Output+="<th>$($User.GroupName)</th><th>$($User.RecipientType)</th>"
	if($User.isvalid -eq $false){$Output+="<th>Not Valid</th></tr>"}else{$Output+="<th>Yes</th></tr>"}
	}
	$Output+="</table><br><a href=""#top"">&#9650;</a>"
}
else
{
$Output+="<table border=""1"" bordercolor=""#000000"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"" width=""100%""><tr align=""center""><th>This Group Does Not Have Any Recipients</th></tr></table>"
}

#Close HTML Output
$Output+="<br><font size=""1"" face=""Arial,sans-serif"">Scripted by <a href=""http://www.get-mailbox.org"">Serkan Varoglu</a>.  
	Elapsed Time To Complete This Report: $($Watch.Elapsed.Minutes.ToString()):$($Watch.Elapsed.Seconds.ToString())</font></body></html>"


#Output HTML Report
$t = Get-Date -UFormat %d%m%H%M
if(!$ReportName)
{
$HTMLReport = "$($ADistGroup.windowsemailaddress.local)-$($t).html"
"Report is ready. $HTMLReport"
}
else
{
$HTMLReport=$ReportName
"Report is ready. $HTMLReport"
}
$Output | Out-File $HTMLReport


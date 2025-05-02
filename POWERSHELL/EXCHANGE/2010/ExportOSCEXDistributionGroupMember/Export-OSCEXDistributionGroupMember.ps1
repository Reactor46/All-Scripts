<#
The sample scripts are not supported under any Microsoft standard support 
program or service. The sample scripts are provided AS IS without warranty  
of any kind. Microsoft further disclaims all implied warranties including,  
without limitation, any implied warranties of merchantability or of fitness for 
a particular purpose. The entire risk arising out of the use or performance of  
the sample scripts and documentation remains with you. In no event shall 
Microsoft, its authors, or anyone else involved in the creation, production, or 
delivery of the scripts be liable for any damages whatsoever (including, 
without limitation, damages for loss of business profits, business interruption, 
loss of business information, or other pecuniary loss) arising out of the use 
of or inability to use the sample scripts or documentation, even if Microsoft 
has been advised of the possibility of such damages.
#> 

#requires -Version 2

#.EXTERNALHELP Export-OSCEXDistributionGroupMember.xml
[CmdletBinding(DefaultParameterSetName="Filter")]
Param
(
	#Define parameters
	[Parameter(Mandatory=$true,Position=1,ParameterSetName="Filter")]
	[string]$Filter,
	[Parameter(Mandatory=$true,Position=1,ParameterSetName="Identity")]
	[string]$Identity,
	[Parameter(Mandatory=$false,Position=2)]
	[string[]]$RecipientProperty=@("Name","RecipientType"),
	[Parameter(Mandatory=$true,Position=3)]
	[string]$ReportFolder,
	[Parameter(Mandatory=$false)]
	[switch]$Recurse,
	[Parameter(Mandatory=$false)]
	[switch]$MergeReport
)

#Import Localized Data
Import-LocalizedData -BindingVariable Messages
#A Hash table to track all processed groups
$nestedGroups = @{}
#An array to save all nested group members
$nestedGroupMembers = @()

Function New-OSCPSCustomErrorRecord
{
	#This function is used to create a PowerShell ErrorRecord
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true,Position=1)][String]$ExceptionString,
		[Parameter(Mandatory=$true,Position=2)][String]$ErrorID,
		[Parameter(Mandatory=$true,Position=3)][System.Management.Automation.ErrorCategory]$ErrorCategory,
		[Parameter(Mandatory=$true,Position=4)][PSObject]$TargetObject
	)
	Process
	{
		$exception = New-Object System.Management.Automation.RuntimeException($ExceptionString)
		$customError = New-Object System.Management.Automation.ErrorRecord($exception,$ErrorID,$ErrorCategory,$TargetObject)
		return $customError
	}
}

Function Get-OSCEXNestedDistributionGroupMember
{
	[CmdletBinding()]
	Param
	(
		#Define parameters
		[Parameter(Mandatory=$true,Position=1)]
		[string]$Identity
	)
	Process
	{
		$dg = Get-DistributionGroup -Identity $Identity -Verbose:$false
		if (-not $nestedGroups.Contains($dg.SamAccountName)) {
			$nestedGroups.Add($dg.SamAccountName,"")
		}	
		$dgMembers = Get-DistributionGroupMember -Identity $Identity -ResultSize unlimited -Verbose:$false
		if ($dgMembers -ne $null) {
			foreach ($dgMember in $dgMembers) {
				if ($dgMember.RecipientTypeDetails -eq "MailUniversalDistributionGroup") {
					if (-not $nestedGroups.Contains($dgMember.SamAccountName)) {
						Get-OSCEXNestedDistributionGroupMember -Identity $dgMember
					}
				} else {
					$nestedGroupMembers += $dgMember
				}
			}
		}
		return $nestedGroupMembers
	}
}

if (-not (Test-Path -Path $ReportFolder -PathType Container)) {
	$errorMsg = $Messages.InvalidReportFolder
	$errorMsg = $errorMsg -f $ReportFolder
	$customError = New-OSCPSCustomErrorRecord `
	-ExceptionString $errorMsg `
	-ErrorCategory NotSpecified -ErrorID 1 -TargetObject $pscmdlet
	$pscmdlet.ThrowTerminatingError($customError)
}
#Try to get distribution groups
Try
{
	Switch ($pscmdlet.ParameterSetName) {
		"Filter" {
			$distributionGroups = Get-DistributionGroup -Filter $Filter -ResultSize unlimited -Verbose:$false
		}
		"Identity" {
			$distributionGroups = Get-DistributionGroup -Identity $Identity -Verbose:$false
		}
	}
}
Catch
{
	$pscmdlet.ThrowTerminatingError($Error[0])
}
#Try to get distribution group members
if ($distributionGroups -ne $null) {
	$emptyDGReports = @()
	$recipientPropertyNames = @{}
	foreach ($distributionGroup in $distributionGroups) {
		$distributionGroupAlias = $distributionGroup.Alias
		if (-not $Recurse) {
			$dgMembers = Get-DistributionGroupMember -Identity $distributionGroupAlias -ResultSize unlimited -Verbose:$false
		} else {
			$dgMembers = Get-OSCEXNestedDistributionGroupMember -Identity $distributionGroupAlias
			#Reset variables which defined in script scope
			$nestedGroupMembers = @()
			$nestedGroups.Clear()
		}
		$dgMembers = $dgMembers | Select-Object -Unique
		$verboseMsg = $Messages.GetDGMember
		if ($dgMembers -ne $null) {
			$reports = @()
			$dgMembersCount = ($dgMembers | Measure-Object).Count
			$verboseMsg = $verboseMsg -f $distributionGroupAlias,$dgMembersCount
			$pscmdlet.WriteVerbose($verboseMsg)
			foreach ($dgMember in $dgMembers) {
				$report = New-Object -TypeName System.Management.Automation.PSObject
				$report | Add-Member -MemberType NoteProperty -Name "DistributionGroupAlias" -Value $distributionGroupAlias
				#Handle customized properties
				$recipientProperties = $dgMember | Select-Object -Property $RecipientProperty
				if ($recipientPropertyNames.Count -eq 0) {
					$recipientProperties | Get-Member -MemberType NoteProperty | %{$recipientPropertyNames.Add($_.Name,"")}
				}
				foreach ($recipientPropertyName in $recipientPropertyNames.Keys.GetEnumerator()) {
					if ($($recipientProperties.$recipientPropertyName) -is [array]) {
						$report | Add-Member -MemberType NoteProperty `
						-Name $recipientPropertyName -Value $($($recipientProperties.$recipientPropertyName) -join ";")
					} else {
						$report | Add-Member -MemberType NoteProperty -Name $recipientPropertyName -Value $($recipientProperties.$recipientPropertyName)
					}
				}
				$reports += $report
			}
			#Sort report column
			$reports = $reports | Sort-Object
			#Generate report per distribution group basis
			$reports | Export-Csv -Path "$ReportFolder\DistributionGroupMembers - $distributionGroupAlias.csv" -NoTypeInformation
		} else {
			$verboseMsg = $verboseMsg -f $distributionGroupAlias,"0"
			$pscmdlet.WriteVerbose($verboseMsg)
			$emptyDGReport = New-Object -TypeName System.Management.Automation.PSObject
			$emptyDGReport | Add-Member -MemberType NoteProperty -Name "DistributionGroupAlias" -Value $distributionGroupAlias
			$emptyDGReports += $emptyDGReport
		}
	}
	$emptyDGReports | Export-Csv -Path "$ReportFolder\EmptyDistributionGroups.csv" -NoTypeInformation
	if ($MergeReport) {
		$reportFiles = Get-ChildItem -Path $ReportFolder -Filter "DistributionGroupMembers*.csv"
		$reportFiles | Import-Csv | Export-Csv -Path "$ReportFolder\MergedReport.csv" -NoTypeInformation
	}
}

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

#Import Localized Data
Import-LocalizedData -BindingVariable Messages

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

Function New-OSCSPSiteGroup
{
	#.EXTERNALHELP New-OSCSPSiteGroup-Help.xml
	
	[CmdletBinding()]
	Param
	(	
		[Parameter(Mandatory=$true,Position=1)]
		[string]$SiteURL,
		[Parameter(Mandatory=$true,Position=2)]
		[string]$Name,
		[Parameter(Mandatory=$true,Position=3)]
		[string]$Description,
		[Parameter(Mandatory=$true,Position=4)]
		[string]$Owner,
		[Parameter(Mandatory=$false,Position=5)]
		[string[]]$Member,
		[Parameter(Mandatory=$false,Position=6)]
		[string]$PermissionLevel
	)
	Process
	{
		#Try to use Get-SPSite to get an instance of Microsoft.SharePoint.SPSite
		Try
		{
			$spSite = Get-SPSite -Identity $SiteURL -Verbose:$false -ErrorAction Stop
			$spWeb = $spSite.Rootweb
		}
		Catch
		{
			$pscmdlet.ThrowTerminatingError($_)
		}
		
		#Check whether the specified login name belongs to a valid user of the Web site,
		#and if the login name does not already exist, adds it to the Web site.
		Try
		{
			$spGroupOwner = $spWeb.EnsureUser($Owner)
		}
		Catch
		{
			$pscmdlet.ThrowTerminatingError($_)
		}

		#Verify specified Permission Level
		if (-not [System.String]::IsNullOrEmpty($PermissionLevel)) {
			#Verify user specified permission level
			$permissionLevelNames = @()
			$spWeb.RoleDefinitions | %{$permissionLevelNames += $_.Name}
			$permissionLevelNamesString = $permissionLevelNames -join ","
			if ($permissionLevelNames -notcontains $PermissionLevel) {
				$errorMsg = $Messages.InvalidPermissionLevel
				$errorMsg = $errorMsg -f $PermissionLevel,$permissionLevelNamesString
				$customError = New-OSCPSCustomErrorRecord `
				-ExceptionString $errorMsg `
				-ErrorCategory NotSpecified -ErrorID 1 -TargetObject $pscmdlet
				$pscmdlet.ThrowTerminatingError($customError)
			}
		}
		
		#Create a SharePoint Group
        	Try
    		{
			$verboseMsg = $Messages.CreateGroup
			$pscmdlet.WriteVerbose($verboseMsg)
			$spWeb.SiteGroups.Add($Name,$spGroupOwner,$spGroupOwner,$Description)
			$spGroup = $spWeb.SiteGroups.Item($Name)
		}
    		Catch
        	{
			$spWeb.Dispose()
			$spSite.Dispose()        
    	    		$pscmdlet.ThrowTerminatingError($_)
    		}
		
		#Add group members
		if (-not [System.String]::IsNullOrEmpty($Member)) {
			foreach ($groupMember in $Member)	{
				Try
				{
					$spGroupMember = $spWeb.EnsureUser($groupMember)
					$spGroup.AddUser($spGroupMember)
					$spGroup.Update()
				}
				Catch
				{
					$pscmdlet.WriteError($_)
				}
			}
		}
		
		if (-not [System.String]::IsNullOrEmpty($PermissionLevel)) {
			$verboseMsg = $Messages.PermissionLevelPrompt
			$verboseMsg = $verboseMsg -f $spGroup,$PermissionLevel
			$pscmdlet.WriteVerbose($verboseMsg)
			#Get role definition item
			$spRoleDefinition = $spWeb.RoleDefinitions.Item($PermissionLevel)				
			#Create a role assigment object
			$spRoleAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($spGroup)
			$spRoleAssignment.RoleDefinitionBindings.Add($spRoleDefinition)
			$spWeb.RoleAssignments.Add($spRoleAssignment)
		} else {
			#If $PermissionLevel parameter is omitted, this function will create the group only.
			$verboseMsg = $Messages.EmptyPermissionLevelParameter
			$pscmdlet.WriteVerbose($verboseMsg)
		}
		
		$spWeb.Update()
		$spWeb.Dispose()
		$spSite.Dispose()
	}
}

Export-ModuleMember -Function "New-OSCSPSiteGroup"
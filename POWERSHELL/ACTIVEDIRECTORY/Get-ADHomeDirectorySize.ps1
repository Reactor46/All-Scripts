<#
.SYNOPSIS
Returns the Home Directory size of each specified user.
.DESCRIPTION
Get-ADHomeDirectorySize takes a user and queries AD for that user's homeDirectory and then calculates the size of that directory.
.PARAMETER $user
The user account(s) that you would like to know their home directory size
.EXAMPLE
Get-ADHomeDirectorySize Twon.of.An

This will return the size and location of Twon of An's homeDirectory
.EXAMPLE
Get-ADUser -filter * -properties homeDirectory | ? {$_.homedirectory -ne $nul} | Get-ADHomeDirectorySize

This will go through each and every user in AD (that has a homeDirectory) and look for its size.
.INPUTS
Microsoft.ActiveDirectory.Management.ADUser
.OUTPUTS
A custom object consisting of the Users name, HomeDirectory, and HomeDirectorySize.
.NOTES
Author: Town of An
.LINK
Get-AdUser
Get-ChildItem
Measure-Object
#>
Function Get-ADHomeDirectorySize
{
	Param
	(
		[Parameter(ValueFromPipeline=$true,Mandatory=$true)]
		[Microsoft.ActiveDirectory.Management.ADUser]$User
	)
	Begin
	{
		$HomeD = @()
		$size = $nul
	}
	Process
	{
		ForEach($userAccount in  $User)
		{
			$userAccount = Get-ADUser $userAccount -properties homeDirectory
			$size = Get-ChildItem $userAccount.homeDirectory -recurse | Measure-Object -property length -sum
			If($userAccount.homeDirectory -eq $nul)
			{
				Write-Host "`nERROR -- User: $userAccount has no Home Directory`n" -foregroundcolor red
				Return
			}
			$obj = New-Object System.Object
			$obj | add-member -type NoteProperty -name User -value $userAccount.Name
			$obj | add-member -type NoteProperty -name HomeDirectory -value $userAccount.homeDirectory
			$obj | add-member -type NoteProperty -name HomeDirectorySizeMB -value ([Math]::Round($size.sum/1MB))
			$HomeD += $obj
			
		}
	}
	End
	{
		$HomeD
	}
}
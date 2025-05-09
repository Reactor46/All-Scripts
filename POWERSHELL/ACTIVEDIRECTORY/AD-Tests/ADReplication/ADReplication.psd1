@{
	# Script module or binary module file associated with this manifest
	ModuleToProcess = 'ADReplication.psm1'

	# Version number of this module.
	ModuleVersion = '2.0'

	# ID used to uniquely identify this module
	GUID = '7a3df55f-cf9c-4521-80b0-48a63f3dddae'

	# Author of this module
	Author = 'Raimund Andree'

	# Company or vendor of this module
	CompanyName = 'Microsoft'

	# Copyright statement for this module
	Copyright = '2011'

	# Description of the functionality provided by this module
	Description = 'Windows PowerShell Module for managing AD Sites, Site Links, Subets and reading replication information'

	# Minimum version of the Windows PowerShell engine required by this module
	PowerShellVersion = '2.0'

	# Name of the Windows PowerShell host required by this module
	PowerShellHostName = ''

	# Minimum version of the Windows PowerShell host required by this module
	PowerShellHostVersion = ''

	# Minimum version of the .NET Framework required by this module
	DotNetFrameworkVersion = '3.5'

	# Minimum version of the common language runtime (CLR) required by this module
	CLRVersion = ''

	# Processor architecture (None, X86, Amd64, IA64) required by this module
	ProcessorArchitecture = ''

	# Modules that must be imported into the global environment prior to importing this module
	# for example 'ActiveDirectory'
	RequiredModules = @()

	# Assemblies that must be loaded prior to importing this module
	# for example 'System.Management.Configuration'
	RequiredAssemblies = @()

	# Script files (.ps1) that are run in the caller's environment prior to importing this module
	ScriptsToProcess = @()

	# Type files (.ps1xml) to be loaded when importing this module
	TypesToProcess = @('ADReplication.types.ps1xml')

	# Format files (.ps1xml) to be loaded when importing this module
	FormatsToProcess = @('ADReplication.format.ps1xml')

	# Modules to import as nested modules of the module specified in ModuleToProcess
	NestedModules = @()

	# Functions to export from this module
	FunctionsToExport = 'Get-ADDomain',
		'Get-ADForest',
		'Get-ADLastChanges',
		'Get-ADReplicationConnection',
		'Get-ADReplicationLink',
		'Get-ADReplicationMetadata',
		'Get-ADReplicationQueue',
		'Get-ADReplicationSchedule',
		'Get-ADSite',
		'Get-ADSiteLink',
		'Get-ADSubnet',
		'Invoke-KCC',
		'New-ADReplicationConnection',
		'New-ADSite',
		'New-ADSiteLink',
		'New-ADSubnet',
		'Remove-ADReplicationConnection',
		'Remove-ADSite',
		'Remove-ADSiteLink',
		'Remove-ADSubnet',
		'Reset-ADReplicationSchedule',
		'Set-ADReplicationSchedule',
		'Set-ADSite',
		'Set-ADSiteLink',
		'Set-ADSubnet'


	# Variables to export from this module
	VariablesToExport = ''

	# Aliases to export from this module
	AliasesToExport = '*'

	# List of all modules packaged with this module
	ModuleList = @('ADReplication.psm1')

	# List of all files packaged with this module
	FileList = @('ADReplication.psm1', 'ADReplication.format.ps1xml')

	# Private data to pass to the module specified in ModuleToProcess
	PrivateData = ''
}
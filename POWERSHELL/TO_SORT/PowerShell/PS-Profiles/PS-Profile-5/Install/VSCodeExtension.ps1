try {
	Get-Command code -ErrorAction Stop > $null
} catch [System.Management.Automation.CommandNotFoundException] {
	Write-Host ('Command `{0}` is required but not found' -f $Error[0].TargetObject) -ForegroundColor Red
}

function Install-VSCodeExtension {
	[CmdletBinding(DefaultParameterSetName = 'Name')]
	param (
		[Parameter(ParameterSetName = 'Name', Mandatory, Position = 0)]
		[string[]]$Name,
		[Parameter(ParameterSetName = 'Recommended', Mandatory)]
		[switch]$Recommended
	)

	if ($Recommended) {
		$Name = (Get-Content "$(Split-Path $PROFILE)\.vscode\extensions.json" | % { $_ -replace '//.*','' } | ConvertFrom-Json).recommendations
	}
	$Name | % { code --install-extension $_ }
}

function Uninstall-VSCodeExtension {
	[CmdletBinding(DefaultParameterSetName = 'Name')]
	param (
		[Parameter(ParameterSetName = 'Name', Mandatory, Position = 0)]
		[string[]]$Name,
		[Parameter(ParameterSetName = 'All', Mandatory)]
		[switch]$All
	)

	if ($All) {
		$Name = code --list-extensions
	}
	$Name | % { code --uninstall-extension $_ }
}

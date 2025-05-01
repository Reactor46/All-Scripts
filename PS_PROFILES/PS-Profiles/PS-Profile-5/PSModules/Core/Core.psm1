function Grant-Administrator {
	$ps = switch ($PSVersionTable.PSVersion.Major) {
		{ $_ -le 5 } { 'powershell' }
		{ $_ -ge 6 } { 'pwsh' }
	}
	saps $ps -ArgumentList "-ExecutionPolicy $(Get-ExecutionPolicy)" -Verb RunAs
}

function New-Directory([string[]]$Path) { New-Item $Path -Force -ItemType Directory }
Set-Alias nd New-Directory
function Remove-Directory([string[]]$Path) { Remove-Item $Path -Recurse -Force -Confirm }
Set-Alias rd Remove-Directory

function Get-Accelerator([string]$Name = '*') {
	$accelerators = [powershell].Assembly.GetType('System.Management.Automation.TypeAccelerators')::Get
	$accelerators.GetEnumerator() | Where-Object Key -Like $Name | Sort-Object -Property Key
}

function Edit-Hosts { saps notepad.exe $env:SystemRoot\System32\drivers\etc\hosts -Verb RunAs }
function Get-CommandDefinition([string[]]$Name) { (gcm $Name).Definition }
function Test-Administrator { ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator) }

function Get-Constructor([Parameter(ValueFromPipeline = $true)][type]$Type) {
	foreach ($c in $Type.GetConstructors()) {
		$c.DeclaringType.FullName
		$c.GetParameters() | select Name, ParameterType | ft
	}
}

function Compare-PSObjectProperties {
	param (
		$ReferenceObject,
		$DifferenceObject
	)

	$referenceProperties = $ReferenceObject.psobject.Properties
	$differenceProperties = $DifferenceObject.psobject.Properties
	$names = $referenceProperties.Name + $differenceProperties.Name | Sort-Object -Unique
	foreach ($n in $names) {
		$result = Compare-Object $ReferenceObject $DifferenceObject -Property $n -ErrorAction Ignore
		if ($result) {
			[PSCustomObject]@{
				Property   = $n
				Reference  = $result | Where-Object SideIndicator -EQ '<=' | ForEach-Object $n
				Difference = $result | Where-Object SideIndicator -EQ '=>' | ForEach-Object $n
			}
		}
	}
}

function Save-WebResource {
	param (
		[uri]$Uri,
		[string]$OutFile = (Join-Path $HOME\Downloads ((Split-Path $Uri -Leaf) -replace "[$([System.IO.Path]::GetInvalidFileNameChars())]"))
	)

	(Test-Path $OutFile) -or (Invoke-WebRequest $Uri -OutFile $OutFile -Verbose) > $null
	Get-Item $OutFile
}

function New-ModuleItem {
	param (
		[Parameter(Mandatory)]
		[string]$Name
	)

	if (Test-Path $Name) {
		Write-Warning "$Name already exists."
		return
	}

	New-Item -Path (Join-Path $Name "$Name.psm1") -Force

	$params = @{
		Path                 = Join-Path $Name "$Name.psd1"
		RootModule           = $Name
		ModuleVersion        = '0.0.0.{0}' -f (Get-Date -Format 'yyMMdd')
		CompatiblePSEditions = 'Core'
		Author               = 'Masatoshi Higuchi'
		CompanyName          = 'N/A'
		Copyright            = '(c) Masatoshi Higuchi. All rights reserved.'
		PowerShellVersion    = '6.0'
		LicenseUri   = 'https://github.com/matt9ucci/{0}/blob/master/LICENSE' -f $Name
		ProjectUri   = 'https://github.com/matt9ucci/{0}' -f $Name
		ReleaseNotes = 'Initial release'
	}
	New-ModuleManifest @params
}

filter ToHex {
	'{0:X2}' -f $_
}

filter FromHex {
	[System.Convert]::ToInt64($_, 16)
}

filter ToBit {
	[System.Convert]::ToString($_, 2)
}

filter FromBit {
	[System.Convert]::ToInt64($_, 2)
}

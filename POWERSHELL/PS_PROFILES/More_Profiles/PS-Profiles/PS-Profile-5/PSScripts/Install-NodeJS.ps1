${desktop.ini} = { @"
[.ShellClassInfo]
IconResource=$IconFile,$IconIndex
InfoTip=$InfoTip
[ViewState]
Mode=
Vid=
FolderType=$FolderType
"@ }

function New-Desktop.ini {
	Param (
		[Parameter(Mandatory, Position = 0)]
		[string]$Directory,
		[string]$IconFile = '',
		[int]$IconIndex = 0,
		[string]$FolderType = '',
		[string]$InfoTip = ''
	)

	$ini = Join-Path $Directory 'desktop.ini'

	Set-Content $ini (Invoke-Command ${desktop.ini} -ArgumentList $PSBoundParameters)

	(Get-Item $ini -Force).Attributes += 'Archive, Hidden, System'
	(Get-Item $ini -Force).Directory.Attributes += 'ReadOnly'
}

$zip = "$DOWNLOADS\node-v8.11.4-win-x64.zip"
$dst = "$HOME\Apps\NodeJS"

$expanded = Expand-Archive -Path $zip -DestinationPath (Split-Path $dst) -PassThru
$root = ($expanded | ? PSIsContainer)[0]

Rename-Item -Path $dst -NewName "$dst-$(Get-Date -Format yyyyMMddHHmmss)" -ErrorAction Ignore
Rename-Item -Path $root.FullName -NewName $dst

New-Desktop.ini $dst -IconFile node.exe

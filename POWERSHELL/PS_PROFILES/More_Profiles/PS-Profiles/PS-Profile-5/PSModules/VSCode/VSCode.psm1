switch ($env:OS) {
	'Windows_NT' {
		Set-Variable VSCODE_USER_SETTINGS_JSON "$env:APPDATA\Code\User\settings.json" -Option ReadOnly, AllScope -Scope Global -Force
		Set-Variable VSCODE_ZIP_URL https://go.microsoft.com/fwlink/?Linkid=850641 -Option ReadOnly, AllScope -Scope Global -Force
	}
	# Mac $HOME/Library/Application Support/Code/User/settings.json
	# Linux $HOME/.config/Code/User/settings.json
	# See https://code.visualstudio.com/docs/getstarted/settings#_settings-file-locations
}

function Save-VSCodeBinary([string]$Path = "$DOWNLOADS\vscode.zip") {
	if (Test-Path $Path) {
		Write-Warning "$Path already exists"
	} else {
		Invoke-WebRequest -Uri $VSCODE_ZIP_URL -OutFile $Path -Verbose
	}
	$Path
}

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

function Install-VSCode([string]$Zip) {
	if (!$Zip) {
		$Zip = Save-VSCodeBinary
	}

	$VSCODE_HOME = "$HOME\Apps\VSCode"

	if (Test-Path $VSCODE_HOME) {
		Rename-Item $VSCODE_HOME "$(Split-Path $VSCODE_HOME -Leaf)-$(Get-Date -Format yyMMddHHmmss)"
	}

	Expand-Archive -Path $Zip -DestinationPath $VSCODE_HOME

	New-Desktop.ini $VSCODE_HOME -IconFile 'Code.exe'
}

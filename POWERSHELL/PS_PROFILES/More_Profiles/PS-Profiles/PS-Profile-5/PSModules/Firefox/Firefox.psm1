Set-Variable FIREFOX_HOME "$env:ProgramFiles\Mozilla Firefox" -Scope Global -Option ReadOnly, AllScope
Set-Variable FIREFOX_EXE (Join-Path $FIREFOX_HOME firefox.exe) -Scope Global -Option ReadOnly, AllScope

function Get-FirefoxProfile($Name = '*') {
	if ($IsLinux) {
		Get-ChildItem -Directory $HOME/.mozilla/firefox -Filter *.$Name
	} else {
		Get-ChildItem -Directory $env:APPDATA\Mozilla\Firefox\Profiles -Filter *.$Name
	}
}

function Install-Firefox([Parameter(Mandatory = $true)][string]$Exe) {
	$ini = Join-Path $PSScriptRoot firefox.ini
	Start-Process -FilePath $Exe -ArgumentList "/INI=$ini"
}

function Uninstall-Firefox {
	Start-Process -FilePath "$FIREFOX_HOME\uninstall\helper.exe" -ArgumentList '/S' -Wait
	Remove-Item -Confirm -Recurse $env:APPDATA\Mozilla\Firefox
	Remove-Item -Confirm -Recurse $env:LOCALAPPDATA\Mozilla\Firefox
}

function New-FirefoxProfile([string[]]$Name) {
	$Name | % { Start-Process -FilePath $FIREFOX_EXE -ArgumentList "-CreateProfile $_" -Wait }
}

function Start-Firefox([string]$Name) {
	if ($Name) {
		Start-Process -FilePath $FIREFOX_EXE -ArgumentList "-P $Name", '--no-remote'
	} else {
		Start-Process -FilePath $FIREFOX_EXE
	}
}

function Copy-FirefoxUserJs([string]$Name) {
	$js = Join-Path $PSScriptRoot user.js
	Copy-Item $js (Join-Path (Get-FirefoxProfile $Name).FullName user.js)
}

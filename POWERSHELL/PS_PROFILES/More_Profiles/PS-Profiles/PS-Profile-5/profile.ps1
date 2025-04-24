sv APPS (Join-Path $HOME Apps) -Option ReadOnly, AllScope
sv DESKTOP (Join-Path $HOME Desktop) -Option ReadOnly, AllScope
sv DOWNLOADS (Join-Path $HOME Downloads) -Option ReadOnly, AllScope
sv PROFILEDIR (Split-Path $PROFILE) -Option ReadOnly, AllScope
sv SCRIPTS (Join-Path $HOME Scripts) -Option ReadOnly, AllScope

sal gh help
sal gcb Get-Clipboard
sal scb Set-Clipboard
sal FromJson ConvertFrom-Json
sal ToJson ConvertTo-Json

sal d docker
sal dc docker-compose
sal dm docker-machine
sal g git

function ll { gci @Args | ? Name -NotLike .* }
function sl.. { sl .. }
function sl~ { sl ~ }
function Get-FileHash { $fh = Microsoft.PowerShell.Utility\Get-FileHash @Args; Set-Clipboard $fh.hash; $fh }
function codep { code $PROFILEDIR $PROFILE.CurrentUserAllHosts }
function codes { code $SCRIPTS }
if (!$IsLinux -and !$IsMacOS -and !(gcm Set-Clipboard -ErrorAction Ignore)) {
	function Set-Clipboard { param ([Parameter(Mandatory, ValueFromPipeline)][string]$Value) $Value | clip.exe }
}

if (gcm Set-PSReadlineOption -Module PSReadline -ea Ignore) {
	Set-PSReadlineOption -BellStyle None
	Set-PSReadlineOption -EditMode Windows
}

$env:PSModulePath = -join ((Join-Path $PROFILEDIR PSModules), [System.IO.Path]::PathSeparator, $env:PSModulePath)

if ($IsLinux) {
	$env:GOROOT = "$HOME/go"
	$env:GOPATH = "$HOME/gopath"
} else {
	$env:GOROOT = "C:\Go"
	$env:GOPATH = "$HOME\Gopath"
}

$env:Path = @(
	$SCRIPTS
	"$APPS\Git\cmd"
	"$APPS\Git\usr\bin"
	"$APPS\VSCode\bin"
	"$APPS\NodeJS"
	"$APPS\Dotnet"
	"$APPS\Geth"
	"$APPS\Solidity"
	"$APPS\GoIpfs"
	"$env:GOPATH\bin"
	"$env:GOROOT\bin"
	"${env:ProgramFiles}\Oracle\VirtualBox"
	$env:Path
) -join ";"

ipmo DockerCompletion -ea Ignore
ipmo DockerComposeCompletion -ea Ignore
ipmo DockerMachineCompletion -ea Ignore

ipmo Core
ipmo EnvironmentVariable
ipmo EthereumCompletion
ipmo Ipfs

. $PSScriptRoot\Completers.ps1

if ([System.Net.ServicePointManager]::SecurityProtocol -ne [System.Net.SecurityProtocolType]::SystemDefault) {
	[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
}

function prompt {
	$Host.UI.RawUI.WindowTitle = "PS $($ExecutionContext.SessionState.Path.CurrentLocation)$('>' * ($NestedPromptLevel + 1))"
	return "$($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)> "
}

if (Test-Path $HOME\.psprofiles\profile.ps1) {
	function codep { code $PROFILEDIR $PROFILE.CurrentUserAllHosts $HOME\.psprofiles\profile.ps1 }
	. $HOME\.psprofiles\profile.ps1
}

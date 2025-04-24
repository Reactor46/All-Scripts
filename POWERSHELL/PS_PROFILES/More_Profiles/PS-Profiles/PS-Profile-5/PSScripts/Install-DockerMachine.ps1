[CmdletBinding()]
param (
	[string]$Version,
	[switch]$Force
)

if (Get-Command docker-machine -ErrorAction Ignore) {
	Write-Warning "$(docker-machine version) already exists"
	if (!$Force) {
		return
	}
}

if (!$Version) {
	$latest = Invoke-RestMethod -Uri "https://api.github.com/repos/docker/machine/releases/latest" -Method Get -Verbose
	$Version = $latest.tag_name
}

$os = 'Windows'
$architecture = 'x86_64'
$extension = '.exe'
$machineUrl = "https://github.com/docker/machine/releases/download/${Version}/docker-machine-${os}-${architecture}${extension}"
$outFile = Join-Path $HOME\Scripts docker-machine.exe

Invoke-WebRequest -Uri $machineUrl -OutFile $outFile -Verbose

$sumUrl = "https://github.com/docker/machine/releases/download/${Version}/sha256sum.txt"
(Invoke-WebRequest -Uri $sumUrl -Verbose).RawContent

Get-FileHash -Path $outFile -Algorithm SHA256

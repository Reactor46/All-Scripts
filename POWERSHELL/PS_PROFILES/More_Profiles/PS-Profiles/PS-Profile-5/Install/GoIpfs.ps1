function Install-GoIpfs {
	param (
		[Alias('Tag')]
		$Version
	)

	if ($Version) {
		$uri = 'https://api.github.com/repos/ipfs/go-ipfs/releases/tags/{0}' -f $Version
	} else {
		$uri = 'https://api.github.com/repos/ipfs/go-ipfs/releases/latest'
	}

	$assets = (Invoke-RestMethod $uri -Verbose).assets

	Write-Host 'Select a number for download'
	for ($i = 0; $i -lt $assets.Count; $i++) {
		Write-Host ('{0}) {1}' -f $i, $assets[$i].name)
	}
	$n = Read-Host 'Download'
	$browserDownloadUrl = $assets[$n].browser_download_url

	$outFile = New-TemporaryFile
	Invoke-WebRequest $browserDownloadUrl -OutFile $outFile -Verbose
	Rename-Item $HOME\Apps\GoIpfs $HOME\Apps\GoIpfs-$(Get-Date -Format yyMMddHHmmss) -Verbose
	Expand-Archive $outFile $HOME\Apps\GoIpfs -Verbose
	Remove-Item $outFile
}

function Install-Go {
	param (
		[Parameter(Mandatory)]
		[string]$Zip
	)

	if (!$env:GOROOT) {
		Write-Host '$env:GOROOT is required.'
		return
	}

	try {
		(Invoke-Expression "$env:GOROOT\bin\go version") + ' is already installed.'
		return
	} catch {}

	$tempPath = Join-Path -Path $env:TEMP -ChildPath ('gozip' + $(Get-Date -Format 'yyMMddHHmmss'))
	Expand-Archive -Path $Zip -DestinationPath $tempPath

	$tempGoPath = Join-Path -Path $tempPath -ChildPath 'go'
	Get-ChildItem -Path $tempGoPath -Recurse | Move-Item -Destination $env:GOROOT
	Remove-Item -Path $tempPath -Recurse -Force

	(Invoke-Expression "$env:GOROOT\bin\go version") + ' has been installed.'
}

function Uninstall-Go {
	if (!$env:GOROOT) {
		Write-Host '$env:GOROOT is required.'
		return
	}
	Remove-Item -Path $env:GOROOT -Force -Recurse
}

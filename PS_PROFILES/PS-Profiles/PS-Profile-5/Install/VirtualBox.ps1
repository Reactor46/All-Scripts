function Install-VirtualBox {
	$uri = 'https://download.virtualbox.org/virtualbox/5.2.20/VirtualBox-5.2.20-125813-Win.exe'
	$checksum = '56ce706480d3ea411bcbf8932122633eb98e79202f1a4460d255a51997ba84f6'
	$checksumAlgorithm = 'SHA256'

	$installer = "$DOWNLOADS\$(Split-Path $uri -Leaf)"
	if (Test-Path $installer) {
		Write-Verbose "$installer already exists."
	} else {
		Invoke-WebRequest -Uri $uri -OutFile $installer
	}

	& "$installer" -msiparams VBOX_INSTALLDESKTOPSHORTCUT=0
}

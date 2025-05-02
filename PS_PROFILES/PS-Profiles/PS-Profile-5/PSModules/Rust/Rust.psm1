function Get-VersionInfo([switch]$All) {
	$info = [ordered]@{}
	rustc --version --verbose | Select-String ': ' | Sort-Object | % {
		$key, $value = ($_ -split ': ').Trim()
		$info.Add($key, $value)
	}

	if ($All) {
		[pscustomobject]$info
	} else {
		$info.release
	}
}

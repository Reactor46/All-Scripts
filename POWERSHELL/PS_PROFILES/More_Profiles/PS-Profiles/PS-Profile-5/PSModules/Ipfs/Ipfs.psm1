if (!(Get-Command -Name ipfs -ErrorAction Ignore)) {
	Write-Warning 'ipfs binary must be in $env:Path'
	return
}

[string[]]$cmds = ipfs commands
[string[]]$flgs = ipfs commands --flags | ForEach-Object {
	if ($_ -notin $cmds) {
		$_ -split ' / '
	}
}

[scriptblock]$completer = {
	param ([string]$wordToComplete, $commandAst, $cursorPosition)

	# replace 'ipfs.exe' and aliases with 'ipfs'
	$astString = "$commandAst" -replace '^\S+', 'ipfs'
	$astStringElements = $astString -split ' '

	$candidates = if ($wordToComplete.StartsWith('-')) { $flgs } else { $cmds }
	$candidates | Where-Object { $_ -match "$astString.+" } | ForEach-Object {
		$words = $_ -split ' '
		for ($i = 0; $i -lt $words.Count; $i++) {
			if (($words[$i] -ne $astStringElements[$i]) -and ($words[$i] -like "$wordToComplete*")) {
				New-Object System.Management.Automation.CompletionResult $words[$i]
				break
			}
		}
	} | Sort-Object CompletionText -Unique
}

'ipfs', (Get-Alias -Definition 'ipfs' -ErrorAction Ignore).Name | ForEach-Object {
	if ($_) { Register-ArgumentCompleter -Native -CommandName $_ -ScriptBlock $completer }
}

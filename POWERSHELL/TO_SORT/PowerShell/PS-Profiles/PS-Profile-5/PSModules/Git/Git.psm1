function Start-Rebase ([int]$Last) {
	git rebase -i "HEAD~$Last"
}

function g-clone-depth-1 ($Repository) { git clone --depth 1 $Repository }

function Get-GitConfig {
	param ([switch]$Global)

	$config = @{}
	$section = ''
	$name = ''
	$configPath = if ($Global) { "$HOME\.gitconfig" } else { '.git\config' }

	Get-Content $configPath | ForEach-Object {
		$line = $_.Trim()
		switch -Regex ($line) {
			'\[(\w*)\s*"?(\w*)"?\]' {
				$section = $Matches[1]
				$name = $Matches[2]

				if (!$config[$section]) {
					$config[$section] = @{}
				}

				if ($name) {
					$config[$section][$name] = @{}
				}
			}
			'(\S+)\s*=\s*(\S+)' {
				$key = $Matches[1]
				$value = switch ($Matches[2]) {
					'true' { $true }
					'false' { $false }
					Default { $_ }
				}

				if ($name) {
					$config[$section][$name][$key] = $value
				} else {
					$config[$section][$key] = $value
				}
			}
			Default { $_ }
		}
	}
	$config
}

function Get-RemoteRepositoryInfo {
	$remote = (Get-GitConfig).remote
	foreach ($key in $remote.Keys) {
		$result = [ordered]@{
			Name = $key
		}

		$r = $remote[$key]
		foreach ($k in $r.Keys) {
			$k = (Get-Culture).TextInfo.ToTitleCase($k)
			$result[$k] = $r[$k]
		}

		[pscustomobject]$result
	}
}

function Get-Branch {
	git branch --list | % Trim '* '
}

function Remove-Branch {
	param (
		[string]$Name,
		[switch]$Force
	)

	if ($Force) {
		git branch -d -f $Name
	} else {
		git branch -d $Name
	}
}

$commandNames = (Get-Command -Module Git -Name *-Branch).Name
Register-ArgumentCompleter -ParameterName Name -CommandName $commandNames -ScriptBlock {
	param ($commandName, $parameterName, $wordToComplete)
	(Get-Branch) -like "$wordToComplete*"
}

function Get-PathEnv {
	param (
		[System.EnvironmentVariableTarget]$Target = [System.EnvironmentVariableTarget]::Process
	)

	(Get-Env 'PATH' $Target).Value -split ';'
}

function Set-PathEnv {
	param (
		[string[]]$Path,
		[System.EnvironmentVariableTarget]$Target = [System.EnvironmentVariableTarget]::Process
	)

	Set-Env 'PATH' ($Path -join ';') $Target
}

function Add-PathEnv {
	param (
		[string]$Path,
		[System.EnvironmentVariableTarget]$Target = [System.EnvironmentVariableTarget]::Process,
		[switch]$First
	)

	if ((Get-PathEnv $Target) -contains $Path) {
		Write-Host $Target "PATH contains" $Path -ForegroundColor Green
	} elseif ($First) {
		Set-PathEnv ((@(,$Path) + (Get-PathEnv $Target)) -join ';') $Target
	} else {
		Set-PathEnv (((Get-PathEnv $Target) + $Path) -join ';') $Target
	}
}

function Remove-PathEnv {
	param (
		[string]$Path,
		[System.EnvironmentVariableTarget]$Target = [System.EnvironmentVariableTarget]::Process
	)

	if ((Get-PathEnv $Target) -contains $Path) {
		Set-PathEnv ((Get-PathEnv $Target) -ne @($Path)) $Target
	} else {
		Write-Host $Target "PATH does not contain" $Path -ForegroundColor Green
	}
}

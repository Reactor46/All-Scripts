function New-DockerMachine {
	param (
		[string]$Name = 'default',
		[ValidateSet(
			'amazonec2',
			'azure',
			'digitalocean',
			'exoscale',
			'generic',
			'google',
			'hyperv',
			'none',
			'openstack',
			'rackspace',
			'softlayer',
			'virtualbox',
			'vmwarefusion',
			'vmwarevcloudair',
			'vmwarevsphere'
		)]
		[string]$Driver = 'virtualbox',
		[string]$Version,
		[uint16]$Cpu,
		[uint64]$Memory,
		[switch]$BehindProxy
	)

	$options = @()
	if ($BehindProxy) {
		Get-Item -Path 'Env:HTTP_PROXY', 'Env:HTTPS_PROXY', 'Env:NO_PROXY' -ErrorAction Ignore| ForEach-Object {
			$options += '--engine-env {0}={1}' -f $_.Key, $_.Value
		}
	}
	if ($Version) {
		$options += "--virtualbox-boot2docker-url https://github.com/boot2docker/boot2docker/releases/download/$Version/boot2docker.iso"
	}
	if ($Cpu) {
		$options += "--virtualbox-cpu-count $Cpu"
	}
	if ($Memory) {
		switch ($Driver) {
			virtualbox { $options += "--virtualbox-memory $($Memory/1MB)" }
			hyperv { $options += "--hyperv-memory $($Memory/1MB)" }
		}
	}
	$options = $options -join ' '

	Invoke-Expression -Command "docker-machine create --driver $Driver $options $Name"
}

function Connect-DockerMachine([string]$Name = 'default') {
	docker-machine env --shell powershell --no-proxy $Name | Invoke-Expression
}

Register-ArgumentCompleter -CommandName Connect-DockerMachine -ParameterName Name -ScriptBlock {
	param ($commandName, $parameterName, [string]$wordToComplete)

	if ($wordToComplete) {
		docker-machine ls --quiet --filter name=$wordToComplete
	} else {
		docker-machine ls --quiet
	}
}

function New-SwarmMachine {
	param (
		[int]$ManagerSize = 1,
		[Alias('ClusterSize')]
		[int]$WorkerSize = 2
	)

	1..$ManagerSize | ForEach-Object {
		New-DockerMachine "manager$i"
	}

	1..$WorkerSize | ForEach-Object {
		New-DockerMachine "worker$i"
	}
}

Set-Alias ccdm Connect-DockerMachine

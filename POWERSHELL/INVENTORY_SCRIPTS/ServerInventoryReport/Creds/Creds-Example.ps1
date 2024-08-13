Get-Credential | Export-Clixml -path .\Creds\SQLCredsSA.xml
$cred = import-clixml -path .\Creds\SQLCredsSA.xml

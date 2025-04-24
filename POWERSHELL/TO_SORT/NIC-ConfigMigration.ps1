# Define source and destination network adapter names
$sourceAdapterName = "Ethernet0"
$destinationAdapterName = "Ethernet1"

# Get the network configuration of the source network adapter
$sourceConfiguration = Get-NetIPConfiguration -InterfaceAlias $sourceAdapterName

# Check if the destination adapter is already enabled
$destinationAdapter = Get-NetAdapter -Name $destinationAdapterName
if (-not $destinationAdapter.Status -eq 'Up') {
    # Enable the destination network adapter if it's not already enabled
    Enable-NetAdapter -Name $destinationAdapterName
}

# Apply the network configuration to the destination network adapter
foreach ($config in $sourceConfiguration) {
    # Configure IPv4 address
    $sourceIPv4Config = $config.IPv4Address
    if ($sourceIPv4Config -ne $null) {
        $ipv4Config = @{
            IPAddress = $sourceIPv4Config.IPAddress
            PrefixLength = $sourceIPv4Config.PrefixLength
            InterfaceAlias = $destinationAdapterName
        }
        New-NetIPAddress @ipv4Config -Confirm:$false
    }

    # Configure DNS server addresses
    $dnsServers = $config.DNSServer.ServerAddresses
    if ($dnsServers -ne $null) {
        Set-DnsClientServerAddress -InterfaceAlias $destinationAdapterName -ServerAddresses $dnsServers -Confirm:$false
    }

    # Configure gateway
    $gateway = $config.IPv4DefaultGateway.NextHop
    if ($gateway -ne $null) {
        Set-NetIPInterface -InterfaceAlias $destinationAdapterName -IPv4DefaultGateway $gateway -Confirm:$false
    }

    # Disable the source network adapter
    Disable-NetAdapter -Name $sourceAdapterName -Confirm:$false
}

Write-Output "Network settings moved successfully from $sourceAdapterName to $destinationAdapterName. $sourceAdapterName is now disabled."

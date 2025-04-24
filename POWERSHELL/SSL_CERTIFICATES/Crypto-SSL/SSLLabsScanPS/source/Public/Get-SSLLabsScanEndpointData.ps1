function Get-SSLLabsScanEndpointData
{
    <#
    .SYNOPSIS
        Gets the endpoint data for a scan

    .DESCRIPTION
        This function gets the endpoint data for a scan

    .PARAMETER HostName
        Specifies the hostname of the scan.

    .PARAMETER IPAddress
        Specifies the ip address of the host in the scan.

    .PARAMETER HostData
        Specifies the HostData object containing details of the scan

    .PARAMETER FromCache
        Specifies whether to retrieve the scan from the cache.

    .INPUTS
        None

    .OUTPUTS
        SSLLabsScan.EndPointData

    .EXAMPLE
        Get-SSLLabsScanEndpointData -HostName www.bbc.co.uk -IPAddress 1.1.1.1

        Gets the endpoint data for a scan on the bbc website for the specified IP address.
#>
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [OutputType( { ($script:ApiPropertes | Where-Object -Property ApiName -eq 'getEndpointData').TypeName })]
    param(
        [Parameter(
            Mandatory,
            ParameterSetName = 'Default')]
        [System.String]
        $HostName,

        [Parameter(
            Mandatory,
            ParameterSetName = 'Default')]
        [System.String[]]
        $IPAddress,

        [Parameter(
            Mandatory,
            ParameterSetName = 'HostData')]
        [PSCustomObject]
        $HostData,

        [Parameter()]
        [System.Management.Automation.SwitchParameter]
        $FromCache
    )

    $apiName = 'getEndpointData'

    if ($PSCmdlet.ParameterSetName -eq 'HostData')
    {
        $IPAddress = $HostData.endpoints.ipaddress
        $HostName = $HostData.host
    }

    $baseQueryParams = @()

    $baseQueryParams += "host=$HostName"

    if ($PSBoundParameters.ContainsKey('FromCache'))
    {
        $baseQueryParams += 'fromCache=on'
    }

    foreach ($ip in $IPAddress)
    {
        $queryParams = $baseQueryParams + "s=$ip"
        Write-Verbose "Getting SSL Labs Scan endpoint data on host $HostName, IP address $ip"

        try
        {
            $result = Invoke-SSLLabsScanApi -ApiName $apiName -QueryParameters $queryParams -Verbose:$false
        }
        catch
        {
            $errorRecord = New-ErrorRecord -Exception $_.Exception
            $PSCmdlet.ThrowTerminatingError($errorRecord)
        }

        $result | Add-Member -Name 'host' -Value $HostName -MemberType NoteProperty
        $result.details.hostStartTime = ([System.DateTimeOffset]::FromUnixTimeMilliSeconds($result.details.hostStartTime)).UtcDateTime

        return $result
    }
}

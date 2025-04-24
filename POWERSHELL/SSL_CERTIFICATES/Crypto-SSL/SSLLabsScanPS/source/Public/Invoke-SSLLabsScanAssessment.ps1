function Invoke-SSLLabsScanAssessment
{
    <#
    .SYNOPSIS
        Invokes an SSL Labs scan assessment of a website

    .DESCRIPTION
        This function invokes an SSL Labs scan assessment of a website

    .PARAMETER HostName
        Specifies the hostname for the scan

    .PARAMETER Publish
        Specifies whether to publish the scan results

    .PARAMETER StartNew
        Specifies whether to start a new scan

    .PARAMETER FromCache
        Specifies whether to retrieve a scan from the cache

    .PARAMETER MaxAge
        Specifies the maximum age in hours of a scan to retrieve from the cache

    .PARAMETER All
        If specified with a value of 'on', full information on individual endpoints will be returned.
        If specified with a value of 'done', full information on individual endpoints will only be returned if the
        assessment is complete.

    .PARAMETER IgnoreMismatch
        Specifies whether to ignore mismatches between the server certificate and the assessment hostname.

    .PARAMETER PollingInterval
        Specifies the polling interval in seconds between scan status checks.

    .INPUTS
        None

    .OUTPUTS
        SSLLabsScan.HostData

    .EXAMPLE
        Invoke-SSLLabsScanAssessment -HostName 'www.bbc.co.uk' -StartNew

        Invokes a new SSL Labs scan assessment of a the bbc web site
#>
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [OutputType( { ($script:ApiPropertes | Where-Object -Property ApiName -eq 'Analyze').TypeName })]
    param(
        [Parameter(Mandatory)]
        [System.String]
        $HostName,

        [Parameter()]
        [System.Management.Automation.SwitchParameter]
        $Publish,

        [Parameter(ParameterSetName = 'StartNew')]
        [System.Management.Automation.SwitchParameter]
        $StartNew,

        [Parameter(ParameterSetName = 'FromCache')]
        [System.Management.Automation.SwitchParameter]
        $FromCache,

        [Parameter(ParameterSetName = 'FromCache')]
        [System.Int32]
        $MaxAge,

        [Parameter()]
        [ValidateSet('On', 'Done')]
        [System.String]
        $All,

        [Parameter()]
        [System.Management.Automation.SwitchParameter]
        $IgnoreMismatch,

        [Parameter()]
        [System.Int32]
        $PollingInterval = 10

    )

    $apiName = 'analyze'

    $queryParams = @()

    $queryParams += "host=$HostName"

    if ($PSBoundParameters.ContainsKey('Publish'))
    {
        $queryParams += 'publish=on'
    }

    if ($PSBoundParameters.ContainsKey('FromCache'))
    {
        $queryParams += 'fromCache=on'
    }

    if ($PSBoundParameters.ContainsKey('MaxAge'))
    {
        $queryParams += "maxAge=$MaxAge"
    }

    if ($PSBoundParameters.ContainsKey('All'))
    {
        $queryParams += "all=$All"
    }

    if ($PSBoundParameters.ContainsKey('IgnoreMismatch'))
    {
        $queryParams += 'ignoreMismatch=on'
    }

    $initialQueryParams = $queryParams

    # Only add the 'startNew' parameter on the initial query
    if ($PSBoundParameters.ContainsKey('StartNew'))
    {
        $initialQueryParams += 'startNew=on'
    }

    Write-Verbose "Invoking SSL Labs Scan API Analysis on host $HostName"

    $progressActivityMessage = "Checking SSL Labs Scan API Analysis on host $HostName"

    try
    {
        $result = Invoke-SSLLabsScanApi -ApiName $apiName -QueryParameters $initialQueryParams -Verbose:$false
    }
    catch
    {
        $errorRecord = New-ErrorRecord -Exception $_.Exception
        $PSCmdlet.ThrowTerminatingError($errorRecord)
    }

    $retryCount = 0

    while ($result.status -ne 'READY' -and $result.status -ne 'ERROR')
    {
        $retryCount++

        Write-Progress -Activity $progressActivityMessage -Status "Status: $($result.status), $retryCount"

        Start-Sleep -Seconds $PollingInterval

        try
        {
            $result = Invoke-SSLLabsScanApi -ApiName $apiName -QueryParameters $queryParams -Verbose:$false
        }
        catch
        {
            $errorRecord = New-ErrorRecord -Exception $_.Exception
            $PSCmdlet.ThrowTerminatingError($errorRecord)
        }
    }

    # Convert Unix time fields to PowerShell DateTime objects
    $result.startTime = ([System.DateTimeOffset]::FromUnixTimeMilliSeconds($result.startTime)).UtcDateTime
    $result.testTime = ([System.DateTimeOffset]::FromUnixTimeMilliSeconds($result.testTime)).UtcDateTime

    foreach ($endpoint in $result.endpoints)
    {
        $endpoint | Add-Member -Name 'host' -Value $HostName -MemberType NoteProperty

        if ($endpoint.details.hostStartTime)
        {
            $endpoint.details.hostStartTime = ([System.DateTimeOffset]::FromUnixTimeMilliSeconds($endpoint.details.hostStartTime)).UtcDateTime
        }
    }

    Write-Progress -Activity $progressActivityMessage -Completed

    return $result
}

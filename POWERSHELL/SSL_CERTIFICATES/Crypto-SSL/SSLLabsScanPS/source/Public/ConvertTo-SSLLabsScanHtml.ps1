function ConvertTo-SSLLabsScanHtml
{
    <#
    .SYNOPSIS
        Converts an SSL Labs Scan to an HTML report

    .DESCRIPTION
        This function converts an SSL Labs Scan to an HTML report

    .PARAMETER EndPointData
        Specifies the endpoint data to use for the report.

    .PARAMETER Path
        Specifies the output path for the report. If not specified, this defaults to the user's 'Documents' folder,
        with a file name of <hostName>-SSLLabsScanReport-<yyyyMMdd-HHmmss>.html'.

    .INPUTS
        None

    .OUTPUTS
        None

    .EXAMPLE
        ConvertTo-SSLLabsScanHtml -EndPointData $endpointData

        Converts an SSL Labs Scan to an HTML report
#>
    [CmdletBinding(PositionalBinding = $false)]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject[]]
        $EndPointData,

        [Parameter()]
        [System.String]
        $Path
    )

    if ((-not $EndPointData[0].host) -or (-not $EndPointData[0].details.hostStartTime))
    {
        Write-Error 'Invalid EndPointData.' -ErrorAction Stop
    }

    $hostName = $EndPointData[0].host
    $scanDate = $EndPointData[0].details.hostStartTime

    $htmlBody = [System.String]::Empty

    if (-not $PSBoundParameters.ContainsKey('Path'))
    {
        $fileName = "$hostName-SSLLabsScanReport-$($scanDate.ToString('yyyyMMdd-HHmmss')).html"
        $Path = Join-Path -Path ([Environment]::GetFolderPath('MyDocuments')) -ChildPath $fileName
    }

    $inlineStyleSheet = @(
        '<style>'
        (Get-Content -Path "$script:resourceDirectoryName\default.css")
        '</style>'
    )

    $header = "SSLLabs Scan for $hostName"

    $preContent = "<h2>$header</h2>"

    foreach ($endpoint in $EndpointData)
    {
        Write-Verbose -Message "Converting scan for endpoint $($endpoint.ipAddress)"

        $protocols = @()
        foreach ($protocol in $endpoint.details.protocols)
        {
            $protocols += "$($protocol.name) v$($protocol.version)"
        }

        $cipherSuites = ($endpoint.details.suites.list.name | Out-String).Trim() -replace ('\r\n', '<br/>')

        $openSslCcsStatusMapping = @{
            -1 = 'Test Failed'
            0  = 'Unknown'
            1  = 'Not Vulnerable'
            2  = 'Possible Vulnerable, but not Exploitable'
            3  = 'Vulnerable and Exploitable'
        }

        $openSslCcsStatus = $openSslCcsStatusMapping[$endpoint.details.openSslCcs]

        $openSslLuckyMinus20StatusMapping = @{
            -1 = 'Test Failed'
            0  = 'Unknown'
            1  = 'Not Vulnerable'
            2  = 'Vulnerable and Insecure'
        }

        $openSslLuckyMinus20Status = $openSslLuckyMinus20StatusMapping[$endpoint.details.openSSLLuckyMinus20]

        $poodleTlsStatusMapping = @{
            -3 = 'Timeout'
            -2 = 'TLS Not Supported'
            -1 = 'Test Failed'
            0  = 'Unknown'
            1  = 'Not Vulnerable'
            2  = 'Vulnerable'
        }

        $poodleTlsStatus = $poodleTlsStatusMapping[$endpoint.details.poodleTls]

        $reportData = [PSCustomObject][Ordered]@{
            'Server Name'                   = $endpoint.serverName
            'Grade'                         = $endpoint.grade
            'Grade Ignoring Trust'          = $endpoint.gradeTrustIgnored
            'Has Warnings'                  = $endpoint.hasWarnings
            'Is Exceptional'                = $endpoint.isExceptional
            'Certificate Subject'           = $endpoint.details.cert.subject
            'Supported Protocols'           = $protocols -join ', '
            'Supported Cipher Suites'       = $cipherSuites
            'BEAST Vulnerable'              = $endpoint.details.vulnBeast
            'Heartbleed Vulnerable'         = $endpoint.details.Heartbleed
            'Poodle Vulnerable'             = $endpoint.details.poodle
            'PoodleTLS Status'              = $poodleTlsStatus
            'FREAK Vulnerable'              = $endpoint.details.freak
            'Drown Vulnerable'              = $endpoint.details.drownVulnerable
            'OpenSSL CCS Status'            = $openSslCcsStatus
            'OpenSSL Lucky Minus 20 Status' = $openSslLuckyMinus20Status
        }

        $endpointPreContent = @(
            '<h3>'
            "IP Address: $($endpoint.ipAddress)"
            '</h3>'
        )

        $htmlBody += $reportData | ConvertTo-Html -As List -PreContent $endpointPreContent -Fragment
    }

    $htmlReport = ConvertTo-Html -Head $inlineStyleSheet -PreContent $preContent -PostContent $htmlBody

    [System.Net.WebUtility]::HtmlDecode($htmlReport) | Out-File -FilePath $Path -Encoding ascii
}

function Get-SSLLabsScanInfo
{
    <#
    .SYNOPSIS
        Gets details of the SSLLabs Scan API

    .DESCRIPTION
        This function gets details of the SSLLabs Scan API

    .INPUTS
        None

    .OUTPUTS
        SSLLabsScan.Info

    .EXAMPLE
        Get-SSLLabsScanInfo

        Gets the SSLLabs Scan API info.
#>
    [CmdletBinding()]
    [OutputType( { ($script:ApiPropertes | Where-Object -Property ApiName -eq 'info').TypeName })]
    param()

    $apiName = 'info'

    Write-Verbose 'Getting SSL Labs Scan API Info'

    try
    {
        $result = Invoke-SSLLabsScanApi -ApiName $apiName
    }
    catch
    {
        $errorRecord = New-ErrorRecord -Exception $_.Exception
        $PSCmdlet.ThrowTerminatingError($errorRecord)
    }

    $result
}

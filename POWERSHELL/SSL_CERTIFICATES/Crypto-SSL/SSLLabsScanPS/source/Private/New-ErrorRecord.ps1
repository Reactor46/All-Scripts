function New-ErrorRecord
{
    <#
    .SYNOPSIS
        Creates an error record from an exception object.

    .DESCRIPTION
        This function creates an error record from an exception object.

    .PARAMETER Exception
        Specifies the exception object to include in the error record.

    .INPUTS
        None

    .OUTPUTS
        System.Management.Automation.ErrorRecord

    .EXAMPLE
        New-ErrorRecord -Exception $exception

        Creates an error record containing the specified exception.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Non state changing')]
    [CmdletBinding()]
    param(
        [Parameter(
            Mandatory,
            Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.Exception]
        $Exception
    )

    $exceptionTypeName = $Exception.GetType().FullName

    if ($exceptionTypeName -eq 'System.Net.WebException' -or
        $exceptionTypeName -eq 'Microsoft.PowerShell.Commands.HttpResponseException')
    {
        $fullyQualifiedErrorId = ("$($errorRecord.Exception.Response.StatusCode.value__)-" +
            $errorRecord.Exception.Response.StatusCode)
    }
    else
    {
        $fullyQualifiedErrorId = 'UnknownException'
    }

    $ErrorRecord = [System.Management.Automation.ErrorRecord]::new(
        $Exception,
        $fullyQualifiedErrorId,
        [System.Management.Automation.ErrorCategory]::InvalidOperation,
        $Null
    )

    return $ErrorRecord
}

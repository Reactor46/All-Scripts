function InvokeSslLabsRestMethod {
    [CmdletBinding()]
    param (
        [Parameter()]
        [String]$Server = 'https://api.ssllabs.com',

        [Parameter()]
        [ValidateSet(2,3,4)]
        [Int]$ApiVersion = 3,

        [Parameter(Mandatory)]
        [String]$Endpoint,

        [Parameter()]
        [String]$Path,

        [Parameter()]
        [System.Collections.Specialized.NameValueCollection]$Query,

        [Parameter()]
        [hashtable]$Body,

        [Parameter()]
        [hashtable]$Headers,

        [Parameter()]
        [Microsoft.PowerShell.Commands.WebRequestMethod]$Method = 'GET'
    )

    $Params = @{
        Uri           = '{0}/api/v{1}/{2}' -f $Server, $ApiVersion, $Endpoint
        Method        = $Method
        ContentType   = 'application/json'
        Verbose       = $VerbosePreference
        ErrorAction   = 'Stop'
        ErrorVariable = 'InvokeWebRequestError'
    }

    if ($PSBoundParameters.ContainsKey("Path")) {
        $Params['Uri'] = '{0}/{1}' -f $Params["Uri"], $Path
    }

    if ($PSBoundParameters.ContainsKey("Query")) {
        $Params['Uri'] = '{0}?{1}' -f $Params["Uri"], $Query.ToString()
    }

    if ($PSBoundParameters.ContainsKey("Body")) {
        $Params['Body'] = $Body | ConvertTo-Json -Compress
    }

    if ($PSBoundParameters.ContainsKey("Headers")) {
        $Params['Headers'] = $Headers
    }

    try {
        Write-Verbose ('URL: {0}'     -f $Params["Uri"])
        Write-Verbose ('Method: {0}'  -f $Params["Method"])
        Write-Verbose ('Body: {0}'    -f $Params["Body"])
        Write-Verbose ('Headers: {0}' -f $Params["Headers"])
        Invoke-WebRequest @Params
    }
    catch {
        # The web exception class is different for Core vs Windows
        if ($InvokeWebRequestError.ErrorRecord.Exception.GetType().FullName -match "HttpResponseException|WebException") {
            # Give the user a minified json string as exception message because the json contains two useful properties (field and message)
            # See section 'Error Reporting' https://github.com/ssllabs/ssllabs-scan/blob/master/ssllabs-api-docs-v4.md
            $ExceptionMessage = $InvokeWebRequestError.ErrorRecord.ErrorDetails.Message | ConvertFrom-Json | ConvertTo-Json -Compress
            $ErrorId = "{0}{1}" -f 
                [Int][System.Net.HttpStatusCode]$InvokeWebRequestError.ErrorRecord.Exception.Response.StatusCode, 
                [String][System.Net.HttpStatusCode]$InvokeWebRequestError.ErrorRecord.Exception.Response.StatusCode

            switch -Regex ($InvokeWebRequestError.ErrorRecord.Exception.StatusCode) {
                'TooManyRequests' {
                    $Exception = [System.ServiceModel.ServerTooBusyException]::new($ExceptionMessage)
                    $ErrorRecord = [System.Management.Automation.ErrorRecord]::new(
                        $Exception,
                        $ErrorId,
                        [System.Management.Automation.ErrorCategory]::ResourceBusy,
                        $Params['Uri']
                    )
                }
                'BadRequest' {
                    $Exception = [System.ArgumentException]::new($ExceptionMessage)
                    $ErrorRecord = [System.Management.Automation.ErrorRecord]::new(
                        $Exception,
                        $ErrorId,
                        [System.Management.Automation.ErrorCategory]::InvalidArgument,
                        $Params['Uri']
                    )
                }
                'NotFound' {
                    $Exception = [System.Management.Automation.ItemNotFoundException]::new($ExceptionMessage)
                    $ErrorRecord = [System.Management.Automation.ErrorRecord]::new(
                        $Exception,
                        $ErrorId,
                        [System.Management.Automation.ErrorCategory]::ObjectNotFound,
                        $Params['Uri']
                    )
                }
                'ServiceUnavailable' {
                    $Exception = [System.InvalidOperationException]::new($ExceptionMessage)
                    $ErrorRecord = [System.Management.Automation.ErrorRecord]::new(
                        $Exception,
                        $ErrorId,
                        [System.Management.Automation.ErrorCategory]::ResourceUnavailable,
                        $Params['Uri']
                    )
                }
                default {
                    $Exception = [System.InvalidOperationException]::new($ExceptionMessage)
                    $ErrorRecord = [System.Management.Automation.ErrorRecord]::new(
                        $Exception,
                        $ErrorId,
                        [System.Management.Automation.ErrorCategory]::InvalidOperation,
                        $Params['Uri']
                    )
                }
            }

            $PSCmdlet.ThrowTerminatingError($ErrorRecord)
        }
        else {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}
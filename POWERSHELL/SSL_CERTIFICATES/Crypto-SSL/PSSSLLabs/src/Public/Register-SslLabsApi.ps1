function Register-SslLabsApi {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$FirstName,

        [Parameter(Mandatory)]
        [String]$LastName,

        [Parameter(Mandatory)]
        [String]$Email,

        [Parameter(Mandatory)]
        [Alias('Organization')]
        [String]$Organisation
    )

    try {
        InvokeSslLabsRestMethod -Endpoint 'register' -Method 'POST' -Body @{
            firstName    = $FirstName
            lastName     = $LastName
            email        = $Email
            organization = $Organisation
        }
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}
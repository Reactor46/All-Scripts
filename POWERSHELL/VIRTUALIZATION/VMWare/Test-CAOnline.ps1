function Test-CAOnline {
[CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [string]$Config,
        [switch]$ShowUI
    )
$signature = @"
[DllImport("Certadm.dll", CharSet=CharSet.Auto, SetLastError=true)]
public static extern bool CertSrvIsServerOnline(
    string pwszServerName,
    ref bool pfServerOnline
);
"@
    Add-Type -MemberDefinition $signature -Namespace CryptoAPI -Name CertAdm
    $CertConfig = New-Object -ComObject CertificateAuthority.Config
    if ($Config -ne "" -and !$Config.Contains("\")) {
        Write-Error -Category InvalidArgument -ErrorId InvalidArgumentException `

        -Message "Config string must be passed in 'CAHostName\CAName' form."
        break
    } elseif ($Config -eq "" -and !$ShowUI) {
        try {$Config = $CertConfig.GetConfig(0x3)}
        catch {
            Write-Error -Category ObjectNotFound -ErrorId ObjectNotFoundElement `

            -Message "Certificate Services are not installed on local computer."
            break
        }
    } elseif ($Config -eq "" -and $ShowUI) {
        $Config = $CertConfig.GetConfig(0x1)
    }
    if ($Config) {
        [void]($Config -match "(.+)\\(.+)")
        $Server = $matches[1]
        $CAName = $matches[2]
        $ServerStatus = $false
        $hresult = [CryptoAPI.CertAdm]::CertSrvIsServerOnline($Server,[ref]$ServerStatus)
        if ($ServerStatus) {
            $CertAdmin = New-Object -ComObject CertificateAuthority.Admin
            $CertRequest = New-Object -ComObject CertificateAuthority.Request
            $CA = New-Object psobject -Property @{
                Name = $CAName;
                ICertAdmin = $true;
                ICertRequest = $true
            }
            try {$retn = $CertAdmin.GetCAProperty($Config,0x6,0,4,0)}
            catch {$CA.ICertAdmin = $false}
            try {$retn = $CertRequest.GetCAProperty($Config,0x6,0,4,0)}
            catch {$CA.ICertRequest = $false}
            $CA
        } else {
            Write-Error -Category ObjectNotFound -ErrorId ObjectNotFoundException `

            -Message "Unable to find a Certification Authority server on '$Server'."
        }
    } else {return}
}
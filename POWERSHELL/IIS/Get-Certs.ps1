function Get-RemoteIISCertificate {
Param([Parameter(Mandatory=$true)]
	[string[]] $ComputerName,
    [int]
    $Port = 443
)

$Connection = New-Object System.Net.Sockets.TcpClient($ComputerName,$Port)
$Connection.SendTimeout = 5000
$Connection.ReceiveTimeout = 5000
$Stream = $Connection.GetStream()

try {
    $sslStream = New-Object System.Net.Security.SslStream($Stream,$False,([Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}))
    #$sslStream = New-Object System.Net.Security.sslStream.AuthenticateAsServer(certificate, false, SslProtocols.Default, true);
    $sslStream.AuthenticateAsClient('')

    #$Certificate = [Security.Cryptography.X509Certificates.X509Certificate2]$sslStream.RemoteCertificate

    $cert = $sslStream.get_remotecertificate()
    $cert2 = New-Object system.security.cryptography.x509certificates.x509certificate2($cert)

    $validto = [datetime]::Parse($cert.getexpirationdatestring())
    $validfrom = [datetime]::Parse($cert.geteffectivedatestring())

    if ($cert.get_issuer().CompareTo($cert.get_subject())) {
        $selfsigned = "no";
    } else {
        $selfsigned = "yes";
    }

    Write-Host '"' -nonewline; Write-Host $ComputerName -nonewline; Write-Host '",' -nonewline;
    Write-Host '"' -nonewline; Write-Host $Port -nonewline; Write-Host '",' -nonewline;
    Write-Host '"' -nonewline; Write-Host $cert.get_subject() -nonewline; Write-Host '",' -nonewline;
    Write-Host '"' -nonewline; Write-Host $cert.get_issuer() -nonewline; Write-Host '",' -nonewline;
    Write-Host '"' -nonewline; Write-Host $cert2.PublicKey.Key.KeySize -nonewline; Write-Host '",' -nonewline;
    Write-Host '"' -nonewline; Write-Host $cert.getserialnumberstring() -nonewline; Write-Host '",' -nonewline;
    Write-Host '"' -nonewline; Write-Host $validfrom -nonewline; Write-Host '",' -nonewline;
    Write-Host '"' -nonewline; Write-Host $validto -nonewline; Write-Host '",' -nonewline;
    Write-Host '"' -nonewline; Write-Host $selfsigned -nonewline; Write-Host '",' -nonewline;
    Write-Host '"' -nonewline; Write-Host $cert2.SignatureAlgorithm.FriendlyName -nonewline; Write-Host '"';

} finally {
    $Connection.Close()
} }
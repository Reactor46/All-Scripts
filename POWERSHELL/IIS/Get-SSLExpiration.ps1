param(
[parameter(Mandatory=$true)]
[string]$computername,
[parameter(Mandatory=$true)]
[int]$port)

#Create a TCP Socket to the computer and a port number
[Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
[System.Net.ServicePointManager]::SecurityProtocol =[System.Net.SecurityProtocolType]'Tls, Tls11, Tls12, Tls13'
$tcpsocket = New-Object Net.Sockets.TcpClient($computerName, $port)
$ProtocolVersion='TLS12'
#test if the socket got connected
if(!$tcpsocket)
{
    Write-Error "Error Opening Connection: $port on $computername Unreachable"
    exit 1
}
else
{
    #Socket Got connected get the tcp stream ready to read the certificate
    write-host "Successfully Connected to $computername on $port" -ForegroundColor Green -BackgroundColor Black
    $tcpstream = $tcpsocket.GetStream()
    Write-host "Reading SSL Certificate...." -ForegroundColor Yellow -BackgroundColor Black 
    #Create an SSL Connection 
    $sslStream = New-Object System.Net.Security.SslStream($stream, $false, ({ $True } -as [Net.Security.RemoteCertificateValidationCallback]))
    #Force the SSL Connection to send us the certificate
    $sslStream.AuthenticateAsClient($computername, $null, [System.Security.Authentication.SslProtocols]$ProtocolVersion, $false)     

    #Read the certificate
    $certinfo = New-Object System.Net.Security.SslStream($sslStream.RemoteCertificate)
}

$returnobj = new-object psobject
$returnobj |Add-Member -MemberType NoteProperty -Name "FriendlyName" -Value $certinfo.FriendlyName
$returnobj |Add-Member -MemberType NoteProperty -Name "SubjectName" -Value $certinfo.SubjectName
$returnobj |Add-Member -MemberType NoteProperty -Name "HasPrivateKey" -Value $certinfo.HasPrivateKey
$returnobj |Add-Member -MemberType NoteProperty -Name "EnhancedKeyUsageList" -Value $certinfo.EnhancedKeyUsageList
$returnobj |Add-Member -MemberType NoteProperty -Name "DnsNameList" -Value $certinfo.DnsNameList
$returnobj |Add-Member -MemberType NoteProperty -Name "SerialNumber" -Value $certinfo.SerialNumber
$returnobj |Add-Member -MemberType NoteProperty -Name "Thumbprint" -Value $certinfo.Thumbprint
$returnobj
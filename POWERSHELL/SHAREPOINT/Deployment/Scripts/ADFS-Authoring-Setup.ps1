$certPath = 'D:\Certificates\TokenSign.cer'

$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certPath)

$map1 = New-SPClaimTypeMapping 'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/windowsaccountname' -IncomingClaimTypeDisplayName 'windowsaccountname' -SameAsIncoming

$realm = 'urn:' + $env:ComputerName + ':adfs'

$signinurl = 'https://login.ksnet.com/adfs/ls/'

$ap = New-SPTrustedIdentityTokenIssuer -Name 'ADFS20' -Description 'ADFS 2.0 Federated Server' -Realm $realm -ImportTrustCertificate $cert -ClaimsMappings $map1 -SignInUrl $signinurl -IdentifierClaim $map1.InputClaimType

$uri = new-object System.Uri('https://authoring2.kelsey-seybold.com')

$ap.ProviderRealms.Add($uri, 'urn:' + $env:ComputerName + ':authoring')

$ap.Update()

New-SPTrustedRootAuthority 'ADFS Token Signing Trusted Root Authority' -Certificate $cert

$certPath = 'D:\Certificates\GODaddySecureCA.cer'
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certPath)
New-SPTrustedRootAuthority 'GoDaddy Secure Trusted Root Authority' -Certificate $cert

$certPath = 'D:\Certificates\GODaddyClass2CA.cer'
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certPath)
New-SPTrustedRootAuthority 'GoDaddy Class2 Trusted Root Authority' -Certificate $cert


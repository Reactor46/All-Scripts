New-SelfSignedCertificate -CertStoreLocation Cert:\CurrentUser\My -Type CodeSigningCert -Subject "CN=Go Live UK Ltd" -FriendlyName "Go Live UK Ltd PowerShell Scripts Signing" -NotAfter 16-07-2027
$cert = New-SelfSignedCertificate -CertStoreLocation Cert:\CurrentUser\My -Type CodeSigningCert -Subject "CN=Go Live UK Ltd" -FriendlyName "Go Live UK Ltd PowerShell Scripts Signing" -NotAfter 16-07-2027
Move-Item -Path $cert.PSPath -Destination "Cert:\CurrentUser\Root"
$CodeCert = Get-ChildItem -Path "Cert:\CurrentUser\Root" -CodeSigningCert | Where-Object {$_.Subject -Match "CN=Go Live UK Ltd"}
Set-AuthenticodeSignature -Certificate $CodeCert -FilePath Location_to_Script.ps1

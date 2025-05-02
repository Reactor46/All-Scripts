$OLDCertificateThumbprint = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
$NEWCertificateThumbprint = "YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY"

#Show bindings where the old certificate is in use
Get-WebBinding | Where-Object { $_.certificateHash -eq $OLDCertificateThumbprint} | Format-Table

#Select bindings where the old certificate is in use and attach the new certificate
Get-WebBinding | Where-Object { $_.certificateHash -eq $OLDCertificateThumbprint} | ForEach-Object {
        Write-Host "Working on"  $_ 
        $_.RemoveSslCertificate()
        $_.AddSslCertificate($NEWCertificateThumbprint, 'My') 
        }

#Show bindings where the new certificate is in use
Get-WebBinding | Where-Object { $_.certificateHash -eq $NEWCertificateThumbprint} | Format-Table

#Validate if the old certificate is in use. If it is not in use proceed to delete it
$OldCertUse = Get-WebBinding | Where-Object { $_.certificateHash -eq $OLDCertificateThumbprint}

   if ($OldCertUse -eq $null){
   Get-ChildItem Cert:\LocalMachine\My\$OLDCertificateThumbprint | Remove-Item -Verbose
	      
}
  else {
      Write-Host "Error: The certificate could not be deleted because it is in use by the following Bindings:" -fore Red
      Get-WebBinding | Where-Object { $_.certificateHash -eq $OLDCertificateThumbprint} | Format-Table
     }
#Reference 'Mono.Security.dll' from http://www.mono-project.com/download/

param(
[switch]$CA,
[switch]$NDES,
[switch]$HSM
)

[string] $Server = $env:COMPUTERNAME
$RootCAName = "*NameOfYourRootCA*"

$ArrayOfChecksWithWarningState = @()
$ArrayOfCheckWithUnknowState = @()
$ArrayOfCheckWithOKState = @()

$cadmsignature = @"
[DllImport("Certadm.dll", CharSet=CharSet.Auto, SetLastError=true)]
public static extern bool CertSrvIsServerOnline(
    string pwszServerName,
    ref bool pfServerOnline
);
"@

#PSScriptRoot variable only available in PS>3.0
if (!(Get-Variable -Name "PSScriptRoot" -ErrorAction SilentlyContinue))
{
 $PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
}

function WriteLog ([string] $logline)
{
 $LogFile = $PSScriptRoot + "\ADCS_Health_Check.log"
 $BackupFile = $PSScriptRoot + "\ADCS__Health_Check.bck"
 
 if (Test-Path $LogFile)
 {
  if ((Get-ChildItem $LogFile).Length -gt 2097152)
  {
   Copy-Item $LogFile $BackupFile
   Remove-Item $LogFile
  }
 }
 $Output = (get-date).ToShortDateString() + " " + (get-date).ToShortTimeString() + ": " + $logline
 Out-File -InputObject $Output -Append -FilePath $LogFile
}

if ($CA)
{
 WriteLog ("ADCS CA health check started.")

 #Get type of the certificate server - http://www.systemcentercentral.com/how-to-determine-the-type-of-certificate-authority-ca-you-have/
 [int] $CATypeNr = $null
 $CAName = $null
 $CAType = $null

 if (Test-Path "HKLM:SYSTEM\CurrentControlSet\Services\CertSvc\Configuration")
 {
  $CAName = (Get-ItemProperty "HKLM:SYSTEM\CurrentControlSet\Services\CertSvc\Configuration" "Active").Active
  WriteLog ("CA name..." + $CAName)
  if (Test-Path "HKLM:SYSTEM\CurrentControlSet\Services\CertSvc\Configuration\$CAName")
  {
   $CATypeNr = (Get-ItemProperty "HKLM:SYSTEM\CurrentControlSet\Services\CertSvc\Configuration\$CAName" "CAType").CAType
   switch ($CATypeNr)
   {
    0 {$CAType = "Enterprise Root CA"}
    1 {$CAType = "Enterprise Subordinate CA" }
    3 {$CAType = "Stand Alone CA" }
    4 {$CAType = "Stand Alone Subordinate CA"}
    Default {$CAType = "Unknown"}
   }
   WriteLog ("CA type..." + $CAType + ".")
  }
 }
 else
 {
  Writelog ("CA name...unknown.")
  Writelog ("CA type...unknown.")
  Writelog ("Maybe this is a member server with one or more installed ADCS services.")
 }

 #Check if Sub CA is alive
 $ServerStatus = $false
 $hresult = $null
 $CertConfig = $null
 $Config = $null
 $CertAdmin = $null
 $CertRequest = $null
 $retn = $null

 if (($CATypeNr -eq 1) -or ($CATypeNr -eq 4))
 {
  try
  {
   Add-Type -MemberDefinition $cadmsignature -Namespace PKI -Name Certadm #-ErrorAction SilentlyContinue
  }
  catch
  {
   WriteLog ("Can not load certadm.dll (error).")
   Write-Host "CRITICAL: Can not DLL files."
   Exit 2
  }
  $hresult = [PKI.CertAdm]::CertSrvIsServerOnline($Server,[ref]$ServerStatus)
  if ($ServerStatus)
  {
   Writelog ("CA status and service is...online and running.")
   $ArrayOfCheckWithOKState += "CAStatus=OK"
   try
   {
    $CertConfig = New-Object -ComObject CertificateAuthority.Config
    $Config = $CertConfig.GetConfig(0)
    $CertAdmin = New-Object -ComObject CertificateAuthority.Admin.1
    $CertRequest = New-Object -ComObject CertificateAuthority.Request.1
   }
   catch
   {
    WriteLog ("Can not load certadm.dll (error).")
    Write-Host "CRITICAL: Can not CA DLL files."
    Exit 2
   }
   try 
   {
    $retn = $CertAdmin.GetCAProperty($Config,0x6,0,4,0)
    Writelog ("ICertAdminD interface...online.")
	$ArrayOfCheckWithOKState += "ICertAdminInterface=OK"
   }
   catch 
   {
    WriteLog ("ICertAdminD interface error.")
    Write-Host "CRITICAL: ICertAdminD interface down."
    Exit 2
   }
   try 
   {
    $retn = $CertRequest.GetCAProperty($Config,0x6,0,4,0)
    Writelog ("ICertRequestD interface...online.")
	$ArrayOfCheckWithOKState += "ICertRequestInterface=OK"
   }
   catch 
   {
    WriteLog ("ICertRequestD interface error.")
    Write-Host "CRITICAL: ICertRequestD interface down."
    Exit 2
   }
    
   #Check if CA signing certificate is not expired and valid
   $CACert = $null
   $Cert = $null
   $CAChain = $null
  
   Foreach ($CACert in (Get-ChildItem -Path "Cert:\LocalMachine\CA" | Where-Object {$_.Subject -like "*$CAName*"} | Get-Unique))
   {
    if (($CACert.NotAfter).Subtract((Get-Date).Date).Days -lt 14)
    { 
     Writelog ("CA signing certificate expiry date...less then 14 days (error).")
	 Write-Host "CRITICAL: CA certificate expires in less then 14 days."
     Exit 2
    }
	else
	{
	 Writelog ("CA signing certificate expiry date..." + $CACert.NotAfter + ".")
	 $ArrayOfCheckWithOKState += "CACertExpireDate=" + $CACert.NotAfter
	}
    $chain = New-Object System.Security.Cryptography.X509Certificates.X509Chain;
    $chain.Build($CACert) | Out-Null
    if ($chain.ChainElements)
    {
	 Foreach ($Cert in $chain.ChainElements)
	 {
	  if ($Cert.certificate.verify())
	  {
	   Writelog ("CA certificate chain for certificate with thumbprint " + $Cert.certificate.Thumbprint + " is valid.")
	  }
	  else
	  {
	   Writelog ("CA certificate chain for certificate with thumbprint " + $Cert.certificate.Thumbprint + " is invalid (error).")   
	  }
	 }
    }    
    else
    {
     Writelog ("Unable to get certificate chain for certificate with thumbprint " + $CACert.Thumbprint + " (error).")
	 Write-Host "CRITICAL: CA certificate chain broken."
     Exit 2
    } 
   }
  
   #Check if Root signing certificate is not expired and valid
   Foreach ($CACert in (Get-ChildItem -Path "Cert:\LocalMachine\Root" | Where-Object {$_.Subject -like $RootCAName} | Get-Unique))
   {
    if (($CACert.NotAfter).Subtract((Get-Date).Date).Days -lt 14)
    {
     Writelog ("Root CA signing certificate expiry date...less then 14 days (error).")
	 Write-Host "CRITICAL: Root CA certificate expires in less then 14 days."
     Exit 2
    }
	else
	{
	 Writelog ("Root CA signing certificate expiry date..." + $CACert.NotAfter + ".")
	 $ArrayOfCheckWithOKState += "RootCACertExpireDate=" + $CACert.NotAfter
	}
    $chain = New-Object System.Security.Cryptography.X509Certificates.X509Chain;
    $chain.Build($CACert) | Out-Null
    if ($chain.ChainElements)
    {
	 Foreach ($Cert in $chain.ChainElements)
	 {
	  if ($Cert.certificate.verify())
	  {
	   Writelog ("Root CA certificate chain for certificate with thumbprint " + $Cert.certificate.Thumbprint + " is valid.")  
	  }
	  else
	  {
	   Writelog ("Root CA certificate chain for certificate with thumbprint " + $Cert.certificate.Thumbprint + " is invalid (error).")   
	  }
	 } 
    }    
    else
    {
     Writelog ("Unable to get certificate chain for Root CA certificate with thumbprint " + $CACert.Thumbprint + " (error).")
	 Write-Host "CRITICAL: Root CA certificate chain broken."
     Exit 2
    }
   }
   
   #Check CRL in http and file locations
   $CRLLocations = @()
   $IPAdresses = $null
   $Hostname = $null
   $CRLAdress = $null
   $Download = $null
   $CRL = $null
  
   try
   {
    $CRLLocations = $CertAdmin.GetCAProperty($config,0x29,0,4,0).split("`n", [StringSplitOptions]::RemoveEmptyEntries)
   }
   catch
   {
    WriteLog ("Can not get CRL information from the CA properties (error).")
   }
   if ($CRLLocations.Length -gt 0)
   {
    For ($i=0; $i -ne $CRLLocations.Length; $i++)
    {
	 if (!($CRLLocations[$i].Contains("ldap")))
	 {
	  $Hostname = $CRLLocations[$i].Substring($CRLLocations[$i].IndexOf("//") + 2)
	  $Hostname = $Hostname.Remove($Hostname.IndexOf("/"))
	  $IPAdresses = [System.Net.Dns]::Resolve($Hostname)
	  Foreach ($IP in $IPAdresses.AddressList)
	  {
	   if (Test-Connection $IP.IPAddressToString -Count 1 -Quiet)
	   {
	    if ($CRLLocations[$i].Contains("http"))
	    {
		 #CRL file download via http
		 $CRLAdress = "http://" + $IP.IPAddressToString + "/" + $CRLLocations[$i].SubString($CRLLocations[$i].LastIndexOf("/") +1)
		 try
		 {
		  $StatusCode = (Invoke-WebRequest $CRLAdress -DisableKeepAlive -TimeoutSec 10 -UseBasicParsing).StatusCode
		  if ($StatusCode -eq 200)
		  {
		   Add-Type -Path ($PSScriptRoot + "\Mono.Security.dll")
		   if (Test-Path "C:\Windows\temp\crlfile.crl") {Remove-Item -Path "C:\Windows\temp\crlfile.crl"}
		   Invoke-WebRequest $CRLAdress -DisableKeepAlive -TimeoutSec 10 -OutFile "C:\Windows\temp\crlfile.crl" -UseBasicParsing
		   Writelog ("Successful download of the CRL file " + $CRLAdress + " - Statuscode " + $StatusCode + ".")
		   $ArrayOfCheckWithOKState += "CRLDownload=OK"
		   $CRL = [Mono.Security.X509.X509Crl]::CreateFromFile("C:\Windows\temp\crlfile.crl")
		   if ($CRL.NextUpdate.Subtract((Get-Date).Date).Days -lt 0)
		   {
		    Writelog ("CRL file " + $CRLAdress + " is expired (" + $CRL.NextUpdate.ToShortDateString() + ") (error).")
			Write-Host "CRITICAL: CRL" + $CRLAdress + "  is expired."
     	    Exit 2
		   }
		   else
		   {
		    Writelog ("CRL file " + $CRLAdress + " is valid, next publishment on " + $CRL.NextUpdate.ToShortDateString() + ".")
			$ArrayOfCheckWithOKState += "CRLNextPublishDate=" + $CRL.NextUpdate.ToShortDateString()
		   }
		  }
		  else
		  {
		   WriteLog ("Can not download CRL file " + $CRLAdress + " - Statuscode " + $StatusCode + " (error).")
		   Write-Host "CRITICAL: CRL file " + $CRLAdress + " download not possible."
           Exit 2
		  }
		 }
		 catch
		 {
		  WriteLog ("Can not download and check CRL file(s) from http location(s).")
		  Write-Host "CRITICAL: CRL file download not possible."
          Exit 2
		 }
	    }
	    else
	    {
	     try
		 {
		  $CRL = [Mono.Security.X509.X509Crl]::CreateFromFile($CRLLocations[$i])
		  if ($CRL.NextUpdate.Subtract((Get-Date).Date).Days -lt 0)
		   {
		    Writelog ("CRL file " + $CRLLocations[$i] + " is expired (" + $CRL.NextUpdate.ToShortDateString() + ") (error).")
			Write-Host "CRITICAL: CRL " + $CRLLocations[$i] + " is expired " + $CRL.NextUpdate.ToShortDateString()
            Exit 2
		   }
		   else
		   {
		    Writelog ("CRL file " + $CRLLocations[$i] + " is valid, next publishment on " + $CRL.NextUpdate.ToShortDateString() + ".")
			$ArrayOfCheckWithOKState += "CRLExpireDate=" + $CRL.NextUpdate.ToShortDateString()
		   }
		 }
		 catch
		 { 
		  WriteLog ("Can not download and check CRL file(s) from file location(s).")
		  Write-Host "CRITICAL: Can not verify CRL " + $CRLLocations[$i]
          Exit 2
		 }
	    } 
	   }
	   else
	   {
	    Writelog ("CRL distribution point $Hostname with IP " + $IP.IPAddressToString + " is unavailable (error).")
		Write-Host "CRITICAL: CRL distribution point with IP " + $IP.IPAddressToString + " is unavailable."
        Exit 2
	   }
	  }
	 }
	 else
	 {
	  if ($CRLLocations[$i].Contains("?"))
	  {
	   $CRLAdress = $CRLLocations[$i].Remove($CRLLocations[$i].IndexOf("?"))
	   $CRLAdress = $CRLAdress.Replace("%20", " ")
	   $CRLAdress = $CRLAdress.Replace("ldap:///", "")
	  }
	  try
	  {
	   if (((([Mono.Security.X509.X509Crl]([ADSI]"LDAP://$CRLAdress").certificateRevocationList.Item(0)).NextUpdate).Subtract((Get-Date).Date).Days -lt 0))
       { 
	    Writelog ("CRL file " + $CRLAdress + " is expired (" + $CRL.NextUpdate.ToShortDateString() + ") (error).")
		Write-Host "CRITICAL: CRL" + $CRLAdress + "  is expired."
     	Exit 2
	   }
	   else
	   {
	    Writelog ("CRL file " + $CRLAdress + " is valid, next publishment on " + $CRL.NextUpdate.ToShortDateString() + ".")
		$ArrayOfCheckWithOKState += "CRLNextPublishDate=" + $CRL.NextUpdate.ToShortDateString()
	   }
	  }
	  catch
	  {
	   Writelog ("Can not download and check CRL file from LDAP location (error).")
	   Write-Host "CRITICAL: Can not verify CRL " + $CRLLocations[$i]
       Exit 2
	  }
     }
    }
   }
   else
   {
    Writelog ("No CRL locations found (error).")
    Write-Host "CRITICAL: Can not get CRL locations from the CA configuration."
    Exit 2
   }
  
   #OCSP Checking - the system account of the monitoring system needs DCOM security permissions on the OCSP systems!
   $OCSPURLS = $null
   $Hostname = $null
   $OCSPAdmin = $null
  
   try
   {
    $OCSPURLS = $CertAdmin.GetCAProperty($Config,0x2B,0,4,0).split("`n", [StringSplitOptions]::RemoveEmptyEntries)
   }
   catch
   {
    Writelog ("Can not get OCSP URLs (error).")
	Write-Host "CRITICAL: Can not get OCSP URLs from CA configuration."
    Exit 2
   }
   if ($OCSPURLS -ne $null)
   {
    $OCSPAdmin = New-Object -ComObject CertAdm.OCSPAdmin
    Foreach ($OCSPURL in $OCSPURLS)
    {
     $Hostname = $OCSPURL.Substring($OCSPURL.IndexOf("//") + 2)
	 $Hostname = $Hostname.Remove($Hostname.IndexOf("/"))
	 $IPAdresses = [System.Net.Dns]::Resolve($Hostname)
	 Foreach ($IP in $IPAdresses.AddressList)
	 {
	  try
	  {
	   $OCSPAdmin.Ping($IP)
	   Writelog ("OCSP service with IP adress " + $IP + " successful responding.")
	   $ArrayOfCheckWithOKState += "OCSPInterface=OK"
	  }
	  catch
	  {
	   Writelog ("OCSP service with IP adress " + $IP + " not responding (error).")
	   Write-Host "CRITICAL: OCSP interface not responding."
       Exit 2
	  }
	 }
    } 
   }
  }
  else
  {
   Writelog ("CA status and service is...offline and stopped.")
   Write-Host "CRITICAL: CA service offline and stopped."
   Exit 2
  }
 }
 Writelog ("ADCS CA health check finished.")
}

if ($NDES)
{
 Writelog ("ADCS NDES health check started.")
 #MSCEP check
 $HTMLOutput = $null
  
 [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
 $wc = New-Object System.Net.WebClient
 try
 {
  if (new-object System.Net.Sockets.TcpClient($Server, "443").Connected)
  {
    $HTMLOutput = $wc.DownloadString("https://$Server/CertSrv/mscep")
    if ($HTMLOutput.Contains("This URL is used by network devices to submit certificate requests."))
    { 
     WriteLog ("Successful connection to MSCEP web interface on " + $Server + ".")
	 $ArrayOfCheckWithOKState += "NDESWebInterface=OK"
	 if (Test-Path "C:\Windows\temp\pkiclient") {Remove-Item "C:\Windows\temp\pkiclient"}
	 try
	 {
	  $wc.DownloadFile("https://$Server/certsrv/mscep/mscep.dll/pkiclient.exe?operation=GetCACert&message=any", "C:\Windows\Temp\pkiclient")
	  if ((Test-Path "C:\Windows\Temp\pkiclient") -and ((Get-ChildItem "C:\Windows\temp\pkiclient").Length -ge 2000))
	  {
	   WriteLog ("Successful file download from MSCEP web interface on " + $Server + ".")
	   $ArrayOfCheckWithOKState += "NDESWebDownload=OK"
	  }
	  else
	  {
	   WriteLog ("Can not download file from MSCEP web interface on " + $Server + " (error).")
	   Write-Host "CRITICAL: Can not download file from NDES web interface."
   	   Exit 2
	  }
	 }
	 catch
	 {
	  WriteLog ("Can not access MSECP web interface on " + $Server + " for file download (error).")
	  Write-Host "CRITICAL: Can not access NDES web interface for file download."
   	  Exit 2
	 }
    }
    else
    {
     WriteLog ("Can not access MSECP web interface on " + $Server + " (error).")
	 Write-Host "CRITICAL: Can not access NDES web interface."
   	 Exit 2
    }
   }
   else
   {
    Writelog ("Can not access MSECP web interface on " + $Server + " (error).")
	Write-Host "CRITICAL: Can not access NDES web interface port."
   	Exit 2
   }
  }
 catch
 {
  Writelog ("Can not access MSCEP web interface on " + $Server)
  Write-Host "CRITICAL: Can not access NDES web interface."
  Exit 2
 }
  
 #Verify MSCEP certificates
 $MSCEPCert = $null
 
 Foreach ($MSCEPCert in (Get-ChildItem -Path "Cert:\LocalMachine\MY" | Get-Unique))
 {
  if (($MSCEPCert.NotAfter).Subtract((Get-Date).Date).Days -lt 14)
  {
   Writelog ("MSCEP signing certificate expiry date...less then 14 days (error).")
   Write-Host "CRITICAL: NDES certificate expired."
   Exit 2
  }
  else
  {
   Writelog ("MSCEP signing certificate expiry date..." + $MSCEPCert.NotAfter + ".")
   $ArrayOfCheckWithOKState += "NDESCertExpireDate=" + $MSCEPCert.NotAfter
  }
  $chain = New-Object System.Security.Cryptography.X509Certificates.X509Chain;
  $chain.Build($MSCEPCert) | Out-Null
  if ($chain.ChainElements)
  {
   Foreach ($Cert in $chain.ChainElements)
   {
	if ($Cert.certificate.verify())
	{
	 Writelog ("MSCEP certificate chain for certificate with thumbprint " + $Cert.certificate.Thumbprint + " is valid.")
	 $ArrayOfCheckWithOKState += "NDESCertChain=OK"
	}
	else
	{
	 Writelog ("MSCEP certificate chain for certificate with thumbprint " + $Cert.certificate.Thumbprint + " is invalid (error).")
	 Write-Host "CRITICAL: NDES certificate chain invalid."
     Exit 2
	}
   }
  }   
  else
  {
   Writelog ("Unable to get certificate chain for certificate with thumbprint " + $MSCEPCert.Thumbprint + " (error).")
  }
 }
 Writelog ("ADCS NDES health check finished.")
}

if ($HSM)
{
 $HSMIPAdresses = $null
 $HSMPorts = $null
 $socket = $null
 Writelog ("ADCS HSM health check started.")
 
 if (Test-Path -Path "XXX\nCipher\Key Management Data\config\config")
 {
  $HSMIPAdresses = ((select-string "XXX\nCipher\Key Management Data\config\config" -pattern "remote_ip=10.").Line | sort-object | Get-Unique | Where-Object {$_ -notlike "*#*"}).Replace("remote_ip=", "")
  $HSMPorts = ((select-string "XXX\nCipher\Key Management Data\config\config" -pattern "remote_port=").Line | sort-object | Get-Unique | Where-Object {$_ -notlike "*#*"}).Replace("remote_port=", "")
  if ($HSMIPAdresses.length -gt 0)
  {
   Foreach ($HSMIPAdress in $HSMIPAdresses)
   {
    Foreach ($HSMPort in $HSMPorts)
	{
	 try
	 {
	  $socket = new-object System.Net.Sockets.TcpClient($HSMIPAdress, $HSMPort)
	  if ($socket.Connected)
	  {
	   Writelog ("Succesful bind to HSM " + $HSMIPAdress + " on tcp port " + $HSMPort)
	   $ArrayOfCheckWithOKState += "HSMInterface=OK"
	   $socket.Close()
	  }
	  else
	  {
	   Writelog ("Can not bind to HSM " + $HSMIPAdress + " on tcp port " + $HSMPort + " (error).")
	   Write-Host "CRITICAL: Can not bind to HSM $HSMIPAdress on tcp port $HSMPort"
   	   Exit 2
	  }
	 }
	 catch
	 {
	  Writelog ("Can not bind to HSM " + $HSMIPAdress + " on tcp port " + $HSMPort + " (error).")
	  Write-Host "CRITICAL: Can not bind to HSM $HSMIPAdress on tcp port $HSMPort"
   	  Exit 2
	 }
	}
   }
  }
  else
  {
   Writelog ("Can not find any ip adress in the HSM config file.")
  }
 }
 else
 { 
  Writelog ("No nCipher config file found.")
 } 
 if (Test-Path -Path "XXX\nCipher\nfast\bin\enquiry.exe")
 {
  try
  {
   $HSMenquiryOutput = $null
   $pinfo = New-Object System.Diagnostics.ProcessStartInfo
   $pinfo.FileName = "XXX\nCipher\nfast\bin\enquiry.exe"
   $pinfo.Arguments = ""
   $pinfo.UseShellExecute = $false
   $pinfo.CreateNoWindow = $true
   $pinfo.RedirectStandardOutput = $true
   $pinfo.RedirectStandardError = $true
   $process = New-Object System.Diagnostics.Process
   $process.StartInfo = $pinfo
   $process.Start() | Out-Null
   sleep -Seconds 5
   if (!$process.HasExited) {$process.Kill()}
   if (Test-Path -Path "C:\Windows\temp\enquiry.txt") {Remove-Item -Path "C:\Windows\temp\enquiry.txt"}
   $process.StandardOutput.ReadToEnd() | Out-File "C:\Windows\temp\enquiry.txt"
  }
  catch
  {
   Writelog ("Thales HSM enquiry tool throws an error.")
  }
  if (Test-Path -Path "C:\Windows\temp\enquiry.txt") {$HSMenquiryOutput = Get-Content "C:\Windows\temp\enquiry.txt" | Select-String -Pattern "mode" | Sort-Object | Get-Unique}
  if ($HSMenquiryOutput.length -gt 0) 
  {
   Foreach ($Output in $HSMenquiryOutput.Line)
   {
    if ($Output.Contains("operational"))
    {
	 Writelog ("Thales HSM mode is operational.")
	 $ArrayOfCheckWithOKState += "HSMMode=OK"
    }
    else
    {
	 Writelog ("Thales HSM mode is NOT operational, run the Thales HSM enquiry tool on the CA and verify the output (error).")
	 Write-Host "CRITICAL: Thales HSM mode is NOT operational"
	 Exit 2
    }
   }
  }
  if (Test-Path -Path "C:\Windows\temp\enquiry.txt") {$HSMenquiryOutput = Get-Content "C:\Windows\temp\enquiry.txt" | Select-String -Pattern "connection status" | Sort-Object | Get-Unique}
  if ($HSMenquiryOutput.length -gt 0) 
  {
   Foreach ($Output in $HSMenquiryOutput.Line)
   {
    if ($Output.Contains("OK"))
   {
	 Writelog ("Thales HSM connection status is OK.") 
	 $ArrayOfCheckWithOKState += "HSMConnectionStatus=OK"
   }
   else
   {
	 Writelog ("Thales HSM connection status is NOT OK, run the Thales HSM enquiry tool on the CA and verify the output (error).")
	 Write-Host "CRITICAL: Thales HSM connection status is NOT OK"
	 Exit 2
   }
   }
  }
 }
 else
 {
  Writelog ("Thales HSM enquiry tool not found.")
 }
 if (Test-Path -Path "XXX\nCipher\nfast\bin\nfkminfo.exe")
 {
  try
  {
   $HSMnfkminfoOutput = $null
   $Flag = $true
   $pinfo = New-Object System.Diagnostics.ProcessStartInfo
   $pinfo.FileName = "XXX\nCipher\nfast\bin\nfkminfo.exe"
   $pinfo.Arguments = ""
   $pinfo.UseShellExecute = $false
   $pinfo.CreateNoWindow = $true
   $pinfo.RedirectStandardOutput = $true
   $pinfo.RedirectStandardError = $true
   $process = New-Object System.Diagnostics.Process
   $process.StartInfo = $pinfo
   $process.Start() | Out-Null
   sleep -Seconds 5
   if (!$process.HasExited) {$process.Kill()}
   if (Test-Path -Path "C:\Windows\temp\nfkminfo.txt") {Remove-Item -Path "C:\Windows\temp\nfkminfo.txt"}
   $process.StandardOutput.ReadToEnd() | Out-File "C:\Windows\temp\nfkminfo.txt"
  }
  catch
  {
   Writelog ("Thales HSM nfkminfo tool throws an error.")
  }
  if (Test-Path -Path "C:\Windows\temp\nfkminfo.txt") {$HSMnfkminfoOutput = Get-Content "C:\Windows\temp\nfkminfo.txt" | Select-String -Pattern "state" | Sort-Object | Get-Unique}
  if ($HSMnfkminfoOutput.length -gt 0) 
  {
   Foreach ($Output in $HSMnfkminfoOutput.Line)
   {
    if ($Output.Contains("Initialised Usable"))
    {
	 Writelog ("Thales HSM is Initialised Usable.")
	 $ArrayOfCheckWithOKState += "HSMState=OK"
	 $Flag = $false
	 break
    }
   }
   if ($Flag)
   {
    Writelog ("Thales HSM is NOT Initialised Usable, run the Thales HSM nfkminfo tool on the CA and verify the output (error).")
	Write-Host "CRITICAL: Thales HSM NOT Initialised Usable"
	Exit 2
   }
  }
  if (Test-Path -Path "C:\Windows\temp\nfkminfo.txt") {$HSMnfkminfoOutput = Get-Content "C:\Windows\temp\nfkminfo.txt" | Select-String -Pattern "error" | Sort-Object | Get-Unique}
  if ($HSMnfkminfoOutput.length -gt 0) 
  {
   Foreach ($Output in $HSMnfkminfoOutput.Line)
   {
    if ($Output.Contains("OK"))
    {
	 Writelog ("Thales HSM error state is OK.")
	 $ArrayOfCheckWithOKState += "HSMErrorState=OK"
    }
	else
	{
	 Writelog ("Thales HSM error state is NOT OK, run the Thales HSM nfkminfo tool on the CA and verify the output (error).")
	 Write-Host "CRITICAL: Thales HSM error state is NOT OK"
	 Exit 2
	}
   }
  }
 }
 else
 {
  Writelog ("Thales HSM nfkminfo tool not found.")
 } 
 Writelog ("ADCS HSM health check finished.")
}


 #Nagius output
 $Element = $null
 $FirstLine = $null
 $SecondLine = $null
 
 if ($ArrayOfChecksWithWarningState.Count -gt 0) 
 {
  $FirstLine = "WARNING - TOTAL=" + ($ArrayOfCheckWithOKState.Count + $ArrayOfChecksWithWarningState.Count + $ArrayOfCheckWithUnknowState.Count) + ", OK="+$ArrayOfCheckWithOKState.Count+", WARNING="+$ArrayOfChecksWithWarningState.Count+", UNKNOWN="+$ArrayOfCheckWithUnknowState.Count
  Write-Host -Object $FirstLine
  #if ($ArrayOfCheckWithOKState.Count -gt 0) {Write-Host [system.string]::Join(",", $ArrayOfCheckWithOKState)}
  if ($ArrayOfChecksWithWarningState.Count -gt 0) 
  {
   $SecondLine = [system.string]::Join(", ", $ArrayOfChecksWithWarningState)
   Write-Host -Object $SecondLine
  }
  if ($ArrayOfCheckWithUnknowState.Count -gt 0) 
  {
   $SecondLine = [system.string]::Join(", ", $ArrayOfCheckWithUnknowState)
   Write-Host -Object $SecondLine
  }
  exit 1
 }
 else
 {
  if ($ArrayOfCheckWithUnknowState.Count -gt 0) 
  {
   $FirstLine = "UNKNOWN - TOTAL=" + ($ArrayOfCheckWithOKState.Count + $ArrayOfChecksWithWarningState.Count + $ArrayOfCheckWithUnknowState.Count) + ", OK="+$ArrayOfCheckWithOKState.Count+", WARNING="+$ArrayOfChecksWithWarningState.Count+", UNKNOWN="+$ArrayOfCheckWithUnknowState.Count
   Write-Host -Object $FirstLine 
   #if ($ArrayOfCheckWithOKState.Count -gt 0) {Write-Host [system.string]::Join(",", $ArrayOfCheckWithOKState)}
   #if ($ArrayOfChecksWithWarningState.Count -gt 0) {Write-Host [system.string]::Join(",", $ArrayOfChecksWithWarningState)}
   if ($ArrayOfCheckWithUnknowState.Count -gt 0) 
   {
    $SecondLine = [system.string]::Join(", ", $ArrayOfCheckWithUnknowState)
    Write-Host -Object $SecondLine 
   }   
   exit 3
  } 
  else 
  {
   $FirstLine = "OK - TOTAL=" + ($ArrayOfCheckWithOKState.Count + $ArrayOfChecksWithWarningState.Count + $ArrayOfCheckWithUnknowState.Count) + ", OK="+$ArrayOfCheckWithOKState.Count+", WARNING="+$ArrayOfChecksWithWarningState.Count+", UNKNOWN="+$ArrayOfCheckWithUnknowState.Count
   Write-Host -Object $FirstLine
   if ($ArrayOfCheckWithOKState.Count -gt 0) 
   {
    $SecondLine = [system.string]::Join(", ", $ArrayOfCheckWithOKState)
    Write-Host -Object $SecondLine
   }
   #if ($ArrayOfChecksWithWarningState.Count -gt 0) {Write-Host [system.string]::Join(",", $ArrayOfChecksWithWarningState)}
   #if ($ArrayOfCheckWithUnknowState.Count -gt 0) {Write-Host [system.string]::Join(",", $ArrayOfCheckWithUnknowState)}  
   exit 0
  }
 }
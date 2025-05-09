#************************************************
# ListMyStoreCertDetails.ps1
# Version 1.0
# Date: 08/24/2017
# Author: Tim Springston [MS]
# Description:  This script exports all user and computer My store certificate details to a text file for review.
#************************************************

function ExportAllMyStoreCerts
{ 
	$ExportFile = $Pwd.path + "\CertificatesList.txt"
	$MyComputerStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("My","LocalMachine")    
	$MyUserStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("My","CurrentUser")
	$MyComputerStore.Open("ReadOnly")
	$MyUserStore.Open("ReadOnly")
	$MyComputerStoreCount = $MyComputerStore.Certificates.Count
	$MyUserStoreCount = $MyUserStore.Certificates.Count
	$Now = Get-Date


	"Certificate My Store Details $Now" | Out-File -FilePath $ExportFile 
	"Logged on user $env:USERNAME" | Out-File -FilePath $ExportFile -Append 
	"Logged on user $env:USERDNSDOMAIN" | Out-File -FilePath $ExportFile -Append 
	"Certificate results from $env:Computername"| Out-File -FilePath $ExportFile -Append 
	"*******************************" | Out-File -FilePath $ExportFile -Append 
	"User My Store Certificate Count: $MyUserStoreCount" | Out-File -FilePath $ExportFile -Append 
	"Computer My Store Certificate Count: $MyComputerStoreCount" | Out-File -FilePath $ExportFile -Append 
	
	$CheckStores = @("My")
	$Counter = 1
	get-childitem -path cert:\ -recurse | Where-Object {($_.PSParentPath -ne $null)  -and `
		($_.IssuerName.Name -ne "CN=Root Agency") -and (-not($_.NotAfter -lt $Now)) -and (-not($_.NotBefore -gt $Now)) `
		-and ($CheckStores -contains (Split-Path ($_.PSParentPath) -Leaf))} | % {
	     
		$CertObject = New-Object PSObject 
		$Store = (Split-Path ($_.PSParentPath) -Leaf)
		$StorePath = (($_.PSParentPath).Split("\"))     
		$StoreWorkingContext = $Store
		$StoreContext = Split-Path $_.PSParentPath.Split("::")[-1] -Leaf

		if ($Store -match "My")
	      {
		  add-member -inputobject $CertObject -membertype noteproperty -name "Certificate Number" -value $Counter
	      if ($_.FriendlyName.length -gt 0)
	      {add-member -inputobject $CertObject -membertype noteproperty -name "Friendly Name" -value $_.FriendlyName}
	      else
	      {add-member -inputobject $CertObject -membertype noteproperty -name "Friendly Name" -value "[None]"}
	      #Determine the context (User or Computer) of the certificate store.
	      $StoreWorkingContext = (($_.PSParentPath).Split("\"))
	      $StoreContext = ($StoreWorkingContext[1].Split(":"))
	      add-member -inputobject $CertObject -membertype noteproperty -name "Path" -value $StoreContext[2]
	      add-member -inputobject $CertObject -membertype noteproperty -name "Store" -value $StorePath[$StorePath.count-1]
	      add-member -inputobject $CertObject -membertype noteproperty -name "Has Private Key" -value $_.HasPrivateKey
	      add-member -inputobject $CertObject -membertype noteproperty -name "Serial Number" -value $_.SerialNumber
	      add-member -inputobject $CertObject -membertype noteproperty -name "Thumbprint" -value $_.Thumbprint
	      add-member -inputobject $CertObject -membertype noteproperty -name "Issuer" -value $_.IssuerName.Name
	      add-member -inputobject $CertObject -membertype noteproperty -name "Not Before" -value $_.NotBefore
	      add-member -inputobject $CertObject -membertype noteproperty -name "Not After" -value $_.NotAfter
	      add-member -inputobject $CertObject -membertype noteproperty -name "Subject Name" -value $_.SubjectName.Name
	      if (($_.Extensions | Where-Object {$_.Oid.FriendlyName -match "subject alternative name"}) -ne $null)
	            {add-member -inputobject $CertObject -membertype noteproperty -name "Subject Alternative Name" -value ($_.Extensions | Where-Object {$_.Oid.FriendlyName -match "subject alternative name"}).Format(1)
	            }
	            else
	            {add-member -inputobject $CertObject -membertype noteproperty -name "Subject Alternative Name" -value "[None]"}
	      if (($_.Extensions | Where-Object {$_.Oid.FriendlyName -like "Key Usage"}) -ne $null) 
	            {add-member -inputobject $CertObject -membertype noteproperty -name "Key Usage" -value ($_.Extensions | Where-Object {$_.Oid.FriendlyName -like "Key Usage"}).Format(1)
	            }
	            else
	            {add-member -inputobject $CertObject -membertype noteproperty -name "Key Usage" -value "[None]"}
	      if (($_.Extensions | Where-Object {$_.Oid.FriendlyName -like "Enhanced Key Usage"}) -ne $null)
	            {add-member -inputobject $CertObject -membertype noteproperty -name "Enhanced Key Usage" -value ($_.Extensions | Where-Object {$_.Oid.FriendlyName -like "Enhanced Key Usage"}).Format(1)
	            }
	            else
	            {add-member -inputobject $CertObject -membertype noteproperty -name "Enhanced Key Usage" -value "[None]"}
	      if (($_.Extensions | Where-Object {$_.Oid.FriendlyName -match "Certificate Template Information"}) -ne $null)
	            {add-member -inputobject $CertObject -membertype noteproperty -name "Certificate Template Information" -value ($_.Extensions | Where-Object {$_.Oid.FriendlyName -match "Certificate Template Information"}).Format(1)
	            }
	            else
	            {add-member -inputobject $CertObject -membertype noteproperty -name "Certificate Template Information" -value "[None]"}

	      $ChainObject = New-Object System.Security.Cryptography.X509Certificates.X509Chain($True)
	      $ChainObject.ChainPolicy.RevocationFlag = "EntireChain" #Possible: EndCertificateOnly, EntireChain, ExcludeRoot (default)
	      $ChainObject.ChainPolicy.VerificationFlags = "NoFlag" #http://msdn.microsoft.com/en-us/library/system.security.cryptography.x509certificates.x509verificationflags.aspx 
	      $ChainObject.ChainPolicy.RevocationMode = "Online" #NoCheck, Online (default), Offline.
	      $ChainResult = $ChainObject.Build($_)
	      $ChainCounter = 1
	      ForEach ($ChainResult in $ChainObject.ChainStatus)
	            {
	            $ChainResultStatusString = $ChainResult.Status.ToString()
	            $ChainStatusString = "ChainStatus " + $ChainCounter
	            $ChainResultStatusInfoString = $ChainResult.StatusInformation.ToString()
	            $ChainStatusInfoString = "Chain Status Info " + $ChainCounter
	            add-member -inputobject $CertObject -membertype noteproperty -name $ChainResultStatusString -value $ChainResultStatusInfoString
	            $ChainCounter++
	            if ($ChainResultStatusString -eq 'RevocationStatusUnknown')
	                  {$ChainRevocationProblem = $True}
	            
	            }
	      ForEach ($Extension in $_.Extensions)
	            {
	            if ($Extension.OID.FriendlyName -eq 'Authority Information Access')
	                  {
	                  #Convert the RawData in the extension to readable form.
	                  $FormattedExtension = $Extension.Format(1)
	                  add-member -inputobject $CertObject -membertype noteproperty -name "AIA URLs" -value $FormattedExtension
	                  }
	            if ($Extension.OID.FriendlyName -eq 'CRL Distribution Points')
	                  {
	                  #Convert the RawData in the extension to readable form.
	                  $FormattedExtension = $Extension.Format(1)
	                  add-member -inputobject $CertObject -membertype noteproperty -name "CDP URLs" -value $FormattedExtension
	                  }
	            if ($Extension.OID.Value -eq '1.3.6.1.5.5.7.48.1')
	                  {
	                  #Convert the RawData in the extension to readable form.
	                  $FormattedExtension = $Extension.Format(1)
	                  add-member -inputobject $CertObject -membertype noteproperty -name "OCSP URLs" -value $FormattedExtension
	                  }
	            }

		$CertObject | Out-File -FilePath $ExportFile -Append 
		$CertObject = $null
		$Counter++
		$ChainRevocationProblem = $False
		}
	}
	Write-host "Certificate details exported to $Exportfile`."
}

Clear-Host
ExportAllMyStoreCerts

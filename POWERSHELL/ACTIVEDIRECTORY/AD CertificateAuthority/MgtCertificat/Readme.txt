# =======================================================
 # NAME: MgtCertificate
 # AUTHOR: Olivier GIGANT
 # DATE: 25/05/2016
 #
 # KEYWORDS: Certificate, CA, OpenSSL
 # VERSION 1.1
 #
 #Requirement : 
 # Active Directory Certification Authority Enterprise
 # OpenSSL installed
 # MgtCertificate.psm1 Module Loaded
 # Powershell v2.0
 # =======================================================

 This script usage is : certificate generation with an enterprise Certification Authority
 You use it at your own risk, and you need to adapt it for your usage.
 
 How to Use it : 
 * Unzip Zip File 
 * Adapt your parameters in \Script\MgtCertificate.ps1 on line 19
		### Customize your script here ###
		$PathOpen = "C:\ScriptsToEx\OpenSSLPort\bin\" # Path to OpenSSL Bin Folder with \ at the end
		$ConfOpenSSL = "openssl.cfg" #Name of the OpenSSL Config File in folder above
		$PathFile = "C:\ScriptsToEx\Certificats\" #Folder where certificates will be saved with \ at the end
		$CertO = "MyOrganization" #Will be in the certificate request
		$CertOU = "MyOU" #Will be in the certificate request
		$CertL = "France" #Will be in the certificate request
		$CertST = "France" #Will be in the certificate request
		$CertC="FR" #Will be in the certificate request
		$CertTemplate = "webserver" #Name Of the Certificate Template To Use
		$CAName = "Hostname.Your.domain\caname.domaine.Name #Name of your Certification authority ComputerName\NameOfCa
		$Password = "MyPassword" #Password used when exporting to PFX
		### End of customization ###
 * Launch a Windows Powershell  (if needed authorize the execution of the script)
 * Browse to your folder with Powershell
 
 Case 1 : One certificate to Generate 
	* Call .\MgtCertificate.ps1 your.url.domain.com

 Case 2 : Several certificates to Generate 
	* Edit \Scripts\CertList.csv to enter all your URL, one by line
	* Call .\MgtCertificate.ps1 csv .\CertList.csv

 * All your certificates will be placed in the folder specified in $PathFile
 * This script will generate a request file, export the key file, export to pfx, export to cer 
 
Best Regards

Olivier

 
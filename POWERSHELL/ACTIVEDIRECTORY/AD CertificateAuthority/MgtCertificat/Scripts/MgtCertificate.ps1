# =======================================================
 # NAME: MgtCertificate.ps1
 # AUTHOR: Olivier GIGANT
 # DATE: 17/02/2016
 #
 # KEYWORDS: Certificate, CA, OpenSSL
 # VERSION 1.1
 # 17/02/2016 Module Creation
 # COMMENTS: Desription des traitements
 #
 #Requirement : 
 # Active Directory Certification Authority Enterprise
 # OpenSSL installed
 # MgtCertificate.psm1 Module Loaded
 # Powershell v2.0
 # =======================================================


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


### Defining some parameters for the script ###
$NbErr = 0
$OpenSSL = $PathOpen + "OpenSSL.exe"
$OpenConf = $PathOpen + $ConfOpenSSL
$DateJ = (Get-Date).ToString('yyyy')
$DateCSV = (Get-Date).ToString('yyyyddMM-HHmmss')

cls
Import-Module .\MgtCertificate.psm1 -Force -Verbose

 if (Get-Module -Name "MgtCertificate") {
        Write-Host "Module MgtCertificate Loaded Successfully !" -ForegroundColor Green
    } else {
        Write-Host "Module MgtCertificate Not Loaded - Please check above !" -ForegroundColor Red
        exit

    }

### Reading Parameters from command line ###
$Arg0 = $Args[0]



if ($arg0 -eq "csv") {
    Write-Host "Requesting Multiple Certificate - CSV File will be imported ..." -ForegroundColor Yellow
    $File = $Args[1]

    if ((Test-File $File yes) -eq 1) { $NbErr = 1}
        if ($NbErr -eq 1) {
                Write-Host "CSV File can't be found - Please check your parameters !" -ForegroundColor Red
                Exit
        } Else {
                Write-Host "CSV File is present - Importing ..." -ForegroundColor Green
                $Liste = Import-Csv -Path $File
        }
} else {
    Write-Host "Requesting 1 Certificate .." -ForegroundColor Yellow
    $CSVTemp = $PathFile + $DateCSV + "-Temp.csv"
    Add-Content -path $CSVTemp -value "CertificatName"
    Add-Content -Path $CSVTemp -value $Args[0]
    $Liste = Import-Csv -Path $CSVTemp
}

foreach ($CertificatName in $Liste) {

    ### Variable ###
    $CertificatCN = $CertificatName.CertificatName
    $BaseFile = $PathFile + $CertificatCN + "_" + $DateJ
    $KeyFile = $BaseFile + ".key"
    $ReqFile = $BaseFile + ".req"
    $CerFile = $BaseFile + ".cer"
    $PfxFile = $BaseFile + ".pfx"
    $CnfFile = $BaseFile + ".cfg"
    
    
    if ($CertificatName.OU -ne "") {
        $CertOU2 = $CertificatName.OU
    } else {
        $CertOU2 = $CertOU
    }
    
    $CertSubj = "/C=$CertC/ST=$CertST/L=$CertL/O=$CertO/OU=$CertOU2/CN=$CertificatCN"
    $FriendName = $CertificatCN + "_" + $DateJ

    Write-Host "Generating Certificate for : " $CertificatCN -ForegroundColor DarkMagenta
    Write-host "Certificate OU is : " $CertOU2
    
### Doing some check to prevent errors ###
if ((Test-File $OpenSSL yes) -eq 1) { $NbErr = 1}
if ((Test-File $OpenConf yes) -eq 1) { $NbErr = 1}
if ((Test-File $KeyFile) -eq 1) { $NbErr = 1}
if ((Test-File $ReqFile) -eq 1) { $NbErr = 1}
if ((Test-File $CerFile) -eq 1) { $NbErr = 1}
if ((Test-File $PfxFile) -eq 1) { $NbErr = 1}
if ((Test-File $CnfFile) -eq 1) { $NbErr = 1} 


if ($NbErr -eq 1) {
    Write-Host "Some errors were encountered - please check log above !" -ForegroundColor Red
    Exit
} Else {
    Write-Host "No errors during File checks - Next Steps ..." -ForegroundColor Green
    }

Write-Host "Generating certificate for $CertificatCN" -ForegroundColor Yellow

### Configuring Config File ###
Write-Host "Configuring Config File ..." -ForegroundColor Yellow
New-OpenSSLConf $CnfFile $CertificatCN

### Request File ###
Write-host "Generating Request File ... " -ForegroundColor Yellow

    New-CertRequest $KeyFile $ReqFile $CertSubj $CnfFile $OpenSSL

    if ((Test-File $ReqFile yes) -eq 1) { $NbErr = 1}

if ($NbErr -eq 1) {
    Write-Host "Some errors were encountered - please check log above !" -ForegroundColor Red
    Exit
} Else {
    Write-Host "No errors during request file - Next Steps ..." -ForegroundColor Green
}


### Submit Request to CA ###
Write-host "Generating Certifcate on Certification Authority ... " -ForegroundColor Yellow

    $TemplateName = "CertificateTemplate:"+$CertTemplate
    New-CertSign $TemplateName $CAName $ReqFile $CerFile

    if ((Test-File $CerFile yes) -eq 1) { $NbErr = 1}

    if ($NbErr -eq 1) {
        Write-Host "Some errors were encountered - please check log above !" -ForegroundColor Red
        Exit
    } Else {
        Write-Host "No errors during submission to Certificate Authority - Next Steps ..." -ForegroundColor Green
    }


### Exporting Certificate to PFX ###
Write-host "Exporting Certifcate to pfx... " -ForegroundColor Yellow
    Export-CertToPfx $KeyFile $CerFile $PfxFile $FriendName $Password $OpenSSL
    
        if ((Test-File $PfxFile yes) -eq 1) { $NbErr = 1}

    if ($NbErr -eq 1) {
        Write-Host "Some errors were encountered - please check log above !" -ForegroundColor Red
        Exit
    } Else {
        Write-Host "No errors during the export - Next Steps ..." -ForegroundColor Green
    }
}


### Cleaning CSV Temp File ###
if ($CSVTemp -ne $null) {
    if (Test-Path $CSVTemp) {
        Remove-Item $CSVTemp -Force
        Write-Host "File $CSVTemp removed ... " -ForegroundColor Yellow
        }
    }

### Cleaning CFG Temp File ###
Remove-Item $CnfFile -Force

Write-Host "Thank you for using this script ..." -ForegroundColor Yellow







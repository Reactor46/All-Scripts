# =======================================================
 # NAME: MgtCertificate.psm1
 # AUTHOR: Olivier GIGANT
 # DATE: 17/02/2016
 #
 # KEYWORDS: Certificate, CA, OpenSSL
 # VERSION 1.2
 # V1.1 - 17/02/2016 Module Creation
 # COMMENTS: Desription des traitements
 # V1.2 - 08/05/2017 Adding Config File 
 # COMMENTS : Since Chrome 58 Can't use anymore the SubjectName adding new function
 #
 #Requirement : 
 # Active Directory Certification Authority Enterprise
 # OpenSSL installed
 # Powershell v2.0
 # =======================================================

 function global:Test-File {
 <#
 .SYNOPSIS 
     Check if a file exist and return error and exit script
 .DESCRIPTION 
     Check if a file exist in a specific place. If the file is there, return an error message and quit script
     Otherwise script will continue
 .PARAMETER File
        Complete Filepath
 .NOTES 
     Author : Olivier GIGANT
     Requires : PowerShell V2 
 .EXAMPLE 
     [ps] c:\foo> Test-File C:\Windows\notepad.exe
     true
  #>
  param(
    [Parameter(Mandatory=$true,Position=0)][String[]]$File,
    [Parameter(Mandatory=$false,Position=1)][String[]]$Exist
    )
    if ($Exist -eq "yes") {
        if (Test-Path $File) {
            Write-Host "File : $File exist ! Next Step ..." -ForegroundColor Green

        } else {
            Write-Host "ERROR File : $file is missing ! Please check." -ForegroundColor Red
            Return 1
            
        }


    } else {
            if (Test-Path $File) {
                Write-Host "ERROR File : $File already exist ! Please delete it and try again." -ForegroundColor Red
                Return 1
            
            } else {
                Write-Host "File : $File doesn't exist ! Next Check ..." -ForegroundColor Green
            
            }
    }
 }

 function New-CertRequest {
 <#
 .SYNOPSIS 
     Generate a request for a certificate
 .DESCRIPTION 
     Generate a request for a certificate
 .PARAMETER File
     KeyFile : Key File part of Certificate
     ReqFile : File containing the request
     CertSubj : Subject of Certificate
     OpenConf : OpenSSL Configuration File
     OpenSSL : OpenSSL Executable File

 .NOTES 
     Author : Olivier GIGANT
     Requires : PowerShell V2 - OpenSSL Installed
 .EXAMPLE 
     [ps] c:\foo> New-CertRequest $KeyFile $ReqFile $CertSubj $OpenConf $OpenSSL

  #>
  param(
    [Parameter(Mandatory=$true,Position=0)][String[]]$KeyFile,
    [Parameter(Mandatory=$true,Position=1)][String[]]$ReqFile,
    [Parameter(Mandatory=$true,Position=2)][String[]]$CertSubj,
    [Parameter(Mandatory=$true,Position=3)][String[]]$OpenConf,
    [Parameter(Mandatory=$true,Position=4)][String[]]$OpenSSL
    )
    $ExeFile = ".\" + $OpenSSL
    $CommandLine = "$OpenSSL req -nodes -sha256 -newkey rsa:2048 -keyout `"$KeyFile`" -out `"$ReqFile`" -subj `"$CertSubj`" -config `"$OpenConf`""
    #$CommandLine = "$OpenSSL req -nodes -sha256 -newkey rsa:2048 -keyout `"$KeyFile`" -out `"$ReqFile`" -subj `"$CertSubj`" -config `"$OpenConf`""
    Write-Host "Executing : $CommandLine" -ForegroundColor Cyan
    cmd.exe /C $CommandLine
 
 }

 function New-CertSign {
 <#
 .SYNOPSIS 
     Submit a Certificate request to a specific CA
 .DESCRIPTION 
     Submit a Certificate request to a specific CA
 .PARAMETER File
     TemplateName : Template Name To Use
     CaName : Certication Authority Name to Use
     ReqFile : File containing the request
     CerFile : Certificate File
     OpenConf : OpenSSL Configuration File
     OpenSSL : OpenSSL Executable File
 .NOTES 
     Author : Olivier GIGANT
     Requires : PowerShell V2 - OpenSSL Installed
 .EXAMPLE 
     [ps] c:\foo> New-CertSign $TemplateName $CAName $ReqFile $CerFile $OpenConf $OpenSSL

  #>
  param(
    [Parameter(Mandatory=$true,Position=0)][String[]]$TemplateName,
    [Parameter(Mandatory=$true,Position=1)][String[]]$CAName,
    [Parameter(Mandatory=$true,Position=2)][String[]]$ReqFile,
    [Parameter(Mandatory=$true,Position=3)][String[]]$CerFile

    )
    $CommandLine = "certreq -attrib `"$TemplateName`" -config `"$CAName`" -submit $ReqFile $CerFile"

    Write-Host "Executing : $CommandLine" -ForegroundColor Cyan
    cmd.exe /C $CommandLine

    
 }

 function Export-CertToPfx {
 <#
 .SYNOPSIS 
     Export a certificate to PFX Format
 .DESCRIPTION 
     Export a certificate to PFX Format
 .PARAMETER File
     KeyFile : File containing Private Key of Certificate
     CerFile : Certificate File
     PfxFile : Certificate exported File
     FriendName : Friendly Name of Certificate
     OpenSSL : OpenSSL Executable File
     Password : Password for encrypting pfx
 .NOTES 
     Author : Olivier GIGANT
     Requires : PowerShell V2 - OpenSSL Installed
 .EXAMPLE 
     [ps] c:\foo> Export-CertToPfx $KeyFile $CerFile $PfxFile $FriendName $Password $OpenSSL

  #>
  param(
    [Parameter(Mandatory=$true,Position=0)][String[]]$KeyFile,
    [Parameter(Mandatory=$true,Position=1)][String[]]$CerFile,
    [Parameter(Mandatory=$true,Position=2)][String[]]$PfxFile,
    [Parameter(Mandatory=$true,Position=3)][String[]]$FriendName,
    [Parameter(Mandatory=$true,Position=4)][String[]]$Password,
    [Parameter(Mandatory=$true,Position=5)][String[]]$OpenSSL
    )
    $commandLine = "$OpenSSL pkcs12 -inkey `"$KeyFile`" -in `"$CerFile`" -export -out `"$PfxFile`" -name `"$FriendName`" -password pass:`"$Password`""
    Write-Host "Executing : $CommandLine" -ForegroundColor Cyan
    cmd.exe /C $CommandLine
 }

 function New-OpenSSLConf {
 <#
 .SYNOPSIS 
     Create an OpenSSL Config File for using Subject Alternate Names
 .DESCRIPTION 
     Create an OpenSSL Config File for using Subject Alternate Names
 .PARAMETER File
     Complete Filepath and Subject for Certificat
 .NOTES 
     Author : Olivier GIGANT
     Requires : PowerShell V2 
 .EXAMPLE 
     [ps] c:\foo> New-OpenSSLConf c:\certificate\ca2\test.dev.apps.bdl.cfg test.dev.apps.bdl
     true
  #>
  param(
    [Parameter(Mandatory=$true,Position=0)][String[]]$CnfFile,
    [Parameter(Mandatory=$false,Position=1)][String[]]$CertificatCN
    )
    Add-Content -path $CnfFile -Value "[req]"
    Add-Content -path $CnfFile -Value "distinguished_name = req_distinguished_name"
    Add-Content -path $CnfFile -Value "req_extensions = v3_req"
    Add-Content -path $CnfFile -Value ""
    Add-Content -path $CnfFile -Value "[req_distinguished_name]"
    Add-Content -path $CnfFile -Value "countryName = Country Name (2 letter code)"
    Add-Content -path $CnfFile -Value "countryName_default = LU"
    Add-Content -path $CnfFile -Value "stateOrProvinceName = State or Province Name (full name)"
    Add-Content -path $CnfFile -Value "stateOrProvinceName_default = Luxembourg"
    Add-Content -path $CnfFile -Value "localityName = Locality Name (eg, city)"
    Add-Content -path $CnfFile -Value "localityName_default = Luxembourg"
    Add-Content -path $CnfFile -Value "organizationName = Organization Name (eg, company)"
    Add-Content -path $CnfFile -Value "organizationName_default = Banque de Luxembourg"
    Add-Content -path $CnfFile -Value "organizationalUnitName  = Organizational Unit Name (eg, section)"
    Add-Content -path $CnfFile -Value "organizationalUnitName_default  = CAT"
    Add-Content -path $CnfFile -Value "commonName = Common Name (eg, YOUR name)"
    Add-Content -path $CnfFile -Value "commonName_max  = 64"
    Add-Content -path $CnfFile -Value "emailAddress = Email Address"
    Add-Content -path $CnfFile -Value "emailAddress_max = 128"
    Add-Content -path $CnfFile -Value "emailAddress_default = infra.servers@bdl.lu"
    Add-Content -path $CnfFile -Value ""
    Add-Content -path $CnfFile -Value "[ v3_req ]"
    Add-Content -path $CnfFile -Value "# Extensions to add to a certificate request"
    Add-Content -path $CnfFile -Value "basicConstraints = CA:FALSE"
    Add-Content -path $CnfFile -Value "keyUsage = nonRepudiation, digitalSignature, keyEncipherment"
    Add-Content -path $CnfFile -Value "subjectAltName = @alt_names"
    Add-Content -path $CnfFile -Value ""
    Add-Content -path $CnfFile -Value "[alt_names]"
    $line = "DNS.1 = "+$CertificatCN
    Add-Content -path $CnfFile -Value $line

 }

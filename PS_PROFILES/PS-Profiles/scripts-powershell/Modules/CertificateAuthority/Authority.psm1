$script:cnf_ca = "ca.cnf"
$script:cnf_crl_info = @"
[crl_info]
URI.0                   = `$crl_url
"@
$script:cnf_default_ca = @"
[ca_default]
home                    = .
database                = `$home/db/index
serial                  = `$home/db/serial
crlnumber               = `$home/db/crlnumber
certificate             = `$home/certs/ca.pem
private_key             = `$home/private/ca.key
RANDFILE                = `$home/private/random
new_certs_dir           = `$home/certs
unique_subject          = no
default_md              = sha256
"@
$script:cnf_issuer_info = @"
[issuer_info]
caIssuers;URI.0         = `$aia_url
"@
$script:cnf_ocsp_info = @"
[issuer_info]
caIssuers;URI.0         = `$aia_url
OCSP;URI.0              = `$ocsp_url
"@
$script:cnf_name_constraints = @"
[name_constraints]
permitted;DNS.0=`$domain_suffix
excluded;IP.0=0.0.0.0/0.0.0.0
excluded;IP.1=0:0:0:0:0:0:0:0/0:0:0:0:0:0:0:0
"@

#--------------------------------------------------------------------------------------------------

function assertAuthority($path=$PWD.Path) {
  if (-not (Test-CertificateAuthority -Path $path)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "This is not a certificate authority that can be managed by this module." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }
}

function cnf_default($ocsp=$false) {
  return @"
aia_url                 = http://pki.`$domain_suffix/`$name.crt
crl_url                 = http://pki.`$domain_suffix/`$name.crl
$(if ($ocsp) { "ocsp_url                = http://ocsp-`$name.`$domain_suffix" })
default_ca              = ca_default
name_opt                = utf8,esc_ctrl,multiline,lname,align
"@
}
function cnf_policy($policy="policy_c_o_match") {
  return @"
policy                  = $policy

[policy_c_o_match]
countryName             = match
stateOrProvinceName     = optional
organizationName        = match
organizationalUnitName  = optional
commonName              = supplied
emailAddress            = optional

[policy_o_match]
countryName             = optional
stateOrProvinceName     = optional
organizationName        = match
organizationalUnitName  = optional
commonName              = supplied
emailAddress            = optional
"@
}

function ext_ca {
  return @"
[ca_ext]
basicConstraints        = critical,CA:true
keyUsage                = critical,keyCertSign,cRLSign
subjectKeyIdentifier    = hash
"@
}

function ext_client($public) {
  return @"
[client_ext]
authorityInfoAccess     = @issuer_info
authorityKeyIdentifier  = keyid:always, issuer:always
basicConstraints        = critical,CA:false
crlDistributionPoints   = @crl_info
extendedKeyUsage        = clientAuth,codeSigning,emailProtection
keyUsage                = critical,digitalSignature
$(if (-not ($public)) { "nameConstraints         = @name_constraints" })
subjectAltName          = email:move
subjectKeyIdentifier    = hash
"@
}

function ext_ocsp {
  return @"
[ocsp_ext]
authorityInfoAccess     = @issuer_info
authorityKeyIdentifier  = keyid:always
basicConstraints        = critical,CA:false
extendedKeyUsage        = critical,OCSPSigning
keyUsage                = critical,digitalSignature
subjectKeyIdentifier    = hash
"@
}

function ext_server($public) {
  return @"
[server_ext]
authorityInfoAccess     = @issuer_info
authorityKeyIdentifier  = keyid:always, issuer:always
basicConstraints        = critical,CA:false
crlDistributionPoints   = @crl_info
extendedKeyUsage        = clientAuth,serverAuth
keyUsage                = critical,digitalSignature,keyEncipherment
$(if (-not ($public)) { "nameConstraints         = @name_constraints" })
subjectKeyIdentifier    = hash
"@
}

function ext_subca($public) {
 return @"
[subca_ext]
authorityInfoAccess     = @issuer_info
authorityKeyIdentifier  = keyid:always
basicConstraints        = critical,CA:true,pathlen:0
crlDistributionPoints   = @crl_info
extendedKeyUsage        = clientAuth,serverAuth
keyUsage                = critical,keyCertSign,cRLSign
subjectKeyIdentifier    = hash
$(if (-not ($public)) { "nameConstraints         = @name_constraints" })
"@
}

function ext_timestamp {
  return @"
[timestamp_ext]
authorityInfoAccess     = @issuer_info
authorityKeyIdentifier  = keyid:always
basicConstraints        = CA:false
crlDistributionPoints   = @crl_info
extendedKeyUsage        = critical,timeStamping
keyUsage                = critical,digitalSignature
subjectKeyIdentifier    = hash
"@
}

#--------------------------------------------------------------------------------------------------

function Get-SubordinateAuthority {
  if (-not (Test-CertificateAuthority -Root)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "Subordinate authorities are typically created within a root authority managed by this module." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  $sub = @()
  $subca = Get-CertificateAuthoritySetting subca

  foreach ($name in $subca) {
    $sn = Get-CertificateAuthoritySetting "subca_$name"
    $cert = Get-IssuedCertificate | Where-Object {$_.SerialNumber -eq $sn}
    $sub += [PSCustomObject]@{
      Name              = $name
      Mounted           = $(Test-SubordinateAuthorityMounted -Name $name)
      Status            = $cert.Status
      SerialNumber      = $sn
      DistinguishedName = $cert.DistinguishedName
      NotValidAfter     = $(if ($cert.RevocationDate) { $cert.RevocationDate } else { $cert.ExpirationDate })
    }
  }

  return $sub
}

function New-CertificateAuthority {
  [CmdletBinding()]
  param (
    [Parameter(Position = 0)]
    [ValidateScript({ Test-Path })]
    [string] $Path = ($PWD.Path),
    [string] $Name = "root",
    [string] $Domain = "contoso.local",
    [Parameter(Mandatory = $true)]
    [securestring] $KeyPassword,
    [string] $Country = "US",
    [string] $Organization = "Contoso",
    [string] $CommonName = "RootCA",
    [switch] $Public,
    [switch] $UseOcsp
  )

  if (Test-CertificateAuthority $Path) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "'$Path' is already an Certificate Authority." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  $origialErrorActionPreference = $ErrorActionPreference
  $ErrorActionPreference = "Stop"

  Write-Output "`nCreating certificate authority directories..."

  if (-not (Test-Path $Path)) {
    New-Folder $Path
  }

  Push-Location $Path

  @("certs", "csr", "db", "private") | ForEach-Object {
    New-Folder $_
  }

  Write-Output "Initalizing certificate authority..."

  New-Item -Path "db/index" -ItemType File | Out-Null
  New-Item -Path "db/serial" -ItemType File `
    -Value $(Get-OpenSslRandom 15 -Hex)  | Out-Null
  New-Item -Path "db/crlnumber" -Value "1001" | Out-Null

  Set-Content -Path "$($script:cnf_ca)" -Value @"
[default]
name                    = $Name
domain_suffix           = $Domain
$(cnf_default $UseOcsp)

[ca_dn]
countryName             = "$Country"
organizationName        = "$Organization"
commonName              = "$CommonName"

$($script:cnf_default_ca)
copy_extensions         = none
default_days            = 7300
default_crl_days        = 365
$(if (-not ($Public)) { cnf_policy })

$($script:cnf_crl_info)

$(if ($UseOcsp) { $script:cnf_ocsp_info } else { $script:cnf_issuer_info })

$(if (-not ($Public)) { $script:cnf_name_constraints })

[req]
encrypt_key             = yes
default_md              = sha256
utf8                    = yes
string_mask             = utf8only
prompt                  = no
distinguished_name      = ca_dn
req_extensions          = ca_ext

$(ext_ca)

$(if ($UseOcsp) { ext_ocsp })

$(ext_subca($Public))
"@

  $cred = New-Object System.Management.Automation.PSCredential -ArgumentList "ni", $KeyPassword
  $passin = "-passin pass:$(($cred.GetNetworkCredential().Password).Trim())"

  Write-Output "`nGenerating the root certificate private key..."

  New-OpenSslEdwardsCurveKeypair -Path "./private/ca.key" -Password $KeyPassword -NoPublicFile

  Write-Output "Generating the root certificate request..."

  Invoke-OpenSsl "req -new -config $($script:cnf_ca) -out csr/ca.csr -key private/ca.key $passin"

  Write-Output "`n`nGenerating the root certificate for this authority..."

  Invoke-OpenSsl "ca -selfsign -config $($script:cnf_ca) -in csr/ca.csr -out certs/ca.pem -extensions ca_ext -notext $passin"

  Set-Content -Path ".openssl_ca" -Encoding UTF8 -Value @"
.type=root
.public=$Public
.name=$Name
.domain=$Domain
.c=$Country
.org=$Organization
.cn=$CommonName
.ocsp=$UseOcsp
.timestamp=False
"@

  if ($UseOcsp) {
    Update-OcspCertificate -AuthorityPassword $KeyPassword -Reset
  }

  Set-Content -Path ".\.gitignore" -Value @"
/**/private/*
/**/secrets/*
/**/*.old
.publish/
"@

  Set-Content -Path ".\.gitattributes" -Value @"
*       text eol=lf
*.cer   binary
*.csr   text
*.crl   text
*.crt   binary
*.der   binary
*.pem   text
*.p12   binary
*.pfx   binary
*.key   text
"@

  Set-Content -Path ".\.editorconfig" -Value @"
root = true

[*]
end_of_line = lf
indent_style = space
indent_size = 2
trim_trailing_whitespace = true
insert_final_newline = true

[{*.pem,*.crl,*.csr,*.key}]
insert_final_newline = false
"@

  Write-Output "`n`nCreation of a root certificate authority complete...`n"

  Pop-Location

  $ErrorActionPreference = $origialErrorActionPreference

  Write-Output "`n~~~~~~`n"
  Write-Output "A root certificate authority should only have subordinate authorities"
  Write-Output "so create at least one subordinate certificate authority to sign"
  Write-Output "certificates within this authority...`n"
}

function New-SubordinateAuthority {
  [CmdletBinding()]
  param (
    [Parameter(Position = 0)]
    [string] $Name = "subca1",
    [string] $Domain,
    [Parameter(Mandatory = $true)]
    [securestring] $KeyPassword,
    [ValidateSet("Edwards", "Eliptic", "RSA")]
    [string] $KeyEncryption = "RSA",
    [string] $Country,
    [string] $Organization,
    [string] $CommonName = "SubCA",
    [Parameter(Mandatory = $true)]
    [securestring] $AuthorityPassword,
    [switch] $Public,
    [switch] $Force,
    [switch] $UseOcsp,
    [switch] $UseTimestamp
  )

  if (-not (Test-CertificateAuthority -Root)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "Subordinate authorities must be created within a root authority managed by this module." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  $origialErrorActionPreference = $ErrorActionPreference
  $ErrorActionPreference = "Stop"

  # if ((Get-CertificateAuthoritySetting SubCA) -contains $Name) {
  #   if (-not $Force) {
  #   <# Settings say subca is issued and not revoke. must removed before it can be overwritten anf force wasn't provided. #>
  #   } else {
  #     Remove-SubordinateAuthority $Name
  #   }
  # }

  if (Test-Path $Name) {
    if ($Force) {
      Remove-Item -Path $Name -Recurse -Force
    } else {
      $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
        -Message "'$Name' folder exists and the Force parameter was not specified. Aborting creation." `
        -ExceptionType "System.InvalidOperationException" `
        -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
    }
  }

  Write-Output "`nCreating subordinate authority directories..."

  if ($Domain.Length -eq 0) {
    $Domain = Get-CertificateAuthoritySetting Domain
  }

  if ($Country.Length -eq 0) {
    $Country = Get-CertificateAuthoritySetting c
  }

  if ($Organization.Length -eq 0) {
    $Organization = Get-CertificateAuthoritySetting org
  }

  if (-not (Test-Path $Name)) {
    New-Folder $Name
  }

  Push-Location $Name

  @("certs", "csr", "db", "private") | ForEach-Object {
    New-Folder $_
  }

  Write-Output "Initalizing subordinate authority..."

  New-Item -Path "db/index" -ItemType File | Out-Null
  New-Item -Path "db/serial" -ItemType File `
    -Value $(Get-OpenSslRandom 15 -Hex)  | Out-Null
  New-Item -Path "db/crlnumber" -Value "1001" | Out-Null

  Set-Content -Path "$($script:cnf_ca)" -Value @"
[default]
name                    = $Name
domain_suffix           = $Domain
$(cnf_default $UseOcsp)

[ca_dn]
countryName             = "$Country"
organizationName        = "$Organization"
commonName              = "$CommonName"

$($script:cnf_default_ca)
copy_extensions         = copy
default_days            = 365
default_crl_days        = 30
$(if (-not ($Public)) { cnf_policy })

$($script:cnf_crl_info)

$(if ($UseOcsp) { $script:cnf_ocsp_info } else { $script:cnf_issuer_info })

$(if (-not ($Public)) { $script:cnf_name_constraints })

[req]
encrypt_key             = yes
default_md              = sha256
utf8                    = yes
string_mask             = utf8only
prompt                  = no
distinguished_name      = ca_dn
req_extensions          = ca_ext

$(ext_ca)

$(if ($UseOcsp) { ext_ocsp })

$(if ($UseTimestamp) { ext_timestamp })

$(ext_server $Public)

$(ext_client $Public)
"@

  $cred = New-Object System.Management.Automation.PSCredential -ArgumentList "ni", $KeyPassword
  $passin = "-passin pass:$(($cred.GetNetworkCredential().Password).Trim())"

  Write-Output "`nGenerating the subordinate certificate private key..."

  switch ($KeyEncryption) {
    "Edwards" {
      New-OpenSslEdwardsCurveKeypair -Path "./private/ca.key" -Password $KeyPassword -NoPublicFile
    }
    "Eliptic" {
      New-OpenSslElipticCurveKeypair -Path "./private/ca.key" -Password $KeyPassword -NoPublicFile
    }
    "RSA" {
      New-OpenSslRsaKeypair -Path "./private/ca.key" -Password $KeyPassword -NoPublicFile
    }
  }

  Write-Output "`nGenerating the subordinate certificate request..."

  Invoke-OpenSsl "req -new -config $($script:cnf_ca) -out csr/ca.csr -key private/ca.key $passin"

  Set-Content -Path ".openssl_ca" -Encoding UTF8 -Value @"
.type=subordinate
.public=$Public
.name=$Name
.domain=$Domain
.c=$Country
.org=$Organization
.cn=$CommonName
.ocsp=$UseOcsp
.timestamp=$UseTimestamp
"@

  Write-Output "Using Root CA to sign the certificate for this authority..."

  Pop-Location

  Approve-SubordinateAuthority -Name $Name -KeyPassword $AuthorityPassword

  Set-Content -Path "./$Name/certs/ca-chain.pem" -Value @"
$(foreach ($line in (Get-Content -Path "./certs/ca.pem")) { "$line`n" })
$(foreach ($line in (Get-Content -Path "./$Name/certs/ca.pem")) { "$line`n" })
"@

  Push-Location $Name

  if ($UseOcsp) {
    Update-OcspCertificate -AuthorityPassword $KeyPassword -Reset
  }

  if ($UseTimestamp) {
    Update-TimestampCertificate -AuthorityPassword $KeyPassword -Reset
  }

  Write-Output "`n`nCreation of a subordinate authority complete...`n"

  $sn = Get-CertificateSerialNumber -Path "./certs/ca.pem"

  Pop-Location

  $subca = (Get-CertificateAuthoritySetting subca | Where-Object { $_ -like $Name })

  if ($subca.Count -gt 0) {
    Set-CertificateAuthoritySetting -Name "subca" -Value $Name -Remove
    Set-CertificateAuthoritySetting -Name "subca_$Name" -Remove
  }

  Set-CertificateAuthoritySetting -Name "subca" -Value "$Name"
  Set-CertificateAuthoritySetting -Name "subca_$Name" -Value $sn

  $ErrorActionPreference = $origialErrorActionPreference

  Write-Output "`n~~~~~~`n"
  Write-Output "The '$Name' subordinate certificate authority can only be used to sign"
  Write-Output "certificates within this authority...`n"
}

function Publish-CertificateAuthority {
  [CmdletBinding()]
  [Alias("ca-publish")]
  param (
    [Parameter(Position = 0, ValueFromPipeline = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $Path = "$($PWD.Path)/.publish",
    [switch] $Force
  )

  assertAuthority

  if (-not (Test-CertificateAuthority -Root)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "Certificate Authorities can only be published by the root authority." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  $origialErrorActionPreference = $ErrorActionPreference
  $ErrorActionPreference = "Stop"

  Write-Verbose "Path: $Path"
  if (-not (Test-Path $Path)) {
    New-Folder $Path
  }

  $Destination = "$((Resolve-Path -Relative -Path $Path) -replace '\\', '/')"
  Write-Verbose "Destination: $Destination"

  if ((Get-ChildItem $Destination).Count -gt 0) {
    if ($Force) {
      Remove-Item -Path "$Destination/*" -Recurse -Force
    } else {
      $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
        -Message "Files present in publish path and -Force was not supplied." `
        -ExceptionType "System.InvalidOperationException" `
        -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
    }
  }

  Set-Content -Path "$Destination/mime.types" -Value @"
application/pkcs7-mime              .p7c
application/pkcs8                   .p8  .key
application/pkcs10                  .p10 .csr
application/pkix-cert               .cer
application/pkix-crl                .crl
application/x-pem-file              .pem
application/x-pkcs7-certificates    .p7b .spc
application/x-pkcs7-certreqresp     .p7r
application/x-pkcs7-crl             .crl
application/x-pkcs12                .p12 .pfx
application/x-x509-ca-cert          .crt .der
application/x-x509-user-cert        .crt
"@

  $root = Get-CertificateAuthoritySetting name
  $authorities = @($root, $(Get-CertificateAuthoritySetting subca))

  Write-Verbose "root: $root"
  Write-Verbose "authorities: $authorities"

  foreach ($authority in $authorities) {
    if (($authority -eq $root) -or (Test-SubordinateAuthorityMounted -Name $authority)) {
      Write-Verbose "Authority is mounted. Proceeding to publish..."
      if ($authority -eq $root) {
        $subdir = "./"
      } else {
        $subdir = "$authority/"
      }

      Push-Location $subdir

      if ( "imported" -ne (Get-CertificateAuthoritySetting type)) {
        $cn = Get-CertificateAuthoritySetting cn

        Write-Output ">>>---------------- '$cn' Certificate Authority`n"

        $pass = Read-Host "Enter password for '$cn' private key" -AsSecureString

        Write-Output "Updating the certificate authority database..."
        Update-CertificateAuthorityDatabase -AuthorityPassword $pass

        if (Get-CertificateAuthoritySetting ocsp) {
          Update-OcspCertificate -AuthorityPassword $pass
        }

        if (Get-CertificateAuthoritySetting timestamp) {
          Update-TimestampCertificate -AuthorityPassword $pass
        }

        Write-Output "Updating the certificate revocation list..."
        Update-CertificateAuthorityRevocationList -AuthorityPassword $pass

        $name = Get-CertificateAuthoritySetting name
      } else {
        Write-Output ">>>---------------- '$cn' Certificate Authority"
        Write-Output "                     This is an imported authority.`n"
      }

      Pop-Location

      if (Test-Path "$($subdir)certs/ca.pem" ) {
        ConvertFrom-PemCertificate -CertPath "$($subdir)certs/ca.pem" `
        -Destination "$Destination/$name.crt" -To DER
      }

      if (Test-Path "$subdir$name.crl") {
        Invoke-OpenSsl "crl -in $subdir$name.crl -out $Destination/$name.crl -outform der"
      }

      if (Test-Path "$($subdir)certs/ocsp.pem" ) {
        Copy-Item -Path "$($subdir)certs/ocsp.pem" -Destination "$Destination/$name-ocsp.pem"
      }

      if (Test-Path "$($subdir)certs/timestamp.pem" ) {
        Copy-Item -Path "$($subdir)certs/timestamp.pem" -Destination "$Destination/$name-timestamp.pem"
      }

      $subdir = ""
    }
  }

  $ErrorActionPreference = $origialErrorActionPreference

  Write-Output "`n~~~~~~`n"
  Write-Output "This certificate authority has been published to '$Destination'"
}

function Remove-SubordinateAuthority {
  [CmdletBinding()]
  [Alias("remove-subca")]
  param (
    [ValidateScript({ Test-Path $(Resolve-Path $_) })]
    [string] $Path = ($PWD.Path),
    [string] $Name = $(Split-Path -Path $Path -Leaf)
  )

  assertAuthority $Path

  if (Test-CertificateAuthority $Path -Subordinate) {
    Push-Location -Path "$((Get-Item -Path $Path).Parent.FullName)"
  } else {
    Push-Location $Path
  }

  if (-not (Test-CertificateAuthority -Root)) {
    Pop-Location
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "'$Path' is not part of a certificate authority that includes the root authority." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  $subca = (Get-CertificateAuthoritySetting subca | Where-Object { $_ -like $Name })

  if ($subca.Length -gt 0) {
    $sn = Get-CertificateAuthoritySetting "subca_$Name"

    if ($sn -eq "~REVOKED~") {
      Set-CertificateAuthoritySetting -Name "subca" -Value $Name -Remove
      Set-CertificateAuthoritySetting -Name "subca_$Name" -Remove

      if (Test-Path "$Name/") {
        Remove-Item -Path $Name -Recurse -Force
      }
    } else {
      Pop-Location
      $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
        -Message "'$Name' authority must be revoked before it can be removed." `
        -ExceptionType "System.InvalidOperationException" `
        -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
    }
  } else {
    Pop-Location
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
      -Message "'$Name' authority is not currently managed by this root authority." `
      -ExceptionType "System.InvalidOperationException" `
      -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  Pop-Location
}

function Revoke-SubordinateAuthority {
  [CmdletBinding()]
  [Alias("revoke-subca")]
  param (
    [ValidateScript({ Test-Path $(Resolve-Path $_) })]
    [string] $Path = ($PWD.Path),
    [string] $Name = $(Split-Path -Path $Path -Leaf),
    [ValidateSet("unspecified", "keyCompromise", "CACompromise", "affiliationChanged", "superseded", "cessationOfOperation", "certificateHold", "removeFromCRL")]
    [string] $Reason = "unspecified"
  )

  assertAuthority $Path

  if (Test-CertificateAuthority $Path -Subordinate) {
    Push-Location -Path "$((Get-Item -Path $Path).Parent.FullName)"
  } else {
    Push-Location $Path
  }

  if (-not (Test-CertificateAuthority -Root)) {
    Pop-Location
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "'$Path' is not part of a certificate authority that includes the root authority." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  $subca = (Get-CertificateAuthoritySetting subca | Where-Object { $_ -like $Name })

  if ($subca.Length -gt 0) {
    $sn = Get-CertificateAuthoritySetting "subca_$Name"

    if ($sn -eq "~REVOKED~") {
      Pop-Location
      $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
        -Message "'$Name' authority has already been revoked." `
        -ExceptionType "System.InvalidOperationException" `
        -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
    }

    if ($sn.Length -gt 0) {
      Invoke-OpenSsl "ca -config $($script:cnf_ca) -revoke certs/$sn.pem -crl_reason $Reason"

      Move-Item -Path "certs/$sn.pem" "certs/$sn.pem.revoked"

      Set-CertificateAuthoritySetting -Name "subca_$Name" -Value "~REVOKED~"

      if (Test-Path "$Name/") {
        ###TODO: If subordinate authority is mounted (directly below root), cycle through each issued certificate and revoke them as well
      }
    }
  } else {
    Pop-Location
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
      -Message "'$Name' authority is not currently managed by this root authority." `
      -ExceptionType "System.InvalidOperationException" `
      -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  Pop-Location
}

function Test-CertificateAuthority {
  [CmdletBinding()]
  [Alias("ca-test")]
  param (
    [Parameter(Position = 0)]
    [ValidateScript({ Test-Path $(Resolve-Path $_) })]
    [string] $Path = ($PWD.Path),
    [switch] $Root,
    [switch] $Subordinate
  )

  $result = Test-Path "$Path/.openssl_ca"

  if ($result) {
    if ($root) {
      return (Get-Content .openssl_ca | Select-String ".type=root").Matches.Count -gt 0
    }

    if ($Subordinate) {
      return (Get-Content .openssl_ca | Select-String ".type=subordinate").Matches.Count -gt 0
    }
  }

  return $result
}

function Test-SubordinateAuthorityMounted {
  [CmdletBinding()]
  param (
    [ValidateScript({ Test-Path })]
    [string] $Path = ($PWD.Path),
    [string] $Name = $(Split-Path -Path $Path -Leaf)
  )

  assertAuthority $Path

  if (Test-CertificateAuthority $Path -Subordinate) {
    Push-Location -Path "$((Get-Item -Path $Path).Parent.FullName)"
  } else {
    Push-Location $Path
  }


  if ($Path -like "*$Name") {
    $testfile = "$Path/.openssl_ca"
  } else {
    $testfile = "$Path/$Name/.openssl_ca"
  }

  if (Test-Path $testfile ) {
    return $true
  } else {
    return $false
  }
}

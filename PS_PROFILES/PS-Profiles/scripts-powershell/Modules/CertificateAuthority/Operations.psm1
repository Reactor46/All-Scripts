$script:cnf_ca = "ca.cnf"
$script:cnf_req = @"
[req]
default_md = sha256
utf8 = yes
string_mask = utf8only
prompt = no
distinguished_name = req_subj
"@

function cnf_san($cn, $names) {
  $URI = @()
  $DNS = @()
  $IP = @()
  $EMail = @()

  if ($names -notcontains $cn) {
    $DNS += $cn
  }

  foreach ($name in $names) {
    if ($name -like "*://*") {
      $URI += $name
      continue
    }

    if ($name -like "*@*") {
      $EMail += $name
      continue
    }

    if (Test-IPAddress $name) {
      $IP += $name
      continue
    }

    $DNS += $name
  }

  $text = @"
[server_req]
subjectAltName = @san_list

[san_list]`n
"@
  $items = @()

  for ($i = 0; $i -lt $URI.Count; $i++)   { $items += "URI.$i = $($URI[$i])" }
  for ($i = 0; $i -lt $DNS.Count; $i++)   { $items += "DNS.$i = $($DNS[$i])" }
  for ($i = 0; $i -lt $IP.Count; $i++)    { $items += "IP.$i = $($IP[$i])" }
  for ($i = 0; $i -lt $EMail.Count; $i++) { $items += "email.$i = $($EMail[$i])" }

  foreach ($item in $Items) { $text += "$item`n" }

  return $text
}

function getConfigFileName($Id) {
  return ".san.$Id.cnf"
}

function req_subj($Country,$State,$Locality,$Organization,$OrganizationUnit,$Name) {
  return (@"
[req_subj]
countryName = $Country
$(if ($State.Length -gt 0) { "stateOrProvinceName = $State" })
$(if ($Locality.Length -gt 0) { "localityName = $Locality" })
organizationName = $Organization
$(if ($OrganizationUnit.Length -gt 0) { "organizationUnitName = $OrganizationUnit" })
commonName = $Name
"@) -replace '(?m)^\s*?\n'
}

function generateRequestConfig($path, $country, $org, $cn) {
  Set-Content -Path $path -Value @"
[req]
default_bits            = 2048
encrypt_key             = no
default_md              = sha256
utf8                    = yes
string_mask             = utf8only
prompt                  = no
distinguished_name      = req_subj

[req_subj]
countryName             = $country
organizationName        = $org
commonName              = $cn
"@
}

function signCertificate($path, $name, $pword, $extension, $days,$subca=$false) {
  if ($PsCmdlet.ParameterSetName -eq "path") {
    $name = Import-CertificateRequest $path
  }

  $cred = New-Object System.Management.Automation.PSCredential -ArgumentList "ni", $pword
  $passin = "-passin pass:$(($cred.GetNetworkCredential().Password).Trim())"

  $cmd = "ca"

  if (-not $subca) {
    $cmd = $cmd + " -batch"
  }

  $cmd = $cmd + " -config $($script:cnf_ca) -notext -extensions $extension -days $days"

  if ($subca) {
    $cmd = $cmd + " $passin -out $name/certs/ca.pem -in $name/csr/ca.csr"
  } else {
    $cmd = $cmd + " $passin -infiles csr/$name.csr"
  }

  Invoke-OpenSsl $cmd

  return $Name
}

#--------------------------------------------------------------------------------------------------

function Approve-ServerCertificate {
  [CmdletBinding(DefaultParameterSetName = "path")]
  [Alias("sign-server-certificate")]
  param (
    [Parameter(ParameterSetName = "path", Mandatory = $true)]
    [ValidateScript({ Test-Path $(Resolve-Path $_) })]
    [string] $Path,
    [Parameter(ParameterSetName = "name", Mandatory = $true)]
    [string] $Name,
    [Parameter(ParameterSetName = "path", Mandatory = $true)]
    [Parameter(ParameterSetName = "name", Mandatory = $true)]
    [securestring] $KeyPassword,
    [int] $Days = 365 # 1 year
  )

  if (-not (Test-CertificateAuthority -Subordinate)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "Certificates can only be signed by a subordinate authority that this module can manage." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  return signCertificate $Path $Name $KeyPassword "server_ext" $Days
}

function Approve-SubordinateAuthority {
  [CmdletBinding(DefaultParameterSetName = "path")]
  [Alias("sign-subordinate-authority", "sign-subca")]
  param (
    [Parameter(ParameterSetName = "path", Mandatory = $true)]
    [ValidateScript({ Test-Path $(Resolve-Path $_) })]
    [string] $Path,
    [Parameter(ParameterSetName = "name", Mandatory = $true)]
    [string] $Name,
    [Parameter(ParameterSetName = "path", Mandatory = $true)]
    [Parameter(ParameterSetName = "name", Mandatory = $true)]
    [securestring] $KeyPassword,
    [int] $Days = 1825 # 5 years
  )

  if (-not (Test-CertificateAuthority -Root)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "Subordinate authorities can only be approved by a root authority." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  return signCertificate $Path $Name $KeyPassword "subca_ext" $Days $true
}

function Approve-UserCertificate {
  [CmdletBinding(DefaultParameterSetName = "path")]
  [Alias("sign-user-certificate")]
  param (
    [Parameter(ParameterSetName = "path", Mandatory = $true)]
    [ValidateScript({ Test-Path $(Resolve-Path $_) })]
    [string] $Path,
    [Parameter(ParameterSetName = "name", Mandatory = $true)]
    [string] $Name,
    [Parameter(ParameterSetName = "path", Mandatory = $true)]
    [Parameter(ParameterSetName = "name", Mandatory = $true)]
    [securestring] $KeyPassword,
    [int] $Days = 365 # 1 year
  )

  if (-not (Test-CertificateAuthority -Subordinate)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "Certificates can only be signed by a subordinate authority that this module can manage." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  return signCertificate $Path $Name $KeyPassword "client_ext" $Days
}

function Get-ImportedCertificateRequest {
  [CmdletBinding()]
  [Alias("list-imported-requests")]
  param (
    [Parameter(ValueFromPipeline = $true, Position = 0)]
    [string] $Name
  )

  if (-not (Test-CertificateAuthority)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "This is not a certificate authority that can be managed by this module." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  if ($Name.Length -gt 0) {
    if (Test-Path "csr/$Name.csr") {
      return Get-CertificateRequest -Path "csr/$Name.csr"
    } else {
      $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
        -Message "A certificate request with named '$Name' does not exists in this authority." `
        -ExceptionType "System.Management.Automation.ItemNotFoundException" `
        -ErrorId "ItemNotFoundException" -ErrorCategory ObjectNotFound))
    }
  }

  $requests = @()
  $regex_version = "Version:\s(\d+)\s"
  $regex_subject = "Subject:\s(.+)"
  $regex_pub = "Public Key Algorithm:\s(\w+)"
  $regex_sig = "Signature Algorithm:\s(\w+)"
  $regex_valid = "self-signature verify OK"
  $files = (Get-ChildItem "csr/").Name

  foreach ($file in $files) {
    $content = Get-CertificateRequest -Path "csr/$file"

    $version = Select-String -InputObject $content -Pattern $regex_version
    $subject = Select-String -InputObject $content -Pattern $regex_subject
    $pub = Select-String -InputObject $content -Pattern $regex_pub
    $sig = Select-String -InputObject $content -Pattern $regex_sig
    $valid = Select-String -InputObject $content -Pattern $regex_valid

    if ($subject.Matches) {
      $subject = (($subject.Matches[0].Groups[1].Value) -split "\s\s+")[0]
      $subject = $subject -replace ' = ', '=' -replace ', ', ','
    } else {
      $subject = "**Unknown**"
    }

    $request = [PSCustomObject]@{
      Name = (Get-Item "csr/$file").BaseName
      Subject = $subject
      Version = $(if ($version.Matches) { $version.Matches[0].Groups[1].Value } else { 0 })
      PublicKey = $(if ($pub.Matches) { $pub.Matches[0].Groups[1].Value } else { "" })
      Signature = $(if ($sig.Matches) { $sig.Matches[0].Groups[1].Value } else { "" })
      Valid = $(if ($valid.Matches) { $true } else { $false })
    }

    $requests += $request
    $request = $null
  }

  return $requests
}

function Get-IssuedCertificate {
  [CmdletBinding(DefaultParameterSetName = "name")]
  [Alias("Get-RevokedIssuedCertificate", "list-issued-certificates", "list-revoked-certificates")]
  param (
    [Parameter(ParameterSetName = "name", ValueFromPipeline = $true, Position = 0)]
    [string] $Name,
    [Parameter(ParameterSetName = "revoked")]
    [switch] $Revoked
  )

  if (-not (Test-CertificateAuthority)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "This is not a certificate authority that can be managed by this module." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  if ($Name.Length -gt 0) {
    if ($MyInvocation.InvocationName -like "*revoked*") {
      if (Test-Path "certs/$Name.pem.revoked") {
        return Get-Certificate -Path "certs/$Name.pem.revoked"
      } else {
        $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
          -Message "A revoked certificate with the name '$Name' does not exists in this authority." `
          -ExceptionType "System.Management.Automation.ItemNotFoundException" `
          -ErrorId "ItemNotFoundException" -ErrorCategory ObjectNotFound))
      }
    } else {
      if (Test-Path "certs/$Name.pem") {
        return Get-Certificate -Path "certs/$Name.pem"
      } else {
        $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
          -Message "A certificate with the name '$Name' does not exists in this authority." `
          -ExceptionType "System.Management.Automation.ItemNotFoundException" `
          -ErrorId "ItemNotFoundException" -ErrorCategory ObjectNotFound))
      }
    }
  }

  $certs = @()

  $content = Get-Content "db/index"
  $regex = "(\w)\s+(\w+)(\s+\w+,\w+\s+|\s+)(\w+)\s+(\w+)\s(.+)"

  foreach ($line in $content) {
    $result = Select-String -InputObject $line -Pattern $regex

    $cert = [PSCustomObject]@{
      SerialNumber      = ($result.Matches[0].Groups[4].Value | Out-String).Trim()
      DistinguishedName = ($result.Matches[0].Groups[6].Value | Out-String).Trim()
      Filename          = ($result.Matches[0].Groups[5].Value | Out-String).Trim()
      Status            = switch (($result.Matches[0].Groups[1].Value | Out-String).Trim()) {
        "V" { "Valid" }
        "E" { "Expired" }
        "R" { "Revoked" }
        Default { "Unknown" }
      }
      ExpirationDate    = `
        [datetime]::ParseExact(($result.Matches[0].Groups[2].Value `
          | Out-String).Trim(), "yyMMddHHmmssZ", $null)
      RevocationDate    = ($result.Matches[0].Groups[3].Value | Out-String).Trim()
      RevocationReason  = $null
    }

    if ($cert.RevocationDate.Length -gt 0) {
      $cert.RevocationReason = ($cert.RevocationDate -split ',')[1]
      $cert.RevocationDate = `
        [datetime]::ParseExact(($cert.RevocationDate -split ',')[0], "yyMMddHHmmssZ", $null)
    }

    $certs += $cert
    $result = $null
  }

  if (($PsCmdlet.ParameterSetName -eq "revoked") -or ($MyInvocation.InvocationName -like "*revoked*")) {
    $certs =  $certs | Where-Object { $_.Status -eq "Revoked" }
  }

  return $certs
}

function Get-IssuedCertificateValidity {
  [CmdletBinding()]
  param (
    [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
    [Alias("Certificate", "Cert")]
    [string] $Name,
    [Parameter(Position = 1)]
    [ValidateScript({ Test-Path $(Resolve-Path $_) })]
    [Alias("CA", "CAFile", "CABundle")]
    [string] $CAPath
  )

  if (-not (Test-CertificateAuthority)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "This is not a certificate authority that can be managed by this module." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  if (Test-Path "certs/$Name.pem") {
    $param = "verify"
    $ca = "./.openssl.$((New-Guid).Guid).cacert.pem"

    if ($CAPath) {
      Copy-Item -Path $CAPath -Destination $ca
    } else {
      $authority = Get-CertificateAuthoritySetting Name

      if (Test-Path "./$authority.crt") {
        $ca = "$authority.crt"
      } else {
        Invoke-WebRequest "https://curl.se/ca/cacert.pem" -OutFile $ca
      }
    }

    if (Test-Path $ca) {
      $param += " -CAfile $ca"
    }

    $param += " certs/$Name.pem"

    Write-Verbose "param: $param"

    Invoke-OpenSsl $param
  } else {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
      -Message "A certificate with the name '$Name' does not exists in this authority." `
      -ExceptionType "System.Management.Automation.ItemNotFoundException" `
      -ErrorId "ItemNotFoundException" -ErrorCategory ObjectNotFound ))
  }
}

function Import-CertificateRequest {
  [CmdletBinding()]
  [Alias("import-csr")]
  param (
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateScript({ Test-Path $(Resolve-Path $_) })]
    [string] $Path
  )

  if (-not (Test-CertificateAuthority $Path)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "This is not a certificate authority that can be managed by this module." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  if (-not (Test-Path csr)) {
    New-Item -Path csr -ItemType Directory | Out-Null
  }

  $id = Get-OpenSslRandom 16 -Hex

  Copy-Item -Path $Path -Destination "csr/$id.csr"

  return $id
}

function New-ServerCertificate {
  [CmdletBinding()]
  [Alias("new-server-certificate")]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [string] $Name,
    [ValidateSet("Edwards", "Eliptic", "RSA")]
    [string] $KeyEncryption = "RSA",
    [securestring] $KeyPassword,
    [string] $Country,
    [Alias("Province", "Region")]
    [string] $State,
    [Alias("City")]
    [string] $Locality,
    [Alias("Company")]
    [string] $Organization,
    [Alias("Section")]
    [string] $OrganizationUnit,
    [string[]] $AdditionalNames,
    [Parameter(Mandatory = $true)]
    [securestring] $AuthorityPassword
  )

  if (-not (Test-CertificateAuthority -Subordinate)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "Certificates can only be requested in a subordinate authority that this module can manage." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  $id = New-ServerCertificateRequest ($PSBoundParameters | Where-Object { $_.Key -ne "AuthorityPassword" })

  Approve-ServerCertificate -Name $id -KeyPassword $AuthorityPassword
}

function New-ServerCertificateRequest {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [string] $Name,
    [ValidateSet("Edwards", "Eliptic", "RSA")]
    [string] $KeyEncryption = "RSA",
    [securestring] $KeyPassword,
    [string] $Country,
    [Alias("Province", "Region")]
    [string] $State,
    [Alias("City")]
    [string] $Locality,
    [Alias("Company")]
    [string] $Organization,
    [Alias("Section")]
    [string] $OrganizationUnit,
    [string[]] $AdditionalNames,
    [switch] $KeepCnf
  )

  if (-not (Test-CertificateAuthority -Subordinate)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "Certificates can only be requested in a subordinate authority that this module can manage." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  if ($Domain.Length -eq 0) {
    $Domain = Get-CertificateAuthoritySetting Domain
  }

  if ($Country.Length -eq 0) {
    $Country = Get-CertificateAuthoritySetting c
  }

  if ($Organization.Length -eq 0) {
    $Organization = Get-CertificateAuthoritySetting org
  }

  if (($Name.Length -eq 0) -and ($AdditionalNames.Count -gt 0)) {
    $Name = $AdditionalNames[0]
  }

  if ($KeyPassword) {
    $cred = New-Object System.Management.Automation.PSCredential -ArgumentList "ni", $KeyPassword
    $passin = "-passin pass:$(($cred.GetNetworkCredential().Password).Trim())"
  } else {
    $passin = ""
  }

  Write-Verbose "Generating the server certificate private key..."

  $id = Get-OpenSslRandom 8 -Hex
  $KeyName = "private/$id.key"
  $CsrName = "csr/$id.csr"

  switch ($KeyEncryption) {
    "Edwards" {
      New-OpenSslEdwardsCurveKeypair -Path $KeyName -Password $KeyPassword -NoPublicFile | Write-Verbose
    }
    "Eliptic" {
      New-OpenSslElipticCurveKeypair -Path $KeyName -Password $KeyPassword -NoPublicFile | Write-Verbose
    }
    "RSA" {
      New-OpenSslRsaKeypair -Path $KeyName -Password $KeyPassword -NoPublicFile | Write-Verbose
    }
  }

  Write-Verbose "Generating the server certificate request..."

  $cnf = getConfigFileName $id
  Set-Content -Path $cnf -Value @"
$($script:cnf_req)
req_extensions = server_req

$(req_subj $Country $State $Locality $Organization $OrganizationUnit $Name)

$(cnf_san $Name $AdditionalNames)
"@

  $param = "req -config $cnf -new -key $KeyName -out $CsrName $passin"

  Write-Verbose "param: $param"

  Invoke-OpenSsl $param

  if ($KeepCnf) {
    Write-Verbose "Not deleting '$cnf' file."
  } else {
    Remove-Item -Path $cnf -Force
  }

  if (Test-Path -Path "csr/$id.csr") {
    return $id
  } else {
    return $null
  }
}

function New-UserCertificate {
  [CmdletBinding()]
  [Alias("new-user-certificate")]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [string] $Name,
    [string] $Domain,
    [ValidateSet("Edwards", "Eliptic", "RSA")]
    [string] $KeyEncryption = "RSA",
    [securestring] $KeyPassword,
    [string] $Country,
    [Alias("Province", "Region")]
    [string] $State,
    [Alias("City")]
    [string] $Locality,
    [Alias("Company")]
    [string] $Organization,
    [Alias("Section")]
    [string] $OrganizationUnit,
    [Parameter(Mandatory = $true)]
    [securestring] $AuthorityPassword
  )

  if (-not (Test-CertificateAuthority -Subordinate)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "Certificates can only be requested in a subordinate authority that this module can manage." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  $param = @{}
  $PSBoundParameters.GetEnumerator() | ForEach-Object {
    if ($_.Key -ne "AuthorityPassword") {
      $param.Add($_.Key, $_.Value)
    }
  }

  $id = New-UserCertificateRequest @param

  Approve-UserCertificate -Name $id -KeyPassword $AuthorityPassword
}

function New-UserCertificateRequest {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [string] $Name,
    [string] $Domain,
    [ValidateSet("Edwards", "Eliptic", "RSA")]
    [string] $KeyEncryption = "RSA",
    [securestring] $KeyPassword,
    [string] $Country,
    [Alias("Province", "Region")]
    [string] $State,
    [Alias("City")]
    [string] $Locality,
    [Alias("Company")]
    [string] $Organization,
    [Alias("Section")]
    [string] $OrganizationUnit,
    [switch] $KeepCnf
  )

  if (-not (Test-CertificateAuthority -Subordinate)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "Certificates can only be requested in a subordinate authority that this module can manage." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  if ($Name -like "*@*") {
    $Domain = ($Name -split '@')[1]
    $Name   = ($Name -split '@')[0]
  }

  if ($Domain.Length -eq 0) {
    $Domain = Get-CertificateAuthoritySetting Domain
  }

  if ($Country.Length -eq 0) {
    $Country = Get-CertificateAuthoritySetting c
  }

  if ($Organization.Length -eq 0) {
    $Organization = Get-CertificateAuthoritySetting org
  }

  if ($KeyPassword) {
    $cred = New-Object System.Management.Automation.PSCredential -ArgumentList "ni", $KeyPassword
    $passin = "-passin pass:$(($cred.GetNetworkCredential().Password).Trim())"
  } else {
    $passin = ""
  }

  Write-Verbose "Generating the client certificate private key..."

  $id = Get-OpenSslRandom 8 -Hex
  $KeyName = "./private/$id.key"
  $CsrName = "./csr/$id.csr"

  switch ($KeyEncryption) {
    "Edwards" {
      New-OpenSslEdwardsCurveKeypair -Path $KeyName -Password $KeyPassword -NoPublicFile | Write-Verbose
    }
    "Eliptic" {
      New-OpenSslElipticCurveKeypair -Path $KeyName -Password $KeyPassword -NoPublicFile | Write-Verbose
    }
    "RSA" {
      New-OpenSslRsaKeypair -Path $KeyName -Password $KeyPassword -NoPublicFile | Write-Verbose
    }
  }

  Write-Verbose "Generating the client certificate request..."

  $cnf = getConfigFileName $id
  Set-Content -Path $cnf -Value @"
$($script:cnf_req)
req_extensions = client_req

[ client_req ]
keyUsage = nonRepudiation, digitalSignature, keyEncipherment
subjectAltName = email:$Name@$Domain

$(req_subj $Country $State $Locality $Organization $OrganizationUnit "$Name@$Domain")
"@

  $param = "req -config $cnf -new -key $KeyName -out $CsrName $passin"

  Write-Verbose "param: $param"

  Invoke-OpenSsl $param

  if ($KeepCnf) {
    Write-Verbose "Not deleting '$cnf' file."
  } else {
    Remove-Item -Path $cnf -Force
  }

  if (Test-Path -Path "csr/$id.csr") {
    return $id
  } else {
    return $null
  }
}

function Revoke-Certificate {
  [CmdletBinding()]
  [Alias("revoke-issued-certificate")]
  param (
    [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
    [Alias("Certificate", "Cert")]
    [string] $Name,
    [securestring] $AuthorityPassword,
    [ValidateSet("unspecified", "keyCompromise", "CACompromise", "affiliationChanged", "superseded", "cessationOfOperation", "certificateHold", "removeFromCRL")]
    [string] $Reason = "unspecified"
  )

  if (-not (Test-CertificateAuthority)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "This is not a certificate authority that can be managed by this module." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  if (Test-Path "certs/$Name.pem") {
    if ($AuthorityPassword) {
      $cred = New-Object System.Management.Automation.PSCredential -ArgumentList "ni", $AuthorityPassword
      $passin = "-passin pass:$(($cred.GetNetworkCredential().Password).Trim())"
    } else {
      $passin = ""
    }

    $param = "ca -config $($script:cnf_ca) -revoke certs/$Name.pem -crl_reason $Reason $passin"

    Write-Verbose "param: $param"

    Invoke-OpenSsl $param

    if ((Get-IssuedCertificate | Where-Object { $_.SerialNumber -eq $name }).Status -eq 'Revoked') {
      Move-Item -Path "certs/$Name.pem" "certs/$Name.pem.revoked"
    }
  } else {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
      -Message "A certificate with the name '$Name' does not exists in this authority." `
      -ExceptionType "System.Management.Automation.ItemNotFoundException" `
      -ErrorId "ItemNotFoundException" -ErrorCategory ObjectNotFound ))
  }
}

function Test-IssuedCertificateValidity {
  [CmdletBinding()]
  param (
    [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
    [Alias("Certificate", "Cert")]
    [string] $Name,
    [Parameter(Position = 1)]
    [ValidateScript({ Test-Path $(Resolve-Path $_) })]
    [Alias("CA", "CAFile", "CABundle")]
    [string] $CAPath
  )

  $result = Get-IssuedCertificateValidity @PsBoundParameters

  return $result.Contains("Verification: OK")
}

function Update-CertificateAuthorityDatabase {
  [Alias("update-ca", "ca-update")]
  param (
    [Parameter(Position = 0)]
    [ValidateScript({ Test-Path })]
    [string] $Path = ($PWD.Path),
    [securestring] $AuthorityPassword
  )

  if (-not (Test-CertificateAuthority $Path)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "This is not a certificate authority that can be managed by this module." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  Push-Location $Path

  if ($AuthorityPassword) {
    $cred = New-Object System.Management.Automation.PSCredential -ArgumentList "ni", $AuthorityPassword
    $passin = "-passin pass:$(($cred.GetNetworkCredential().Password).Trim())"
  } else {
    $passin = ""
  }

  Invoke-OpenSsl "ca -config $($script:cnf_ca) -updatedb $passin"

  Pop-Location
}

function Update-CertificateAuthorityRevocationList {
  [Alias("update-crl", "crl-update")]
  param (
    [Parameter(Position = 0)]
    [ValidateScript({ Test-Path })]
    [string] $Path = ($PWD.Path),
    [securestring] $AuthorityPassword
  )

  if (-not (Test-CertificateAuthority $Path)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "This is not a certificate authority that can be managed by this module." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  Push-Location $Path

  $Name = Get-CertificateAuthoritySetting Name

  if ($AuthorityPassword) {
    $cred = New-Object System.Management.Automation.PSCredential -ArgumentList "ni", $AuthorityPassword
    $passin = "-passin pass:$(($cred.GetNetworkCredential().Password).Trim())"
  } else {
    $passin = ""
  }

  Invoke-OpenSsl "ca -config $($script:cnf_ca) -gencrl -out $Name.crl $passin"

  if (Test-Path "$Name.crl") {
    Invoke-OpenSsl "crl -in $Name.crl -noout -text"
  }

  Pop-Location
}

function Update-OcspCertificate {
  [Alias("update-ocsp", "ocsp-update")]
  param (
    [Parameter(Position = 0)]
    [ValidateScript({ Test-Path $(Resolve-Path $_) })]
    [string] $Path = ($PWD.Path),
    [Parameter(Mandatory = $true)]
    [securestring] $AuthorityPassword,
    [int] $Days = 30,
    [switch] $Reset
  )

  if (-not (Test-CertificateAuthority $Path)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "This is not a certificate authority that can be managed by this module." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  Push-Location $Path

  if ("True" -ne $(Get-CertificateAuthoritySetting ocsp)) {
    return
  }

  $old = Get-IssuedCertificate | `
    Where-Object { $_.DistinguishedName -like "*$(Get-CertificateAuthoritySetting cn) OCSP Responder" }

  if ($Reset) {
    Write-Output "`n~~~~~`nGenerating the OCSP private key for this authority...`n"

    generateRequestConfig "ocsp.cnf" `
    $(Get-CertificateAuthoritySetting c) `
    $(Get-CertificateAuthoritySetting org) `
    "$(Get-CertificateAuthoritySetting cn) OCSP Responder"

    New-OpenSslRsaKeypair -Path "private/ocsp.key" -BitSize 2048 -NoPublicFile

    Write-Output "`nGenerating the OCSP certificate request..."

    Invoke-OpenSsl "req -new -config ocsp.cnf -out csr/ocsp.csr -key private/ocsp.key"

    Write-Output "`nGenerating the OCSP certificate for this authority..."
  }

  $cred = New-Object System.Management.Automation.PSCredential -ArgumentList "ni", $AuthorityPassword
  $passin = "-passin pass:$(($cred.GetNetworkCredential().Password).Trim())"
  $ext = "-extensions ocsp_ext"

  Invoke-OpenSsl `
    "ca -batch -config $($script:cnf_ca) -out certs/ocsp.pem $ext -days $Days $passin -infiles csr/ocsp.csr"

  foreach ($cert in $old) {
    if ($cert.Status -eq 'Valid') {
      Write-Verbose "Revoking superceded OCSP certificate: $($cert.SerialNumber)"
      Revoke-Certificate -Name $cert.SerialNumber -AuthorityPassword $AuthorityPassword -Reason superseded
    }
  }

  Pop-Location
}

function Update-TimestampCertificate {
  [Alias("update-timestamp")]
  param (
    [Parameter(Position = 0)]
    [ValidateScript({ Test-Path $(Resolve-Path $_) })]
    [string] $Path = ($PWD.Path),
    [Parameter(Mandatory = $true)]
    [securestring] $AuthorityPassword,
    [int] $Days = 30,
    [switch] $Reset
  )

  if (-not (Test-CertificateAuthority $Path)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord `
       -Message "This is not a certificate authority that can be managed by this module." `
       -ExceptionType "System.InvalidOperationException" `
       -ErrorId "System.InvalidOperation" -ErrorCategory "InvalidOperation"))
  }

  Push-Location $Path

  if ("True" -ne $(Get-CertificateAuthoritySetting timestamp)) {
    return
  }

  $old = Get-IssuedCertificate | `
    Where-Object { $_.DistinguishedName -like "*$(Get-CertificateAuthoritySetting cn) Timestamp Authority" }

  if ($Reset) {
    Write-Output "`n~~~~~`nGenerating the timestamp private key for this authority...`n"

    generateRequestConfig "timestamp.cnf" `
      $(Get-CertificateAuthoritySetting c) `
      $(Get-CertificateAuthoritySetting org) `
      "$(Get-CertificateAuthoritySetting cn) Timestamp Authority"

    New-OpenSslRsaKeypair -Path "private/timestamp.key" -BitSize 2048 -NoPublicFile

    Write-Output "`nGenerating the Timestamp certificate request..."

    Invoke-OpenSsl "req -new -config timestamp.cnf -out csr/timestamp.csr -key private/timestamp.key"

    Write-Output "`nGenerating the Timestamp certificate for this authority..."
  }

  $cred = New-Object System.Management.Automation.PSCredential -ArgumentList "ni", $AuthorityPassword
  $passin = "-passin pass:$(($cred.GetNetworkCredential().Password).Trim())"
  $ext = "-extensions timestamp_ext"

  Invoke-OpenSsl `
    "ca -batch -config $($script:cnf_ca) -out certs/timestamp.pem $ext -days $Days $passin -infiles csr/timestamp.csr"

    foreach ($cert in $old) {
      if ($cert.Status -eq 'Valid') {
        Write-Verbose "Revoking superceded Timestamp certificate: $($cert.SerialNumber)"
        Revoke-Certificate -Name $cert.SerialNumber -AuthorityPassword $AuthorityPassword -Reason superseded
      }
    }

    Pop-Location
}

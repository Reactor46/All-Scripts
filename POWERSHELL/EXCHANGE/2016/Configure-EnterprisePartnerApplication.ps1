<#
.SYNOPSIS 
    Configures a Partner Application that uses the OAuth protocol to authenticate to Exchange.

.DESCRIPTION 
    Configures a Enterprise Partner Application that self-issues OAuth tokens to successful authenticate to Exchange On-Premises. 
    The Partner Application must publish a Auth Metadata Document for Exchange On-Premises to establish a direct trust to this application and accept authentication request. 
    In addition to establishing the direct trust, RBAC roles are assigned to the Partner Application to authorize it for calling specific Exchange Web Services APIs.

.PARAMETER   AuthMetadataUrl
    The Uri to a Auth Metadata document containing the ApplicationIdentifier, Realm and Public Key information for the Partner Application.

.PARAMETER   ApplicationType
    The type of Application that indicates which RBAC roles should be assigned. 
    For a Microsoft SharePoint farm, this must be specified as 'SharePoint'.
    For Microsoft Lync this must be specified as 'Lync'. 
    For all other applications 'Generic' must be specified.

.PARAMETER   DomainController
    The Domain Controller the configuration objects should get created on.

.EXAMPLE
    .\Configure-EnterprisePartnerApplication.ps1 -AuthMetadataUrl https://mysharpointfarm/metadata/json/1 -ApplicationType Sharepoint
    Creates a Partner Application for Microsoft SharePoint

.EXAMPLE
    .\Configure-EnterprisePartnerApplication.ps1 -AuthMetadataUrl https://mylync/metadata/json/1 -ApplicationType Lync
    Creates a Partner Application for Microsoft Lync

.EXAMPLE
    .\Configure-EnterprisePartnerApplication.ps1 -AuthMetadataUrl https://mylobapp/metadata/json/1 -ApplicationType Generic
    Creates a Partner Application for a LOB Application

#>


[CmdletBinding()]
Param(
  [Parameter(Mandatory = $true)]
  [string] $AuthMetadataUrl,
  
  [Parameter(Mandatory = $true)]
  [string] $ApplicationType,

  [Parameter(Mandatory = $false)]
  [string] $DomainController,

  [Parameter(Mandatory = $false)]
  [bool] $TrustAnySSLCertificate
)

$ApplicationAccountSuffix = "ApplicationAccount";
$MaxUserRetry = 10;
$LyncAppName = "LyncEnterprise";
$SharePointAppName = "SharePointEnterprise"; 
$GenericAppName = "GenericEnterprise";
$UseDomainController = $false; 

function WriteError($str)
{
  write-host "";
  Write-host $str -ForegroundColor Red;
}

function WriteInformation($str)
{
  write-host "";
  Write-host $str -ForegroundColor Yellow;
}

function WriteSuccess($str)
{
  write-host "";
  Write-host $str -ForegroundColor Green;
}

function ConfigureLync()
{
  ### try to get user, if exists use it, if not create one 
  $user = CreateApplicationUser $LyncAppName;

  ### assign the management roles
  [string[]]$roles = @("UserApplication", "ArchiveApplication", "MeetingGraphApplication");
  AssignRolesToApplication $user $roles;

  ### create the partner application
  $app = CreatePartnerApplication $LyncAppName $authMetadataUrl $true $user;
}

function ConfigureSharepoint()
{
  ### try to get user, if exists use it, if not create one 
  $user = CreateApplicationUser $SharePointAppName;

  ### assign the management roles
  [string[]]$roles = @("UserApplication", "LegalHoldApplication", "Mailbox Search", "TeamMailboxLifecycleApplication", "Legal Hold");
  AssignRolesToApplication $user $roles;

  ### create the partner application
  $app = CreatePartnerApplication $SharePointAppName $authMetadataUrl $true $user;
}

function ConfigureGeneric([string]$authMetadataUrl)
{ 
  ### try to get user, if exists use it, if not create one 
  $user = CreateApplicationUser $GenericAppName;

  ### assign the management roles
  [string[]]$roles = @("UserApplication");
  AssignRolesToApplication $user $roles ;

  ### create the partner appliation
  $app = CreatePartnerApplication $GenericAppName $authMetadataUrl $false $user;
}

function CreatePartnerApplication([string]$appPrefix, [string]$authMetadataUrl, [bool]$acceptNii, [object]$user)
{
  $appname = [String]::Format("{0}-{1}", $appPrefix.Trim(), [System.Guid]::NewGuid().ToString("N"));

  WriteInformation ("Creating Partner Application <$appname> using metadata <$authMetadataUrl> with linked account <$($user.Identity)>.");

  if($UseDomainController -eq $true)
  {
	if($TrustAnySSLCertificate -eq $true)
	{   
		$app = New-PartnerApplication -Name $appname -AuthMetadataUrl $authMetadataUrl -AcceptSecurityIdentifierInformation $acceptNii -Enabled $true -LinkedAccount $user.Identity -DomainController $DomainController -TrustAnySSLCertificate 
	}
	else
	{
		$app = New-PartnerApplication -Name $appname -AuthMetadataUrl $authMetadataUrl -AcceptSecurityIdentifierInformation $acceptNii -Enabled $true -LinkedAccount $user.Identity -DomainController $DomainController
	}
  }
  else
  {
	if($TrustAnySSLCertificate -eq $true)
	{
		$app = New-PartnerApplication -Name $appname -AuthMetadataUrl $authMetadataUrl -AcceptSecurityIdentifierInformation $acceptNii -Enabled $true -LinkedAccount $user.Identity -TrustAnySSLCertificate;
	}
	else
	{
		$app = New-PartnerApplication -Name $appname -AuthMetadataUrl $authMetadataUrl -AcceptSecurityIdentifierInformation $acceptNii -Enabled $true -LinkedAccount $user.Identity;
	}
  }

  if($app -ne $null)
  {
    WriteInformation ("Created Partner Application <$($app.Identity)>.");
    return $app;
  }

  exit 1;
}

function CreateApplicationUser([string]$appPrefix)
{
  $username = [String]::Format("{0}-{1}", $appPrefix.Trim(), $ApplicationAccountSuffix);

  for($i=1; $i -le $MaxUserRetry; $i++)
  {
    if($UseDomainController -eq $true)
    {
      $user = Get-User -Identity $username -ErrorAction SilentlyContinue -DomainController $DomainController;
    }
    else
    {
      $user = Get-User -Identity $username -ErrorAction SilentlyContinue;
    }

    if($user -eq $null)
    { ## found a username that can get created
      break;
    }

    if($user.RecipientType -eq "MailUser" -or $user.RecipientType -eq "User")
    {
      WriteInformation ("User <$($user.Identity)> already exists. Using this user for the Partner Application.");
      return $user;
    }
      
    $username = [String]::Format("{0}-{1}", $username, $i);
    WriteInformation ("Cannot use existing User <$($user.Identity)> with RecipientType <$($user.RecipientType)>. Retry with <$username>");
  }
  
  if($user -ne $null)
  {
    WriteError ("Can not create a user for this Partner Application. Max retry reached.");
    exit 1;
  }

  WriteInformation ("Creating User <$username> for Partner Application.");

  $acceptedDomains = Get-AcceptedDomain;

  if ($acceptedDomains -eq $null)
  {
    WriteError ("There is no accepted domain so user can not be created.")
  }
  
  $acceptedDomain = $acceptedDomains[0].Name;

  if($UseDomainController -eq $true)
  {
    $user = New-MailUser -Name $username -DomainController $DomainController -ExternalEmailAddress $username@$acceptedDomain;
	set-mailuser -Identity $user.Identity -HiddenFromAddressListsEnabled $true -DomainController $DomainController
  }
  else
  {
    $user = New-MailUser -Name $username -ExternalEmailAddress $username@$acceptedDomain;
	set-mailuser -Identity $user.Identity -HiddenFromAddressListsEnabled $true; 
  }

  WriteInformation ("Created User <$($user.Identity)> for Partner Application.");

  return $user;
}

function AssignRolesToApplication([object]$user, [string[]]$rolenames)
{
  foreach($rolename in $rolenames)
  {
    $role = [String]::Format("{0}-{1}", $rolename, $user.Name);

    if($UseDomainController -eq $true)
    {
      $roleassignment = Get-ManagementRoleAssignment -DomainController $DomainController | ?{$_.Role -eq $rolename -and $_.RoleAssigneeName -eq $user.name};
    }
    else
    {
      $roleassignment = Get-ManagementRoleAssignment | ?{$_.Role -eq $rolename -and $_.RoleAssigneeName -eq $user.name};
    }

    if($roleassignment -ne $null)
    {
      WriteInformation ("Skipping role assignment: Role <$rolename> already assigned to Application User <$($user.Identity)>.");
      continue;
    } 

    WriteInformation ("Assigning role <$rolename> to Application User <$($user.Identity)>.");
    if($UseDomainController -eq $true)
    {
      $roleassignment = New-ManagementRoleAssignment -Role $rolename -User $user.Identity -DomainController $DomainController;
    }
    else
    {
      $roleassignment = New-ManagementRoleAssignment -Role $rolename -User $user.Identity;
    }

    if($roleassignment -eq $null)
    {
      exit 1;
    }
  }
}

#Make sure we trap the exception if any, and stop processing
trap
{
    break
}

if([String]::IsNullOrEmpty($DomainController) -eq $false)
{ 
    $UseDomainController = $true;
}

switch ($ApplicationType.ToLower())
{	
  "lync"       { ConfigureLync $AuthMetadataUrl; }
  "sharepoint" { ConfigureSharepoint $AuthMetadataUrl; }
  "generic"    { ConfigureGeneric $AuthMetadataUrl; }
  
  default
  {
    WriteError ("Invalid (ServiceName) specified. Possible values are 'Lync', 'SharePoint', or 'Generic'.");
    exit 1
  }
}

WriteSuccess ("THE CONFIGURATION HAS SUCCEEDED.");


# SIG # Begin signature block
# MIId1gYJKoZIhvcNAQcCoIIdxzCCHcMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUN5XmP0SBOqPTnXotYwrNvHbf
# fWmgghhkMIIEwzCCA6ugAwIBAgITMwAAAJqamxbCg9rVwgAAAAAAmjANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI5
# WhcNMTcwNjMwMTkyMTI5WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OkIxQjctRjY3Ri1GRUMyMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEApkZzIcoArX4o
# w+UTmzOJxzgIkiUmrRH8nxQVgnNiYyXy7kx7X5moPKzmIIBX5ocSdQ/eegetpDxH
# sNeFhKBOl13fmCi+AFExanGCE0d7+8l79hdJSSTOF7ZNeUeETWOP47QlDKScLir2
# qLZ1xxx48MYAqbSO30y5xwb9cCr4jtAhHoOBZQycQKKUriomKVqMSp5bYUycVJ6w
# POqSJ3BeTuMnYuLgNkqc9eH9Wzfez10Bywp1zPze29i0g1TLe4MphlEQI0fBK3HM
# r5bOXHzKmsVcAMGPasrUkqfYr+u+FZu0qB3Ea4R8WHSwNmSP0oIs+Ay5LApWeh/o
# CYepBt8c1QIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFCaaBu+RdPA6CKfbWxTt3QcK
# IC8JMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAIl6HAYUhsO/7lN8D/8YoxYAbFTD0plm82rFs1Mff9WBX1Hz
# /PouqK/RjREf2rdEo3ACEE2whPaeNVeTg94mrJvjzziyQ4gry+VXS9ZSa1xtMBEC
# 76lRlsHigr9nq5oQIIQUqfL86uiYglJ1fAPe3FEkrW6ZeyG6oSos9WPEATTX5aAM
# SdQK3W4BC7EvaXFT8Y8Rw+XbDQt9LJSGTWcXedgoeuWg7lS8N3LxmovUdzhgU6+D
# ZJwyXr5XLp2l5nvx6Xo0d5EedEyqx0vn3GrheVrJWiDRM5vl9+OjuXrudZhSj9WI
# 4qu3Kqx+ioEpG9FwqQ8Ps2alWrWOvVy891W8+RAwggYHMIID76ADAgECAgphFmg0
# AAAAAAAcMA0GCSqGSIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNjb20xGTAX
# BgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBSb290
# IENlcnRpZmljYXRlIEF1dGhvcml0eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0MDMx
# MzAzMDlaMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAf
# BgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQTCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJ+hbLHf20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP7tGn
# 0UytdDAgEesH1VSVFUmUG0KSrphcMCbaAGvoe73siQcP9w4EmPCJzB/LMySHnfL0
# Zxws/HvniB3q506jocEjU8qN+kXPCdBer9CwQgSi+aZsk2fXKNxGU7CG0OUoRi4n
# rIZPVVIM5AMs+2qQkDBuh/NZMJ36ftaXs+ghl3740hPzCLdTbVK0RZCfSABKR2YR
# JylmqJfk0waBSqL5hKcRRxQJgp+E7VV4/gGaHVAIhQAQMEbtt94jRrvELVSfrx54
# QTF3zJvfO4OToWECtR0Nsfz3m7IBziJLVP/5BcPCIAsCAwEAAaOCAaswggGnMA8G
# A1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFCM0+NlSRnAK7UD7dvuzK7DDNbMPMAsG
# A1UdDwQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQOrIJg
# QFYnl+UlE/wq4QpTlVnkpKFjpGEwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcG
# CgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3Qg
# Q2VydGlmaWNhdGUgQXV0aG9yaXR5ghB5rRahSqClrUxzWPQHEy5lMFAGA1UdHwRJ
# MEcwRaBDoEGGP2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1
# Y3RzL21pY3Jvc29mdHJvb3RjZXJ0LmNybDBUBggrBgEFBQcBAQRIMEYwRAYIKwYB
# BQUHMAKGOGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9z
# b2Z0Um9vdENlcnQuY3J0MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEB
# BQUAA4ICAQAQl4rDXANENt3ptK132855UU0BsS50cVttDBOrzr57j7gu1BKijG1i
# uFcCy04gE1CZ3XpA4le7r1iaHOEdAYasu3jyi9DsOwHu4r6PCgXIjUji8FMV3U+r
# kuTnjWrVgMHmlPIGL4UD6ZEqJCJw+/b85HiZLg33B+JwvBhOnY5rCnKVuKE5nGct
# xVEO6mJcPxaYiyA/4gcaMvnMMUp2MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tOi3/F
# NSteo7/rvH0LQnvUU3Ih7jDKu3hlXFsBFwoUDtLaFJj1PLlmWLMtL+f5hYbMUVbo
# nXCUbKw5TNT2eb+qGHpiKe+imyk0BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXjpKh0
# NbhOxXEjEiZ2CzxSjHFaRkMUvLOzsE1nyJ9C/4B5IYCeFTBm6EISXhrIniIh0EPp
# K+m79EjMLNTYMoBMJipIJF9a6lbvpt6Znco6b72BJ3QGEe52Ib+bgsEnVLaxaj2J
# oXZhtG6hE6a/qkfwEm/9ijJssv7fUciMI8lmvZ0dhxJkAj0tr1mPuOQh5bWwymO0
# eFQF1EEuUKyUsKV4q7OglnUa2ZKHE3UiLzKoCG6gW4wlv6DvhMoh1useT8ma7kng
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TCCBhAwggP4
# oAMCAQICEzMAAABkR4SUhttBGTgAAAAAAGQwDQYJKoZIhvcNAQELBQAwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMTAeFw0xNTEwMjgyMDMxNDZaFw0xNzAx
# MjgyMDMxNDZaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29ycG9yYXRpb24w
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCTLtrY5j6Y2RsPZF9NqFhN
# FDv3eoT8PBExOu+JwkotQaVIXd0Snu+rZig01X0qVXtMTYrywPGy01IVi7azCLiL
# UAvdf/tqCaDcZwTE8d+8dRggQL54LJlW3e71Lt0+QvlaHzCuARSKsIK1UaDibWX+
# 9xgKjTBtTTqnxfM2Le5fLKCSALEcTOLL9/8kJX/Xj8Ddl27Oshe2xxxEpyTKfoHm
# 5jG5FtldPtFo7r7NSNCGLK7cDiHBwIrD7huTWRP2xjuAchiIU/urvzA+oHe9Uoi/
# etjosJOtoRuM1H6mEFAQvuHIHGT6hy77xEdmFsCEezavX7qFRGwCDy3gsA4boj4l
# AgMBAAGjggF/MIIBezAfBgNVHSUEGDAWBggrBgEFBQcDAwYKKwYBBAGCN0wIATAd
# BgNVHQ4EFgQUWFZxBPC9uzP1g2jM54BG91ev0iIwUQYDVR0RBEowSKRGMEQxDTAL
# BgNVBAsTBE1PUFIxMzAxBgNVBAUTKjMxNjQyKzQ5ZThjM2YzLTIzNTktNDdmNi1h
# M2JlLTZjOGM0NzUxYzRiNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzcitW2oynUC
# lTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# b3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEGCCsGAQUF
# BwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0MAwGA1Ud
# EwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAIjiDGRDHd1crow7hSS1nUDWvWas
# W1c12fToOsBFmRBN27SQ5Mt2UYEJ8LOTTfT1EuS9SCcUqm8t12uD1ManefzTJRtG
# ynYCiDKuUFT6A/mCAcWLs2MYSmPlsf4UOwzD0/KAuDwl6WCy8FW53DVKBS3rbmdj
# vDW+vCT5wN3nxO8DIlAUBbXMn7TJKAH2W7a/CDQ0p607Ivt3F7cqhEtrO1Rypehh
# bkKQj4y/ebwc56qWHJ8VNjE8HlhfJAk8pAliHzML1v3QlctPutozuZD3jKAO4WaV
# qJn5BJRHddW6l0SeCuZmBQHmNfXcz4+XZW/s88VTfGWjdSGPXC26k0LzV6mjEaEn
# S1G4t0RqMP90JnTEieJ6xFcIpILgcIvcEydLBVe0iiP9AXKYVjAPn6wBm69FKCQr
# IPWsMDsw9wQjaL8GHk4wCj0CmnixHQanTj2hKRc2G9GL9q7tAbo0kFNIFs0EYkbx
# Cn7lBOEqhBSTyaPS6CvjJZGwD0lNuapXDu72y4Hk4pgExQ3iEv/Ij5oVWwT8okie
# +fFLNcnVgeRrjkANgwoAyX58t0iqbefHqsg3RGSgMBu9MABcZ6FQKwih3Tj0DVPc
# gnJQle3c6xN3dZpuEgFcgJh/EyDXSdppZzJR4+Bbf5XA/Rcsq7g7X7xl4bJoNKLf
# cafOabJhpxfcFOowMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkqhkiG9w0B
# AQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEyMDAG
# A1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IDIwMTEw
# HhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBT
# aWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA
# q/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03a8YS2Avw
# OMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akrrnoJr9eW
# WcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0RrrgOGSsbmQ1
# eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy4BI6t0le
# 2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9sbKvkjh+
# 0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAhdCVfGCi2
# zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8kA/DRelsv
# 1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTBw3J64HLn
# JN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmnEyimp31n
# gOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90lfdu+Hgg
# WCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0wggHpMBAG
# CSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2oynUClTAZ
# BgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/
# BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBaBgNVHR8E
# UzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9k
# dWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsGAQUFBwEB
# BFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9j
# ZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNVHSAEgZcw
# gZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsGAQUFBwIC
# MDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABlAG0AZQBu
# AHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKbC5YR4WOS
# mUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11lhJB9i0ZQ
# VdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6I/MTfaaQ
# dION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0wI/zRive
# /DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560STkKxgrC
# xq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQamASooPoI/
# E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGaJ+HNpZfQ
# 7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ahXJbYANah
# Rr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA9Z74v2u3
# S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33VtY5E90Z1W
# Tk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr/Xmfwb1t
# bWrJUnMTDXpQzTGCBNwwggTYAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB8DAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUNOP/W1xc7Xvg5TAKYF7pIuKHukswgY8GCisG
# AQQBgjcCAQwxgYAwfqBWgFQAQwBvAG4AZgBpAGcAdQByAGUALQBFAG4AdABlAHIA
# cAByAGkAcwBlAFAAYQByAHQAbgBlAHIAQQBwAHAAbABpAGMAYQB0AGkAbwBuAC4A
# cABzADGhJIAiaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdlIDANBgkq
# hkiG9w0BAQEFAASCAQB8THLexH1VgsvoQo4/JQPB+GfBw3h0RlzdUr95L3rSI1d2
# PoLrnU6mBIbCPeHW2JA5UchJwoQydbqI5QcolWJ6LGZazjVHmyo4Nliq/MyMagEA
# 0mRZLGBkAYFQ/9beu6KjJjbJRpKY2/eFmRbOE3j4gVHuxembex40d6jCDeKzObME
# UyxhgC7q4gi7a7nkm9NDUkJ9/Y0/zgZaQ2egWFrsvc3os9JV6brFmBRHkuIGAUEY
# Yo7Hc/u39EgY7TJWtLaG0t9UESO8N8PjH0nAZkIgDO3cEYM+1OlP9pu4u0iWWvtu
# GbGls0aMMuUKWq7d/QGpS3gKXG+aXoYjzEIrnOkXoYICKDCCAiQGCSqGSIb3DQEJ
# BjGCAhUwggIRAgEBMIGOMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5n
# dG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
# YXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQQITMwAAAJqa
# mxbCg9rVwgAAAAAAmjAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3
# DQEHATAcBgkqhkiG9w0BCQUxDxcNMTYwOTAzMTg0NDM5WjAjBgkqhkiG9w0BCQQx
# FgQUu/lg0F+yrceMi4kxK0gQ+9AHJd4wDQYJKoZIhvcNAQEFBQAEggEAjjGwZ59W
# 0mV5fntDrrCvJHC62806U8DdT/kTvpYIl1E8Oep2Vs0K/N/6D9gXzvf/d4PIQBaY
# IdVWkXRiv0BMqERTDrcWRgq1bfFmzMCjyLtAeWnxkfLbhpq4/FYRCGSgIQJFvEV1
# 3NvE2Q/JfExkSYauayIb8+7enww89hZJ0wz45fI/qYxqU55kHKtMy7g1aDu8roMs
# OEP4lW48ZCWFiveJMv2GZLrEeq3Oho88UINbRon1Ugw0HK7XS7EtvpUQQ4KtQUf7
# E6qqO+sAGDXWHEDyqdETIT3yiT7oFI8yN2iVRTIhruZ6LoF8qz6xwwBTTa2X1cxR
# J3jh+R6DThr46g==
# SIG # End signature block

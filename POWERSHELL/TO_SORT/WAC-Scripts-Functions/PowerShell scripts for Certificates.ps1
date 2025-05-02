function Export-WACCMCertificate {
<#

.SYNOPSIS
Script that exports certificate.

.DESCRIPTION
Script that exports certificate.

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [boolean]
    $fromTaskScheduler,
    [Parameter(Mandatory = $true)]
    [String]
    $certPath,
    [Parameter(Mandatory = $true)]
    [String]
    $exportType,
    [String]
    $fileName,
    [String]
    $exportChain,
    [String]
    $exportProperties,
    [String]
    $usersAndGroups,
    [String]
    $password,
    [String]
    $tempPath,
    [String]
    $resultFile,
    [String]
    $errorFile
)

BEGIN {
    Set-StrictMode -Version 5.0

    Import-Module -Name Microsoft.PowerShell.Management -ErrorAction SilentlyContinue

    Set-Variable -Name LogName -Option ReadOnly -Value "Microsoft-ServerManagementExperience" -Scope Script -ErrorAction SilentlyContinue
    Set-Variable -Name LogSource -Option ReadOnly -Value "SMEScripts" -Scope Script -ErrorAction SilentlyContinue
    Set-Variable -Name WaitTimeOut -Option ReadOnly -Value 30000 -Scope Script -ErrorAction SilentlyContinue        # 30 second timeout
    Set-Variable -Name ScriptName -Option ReadOnly -Value "Export-Certificate.ps1" -Scope Script -ErrorAction SilentlyContinue
    Set-Variable -Name RsaProviderInstanceName -Option ReadOnly -Value "RSA" -Scope Script -ErrorAction SilentlyContinue
}
PROCESS {
    <#

    .SYNOPSIS
    Helper function to write the info logs to info stream.

    .DESCRIPTION
    Helper function to write the info logs to info stream.

    .PARAMETER logMessage
    log message

    #>

    function writeInfoLog($logMessage) {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Information `
            -Message $logMessage -ErrorAction SilentlyContinue
    }

    <#

    .SYNOPSIS
    Helper function to write the info logs to info stream.

    .DESCRIPTION
    Helper function to write the info logs to info stream.

    .PARAMETER logMessage
    log message

    #>

    function writeErrorLog($errorMessage) {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message $errorMessage -ErrorAction SilentlyContinue
    }

    function exportCertificate() {
        param (
            [String]
            $certPath,
            [String]
            $tempPath,
            [String]
            $exportType,
            [String]
            $exportChain,
            [String]
            $exportProperties,
            [String]
            $usersAndGroups,
            [String]
            $password,
            [String]
            $resultFile,
            [String]
            $errorFile
        )
        try {
            Import-Module PKI
            if ($exportChain -eq "CertificateChain") {
                $chainOption = "BuildChain";
            }
            else {
                $chainOption = "EndEntityCertOnly";
            }

            $ExportPfxCertParams = @{ Cert = $certPath; FilePath = $tempPath; ChainOption = $chainOption }
            if ($exportProperties -ne "Extended") {
                $ExportPfxCertParams.NoProperties = $true
            }

            # Decrypt user encrypted password
            if ($password) {
                Add-Type -AssemblyName System.Security
                $encode = New-Object System.Text.UTF8Encoding
                $encrypted = [System.Convert]::FromBase64String($password)
                $decrypted = [System.Security.Cryptography.ProtectedData]::Unprotect($encrypted, $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
                $newPassword = $encode.GetString($decrypted)
                $securePassword = ConvertTo-SecureString -String $newPassword -Force -AsPlainText;
                $ExportPfxCertParams.Password = $securePassword
            }

            if ($usersAndGroups) {
                $ExportPfxCertParams.ProtectTo = $usersAndGroups
            }

            Export-PfxCertificate @ExportPfxCertParams | ConvertTo-Json -Depth 10 | Out-File $ResultFile
        }
        catch {
            $_.Exception.Message | ConvertTo-Json | Out-File $ErrorFile
        }
    }

    function CalculateFilePath {
        param (
            [Parameter(Mandatory = $true)]
            [String]
            $exportType,
            [Parameter(Mandatory = $true)]
            [String]
            $certPath
        )

        $extension = $exportType.ToLower();
        if ($exportType -ieq "cert") {
            $extension = "cer";
        }

        if (!$fileName) {
            try {
                $fileName = [IO.Path]::GetFileName($certPath);
            }
            catch {
                $err = $_.Exception.Message
                writeErrorLog "An error occured attempting to extract file name from certificate path. Exception: $err"
                throw $err
            }
        }

        try {
            $path = Join-Path $env:TEMP ([IO.Path]::ChangeExtension($filename, $extension))
        }
        catch {
            $err = $_.Exception.Message
            writeErrorLog "An error occured attempting to join file name to file extension. Exception: $err"
            throw $err
        }

        writeInfoLog "Calculated file name: $fileName."
        return $path
    }

    function DecryptPasswordWithJWKOnNode($encryptedJWKPassword) {
        if (Get-Variable -Scope Script -Name $RsaProviderInstanceName -ErrorAction SilentlyContinue) {
            $rsaProvider = (Get-Variable -Scope Script -Name $RsaProviderInstanceName).Value
            $decryptedBytes = $rsaProvider.Decrypt([Convert]::FromBase64String($encryptedJWKPassword), [System.Security.Cryptography.RSAEncryptionPadding]::OaepSHA1)

            return [System.Text.Encoding]::UTF8.GetString($decryptedBytes)
        }
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message "[$ScriptName]: Password decryption failed. RSACryptoServiceProvider Instance not found" -ErrorAction SilentlyContinue
        # TODO: Localize this message!
        throw [System.InvalidOperationException] "Password decryption failed. RSACryptoServiceProvider Instance not found"
    }

    ###########################################################################
    # Script execution starts here
    ###########################################################################

    $tempPath = CalculateFilePath -exportType $exportType -certPath $certPath;
    if ($exportType -ne "Pfx") {
        Export-Certificate -Cert $certPath -FilePath $tempPath -Type $exportType -Force
        return;
    }

    $stopwatch = [System.Diagnostics.Stopwatch]::new()

    function isSystemLockdownPolicyEnforced() {
        return [System.Management.Automation.Security.SystemPolicy]::GetSystemLockdownPolicy() -eq [System.Management.Automation.Security.SystemEnforcementMode]::Enforce
    }
    $isWdacEnforced = isSystemLockdownPolicyEnforced;

    # In WDAC environment script file will already be available on the machine
    # In WDAC mode the same script is executed - once normally and once through task Scheduler
    if ($isWdacEnforced) {
        if ($fromTaskScheduler) {
            exportCertificate $certPath $tempPath $exportType $exportChain $exportProperties $usersAndGroups $password $resultFile $errorFile;
            return;
        }
    }
    else {
        # In non-WDAC environment script file will not be available on the machine
        # Hence, a dynamic script is created which is executed through the task Scheduler
        $ScriptFile = $env:temp + "\export-certificate.ps1"
    }

    # PFX private key handlings
    if ($password) {
        $decryptedJWKPassword = DecryptPasswordWithJWKOnNode $password
        # encrypt password with current user.
        Add-Type -AssemblyName System.Security
        $encode = New-Object System.Text.UTF8Encoding
        $bytes = $encode.GetBytes($decryptedJWKPassword)
        $encrypt = [System.Security.Cryptography.ProtectedData]::Protect($bytes, $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
        $userEncryptedPassword = [System.Convert]::ToBase64String($encrypt)
    }

    # Pass parameters to script and generate script file in temp folder
    $resultFile = $env:temp + "\export-certificate_result.json"
    $errorFile = $env:temp + "\export-certificate_error.json"
    if (Test-Path $errorFile) {
        Remove-Item $errorFile
    }

    if (Test-Path $resultFile) {
        Remove-Item $resultFile
    }

    # Create a scheduled task
    $TaskName = "SMEExportCertificate"
    $User = [Security.Principal.WindowsIdentity]::GetCurrent()

    $HashArguments = @{};
    if ($exportChain) {
        $HashArguments.Add("exportChain", $exportChain)
    }

    if ($exportProperties) {
        $HashArguments.Add("exportProperties", $exportProperties)
    }

    if ($usersAndGroups) {
        $HashArguments.Add("usersAndGroups", $usersAndGroups)
    }

    if ($userEncryptedPassword) {
        $HashArguments.Add("password", $userEncryptedPassword)
    }

    $tempArgs = ""
    foreach ($key in $HashArguments.Keys) {
        $value = $HashArguments[$key]
        $value = """$value"""
        $tempArgs += " -$key $value"
    }

    if ($isWdacEnforced) {
        $arg = "-NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -command ""&{Import-Module Microsoft.SME.CertificateManager; Export-WACCMCertificate -fromTaskScheduler `$true -exportType $exportType $tempArgs -certPath $certPath -tempPath $tempPath -resultFile $resultFile -errorFile $errorFile}"""
    }
    else {
        (Get-Command exportCertificate).ScriptBlock | Set-Content -Path $ScriptFile
        $arg = "-NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -File $ScriptFile -certPath $certPath -exportType $exportType $tempArgs -tempPath $tempPath -resultFile $resultFile -errorFile $errorFile"
    }

    if ($null -eq (New-Object Security.Principal.WindowsPrincipal $User).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
        Write-Warning "To perform some operations you must run an elevated Windows PowerShell console."
    }

    $Scheduler = New-Object -ComObject Schedule.Service

    # Try to connect to schedule service 3 time since it may fail the first time
    for ($i = 1; $i -le 3; $i++) {
        try {
            $Scheduler.Connect()
            Break
        }
        catch {
            if ($i -ge 3) {
                $message = $_.Exception.Message
                writeErrorLog $message
                Write-Error $message -ErrorAction Stop
            }
            else {
                Start-Sleep -s 1
            }
        }
    }

    $RootFolder = $Scheduler.GetFolder("\")
    # Delete existing task
    if ($RootFolder.GetTasks(0) | Where-Object { $_.Name -eq $TaskName }) {
        Write-Debug("Deleting existing task" + $TaskName)
        $RootFolder.DeleteTask($TaskName, 0)
    }

    $Task = $Scheduler.NewTask(0)
    $RegistrationInfo = $Task.RegistrationInfo
    $RegistrationInfo.Description = $TaskName
    $RegistrationInfo.Author = $User.Name

    $Triggers = $Task.Triggers
    $Trigger = $Triggers.Create(7) #TASK_TRIGGER_REGISTRATION: Starts the task when the task is registered.
    $Trigger.Enabled = $true

    $Settings = $Task.Settings
    $Settings.Enabled = $True
    $Settings.StartWhenAvailable = $True
    $Settings.Hidden = $False

    $Action = $Task.Actions.Create(0)
    $Action.Path = "powershell"
    $Action.Arguments = $arg

    # Tasks will be run with the highest privileges
    $Task.Principal.RunLevel = 1
    $RootFolder.RegisterTaskDefinition($TaskName, $Task, 6, "SYSTEM", $Null, 1) | Out-Null

    # Wait for running task finished
    $stopWatch.Start()
    $RootFolder.GetTask($TaskName).Run(0) | Out-Null
    while ($Scheduler.GetRunningTasks(0) | Where-Object { $_.Name -eq $TaskName }) {
        Start-Sleep -s 2

        $now = $stopWatch.Elapsed.Milliseconds
        if ($now -ge $WaitTimeOut) {
            $message = 'Timed out waiting for the the scheduled task that exports the certificate to complete.'

            writeErrorLog $message
            throw $message
        }
    }

    # Clean up
    $RootFolder.DeleteTask($TaskName, 0)
    if (!$isWdacEnforced) {
        Remove-Item $ScriptFile
    }

    if (Test-Path $ErrorFile) {
        $result = Get-Content -Raw -Path $ErrorFile | ConvertFrom-Json
        Remove-Item $ErrorFile
        Remove-Item $ResultFile
        throw $result
    }

    # Return result
    if (Test-Path $ResultFile) {
        $result = Get-Content -Raw -Path $ResultFile | ConvertFrom-Json
        Remove-Item $ResultFile
        return $result
    }
}
END {
    Remove-Variable -Name LogName -Scope Script -Force
    Remove-Variable -Name LogSource -Scope Script -Force
    Remove-Variable -Name WaitTimeOut -Scope Script -Force
    Remove-Variable -Name ScriptName -Scope Script -Force
    Remove-Variable -Name RsaProviderInstanceName -Scope Script -Force
}

}
## [END] Export-WACCMCertificate ##
function Get-WACCMCertificateOverview {
<#

.SYNOPSIS
Script that get the certificates overview (total, ex) in the system.

.DESCRIPTION
Script that get the certificates overview (total, ex) in the system.

.ROLE
Readers

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $EventChannels,
    [String]
    $CertificatePath = "Cert:\",
    [int]
    $NearlyExpiredThresholdInDays = 60
)

BEGIN {
    Set-StrictMode -Version 5.0

    Import-Module Microsoft.PowerShell.Diagnostics -ErrorAction SilentlyContinue

    Set-Variable -Name LogName -Option ReadOnly -Value "Microsoft-ServerManagementExperience" -Scope Script -ErrorAction SilentlyContinue
    Set-Variable -Name LogSource -Option ReadOnly -Value "SMEScripts" -Scope Script -ErrorAction SilentlyContinue
}
PROCESS {
    # Notes: $channelList must be in this format:
    # "Microsoft-Windows-CertificateServicesClient-Lifecycle-System*,
    # Microsoft-Windows-CertificateServices-Deployment*,
    # Microsoft-Windows-CertificateServicesClient-CredentialRoaming*,
    # Microsoft-Windows-CertificateServicesClient-Lifecycle-User*,
    # Microsoft-Windows-CAPI2*,Microsoft-Windows-CertPoleEng*"

    function Get-ChildLeafRecurse($psPath) {
        try {
            Get-ChildItem -Path $pspath -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object NotAfter, PSIsContainer, Location, PSChildName | ForEach-Object {
                if (!$_.PSIsContainer) {
                    $_
                }
                else {
                    $location = "Cert:\$($_.Location)"

                    if ($_.psChildName -ne $_.Location) {
                        $location += "\$($_.PSChildName)"
                    }

                    Get-ChildLeafRecurse $location
                }
            } | Microsoft.PowerShell.Utility\Select-Object NotAfter
        } 
        catch [System.ComponentModel.Win32Exception] {
            # When running this script remotely/non-elevated at least one store 'Cert:\CurrentUser\UserDS' cannot be
            # opened and traversed.  Logging an info record in case we ever need to investigate further.  An Error
            # record would be too chatty since this happens almost every time we run this script remotely/non-elevated.
            $message = "Error: '$_' for certificate store $location"
            Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Information `
                -Message $message -ErrorAction SilentlyContinue
        }
    }

    function main([String] $eventChannels, [String] $path, [int] $earlyExpiredThresholdInDays) {
        $stopwatch = [System.Diagnostics.Stopwatch]::new()

        $payload = New-Object -TypeName psobject

        $stopwatch.Start()
        $certs = Get-ChildLeafRecurse -pspath $path
        $stopwatch.Stop()
        $queryTime = $stopwatch.Elapsed.TotalMilliseconds

        $stopwatch.Restart()

        $expiredCount = @($certs | Where-Object { $_.NotAfter -lt [DateTime]::Now })
        $nearlyExpiredCount = @($certs | Where-Object { ($_.NotAfter -gt [DateTime]::Now ) -and ($_.NotAfter -lt [DateTime]::Now.AddDays($nearlyExpiredThresholdInDays) ) })

        $channelList = @($eventChannels.split(","))
        $eventCount = 0
        Get-WinEvent -ListLog $channelList -Force -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object RecordCount | ForEach-Object {
            $eventCount += $_.RecordCount
        }

        $payload | add-member -Name "total" -Value $certs.length -MemberType NoteProperty
        $payload | add-member -Name "expired" -Value $expiredCount.length -MemberType NoteProperty
        $payload | add-member -Name "nearlyExpired" -Value $nearlyExpiredCount.length -MemberType NoteProperty
        $payload | add-member -Name "eventCount" -Value $eventCount -MemberType NoteProperty

        $stopwatch.Stop()
        $totalTime = $queryTime + $stopwatch.Elapsed.TotalMilliseconds

        $payload | add-member -Name "certQueryTime" -Value "$queryTime ms" -MemberType NoteProperty
        $payload | add-member -Name "totalTimeInScript" -Value "$totalTime ms" -MemberType NoteProperty

        return $payload
    }

    ###########################################################################
    # Script execution start here
    ###########################################################################

    return main $EventChannels $CertificatePath $NearlyExpiredThresholdInDays
}
END {
    Remove-Variable -Name LogName -Scope Script -Force
    Remove-Variable -Name LogSource -Scope Script -Force
}
}
## [END] Get-WACCMCertificateOverview ##
function Get-WACCMCertificateScopes {
<#

.SYNOPSIS
Script that enumerates all the certificate scopes/locations in the system.

.DESCRIPTION
Script that enumerates all the certificate scopes/locations in the system.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

Get-ChildItem | Microsoft.PowerShell.Utility\Select-Object -Property @{name ="Name";expression= {$($_.LocationName)}}


}
## [END] Get-WACCMCertificateScopes ##
function Get-WACCMCertificateStores {
<#

.SYNOPSIS
Script that enumerates all the certificate stores in the system inside the scope/location.

.DESCRIPTION
Script that enumerates all the certificate stores in the system inside the scope/location.

.ROLE
Readers

#>

Param([string]$scope)

Set-StrictMode -Version 5.0

Get-ChildItem $('Cert:' + $scope) | Microsoft.PowerShell.Utility\Select-Object Name, @{name ="Path";expression= {$($_.Location.toString() + '\' + $_.Name)}}

}
## [END] Get-WACCMCertificateStores ##
function Get-WACCMCertificates {
<#

.SYNOPSIS
Script that enumerates all the certificates in the system.

.DESCRIPTION
Script that enumerates all the certificates in the system.

.ROLE
Readers

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $path,
    [int]
    $nearlyExpiredThresholdInDays = 60
)

Set-StrictMode -Version 5.0

<#
.Synopsis
    Name: GetChildLeafRecurse
    Description: Recursively enumerates each scope and store in Cert:\ drive.

.Parameters
    $pspath: The initial pspath to use for creating whole path to certificate store.

.Returns
    The constructed ps-path object.
#>
function GetChildLeafRecurse {
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $pspath
    )
    try {
        Get-ChildItem -Path $pspath -ErrorAction SilentlyContinue | ForEach-Object {
            if (!$_.PSIsContainer) {
                $_
            }
            else {
                $location = "Cert:\$($_.location)";
                if ($_.psChildName -ne $_.location) {
                    $location += "\$($_.PSChildName)";
                }

                GetChildLeafRecurse $location
            }
        }
    }
    catch {}
}

<#
.Synopsis
    Name: ComputePublicKey
    Description: Computes public key algorithm and public key parameters

.Parameters
    $cert: The original certificate object.

.Returns
    A hashtable object of public key algorithm and public key parameters.
#>
function ComputePublicKey {
    param (
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]
        $cert
    )

    $publicKeyInfo = @{}

    $publicKeyInfo["PublicKeyAlgorithm"] = ""
    $publicKeyInfo["PublicKeyParameters"] = ""

    if ($cert.PublicKey) {
        $publicKeyInfo["PublicKeyAlgorithm"] = $cert.PublicKey.Oid.FriendlyName
        $publicKeyInfo["PublicKeyParameters"] = $cert.PublicKey.EncodedParameters.Format($true)
    }

    $publicKeyInfo
}

<#
.Synopsis
    Name: ComputeSignatureAlgorithm
    Description: Computes signature algorithm out of original certificate object.

.Parameters
    $cert: The original certificate object.

.Returns
    The signature algorithm friendly name.
#>
function ComputeSignatureAlgorithm {
    param (
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]
        $cert
    )

    $signatureAlgorithm = [System.String]::Empty

    if ($cert.SignatureAlgorithm) {
        $signatureAlgorithm = $cert.SignatureAlgorithm.FriendlyName;
    }

    $signatureAlgorithm
}

<#
.Synopsis
    Name: ComputePrivateKeyStatus
    Description: Computes private key exportable status.
.Parameters
    $hasPrivateKey: A flag indicating certificate has a private key or not.
    $canExportPrivateKey: A flag indicating whether certificate can export a private key.

.Returns
    Enum values "Exported" or "NotExported"
#>
function ComputePrivateKeyStatus {
    param (
        [Parameter(Mandatory = $true)]
        [bool]
        $hasPrivateKey,

        [Parameter(Mandatory = $true)]
        [bool]
        $canExportPrivateKey
    )

    if (-not ($hasPrivateKey)) {
        $privateKeystatus = "None"
    }
    else {
        if ($canExportPrivateKey) {
            $privateKeystatus = "Exportable"
        }
        else {
            $privateKeystatus = "NotExportable"
        }
    }

    $privateKeystatus
}

<#
.Synopsis
    Name: ComputeExpirationStatus
    Description: Computes expiration status based on notAfter date.
.Parameters
    $notAfter: A date object refering to certificate expiry date.

.Returns
    Enum values "Expired", "NearlyExpired" and "Healthy"
#>
function ComputeExpirationStatus {
    param (
        [Parameter(Mandatory = $true)]
        [DateTime]$notAfter
    )

    if ([DateTime]::Now -gt $notAfter) {
        $expirationStatus = "Expired"
    }
    else {
        $nearlyExpired = [DateTime]::Now.AddDays($nearlyExpiredThresholdInDays);

        if ($nearlyExpired -ge $notAfter) {
            $expirationStatus = "NearlyExpired"
        }
        else {
            $expirationStatus = "Healthy"
        }
    }

    $expirationStatus
}

<#
.Synopsis
    Name: ComputeArchivedStatus
    Description: Computes archived status of certificate.
.Parameters
    $archived: A flag to represent archived status.

.Returns
    Enum values "Archived" and "NotArchived"
#>
function ComputeArchivedStatus {
    param (
        [Parameter(Mandatory = $true)]
        [bool]
        $archived
    )

    if ($archived) {
        $archivedStatus = "Archived"
    }
    else {
        $archivedStatus = "NotArchived"
    }

    $archivedStatus
}

<#
.Synopsis
    Name: ComputeIssuedTo
    Description: Computes issued to field out of the certificate subject.
.Parameters
    $subject: Full subject string of the certificate.

.Returns
    Issued To authority name.
#>
function ComputeIssuedTo {
    param (
        [String]
        $subject
    )

    $issuedTo = [String]::Empty

    $issuedToRegex = "CN=(?<issuedTo>[^,?]+)"
    $matched = $subject -match $issuedToRegex

    if ($matched -and $Matches) {
        $issuedTo = $Matches["issuedTo"]
    }

    $issuedTo
}

<#
.Synopsis
    Name: ComputeIssuerName
    Description: Computes issuer name of certificate.
.Parameters
    $cert: The original cert object.

.Returns
    The Issuer authority name.
#>
function ComputeIssuerName {
    param (
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]
        $cert
    )

    $issuerName = $cert.GetNameInfo([System.Security.Cryptography.X509Certificates.X509NameType]::SimpleName, $true)

    $issuerName
}

<#
.Synopsis
    Name: ComputeCertificateName
    Description: Computes certificate name of certificate.
.Parameters
    $cert: The original cert object.

.Returns
    The certificate name.
#>
function ComputeCertificateName {
    param (
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]
        $cert
    )

    $certificateName = $cert.GetNameInfo([System.Security.Cryptography.X509Certificates.X509NameType]::SimpleName, $false)
    if (!$certificateName) {
        $certificateName = $cert.GetNameInfo([System.Security.Cryptography.X509Certificates.X509NameType]::DnsName, $false)
    }

    $certificateName
}

<#
.Synopsis
    Name: ComputeStore
    Description: Computes certificate store name.
.Parameters
    $pspath: The full certificate ps path of the certificate.

.Returns
    The certificate store name.
#>
function ComputeStore {
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $pspath
    )

    $pspath.Split('\')[2]
}

<#
.Synopsis
    Name: ComputeScope
    Description: Computes certificate scope/location name.
.Parameters
    $pspath: The full certificate ps path of the certificate.

.Returns
    The certificate scope/location name.
#>
function ComputeScope {
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $pspath
    )

    $pspath.Split('\')[1].Split(':')[2]
}

<#
.Synopsis
    Name: ComputePath
    Description: Computes certificate path. E.g. CurrentUser\My\<thumbprint>
.Parameters
    $pspath: The full certificate ps path of the certificate.

.Returns
    The certificate path.
#>
function ComputePath {
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $pspath
    )

    $pspath.Split(':')[2]
}


<#
.Synopsis
    Name: EnhancedKeyUsageList
    Description: Enhanced KeyUsage
.Parameters
    $cert: The original cert object.

.Returns
    Enhanced Key Usage.
#>
function EnhancedKeyUsageList {
    param (
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]
        $cert
    )

    $usageString = ''
    foreach ( $usage in $cert.EnhancedKeyUsageList) {
        $usageString = $usageString + $usage.FriendlyName + ' ' + $usage.ObjectId + "`n"
    }

    $usageString
}

<#
.Synopsis
    Name: ComputeTemplate
    Description: Compute template infomation of a certificate
    $certObject: The original certificate object.

.Returns
    The certificate template if there is one otherwise empty string
#>
function ComputeTemplate {
    param (
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]
        $cert
    )

    $template = $cert.Extensions | Where-Object { $_.Oid.FriendlyName -match "Template" }
    if ($template) {
        $name = $template.Format(1).split('(')[0]
        if ($name) {
            $name -replace "Template="
        }
        else {
            ''
        }
    }
    else {
        ''
    }
}

<#
.Synopsis
    Name: ExtractCertInfo
    Description: Extracts certificate info by decoding different field and create a custom object.
.Parameters
    $certObject: The original certificate object.

.Returns
    The custom object for certificate.
#>
function ExtractCertInfo {
    param (
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]
        $certObject,
        [Parameter(Mandatory = $true)]
        [boolean]
        $hasPrivateKey,
        [Parameter(Mandatory = $true)]
        [boolean]
        $canExportPrivateKey
    )

    $publicKeyInfo = $(ComputePublicKey $certObject)
    return @{
        Archived            = $(ComputeArchivedStatus $certObject.Archived)
        CertificateName     = $(ComputeCertificateName $certObject)

        EnhancedKeyUsage    = $(EnhancedKeyUsageList $certObject)
        FriendlyName        = $certObject.FriendlyName
        IssuerName          = $(ComputeIssuerName $certObject)
        IssuedTo            = $(ComputeIssuedTo $certObject.Subject)
        Issuer              = $certObject.Issuer

        NotAfter            = $certObject.NotAfter
        NotBefore           = $certObject.NotBefore

        Path                = $(ComputePath  $certObject.PsPath)
        PrivateKey          = $(ComputePrivateKeyStatus -hasPrivateKey $hasPrivateKey -canExportPrivateKey  $canExportPrivateKey)
        PublicKey           = $publicKeyInfo.PublicKeyAlgorithm
        PublicKeyParameters = $publicKeyInfo.PublicKeyParameters

        Scope               = $(ComputeScope  $certObject.PsPath)
        Store               = $(ComputeStore  $certObject.PsPath)
        SerialNumber        = $certObject.SerialNumber
        Subject             = $certObject.Subject
        Status              = $(ComputeExpirationStatus $certObject.NotAfter)
        SignatureAlgorithm  = $(ComputeSignatureAlgorithm $certObject)

        Thumbprint          = $certObject.Thumbprint
        Version             = $certObject.Version

        Template            = $(ComputeTemplate $certObject)
    }
}

###############################################################################
# Script execution starts here
###############################################################################

GetChildLeafRecurse $path | ForEach-Object {
    $canExportPrivateKey = $false

    if ($_.HasPrivateKey) {
        [System.Security.Cryptography.CspParameters] $cspParams = new-object System.Security.Cryptography.CspParameters
        $contextField = $_.GetType().GetField("m_safeCertContext", [Reflection.BindingFlags]::NonPublic -bor [Reflection.BindingFlags]::Instance)
        $privateKeyMethod = $_.GetType().GetMethod("GetPrivateKeyInfo", [Reflection.BindingFlags]::NonPublic -bor [Reflection.BindingFlags]::Static)

        if ($contextField -and $privateKeyMethod) {
            $contextValue = $contextField.GetValue($_)
            $privateKeyInfoAvailable = $privateKeyMethod.Invoke($_, @($ContextValue, $cspParams))

            if ($privateKeyInfoAvailable) {
                $csp = new-object System.Security.Cryptography.CspKeyContainerInfo -ArgumentList @($cspParams)

                if ($csp.Exportable) {
                    $canExportPrivateKey = $true
                }
            }
        }
        else {
            $canExportPrivateKey = $true
        }
    }

    ExtractCertInfo $_ $_.HasPrivateKey $canExportPrivateKey
}

}
## [END] Get-WACCMCertificates ##
function Get-WACCMCertificatesForStore {
<#

.SYNOPSIS
Script that enumerates all the certificate scopes/locations in the system.

.DESCRIPTION
Script that enumerates all the certificate scopes/locations in the system.

.ROLE
Readers

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $certificatesStorePath
)

Set-StrictMode -Version 5.0

$treeNodes = @()

$treeNodes = @(Get-ChildItem  $certificatesStorePath | `
    Microsoft.PowerShell.Utility\Select-Object Name, @{name="Path"; expression={$($_.Location.toString() + '\' + $_.Name)}})

$treeNodes

}
## [END] Get-WACCMCertificatesForStore ##
function Get-WACCMTempFolder {
<#

.SYNOPSIS
Script that gets temp folder based on the target node.

.DESCRIPTION
Script that gets temp folder based on the target node.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

Get-Childitem -Path Env:* | where-Object {$_.Name -eq "TEMP"}

}
## [END] Get-WACCMTempFolder ##
function Import-WACCMCertificate {
<#

.SYNOPSIS
Script that imports certificate.

.DESCRIPTION
Script that imports certificate.

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [boolean]
    $fromTaskScheduler,
    [Parameter(Mandatory = $true)]
    [String]
    $storePath,
    [Parameter(Mandatory = $true)]
    [String]
    $filePath,
    [string]
    $exportable,
    [String]
    $password,
    [String]
    $resultFile,
    [String]
    $errorFile
)

BEGIN {
    Set-StrictMode -Version 5.0

    Import-Module -Name Microsoft.PowerShell.Management -ErrorAction SilentlyContinue

    Set-Variable -Name WaitTimeOut -Option ReadOnly -Value 30000 -Scope Script -ErrorAction SilentlyContinue        # 30 second timeout
    Set-Variable -Name LogName -Option ReadOnly -Value "Microsoft-ServerManagementExperience" -Scope Script -ErrorAction SilentlyContinue
    Set-Variable -Name LogSource -Option ReadOnly -Value "SMEScripts" -Scope Script -ErrorAction SilentlyContinue
    Set-Variable -Name ScriptName -Option ReadOnly -Value "Import-Certificate.ps1" -Scope Script -ErrorAction SilentlyContinue
    Set-Variable -Name RsaProviderInstanceName -Option ReadOnly -Value "RSA" -Scope Script -ErrorAction SilentlyContinue
}
PROCESS {
    function importCertificate() {
        param (
            [String]
            $storePath,
            [String]
            $filePath,
            [string]
            $exportable,
            [string]
            $password,
            [String]
            $resultFile,
            [String]
            $errorFile
        )

        try {
            Import-Module PKI
            $params = @{ CertStoreLocation = $storePath; FilePath = $filePath }
            if ($password) {
                Add-Type -AssemblyName System.Security
                $encode = New-Object System.Text.UTF8Encoding
                $encrypted = [System.Convert]::FromBase64String($password)
                $decrypted = [System.Security.Cryptography.ProtectedData]::Unprotect($encrypted, $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
                $newPassword = $encode.GetString($decrypted)
                $securePassword = ConvertTo-SecureString -String $newPassword -Force -AsPlainText
                $params.Password = $securePassword
            }

            if ($exportable -eq "Export") {
                $params.Exportable = $true;
            }

            Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Information `
                -Message "[$ScriptName]: Calling ImportPfx Certificate" -ErrorAction SilentlyContinue

            Import-PfxCertificate @params | ConvertTo-Json | Out-File $ResultFile
        }
        catch {
            $_.Exception.Message | ConvertTo-Json | Out-File $ErrorFile
        }
    }

    function DecryptPasswordWithJWKOnNode($encryptedJWKPassword) {
        if (Get-Variable -Scope Global -Name $RsaProviderInstanceName -ErrorAction SilentlyContinue) {
            $rsaProvider = (Get-Variable -Scope Global -Name $RsaProviderInstanceName).Value
            $decryptedBytes = $rsaProvider.Decrypt([Convert]::FromBase64String($encryptedJWKPassword), [System.Security.Cryptography.RSAEncryptionPadding]::OaepSHA1)

            return [System.Text.Encoding]::UTF8.GetString($decryptedBytes)
        }

        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message "[$ScriptName]: Password decryption failed. RSACryptoServiceProvider Instance not found" -ErrorAction SilentlyContinue

        # TODO: Localize this message!
        throw [System.InvalidOperationException] "Password decryption failed. RSACryptoServiceProvider Instance not found"
    }

    ###########################################################################
    # Script execution starts here
    ###########################################################################

    if ([System.IO.Path]::GetExtension($filePath) -ne ".pfx") {
        Import-Module PKI
        Import-Certificate -CertStoreLocation $storePath -FilePath $filePath
        return;
    }

    $stopwatch = [System.Diagnostics.Stopwatch]::new()

    function isSystemLockdownPolicyEnforced() {
        return [System.Management.Automation.Security.SystemPolicy]::GetSystemLockdownPolicy() -eq [System.Management.Automation.Security.SystemEnforcementMode]::Enforce
    }
    $isWdacEnforced = isSystemLockdownPolicyEnforced;

    #In WDAC environment script file will already be available on the machine
    #In WDAC mode the same script is executed - once normally and once through task Scheduler
    if ($isWdacEnforced) {
        if ($fromTaskScheduler) {
            importCertificate $storePath $filePath $exportable $password $resultFile $errorFile;
            return;
        }
    }
    else {
        #In non-WDAC environment script file will not be available on the machine
        #Hence, a dynamic script is created which is executed through the task Scheduler
        $ScriptFile = $env:temp + "\import-certificate.ps1"
    }

    # PFX private key handlings
    if ($password) {
        $decryptedJWKPassword = DecryptPasswordWithJWKOnNode $password
        # encrypt password with current user.
        Add-Type -AssemblyName System.Security
        $encode = New-Object System.Text.UTF8Encoding
        $bytes = $encode.GetBytes($decryptedJWKPassword)
        $encrypt = [System.Security.Cryptography.ProtectedData]::Protect($bytes, $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
        $userEncryptedPassword = [System.Convert]::ToBase64String($encrypt)
    }

    # Pass parameters to script and generate script file in temp folder
    $resultFile = $env:temp + "\import-certificate_result.json"
    $errorFile = $env:temp + "\import-certificate_error.json"
    if (Test-Path $errorFile) {
        Remove-Item $errorFile
    }

    if (Test-Path $resultFile) {
        Remove-Item $resultFile
    }

    # Create a scheduled task
    $TaskName = "SMEImportCertificate"

    $User = [Security.Principal.WindowsIdentity]::GetCurrent()
    $HashArguments = @{};

    if ($exportable) {
        $HashArguments.Add("exportable", $exportable)
    }
    if ($userEncryptedPassword) {
        $HashArguments.Add("password", $userEncryptedPassword)
    }

    $tempArgs = ""
    foreach ($key in $HashArguments.Keys) {
        $value = $HashArguments[$key]
        $value = """$value"""
        $tempArgs += " -$key $value"
    }

    if ($isWdacEnforced) {
        $arg = "-NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -command ""&{Import-Module Microsoft.SME.CertificateManager; Import-WACCMCertificate -fromTaskScheduler `$true -storePath $storePath $tempArgs -filePath $filePath -resultFile $resultFile -errorFile $errorFile}"""
    }
    else {
    (Get-Command importCertificate).ScriptBlock | Set-Content -Path $ScriptFile
        $arg = "-NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -File $ScriptFile -storePath $storePath $tempArgs -filePath $filePath -resultFile $resultFile -errorFile $errorFile"
    }

    if ($null -eq (New-Object Security.Principal.WindowsPrincipal $User).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
        Write-Warning "To perform some operations you must run an elevated Windows PowerShell console."
    }

    $Scheduler = New-Object -ComObject Schedule.Service

    #Try to connect to schedule service 3 time since it may fail the first time
    for ($i = 1; $i -le 3; $i++) {
        try {
            $Scheduler.Connect()
            Break
        }
        catch {
            if ($i -ge 3) {
                $message = $_.Exception.Message

                Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
                    -Message "[$ScriptName]:Can't connect to Schedule service. Error: $message"  -ErrorAction SilentlyContinue
                Write-Error $message -ErrorAction Stop
            }
            else {
                Start-Sleep -s 1
            }
        }
    }

    $RootFolder = $Scheduler.GetFolder("\")

    #Delete existing task
    if ($RootFolder.GetTasks(0) | Where-Object { $_.Name -eq $TaskName }) {
        Write-Debug("Deleting existing task" + $TaskName)
        $RootFolder.DeleteTask($TaskName, 0)
    }

    $Task = $Scheduler.NewTask(0)
    $RegistrationInfo = $Task.RegistrationInfo
    $RegistrationInfo.Description = $TaskName
    $RegistrationInfo.Author = $User.Name

    $Triggers = $Task.Triggers
    $Trigger = $Triggers.Create(7) #TASK_TRIGGER_REGISTRATION: Starts the task when the task is registered.
    $Trigger.Enabled = $true

    $Settings = $Task.Settings
    $Settings.Enabled = $True
    $Settings.StartWhenAvailable = $True
    $Settings.Hidden = $False

    $Action = $Task.Actions.Create(0)
    $Action.Path = "powershell"
    $Action.Arguments = $arg

    #Tasks will be run with the highest privileges
    $Task.Principal.RunLevel = 1

    # Start the task with SYSTEM creds
    $RootFolder.RegisterTaskDefinition($TaskName, $Task, 6, "SYSTEM", $Null, 1) | Out-Null
    $RootFolder.DeleteTask($TaskName, 0)
    $stopWatch.Start()
    if (!$isWdacEnforced) {
        Remove-Item $ScriptFile
    }

    $now = $stopWatch.Elapsed.Milliseconds
    if ($now -ge $WaitTimeOut) {
        $message = 'Timed out waiting for the the scheduled task that exports the certificate to complete.'

        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message "[$ScriptName]:$message"  -ErrorAction SilentlyContinue

        throw $message
    }

    if (Test-Path $ErrorFile) {
        $result = Get-Content -Raw -Path $ErrorFile | ConvertFrom-Json
        Remove-Item $ErrorFile
        Remove-Item $ResultFile
        throw $result
    }

    #Return result
    if (Test-Path $ResultFile) {
        $result = Get-Content -Raw -Path $ResultFile | ConvertFrom-Json
        Remove-Item $ResultFile
        return $result
    }
}
END {
    Remove-Variable -Name WaitTimeOut -Scope Script -Force
    Remove-Variable -Name LogName -Scope Script -Force
    Remove-Variable -Name LogSource -Scope Script -Force
    Remove-Variable -Name ScriptName -Scope Script -Force
    Remove-Variable -Name RsaProviderInstanceName -Scope Script -Force
}

}
## [END] Import-WACCMCertificate ##
function Remove-WACCMCertificate {
 <#

.SYNOPSIS
Script that deletes certificate.

.DESCRIPTION
Script that deletes certificate.

.ROLE
Administrators

#>

 param (
    [Parameter(Mandatory = $true)]
    [string]
    $thumbprintPath
    )

Set-StrictMode -Version 5.0

Get-ChildItem $thumbprintPath | Remove-Item

}
## [END] Remove-WACCMCertificate ##
function Remove-WACCMItemByPath {
<#

.SYNOPSIS
Script that deletes certificate based on the path.

.DESCRIPTION
Script that deletes certificate based on the path.

.ROLE
Administrators

#>

 Param(
    [Parameter(Mandatory = $true)]
    [string]
    $path
    )

Set-StrictMode -Version 5.0

Remove-Item -Path $path;

}
## [END] Remove-WACCMItemByPath ##
function Update-WACCMCertificate {
<#

.SYNOPSIS
Renew Certificate

.DESCRIPTION
Renew Certificate

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String]
    $CertificatePath,
    [Parameter(Mandatory = $true)]
    [Boolean]
    $UseSameCertificateKey,
    [Parameter(Mandatory = $true)]
    [String]
    $UserName,
    [Parameter(Mandatory = $true)]
    [String]
    $Password
)

BEGIN {
    Set-StrictMode -Version 5.0

    Import-Module -Name Microsoft.PowerShell.Management -ErrorAction SilentlyContinue

    Set-Variable -Name LogName -Option ReadOnly -Value "Microsoft-ServerManagementExperience" -Scope Script -ErrorAction SilentlyContinue
    Set-Variable -Name LogSource -Option ReadOnly -Value "SMEScripts" -Scope Script -ErrorAction SilentlyContinue
    Set-Variable -Name rsaProviderInstanceName -Option ReadOnly -Value "RSA" -Scope Script -ErrorAction SilentlyContinue
    Set-Variable -Name ScriptName -Option ReadOnly -Value "Update-Certificate.ps1" -Scope Script -ErrorAction SilentlyContinue
}
PROCESS {
    function DecryptDataWithJWKOnNode($encryptedData) {
        if (Get-Variable -Scope Global -Name $rsaProviderInstanceName -ErrorAction SilentlyContinue) {
            $rsaProvider = (Get-Variable -Scope Global -Name $rsaProviderInstanceName).Value
            $decryptedBytes = $rsaProvider.Decrypt([Convert]::FromBase64String($encryptedData), [System.Security.Cryptography.RSAEncryptionPadding]::OaepSHA1)

            return [System.Text.Encoding]::UTF8.GetString($decryptedBytes)
        }

        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message "[$ScriptName]: Password decryption failed. RSACryptoServiceProvider Instance not found" -ErrorAction SilentlyContinue

        # TODO: Localize this message!
        throw [System.InvalidOperationException] "Password decryption failed. RSACryptoServiceProvider Instance not found"
    }

    ###############################################################################
    # Script execution starts here...
    ###############################################################################

    # TODO: Figure out if this script really needs alternate credentials!

    $decryptedPassword = DecryptDataWithJWKOnNode $Password
    $securePassword = ConvertTo-SecureString $decryptedPassword -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($UserName, $securePassword)
    $thisComputer = [System.Net.DNS]::GetHostByName('').HostName

    Invoke-Command -ComputerName $thisComputer -ScriptBlock {
        param(
            [string]
            $CertificatePath,
            [boolean]
            $UseSameCertificateKey
        )

        BEGIN {
            Set-StrictMode -Version 5.0

            Import-Module -Name Microsoft.PowerShell.Management -ErrorAction SilentlyContinue

            #https://msdn.microsoft.com/en-us/library/windows/desktop/aa374936(v=vs.85).aspx
            enum EncodingType {
                XCN_CRYPT_STRING_BASE64HEADER = 0
                XCN_CRYPT_STRING_BASE64 = 0x1
                XCN_CRYPT_STRING_BINARY = 0x2
                XCN_CRYPT_STRING_BASE64REQUESTHEADER = 0x3
                XCN_CRYPT_STRING_HEX = 0x4
                XCN_CRYPT_STRING_HEXASCII = 0x5
                XCN_CRYPT_STRING_BASE64_ANY = 0x6
                XCN_CRYPT_STRING_ANY = 0x7
                XCN_CRYPT_STRING_HEX_ANY = 0x8
                XCN_CRYPT_STRING_BASE64X509CRLHEADER = 0x9
                XCN_CRYPT_STRING_HEXADDR = 0xa
                XCN_CRYPT_STRING_HEXASCIIADDR = 0xb
                XCN_CRYPT_STRING_HEXRAW = 0xc
                XCN_CRYPT_STRING_NOCRLF = 0x40000000
                XCN_CRYPT_STRING_NOCR = 0x80000000
            }

            #https://msdn.microsoft.com/en-us/library/windows/desktop/aa379399(v=vs.85).aspx
            enum X509CertificateEnrollmentContext {
                ContextUser = 0x1
                ContextMachine = 0x2
                ContextAdministratorForceMachine = 0x3
            }

            #https://msdn.microsoft.com/en-us/library/windows/desktop/aa379430(v=vs.85).aspx
            enum X509RequestInheritOptions {
                InheritDefault = 0x00000000
                InheritNewDefaultKey = 0x00000001
                InheritNewSimilarKey = 0x00000002
                InheritPrivateKey = 0x00000003
                InheritPublicKey = 0x00000004
                InheritKeyMask = 0x0000000f
                InheritNone = 0x00000010
                InheritRenewalCertificateFlag = 0x00000020
                InheritTemplateFlag = 0x00000040
                InheritSubjectFlag = 0x00000080
                InheritExtensionsFlag = 0x00000100
                InheritSubjectAltNameFlag = 0x00000200
                InheritValidityPeriodFlag = 0x00000400
            }

            Set-Variable -Name LogName -Option ReadOnly -Value "Microsoft-ServerManagementExperience" -Scope Script -ErrorAction SilentlyContinue
            Set-Variable -Name LogSource -Option ReadOnly -Value "SMEScripts" -Scope Script -ErrorAction SilentlyContinue
            Set-Variable -Name ScriptName -Option ReadOnly -Value "Update-Certificate.ps1" -Scope Script -ErrorAction SilentlyContinue
            Set-Variable -Name LocalMachineStoreName -Option ReadOnly -Value "LocalMachine" -Scope Script -ErrorAction SilentlyContinue
            Set-Variable -Name CurrentUserStoreName -Option ReadOnly -Value "CurrentUser" -Scope Script -ErrorAction SilentlyContinue
            Set-Variable -Name ResultPropertyName -Option ReadOnly -Value "Result" -Scope Script -ErrorAction SilentlyContinue
            Set-Variable -Name ErrorPropertyName -Option ReadOnly -Value "Error" -Scope Script -ErrorAction SilentlyContinue
            Set-Variable -Name ErrorMessagePropertyName -Option ReadOnly -Value "ErrorMessage" -Scope Script -ErrorAction SilentlyContinue
            Set-Variable -Name StatusPropertyName -Option ReadOnly -Value "Status" -Scope Script -ErrorAction SilentlyContinue
            Set-Variable -Name StatusError -Option ReadOnly -Value "Error" -Scope Script -ErrorAction SilentlyContinue
            Set-Variable -Name StatusSuccess -Option ReadOnly -Value "Success" -Scope Script -ErrorAction SilentlyContinue
            Set-Variable -Name ErrorNoContext -Option ReadOnly -Value "NoContext" -Scope Script -ErrorAction SilentlyContinue
            Set-Variable -Name ErrorUpdateFailed -Option ReadOnly -Value "UpdateFailed" -Scope Script -ErrorAction SilentlyContinue
        }
        PROCESS {
            function main([String] $path, [Boolean] $useSameKey) {
                $global:result = ""
        
                $cert = Get-Item -Path $path
        
                if ($path -match $LocalMachineStoreName) {
                    $context = [X509CertificateEnrollmentContext]::ContextAdministratorForceMachine
                }
        
                if ($path -match $CurrentUserStoreName) {
                    $context = [X509CertificateEnrollmentContext]::ContextUser
                }
        
                if (!$context) {
                    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
                        -Message "[$ScriptName]:The certificate store name in certificate path $path is not known."  -ErrorAction SilentlyContinue
        
                    $global:result = @{ $StatusPropertyName = $StatusError; $ResultPropertyName = ''; $ErrorMessagePropertyName = ''; $ErrorPropertyName = $ErrorNoContext; }
                    $global:result
        
                    return
                }
        
                $x509RequestInheritOptions = [X509RequestInheritOptions]::InheritTemplateFlag
        
                $x509RequestInheritOptions += [X509RequestInheritOptions]::InheritRenewalCertificateFlag
        
                if ($useSameKey) {
                    $x509RequestInheritOptions += [X509RequestInheritOptions]::InheritPrivateKey
                }
        
                try {
                    $encodingType = [EncodingType]::XCN_CRYPT_STRING_BASE64
        
                    $pkcs10 = New-Object -ComObject X509Enrollment.CX509CertificateRequestPkcs10
                    $pkcs10.Silent = $true
                    $pkcs10.InitializeFromCertificate($context, [System.Convert]::ToBase64String($cert.RawData), $encodingType, $x509RequestInheritOptions)
                    $pkcs10.AlternateSignatureAlgorithm = $false
                    $pkcs10.SmimeCapabilities = $false
                    $pkcs10.SuppressDefaults = $true
                    $pkcs10.Encode()
        
                    #https://msdn.microsoft.com/en-us/library/windows/desktop/aa377809(v=vs.85).aspx
                    $enrolledCert = New-Object -ComObject X509Enrollment.CX509Enrollment
                    $enrolledCert.InitializeFromRequest($pkcs10)
        
                    $enrolledCert.Enroll()
        
                    $cert = New-Object Security.Cryptography.X509Certificates.X509Certificate2
                    $cert.Import([System.Convert]::FromBase64String($enrolledCert.Certificate(1)))
        
                    $global:result = @{ $StatusPropertyName = $StatusSuccess; $ResultPropertyName = $cert.Thumbprint; $ErrorMessagePropertyName = ''; $ErrorPropertyName = ''; }
                } catch [System.Runtime.InteropServices.COMException] {
                    $exceptionMessage = $_.Exception.Message
                    $friendlyName = if ($cert.FriendlyName) { $cert.FriendlyName } else { $cert.Subject }
                    $global:result = @{ $StatusPropertyName = $StatusError; $ResultPropertyName = ''; $ErrorMessagePropertyName = $exceptionMessage; $ErrorPropertyName = $ErrorUpdateFailed; }
                
                    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
                        -Message "[$ScriptName]: Couldn't update certificate $friendlyName.  Error: $exceptionMessage ; Certificate path: $path"  -ErrorAction SilentlyContinue
                }
        
                $global:result
            }

            main $CertificatePath $UseSameCertificateKey
        }
        END {
            Remove-Variable -Name LogName -Scope Script -Force
            Remove-Variable -Name LogSource -Scope Script -Force
            Remove-Variable -Name ScriptName -Scope Script -Force
            Remove-Variable -Name LocalMachineStoreName -Scope Script -Force
            Remove-Variable -Name CurrentUserStoreName -Scope Script -Force
            Remove-Variable -Name ResultPropertyName -Scope Script -Force
            Remove-Variable -Name ErrorPropertyName -Scope Script -Force
            Remove-Variable -Name ErrorMessagePropertyName -Scope Script -Force
            Remove-Variable -Name StatusPropertyName -Scope Script -Force
            Remove-Variable -Name StatusError -Scope Script -Force
            Remove-Variable -Name StatusSuccess -Scope Script -Force
            Remove-Variable -Name ErrorNoContext -Scope Script -Force
            Remove-Variable -Name ErrorUpdateFailed -Scope Script -Force
        }

    } -Credential $credential -ArgumentList $CertificatePath, $UseSameCertificateKey 
}
END {
    Remove-Variable -Name LogName -Scope Script -Force
    Remove-Variable -Name LogSource -Scope Script -Force
    Remove-Variable -Name rsaProviderInstanceName -Scope Script -Force
}

}
## [END] Update-WACCMCertificate ##
function Clear-WACCMEventLogChannel {
<#

.SYNOPSIS
Clear the event log channel specified.

.DESCRIPTION
Clear the event log channel specified.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>
 
Param(
    [string]$channel
)

[System.Diagnostics.Eventing.Reader.EventLogSession]::GlobalSession.ClearLog("$channel") 
}
## [END] Clear-WACCMEventLogChannel ##
function Clear-WACCMEventLogChannelAfterExport {
<#

.SYNOPSIS
Clear the event log channel after export the event log channel file (.evtx).

.DESCRIPTION
Clear the event log channel after export the event log channel file (.evtx).
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>

Param(
    [string]$channel
)

$segments = $channel.Split("-")
$name = $segments[-1]

$randomString = [GUID]::NewGuid().ToString()
$ResultFile = $env:temp + "\" + $name + "_" + $randomString + ".evtx"
$ResultFile = $ResultFile -replace "/", "-"

wevtutil epl "$channel" "$ResultFile" /ow:true

[System.Diagnostics.Eventing.Reader.EventLogSession]::GlobalSession.ClearLog("$channel") 

return $ResultFile

}
## [END] Clear-WACCMEventLogChannelAfterExport ##
function Export-WACCMEventLogChannel {
<#

.SYNOPSIS
Export the event log channel file (.evtx) with filter XML.

.DESCRIPTION
Export the event log channel file (.evtx) with filter XML.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>

Param(
    [string]$channel,
    [string]$filterXml
)

$segments = $channel.Split("-")
$name = $segments[-1]

$randomString = [GUID]::NewGuid().ToString()
$ResultFile = $env:temp + "\" + $name + "_" + $randomString + ".evtx"
$ResultFile = $ResultFile -replace "/", "-"

wevtutil epl "$channel" "$ResultFile" /q:"$filterXml" /ow:true

return $ResultFile

}
## [END] Export-WACCMEventLogChannel ##
function Get-WACCMCimEventLogRecords {
<#

.SYNOPSIS
Get Log records of event channel by using Server Manager CIM provider.

.DESCRIPTION
Get Log records of event channel by using Server Manager CIM provider.

.ROLE
Readers

#>

Param(
    [string]$FilterXml,
    [bool]$ReverseDirection
)

import-module CimCmdlets

$machineName = [System.Net.DNS]::GetHostByName('').HostName
Invoke-CimMethod -Namespace root/Microsoft/Windows/ServerManager -ClassName MSFT_ServerManagerTasks -MethodName GetServerEventDetailEx -Arguments @{FilterXml = $FilterXml; ReverseDirection = $ReverseDirection; } |
    ForEach-Object {
        $result = $_
        if ($result.PSObject.Properties.Match('ItemValue').Count) {
            foreach ($item in $result.ItemValue) {
                @{
                    ItemValue = 
                    @{
                        Description  = $item.description
                        Id           = $item.id
                        Level        = $item.level
                        Log          = $item.log
                        Source       = $item.source
                        Timestamp    = $item.timestamp
                        __ServerName = $machineName
                    }
                }
            }
        }
    }

}
## [END] Get-WACCMCimEventLogRecords ##
function Get-WACCMCimWin32LogicalDisk {
<#

.SYNOPSIS
Gets Win32_LogicalDisk object.

.DESCRIPTION
Gets Win32_LogicalDisk object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_LogicalDisk

}
## [END] Get-WACCMCimWin32LogicalDisk ##
function Get-WACCMCimWin32NetworkAdapter {
<#

.SYNOPSIS
Gets Win32_NetworkAdapter object.

.DESCRIPTION
Gets Win32_NetworkAdapter object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_NetworkAdapter

}
## [END] Get-WACCMCimWin32NetworkAdapter ##
function Get-WACCMCimWin32PhysicalMemory {
<#

.SYNOPSIS
Gets Win32_PhysicalMemory object.

.DESCRIPTION
Gets Win32_PhysicalMemory object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_PhysicalMemory

}
## [END] Get-WACCMCimWin32PhysicalMemory ##
function Get-WACCMCimWin32Processor {
<#

.SYNOPSIS
Gets Win32_Processor object.

.DESCRIPTION
Gets Win32_Processor object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_Processor

}
## [END] Get-WACCMCimWin32Processor ##
function Get-WACCMClusterEvents {
<#
.SYNOPSIS
Gets CIM instance

.DESCRIPTION
Gets CIM instance

.ROLE
Readers

#>

param (
		[Parameter(Mandatory = $true)]
		[string]
    $namespace,

    [Parameter(Mandatory = $true)]
		[string]
    $className

)
Import-Module CimCmdlets
Get-CimInstance -Namespace  $namespace -ClassName $className

}
## [END] Get-WACCMClusterEvents ##
function Get-WACCMClusterInventory {
<#

.SYNOPSIS
Retrieves the inventory data for a cluster.

.DESCRIPTION
Retrieves the inventory data for a cluster.

.ROLE
Readers

#>

Import-Module CimCmdlets -ErrorAction SilentlyContinue

# JEA code requires to pre-import the module (this is slow on failover cluster environment.)
Import-Module FailoverClusters -ErrorAction SilentlyContinue

Import-Module Storage -ErrorAction SilentlyContinue
<#

.SYNOPSIS
Get the name of this computer.

.DESCRIPTION
Get the best available name for this computer.  The FQDN is preferred, but when not avaialble
the NetBIOS name will be used instead.

#>

function getComputerName() {
    $computerSystem = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object Name, DNSHostName

    if ($computerSystem) {
        $computerName = $computerSystem.DNSHostName

        if ($null -eq $computerName) {
            $computerName = $computerSystem.Name
        }

        return $computerName
    }

    return $null
}

<#

.SYNOPSIS
Are the cluster PowerShell cmdlets installed on this server?

.DESCRIPTION
Are the cluster PowerShell cmdlets installed on this server?

#>

function getIsClusterCmdletAvailable() {
    $cmdlet = Get-Command "Get-Cluster" -ErrorAction SilentlyContinue

    return !!$cmdlet
}

<#

.SYNOPSIS
Get the MSCluster Cluster CIM instance from this server.

.DESCRIPTION
Get the MSCluster Cluster CIM instance from this server.

#>
function getClusterCimInstance() {
    $namespace = Get-CimInstance -Namespace root/MSCluster -ClassName __NAMESPACE -ErrorAction SilentlyContinue

    if ($namespace) {
        return Get-CimInstance -Namespace root/mscluster MSCluster_Cluster -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object fqdn, S2DEnabled
    }

    return $null
}


<#

.SYNOPSIS
Determines if the current cluster supports Failover Clusters Time Series Database.

.DESCRIPTION
Use the existance of the path value of cmdlet Get-StorageHealthSetting to determine if TSDB
is supported or not.

#>
function getClusterPerformanceHistoryPath() {
    $storageSubsystem = Get-StorageSubSystem clus* -ErrorAction SilentlyContinue
    $storageHealthSettings = Get-StorageHealthSetting -InputObject $storageSubsystem -Name "System.PerformanceHistory.Path" -ErrorAction SilentlyContinue

    return $null -ne $storageHealthSettings
}

<#

.SYNOPSIS
Get some basic information about the cluster from the cluster.

.DESCRIPTION
Get the needed cluster properties from the cluster.

#>
function getClusterInfo() {
    $returnValues = @{}

    $returnValues.Fqdn = $null
    $returnValues.isS2DEnabled = $false
    $returnValues.isTsdbEnabled = $false

    $cluster = getClusterCimInstance
    if ($cluster) {
        $returnValues.Fqdn = $cluster.fqdn
        $isS2dEnabled = !!(Get-Member -InputObject $cluster -Name "S2DEnabled") -and ($cluster.S2DEnabled -eq 1)
        $returnValues.isS2DEnabled = $isS2dEnabled

        if ($isS2DEnabled) {
            $returnValues.isTsdbEnabled = getClusterPerformanceHistoryPath
        } else {
            $returnValues.isTsdbEnabled = $false
        }
    }

    return $returnValues
}

<#

.SYNOPSIS
Are the cluster PowerShell Health cmdlets installed on this server?

.DESCRIPTION
Are the cluster PowerShell Health cmdlets installed on this server?

s#>
function getisClusterHealthCmdletAvailable() {
    $cmdlet = Get-Command -Name "Get-HealthFault" -ErrorAction SilentlyContinue

    return !!$cmdlet
}
<#

.SYNOPSIS
Are the Britannica (sddc management resources) available on the cluster?

.DESCRIPTION
Are the Britannica (sddc management resources) available on the cluster?

#>
function getIsBritannicaEnabled() {
    return $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_Cluster -ErrorAction SilentlyContinue)
}

<#

.SYNOPSIS
Are the Britannica (sddc management resources) virtual machine available on the cluster?

.DESCRIPTION
Are the Britannica (sddc management resources) virtual machine available on the cluster?

#>
function getIsBritannicaVirtualMachineEnabled() {
    return $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_VirtualMachine -ErrorAction SilentlyContinue)
}

<#

.SYNOPSIS
Are the Britannica (sddc management resources) virtual switch available on the cluster?

.DESCRIPTION
Are the Britannica (sddc management resources) virtual switch available on the cluster?

#>
function getIsBritannicaVirtualSwitchEnabled() {
    return $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_VirtualSwitch -ErrorAction SilentlyContinue)
}

###########################################################################
# main()
###########################################################################

$clusterInfo = getClusterInfo

$result = New-Object PSObject

$result | Add-Member -MemberType NoteProperty -Name 'Fqdn' -Value $clusterInfo.Fqdn
$result | Add-Member -MemberType NoteProperty -Name 'IsS2DEnabled' -Value $clusterInfo.isS2DEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsTsdbEnabled' -Value $clusterInfo.isTsdbEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsClusterHealthCmdletAvailable' -Value (getIsClusterHealthCmdletAvailable)
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaEnabled' -Value (getIsBritannicaEnabled)
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaVirtualMachineEnabled' -Value (getIsBritannicaVirtualMachineEnabled)
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaVirtualSwitchEnabled' -Value (getIsBritannicaVirtualSwitchEnabled)
$result | Add-Member -MemberType NoteProperty -Name 'IsClusterCmdletAvailable' -Value (getIsClusterCmdletAvailable)
$result | Add-Member -MemberType NoteProperty -Name 'CurrentClusterNode' -Value (getComputerName)

$result

}
## [END] Get-WACCMClusterInventory ##
function Get-WACCMClusterNodes {
<#

.SYNOPSIS
Retrieves the inventory data for cluster nodes in a particular cluster.

.DESCRIPTION
Retrieves the inventory data for cluster nodes in a particular cluster.

.ROLE
Readers

#>

import-module CimCmdlets

# JEA code requires to pre-import the module (this is slow on failover cluster environment.)
import-module FailoverClusters -ErrorAction SilentlyContinue

###############################################################################
# Constants
###############################################################################

Set-Variable -Name LogName -Option Constant -Value "Microsoft-ServerManagementExperience" -ErrorAction SilentlyContinue
Set-Variable -Name LogSource -Option Constant -Value "SMEScripts" -ErrorAction SilentlyContinue
Set-Variable -Name ScriptName -Option Constant -Value $MyInvocation.ScriptName -ErrorAction SilentlyContinue

<#

.SYNOPSIS
Are the cluster PowerShell cmdlets installed?

.DESCRIPTION
Use the Get-Command cmdlet to quickly test if the cluster PowerShell cmdlets
are installed on this server.

#>

function getClusterPowerShellSupport() {
    $cmdletInfo = Get-Command 'Get-ClusterNode' -ErrorAction SilentlyContinue

    return $cmdletInfo -and $cmdletInfo.Name -eq "Get-ClusterNode"
}

<#

.SYNOPSIS
Get the cluster nodes using the cluster CIM provider.

.DESCRIPTION
When the cluster PowerShell cmdlets are not available fallback to using
the cluster CIM provider to get the needed information.

#>

function getClusterNodeCimInstances() {
    # Change the WMI property NodeDrainStatus to DrainStatus to match the PS cmdlet output.
    return Get-CimInstance -Namespace root/mscluster MSCluster_Node -ErrorAction SilentlyContinue | `
        Microsoft.PowerShell.Utility\Select-Object @{Name="DrainStatus"; Expression={$_.NodeDrainStatus}}, DynamicWeight, Name, NodeWeight, FaultDomain, State
}

<#

.SYNOPSIS
Get the cluster nodes using the cluster PowerShell cmdlets.

.DESCRIPTION
When the cluster PowerShell cmdlets are available use this preferred function.

#>

function getClusterNodePsInstances() {
    return Get-ClusterNode -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object DrainStatus, DynamicWeight, Name, NodeWeight, FaultDomain, State
}

<#

.SYNOPSIS
Use DNS services to get the FQDN of the cluster NetBIOS name.

.DESCRIPTION
Use DNS services to get the FQDN of the cluster NetBIOS name.

.Notes
It is encouraged that the caller add their approprate -ErrorAction when
calling this function.

#>

function getClusterNodeFqdn([string]$clusterNodeName) {
    return ([System.Net.Dns]::GetHostEntry($clusterNodeName)).HostName
}

<#

.SYNOPSIS
Writes message to event log as warning.

.DESCRIPTION
Writes message to event log as warning.

#>

function writeToEventLog([string]$message) {
    Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Warning `
        -Message $message  -ErrorAction SilentlyContinue
}

<#

.SYNOPSIS
Get the cluster nodes.

.DESCRIPTION
When the cluster PowerShell cmdlets are available get the information about the cluster nodes
using PowerShell.  When the cmdlets are not available use the Cluster CIM provider.

#>

function getClusterNodes() {
    $isClusterCmdletAvailable = getClusterPowerShellSupport

    if ($isClusterCmdletAvailable) {
        $clusterNodes = getClusterNodePsInstances
    } else {
        $clusterNodes = getClusterNodeCimInstances
    }

    $clusterNodeMap = @{}

    foreach ($clusterNode in $clusterNodes) {
        $clusterNodeName = $clusterNode.Name.ToLower()
        try 
        {
            $clusterNodeFqdn = getClusterNodeFqdn $clusterNodeName -ErrorAction SilentlyContinue
        }
        catch 
        {
            $clusterNodeFqdn = $clusterNodeName
            writeToEventLog "[$ScriptName]: The fqdn for node '$clusterNodeName' could not be obtained. Defaulting to machine name '$clusterNodeName'"
        }

        $clusterNodeResult = New-Object PSObject

        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'FullyQualifiedDomainName' -Value $clusterNodeFqdn
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'Name' -Value $clusterNodeName
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'DynamicWeight' -Value $clusterNode.DynamicWeight
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'NodeWeight' -Value $clusterNode.NodeWeight
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'FaultDomain' -Value $clusterNode.FaultDomain
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'State' -Value $clusterNode.State
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'DrainStatus' -Value $clusterNode.DrainStatus

        $clusterNodeMap.Add($clusterNodeName, $clusterNodeResult)
    }

    return $clusterNodeMap
}

###########################################################################
# main()
###########################################################################

getClusterNodes

}
## [END] Get-WACCMClusterNodes ##
function Get-WACCMDecryptedDataFromNode {
<#

.SYNOPSIS
Gets data after decrypting it on a node.

.DESCRIPTION
Decrypts data on node using a cached RSAProvider used during encryption within 3 minutes of encryption and returns the decrypted data.
This script should be imported or copied directly to other scripts, do not send the returned data as an argument to other scripts.

.PARAMETER encryptedData
Encrypted data to be decrypted (String).

.ROLE
Readers

#>
param (
  [Parameter(Mandatory = $true)]
  [String]
  $encryptedData
)

Set-StrictMode -Version 5.0

New-Variable -Name rsaProviderInstanceName -Value "RSA" -Option Constant

function DecryptDataWithJWKOnNode {
  if(Get-Variable -Scope Global -Name $rsaProviderInstanceName -EA SilentlyContinue) {
    $rsaProvider = (Get-Variable -Scope Global -Name $rsaProviderInstanceName).Value
    $decryptedBytes = $rsaProvider.Decrypt([Convert]::FromBase64String($encryptedData), [System.Security.Cryptography.RSAEncryptionPadding]::OaepSHA1)
    return [System.Text.Encoding]::UTF8.GetString($decryptedBytes)
  }
  # If you copy this script directly to another, you can get rid of the throw statement and add custom error handling logic such as "Write-Error"
  throw [System.InvalidOperationException] "Password decryption failed. RSACryptoServiceProvider Instance not found"
}

}
## [END] Get-WACCMDecryptedDataFromNode ##
function Get-WACCMEncryptionJWKOnNode {
<#

.SYNOPSIS
Gets encrytion JSON web key from node.

.DESCRIPTION
Gets encrytion JSON web key from node.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

New-Variable -Name rsaProviderInstanceName -Value "RSA" -Option Constant

function Get-RSAProvider
{
    if(Get-Variable -Scope Global -Name $rsaProviderInstanceName -EA SilentlyContinue)
    {
        return (Get-Variable -Scope Global -Name $rsaProviderInstanceName).Value
    }

    $Global:RSA = New-Object System.Security.Cryptography.RSACryptoServiceProvider -ArgumentList 4096
    return $RSA
}

function Get-JsonWebKey
{
    $rsaProvider = Get-RSAProvider
    $parameters = $rsaProvider.ExportParameters($false)
    return [PSCustomObject]@{
        kty = 'RSA'
        alg = 'RSA-OAEP'
        e = [Convert]::ToBase64String($parameters.Exponent)
        n = [Convert]::ToBase64String($parameters.Modulus).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    }
}

$jwk = Get-JsonWebKey
ConvertTo-Json $jwk -Compress

}
## [END] Get-WACCMEncryptionJWKOnNode ##
function Get-WACCMEventLogDisplayName {
<#

.SYNOPSIS
Get the EventLog log name and display name by using Get-EventLog cmdlet.

.DESCRIPTION
Get the EventLog log name and display name by using Get-EventLog cmdlet.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>


return (Get-EventLog -LogName * | Microsoft.PowerShell.Utility\Select-Object Log,LogDisplayName)
}
## [END] Get-WACCMEventLogDisplayName ##
function Get-WACCMEventLogFilteredCount {
<#

.SYNOPSIS
Get the total amout of events that meet the filters selected by using Get-WinEvent cmdlet.

.DESCRIPTION
Get the total amout of events that meet the filters selected by using Get-WinEvent cmdlet.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>

Param(
    [string]$filterXml
)

return (Get-WinEvent -FilterXml "$filterXml" -ErrorAction 'SilentlyContinue').count
}
## [END] Get-WACCMEventLogFilteredCount ##
function Get-WACCMEventLogRecords {
<#

.SYNOPSIS
Get Log records of event channel by using Get-WinEvent cmdlet.

.DESCRIPTION
Get Log records of event channel by using Get-WinEvent cmdlet.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers
#>

Param(
    [string]
    $filterXml,
    [bool]
    $reverseDirection
)

$ErrorActionPreference = 'SilentlyContinue'
Import-Module Microsoft.PowerShell.Diagnostics;

#
# Prepare parameters for command Get-WinEvent
#
$winEventscmdParams = @{
    FilterXml = $filterXml;
    Oldest    = !$reverseDirection;
}

Get-WinEvent  @winEventscmdParams -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object recordId,
id, 
@{Name = "Log"; Expression = {$_."logname"}}, 
level, 
timeCreated, 
machineName, 
@{Name = "Source"; Expression = {$_."ProviderName"}}, 
@{Name = "Description"; Expression = {$_."Message"}}



}
## [END] Get-WACCMEventLogRecords ##
function Get-WACCMEventLogSummary {
<#

.SYNOPSIS
Get the log summary (Name, Total) for the channel selected by using Get-WinEvent cmdlet.

.DESCRIPTION
Get the log summary (Name, Total) for the channel selected by using Get-WinEvent cmdlet.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>

Param(
    [string]$channel
)

Import-Module Microsoft.PowerShell.Diagnostics

$channelList = $channel.split(",")

Get-WinEvent -ListLog $channelList -Force -ErrorAction SilentlyContinue |`
    Microsoft.PowerShell.Utility\Select-Object LogName, IsEnabled, RecordCount, IsClassicLog, LogType, OwningProviderName
}
## [END] Get-WACCMEventLogSummary ##
function Get-WACCMServerInventory {
<#

.SYNOPSIS
Retrieves the inventory data for a server.

.DESCRIPTION
Retrieves the inventory data for a server.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

Import-Module CimCmdlets

Import-Module Storage -ErrorAction SilentlyContinue

<#

.SYNOPSIS
Converts an arbitrary version string into just 'Major.Minor'

.DESCRIPTION
To make OS version comparisons we only want to compare the major and
minor version.  Build number and/os CSD are not interesting.

#>

function convertOsVersion([string]$osVersion) {
  [Ref]$parsedVersion = $null
  if (![Version]::TryParse($osVersion, $parsedVersion)) {
    return $null
  }

  $version = [Version]$parsedVersion.Value
  return New-Object Version -ArgumentList $version.Major, $version.Minor
}

<#

.SYNOPSIS
Determines if CredSSP is enabled for the current server or client.

.DESCRIPTION
Check the registry value for the CredSSP enabled state.

#>

function isCredSSPEnabled() {
  Set-Variable credSSPServicePath -Option Constant -Value "WSMan:\localhost\Service\Auth\CredSSP"
  Set-Variable credSSPClientPath -Option Constant -Value "WSMan:\localhost\Client\Auth\CredSSP"

  $credSSPServerEnabled = $false;
  $credSSPClientEnabled = $false;

  $credSSPServerService = Get-Item $credSSPServicePath -ErrorAction SilentlyContinue
  if ($credSSPServerService) {
    $credSSPServerEnabled = [System.Convert]::ToBoolean($credSSPServerService.Value)
  }

  $credSSPClientService = Get-Item $credSSPClientPath -ErrorAction SilentlyContinue
  if ($credSSPClientService) {
    $credSSPClientEnabled = [System.Convert]::ToBoolean($credSSPClientService.Value)
  }

  return ($credSSPServerEnabled -or $credSSPClientEnabled)
}

<#

.SYNOPSIS
Determines if the Hyper-V role is installed for the current server or client.

.DESCRIPTION
The Hyper-V role is installed when the VMMS service is available.  This is much
faster then checking Get-WindowsFeature and works on Windows Client SKUs.

#>

function isHyperVRoleInstalled() {
  $vmmsService = Get-Service -Name "VMMS" -ErrorAction SilentlyContinue

  return $vmmsService -and $vmmsService.Name -eq "VMMS"
}

<#

.SYNOPSIS
Determines if the Hyper-V PowerShell support module is installed for the current server or client.

.DESCRIPTION
The Hyper-V PowerShell support module is installed when the modules cmdlets are available.  This is much
faster then checking Get-WindowsFeature and works on Windows Client SKUs.

#>
function isHyperVPowerShellSupportInstalled() {
  # quicker way to find the module existence. it doesn't load the module.
  return !!(Get-Module -ListAvailable Hyper-V -ErrorAction SilentlyContinue)
}

<#

.SYNOPSIS
Determines if Windows Management Framework (WMF) 5.0, or higher, is installed for the current server or client.

.DESCRIPTION
Windows Admin Center requires WMF 5 so check the registey for WMF version on Windows versions that are less than
Windows Server 2016.

#>
function isWMF5Installed([string] $operatingSystemVersion) {
  Set-Variable Server2016 -Option Constant -Value (New-Object Version '10.0')   # And Windows 10 client SKUs
  Set-Variable Server2012 -Option Constant -Value (New-Object Version '6.2')

  $version = convertOsVersion $operatingSystemVersion
  if (-not $version) {
    # Since the OS version string is not properly formatted we cannot know the true installed state.
    return $false
  }

  if ($version -ge $Server2016) {
    # It's okay to assume that 2016 and up comes with WMF 5 or higher installed
    return $true
  }
  else {
    if ($version -ge $Server2012) {
      # Windows 2012/2012R2 are supported as long as WMF 5 or higher is installed
      $registryKey = 'HKLM:\SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine'
      $registryKeyValue = Get-ItemProperty -Path $registryKey -Name PowerShellVersion -ErrorAction SilentlyContinue

      if ($registryKeyValue -and ($registryKeyValue.PowerShellVersion.Length -ne 0)) {
        $installedWmfVersion = [Version]$registryKeyValue.PowerShellVersion

        if ($installedWmfVersion -ge [Version]'5.0') {
          return $true
        }
      }
    }
  }

  return $false
}

<#

.SYNOPSIS
Determines if the current usser is a system administrator of the current server or client.

.DESCRIPTION
Determines if the current usser is a system administrator of the current server or client.

#>
function isUserAnAdministrator() {
  return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}

<#

.SYNOPSIS
Get some basic information about the Failover Cluster that is running on this server.

.DESCRIPTION
Create a basic inventory of the Failover Cluster that may be running in this server.

#>
function getClusterInformation() {
  $returnValues = @{ }

  $returnValues.IsS2dEnabled = $false
  $returnValues.IsCluster = $false
  $returnValues.ClusterFqdn = $null
  $returnValues.IsBritannicaEnabled = $false

  $namespace = Get-CimInstance -Namespace root/MSCluster -ClassName __NAMESPACE -ErrorAction SilentlyContinue
  if ($namespace) {
    $cluster = Get-CimInstance -Namespace root/MSCluster -ClassName MSCluster_Cluster -ErrorAction SilentlyContinue
    if ($cluster) {
      $returnValues.IsCluster = $true
      $returnValues.ClusterFqdn = $cluster.Fqdn
      $returnValues.IsS2dEnabled = !!(Get-Member -InputObject $cluster -Name "S2DEnabled") -and ($cluster.S2DEnabled -gt 0)
      $returnValues.IsBritannicaEnabled = $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_Cluster -ErrorAction SilentlyContinue)
    }
  }

  return $returnValues
}

<#

.SYNOPSIS
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the passed in computer name.

.DESCRIPTION
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the passed in computer name.

#>
function getComputerFqdnAndAddress($computerName) {
  $hostEntry = [System.Net.Dns]::GetHostEntry($computerName)
  $addressList = @()
  foreach ($item in $hostEntry.AddressList) {
    $address = New-Object PSObject
    $address | Add-Member -MemberType NoteProperty -Name 'IpAddress' -Value $item.ToString()
    $address | Add-Member -MemberType NoteProperty -Name 'AddressFamily' -Value $item.AddressFamily.ToString()
    $addressList += $address
  }

  $result = New-Object PSObject
  $result | Add-Member -MemberType NoteProperty -Name 'Fqdn' -Value $hostEntry.HostName
  $result | Add-Member -MemberType NoteProperty -Name 'AddressList' -Value $addressList
  return $result
}

<#

.SYNOPSIS
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the current server or client.

.DESCRIPTION
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the current server or client.

#>
function getHostFqdnAndAddress($computerSystem) {
  $computerName = $computerSystem.DNSHostName
  if (!$computerName) {
    $computerName = $computerSystem.Name
  }

  return getComputerFqdnAndAddress $computerName
}

<#

.SYNOPSIS
Are the needed management CIM interfaces available on the current server or client.

.DESCRIPTION
Check for the presence of the required server management CIM interfaces.

#>
function getManagementToolsSupportInformation() {
  $returnValues = @{ }

  $returnValues.ManagementToolsAvailable = $false
  $returnValues.ServerManagerAvailable = $false

  $namespaces = Get-CimInstance -Namespace root/microsoft/windows -ClassName __NAMESPACE -ErrorAction SilentlyContinue

  if ($namespaces) {
    $returnValues.ManagementToolsAvailable = !!($namespaces | Where-Object { $_.Name -ieq "ManagementTools" })
    $returnValues.ServerManagerAvailable = !!($namespaces | Where-Object { $_.Name -ieq "ServerManager" })
  }

  return $returnValues
}

<#

.SYNOPSIS
Check the remote app enabled or not.

.DESCRIPTION
Check the remote app enabled or not.

#>
function isRemoteAppEnabled() {
  Set-Variable key -Option Constant -Value "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Terminal Server\\TSAppAllowList"

  $registryKeyValue = Get-ItemProperty -Path $key -Name fDisabledAllowList -ErrorAction SilentlyContinue

  if (-not $registryKeyValue) {
    return $false
  }
  return $registryKeyValue.fDisabledAllowList -eq 1
}

<#

.SYNOPSIS
Check the remote app enabled or not.

.DESCRIPTION
Check the remote app enabled or not.

#>

<#
c
.SYNOPSIS
Get the Win32_OperatingSystem information as well as current version information from the registry

.DESCRIPTION
Get the Win32_OperatingSystem instance and filter the results to just the required properties.
This filtering will make the response payload much smaller. Included in the results are current version
information from the registry

#>
function getOperatingSystemInfo() {
  $operatingSystemInfo = Get-CimInstance Win32_OperatingSystem | Microsoft.PowerShell.Utility\Select-Object csName, Caption, OperatingSystemSKU, Version, ProductType, OSType, LastBootUpTime, SerialNumber
  $currentVersion = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" | Microsoft.PowerShell.Utility\Select-Object CurrentBuild, UBR, DisplayVersion

  $operatingSystemInfo | Add-Member -MemberType NoteProperty -Name CurrentBuild -Value $currentVersion.CurrentBuild
  $operatingSystemInfo | Add-Member -MemberType NoteProperty -Name UpdateBuildRevision -Value $currentVersion.UBR
  $operatingSystemInfo | Add-Member -MemberType NoteProperty -Name DisplayVersion -Value $currentVersion.DisplayVersion

  return $operatingSystemInfo
}

<#

.SYNOPSIS
Get the Win32_ComputerSystem information

.DESCRIPTION
Get the Win32_ComputerSystem instance and filter the results to just the required properties.
This filtering will make the response payload much smaller.

#>
function getComputerSystemInfo() {
  return Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue | `
    Microsoft.PowerShell.Utility\Select-Object TotalPhysicalMemory, DomainRole, Manufacturer, Model, NumberOfLogicalProcessors, Domain, Workgroup, DNSHostName, Name, PartOfDomain, SystemFamily, SystemSKUNumber
}

<#

.SYNOPSIS
Get SMBIOS locally from the passed in machineName


.DESCRIPTION
Get SMBIOS locally from the passed in machine name

#>
function getSmbiosData($computerSystem) {
  <#
    Array of chassis types.
    The following list of ChassisTypes is copied from the latest DMTF SMBIOS specification.
    REF: https://www.dmtf.org/sites/default/files/standards/documents/DSP0134_3.1.1.pdf
  #>
  $ChassisTypes =
  @{
    1  = 'Other'
    2  = 'Unknown'
    3  = 'Desktop'
    4  = 'Low Profile Desktop'
    5  = 'Pizza Box'
    6  = 'Mini Tower'
    7  = 'Tower'
    8  = 'Portable'
    9  = 'Laptop'
    10 = 'Notebook'
    11 = 'Hand Held'
    12 = 'Docking Station'
    13 = 'All in One'
    14 = 'Sub Notebook'
    15 = 'Space-Saving'
    16 = 'Lunch Box'
    17 = 'Main System Chassis'
    18 = 'Expansion Chassis'
    19 = 'SubChassis'
    20 = 'Bus Expansion Chassis'
    21 = 'Peripheral Chassis'
    22 = 'Storage Chassis'
    23 = 'Rack Mount Chassis'
    24 = 'Sealed-Case PC'
    25 = 'Multi-system chassis'
    26 = 'Compact PCI'
    27 = 'Advanced TCA'
    28 = 'Blade'
    29 = 'Blade Enclosure'
    30 = 'Tablet'
    31 = 'Convertible'
    32 = 'Detachable'
    33 = 'IoT Gateway'
    34 = 'Embedded PC'
    35 = 'Mini PC'
    36 = 'Stick PC'
  }

  $list = New-Object System.Collections.ArrayList
  $win32_Bios = Get-CimInstance -class Win32_Bios
  $obj = New-Object -Type PSObject | Microsoft.PowerShell.Utility\Select-Object SerialNumber, Manufacturer, UUID, BaseBoardProduct, ChassisTypes, Chassis, SystemFamily, SystemSKUNumber, SMBIOSAssetTag
  $obj.SerialNumber = $win32_Bios.SerialNumber
  $obj.Manufacturer = $win32_Bios.Manufacturer
  $computerSystemProduct = Get-CimInstance Win32_ComputerSystemProduct
  if ($null -ne $computerSystemProduct) {
    $obj.UUID = $computerSystemProduct.UUID
  }
  $baseboard = Get-CimInstance Win32_BaseBoard
  if ($null -ne $baseboard) {
    $obj.BaseBoardProduct = $baseboard.Product
  }
  $systemEnclosure = Get-CimInstance Win32_SystemEnclosure
  if ($null -ne $systemEnclosure) {
    $obj.SMBIOSAssetTag = $systemEnclosure.SMBIOSAssetTag
  }
  $obj.ChassisTypes = Get-CimInstance Win32_SystemEnclosure | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty ChassisTypes
  $obj.Chassis = New-Object -TypeName 'System.Collections.ArrayList'
  $obj.ChassisTypes | ForEach-Object -Process {
    $obj.Chassis.Add($ChassisTypes[[int]$_])
  }
  $obj.SystemFamily = $computerSystem.SystemFamily
  $obj.SystemSKUNumber = $computerSystem.SystemSKUNumber
  $list.Add($obj) | Out-Null

  return $list

}

<#

.SYNOPSIS
Get the azure arc status information

.DESCRIPTION
Get the azure arc status information

#>
function getAzureArcStatus() {

  $LogName = "Microsoft-ServerManagementExperience"
  $LogSource = "SMEScript"
  $ScriptName = "Get-ServerInventory.ps1 - getAzureArcStatus()"

  Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

  Get-Service -Name himds -ErrorVariable Err -ErrorAction SilentlyContinue | Out-Null

  if (!!$Err) {

    $Err = "The Azure arc agent is not installed. Details: $Err"

    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Information `
    -Message "[$ScriptName]: $Err"  -ErrorAction SilentlyContinue

    $status = "NotInstalled"
  }
  else {
    $status = (azcmagent show --json | ConvertFrom-Json -ErrorAction Stop).status
  }

  return $status
}

<#

.SYNOPSIS
Gets an EnforcementMode that describes the system lockdown policy on this computer.

.DESCRIPTION
By checking the system lockdown policy, we can infer if PowerShell is in ConstrainedLanguage mode as a result of an enforced WDAC policy.
Note: $ExecutionContext.SessionState.LanguageMode should not be used within a trusted (by the WDAC policy) script context for this purpose because
the language mode returned would potentially not reflect the system-wide lockdown policy/language mode outside of the execution context.

#>
function getSystemLockdownPolicy() {
  return [System.Management.Automation.Security.SystemPolicy]::GetSystemLockdownPolicy().ToString()
}

<#

.SYNOPSIS
Determines if the operating system is HCI.

.DESCRIPTION
Using the operating system 'Caption' (which corresponds to the 'ProductName' registry key at HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion) to determine if a server OS is HCI.

#>
function isServerOsHCI([string] $operatingSystemCaption) {
  return $operatingSystemCaption -eq "Microsoft Azure Stack HCI"
}

###########################################################################
# main()
###########################################################################

$operatingSystem = getOperatingSystemInfo
$computerSystem = getComputerSystemInfo
$isAdministrator = isUserAnAdministrator
$fqdnAndAddress = getHostFqdnAndAddress $computerSystem
$hostname = [Environment]::MachineName
$netbios = $env:ComputerName
$managementToolsInformation = getManagementToolsSupportInformation
$isWmfInstalled = isWMF5Installed $operatingSystem.Version
$clusterInformation = getClusterInformation -ErrorAction SilentlyContinue
$isHyperVPowershellInstalled = isHyperVPowerShellSupportInstalled
$isHyperVRoleInstalled = isHyperVRoleInstalled
$isCredSSPEnabled = isCredSSPEnabled
$isRemoteAppEnabled = isRemoteAppEnabled
$smbiosData = getSmbiosData $computerSystem
$azureArcStatus = getAzureArcStatus
$systemLockdownPolicy = getSystemLockdownPolicy
$isHciServer = isServerOsHCI $operatingSystem.Caption

$result = New-Object PSObject
$result | Add-Member -MemberType NoteProperty -Name 'IsAdministrator' -Value $isAdministrator
$result | Add-Member -MemberType NoteProperty -Name 'OperatingSystem' -Value $operatingSystem
$result | Add-Member -MemberType NoteProperty -Name 'ComputerSystem' -Value $computerSystem
$result | Add-Member -MemberType NoteProperty -Name 'Fqdn' -Value $fqdnAndAddress.Fqdn
$result | Add-Member -MemberType NoteProperty -Name 'AddressList' -Value $fqdnAndAddress.AddressList
$result | Add-Member -MemberType NoteProperty -Name 'Hostname' -Value $hostname
$result | Add-Member -MemberType NoteProperty -Name 'NetBios' -Value $netbios
$result | Add-Member -MemberType NoteProperty -Name 'IsManagementToolsAvailable' -Value $managementToolsInformation.ManagementToolsAvailable
$result | Add-Member -MemberType NoteProperty -Name 'IsServerManagerAvailable' -Value $managementToolsInformation.ServerManagerAvailable
$result | Add-Member -MemberType NoteProperty -Name 'IsWmfInstalled' -Value $isWmfInstalled
$result | Add-Member -MemberType NoteProperty -Name 'IsCluster' -Value $clusterInformation.IsCluster
$result | Add-Member -MemberType NoteProperty -Name 'ClusterFqdn' -Value $clusterInformation.ClusterFqdn
$result | Add-Member -MemberType NoteProperty -Name 'IsS2dEnabled' -Value $clusterInformation.IsS2dEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaEnabled' -Value $clusterInformation.IsBritannicaEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsHyperVRoleInstalled' -Value $isHyperVRoleInstalled
$result | Add-Member -MemberType NoteProperty -Name 'IsHyperVPowershellInstalled' -Value $isHyperVPowershellInstalled
$result | Add-Member -MemberType NoteProperty -Name 'IsCredSSPEnabled' -Value $isCredSSPEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsRemoteAppEnabled' -Value $isRemoteAppEnabled
$result | Add-Member -MemberType NoteProperty -Name 'SmbiosData' -Value $smbiosData
$result | Add-Member -MemberType NoteProperty -Name 'AzureArcStatus' -Value $azureArcStatus
$result | Add-Member -MemberType NoteProperty -Name 'SystemLockdownPolicy' -Value $systemLockdownPolicy
$result | Add-Member -MemberType NoteProperty -Name 'IsHciServer' -Value $isHciServer

$result

}
## [END] Get-WACCMServerInventory ##
function Set-WACCMEventLogChannelStatus {
 <#

.SYNOPSIS
 Change the current status (Enabled/Disabled) for the selected channel.

.DESCRIPTION
Change the current status (Enabled/Disabled) for the selected channel.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>

Param(
    [string]$channel,
    [boolean]$status
)

$ch = Get-WinEvent -ListLog $channel
$ch.set_IsEnabled($status)
$ch.SaveChanges()
}
## [END] Set-WACCMEventLogChannelStatus ##

# SIG # Begin signature block
# MIInvgYJKoZIhvcNAQcCoIInrzCCJ6sCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBc7Y439qObEIi6
# Wuir0RhXNx6I6ky1VYMZ9xHqVbU3fKCCDXYwggX0MIID3KADAgECAhMzAAADTrU8
# esGEb+srAAAAAANOMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMwMzE2MTg0MzI5WhcNMjQwMzE0MTg0MzI5WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDdCKiNI6IBFWuvJUmf6WdOJqZmIwYs5G7AJD5UbcL6tsC+EBPDbr36pFGo1bsU
# p53nRyFYnncoMg8FK0d8jLlw0lgexDDr7gicf2zOBFWqfv/nSLwzJFNP5W03DF/1
# 1oZ12rSFqGlm+O46cRjTDFBpMRCZZGddZlRBjivby0eI1VgTD1TvAdfBYQe82fhm
# WQkYR/lWmAK+vW/1+bO7jHaxXTNCxLIBW07F8PBjUcwFxxyfbe2mHB4h1L4U0Ofa
# +HX/aREQ7SqYZz59sXM2ySOfvYyIjnqSO80NGBaz5DvzIG88J0+BNhOu2jl6Dfcq
# jYQs1H/PMSQIK6E7lXDXSpXzAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUnMc7Zn/ukKBsBiWkwdNfsN5pdwAw
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMDUxNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAD21v9pHoLdBSNlFAjmk
# mx4XxOZAPsVxxXbDyQv1+kGDe9XpgBnT1lXnx7JDpFMKBwAyIwdInmvhK9pGBa31
# TyeL3p7R2s0L8SABPPRJHAEk4NHpBXxHjm4TKjezAbSqqbgsy10Y7KApy+9UrKa2
# kGmsuASsk95PVm5vem7OmTs42vm0BJUU+JPQLg8Y/sdj3TtSfLYYZAaJwTAIgi7d
# hzn5hatLo7Dhz+4T+MrFd+6LUa2U3zr97QwzDthx+RP9/RZnur4inzSQsG5DCVIM
# pA1l2NWEA3KAca0tI2l6hQNYsaKL1kefdfHCrPxEry8onJjyGGv9YKoLv6AOO7Oh
# JEmbQlz/xksYG2N/JSOJ+QqYpGTEuYFYVWain7He6jgb41JbpOGKDdE/b+V2q/gX
# UgFe2gdwTpCDsvh8SMRoq1/BNXcr7iTAU38Vgr83iVtPYmFhZOVM0ULp/kKTVoir
# IpP2KCxT4OekOctt8grYnhJ16QMjmMv5o53hjNFXOxigkQWYzUO+6w50g0FAeFa8
# 5ugCCB6lXEk21FFB1FdIHpjSQf+LP/W2OV/HfhC3uTPgKbRtXo83TZYEudooyZ/A
# Vu08sibZ3MkGOJORLERNwKm2G7oqdOv4Qj8Z0JrGgMzj46NFKAxkLSpE5oHQYP1H
# tPx1lPfD7iNSbJsP6LiUHXH1MIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGZ4wghmaAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAANOtTx6wYRv6ysAAAAAA04wDQYJYIZIAWUDBAIB
# BQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEICKkziFT3yLMxv7uhI/inG0w
# 8+MDtTFZDv+Wx6myGjgNMEIGCisGAQQBgjcCAQwxNDAyoBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEB
# BQAEggEASrYFpmqc8Cila9corChlOq7Z5FSd5yD719+LhcagGcVxvpy1D7mG1QHl
# DQmtKhk0M11hVos5fn/+LyOZQKsblelsZuIRPztnU+BJN2QB/gXOXAO4hf+3EKP3
# laUNPkQc8t1DJD4L5Lt4UUjO/stH3dRANTI2GW09Rq90FBzWA5WZ9d7lVb1JFBtG
# 3UJKnhYLd/Si0rEw9CdfVBPu1X93dBdcwvVQ001fIrkgXZWYPoVrddJEobPD/ME7
# WnsQUjlHHXvO2XQzdYB6EqBcO2NZs2JhlU7M43kK2Bdo5a92e0MKXc4oacktSPqi
# ofrxXG1DzfyoBFelm/vIH+Y0mQQtm6GCFygwghckBgorBgEEAYI3AwMBMYIXFDCC
# FxAGCSqGSIb3DQEHAqCCFwEwghb9AgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFYBgsq
# hkiG9w0BCRABBKCCAUcEggFDMIIBPwIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFl
# AwQCAQUABCCdGTOh4EhqobFGuPj9Ruuy/lhxbMarTr6yTPveh+7iQwIGZWdIIW3r
# GBIyMDIzMTIwNzAzNDQ0My40MVowBIACAfSggdikgdUwgdIxCzAJBgNVBAYTAlVT
# MRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQK
# ExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVs
# YW5kIE9wZXJhdGlvbnMgTGltaXRlZDEmMCQGA1UECxMdVGhhbGVzIFRTUyBFU046
# RkM0MS00QkQ0LUQyMjAxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNl
# cnZpY2WgghF4MIIHJzCCBQ+gAwIBAgITMwAAAeKZmZXx3OMg6wABAAAB4jANBgkq
# hkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0yMzEw
# MTIxOTA3MjVaFw0yNTAxMTAxOTA3MjVaMIHSMQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVy
# YXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOkZDNDEtNEJE
# NC1EMjIwMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNlMIIC
# IjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAtWO1mFX6QWZvxwpCmDabOKwO
# VEj3vwZvZqYa9sCYJ3TglUZ5N79AbMzwptCswOiXsMLuNLTcmRys+xaL1alXCwhy
# RFDwCRfWJ0Eb0eHIKykBq9+6/PnmSGXtus9DHsf31QluwTfAyamYlqw9amAXTnNm
# W+lZANQsNwhjKXmVcjgdVnk3oxLFY7zPBaviv3GQyZRezsgLEMmvlrf1JJ48AlEj
# LOdohzRbNnowVxNHMss3I8ETgqtW/UsV33oU3EDPCd61J4+DzwSZF7OvZPcdMUSW
# d4lfJBh3phDt4IhzvKWVahjTcISD2CGiun2pQpwFR8VxLhcSV/cZIRGeXMmwruz9
# kY9Th1odPaNYahiFrZAI6aSCM6YEUKpAUXAWaw+tmPh5CzNjGrhzgeo+dS7iFPhq
# qm9Rneog5dt3JTjak0v3dyfSs9NOV45Sw5BuC+VF22EUIF6nF9vqduynd9xlo8F9
# Nu1dVryctC4wIGrJ+x5u6qdvCP6UdB+oqmK+nJ3soJYAKiPvxdTBirLUfJidK1OZ
# 7hP28rq7Y78pOF9E54keJKDjjKYWP7fghwUSE+iBoq802xNWbhBuqmELKSevAHKq
# isEIsfpuWVG0kwnCa7sZF1NCwjHYcwqqmES2lKbXPe58BJ0+uA+GxAhEWQdka6KE
# vUmOPgu7cJsCaFrSU6sCAwEAAaOCAUkwggFFMB0GA1UdDgQWBBREhA4R2r7tB2yW
# m0mIJE2leAnaBTAfBgNVHSMEGDAWgBSfpxVdAF5iXYP05dJlpxtTNRnpcjBfBgNV
# HR8EWDBWMFSgUqBQhk5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2Ny
# bC9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcmwwbAYI
# KwYBBQUHAQEEYDBeMFwGCCsGAQUFBzAChlBodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NlcnRzL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAy
# MDEwKDEpLmNydDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMI
# MA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEA5FREMatVFNue6V+y
# DZxOzLKHthe+FVTs1kyQhMBBiwUQ9WC9K+ILKWvlqneRrvpjPS3/qXG5zMjrDu1e
# ryfhbFRSByPnACGc2iuGcPyWNiptyTft+CBgrf7ATAuE/U8YLm29crTFiiZTWdT6
# Vc7L1lGdKEj8dl0WvDayuC2xtajD04y4ANLmWDuiStdrZ1oI4afG5oPUg77rkTuq
# /Y7RbSwaPsBZ06M12l7E+uykvYoRw4x4lWaST87SBqeEXPMcCdaO01ad5TXVZDoH
# G/w6k3V9j3DNCiLJyC844kz3eh3nkQZ5fF8Xxuh8tWVQTfMiKShJ537yzrU0M/7H
# 1EzJrabAr9izXF28OVlMed0gqyx+a7e+79r4EV/a4ijJxVO8FCm/92tEkPrx6jjT
# WaQJEWSbL/4GZCVGvHatqmoC7mTQ16/6JR0FQqZf+I5opnvm+5CDuEKIEDnEiblk
# hcNKVfjvDAVqvf8GBPCe0yr2trpBEB5L+j+5haSa+q8TwCrfxCYqBOIGdZJL+5U9
# xocTICufIWHkb6p4IaYvjgx8ScUSHFzexo+ZeF7oyFKAIgYlRkMDvffqdAPx+fjL
# rnfgt6X4u5PkXlsW3SYvB34fkbEbM5tmab9zekRa0e/W6Dt1L8N+tx3WyfYTiCTh
# bUvWN1EFsr3HCQybBj4Idl4xK8EwggdxMIIFWaADAgECAhMzAAAAFcXna54Cm0mZ
# AAAAAAAVMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0
# ZSBBdXRob3JpdHkgMjAxMDAeFw0yMTA5MzAxODIyMjVaFw0zMDA5MzAxODMyMjVa
# MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMT
# HU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMIICIjANBgkqhkiG9w0BAQEF
# AAOCAg8AMIICCgKCAgEA5OGmTOe0ciELeaLL1yR5vQ7VgtP97pwHB9KpbE51yMo1
# V/YBf2xK4OK9uT4XYDP/XE/HZveVU3Fa4n5KWv64NmeFRiMMtY0Tz3cywBAY6GB9
# alKDRLemjkZrBxTzxXb1hlDcwUTIcVxRMTegCjhuje3XD9gmU3w5YQJ6xKr9cmmv
# Haus9ja+NSZk2pg7uhp7M62AW36MEBydUv626GIl3GoPz130/o5Tz9bshVZN7928
# jaTjkY+yOSxRnOlwaQ3KNi1wjjHINSi947SHJMPgyY9+tVSP3PoFVZhtaDuaRr3t
# pK56KTesy+uDRedGbsoy1cCGMFxPLOJiss254o2I5JasAUq7vnGpF1tnYN74kpEe
# HT39IM9zfUGaRnXNxF803RKJ1v2lIH1+/NmeRd+2ci/bfV+AutuqfjbsNkz2K26o
# ElHovwUDo9Fzpk03dJQcNIIP8BDyt0cY7afomXw/TNuvXsLz1dhzPUNOwTM5TI4C
# vEJoLhDqhFFG4tG9ahhaYQFzymeiXtcodgLiMxhy16cg8ML6EgrXY28MyTZki1ug
# poMhXV8wdJGUlNi5UPkLiWHzNgY1GIRH29wb0f2y1BzFa/ZcUlFdEtsluq9QBXps
# xREdcu+N+VLEhReTwDwV2xo3xwgVGD94q0W29R6HXtqPnhZyacaue7e3PmriLq0C
# AwEAAaOCAd0wggHZMBIGCSsGAQQBgjcVAQQFAgMBAAEwIwYJKwYBBAGCNxUCBBYE
# FCqnUv5kxJq+gpE8RjUpzxD/LwTuMB0GA1UdDgQWBBSfpxVdAF5iXYP05dJlpxtT
# NRnpcjBcBgNVHSAEVTBTMFEGDCsGAQQBgjdMg30BATBBMD8GCCsGAQUFBwIBFjNo
# dHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9yeS5o
# dG0wEwYDVR0lBAwwCgYIKwYBBQUHAwgwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBD
# AEEwCwYDVR0PBAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU1fZW
# y4/oolxiaNE9lJBb186aGMQwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5t
# aWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIwMTAt
# MDYtMjMuY3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0y
# My5jcnQwDQYJKoZIhvcNAQELBQADggIBAJ1VffwqreEsH2cBMSRb4Z5yS/ypb+pc
# FLY+TkdkeLEGk5c9MTO1OdfCcTY/2mRsfNB1OW27DzHkwo/7bNGhlBgi7ulmZzpT
# Td2YurYeeNg2LpypglYAA7AFvonoaeC6Ce5732pvvinLbtg/SHUB2RjebYIM9W0j
# VOR4U3UkV7ndn/OOPcbzaN9l9qRWqveVtihVJ9AkvUCgvxm2EhIRXT0n4ECWOKz3
# +SmJw7wXsFSFQrP8DJ6LGYnn8AtqgcKBGUIZUnWKNsIdw2FzLixre24/LAl4FOmR
# sqlb30mjdAy87JGA0j3mSj5mO0+7hvoyGtmW9I/2kQH2zsZ0/fZMcm8Qq3UwxTSw
# ethQ/gpY3UA8x1RtnWN0SCyxTkctwRQEcb9k+SS+c23Kjgm9swFXSVRk2XPXfx5b
# RAGOWhmRaw2fpCjcZxkoJLo4S5pu+yFUa2pFEUep8beuyOiJXk+d0tBMdrVXVAmx
# aQFEfnyhYWxz/gq77EFmPWn9y8FBSX5+k77L+DvktxW/tM4+pTFRhLy/AsGConsX
# HRWJjXD+57XQKBqJC4822rpM+Zv/Cuk0+CQ1ZyvgDbjmjJnW4SLq8CdCPSWU5nR0
# W2rRnj7tfqAxM328y+l7vzhwRNGQ8cirOoo6CGJ/2XBjU02N7oJtpQUQwXEGahC0
# HVUzWLOhcGbyoYIC1DCCAj0CAQEwggEAoYHYpIHVMIHSMQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFu
# ZCBPcGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOkZD
# NDEtNEJENC1EMjIwMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2
# aWNloiMKAQEwBwYFKw4DAhoDFQAWm5lp+nRuekl0iF+IHV3ylOiGb6CBgzCBgKR+
# MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMT
# HU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBBQUAAgUA
# 6RupSTAiGA8yMDIzMTIwNzEwMTUzN1oYDzIwMjMxMjA4MTAxNTM3WjB0MDoGCisG
# AQQBhFkKBAExLDAqMAoCBQDpG6lJAgEAMAcCAQACAhBcMAcCAQACAhPrMAoCBQDp
# HPrJAgEAMDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMH
# oSChCjAIAgEAAgMBhqAwDQYJKoZIhvcNAQEFBQADgYEAMe3+6BCtTC6QpBCrwuna
# eLERlFRJOY/CqJf9THf8VPagYhhr/1k2guhpcRfzRk5L5LaLAU51iR7IKKunQkxt
# jXgq76lebJbs7HPmkJ99UKogzYny8GvjfMJQbDlyx9bW6FzAEhBz0lmhu9yIoe3l
# kLOomY7ArUiLSVS9UtH8d4IxggQNMIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1T
# dGFtcCBQQ0EgMjAxMAITMwAAAeKZmZXx3OMg6wABAAAB4jANBglghkgBZQMEAgEF
# AKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEi
# BCCTHQTwO5ClGvQXSCVxh+6xffjgDbFaxGVf1nvpFPeG8DCB+gYLKoZIhvcNAQkQ
# Ai8xgeowgecwgeQwgb0EICuJKkoQ/Sa4xsFQRM4Ogvh3ktToj9uO5whmQ4kIj3//
# MIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAO
# BgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEm
# MCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAHimZmV
# 8dzjIOsAAQAAAeIwIgQg1DaaqNKTqy+3UJpUIW7xt+ppnaS+9inVeqdmuIZp6oQw
# DQYJKoZIhvcNAQELBQAEggIAicpid41IUSaJKDE0egy7myRzdeDWJXD/Q3AW1Zeg
# Etu9WTeHZfyIi7l5l7T7UJ+uG3oB5vi6jLczyo3TVOpSWtNKZ5chejCSZJoxvRuv
# oSGrqReoyFZRX8L64HvETakqip+f9DQsrkOGSVquCpPfNz/mDFwUPaETr15L4lKn
# KSbJTRMOHUr2bM4z6KrzwL+OVCtdiraHD9c6W2h/0juFNXDm9oecLUtqruGCbey+
# Pq7cE/T4oNrCjAYeTjoUVf0DhQMqbb70er4zxF3CSTDDbfJwOQ8hRQffYe40fbKu
# T1Akf3tR0ZcRKAy7Dw3rP5fBrLy7DeFnZQaCpBYzXLdsTk6ovKy/dsomlrgWi2VQ
# hPPaarcvslUWCLIUcR97qht6G0kDXBybtkdhBL5nPyi/TDEDR+Ykj9R0NIO6Ckor
# Sl4tzDNBR+GvtIlBZ31KvaTBuOp8w43Jk1uWRD4eaXqU9XhYDe82nMqOI4dPsghW
# YYRTW3pTw0aqNZifn1fJnMeUZwAZdrbrIVhdOdVIA8hQbj0ePKXL2EpsYmYr00H5
# cYg8dv+wLjxSdL6gH7bQ1Q+h5Ec/yiqFqh7+D/iWrB5yStoWXyHkK3Dxn2DzDoY5
# c4OioRscK8QYppB6WBT4cdccMfq0d63l/BoMa5mnVxh0Qc/5BXc4c4qjXoOBN+U2
# xkQ=
# SIG # End signature block

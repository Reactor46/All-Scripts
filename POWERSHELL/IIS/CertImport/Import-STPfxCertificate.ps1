#requires -version 2
function Import-STPfxCertificate
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Alias('Cn')] [string[]] $ComputerName,
        [Parameter(Mandatory=$true)][string] $CertFilePath,
        [string] $CertRootStore = 'localmachine',
        [string] $CertStore = 'My',
        [string] $X509Flags = 'PersistKeySet,MachineKeySet',
        [System.Security.SecureString] $Password = $null)
    $ErrorActionPreference = 'Continue'
    $TempCertFileName = 'TempC1B_Cert.pfx'
    if ($Password -eq $null)
    {
        $Password = Read-Host -Prompt 'Enter PFX cert password' -AsSecureString
    }
    foreach ($Computer in $ComputerName)
    {
        $Destination = "\\$Computer\admin$\$TempCertFileName"
        try
        {
            Copy-Item -LiteralPath $CertFilePath -Destination $Destination -ErrorAction Stop
        }
        catch
        {
            Write-Error -Message "${Computer}: Unable to copy '$CertFilePath' to '$Destination'. Aborting further processing of this computer."
            continue
        }
        Invoke-Command -ComputerName $Computer -ScriptBlock {
            param(
                [string] $CertFileName,
                [string] $CertRootStore,
                [string] $CertStore,
                [string] $X509Flags,
                $PfxPass)
            $CertPath = "$Env:SystemRoot\$CertFileName"
            $Pfx = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
            # Flags to send in are documented here: https://msdn.microsoft.com/en-us/library/system.security.cryptography.x509certificates.x509keystorageflags%28v=vs.110%29.aspx
            $Pfx.Import($CertPath, $PfxPass, $X509Flags) #"Exportable,PersistKeySet")
            $Store = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Store -ArgumentList $CertStore, $CertRootStore
            $Store.Open("MaxAllowed")
            $Store.Add($Pfx)
            if ($?)
            {
                "${Env:ComputerName}: Successfully added certificate."
            }
            else
            {
                "${Env:ComputerName}: Failed to add certificate! $($Error[0].ToString() -replace '[\r\n]+', ' ')"
            }
            $Store.Close()
            Remove-Item -LiteralPath $CertPath
        } -ArgumentList $TempCertFileName, $CertRootStore, $CertStore, $X509Flags, $Password
    }
}

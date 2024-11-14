[CmdletBinding()] 
Param 
( 
    # Name of the server, defaults to local 
    [Parameter(Mandatory=$false, 
                ValueFromPipelineByPropertyName=$true, 
                Position=0)] 
    [string]$ComputerName=$env:COMPUTERNAME, 
    [int]$returnStateOK = 0, 
    [int]$returnStateWarning = 1, 
    [int]$returnStateCritical = 2, 
    [int]$returnStateUnknown = 3, 
    [int]$WarningDays = 60, 
    [int]$CriticalDays = 30, 
    [string]$CertificatePath = 'Cert:\LocalMachine\My', 
    [string[]]$ExcludedThumbprint#=@('DFE816240B40151BBCD7529D4C55627A8CE1671C') 
) 
 
Begin 
{ 
} 
Process 
{ 
     
    # Get the certificates from the specified computer. 
    try { 
        # Use local path if it is localhost 
        if ($ComputerName -eq $env:COMPUTERNAME) { 
            $Certificates = Get-ChildItem -Path $CertificatePath -ErrorAction Stop -Exclude $ExcludedThumbprint 
            } 
        # Use PSRP if computer is not localhost 
        else { 
            $Certificates = Invoke-Command -ComputerName $ComputerName -ScriptBlock {param($CertificatePath,$ExcludedThumbprint) Get-ChildItem -Path $CertificatePath -Exclude $ExcludedThumbprint} -ArgumentList $CertificatePath,$ExcludedThumbprint -ErrorAction Stop 
            } 
        } 
    # Catch all exceptions 
    catch { 
        Write-Output "Unable to get certificates from $ComputerName.|" ; exit $returnStateUnknown 
        } 
     
    # Filter warning and critical certificates. 
    $WarningCertificates = $Certificates | Where-Object -FilterScript {$_.NotAfter -le (Get-Date).AddDays($WarningDays) -and $_.NotAfter -gt (Get-Date).AddDays($CriticalDays)} | Select Subject, NotAfter, @{Label="Days";Expression={($_.NotAfter - (Get-Date)).Days}} 
    $CriticalCertificates = $Certificates | Where-Object -FilterScript {$_.NotAfter -le (Get-Date).AddDays($CriticalDays)} | Select Subject, NotAfter, @{Label="Days";Expression={($_.NotAfter - (Get-Date)).Days}} 
 
    # If we have either warning or critical certificates, generate list and output status code. 
    if ($WarningCertificates -or $CriticalCertificates) { 
        # If we have critical AND warning certificates, generate list and output status code. 
        if ($CriticalCertificates -and $WarningCertificates) { 
            $CertificatesMessage = "Critical Certificates:`n" 
            foreach ($CriticalCertificate in $CriticalCertificates) { 
                $CertificatesMessage += "$($CriticalCertificate.Subject.Split(',')[0]) expires $($CriticalCertificate.NotAfter) $($CriticalCertificate.Days) days.`n" 
                } 
            $CertificatesMessage += "Warning Certificates:`n" 
            foreach ($WarningCertificate in $WarningCertificates) { 
                $CertificatesMessage += "$($WarningCertificate.Subject.Split(',')[0]) expires $($WarningCertificate.NotAfter) $($WarningCertificate.Days) days.`n" 
                } 
            Write-Output "$CertificatesMessage|" ; exit $returnStateCritical 
            } 
        # If we have only critical certificates. 
        elseif ($CriticalCertificates) { 
            $CertificatesMessage = "Critical Certificates:`n" 
            foreach ($CriticalCertificate in $CriticalCertificates) { 
                $CertificatesMessage += "$($CriticalCertificate.Subject.Split(',')[0]) expires $($CriticalCertificate.NotAfter) $($CriticalCertificate.Days) days.`n" 
                } 
            Write-Output "$CertificatesMessage|" ; exit $returnStateCritical 
            } 
        # If we have only warning certificates.   
        elseif ($WarningCertificates) { 
            $CertificatesMessage = "Warning Certificates:`n" 
            foreach ($WarningCertificate in $WarningCertificates) { 
                $CertificatesMessage += "$($WarningCertificate.Subject.Split(',')[0]) expires $($WarningCertificate.NotAfter) $($WarningCertificate.Days) days.`n" 
                } 
            Write-Output "$CertificatesMessage|" ; exit $returnStateWarning 
            } 
        else {} 
        } 
    else { 
        # No problems found 
        Write-Output "Certificates OK.|" ; exit $returnStateOK 
        } 
} 
End 
{ 
}
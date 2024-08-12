function hex ([string]$delimiter = '') {
    # Helper function to convert bytes into hex strings for display purposes
    # Often used to display WWN
    Begin { $hex = ''}
    Process { $hex += $_.toString('X2') + $delimiter }
    End { $hex.trimEnd($delimiter) }
}

function Add-MyHost2Vipr {
<#
.DESCRIPTION
  If run from a Windows Host, this will add the host run from and its initiators.

.EXAMPLE
  Add-MyHost2Vipr

  Run with no parameters, will get host name and HBA WWNs from the host this is run on, then add it all to ViPR Controller.
#>

    Write-Host "Adding host $env:COMPUTERNAME to Vipr"
    $myHostId = Get-ViprObjectByName -ObjectName $env:COMPUTERNAME -ObjectType host
    if ($myHostId) {
        write-host "Host $env:COMPUTERNAME already exists in Vipr"

    }
    else {
        $hostSpec = @{
            type = 'Windows'
            host_name = $env:COMPUTERNAME
            name = $env:COMPUTERNAME
            discoverable = $false
        }
        $myHostId = Invoke-ViprCall "/tenants/$script:tenantId/hosts" -message $hostSpec -method POST | Wait-ForViprTask
    }
    $wwns = Get-WmiObject -Namespace root\wmi -Class MSFC_FibrePortHBAAttributes -ErrorAction Ignore | ForEach-Object { $_.Attributes.Portwwn | Hex : }
    if ($wwns) {
        foreach ($wwn in $wwns) {
            Write-Host "Adding initiator $wwn to Vipr"
            $wwnId = Get-ViprObjectByName -ObjectName $wwn -ObjectType initiator
            if (!$wwnId) {
                $initiatorSpec = @{
                    protocol = 'FC'
                    initiator_port = $wwn
                    initiator_node = $wwn
                }
                $null = Invoke-ViprCall "/compute/hosts/$myHostId/initiators" -message $initiatorSpec -method POST
            }
            else {
                # wwn aleady exists in Vipr; get its details
                $initiator = Invoke-ViprCall /compute/initiators/$wwnId
                if ($initiator.hostname -ne $env:COMPUTERNAME) {
                    Write-Host "Wwn $wwn is already registered to another host: $($initiator.hostname) and cannot be used"
                }
                elseif ($initiator.'registration_status' -ne 'registered') {
                    Write-Host "Wwn $wwn status is showing UNREGISTERED in Vipr and cannot be used"
                }
                else { Write-Host "Wwn $wwn is already present and registered in Vipr for this host" }
            }
        }
    } else { 'No HBAs found; skipping initiator add to Vipr' }
}
[CmdletBinding(DefaultParameterSetName = "MoveUser")]
param(
    # The mailbox to move (if moving mailboxes as opposed to MDBs)
    [Parameter(Mandatory = $true, ParameterSetName = "MoveUser", ValueFromPipeline = $true, Position = 0)]
    $identity,

    # The MDB to move (if moving MDBs as opposed to mailboxes)
    [Parameter(Mandatory = $true, ParameterSetName = "MoveDatabase")]
    $mailboxDatabase,

    # An optional map of source databases to target databases 
    [Parameter(Mandatory = $false, ParameterSetName = "MoveUser")]
    [Parameter(Mandatory = $false, ParameterSetName = "MoveDatabase")]
    [hashtable]$DatabaseMap,

    # An optinal target database (can't be used with -DatabaseMap)
    [Parameter(Mandatory = $false, ParameterSetName = "MoveUser")]
    [Parameter(Mandatory = $false, ParameterSetName = "MoveDatabase")]
    $TargetDatabase,
    
    # Batch size to start with (will start moves when this many users are found to the same target database)
    [Parameter(Mandatory = $false, ParameterSetName = "MoveUser")]
    [Parameter(Mandatory = $false, ParameterSetName = "MoveDatabase")]
    [int]$startBatchSize = 10,

    # BadItemLimit
    [Parameter(Mandatory = $false, ParameterSetName = "MoveUser")]
    [Parameter(Mandatory = $false, ParameterSetName = "MoveDatabase")]
    [int]$BadItemLimit = 0,

    # Do not wait for moves to be completed
    [Parameter(Mandatory = $false, ParameterSetName = "MoveUser")]
    [Parameter(Mandatory = $false, ParameterSetName = "MoveDatabase")]
    [switch]$SuspendWhenReadyToComplete,

    # When waiting for moves to complete, scan the move status this often (in seconds)
    [int]$pollInterval = 10,
    
    # Domain controller to query
    [Parameter(Mandatory = $false, ParameterSetName = "MoveUser")]
    [Parameter(Mandatory = $false, ParameterSetName = "MoveDatabase")]
    $DomainController
)

## Define common functions and script variables
begin
{
    #load hashtable of localized string
    Import-LocalizedData -BindingVariable MoveMailbox_LocalizedStrings -FileName MoveMailbox.strings.psd1

    # Map of each target DB and their move handlers to provide batching
    $script:databaseHandlerMap = @{}
    
    # flag to track if mailboxes were passed on the pipeline
    [bool]$script:shouldStartMoves = $false

    ## Creates a custom object to serve as bucket for moves targetted to the same MDB 
    ## (can handle provisioning with the special databaseName of ""
    function _CreateDatabaseMoveHandler([string]$databaseName)
    {
        $handler = new-object PSObject
        Add-Member -InputObject $handler -MemberType NoteProperty -Name Name -Value $databaseName
        Add-Member -InputObject $handler -MemberType NoteProperty -Name NewAlias -Value @()
        Add-Member -InputObject $handler -MemberType NoteProperty -Name Started -Value @()
        Add-Member -InputObject $handler -MemberType ScriptMethod -Name Start -Value {
            [array]$movesToStart = @()
            $movesToStart += $this.NewAlias

            Write-Verbose ($MoveMailbox_LocalizedStrings.res_0000 -f $this.NewAlias.Count)

            Write-Progress -Activity ($MoveMailbox_LocalizedStrings.progress_0000 -f $this.Name) -Status ($MoveMailbox_LocalizedStrings.progress_0001 -f $this.NewAlias.Count) -PercentComplete 0 -Id 100
            
            if ("$($this.Name)" -eq "")
            {
                Write-Verbose ($MoveMailbox_LocalizedStrings.res_0001 -f $BadItemLimit)
                # random database - some of the moves may fail starting this way because the database picked by the provisioning layer is the same as the source db
                # need to handle these conditions later :)
                if ($DomainController)
                {
                    [array]$moveRequests=$movesToStart | new-MoveRequest -BadItemLimit $BadItemLimit -SuspendWhenReadyToComplete:$SuspendWhenReadyToComplete -DomainController $DomainController
                }
                else
                {
                    [array]$moveRequests=$movesToStart | new-MoveRequest -BadItemLimit $BadItemLimit -SuspendWhenReadyToComplete:$SuspendWhenReadyToComplete
                }
            }
            else
            {
                Write-Verbose ($MoveMailbox_LocalizedStrings.res_0002 -f $this.Name,$BadItemLimit)

                if ($DomainController)
                {
                    [array]$moveRequests=$movesToStart | new-MoveRequest -TargetDatabase $this.Name -BadItemLimit $BadItemLimit -SuspendWhenReadyToComplete:$SuspendWhenReadyToComplete -DomainController $DomainController
                }
                else
                {
                    [array]$moveRequests=$movesToStart | new-MoveRequest -TargetDatabase $this.Name -BadItemLimit $BadItemLimit -SuspendWhenReadyToComplete:$SuspendWhenReadyToComplete
                }
            }
           
            Write-Verbose $MoveMailbox_LocalizedStrings.res_0003
            # just double check that all moves started (see comment on random database selection :)
            Write-Progress -Activity ($MoveMailbox_LocalizedStrings.progress_0002 -f $this.Name) -Status $MoveMailbox_LocalizedStrings.progress_0003 -PercentComplete 50 -Id 100
            $this.NewAlias = @()
            $startedDNs = $moveRequests | ?{ $_ -ne $null} | %{ $_.DistinguishedName }
            [array]$movesThatFailedToStart = $movesToStart | ?{ $startedDNs -notcontains $_.DistinguishedName }
            if ($movesThatFailedToStart -ne $null -and $movesThatFailedToStart.Count -gt 0)
            {
                $movesToStart | ?{ $movesThatFailedToStart -contains $_ } | %{
                    Write-Warning ($MoveMailbox_LocalizedStrings.res_0004 -f $_)
                }
            }
                        
            Write-Verbose $MoveMailbox_LocalizedStrings.res_0005
            # Just organize the house, get ready for the next batch
            Write-Progress -Activity ($MoveMailbox_LocalizedStrings.progress_0004 -f $this.Name) -Status $MoveMailbox_LocalizedStrings.progress_0005 -PercentComplete 90 -Id 100
            $this.Started += $movesToStart
            Write-Progress -Activity ($MoveMailbox_LocalizedStrings.progress_0006 -f $this.Name) -Status $MoveMailbox_LocalizedStrings.progress_0007 -Completed -Id 100
        }
        
        # put the handler in the pipeline so the caller may consume it
        write-output $handler
    }
    
    ## This is a convenience function for getting the correct move handler based on a *source* database
    # it'll first convert the source db to a target (based on the parameters), and then get the handler for that target
    function _AcquireMoveHandler([string]$database)
    {
        Write-Verbose ($MoveMailbox_LocalizedStrings.res_0006 -f $database)
        # get the correct target DB based on the source
        if ($DatabaseMap)
        {
            Write-Verbose ($MoveMailbox_LocalizedStrings.res_0007 -f $DatabaseMap,$database)
            $realDatabase = $DatabaseMap[$database]
            if ($realDatabase -eq $null)
            {
                throw ($MoveMailbox_LocalizedStrings.resthrow_0000 -f $database)
            }
            write-verbose ($MoveMailbox_LocalizedStrings.res_0023 -f $realDatabase)
        }
        elseif ($TargetDatabase)
        {
            $realDatabase = $TargetDatabase
            write-verbose ($MoveMailbox_LocalizedStrings.res_0024 -f $realDatabase)
        }
        else
        {
            write-verbose $MoveMailbox_LocalizedStrings.res_0025
            $realDatabase = ""
        }
        
        # try to get a handler from the "cache", if not available, create one
        if ($script:databaseHandlerMap[$realDatabase])
        {
            write-output $script:databaseHandlerMap[$realDatabase]
        }
        else
        {
            $dbObj = _CreateDatabaseMoveHandler -DatabaseName $realDatabase
            $script:databaseHandlerMap[$realDatabase] = $dbObj
            write-output $dbObj
        }
    }
    
    ## The main function for processing a single mailbox
    ## Gets the target handler, put the mailbox in the queue, if the queue is big enough, start moves
    function _ProcessMailbox([Parameter(ValueFromPipeline=$true, Mandatory=$true)] [Microsoft.Exchange.Data.Directory.Management.Mailbox]$mailbox)
    {
        Process {
            try
            {
                Write-Verbose ($MoveMailbox_LocalizedStrings.res_0008 -f $mailbox.Alias,$mailbox.Database.Name)
                $dbMoveHandler = _AcquireMoveHandler -Database $mailbox.Database.Name
                Write-Verbose ($MoveMailbox_LocalizedStrings.res_0009 -f $mailbox.Alias,$dbMoveHandler.Name)
                $dbMoveHandler.NewAlias += $mailbox
                if ($dbMoveHandler.NewAlias.Count -ge $startBatchSize)
                {
                    Write-Verbose ($MoveMailbox_LocalizedStrings.res_0010 -f $dbMoveHandler.Name)
                    $dbMoveHandler.Start()
                }
                
                # processed at least one mailbox, mark the flag as true
                $script:shouldStartMoves = $true
            }
            catch
            {
                $warningMessage = ($MoveMailbox_LocalizedStrings.res_0026 -f $mailbox.Alias, $error[0].Exception.Message)
                Write-Warning $warningMessage
            }
        }
    }
}

## Just pick one mailbox or database from the pipeline, make sure we have the mailbox objects, pass it in to _ProcessMailbox
Process {
    if ($PSCmdlet.ParameterSetName -eq "MoveDatabase")
    {
        Write-Verbose ($MoveMailbox_LocalizedStrings.res_0011 -f $mailboxDatabase)
        Get-Mailbox -Database $mailboxDatabase -ResultSize unlimited -Filter "MailboxMoveStatus -eq 'None'" | _ProcessMailbox
        Get-Mailbox -Arbitration -Database $mailboxDatabase -ResultSize unlimited -Filter "MailboxMoveStatus -eq 'None'" | _ProcessMailbox
    }
    elseif ($PSCmdlet.ParameterSetName -eq "MoveUser")
    {
        if ($identity -is [Microsoft.Exchange.Data.Directory.Management.Mailbox])
        {
            Write-Verbose ($MoveMailbox_LocalizedStrings.res_0012 -f $identity)
            # avoid remote PS trouble and pass the object straight to the processing function
            _ProcessMailbox -mailbox $identity
        }
        else
        {
            Write-Verbose ($MoveMailbox_LocalizedStrings.res_0013 -f $Identity)
            $temp = Get-Mailbox $identity -ErrorAction "SilentlyContinue"
            if ($temp)
            {
                $temp | _ProcessMailbox
            }
            $temp = Get-Mailbox -Arbitration $identity -ErrorAction "SilentlyContinue"
            if ($temp)
            {
                $temp | _ProcessMailbox 
            }
        }       
    }
}

End
{
    # start the moves on databases that didn't reach the batch size
    Write-Verbose $MoveMailbox_LocalizedStrings.res_0014
    $script:databaseHandlerMap.Values | %{ Write-Verbose ($MoveMailbox_LocalizedStrings.res_0015 -f $_.Name);$_.Start() }

    if ($SuspendWhenReadyToComplete)
    {
        Write-Warning $MoveMailbox_LocalizedStrings.res_0016
    }
    else
    {
        # complete all moves and capture stats
        
        ## Loop while there's at least one of the mailboxes we started that is moving
        Write-Progress -Id 500 -Activity $MoveMailbox_LocalizedStrings.progress_0008 -Status $MoveMailbox_LocalizedStrings.progress_0009 -PercentComplete 0

        [array]$mailboxList = @()
        $script:databaseHandlerMap.Values | %{ $mailboxList += $_.Started }

        [array]$inTransit = $mailboxList | Get-MoveRequest  -ErrorAction "SilentlyContinue" | ?{ $_.Status -ne 'Failed' -and $_.Status -ne 'Completed' -and $_.Status -ne 'CompletedWithWarning' }    

        if ($script:shouldStartMoves)
        {
            Write-Verbose $MoveMailbox_LocalizedStrings.res_0017
            # Check sanity, if there isn't at least one move to complete, throw        
            if ($mailboxList.Count -lt 1)
            {
                throw $MoveMailbox_LocalizedStrings.resthrow_0001
            }
            
            # capture the total number of moves that we'll need to complete (just for progress reporting sake :)
            # we'll be shrinking the $inTransit list to hammer AD less, so can't really count on that
            $totalMoves = $inTransit.Count
            
            if ($totalMoves -lt 1)
            {
                $totalMoves = 1
            }

            [int]$percentComplete=0
            [int]$movesCompleted=0
            [array]$readyToComplete=@()

            Write-Verbose ($MoveMailbox_LocalizedStrings.res_0018 -f $inTransit.Count)
            # loop here while there's at least one element inTransit, keep running this loop seeking for moves to complete
            while ($inTransit)
            {
                Write-Progress -Id 500 -Activity ($MoveMailbox_LocalizedStrings.progress_0010 -f $totalMoves) -Status $MoveMailbox_LocalizedStrings.progress_0011 -PercentComplete $percentComplete
                
                # update the inTransit array so our next calls to check for readiness are cheaper
                [array]$inTransit = $inTransit | Get-MoveRequest -ErrorAction "SilentlyContinue" | ?{ $_.Status -ne 'Failed' -and $_.Status -ne 'Completed' -and $_.Status -ne 'CompletedWithWarning' }
                
                $movesCompleted = $totalMoves - $inTransit.Count
                $percentComplete = $movesCompleted/$totalMoves*100
                
                Write-Debug "InTransit=$($inTransit.Count) { $InTransit }, $regularMailboxList , $arbitrationMailboxList, count -gt 0 = $($inTransit.Count -gt 0) and -ne null $($inTrasnit -ne $null)"
            }
            Write-Progress -Id 500 -Activity $MoveMailbox_LocalizedStrings.progress_0012 -Status $MoveMailbox_LocalizedStrings.progress_0013 -Completed
            Write-Verbose $MoveMailbox_LocalizedStrings.res_0019

            $movesCompleted = 0

            Write-Verbose ($MoveMailbox_LocalizedStrings.res_0020 -f $mailboxList)
            $completedMoves = $mailboxList | Get-MoveRequest -ErrorAction SilentlyContinue | ?{ $_.Status -eq 'Completed' -or $_.Status -eq 'CompletedWithWarning' -or $_.Status -eq 'Failed' }
            Write-Verbose ($MoveMailbox_LocalizedStrings.res_0021 -f $completedMoves.Count)
            $completedMoves | %{ 
                Write-Progress -Id 500 -Activity $MoveMailbox_LocalizedStrings.progress_0014 -Status ($MoveMailbox_LocalizedStrings.progress_0015 -f $_.Alias) -PercentComplete $([System.Math]::Min($($movesCompleted/$totalMoves*100), 100))
                $movesCompleted += 1
                Write-Verbose ($MoveMailbox_LocalizedStrings.res_0022 -f $_)
                $history = Get-MoveRequestStatistics $_ -IncludeReport
                    
                $obj = new-Object -TypeName PSObject

                Add-Member -InputObject $obj -MemberType NoteProperty -Name Status -Value $history.Status
                Add-Member -InputObject $obj -MemberType NoteProperty -Name QueuedTimestamp -Value $history.QueuedTimestamp
                Add-Member -InputObject $obj -MemberType NoteProperty -Name StartTimestamp -Value $history.StartTimestamp
                Add-Member -InputObject $obj -MemberType NoteProperty -Name InitialSeedingCompletedTimestamp -Value $history.InitialSeedingCompletedTimestamp
                Add-Member -InputObject $obj -MemberType NoteProperty -Name FinalSyncTimeStamp -Value $history.FinalSyncTimestamp
                Add-Member -InputObject $obj -MemberType NoteProperty -Name CompletedTimeStamp -Value $history.CompletionTimestamp
                Add-Member -InputObject $obj -MemberType NoteProperty -Name OverallDuration -Value $history.OverallDuration
                $offlineDuration = $history.FinalSyncTimestamp - $history.CompletionTimestamp
                Add-Member -InputObject $obj -MemberType NoteProperty -Name OfflineDuration -Value $offlineDuration
                Add-Member -InputObject $obj -MemberType NoteProperty -Name SecondsQueued  -Value  $history.StartTimestamp.Subtract($history.QueuedTimestamp).TotalSeconds
                Add-Member -InputObject $obj -MemberType NoteProperty -Name InitialSeedingDurationInMinutes -Value $history.InitialSeedingCompletedTimestamp.Subtract($history.StartTimestamp).TotalMinutes
                [double]$size = $history.TotalMailboxSize.ToMb()
                if ($history.TotalArchiveSize -ne $null)
                {
                    $size += $history.TotalArchiveSize.ToMb()
                }
                Add-Member -InputObject $obj -MemberType NoteProperty -Name TotalSize -Value $size
                Add-Member -InputObject $obj -MemberType NoteProperty -Name Mbps -Value $($size/($history.OverallDuration.TotalSeconds))
                if ($history.Status -ne "Completed")
                {
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name Report -Value $history.Report
                }

                Write-Output $obj
            }

            $mailboxList | Get-MoveRequest -ErrorAction SilentlyContinue | ?{ $_.Status -eq 'Completed' } | Remove-MoveRequest -Confirm:$false -ErrorAction "SilentlyContinue"
        }
    }
} 

# SIG # Begin signature block
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUMxAf+dgnKBgomPeE7pqScsSL
# MzOgghhqMIIE2jCCA8KgAwIBAgITMwAAASMwQ40kSDyg1wAAAAABIzANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTgxMDI0MjEwNzQw
# WhcNMjAwMTEwMjEwNzQwWjCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEm
# MCQGA1UECxMdVGhhbGVzIFRTUyBFU046MUE4Ri1FM0MzLUQ2OUQxJTAjBgNVBAMT
# HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQCyYt3Mdjll12pYYKRUadMqDK0tJrlPK1MhEo75C/BI1y2i
# r4HnxDl/BSAvi44lI4IKUFDS40WJPlnCHfEuwgvEUNuFrseh7bDkaWezo6W/F7i9
# PO7GAxmM139zh5p2a1XwfzqYpZE2hEucIkiOR82fg7I0s8itzwl1jWQmdAI4XAZN
# LeeXXof9Qm80uuUfEn6x/pANst2N+WKLRLnCqWLR7o6ZKqofshoYFpPukVLPsvU/
# ik/ch1kj2Ja53Zb+KHctMCk/CpN2p7fNArpLcUA3H7/mdJjlaUFYLY9yy5TBndFF
# I1kBbZEB/Z1kYVnjRsIsV8W2CCp1RCxiIkx6AhIzAgMBAAGjggEJMIIBBTAdBgNV
# HQ4EFgQU2zl1LgtoHHcQXPImRhW0WL0hxPAwHwYDVR0jBBgwFoAUIzT42VJGcArt
# QPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# bDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNV
# HSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAauzxVwcRLuLSItcW
# CHqZFtDSR5Ci4pgS+WrhLSmhfKXQRJekrOwZR7keC/bS7lyqai7y4NK9+ZHc2+F7
# dG3Ngym/92H45M/fRYtP63ObzWY9SNXBxEaZ1l8UP/hvv3uJnPf5/92mws50THX8
# tlvKAkBMWikcuA5y4s6yYy2GBFZIypm+ChZGtswTCst+uZhG8SBeE+U342Tbb3fG
# 5MLS+xuHrvSWdRqVHrWHpPKESBTStNPzR/dJ7pgtmF7RFKAWYLcEpPhr9hjUcf9q
# SJa7D5aghTY2UNFmn3BvKBSON+Dy5nDJA81RyZ/lU9iCOG+hGdpsGsJfvKT5WxsJ
# vEwdjzCCBf8wggPnoAMCAQICEzMAAAFRno2PQHGjDkEAAAAAAVEwDQYJKoZIhvcN
# AQELBQAwfjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYG
# A1UEAxMfTWljcm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMTAeFw0xOTA1MDIy
# MTM3NDZaFw0yMDA1MDIyMTM3NDZaMHQxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xHjAcBgNVBAMTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjCCASIw
# DQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJVaxoZpRx00HvFVw2Z19mJUGFgU
# ZyfwoyrGA0i85lY0f0lhAu6EeGYnlFYhLLWh7LfNO7GotuQcB2Zt5Tw0Uyjj0+/v
# UyAhL0gb8S2rA4fu6lqf6Uiro05zDl87o6z7XZHRDbwzMaf7fLsXaYoOeilW7SwS
# 5/LjneDHPXozxsDDj5Be6/v59H1bNEnYKlTrbBApiIVAx97DpWHl+4+heWg3eTr5
# CXPvOBxPhhGbHPHuMxWk/+68rqxlwHFDdaAH9aTJceDFpjX0gDMurZCI+JfZivKJ
# HkSxgGrfkE/tTXkOVm2lKzbAhhOSQMHGE8kgMmCjBm7kbKEd2quy3c6ORJECAwEA
# AaOCAX4wggF6MB8GA1UdJQQYMBYGCisGAQQBgjdMCAEGCCsGAQUFBwMDMB0GA1Ud
# DgQWBBRXghquSrnt6xqC7oVQFvbvRmKNzzBQBgNVHREESTBHpEUwQzEpMCcGA1UE
# CxMgTWljcm9zb2Z0IE9wZXJhdGlvbnMgUHVlcnRvIFJpY28xFjAUBgNVBAUTDTIz
# MDAxMis0NTQxMzUwHwYDVR0jBBgwFoAUSG5k5VAF04KqFzc3IrVtqMp1ApUwVAYD
# VR0fBE0wSzBJoEegRYZDaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9j
# cmwvTWljQ29kU2lnUENBMjAxMV8yMDExLTA3LTA4LmNybDBhBggrBgEFBQcBAQRV
# MFMwUQYIKwYBBQUHMAKGRWh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMv
# Y2VydHMvTWljQ29kU2lnUENBMjAxMV8yMDExLTA3LTA4LmNydDAMBgNVHRMBAf8E
# AjAAMA0GCSqGSIb3DQEBCwUAA4ICAQBaD4CtLgCersquiCyUhCegwdJdQ+v9Go4i
# Elf7fY5u5jcwW92VESVtKxInGtHL84IJl1Kx75/YCpD4X/ZpjAEOZRBt4wHyfSlg
# tmc4+J+p7vxEEfZ9Vmy9fHJ+LNse5tZahR81b8UmVmUtfAmYXcGgvwTanT0reFqD
# DP+i1wq1DX5Dj4No5hdaV6omslSycez1SItytUXSV4v9DVXluyGhvY5OVmrSrNJ2
# swMtZ2HKtQ7Gdn6iNntR1NjhWcK6iBtn1mz2zIluDtlRL1JWBiSjBGxa/mNXiVup
# MP60bgXOE7BxFDB1voDzOnY2d36ztV0K5gWwaAjjW5wPyjFV9wAyMX1hfk3aziaW
# 2SqdR7f+G1WufEooMDBJiWJq7HYvuArD5sPWQRn/mjMtGcneOMOSiZOs9y2iRj8p
# pnWq5vQ1SeY4of7fFQr+mVYkrwE5Bi5TuApgftjL1ZIo2U/ukqPqLjXv7c1r9+si
# eOcGQpEIn95hO8Ef6zmC57Ol9Ba1Ths2j+PxDDa+lND3Dt+WEfvxGbB3fX35hOaG
# /tNzENtaXK15qPhErbCTeljWhLPYk8Tk8242Z30aZ/qh49mDLsiL0ksurxKdQtXt
# v4g/RRdFj2r4Z1GMzYARfqaxm+88IigbRpgdC73BmwoQraOq9aLz/F1555Ij0U3o
# rXDihVAzgzCCBgcwggPvoAMCAQICCmEWaDQAAAAAABwwDQYJKoZIhvcNAQEFBQAw
# XzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jvc29m
# dDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# MB4XDTA3MDQwMzEyNTMwOVoXDTIxMDQwMzEzMDMwOVowdzELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgUENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAn6Fssd/b
# SJIqfGsuGeG94uPFmVEjUK3O3RhOJA/u0afRTK10MCAR6wfVVJUVSZQbQpKumFww
# JtoAa+h7veyJBw/3DgSY8InMH8szJIed8vRnHCz8e+eIHernTqOhwSNTyo36Rc8J
# 0F6v0LBCBKL5pmyTZ9co3EZTsIbQ5ShGLieshk9VUgzkAyz7apCQMG6H81kwnfp+
# 1pez6CGXfvjSE/MIt1NtUrRFkJ9IAEpHZhEnKWaol+TTBoFKovmEpxFHFAmCn4Tt
# VXj+AZodUAiFABAwRu233iNGu8QtVJ+vHnhBMXfMm987g5OhYQK1HQ2x/PebsgHO
# IktU//kFw8IgCwIDAQABo4IBqzCCAacwDwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4E
# FgQUIzT42VJGcArtQPt2+7MrsMM1sw8wCwYDVR0PBAQDAgGGMBAGCSsGAQQBgjcV
# AQQDAgEAMIGYBgNVHSMEgZAwgY2AFA6sgmBAVieX5SUT/CrhClOVWeSkoWOkYTBf
# MRMwEQYKCZImiZPyLGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0
# MS0wKwYDVQQDEyRNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHmC
# EHmtFqFKoKWtTHNY9AcTLmUwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC5t
# aWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNlcnQu
# Y3JsMFQGCCsGAQUFBwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraS9jZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwEwYDVR0l
# BAwwCgYIKwYBBQUHAwgwDQYJKoZIhvcNAQEFBQADggIBABCXisNcA0Q23em0rXfb
# znlRTQGxLnRxW20ME6vOvnuPuC7UEqKMbWK4VwLLTiATUJndekDiV7uvWJoc4R0B
# hqy7ePKL0Ow7Ae7ivo8KBciNSOLwUxXdT6uS5OeNatWAweaU8gYvhQPpkSokInD7
# 9vzkeJkuDfcH4nC8GE6djmsKcpW4oTmcZy3FUQ7qYlw/FpiLID/iBxoy+cwxSnYx
# PStyC8jqcD3/hQoT38IKYY7w17gX606Lf8U1K16jv+u8fQtCe9RTciHuMMq7eGVc
# WwEXChQO0toUmPU8uWZYsy0v5/mFhsxRVuidcJRsrDlM1PZ5v6oYemIp76KbKTQG
# dxpiyT0ebR+C8AvHLLvPQ7Pl+ex9teOkqHQ1uE7FcSMSJnYLPFKMcVpGQxS8s7Ow
# TWfIn0L/gHkhgJ4VMGboQhJeGsieIiHQQ+kr6bv0SMws1NgygEwmKkgkX1rqVu+m
# 3pmdyjpvvYEndAYR7nYhv5uCwSdUtrFqPYmhdmG0bqETpr+qR/ASb/2KMmyy/t9R
# yIwjyWa9nR2HEmQCPS2vWY+45CHltbDKY7R4VAXUQS5QrJSwpXirs6CWdRrZkocT
# dSIvMqgIbqBbjCW/oO+EyiHW6x5PyZruSeD3AWVviQt9yGnI5m7qp5fOMSn/DsVb
# XNhNG6HY+i+ePy5VFmvJE6P9MIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
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
# /Xmfwb1tbWrJUnMTDXpQzTGCBKgwggSkAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAFRno2PQHGjDkEAAAAAAVEwCQYFKw4DAhoFAKCB
# vDAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYK
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUjT0eEvtEoq7MI7kCPTCVYC98Rxcw
# XAYKKwYBBAGCNwIBDDFOMEygJIAiAEMASQBUAFMAQwBvAG4AcwB0AGEAbgB0AHMA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBAF+LrI1BkiUa2pZQOaXjbEjoieJFPiC9Lq7MnafOGCPE
# 0oHMhAdGZSil/lu7XvtvfqpSGDWZIchjCOY/iMmuyfm9lyOZdmY1z1l0BVIRv2qf
# 0rzfUtu1OdESvSkUf1BD+dCnNfIDUlv7EmJzMZJnvoTbYbIfh+gCcJ15TBFTU1wF
# ukqKnsLVXSIRfYJ1txLpAcCeUvtfJmQOStqXqvPnG+j8YleNzcYjuFb053ai8/l2
# JUxXwJlEZmzWOj5L+YME/O9bsjOKydkzY1XlOHWc0ZXzlfYQqflYpGn+qakI021z
# Z7lG/tW3OoXiXMgzjrTCORlM+Qq0eBtEEWiJEorOIpqhggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAB
# IzBDjSRIPKDXAAAAAAEjMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTA2MjYxNTI2MDBaMCMGCSqGSIb3DQEJ
# BDEWBBS2CrWXfG2DsReF72BHQvDUQYi5rzANBgkqhkiG9w0BAQUFAASCAQCL7OKw
# Cj6McggHvghWwtUlPULv90YO28/HNIaAQ2EYAMNx6IxhjO40tYTb8eeMw2XlMsDl
# pfLMz/dylf80Lppr3ZIm6cAQU0qRANcbwpQXgn3WXRXohrIWT3xmI+gaoj/3utPX
# 223m41TnXMW+4N/04EXjb68h7u/4s0ExXUW2C/D36v45BXHSZ5VSTGwnqjGzWKgL
# qxlW3cbXsgIUk3+hEojWIE2Mjh4miv8ia1uqk/OiGG/3pj45vQ40VkZvH8L4+zoN
# J+w7Av9m5Af9W2f9Oc07vYPyPHpo+1cyLcil2OJsMfkMriGdJaWRtlziZOguBtlA
# In8LrFxsfQI4ABGH
# SIG # End signature block

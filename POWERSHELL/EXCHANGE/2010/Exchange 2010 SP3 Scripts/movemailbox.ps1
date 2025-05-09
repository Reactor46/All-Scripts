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
# MIIaZAYJKoZIhvcNAQcCoIIaVTCCGlECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUMxAf+dgnKBgomPeE7pqScsSL
# MzOgghUvMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
# 9w0BAQUFADB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSMw
# IQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQTAeFw0xMjA5MDQyMTQy
# MDlaFw0xMzAzMDQyMTQyMDlaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQC6pElsEPsi
# nGWiFpg7y2Fi+nQprY0GGdJxWBmKXlcNaWJuNqBO/SJ54B3HGmGO+vyjESUWyMBY
# LDGKiK4yHojbfz50V/eFpDZTykHvabhpnm1W627ksiZNc9FkcbQf1mGEiAAh72hY
# g1tJj7Tf0zXWy9kwn1P8emuahCu3IWd01PZ4tmGHmJR8Ks9n6Rm+2bpj7TxOPn0C
# 6/N/r88Pt4F+9Pvo95FIu489jMgHkxzzvXXk/GMgKZ8580FUOB5UZEC0hKo3rvMA
# jOIN+qGyDyK1p6mu1he5MPACIyAQ+mtZD+Ctn55ggZMDTA2bYhmzu5a8kVqmeIZ2
# m2zNTOwStThHAgMBAAGjggENMIIBCTATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNV
# HQ4EFgQU3lHcG/IeSgU/EhzBvMOzZSyRBZgwHwYDVR0jBBgwFoAUyxHoytK0FlgB
# yTcuMxYWuUyaCh8wVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljQ29kU2lnUENBXzA4LTMxLTIwMTAu
# Y3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraS9jZXJ0cy9NaWNDb2RTaWdQQ0FfMDgtMzEtMjAxMC5jcnQw
# DQYJKoZIhvcNAQEFBQADggEBACqk9+7AwyZ6g2IaeJxbxf3sFcSneBPRF1MoCwwA
# Qj84D4ncZBmENX9Iuc/reomhzU+p4LvtRxD+F9qHiRDRTBWg8BH/2pbPZM+B/TOn
# w3iT5HzVbYdx1hxh4sxOZLdzP/l7JzT2Uj9HQ8AOgXBTwZYBoku7vyoDd3tu+9BG
# ihcoMaUF4xaKuPFKaRVdM/nff5Q8R0UdrsqLx/eIHur+kQyfTwcJ7SaSbrOUGQH4
# X4HnrtqJj39aXoRftb58RuVHr/5YK5F/h9xGH1GVzMNiobXHX+vJaVxxkamNViAs
# Ok6T/ZsGj62K+Gh+O7p5QpM5SfXQXuxwjUJ1xYJVkBu1VWEwggTDMIIDq6ADAgEC
# AhMzAAAAKzkySMGyyUjzAAAAAAArMA0GCSqGSIb3DQEBBQUAMHcxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBU
# aW1lLVN0YW1wIFBDQTAeFw0xMjA5MDQyMTEyMzRaFw0xMzEyMDQyMTEyMzRaMIGz
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMQ0wCwYDVQQLEwRN
# T1BSMScwJQYDVQQLEx5uQ2lwaGVyIERTRSBFU046QzBGNC0zMDg2LURFRjgxJTAj
# BgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCmtjAOA2WuUFqGa4WfSKEeycDuXkkHheBwlny+
# uV9iXwYm04s5uxgipS6SrdhLiDoar5uDrsheOYzCMnsWeO03ODrxYvtoggJo7Ou7
# QIqx/qEsNmJgcDlgYg77xhg4b7CS1kANgKYNeIs2a4aKJhcY/7DrTbq7KRPmXEiO
# cEY2Jv40Nas04ffa2FzqmX0xt00fV+t81pUNZgweDjIXPizVgKHO6/eYkQLcwV/9
# OID4OX9dZMo3XDtRW12FX84eHPs0vl/lKFVwVJy47HwAVUZbKJgoVkzh8boJGZaB
# SCowyPczIGznacOz1MNOzzAeN9SYUtSpI0WyrlxBSU+0YmiTAgMBAAGjggEJMIIB
# BTAdBgNVHQ4EFgQUpRgzUz+VYKFDFu+Oxq/SK7qeWNAwHwYDVR0jBBgwFoAUIzT4
# 2VJGcArtQPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5t
# aWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1w
# UENBLmNybDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# dDATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAfsywe+Uv
# vudWtc9z26pS0RY5xrTN+tf+HmW150jzm0aIBWZqJoZe/odY3MZjjjiA9AhGfCtz
# sQ6/QarLx6qUpDfwZDnhxdX5zgfOq+Ql8Gmu1Ebi/mYyPNeXxTIh+u4aJaBeDEIs
# ETM6goP97R2zvs6RpJElcbmrcrCer+TPAGKJcKm4SlCM7i8iZKWo5k1rlSwceeyn
# ozHakGCQpG7+kwINPywkDcZqJoFRg0oQu3VjRKppCMYD6+LPC+1WOuzvcqcKDPQA
# 0yK4ryJys+fEnAsooIDK4+HXOWYw50YXGOf6gvpZC3q8qA3+HP8Di2OyTRICI08t
# s4WEO+KhR+jPFTCCBbwwggOkoAMCAQICCmEzJhoAAAAAADEwDQYJKoZIhvcNAQEF
# BQAwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jv
# c29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9y
# aXR5MB4XDTEwMDgzMTIyMTkzMloXDTIwMDgzMTIyMjkzMloweTELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEjMCEGA1UEAxMaTWljcm9zb2Z0IENv
# ZGUgU2lnbmluZyBQQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCy
# cllcGTBkvx2aYCAgQpl2U2w+G9ZvzMvx6mv+lxYQ4N86dIMaty+gMuz/3sJCTiPV
# cgDbNVcKicquIEn08GisTUuNpb15S3GbRwfa/SXfnXWIz6pzRH/XgdvzvfI2pMlc
# RdyvrT3gKGiXGqelcnNW8ReU5P01lHKg1nZfHndFg4U4FtBzWwW6Z1KNpbJpL9oZ
# C/6SdCnidi9U3RQwWfjSjWL9y8lfRjFQuScT5EAwz3IpECgixzdOPaAyPZDNoTgG
# hVxOVoIoKgUyt0vXT2Pn0i1i8UU956wIAPZGoZ7RW4wmU+h6qkryRs83PDietHdc
# pReejcsRj1Y8wawJXwPTAgMBAAGjggFeMIIBWjAPBgNVHRMBAf8EBTADAQH/MB0G
# A1UdDgQWBBTLEejK0rQWWAHJNy4zFha5TJoKHzALBgNVHQ8EBAMCAYYwEgYJKwYB
# BAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQU/dExTtMmipXhmGA7qDFvpjy8
# 2C0wGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwHwYDVR0jBBgwFoAUDqyCYEBW
# J5flJRP8KuEKU5VZ5KQwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC5taWNy
# b3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNlcnQuY3Js
# MFQGCCsGAQUFBwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5taWNyb3Nv
# ZnQuY29tL3BraS9jZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwDQYJKoZIhvcN
# AQEFBQADggIBAFk5Pn8mRq/rb0CxMrVq6w4vbqhJ9+tfde1MOy3XQ60L/svpLTGj
# I8x8UJiAIV2sPS9MuqKoVpzjcLu4tPh5tUly9z7qQX/K4QwXaculnCAt+gtQxFbN
# LeNK0rxw56gNogOlVuC4iktX8pVCnPHz7+7jhh80PLhWmvBTI4UqpIIck+KUBx3y
# 4k74jKHK6BOlkU7IG9KPcpUqcW2bGvgc8FPWZ8wi/1wdzaKMvSeyeWNWRKJRzfnp
# o1hW3ZsCRUQvX/TartSCMm78pJUT5Otp56miLL7IKxAOZY6Z2/Wi+hImCWU4lPF6
# H0q70eFW6NB4lhhcyTUWX92THUmOLb6tNEQc7hAVGgBd3TVbIc6YxwnuhQ6MT20O
# E049fClInHLR82zKwexwo1eSV32UjaAbSANa98+jZwp0pTbtLS8XyOZyNxL0b7E8
# Z4L5UrKNMxZlHg6K3RDeZPRvzkbU0xfpecQEtNP7LN8fip6sCvsTJ0Ct5PnhqX9G
# uwdgR2VgQE6wQuxO7bN2edgKNAltHIAxH+IOVN3lofvlRxCtZJj/UBYufL8FIXri
# lUEnacOTj5XJjdibIa4NXJzwoq6GaIMMai27dmsAHZat8hZ79haDJLmIz2qoRzEv
# mtzjcT3XAH5iR9HOiMm4GPoOco3Boz2vAkBq/2mbluIQqBC0N1AI1sM9MIIGBzCC
# A++gAwIBAgIKYRZoNAAAAAAAHDANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZImiZPy
# LGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQDEyRN
# aWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMDcwNDAzMTI1
# MzA5WhcNMjEwNDAzMTMwMzA5WjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwggEi
# MA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCfoWyx39tIkip8ay4Z4b3i48WZ
# USNQrc7dGE4kD+7Rp9FMrXQwIBHrB9VUlRVJlBtCkq6YXDAm2gBr6Hu97IkHD/cO
# BJjwicwfyzMkh53y9GccLPx754gd6udOo6HBI1PKjfpFzwnQXq/QsEIEovmmbJNn
# 1yjcRlOwhtDlKEYuJ6yGT1VSDOQDLPtqkJAwbofzWTCd+n7Wl7PoIZd++NIT8wi3
# U21StEWQn0gASkdmEScpZqiX5NMGgUqi+YSnEUcUCYKfhO1VeP4Bmh1QCIUAEDBG
# 7bfeI0a7xC1Un68eeEExd8yb3zuDk6FhArUdDbH895uyAc4iS1T/+QXDwiALAgMB
# AAGjggGrMIIBpzAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBQjNPjZUkZwCu1A
# +3b7syuwwzWzDzALBgNVHQ8EBAMCAYYwEAYJKwYBBAGCNxUBBAMCAQAwgZgGA1Ud
# IwSBkDCBjYAUDqyCYEBWJ5flJRP8KuEKU5VZ5KShY6RhMF8xEzARBgoJkiaJk/Is
# ZAEZFgNjb20xGTAXBgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1p
# Y3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eYIQea0WoUqgpa1Mc1j0
# BxMuZTBQBgNVHR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20v
# cGtpL2NybC9wcm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUH
# AQEESDBGMEQGCCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# L2NlcnRzL01pY3Jvc29mdFJvb3RDZXJ0LmNydDATBgNVHSUEDDAKBggrBgEFBQcD
# CDANBgkqhkiG9w0BAQUFAAOCAgEAEJeKw1wDRDbd6bStd9vOeVFNAbEudHFbbQwT
# q86+e4+4LtQSooxtYrhXAstOIBNQmd16QOJXu69YmhzhHQGGrLt48ovQ7DsB7uK+
# jwoFyI1I4vBTFd1Pq5Lk541q1YDB5pTyBi+FA+mRKiQicPv2/OR4mS4N9wficLwY
# Tp2OawpylbihOZxnLcVRDupiXD8WmIsgP+IHGjL5zDFKdjE9K3ILyOpwPf+FChPf
# wgphjvDXuBfrTot/xTUrXqO/67x9C0J71FNyIe4wyrt4ZVxbARcKFA7S2hSY9Ty5
# ZlizLS/n+YWGzFFW6J1wlGysOUzU9nm/qhh6YinvopspNAZ3GmLJPR5tH4LwC8cs
# u89Ds+X57H2146SodDW4TsVxIxImdgs8UoxxWkZDFLyzs7BNZ8ifQv+AeSGAnhUw
# ZuhCEl4ayJ4iIdBD6Svpu/RIzCzU2DKATCYqSCRfWupW76bemZ3KOm+9gSd0BhHu
# diG/m4LBJ1S2sWo9iaF2YbRuoROmv6pH8BJv/YoybLL+31HIjCPJZr2dHYcSZAI9
# La9Zj7jkIeW1sMpjtHhUBdRBLlCslLCleKuzoJZ1GtmShxN1Ii8yqAhuoFuMJb+g
# 74TKIdbrHk/Jmu5J4PcBZW+JC33Iacjmbuqnl84xKf8OxVtc2E0bodj6L54/LlUW
# a8kTo/0xggSfMIIEmwIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQIT
# MwAAAJ0ejSeuuPPYOAABAAAAnTAJBgUrDgMCGgUAoIG4MBkGCSqGSIb3DQEJAzEM
# BgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqG
# SIb3DQEJBDEWBBSNPR4S+0SirswjuQI9MJVgL3xHFzBYBgorBgEEAYI3AgEMMUow
# SKAggB4ATQBvAHYAZQBNAGEAaQBsAGIAbwB4AC4AcABzADGhJIAiaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQA1N6zU
# jV/r1URiRI6fzO0fBuK8Am0VHYigmFPmpYl1duSbcxZHcHTIy6P27Esq/OxFxUpi
# pj28KrDJwfv0pegPfVjpkMepj6Vs5O7JkmSoo4vj3v3JFWNtb+cB+oAD6YC68jRv
# ffvsZjjsJ1xtU4aEEM2iUqFiuP5rSbO9Vv7NZSaxdqw+2JWV8T5KXVeRHSGaaA1s
# T7BHprkHcjoV2kV9eim2zj+nhghZ8yhy5n3DExIatP1Jpfn7tLZXSiX6JkbDcHff
# BUmeS1e2vdLnNxJVrEXUSMHXFvaMVEWW+xzFx0mwnbh7mfG748R8qGPfry4+5prP
# HcNp9kF/q4NDQVlUoYICKDCCAiQGCSqGSIb3DQEJBjGCAhUwggIRAgEBMIGOMHcx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFBDQQITMwAAACs5MkjBsslI8wAAAAAAKzAJBgUr
# DgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUx
# DxcNMTMwMjA1MDYzNzI1WjAjBgkqhkiG9w0BCQQxFgQUyMqwazVgiCIw2JEJ0IQA
# QGuL0i4wDQYJKoZIhvcNAQEFBQAEggEAmw2+YHj2clgHqqSyMqCmSey1MUcUX23u
# 9LveJqEern9WRlVeLNOJV9zwG43khhCX2Z/cBqoCrbeHI1/xAf/LoyadqZO32nrE
# ljt9lruYiWb2jJauakl8C+59FdxRysvGmuSTAc+eOv/3XNFpWqceIG93o8d0HPkC
# tEOqlHgNF20BK35TTqxWDMmVuXdTXO3Do4mhAoZFyhuUWU4/xEU4cLhgtkJ7VcSC
# aBXMqW1wnh5whY3Nh/V7buZpxfvboUD6akdvgzfrfNgPW8TRFDH653yolnw47EP6
# SLM6QG8zuCQ+iyKAfoB5l6hu6BODYBGu8O0wpkaFwG3v2LjehomIow==
# SIG # End signature block

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
# MIIdngYJKoZIhvcNAQcCoIIdjzCCHYsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUMxAf+dgnKBgomPeE7pqScsSL
# MzOgghhkMIIEwzCCA6ugAwIBAgITMwAAAKxjFufjRlWzHAAAAAAArDANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwNTAzMTcxMzIz
# WhcNMTcwODAzMTcxMzIzWjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OkMwRjQtMzA4Ni1ERUY4MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAnyHdhNxySctX
# +G+LSGICEA1/VhPVm19x14FGBQCUqQ1ATOa8zP1ZGmU6JOUj8QLHm4SAwlKvosGL
# 8o03VcpCNsN+015jMXbhhP7wMTZpADTl5Ew876dSqgKRxEtuaHj4sJu3W1fhJ9Yq
# mwep+Vz5+jcUQV2IZLBw41mmWMaGLahpaLbul+XOZ7wi2+qfTrPVYpB3vhVMwapL
# EkM32hsOUfl+oZvuAfRwPBFxY/Gm0nZcTbB12jSr8QrBF7yf1e/3KSiqleci3GbS
# ZT896LOcr7bfm5nNX8fEWow6WZWBrI6LKPx9t3cey4tz0pAddX2N6LASt3Q0Hg7N
# /zsgOYvrlwIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFCFXLAHtg1Boad3BTWmrjatP
# lDdiMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAEY2iloCmeBNdm4IPV1pQi7f4EsNmotUMen5D8Dg4rOLE9Jk
# d0lNOL5chmWK+d9BLG5SqsP0R/gqph4hHFZM4LVHUrSxQcQLWBEifrM2BeN0G6Yp
# RiGB7nnQqq86+NwX91pLhJ5LBzJo+EucWFKFmEBXLMBL85fyCusCk0RowdHpqh5s
# 3zhkMgjFX+cXWzJXULfGfEPvCXDKIgxsc5kUalYie/mkCKbpWXEW6gN+FNPKTbvj
# HcCxtcf9mVeqlA5joTFe+JbMygtOTeX0Mlf4rTvCrf3kA0zsRJL/y5JdihdxSP8n
# KX5H0Q2CWmDDY+xvbx9tLeqs/bETpaMz7K//Af4wggYHMIID76ADAgECAgphFmg0
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
# bWrJUnMTDXpQzTGCBKQwggSgAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCBuDAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUjT0eEvtEoq7MI7kCPTCVYC98RxcwWAYKKwYB
# BAGCNwIBDDFKMEigIIAeAE0AbwB2AGUATQBhAGkAbABiAG8AeAAuAHAAcwAxoSSA
# Imh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZIhvcNAQEB
# BQAEggEAR5+J/MHVxGTdyceSgAo7AtHLlG+TnwQ066EXeNOVyh8KUy0BAKUFAaEx
# ia7GlmmjkBzzUxgXA0KUEdwzRxQsBBND9T1FkCcM1909wTFkfkoC/WFf75d5YuRF
# MYo3IGI0xVWaDkDSyvXqcNBjslQpz0HjOkYVdFXIP6QwWDDgxeVcz74aQ1cZuBN0
# uRbX7+G5iVTvMQCNRmrQTXt/D00lcDsdl5zgYRhtbIarL/fyCYSNMgCTc1nwvWq6
# YsSkaUAztrEOFzj2qRSpyJ2jLcxUeTmal19NFixdqpk1wA9R1fXteCpHjvb80ffe
# 9AET0AMxHgT9F1M7ws3PWxgB9Gp9oqGCAigwggIkBgkqhkiG9w0BCQYxggIVMIIC
# EQIBATCBjjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0ECEzMAAACsYxbn40ZVsxwA
# AAAAAKwwCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJ
# KoZIhvcNAQkFMQ8XDTE2MDkwMzE4NDUxOFowIwYJKoZIhvcNAQkEMRYEFMn/po/b
# WlVeipQGs+H/ki3AcFBwMA0GCSqGSIb3DQEBBQUABIIBAAPTSmVjLaPXZEE2ksEY
# mzeYNYsNnL28VtSOf2uLJ2mVEls1TZwyA+hgUme0HwCGueOEKTv+mIjRCOlbj0iB
# fCaV7DmTZRZ1s7f6SY3AScNhDGpUHw0lyRzdFotxagmtNc8xFthahTa8ttkwFn9g
# m8tBICT+vtZliliCRZDW+eInvRYpjjQcP1CnEI5E+LKtuUD5Zp2+msL16o0N7iFy
# LvhlNX/kyip0XaUgokedjRWEMVdtRYgfu4Pd8XEHsQ75alOHDrUR3GqXpf2s7ulH
# H+R39ZHT2cVKDdyMXylXVpGvN2OPE+ySPxaxBGA5oZW9oVVGP32SNexLa0gQZ98O
# GNo=
# SIG # End signature block

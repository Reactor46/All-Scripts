# Copyright (c) Microsoft Corporation. All rights reserved.  
#
# Cmdlet takes IDs of multiple arrays, but operates with every one of them in isolation.
# All password changes to Active Directory are done after processing all arrays. The script will ask with 
# "medium impact" before propagating a password change to the Active Directory, especially if
# there were any failures setting passwords on individual CAS boxes.
#
# Target scenarios:
#   a.	First time setup and distribution of an alternate service account for Kerberos authentication.  Set up whole forest or a specific array of load balanced CAS servers.
#   b.	Standard password rollover of Kerberos alternate service account credential
#   c.	Adding one or more machines to a CAS array or rerunning on a specific computer
#   d.  Creating a scheduled task to perform ongoing password maintenance

<#
  .Synopsis
  Performs a first-time or periodic maintenance on the alternate service account (ASA) credential used to enable Kerberos authentication on Client Access Arrays.
    
  .Description
  The script functions in two mutually exclusive modes.  1. Generation of a new password for the ASA and subsequent propagation this password to destination servers and AD.  2. Copying the ASA credentials from a specified server and distributing to destination servers.  The mode is chosen by specifying either -GenerateNewPasswordFor or -CopyFrom parameters. In both modes, the set of destination servers to which credentials are appended or copied is determined by the mutually exclusive -ToEntireForest, -ToArrayMembers and -ToSpecificServers parameters.

  To have the script run in a scheduled task specify the -CreateScheduledTask parameter. Instead of immediately executing the maintenance procedure with the specified parameters, the script will schedule its own periodic execution.  A .cmd file with the specified name is created in the current directory and a corresponding task in Windows Task Scheduler is created to run the command every 3 weeks.  Use Windows Task Scheduler to run or modify properties of a task once it is created.  Use -Verbose to see the directory the scheduled task will log to.

  The script honors common parameters, like -Confirm, -Debug, -Verbose, -WhatIf.

  .Parameter ToEntireForest
  Specifies all members of all Client Access Arrays within the current forest as the destination servers.

  .Parameter ToArrayMembers
  Specifies all members of the specified Client Access Array(s) as the destination servers.
  Use the -Identity parameter to specify identities of the Client Access Array(s).

  .Parameter ToSpecificServers
  Specifies the destination Client Access Server or Servers.
  Use the -Identity parameter to specify identities of the server(s).

  .Parameter Identity
  Specifies identities of Client Access Arrays or servers, depending on whether -ToArrayMembers is used or -ToSpecificServers.

  .Parameter GenerateNewPasswordFor
  Specifies an account in the DOMAIN\AccountName form.  Acount is to be used as the alternate service account for Kerberos authentication.  The script will generate a new password for the account, distribute this new password to all destination servers then update AD with the newly generated password.  You must pass a trailing '$' for computer accounts.

  .Parameter CopyFrom
  Specifies the identity of a Client Access Server to copy all Alternate Service Account credentials from.  If specified server does not have any Alternate Service Account credentials, any credentials on destination servers will be removed.  There will be a prompt shown in this case on whether to proceed or fail

  .Parameter Unattended
  Suppresses confirmation prompts and provides "safest" answers to all the questions. Always used when executed from a scheduled task created with -CreateScheduledTask.

  .Parameter CreateScheduledTask
  Name of a scheduled task to create or overwrite for periodic execution of this script with the given parameters.
    
  .Inputs
  None. You cannot pipe objects to RollAlternateServiceAccountPassword.ps1.

  .Outputs
  None. All output is printed to a host.

  .Example
  .\RollAlternateServiceAccountPassword.ps1 -ToEntireForest -GenerateNewPasswordFor 'contoso\computerAccount$' -Verbose

  Generate a new password for the specified computer account and use it as an Alternate Service Account on Client Access servers - members of all Client Access Arrays in the entire forest.

  .Example
  .\RollAlternateServiceAccountPassword.ps1 -ToArrayMembers *mailbox* -GenerateNewPasswordFor contoso\userAccount -Verbose

  Generate a new password for the specified user account and use it as an Alternate Service Account on Client Access servers - members of the Client Access Arrays whose names or FQDNs correspond to the *mailbox* wildcard.

  .Example
  .\RollAlternateServiceAccountPassword.ps1 -CopyFrom ServerA -ToSpecificServers ServerB -Verbose

  Copies Alternate Service Account settings from ServerA to ServerB. This is useful if ServerB was down or not yet a member of the array when the password was rolled the last time.

  .Example
  .\RollAlternateServiceAccountPassword.ps1 -CreateScheduledTask "Exchange-RollAsa" -ToEntireForest -GenerateNewPasswordFor 'contoso\computerAccount$'

  Schedules an automated password roll every 3 weeks.

  .Example
  SchTasks.exe /Run /TN "Exchange-RollAsa"

  Immediately runs a previously configured scheduled password maintenance task. Can be useful if the previous scheduled iteration failed due to an error or if array membership has changed.
#>

[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact = 'Medium')]
param
(
  [Parameter(ParameterSetName="Forest-Generate")]
  [Parameter(ParameterSetName="Forest-CopyFrom")]
  [Switch]
  $ToEntireForest, 

  [Parameter(ParameterSetName="Array-Generate")]
  [Parameter(ParameterSetName="Array-CopyFrom")]
  [Switch]
  $ToArrayMembers, 

  [Parameter(ParameterSetName="Server-Generate")]
  [Parameter(ParameterSetName="Server-CopyFrom")]
  [Switch]
  $ToSpecificServers, 

  [Parameter(Mandatory=$true, Position=0, ParameterSetName="Array-Generate")]
  [Parameter(Mandatory=$true, Position=0, ParameterSetName="Array-CopyFrom")]
  [Parameter(Mandatory=$true, Position=0, ParameterSetName="Server-Generate")]
  [Parameter(Mandatory=$true, Position=0, ParameterSetName="Server-CopyFrom")]
  [ValidateNotNullOrEmpty()]
  [string[]]
  $Identity, 

  # Should we disallow Server-Generate?
  [Parameter(ParameterSetName="Forest-Generate")]
  [Parameter(ParameterSetName="Array-Generate")]
  [Parameter(ParameterSetName="Server-Generate")]
  [ValidatePattern('^[^\s\\]+\\[^\s\\]+$')]
  [string]
  $GenerateNewPasswordFor, 

  [Parameter(ParameterSetName="Forest-CopyFrom")]
  [Parameter(ParameterSetName="Array-CopyFrom")]
  [Parameter(ParameterSetName="Server-CopyFrom")]
  [ValidateNotNullOrEmpty()]
  [string]
  $CopyFrom, 

  [Switch]
  $Unattended, 

  [string]
  $CreateScheduledTask
)

#####################
# Library functions #
#####################
function Select-Unique([ScriptBlock]$keySelector = { $_ })
{
  begin { $keys = @{ } }
  process
  {
    $key = $_ | & $keySelector
    if ($keys[$key] -eq $null)
    {
      $keys[$key] = $true
      $_
    }
  }
}

function SubtractSet([object[]]$x, [object[]]$y)
{
  $x | Where { $y -notcontains $_ }
}

function Write-OperationProgress([string]$activity, [ScriptBlock]$inputToDescription = { $_ })
{
  $all_items = @($input)
  $current = 0
  $total = $all_items.Length
  $all_items | foreach { `
    Write-Progress `
      -Activity $activity `
      -Status "Current item (out of $total)" `
      -CurrentOperation ($_ | & $inputToDescription) `
      -PercentComplete ($current*100 / $total); `
    $current++; `
    $_ }
}

function GenerateNewPassword([string]$qualifiedAccountName)
{
  # This password should always satisfy the standard variations of the account policies:
  # 1) minimum length - the maximum value that a policy can set is 14. Generated passwords are 73 chars long
  # 2) 3/4 of the following character sets should be present:
  #   a) non-character
  #   b) digit
  #   c) capital-case latin
  #   d) lower-case latin
  $password = ConvertTo-SecureString -Force -AsPlainText -String ("{0:D}X{1:D}" -f [Guid]::NewGuid(), [Guid]::NewGuid())
  New-Object 'System.Management.Automation.PSCredential' -Arg $qualifiedAccountName, $password    
}

function GetAccount([string]$qualifiedUserName, [bool]$throwOnNotFound = $true)
{
  if (-not $throwOnNotFound)
  {
    try
      { GetAccount $qualifiedUserName $true }
    catch [System.Management.Automation.RuntimeException]
      { Write-Verbose $error[0] }
    return
  }

  $parts = @($qualifiedUserName.Split('\'))

  if ($parts.Count -ne 2 -or $parts[0] -eq '')
    { throw "The username must be fully qualified" } 

  $domain = $parts[0]
  $userName = $parts[1]

  Write-Verbose "Looking up account $userName in domain $domain"

  $null = [System.Reflection.Assembly]::Load('System.DirectoryServices.AccountManagement, Version=3.5.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
  $context = $null
  try
  {
    $context = New-Object -TypeName 'System.DirectoryServices.AccountManagement.PrincipalContext' -Arg ('Domain', $domain)
  }
  catch
  {
    throw "Failed to contact a domain controller for $($domain): $($_.Exception.GetBaseException().Message)"
  }

  # Upon 'not found' error condition, FindByIdentity just returns $null
  $account = $null
  if ($context -ne $null)
  {
    $account = `
      ( `
        [System.DirectoryServices.AccountManagement.ComputerPrincipal]::FindByIdentity($context, 'SamAccountName', $userName), `
        [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($context, 'SamAccountName', $userName) `
      ) `
      | Where { $_ -ne $null }
  }
  else
  {
    $account = $null
  }

  if ($account -ne $null)
    { [System.DirectoryServices.AccountManagement.AuthenticablePrincipal] $account }
  else
    { throw "A user or computer account $qualifiedUserName was not found." }
}

####################
# Helper functions #
####################

# -ErrorVariable doesn't work well with RPS, so we'll have to have our own tracker
# Note, that cmdlets / functions executed with -ErrorAction SilentlyContinue will
# still contribute their errors to this variable.
# Example use / unit test: 
#   RecordErrors { throw "Contoso" }; $criticalErrors
$script:criticalErrors = @()
function RecordErrors([ScriptBlock] $body, [Switch] $exceptionsOnly, [Switch] $reportOutcome)
{
  $errorsBefore = @($error)
  $errorsAfter = $null

  try
  {
    $input | & $body
  }
  # Some exceptions, like from division by 0, cannot be caught
  # The thrown exception has been already captured in $errorsAfter
  catch
  {
    # By the time we enter the catch block, exception is already 
    # added to $error.
    $errorsAfter = @($error)

    # However, it wasn't written out to the host, so do it now.
    Write-Error $_
  }
  finally
  {
    # Don't count the exceptions twice - once when thrown and another 
    # time when output by the Write-Error above.
    if (-not $errorsAfter)
    {
      # We get here only if no exception was caught, which means that
      # for the -ExceptionsOnly mode we'd have to report no problems.
      if ($exceptionsOnly) 
        { $errorsAfter = @() }
      else
        { $errorsAfter = @($error) }
    }

    $newErrors = SubtractSet $errorsAfter $errorsBefore

    if ($newErrors)
      { $script:criticalErrors += $newErrors }

    if ($ReportOutcome)
      { return (-not $newErrors) }
  }
}

function HasCriticalErrors
{
  return [bool]$script:criticalErrors
}

$script:nonCriticalErrors = @()
function RecordNonCriticalError([string]$message)
{
  $script:nonCriticalErrors += $message
  Write-Host -Fore:DarkRed "NON-CRITICAL ERROR: $message"
}

function HasNonCriticalErrors
{
  return [bool]$script:nonCriticalErrors
}

function Confirm([string]$prompt, [string]$title, [bool]$safeAnswer = $false)
{
  if ($Unattended)
  {
    Write-Verbose "Unattended mode: supplying $safeAnswer to the `"$prompt`""
    return $safeAnswer
  }
  else
  {
    return $pscmdlet.ShouldContinue($prompt, $title)
  }
}

function GetCasWithAsa([string]$errorAction = 'Continue', [Switch]$retrievePasswords = $false, [Switch]$executeOnTargetServer = $false)
{
  $ids = @($input)
  Write-Verbose "Retrieving CAS server objects with credentials (passwords=$retrievePasswords):`n`t$($ids -join ', ')"

  if ($retrievePasswords)
    { $executeOnTargetServer = $true }

  # $getCas - a remotable payload that loads CAS object with appropriate credential information.
  # This script block MAY BE sent as-is to the remote runspace, and it should therefore adhere to
  # the rules of the "restricted language" that Exchange PowerShell enforces.
  $getCas = 
    { `
      param($serverId, $retrievePasswords, $verbose) 
      Get-ClientAccessServer $serverId `
        -IncludeAlternateServiceAccountCredentialPassword:$retrievePasswords `
        -IncludeAlternateServiceAccountCredentialStatus:(-not $retrievePasswords) `
        -Verbose:$verbose
    }

  # (ExchangeServer | $getCasInvoke) - a strategy for efficiently calling $getCas payload 
  # for the specified credential retrieval options.
  if ($executeOnTargetServer)
  {
    $getCasInvoke = `
      { `
        Write-Verbose "Retrieving ASA credentials from server $_"
        $session = CreateOrGetExchangeSession $_.Fqdn
        if ($session -ne $null)
        {
          Invoke-Command `
              -Session $session `
              -Arg @($_.Identity, $retrievePasswords, $InnerVerbose) `
              -ScriptBlock $getCas `
              -ErrorAction $errorAction
        }
      }
  }
  else
  {
    $getCasInvoke = { & $getCas $_.Identity $retrievePasswords $InnerVerbose }
  }

  RecordErrors `
    {
      # Make sure we're not executing nested pipelines - we'll get a "nested pipeline" RPS error
      # if ForEach will try to execute a block that calls into its own runspace (the $executeOnTargetServer=$false) case.
      ($ids `
        | Get-ExchangeServer `
        | Where { $_.IsClientAccessServer }) `
        | Write-OperationProgress "Connecting to and retrieving the current Alternate Service Account configuration from" `
        | ForEach $getCasInvoke
    }
}

function CheckServerVersions([Microsoft.Exchange.Data.Directory.Management.ClientAccessServer[]] $casServers)
{
  # AdminDisplayVersion is not present on ClientAccessServer

  foreach ($server in $casServers | Get-ExchangeServer)
  {
    $ver = $server.AdminDisplayVersion
    $prefix = "Server $server is of version $ver"

    if (($ver.Major -lt 14) `
      -or ($ver.Major -eq 14 -and $ver.Minor -eq 0 -and $ver.Build -lt 689) `
      -or ($ver.Major -eq 14 -and $ver.Minor -eq 1 -and $ver.Build -lt 085))
    {
      Write-Error "$prefix, which doesn't support automated Alternate Service Account password maintenance. Install Exchange 2010 RTM RU2 or later."
    }

    if ($ver.Major -gt 14)
    {
      Write-Warning "$prefix, which has not been tested with this script."
    }
  }
}

function GetArrayForServer([Microsoft.Exchange.Data.Directory.Management.ClientAccessServer] $server)
{
  $script:arrays | Where { $_.Members -contains $server.Identity }
}

function CreateTask([string] $taskName)
{
  # Why is this so complicated?! Here are the main reasons:
  # 1)  Windows Scheduled Tasks limit command line length to 261 characters, which is very easy
  #     to trip over if we have to reference Exchange Install path more than once (for addressing
  #     a script and logging) and pass all parameters inline.
  # 2)  Difficult escaping: SchTasks.exe, which is a command-line way to manipulate
  #     scheduled tasks, consumes double-quotes and says that single or escaped quotes should be 
  #     used to build a command line for cmd.exe. Well, the problem is that we need to tell cmd.exe
  #     to execute PowerShell.exe, and we'd like PowerShell.exe to execute a script with parameters
  #     that would also need quotes.
  # 3)  We need to do logging, and for that we need to use output redirection. However, using >>
  #     on the cmd.exe command line is going to make logging happen at the cmd.exe level, and that eats
  #     all line feeds.
  # 4)  We simply want to execute a script with exactly the same arguments. However, it's not that
  #     easy: $MyInvocation.Line will contain the exact line that called into us, which means that
  #     it will be invalid outside of its own scope, since it may reference local variables.
  #     $MyInvocation.BoundParameters is good, but it's a dictionary, which we'd have to "splice"
  #     into the script invocate command during scheduled runs, and so types of the values are 
  #     to strictly match types of the parameters, which makes stringization not an option. 
  #
  # Based on what was said above, we really need some intermediate file to store a lengthy cmdline in, 
  # and we also need to have a way to store the serialized parameters' dictionary. Luckily, we can
  # get by with creating just one intermediate file: .cmd files allow long command lines, and so we
  # can invoke PowerShell.exe with -EncodedCommand option, which contains a parameters dictionary, 
  # a script invocation command and a logging support.
  #
  # Considering that two distribution models are supported - downloadable for RTM RU2+ and shipped with SP1, 
  # we really don't know a good place to put our generated files into, and so we'll just pick a 
  # directory from which the script itself is being executed - at least it's discoverable and accessible by
  # different users, just in case an admin decides to change a run-as account for the task we created.

  # So, here's the action plan:
  # 1. Get arguments to this script's execution
  # 2. Slightly tweak them to avoid task re-creation and ensure the unattended mode
  # 3. Serialize arguments into CliXml and save them into a string
  # 4. Escape the string and make it a part of the PowerShell.exe command line
  # 5. Add script invocation and logging
  # 6. Encode this PowerShell command sequence into Base64
  # 7. Build a PowerShell.exe command line with an encoded command
  # 8. Build a SchTasks.exe command line and execute it

  function EncodeCommand([string] $command)
  {
    return [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($command))
  }

  $inv = $script:MyInvocation
  $scriptFile = Get-Item $inv.MyCommand.Definition
  $batchContent = @()

  # Task names can contain "\" for folders.
  $escapedTaskName = $taskName -replace '\\', '_'

  # Figure out logging - the standard logging directory is a good option, 
  # although it may have permissions problems. We'll make sure it's not
  # the case by creating a directory right away.
  # Make a name of a logging directory out of a script name, by stripping everything
  # that follows the first dot: Foo.ps1 => Foo
  $logDir = Join-Path $script:exchangePath "Logging\$(@($scriptFile.Name.Split('.'))[0])"
  $logFileTemplate = Join-Path $logDir "$($escapedTaskName)_{0}.log"
  $rawLogFile = { $logFileTemplate -f "" }
  $processedLogFile = { $logFileTemplate -f [DateTime]::UtcNow.ToString("yyyyMMdd-HHMMss") }
  Write-Verbose "Logs for the scheduled task will be stored in $logDir"

  # Create a logging directory to check the current user's permissions.
  if (-not (Test-Path $logDir))
    { $null = MkDir $logDir -ErrorAction Stop }

  # .. make sure the directory doesn't vanish by the time we get to writing anything
  # meaningful into it.
  $batchContent += "md `"$logDir`""
  
  # Augment parameters and serialize them.
  # .. copy the parameters dictionary
  $rollAsaArgs = @{ } + $inv.BoundParameters

  # .. tweak the command-line by adding the mandatory parameters (any requests for interaction
  # will cause a script to fail when PowerShell is started in -NonInteractive mode) and removing
  # the prohibited ones.
  $rollAsaArgs.Unattended = $true
  $rollAsaArgs.Remove('CreateScheduledTask')

  # .. serialize & save. XML still needs to be escaped when made a part of the PowerShell command
  # sequence. However, using single-quoted strings leaves ' as the only character that needs escaping.
  $argsFileName = [System.IO.Path]::GetTempFileName()
  $rollAsaArgs | Export-CliXml $argsFileName -ErrorAction Stop
  $rollAsaEscapedArgs = (Get-Content $argsFileName) -replace "'", "''"
  Remove-Item $argsFileName

  # Create a command to deserialize and supply the arguments.
  # 2>&1 means "redirect errors to the standard output stream".
  # @ in @a means parameter "splicing" - binding command arguments from a dictionary.
  $rollAsaCommand = "    
    `$fn = [System.IO.Path]::GetTempFileName(); 
    Set-Content -Path `$fn -Value ('$rollAsaEscapedArgs' -replace `"''`", `"'`");
    `$a = Import-CliXml `$fn;
    Remove-Item `$fn;
    & '$($scriptFile.FullName)' @a  2>&1"
  Write-Verbose "PowerShell password maintenance command that will be executed periodically:`r`n$rollAsaCommand`r`n"

  # This command will load a log file and normalize a bunch of `r`n and `n line-feeds used by PowerShell
  # to the uniform `r`n understood by Notepad.
  $formatLogCommand = "
    `$logFileTemplate = `"$logFileTemplate`";
    `$s = (Get-Content ($rawLogFile)) -join `"`r`n`";
    `$s = (((`$s -replace `"`r`n`", `"`n`") -replace `"`r`", `"`") -replace `"`n`", `"`r`n`");
    New-Item ($processedLogFile) -Type file -Value `$s"
  Write-Verbose "PowerShell log maintenance command that will be executed periodically:`r`n$formatLogCommand`r`n"

  # Encode a command
  $batchContent += "powershell.exe -version 2.0 -NoProfile -NonInteractive -EncodedCommand $(EncodeCommand $rollAsaCommand) > `"$(& $rawLogFile)`""
  $batchContent += "powershell.exe -version 2.0 -NoProfile -NonInteractive -EncodedCommand $(EncodeCommand $formatLogCommand)"
  $batchContentString = $batchContent -join "`r`n"
  # $batchContentString is too big and bizzare-looking to dump it into verbose. Its content will be in a batch file anyway.

  # Save the batch content to a file right next to the script itself.
  $cmdFile = New-Item `
    -Force `
    -Type file `
    -Path (Join-Path $scriptFile.Directory "$escapedTaskName.cmd") `
    -Value $batchContentString `
    -ErrorAction Stop
  Write-Verbose "Created `"$($cmdFile.FullName)`" that contains commands for scheduled execution."

  # Create, force (overwrite), execute even when not logged on, scheduled once every 3 weeks.
  # A user will be able to change any script parameters later, except name (Task Scheduler limitation).
  # Execution frequency can be modified using Windows Task Scheduler.
  #
  # If you wish to customize the script to schedule as a specific user, append /RU `"$env:UserDomain\$env:UserName`" /RP
  # If that is the case you may want to show this prompt:
  # Write-Host -Fore:White "The task will be set up to execute in a background as a currently logged-on user. Please, type in your password:"
  # and issue a prompt for a password that the util you call would wait for.

  $schedulerCommand = "SchTasks.exe /Create /F /SC WEEKLY /MO 3 /TN `"$taskName`" /TR `"'$($cmdFile.Fullname)'`""
  Write-Verbose "Task Scheduler command-line that will be executed as a scheduled task:`r`n$schedulerCommand`r`n"
  Invoke-Expression $schedulerCommand

  # This will print detailed properties of the task we created in a Format-List-like format.
  & SchTasks.exe /Query /TN "$taskName" /FO LIST /V | Out-Host
}

function ChangePasswordInAD([System.Management.Automation.PSCredential] $newCredential, [Microsoft.Exchange.Data.Directory.ADObjectId] $sampleServerId)
{
  # Takes no input.
  # Provides no output.
  # All error conditions are conveyed through either exceptions or logged as non-critical errors.

  Write-Host -Fore:DarkGreen "Preparing to update Active Directory with a new password for $($newCredential.UserName) ..."

  $account = GetAccount $newCredential.UserName
  [string]$newPassword = $newCredential.GetNetworkCredential().Password
  [string]$oldPassword = $null
  [Microsoft.Exchange.Data.Directory.Management.ClientAccessServer]$serverWithCreds = $null

  # We have two options:
  # 1)  .ChangePassword(): only supported on User accounts (BCL limitation?), but requires nothing 
  #     but the old password to succeed. A very nice option permission-wise, but is a subject to
  #     password policies, with "minimum password age" being hit the most during initial configuration of ASA.
  #
  # 2)  .SetPassword(): resets the password. Doesn't need anything but proper permissions to succeed.
  #
  # We'll try to use the former if at all possible. If not - use the latter.

  # Try getting old working creds from a sample server. 
  # First, get all creds from a sample server.
  if ($account -is [System.DirectoryServices.AccountManagement.UserPrincipal])
  {
    if ($sampleServerId)
    {
      # If we got this far without critical errors, then we must have at least one server.
      $serverWithCreds = $sampleServerId | GetCasWithAsa -RetrievePasswords
    }
    else
    {
      RecordNonCriticalError "Password change operation has just failed for all servers. Can't pick a sample server to get the current working credentials from."
    }
  }

  if ($serverWithCreds)
  {
    # Get all credentials for this account from a sample server.
    $allRelevantCreds = @($serverWithCreds `
      | ForEach { $_.AlternateServiceAccountConfiguration.EffectiveCredentials } `
      | Where { $_.QualifiedUserName -eq $newCredential.UserName })

    # After credential clean-up and update with new credentials, there should be
    # exactly two credentials, one working and one not-yet-set-in-AD. The latter
    # should be the oldest of those two (by timestamps, which is how they're ordered). 
    # If not, bail on trying to use ChangePassword().
    switch ($allRelevantCreds.Count)
    {
      1         { Write-Host -Fore:DarkYellow "No working credentials are known to Exchange yet." }
      2         {
                  if ($allRelevantCreds[0].Credential.GetNetworkCredential().Password -eq $newPassword)
                    { $oldPassword = $allRelevantCreds[1].Credential.GetNetworkCredential().Password }
                  else
                    { RecordNonCriticalError "Credentials are ordered in an unexpected way." }
                }
      default   { RecordNonCriticalError "Found $($allRelevantCreds.Count) credentials. Expected 2." }
    }
  }

  # Try changing a password first, if we found the old password
  if ($oldPassword)
  {
    try 
    {
      Write-Host -Fore:DarkGreen "Changing a password in the Active Directory for $($newCredential.UserName) ..."
      $account.ChangePassword($oldPassword, $newPassword)
    }
    catch 
    {
      RecordNonCriticalError "Failed to change a password by supplying an old (current working) password: $_"
      $oldPassword = $null
    }
  }

  # If not, or if changing a password failed (password minimum age requirement is the usual cause), reset it. 
  # This requires more privileges.
  if (-not $oldPassword)
  {
    Write-Host -Fore:DarkGreen "Resetting a password in the Active Directory for $($newCredential.UserName) ..."
    $account.SetPassword($newPassword)
  }

  Write-Host -Fore:DarkGreen "New password was successfully set to Active Directory."
}

function LogEvent([int]$eventId, [string]$entryType, [string]$message)
{
  # Write-EventLog only wraps a single overload of EventLog.WriteEvent(), 
  # which takes an ID of a message and a single argument. Sad, weird, but
  # doens't quite justify going through using CLR directly. That's why
  # all our message templates would have a single insertion point.
  #
  # If a script was downloaded to RTM RU2+, there will be error messages
  # in the EventLog saying message templates cannot be found. However, the
  # strings we insert will be still visible, and the event types will
  # be reflected correctly.
  Write-EventLog `
    -Log "Application" `
    -Source "MSExchange Management Application" `
    -Event $eventId `
    -EntryType $entryType `
    -Category 14 `
    -Message $message
}

function LogStarting
{
  LogEvent 14001 'Information' $script:MyInvocation.Line
}

function LogFinished([bool]$success, [ScriptBlock]$finalStats)
{
  # Creates a string that looks like this on output:
  #  Header
  # ========
  #  Message[0]
  #  Message[1]
  function MessageWithHeader([string]$header, [string[]]$message)
  {
    $header = " $header "
    $hr = [string]::Concat((1..$header.Length | ForEach { "=" }))
    return (("", $header, $hr) + $message + @("")) -join "`r`n"
  }

  $eventId = $null
  $eventType = $null

  if (-not $success) 
    { $eventId = 14004; $eventType = 'Error' }
  elseif ((HasCriticalErrors) -or (HasNonCriticalErrors))
    { $eventid = 14003; $eventType = 'Warning' }
  else
    { $eventid = 14002; $eventType = 'Information' }

  $message = $null
  if (HasCriticalErrors) 
    { $message += MessageWithHeader "Critical errors:" $script:criticalErrors }
  if (HasNonCriticalErrors) 
    { $message += MessageWithHeader "Non-critical errors:" $script:nonCriticalErrors }

  $message += MessageWithHeader "Configuration at the time of script completion:" (& $finalStats)

  LogEvent $eventId $eventType $message 
}

function PrintAsaConfigPerAccount
{
  $script:servers `
    | ForEach { $_.AlternateServiceAccountConfiguration.EffectiveCredentials } `
    | Where { $_ -ne $null } `
    | ForEach { $_.QualifiedUserName } `
    | Select-Unique `
    | ForEach {
      $qualifiedUserName = $_
      GetAccount $qualifiedUserName $false `
        | Where { $_ -ne $null } `
        | Add-Member -Type NoteProperty -Name QualifiedUserName -Value $qualifiedUserName -PassThru } `
    | Format-Table -AutoSize -Wrap `
      StructuralObjectClass, `
      QualifiedUserName, `
      @{ Label='Last Pwd Update'; Expression={ $_.LastPasswordSet.ToLocalTime() }}, `
      @{ Label='SPNs'; Expression={ $_.ServicePrincipalNames | Format-Wide -c 1 | out-string }}
}

function PrintAsaConfigPerServer
{
  $script:servers `
    | Select *, @{ Name='Array'; Expression={ (GetArrayForServer $_).Fqdn }} `
    | Format-Table -AutoSize -Wrap Identity, AlternateServiceAccount* -Group Array `
}

#########################################
# Parameter checking / mode setup logic #
#########################################

function Body
{
  #
  # Load credentials for all affected servers and open sessions to them. 
  # Yes, that could be a lot, but that way we'll at least know where it might fail - consistency is more important.
  #

  Write-Verbose "Preparing the destination ..."

  if ($pscmdlet.ParameterSetName -like 'Server-*')
  {
    $serversInput = $Identity
  }
  else
  {
    RecordErrors `
      {
        if ($pscmdlet.ParameterSetName -like 'Forest-*')
        {
          $script:arrays = Get-ClientAccessArray

          if ($script:arrays -eq $null)
            { throw "No Client Access Arrays are defined in the current forest" }
        }
        else
        {
          $script:arrays = $Identity | Get-ClientAccessArray
        }
      }

    if ($script:arrays -ne $null)
    {
      $serversInput = $script:arrays `
        | Where { $_.Members.Count -gt 0 } `
        | ForEach { $_.Members }
    }
  }

  # Get the list of servers and also force opening connections to all of them
  # to hit any possible errors before we start writing configuration back.
  # First, check if $serversInput is not null: otherwise, GetCasWithAsa will return all servers.
  Write-Verbose "Destination server identities: $serversInput"
  [Microsoft.Exchange.Data.Directory.Management.ClientAccessServer[]]$script:servers = @()
  if ($serversInput)
    { $script:servers = @($serversInput | GetCasWithAsa -ExecuteOnTargetServer) }
  Write-Verbose "Destination servers: $servers"

  Write-Verbose "Checking version requirements for the destination servers ..."
  RecordErrors { CheckServerVersions $script:servers }

  #
  # Figure out the credentials to push
  #

  Write-Verbose "Preparing the credential source ..."
  if ($pscmdlet.ParameterSetName -like '*Generate')
  {
    # Verify that the account exists. If not - hard-fail.
    $null = GetAccount $GenerateNewPasswordFor

    [System.Management.Automation.PSCredential[]]$script:credentialsToSetToAD = @(GenerateNewPassword $GenerateNewPasswordFor)
    [System.Management.Automation.PSCredential[]]$script:credentialsToSetToCas = $script:credentialsToSetToAD
    [bool]$script:removeAllExistingCredentials = $false
  }
  elseif ($pscmdlet.ParameterSetName -like '*CopyFrom')
  {
    [Microsoft.Exchange.Data.Directory.Management.ClientAccessServer[]]$copyFromServers = @(Get-ClientAccessServer $copyFrom)
    if ($copyFromServers.Count -ne 1)
      { throw "-CopyFrom parameter should be set to identity that resolves to exactly one Client Access Server. Got: [$copyFromServers]." }

    [System.Management.Automation.PSCredential[]]$script:credentialsToSetToAD = @()
    [System.Management.Automation.PSCredential[]]$script:credentialsToSetToCas = @( `
      $copyFromServers `
        | GetCasWithAsa -ErrorAction Stop -RetrievePasswords `
        | ForEach { $_.AlternateServiceAccountConfiguration.EffectiveCredentials } `
        | ForEach { $_.Credential })
    [bool]$script:removeAllExistingCredentials = $true
  }
  else
    { throw "Unrecognized parameter set." }

  if (-not $script:servers -or $script:servers.Count -eq 0)
  {
    throw "Couldn't figure out valid servers from the specified destination scope. Check your parameters and try again."
  }

  Write-Host -Fore:DarkGreen "Destination servers that will be updated:"
  $script:servers | Out-Host

  Write-Host -Fore:DarkGreen "Credentials that will be pushed to every server in the specified scope (recent first):"
  $script:credentialsToSetToCas | Out-Host

  # It's okay for $credentialsToSetToCas to have no elements if everything went well: 
  # maybe the source server didn't have any. However, if there were errors and it
  # seems like there's nothing to do, then the script will bail out.
  if ($script:credentialsToSetToCas.Count -eq 0)
  {
    if (HasCriticalErrors)
    {
      throw "No credentials to push to destination servers. The script cannot continue. Check script parameters and errors output above."
    }
    elseif (-not (Confirm `
        "The source server doesn't have any credentials. All credentials on destination servers will therefore be removed." `
        "Error detected"))
    {
      return $false
    }
  }

  if ($removeAllExistingCredentials)
    { Write-Host -Fore:Yellow "Prior to pushing new credentials, all existing credentials will be removed from the destination servers." }
  else
    { Write-Host -Fore:Yellow "Prior to pushing new credentials, all existing credentials that are invalid or no longer work will be removed from the destination servers." }

  if ( `
    (HasCriticalErrors) `
    -and (-not $WhatIfPreference) `
    -and (-not (Confirm "Errors were detected during the preparation stage. Do you want to continue and make changes?" "Error detected")))
    { return $false }

  ####################################
  # Configuration modification logic #
  ####################################

  #
  # Push credentials to destination servers
  #

  [Microsoft.Exchange.Data.Directory.Management.ClientAccessServer[]]$successfulServers = @()

  foreach ($server in $script:servers)
  {
    if ($pscmdlet.ShouldProcess($server.Name) -or $WhatIfPreference)
    {
      Write-Host -Fore:DarkGreen "Pushing credentials to server $server"

      # The script block passed to Invoke-Command will be executed remotely
      # in a runspace with restricted language. Only a very limited subset
      # of language constructs is allowed there.
      if (RecordErrors -ReportOutcome { 
        $session = CreateOrGetExchangeSession $server.Fqdn
        Invoke-Command `
          -Session $session `
          -Arg ($server.Identity, $credentialsToSetToCas, $removeAllExistingCredentials, $WhatIfPreference, $InnerVerbose) `
          -ScriptBlock {
            param($serverId, $creds, $shouldRemoveAll, $whatIf, $verbose)
            Set-ClientAccessServer $serverId `
              -RemoveAlternateServiceAccountCredentials:$shouldRemoveAll `
              -CleanUpInvalidAlternateServiceAccountCredentials:(-not $shouldRemoveAll) `
              -AlternateServiceAccountCredential $creds `
              -Verbose:$verbose `
              -WhatIf:$whatIf }
        })
      {
        $successfulServers += $server
      }
    }
  }

  #
  # Push new passwords to Active Directory
  #

  if ($script:credentialsToSetToAD.Count -gt 0)
  {
    Write-Host -Fore:Green "Setting a new password on Alternate Serice Account in Active Directory"

    if (HasCriticalErrors)
    {
      Write-Host -Fore:Red "Errors were encountered while propagating credentials to Exchange servers. Pushing a password change to Active Directory may cause an authentication outage"
      Write-Host -Fore:DarkRed $script:criticalErrors
    }

    $isSafeToPushCreds = -not (HasCriticalErrors)
    [System.Management.Automation.PSCredential[]]$successfulCreds = @()

    foreach ($cred in $script:credentialsToSetToAD)
    {
      if ( `
        -not $WhatIfPreference `
        -and (Confirm `
          "Do you want to change password for $($cred.UserName) in Active Directory at this time?" `
          "Password change" `
          -SafeAnswer $isSafeToPushCreds) `
        -and (RecordErrors -ReportOutcome -ExceptionsOnly {
          ChangePasswordInAD $cred $successfulServers[0].Identity }))
      {
        $successfulCreds += $cred
      }
    }
  }

  return `
    $successfulServers `
    -and (($pscmdlet.ParameterSetName -notlike 'Server-*') -or -not (SubtractSet $script:servers $successfulServers)) `
    -and -not (SubtractSet $script:credentialsToSetToAD $successfulCreds)
}

###############
# ENTRY POINT #
###############

$InnerVerbose = $DebugPreference -ne 'SilentlyContinue'
# Strict mode is good for verification of this script, but breaks when it has to
# dot-source other non-strict scripts, like RemoteFunctions.ps1.
#Set-StrictMode -Version Latest

Write-Host -Fore:DarkYellow "`r`n`r`n`r`n========== Starting at $([DateTime]::Now) =========="
Write-Verbose ("Effective parameters that were passed to this script:`r`n" + (($script:MyInvocation.BoundParameters | Out-String) -replace "`r`n", ""))

# 
# Verify that we're running in the Exchange Management Shell and if not, load it.
#

$setupRegKeyPath = "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup"
if (test-path $setupRegKeyPath)
  { $script:exchangePath = (Get-ItemProperty $setupRegKeyPath).MsiInstallPath }
else
  { throw "Cannot find registry settings for Microsoft Exchange" }

# 
# Create a scheduled task with a command, if that's what we were asked to do.
# Scheduled task creation doesn't really need anything except for $script:exchangePath.
#

if ($CreateScheduledTask)
{
  if ($Unattended)
  {
    throw "Unattended mode is not supported when creating a scheduled task. Scheduled tasks themselves will always run in the unattended mode. Remove the parameter and retry."
  }

  Write-Host -Fore:Green "Creating a scheduled task"
  CreateTask $CreateScheduledTask
  return
}

#
# Ensure that Exchange Management Shell is loaded.
#

# This block MUST be at the script level, or dot-sourcing will not import
# functions into the script scope.
Write-Verbose "Examining the state of the local runspace ..."
if (-not (Get-Command Connect-ExchangeServer -ErrorAction SilentlyContinue))
{
  Write-Verbose "Exchange Remote Shell was not imported into the current environment. Loading ..."

  . (Join-Path $script:exchangePath "bin\RemoteExchange.ps1")
  Connect-ExchangeServer -Auto
}

#
# Check that the running user has RBAC access to the necessary commands and parameters
#

$getCas = Get-Command Get-ClientAccessServer
if ($getCas -eq $null -or $getCas.Parameters.IncludeAlternateServiceAccountCredentialPassword -eq $null)
{
  throw "The running user doesn't have enough privileges to perform this operation. Verify that he/she has the 'Organization Configuration' role assigned."
}

LogStarting

$script:success = $false

# Catch all exceptions, but print them into the console: 
# we intend to exit the script execution with "exit",
# and that will stop the script execution right away, 
# preventing the exception from bubbling up further and
# thus concealing it from the operator.
RecordErrors -ExceptionsOnly { $script:success = Body }

# 
# Print the final statisics
#

Write-Host -Fore:Green "Retrieving the current Alternate Service Account configuration from servers in scope"
# Update objects
$script:arrays = Get-ClientAccessArray
$script:servers = @($script:servers | GetCasWithAsa)

LogFinished $script:success `
  {
    PrintAsaConfigPerAccount | Out-String
    PrintAsaConfigPerServer | Out-String
  }

Write-Host -Fore:DarkGreen "Alternate Service Account properties:"
PrintAsaConfigPerAccount | Out-Host

Write-Host -Fore:DarkGreen "Per-server Alternate Service Account configuration as of the time of script completion:"
PrintAsaConfigPerServer | Out-Host

Write-Host -Fore:DarkYellow "========== Finished at $([DateTime]::Now) =========="

if ($script:success)
{
  Write-Host -Fore:Green "`r`n`tTHE SCRIPT HAS SUCCEEDED"
  exit 0
}
else
{
  Write-Host -Fore:Red "`r`n`tTHE SCRIPT HAS FAILED"
  exit 1
}

# SIG # Begin signature block
# MIIagwYJKoZIhvcNAQcCoIIadDCCGnACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU9fI/WCTZzX5DmNYaSVun/9/w
# JdegghUmMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
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
# Ok6T/ZsGj62K+Gh+O7p5QpM5SfXQXuxwjUJ1xYJVkBu1VWEwggS6MIIDoqADAgEC
# AgphAo5CAAAAAAAfMA0GCSqGSIb3DQEBBQUAMHcxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQTAeFw0xMjAxMDkyMjI1NThaFw0xMzA0MDkyMjI1NThaMIGzMQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMQ0wCwYDVQQLEwRNT1BSMScwJQYD
# VQQLEx5uQ2lwaGVyIERTRSBFU046RjUyOC0zNzc3LThBNzYxJTAjBgNVBAMTHE1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUAA4IB
# DwAwggEKAoIBAQCW7I5HTVTCXJWA104LPb+XQ8NL42BnES8BTQzY0UYvEEDeC6RQ
# UhKIC0N6LT/uSG5mx5HmA8pu7HmpaiObzWKezWqkP+ejQ/9iR6G0ukT630DBhVR+
# 6KCnLEMjm1IfMjX0/ppWn41jd3swngozhXIbykrIzCXN210RLsewjPGPQ0hHBbV6
# IAvl8+/BuvSz2M04j/shqj0KbYUX0MrnhgPAM4O1JcTMWpzEw9piJU1TJRRhj/sb
# 4Oz3R8aAReY1UyM2d8qw3ZgrOcB1NQ/dgUwhPXYwxbKwZXMpSCfYwtKwhEe7eLrV
# dAPe10sZ91PeeNqG92GIJjO0R8agVIgVKyx1AgMBAAGjggEJMIIBBTAdBgNVHQ4E
# FgQUL+hGyGjTbk+yINDeiU7xR+5IwfIwHwYDVR0jBBgwFoAUIzT42VJGcArtQPt2
# +7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3NvZnQu
# Y29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNybDBY
# BggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9zb2Z0
# LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNVHSUE
# DDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAc/99Lp3NjYgrfH3jXhVx
# 6Whi8Ai2Q1bEXEotaNj5SBGR8xGchewS1FSgdak4oVl/de7G9TTYVKTi0Mx8l6uT
# dTCXBx0EUyw2f3/xQB4Mm4DiEgogOjHAB3Vn4Po0nOyI+1cc5VhiIJBFL11FqciO
# s3xybRAnxUvYb6KoErNtNSNn+izbJS25XbEeBedDKD6cBXZ38SXeBUcZbd5JhaHa
# SksIRiE1qHU2TLezCKrftyvZvipq/d81F8w/DMfdBs9OlCRjIAsuJK5fQ0QSelzd
# N9ukRbOROhJXfeNHxmbTz5xGVvRMB7HgDKrV9tU8ouC11PgcfgRVEGsY9JHNUaeV
# ZTCCBbwwggOkoAMCAQICCmEzJhoAAAAAADEwDQYJKoZIhvcNAQEFBQAwXzETMBEG
# CgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsG
# A1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5MB4XDTEw
# MDgzMTIyMTkzMloXDTIwMDgzMTIyMjkzMloweTELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEjMCEGA1UEAxMaTWljcm9zb2Z0IENvZGUgU2lnbmlu
# ZyBQQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCycllcGTBkvx2a
# YCAgQpl2U2w+G9ZvzMvx6mv+lxYQ4N86dIMaty+gMuz/3sJCTiPVcgDbNVcKicqu
# IEn08GisTUuNpb15S3GbRwfa/SXfnXWIz6pzRH/XgdvzvfI2pMlcRdyvrT3gKGiX
# GqelcnNW8ReU5P01lHKg1nZfHndFg4U4FtBzWwW6Z1KNpbJpL9oZC/6SdCnidi9U
# 3RQwWfjSjWL9y8lfRjFQuScT5EAwz3IpECgixzdOPaAyPZDNoTgGhVxOVoIoKgUy
# t0vXT2Pn0i1i8UU956wIAPZGoZ7RW4wmU+h6qkryRs83PDietHdcpReejcsRj1Y8
# wawJXwPTAgMBAAGjggFeMIIBWjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTL
# EejK0rQWWAHJNy4zFha5TJoKHzALBgNVHQ8EBAMCAYYwEgYJKwYBBAGCNxUBBAUC
# AwEAATAjBgkrBgEEAYI3FQIEFgQU/dExTtMmipXhmGA7qDFvpjy82C0wGQYJKwYB
# BAGCNxQCBAweCgBTAHUAYgBDAEEwHwYDVR0jBBgwFoAUDqyCYEBWJ5flJRP8KuEK
# U5VZ5KQwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC5taWNyb3NvZnQuY29t
# L3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNlcnQuY3JsMFQGCCsGAQUF
# BwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aS9jZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwDQYJKoZIhvcNAQEFBQADggIB
# AFk5Pn8mRq/rb0CxMrVq6w4vbqhJ9+tfde1MOy3XQ60L/svpLTGjI8x8UJiAIV2s
# PS9MuqKoVpzjcLu4tPh5tUly9z7qQX/K4QwXaculnCAt+gtQxFbNLeNK0rxw56gN
# ogOlVuC4iktX8pVCnPHz7+7jhh80PLhWmvBTI4UqpIIck+KUBx3y4k74jKHK6BOl
# kU7IG9KPcpUqcW2bGvgc8FPWZ8wi/1wdzaKMvSeyeWNWRKJRzfnpo1hW3ZsCRUQv
# X/TartSCMm78pJUT5Otp56miLL7IKxAOZY6Z2/Wi+hImCWU4lPF6H0q70eFW6NB4
# lhhcyTUWX92THUmOLb6tNEQc7hAVGgBd3TVbIc6YxwnuhQ6MT20OE049fClInHLR
# 82zKwexwo1eSV32UjaAbSANa98+jZwp0pTbtLS8XyOZyNxL0b7E8Z4L5UrKNMxZl
# Hg6K3RDeZPRvzkbU0xfpecQEtNP7LN8fip6sCvsTJ0Ct5PnhqX9GuwdgR2VgQE6w
# QuxO7bN2edgKNAltHIAxH+IOVN3lofvlRxCtZJj/UBYufL8FIXrilUEnacOTj5XJ
# jdibIa4NXJzwoq6GaIMMai27dmsAHZat8hZ79haDJLmIz2qoRzEvmtzjcT3XAH5i
# R9HOiMm4GPoOco3Boz2vAkBq/2mbluIQqBC0N1AI1sM9MIIGBzCCA++gAwIBAgIK
# YRZoNAAAAAAAHDANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZImiZPyLGQBGRYDY29t
# MRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQDEyRNaWNyb3NvZnQg
# Um9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMDcwNDAzMTI1MzA5WhcNMjEw
# NDAzMTMwMzA5WjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCfoWyx39tIkip8ay4Z4b3i48WZUSNQrc7dGE4k
# D+7Rp9FMrXQwIBHrB9VUlRVJlBtCkq6YXDAm2gBr6Hu97IkHD/cOBJjwicwfyzMk
# h53y9GccLPx754gd6udOo6HBI1PKjfpFzwnQXq/QsEIEovmmbJNn1yjcRlOwhtDl
# KEYuJ6yGT1VSDOQDLPtqkJAwbofzWTCd+n7Wl7PoIZd++NIT8wi3U21StEWQn0gA
# SkdmEScpZqiX5NMGgUqi+YSnEUcUCYKfhO1VeP4Bmh1QCIUAEDBG7bfeI0a7xC1U
# n68eeEExd8yb3zuDk6FhArUdDbH895uyAc4iS1T/+QXDwiALAgMBAAGjggGrMIIB
# pzAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBQjNPjZUkZwCu1A+3b7syuwwzWz
# DzALBgNVHQ8EBAMCAYYwEAYJKwYBBAGCNxUBBAMCAQAwgZgGA1UdIwSBkDCBjYAU
# DqyCYEBWJ5flJRP8KuEKU5VZ5KShY6RhMF8xEzARBgoJkiaJk/IsZAEZFgNjb20x
# GTAXBgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBS
# b290IENlcnRpZmljYXRlIEF1dGhvcml0eYIQea0WoUqgpa1Mc1j0BxMuZTBQBgNV
# HR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9w
# cm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUHAQEESDBGMEQG
# CCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01p
# Y3Jvc29mdFJvb3RDZXJ0LmNydDATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG
# 9w0BAQUFAAOCAgEAEJeKw1wDRDbd6bStd9vOeVFNAbEudHFbbQwTq86+e4+4LtQS
# ooxtYrhXAstOIBNQmd16QOJXu69YmhzhHQGGrLt48ovQ7DsB7uK+jwoFyI1I4vBT
# Fd1Pq5Lk541q1YDB5pTyBi+FA+mRKiQicPv2/OR4mS4N9wficLwYTp2Oawpylbih
# OZxnLcVRDupiXD8WmIsgP+IHGjL5zDFKdjE9K3ILyOpwPf+FChPfwgphjvDXuBfr
# Tot/xTUrXqO/67x9C0J71FNyIe4wyrt4ZVxbARcKFA7S2hSY9Ty5ZlizLS/n+YWG
# zFFW6J1wlGysOUzU9nm/qhh6YinvopspNAZ3GmLJPR5tH4LwC8csu89Ds+X57H21
# 46SodDW4TsVxIxImdgs8UoxxWkZDFLyzs7BNZ8ifQv+AeSGAnhUwZuhCEl4ayJ4i
# IdBD6Svpu/RIzCzU2DKATCYqSCRfWupW76bemZ3KOm+9gSd0BhHudiG/m4LBJ1S2
# sWo9iaF2YbRuoROmv6pH8BJv/YoybLL+31HIjCPJZr2dHYcSZAI9La9Zj7jkIeW1
# sMpjtHhUBdRBLlCslLCleKuzoJZ1GtmShxN1Ii8yqAhuoFuMJb+g74TKIdbrHk/J
# mu5J4PcBZW+JC33Iacjmbuqnl84xKf8OxVtc2E0bodj6L54/LlUWa8kTo/0xggTH
# MIIEwwIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQITMwAAAJ0ejSeu
# uPPYOAABAAAAnTAJBgUrDgMCGgUAoIHpMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3
# AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEW
# BBT+FRN4nKBlWtPhmIn4+FIwfU0h1zCBiAYKKwYBBAGCNwIBDDF6MHigUIBOAFIA
# bwBsAGwAQQBsAHQAZQByAG4AYQB0AGUAUwBlAHIAdgBpAGMAZQBBAGMAYwBvAHUA
# bgB0AFAAYQBzAHMAdwBvAHIAZAAuAHAAcwAxoSSAImh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZIhvcNAQEBBQAEggEAUK9nvZ1Vhyu0Z8n5
# v94xvO3To8MamsbhkF268QXXkGQCWx9R4Cpyg/83sfbRllyOO1mC8k68WcHsBU7r
# M3cVb3WbbWaRpyJN1gjIQlQq9CSDE/d7m5XtS5COD9yvyNf5zeXJDOVbJ2ZjgPDN
# HxQDC7bQHy4nvoCCR7z8zGEiqTAfayKy47UbcCQ4r5KlCxQEa8Fd6pZwgPkdVBiI
# 3C8DrnLWbRWs1uaIYoyHgWU0YkEfG0YXE5Ezy4ijRvQGvUv5WBHFv3VnDoBvTohJ
# HOPDJ91m2A+in/VHfeJtp99t9yqqyqc5WF7t7hzY2To/wR52/4TJzFc68/CD1GeZ
# oVS04aGCAh8wggIbBgkqhkiG9w0BCQYxggIMMIICCAIBATCBhTB3MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQg
# VGltZS1TdGFtcCBQQ0ECCmECjkIAAAAAAB8wCQYFKw4DAhoFAKBdMBgGCSqGSIb3
# DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTEzMDIwNTA2MzcyNlow
# IwYJKoZIhvcNAQkEMRYEFK1wrYguKHUF3TTT4Ar1z1h+/3usMA0GCSqGSIb3DQEB
# BQUABIIBAEkga9deRNBH02V19MrcrgCNXuZ/wybvjyjUAcy9AjKMqFs+U+4VtoL8
# u6+HRGu8lCqBW6FrOC3or3d/xZ2ZMpGRJ7nGjkfVKiMDOBwF4jtUotc01bspJayA
# Cy6e/8hEaU/fKheMvWgBLw/235H+TCwIzm/zj+x3ZBbM4NIsAZC66siq+UsNyXfm
# CB/0pPGm2edVxUFAEifUD/R1B0LB0d+cy1wCtvHP2HL/Njf8GTFdi2LyaUxkaKst
# uxFqx1b1n4vt1n6vRwtueKORJzhRiR6oJP5lxU5aNJUhn6FtEBBWF8cY3kgMtPLK
# vTzndZwf+tj80lrSNS5giTgjlV1MdsA=
# SIG # End signature block

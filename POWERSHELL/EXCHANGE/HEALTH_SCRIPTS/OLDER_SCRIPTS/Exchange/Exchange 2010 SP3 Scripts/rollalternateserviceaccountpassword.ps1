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
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU9fI/WCTZzX5DmNYaSVun/9/w
# JdegghhqMIIE2jCCA8KgAwIBAgITMwAAASIn72vt4vugowAAAAABIjANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTgxMDI0MjEwNzQw
# WhcNMjAwMTEwMjEwNzQwWjCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEm
# MCQGA1UECxMdVGhhbGVzIFRTUyBFU046RUFDRS1FMzE2LUM5MUQxJTAjBgNVBAMT
# HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQDRU0uHr30XLzP0oZTW7fCdslb6OXTQeoCa/IpqzTgDcXyf
# EqI0fdvhtqQ84neZE4vwJUAbJ2S+ajJirzzZIEU/JiTZpJgeeAMtN+MuAbzXrdyS
# ohyUDuGkuN+vVSeCnEZkeGcFf/zrNWWXmS7JsVK2BJR8YvXk0sBUbWVpdj0uvz68
# Y+HUyx8AKKE2nHRu54f6fC4eiwP/hs+L7NejJm+sNo7HXV4Y6edQI36FdY0Sotq8
# 7Lh3U96U4O6X9cD0iqKxr4lxYYkh98AzVUjiiSdWUt65DAMbdjBV6cepatwVVoET
# EtNK/f83bMS3sOL00QMWoyQM1F7+fLoz1TF7qlozAgMBAAGjggEJMIIBBTAdBgNV
# HQ4EFgQUtlFVlkkUKuXnuF3JZxfDlHs2paYwHwYDVR0jBBgwFoAUIzT42VJGcArt
# QPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# bDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNV
# HSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAZldsd7vjji5U30Zj
# cKiJvhDtcmx0b4s4s0E7gd8Lp4VnvAQAnpc3SkknUslyknvHGE77OSdxKdrO8qnn
# T0Tymqvf7/Re2xJcRVcM4f8TeE5hCaffCkB7Gtu90R+6+Eb1BnBDYMbj3b42Jq8K
# 42hnDG0ntrgv4/TmyJWIvmGQORWMCWyM/NraY3Ldi7pDpTfx9Z9s4eNE/cxipoST
# XHMIgPgDgbZcuFBANnWwF+/swj69cv87x+Jv/8HM/Naoawrr8+0yDjiJ90OzLGI5
# RScuGfQUlH0ESbzevO/9PFpoUywmNYhHoEPngLJVT2W6y13jFUx3IS9lnR0r1dCh
# mynB8jCCBf8wggPnoAMCAQICEzMAAAFRno2PQHGjDkEAAAAAAVEwDQYJKoZIhvcN
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
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU/hUTeJygZVrT4ZiJ+PhSMH1NIdcw
# XAYKKwYBBAGCNwIBDDFOMEygJIAiAEMASQBUAFMAQwBvAG4AcwB0AGEAbgB0AHMA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBAFYVmPxFcpY8ks1COVtcyjXr1KhiLyiz8DGQ1LpUuT6C
# s5imKDolpBJkrSeXG5Z0mjEqXXULyILoN+KxTL/iim+L1n+MHWUWSNWVkELZbkEO
# EUF4hSJ/csuu2fOu7+vZjy4bBE0rVuJWRGmdVJX/oN3kzNj/7uRMMtTRWKb+zCPt
# apwDTE9w2i7oKoKW3UHhXQfxvCx+MeprKNwiZI/EqL4KmrEJnWp3k1NEtP5MlIBR
# ImCMNywJcGWUAVWsrZNeZri2VTwbgIvLv/J2gks+S51Dfl1f9ff4FQgZwEEd+TJV
# 1Q/nF8UZcxAEdx6FvDTqNstDeBLV3L1hy7xDujR4tVGhggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAB
# Iifva+3i+6CjAAAAAAEiMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTA2MjYxNTI2MDFaMCMGCSqGSIb3DQEJ
# BDEWBBQ1ahG56JbnFrtmM2BDpWm0IhGaujANBgkqhkiG9w0BAQUFAASCAQBRoZkW
# lon0TU1SSWK/SFXxdm+TkfdS78GZO66xpGGM80ebsK26wFw5nO6Al+N/7XWKJFS+
# F1RNPuI0nCKIRjka2TLQmyUf8Fi0tg7UELwouNSAC+3ouyQFZSvaFo8EMYqlI/0M
# JcTn7SR9daXBigF8cnPJMSYa3O8+1jrGcV47TtAgfTYgjZusDOEyqwRPo2Bbxeyp
# moCURp/d7o/clXUOk1o6SM+Q5/vssFWysj2Yz8L6j1sgoDMKmLVU7dOSLD2KqoTK
# TQDTYNyL1drP47gU213WchZrXa6tAm2/nQukvr5LS/hBRGJLtbT2scTxe8cxWQ63
# jX2coDiFtWAS0rI0
# SIG # End signature block

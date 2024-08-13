function Check-Elevation {
    if ($PSVersionTable.PSEdition -eq "Desktop" -or $PSVersionTable.Platform -eq "Win32NT" -or $PSVersionTable.PSVersion.Major -le 5) {
        [System.Security.Principal.WindowsPrincipal]$currentPrincipal = New-Object System.Security.Principal.WindowsPrincipal(
            [System.Security.Principal.WindowsIdentity]::GetCurrent()
        )

        [System.Security.Principal.WindowsBuiltInRole]$administratorsRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator

        if($currentPrincipal.IsInRole($administratorsRole)) {
            return $true
        }
        else {
            return $false
        }
    }
    
    if ($PSVersionTable.Platform -eq "Unix") {
        if ($(whoami) -eq "root") {
            return $true
        }
        else {
            return $false
        }
    }
}

function Get-NativePath {
    [CmdletBinding()]
    Param( 
        [Parameter(Mandatory=$True)]
        [string[]]$PathAsStringArray
    )

    $PathAsStringArray = foreach ($pathPart in $PathAsStringArray) {
        $SplitAttempt = $pathPart -split [regex]::Escape([IO.Path]::DirectorySeparatorChar)
        
        if ($SplitAttempt.Count -gt 1) {
            foreach ($obj in $SplitAttempt) {
                $obj
            }
        }
        else {
            $pathPart
        }
    }
    $PathAsStringArray = $PathAsStringArray -join [IO.Path]::DirectorySeparatorChar

    $PathAsStringArray

}

function Pause-ForWarning {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [int]$PauseTimeInSeconds,

        [Parameter(Mandatory=$True)]
        $Message
    )

    Write-Warning $Message
    Write-Host "To answer in the affirmative, press 'y' on your keyboard."
    Write-Host "To answer in the negative, press any other key on your keyboard, OR wait $PauseTimeInSeconds seconds"

    $timeout = New-Timespan -Seconds ($PauseTimeInSeconds - 1)
    $stopwatch = [diagnostics.stopwatch]::StartNew()
    while ($stopwatch.elapsed -lt $timeout){
        if ([Console]::KeyAvailable) {
            $keypressed = [Console]::ReadKey("NoEcho").Key
            Write-Host "You pressed the `"$keypressed`" key"
            if ($keypressed -eq "y") {
                $Result = $true
                break
            }
            if ($keypressed -ne "y") {
                $Result = $false
                break
            }
        }

        # Check once every 1 second to see if the above "if" condition is satisfied
        Start-Sleep 1
    }

    if (!$Result) {
        $Result = $false
    }
    
    $Result
}

function Unzip-File {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,Position=0)]
        [string]$PathToZip,
        
        [Parameter(Mandatory=$true,Position=1)]
        [string]$TargetDir,

        [Parameter(Mandatory=$false,Position=2)]
        [string[]]$SpecificItem
    )

    if ($PSVersionTable.PSEdition -eq "Core") {
        [System.Collections.ArrayList]$AssembliesToCheckFor = @("System.Console","System","System.IO",
            "System.IO.Compression","System.IO.Compression.Filesystem","System.IO.Compression.ZipFile"
        )

        [System.Collections.ArrayList]$NeededAssemblies = @()

        foreach ($assembly in $AssembliesToCheckFor) {
            try {
                [System.Collections.ArrayList]$Failures = @()
                try {
                    $TestLoad = [System.Reflection.Assembly]::LoadWithPartialName($assembly)
                    if (!$TestLoad) {
                        throw
                    }
                }
                catch {
                    $null = $Failures.Add("Failed LoadWithPartialName")
                }

                try {
                    $null = Invoke-Expression "[$assembly]"
                }
                catch {
                    $null = $Failures.Add("Failed TabComplete Check")
                }

                if ($Failures.Count -gt 1) {
                    $Failures
                    throw
                }
            }
            catch {
                Write-Host "Downloading $assembly..."
                $NewAssemblyDir = "$HOME\Downloads\$assembly"
                $NewAssemblyDllPath = "$NewAssemblyDir\$assembly.dll"
                if (!$(Test-Path $NewAssemblyDir)) {
                    New-Item -ItemType Directory -Path $NewAssemblyDir
                }
                if (Test-Path "$NewAssemblyDir\$assembly*.zip") {
                    Remove-Item "$NewAssemblyDir\$assembly*.zip" -Force
                }
                $OutFileBaseNamePrep = Invoke-WebRequest "https://www.nuget.org/api/v2/package/$assembly" -DisableKeepAlive -UseBasicParsing
                $OutFileBaseName = $($OutFileBaseNamePrep.BaseResponse.ResponseUri.AbsoluteUri -split "/")[-1] -replace "nupkg","zip"
                Invoke-WebRequest -Uri "https://www.nuget.org/api/v2/package/$assembly" -OutFile "$NewAssemblyDir\$OutFileBaseName"
                Expand-Archive -Path "$NewAssemblyDir\$OutFileBaseName" -DestinationPath $NewAssemblyDir

                $PossibleDLLs = Get-ChildItem -Recurse $NewAssemblyDir | Where-Object {$_.Name -eq "$assembly.dll" -and $_.Parent -notmatch "net[0-9]" -and $_.Parent -match "core|standard"}

                if ($PossibleDLLs.Count -gt 1) {
                    Write-Warning "More than one item within $NewAssemblyDir\$OutFileBaseName matches $assembly.dll"
                    Write-Host "Matches include the following:"
                    for ($i=0; $i -lt $PossibleDLLs.Count; $i++){
                        "$i) $($($PossibleDLLs[$i]).FullName)"
                    }
                    $Choice = Read-Host -Prompt "Please enter the number corresponding to the .dll you would like to load [0..$($($PossibleDLLs.Count)-1)]"
                    if ($(0..$($($PossibleDLLs.Count)-1)) -notcontains $Choice) {
                        Write-Error "The number indicated does is not a valid choice! Halting!"
                        $global:FunctionResult = "1"
                        return
                    }

                    if ($PSVersionTable.Platform -eq "Win32NT") {
                        # Install to GAC
                        [System.Reflection.Assembly]::LoadWithPartialName("System.EnterpriseServices")
                        $publish = New-Object System.EnterpriseServices.Internal.Publish
                        $publish.GacInstall($PossibleDLLs[$Choice].FullName)
                    }

                    # Copy it to the root of $NewAssemblyDir\$OutFileBaseName
                    Copy-Item -Path "$($PossibleDLLs[$Choice].FullName)" -Destination "$NewAssemblyDir\$assembly.dll"

                    # Remove everything else that was extracted with Expand-Archive
                    Get-ChildItem -Recurse $NewAssemblyDir | Where-Object {
                        $_.FullName -ne "$NewAssemblyDir\$assembly.dll" -and
                        $_.FullName -ne "$NewAssemblyDir\$OutFileBaseName"
                    } | Remove-Item -Recurse -Force
                    
                }
                if ($PossibleDLLs.Count -lt 1) {
                    Write-Error "No matching .dll files were found within $NewAssemblyDir\$OutFileBaseName ! Halting!"
                    continue
                }
                if ($PossibleDLLs.Count -eq 1) {
                    if ($PSVersionTable.Platform -eq "Win32NT") {
                        # Install to GAC
                        [System.Reflection.Assembly]::LoadWithPartialName("System.EnterpriseServices")
                        $publish = New-Object System.EnterpriseServices.Internal.Publish
                        $publish.GacInstall($PossibleDLLs.FullName)
                    }

                    # Copy it to the root of $NewAssemblyDir\$OutFileBaseName
                    Copy-Item -Path "$($PossibleDLLs[$Choice].FullName)" -Destination "$NewAssemblyDir\$assembly.dll"

                    # Remove everything else that was extracted with Expand-Archive
                    Get-ChildItem -Recurse $NewAssemblyDir | Where-Object {
                        $_.FullName -ne "$NewAssemblyDir\$assembly.dll" -and
                        $_.FullName -ne "$NewAssemblyDir\$OutFileBaseName"
                    } | Remove-Item -Recurse -Force
                }
            }
            $AssemblyFullInfo = [System.Reflection.Assembly]::LoadWithPartialName($assembly)
            if (!$AssemblyFullInfo) {
                $AssemblyFullInfo = [System.Reflection.Assembly]::LoadFile("$NewAssemblyDir\$assembly.dll")
            }
            if (!$AssemblyFullInfo) {
                Write-Error "The assembly $assembly could not be found or otherwise loaded! Halting!"
                $global:FunctionResult = "1"
                return
            }
            $null = $NeededAssemblies.Add([pscustomobject]@{
                AssemblyName = "$assembly"
                Available = if ($AssemblyFullInfo){$true} else {$false}
                AssemblyInfo = $AssemblyFullInfo
                AssemblyLocation = $AssemblyFullInfo.Location
            })
        }

        if ($NeededAssemblies.Available -contains $false) {
            $AssembliesNotFound = $($NeededAssemblies | Where-Object {$_.Available -eq $false}).AssemblyName
            Write-Error "The following assemblies cannot be found:`n$AssembliesNotFound`nHalting!"
            $global:FunctionResult = "1"
            return
        }

        $Assem = $NeededAssemblies.AssemblyInfo.FullName

        $Source = @"
        using System;
        using System.IO;
        using System.IO.Compression;

        namespace MyCore.Utils
        {
            public static class Zip
            {
                public static void ExtractAll(string sourcepath, string destpath)
                {
                    string zipPath = @sourcepath;
                    string extractPath = @destpath;

                    using (ZipArchive archive = ZipFile.Open(zipPath, ZipArchiveMode.Update))
                    {
                        archive.ExtractToDirectory(extractPath);
                    }
                }

                public static void ExtractSpecific(string sourcepath, string destpath, string specificitem)
                {
                    string zipPath = @sourcepath;
                    string extractPath = @destpath;
                    string itemout = @specificitem.Replace(@"\","/");

                    //Console.WriteLine(itemout);

                    using (ZipArchive archive = ZipFile.OpenRead(zipPath))
                    {
                        foreach (ZipArchiveEntry entry in archive.Entries)
                        {
                            //Console.WriteLine(entry.FullName);
                            //bool satisfied = new bool();
                            //satisfied = entry.FullName.IndexOf(@itemout, 0, StringComparison.CurrentCultureIgnoreCase) != -1;
                            //Console.WriteLine(satisfied);

                            if (entry.FullName.IndexOf(@itemout, 0, StringComparison.CurrentCultureIgnoreCase) != -1)
                            {
                                string finaloutputpath = extractPath + "\\" + entry.Name;
                                entry.ExtractToFile(finaloutputpath, true);
                            }
                        }
                    } 
                }
            }
        }
"@

        Add-Type -ReferencedAssemblies $Assem -TypeDefinition $Source

        if (!$SpecificItem) {
            [MyCore.Utils.Zip]::ExtractAll($PathToZip, $TargetDir)
        }
        else {
            [MyCore.Utils.Zip]::ExtractSpecific($PathToZip, $TargetDir, $SpecificItem)
        }
    }


    if ($PSVersionTable.PSEdition -eq "Desktop" -and $($($PSVersionTable.Platform -and $PSVersionTable.Platform -eq "Win32NT") -or !$PSVersionTable.Platform)) {
        if ($SpecificItem) {
            foreach ($item in $SpecificItem) {
                if ($SpecificItem -match "\\") {
                    $SpecificItem = $SpecificItem -replace "\\","\\"
                }
            }
        }

        ##### BEGIN Native Helper Functions #####
        function Get-ZipChildItems {
            [CmdletBinding()]
            Param(
                [Parameter(Mandatory=$false,Position=0)]
                [string]$ZipFile = $(Read-Host -Prompt "Please enter the full path to the zip file")
            )

            $shellapp = new-object -com shell.application
            $zipFileComObj = $shellapp.Namespace($ZipFile)
            $i = $zipFileComObj.Items()
            Get-ZipChildItems_Recurse $i
        }

        function Get-ZipChildItems_Recurse {
            [CmdletBinding()]
            Param(
                [Parameter(Mandatory=$true,Position=0)]
                $items
            )

            foreach($si in $items) {
                if($si.getfolder -ne $null) {
                    # Loop through subfolders 
                    Get-ZipChildItems_Recurse $si.getfolder.items()
                }
                # Spit out the object
                $si
            }
        }

        ##### END Native Helper Functions #####

        ##### BEGIN Variable/Parameter Transforms and PreRun Prep #####
        if (!$(Test-Path $PathToZip)) {
            Write-Verbose "The path $PathToZip was not found! Halting!"
            Write-Error "The path $PathToZip was not found! Halting!"
            $global:FunctionResult = "1"
            return
        }
        if ($(Get-ChildItem $PathToZip).Extension -ne ".zip") {
            Write-Verbose "The file specified by the -PathToZip parameter does not have a .zip file extension! Halting!"
            Write-Error "The file specified by the -PathToZip parameter does not have a .zip file extension! Halting!"
            $global:FunctionResult = "1"
            return
        }

        $ZipFileNameWExt = $(Get-ChildItem $PathToZip).Name

        ##### END Variable/Parameter Transforms and PreRun Prep #####

        ##### BEGIN Main Body #####

        Write-Verbose "NOTE: PowerShell 5.0 uses Expand-Archive cmdlet to unzip files"

        if (!$SpecificItem) {
            if ($PSVersionTable.PSVersion.Major -ge 5) {
                Expand-Archive -Path $PathToZip -DestinationPath $TargetDir
            }
            if ($PSVersionTable.PSVersion.Major -lt 5) {
                # Load System.IO.Compression.Filesystem 
                [System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null

                # Unzip file
                [System.IO.Compression.ZipFile]::ExtractToDirectory($PathToZip, $TargetDir)
            }
        }
        if ($SpecificItem) {
            $ZipSubItems = Get-ZipChildItems -ZipFile $PathToZip

            foreach ($searchitem in $SpecificItem) {
                [array]$potentialItems = foreach ($item in $ZipSubItems) {
                    if ($item.Path -match $searchitem) {
                        $item
                    }
                }

                $shell = new-object -com shell.application

                if ($potentialItems.Count -eq 1) {
                    $shell.Namespace($TargetDir).CopyHere($potentialItems[0], 0x14)
                }
                if ($potentialItems.Count -gt 1) {
                    Write-Warning "More than one item within $ZipFileNameWExt matches $searchitem."
                    Write-Host "Matches include the following:"
                    for ($i=0; $i -lt $potentialItems.Count; $i++){
                        "$i) $($($potentialItems[$i]).Path)"
                    }
                    $Choice = Read-Host -Prompt "Please enter the number corresponding to the item you would like to extract [0..$($($potentialItems.Count)-1)]"
                    if ($(0..$($($potentialItems.Count)-1)) -notcontains $Choice) {
                        Write-Warning "The number indicated does is not a valid choice! Skipping $searchitem..."
                        continue
                    }
                    for ($i=0; $i -lt $potentialItems.Count; $i++){
                        $shell.Namespace($TargetDir).CopyHere($potentialItems[$Choice], 0x14)
                    }
                }
                if ($potentialItems.Count -lt 1) {
                    Write-Warning "No items within $ZipFileNameWExt match $searchitem! Skipping..."
                    continue
                }
            }
        }
        ##### END Main Body #####
    }
}

function New-SudoSession {
    [CmdletBinding(DefaultParameterSetName='Supply UserName and Password')]
    Param(
        [Parameter(
            Mandatory=$False,
            ParameterSetName='Supply UserName and Password'
        )]
        [string]$UserName = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name -split "\\")[-1],

        [Parameter(
            Mandatory=$False,
            ParameterSetName='Supply UserName and Password'
        )]
        $Password,

        [Parameter(
            Mandatory=$False,
            ParameterSetName='Supply Credentials'
        )]
        [System.Management.Automation.PSCredential]$Credentials

    )

    ##### BEGIN Variable/Parameter Transforms and PreRun Prep #####

    if ($UserName -and !$Password -and !$Credentials) {
        $Password = Read-Host -Prompt "Please enter the password for $UserName" -AsSecureString
    }

    if ($UserName -and $Password) {
        if ($Password.GetType().FullName -eq "System.String") {
            $Password = ConvertTo-SecureString $Password -AsPlainText -Force
        }
        $Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName, $Password
    }

    $Domain = $(Get-CimInstance -ClassName Win32_ComputerSystem).Domain
    $LocalHostFQDN = "$env:ComputerName.$Domain"

    ##### END Variable/Parameter Transforms and PreRunPrep #####

    ##### BEGIN Main Body #####

    $CredDelRegLocation = "HKLM:\Software\Policies\Microsoft\Windows\CredentialsDelegation"
    $CredDelRegLocationParent = $CredDelRegLocation | Split-Path -Parent
    $AllowFreshValue = "WSMAN/$LocalHostFQDN"
    $tmpFileXmlPrep = [IO.Path]::GetTempFileName()
    $UpdatedtmpFileXmlName = $tmpFileXmlPrep -replace "\.tmp",".xml"
    $tmpFileXml = $UpdatedtmpFileXmlName
    $TranscriptPath = "$HOME\Open-SudoSession_Transcript_$UserName_$(Get-Date -Format MM-dd-yyy_hhmm_tt).txt"

    $WSManGPOTempConfig = @"
-noprofile -WindowStyle Hidden -Command "Start-Transcript -Path $TranscriptPath -Append
try {`$CurrentAllowFreshCredsProperties = Get-ChildItem -Path $CredDelRegLocation | ? {`$_.PSChildName -eq 'AllowFreshCredentials'}} catch {}
try {`$CurrentAllowFreshCredsValues = foreach (`$propNum in `$CurrentAllowFreshCredsProperties) {`$(Get-ItemProperty -Path '$CredDelRegLocation\AllowFreshCredentials').`$propNum}} catch {}

if (!`$(Test-WSMan)) {`$WinRMConfigured = 'false'; winrm quickconfig /force; Start-Sleep -Seconds 5} else {`$WinRMConfigured = 'true'}
try {`$CredSSPServiceSetting = `$(Get-ChildItem WSMan:\localhost\Service\Auth\CredSSP).Value} catch {}
try {`$CredSSPClientSetting = `$(Get-ChildItem WSMan:\localhost\Client\Auth\CredSSP).Value} catch {}
if (`$CredSSPServiceSetting -eq 'false') {Enable-WSManCredSSP -Role Server -Force}
if (`$CredSSPClientSetting -eq 'false') {Enable-WSManCredSSP -DelegateComputer localhost -Role Client -Force}

if (!`$(Test-Path $CredDelRegLocation)) {`$Status = 'CredDelKey DNE'}
if (`$(Test-Path $CredDelRegLocation) -and !`$(Test-Path $CredDelRegLocation\AllowFreshCredentials)) {`$Status = 'AllowFreshCreds DNE'}
if (`$(Test-Path $CredDelRegLocation) -and `$(Test-Path $CredDelRegLocation\AllowFreshCredentials)) {`$Status = 'AllowFreshCreds AlreadyExists'}

if (!`$(Test-Path $CredDelRegLocation)) {New-Item -Path $CredDelRegLocation}
if (`$(Test-Path $CredDelRegLocation) -and !`$(Test-Path $CredDelRegLocation\AllowFreshCredentials)) {New-Item -Path $CredDelRegLocation\AllowFreshCredentials}

if (`$CurrentAllowFreshCredsValues -notcontains '$AllowFreshValue') {Set-ItemProperty -Path $CredDelRegLocation -Name ConcatenateDefaults_AllowFresh -Value `$(`$CurrentAllowFreshCredsProperties.Count+1) -Type DWord; Start-Sleep -Seconds 2; Set-ItemProperty -Path $CredDelRegLocation\AllowFreshCredentials -Name `$(`$CurrentAllowFreshCredsProperties.Count+1) -Value '$AllowFreshValue' -Type String}
New-Variable -Name 'OrigAllowFreshCredsState' -Value `$([pscustomobject][ordered]@{OrigAllowFreshCredsProperties = `$CurrentAllowFreshCredsProperties; OrigAllowFreshCredsValues = `$CurrentAllowFreshCredsValues; Status = `$Status; OrigWSMANConfigStatus = `$WinRMConfigured; OrigWSMANServiceCredSSPSetting = `$CredSSPServiceSetting; OrigWSMANClientCredSSPSetting = `$CredSSPClientSetting; PropertyToRemove = `$(`$CurrentAllowFreshCredsProperties.Count+1)})
`$(Get-Variable -Name 'OrigAllowFreshCredsState' -ValueOnly) | Export-CliXml -Path $tmpFileXml
exit"
"@
    $WSManGPOTempConfigFinal = $WSManGPOTempConfig -replace "`n","; "

    # IMPORTANT NOTE: You CANNOT use the RunAs Verb if UseShellExecute is $false, and you CANNOT use
    # RedirectStandardError or RedirectStandardOutput if UseShellExecute is $true, so we have to write
    # output to a file temporarily
    $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
    $ProcessInfo.FileName = "powershell.exe"
    $ProcessInfo.RedirectStandardError = $false
    $ProcessInfo.RedirectStandardOutput = $false
    $ProcessInfo.UseShellExecute = $true
    $ProcessInfo.Arguments = $WSManGPOTempConfigFinal
    $ProcessInfo.Verb = "RunAs"
    $Process = New-Object System.Diagnostics.Process
    $Process.StartInfo = $ProcessInfo
    $Process.Start() | Out-Null
    $Process.WaitForExit()
    $WSManAndRegStatus = Import-CliXML $tmpFileXml

    $ElevatedPSSession = New-PSSession -Name "ElevatedSessionFor$UserName" -Authentication CredSSP -Credential $Credentials

    New-Variable -Name "NewSessionAndOriginalStatus" -Scope Global -Value $(
        [pscustomobject][ordered]@{
            ElevatedPSSession   = $ElevatedPSSession
            OriginalWSManAndRegistryStatus   = $WSManAndRegStatus
        }
    ) -Force
    
    $(Get-Variable -Name "NewSessionAndOriginalStatus" -ValueOnly)

    # Cleanup 
    Remove-Item $tmpFileXml

    ##### END Main Body #####

}

function Remove-SudoSession {
    [CmdletBinding(DefaultParameterSetName='Supply UserName and Password')]
    Param(
        [Parameter(
            Mandatory=$False,
            ParameterSetName='Supply UserName and Password'
        )]
        [string]$UserName = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name -split "\\")[-1],

        [Parameter(
            Mandatory=$False,
            ParameterSetName='Supply UserName and Password'
        )]
        $Password,

        [Parameter(
            Mandatory=$False,
            ParameterSetName='Supply Credentials'
        )]
        [System.Management.Automation.PSCredential]$Credentials,

        [Parameter(Mandatory=$True)]
        $OriginalConfigInfo = $(Get-Variable -Name "NewSessionAndOriginalStatus" -ValueOnly).OriginalWSManAndRegistryStatus,

        [Parameter(
            Mandatory=$True,
            ValueFromPipeline=$true,
            Position=0
        )]
        [System.Management.Automation.Runspaces.PSSession]$SessionToRemove

    )

    ##### BEGIN Variable/Parameter Transforms and PreRun Prep #####

    if ($OriginalConfigInfo -eq $null) {
        Write-Warning "Unable to determine the original configuration of WinRM/WSMan and AllowFreshCredentials Registry prior to using New-SudoSession. No configuration changes will be made/reverted."
        Write-Warning "The only action will be removing the Elevated PSSession specified by the -SessionToRemove parameter."
    }

    if ($UserName -and !$Password -and !$Credentials -and $OriginalConfigInfo -ne $null) {
        $Password = Read-Host -Prompt "Please enter the password for $UserName" -AsSecureString
    }

    if ($UserName -and $Password) {
        if ($Password.GetType().FullName -eq "System.String") {
            $Password = ConvertTo-SecureString $Password -AsPlainText -Force
        }
        $Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName, $Password
    }

    $Domain = $(Get-CimInstance -ClassName Win32_ComputerSystem).Domain
    $LocalHostFQDN = "$env:ComputerName.$Domain"

    ##### END Variable/Parameter Transforms and PreRunPrep #####

    ##### BEGIN Main Body #####

    if ($OriginalConfigInfo -ne $null) {
        $CredDelRegLocation = "HKLM:\Software\Policies\Microsoft\Windows\CredentialsDelegation"
        $CredDelRegLocationParent = $CredDelRegLocation | Split-Path -Parent
        $AllowFreshValue = "WSMAN/$LocalHostFQDN"
        $tmpFileXmlPrep = [IO.Path]::GetTempFileName()
        $UpdatedtmpFileXmlName = $tmpFileXmlPrep -replace "\.tmp",".xml"
        $tmpFileXml = $UpdatedtmpFileXmlName
        $TranscriptPath = "$HOME\Remove-SudoSession_Transcript_$UserName_$(Get-Date -Format MM-dd-yyy_hhmm_tt).txt"

        $WSManGPORevertConfig = @"
-noprofile -WindowStyle Hidden -Command "Start-Transcript -Path $TranscriptPath -Append
if ('$($OriginalConfigInfo.Status)' -eq 'CredDelKey DNE') {Remove-Item -Recurse $CredDelRegLocation -Force}
if ('$($OriginalConfigInfo.Status)' -eq 'AllowFreshCreds DNE') {Remove-Item -Recurse $CredDelRegLocation\AllowFreshCredentials -Force}
if ('$($OriginalConfigInfo.Status)' -eq 'AllowFreshCreds AlreadyExists') {Remove-ItemProperty $CredDelRegLocation\AllowFreshCredentials\AllowFreshCredentials -Name $($WSManAndRegStatus.PropertyToRemove) -Force}
if ('$($OriginalConfigInfo.OrigWSMANConfigStatus)' -eq 'false') {Stop-Service -Name WinRm; Set-Service WinRM -StartupType "Manual"}
if ('$($OriginalConfigInfo.OrigWSMANServiceCredSSPSetting)' -eq 'false') {Set-Item -Path WSMan:\localhost\Service\Auth\CredSSP -Value `$false}
if ('$($OriginalConfigInfo.OrigWSMANClientCredSSPSetting)' -eq 'false') {Set-Item -Path WSMan:\localhost\Client\Auth\CredSSP -Value `$false}
exit"
"@
        $WSManGPORevertConfigFinal = $WSManGPORevertConfig -replace "`n","; "

        $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
        $ProcessInfo.FileName = "powershell.exe"
        $ProcessInfo.RedirectStandardError = $false
        $ProcessInfo.RedirectStandardOutput = $false
        $ProcessInfo.UseShellExecute = $true
        $ProcessInfo.Arguments = $WSManGPORevertConfigFinal
        $ProcessInfo.Verb = "RunAs"
        $Process = New-Object System.Diagnostics.Process
        $Process.StartInfo = $ProcessInfo
        $Process.Start() | Out-Null
        $Process.WaitForExit()

    }

    Remove-PSSession $SessionToRemove

    ##### END Main Body #####

}

function Update-PackageManagement {
    [CmdletBinding()]
    Param( 
        [Parameter(Mandatory=$False)]
        [switch]$UseChocolatey,

        [Parameter(Mandatory=$False)]
        [switch]$InstallNuGetCmdLine
    )


    ##### BEGIN Variable/Parameter Transforms and PreRun Prep #####

    # We're going to need Elevated privileges for some commands below, so might as well try to set this up now.
    if (!$(Check-Elevation)) {
        Write-Error "The Update-PackageManagement function must be run with elevated privileges. Halting!"
        $global:FunctionResult = "1"
        return
    }

    if (!$([Environment]::Is64BitProcess)) {
        Write-Error "You are currently running the 32-bit version of PowerShell. Please run the 64-bit version found under C:\Windows\SysWOW64\WindowsPowerShell\v1.0 and try again. Halting!"
        $global:FunctionResult = "1"
        return
    }

    if ($PSVersionTable.PSEdition -eq "Core" -and $PSVersionTable.Platform -ne "Win32NT" -and $UseChocolatey) {
        Write-Error "The Chocolatey Repo should only be added on a Windows OS! Halting!"
        $global:FunctionResult = "1"
        return
    }

    if ($InstallNuGetCmdLine -and !$UseChocolatey) {
        if ($PSVersionTable.PSEdition -eq "Desktop" -or $PSVersionTable.PSVersion.Major -le 5) {                
            $WarningMessage = "NuGet Command Line Tool cannot be installed without using Chocolatey. Would you like to use the Chocolatey Package Provider (NOTE: This is NOT an installation of the chocolatey command line)?"
            $WarningResponse = Pause-ForWarning -PauseTimeInSeconds 15 -Message $WarningMessage
            if ($WarningResponse) {
                $UseChocolatey = $true
            }
        }
        elseif ($PSVersionTable.PSEdition -eq "Core" -and $PSVersionTable.Platform -eq "Win32NT") {
            $WarningMessage = "NuGet Command Line Tool cannot be installed without using Chocolatey. Would you like to install Chocolatey Command Line Tools in order to install NuGet Command Line Tools?"
            $WarningResponse = Pause-ForWarning -PauseTimeInSeconds 15 -Message $WarningMessage
            if ($WarningResponse) {
                $UseChocolatey = $true
            }
        }
        elseif ($PSVersionTable.PSEdition -eq "Core" -and $PSVersionTable.Platform -eq "Unix") {
            $WarningMessage = "The NuGet Command Line Tools binary nuget.exe can be downloaded, but will not be able to be run without Mono. Do you want to download the latest stable nuget.exe?"
            $WarningResponse = Pause-ForWarning -PauseTimeInSeconds 15 -Message $WarningMessage
            if ($WarningResponse) {
                Write-Host "Downloading latest stable nuget.exe..."
                $OutFilePath = Get-NativePath -PathAsStringArray @($HOME, "Downloads", "nuget.exe")
                Invoke-WebRequest -Uri "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe" -OutFile $OutFilePath
            }
            $UseChocolatey = $false
        }
    }

    if ($PSVersionTable.PSEdition -eq "Desktop" -or $PSVersionTable.PSVersion.Major -le 5) {
        # Check to see if we're behind a proxy
        if ([System.Net.WebProxy]::GetDefaultProxy().Address -ne $null) {
            $ProxyAddress = [System.Net.WebProxy]::GetDefaultProxy().Address
            [system.net.webrequest]::defaultwebproxy = New-Object system.net.webproxy($ProxyAddress)
            [system.net.webrequest]::defaultwebproxy.credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
            [system.net.webrequest]::defaultwebproxy.BypassProxyOnLocal = $true
        }
    }
    # TODO: Figure out how to identify default proxy on PowerShell Core...

    ##### END Variable/Parameter Transforms and PreRun Prep #####


    if ($PSVersionTable.PSVersion.Major -lt 5) {
        if ($(Get-Module -ListAvailable).Name -notcontains "PackageManagement") {
            Write-Host "Downloading PackageManagement .msi installer..."
            $OutFilePath = Get-NativePath -PathAsStringArray @($HOME, "Downloads", "PackageManagement_x64.msi")
            Invoke-WebRequest -Uri "https://download.microsoft.com/download/C/4/1/C41378D4-7F41-4BBE-9D0D-0E4F98585C61/PackageManagement_x64.msi" -OutFile $OutFilePath
            
            $DateStamp = Get-Date -Format yyyyMMddTHHmmss
            $MSIFullPath = $OutFilePath
            $MSIParentDir = $MSIFullPath | Split-Path -Parent
            $MSIFileName = $MSIFullPath | Split-Path -Leaf
            $MSIFileNameOnly = $MSIFileName -replace "\.msi",""
            $logFile = Get-NativePath -PathAsStringArray @($MSIParentDir, "$MSIFileNameOnly$DateStamp.log")
            $MSIArguments = @(
                "/i"
                $MSIFullPath
                "/qn"
                "/norestart"
                "/L*v"
                $logFile
            )
            # Install PowerShell Core
            Start-Process "msiexec.exe" -ArgumentList $MSIArguments -Wait -NoNewWindow
        }
        while ($($(Get-Module -ListAvailable).Name -notcontains "PackageManagement") -and $($(Get-Module -ListAvailable).Name -notcontains "PowerShellGet")) {
            Write-Host "Waiting for PackageManagement and PowerShellGet Modules to become available"
            Start-Sleep -Seconds 1
        }
        Write-Host "PackageManagement and PowerShellGet Modules are ready. Continuing..."
    }

    # Set LatestLocallyAvailable variables...
    $PackageManagementLatestLocallyAvailableVersion = $($(Get-Module -ListAvailable | Where-Object {$_.Name -eq "PackageManagement"}).Version | Measure-Object -Maximum).Maximum
    $PowerShellGetLatestLocallyAvailableVersion = $($(Get-Module -ListAvailable | Where-Object {$_.Name -eq "PowerShellGet"}).Version | Measure-Object -Maximum).Maximum

    if ($(Get-Module).Name -notcontains "PackageManagement") {
        Import-Module "PackageManagement" -RequiredVersion $PackageManagementLatestLocallyAvailableVersion
    }
    if ($(Get-Module).Name -notcontains "PowerShellGet") {
        Import-Module "PowerShellGet" -RequiredVersion $PowerShellGetLatestLocallyAvailableVersion
    }

    if ($(Get-Module -Name PackageManagement).ExportedCommands.Count -eq 0 -or
        $(Get-Module -Name PowerShellGet).ExportedCommands.Count -eq 0
    ) {
        Write-Warning "Either PowerShellGet or PackagementManagement Modules were not able to be loaded Imported successfully due to an update initiated within the current session. Please close this PowerShell Session, open a new one, and run this function again."

        $Result = [pscustomobject][ordered]@{
            PackageManagementUpdated  = $false
            PowerShellGetUpdated      = $false
            NewPSSessionRequired      = $true
        }

        $Result
        return
    }

    # Determine if the NuGet Package Provider is available. If not, install it, because it needs it for some reason
    # that is currently not clear to me. Point is, if it's not installed it will prompt you to install it, so just
    # do it beforehand.
    if ($(Get-PackageProvider).Name -notcontains "NuGet") {
        Install-PackageProvider "NuGet" -Scope CurrentUser -Force
        Register-PackageSource -Name 'nuget.org' -Location 'https://api.nuget.org/v3/index.json' -ProviderName NuGet -Trusted -Force -ForceBootstrap
    }

    if ($UseChocolatey) {
        if ($PSVersionTable.PSEdition -eq "Desktop" -or $PSVersionTable.PSVersion.Major -le 5) {
            # Install the Chocolatey Package Provider to be used with PowerShellGet
            if ($(Get-PackageProvider).Name -notcontains "Chocolatey") {
                Install-PackageProvider "Chocolatey" -Scope CurrentUser -Force
                # The above Install-PackageProvider "Chocolatey" -Force DOES register a PackageSource Repository, so we need to trust it:
                Set-PackageSource -Name Chocolatey -Trusted

                # Make sure packages installed via Chocolatey PackageProvider are part of $env:Path
                [System.Collections.ArrayList]$ChocolateyPathsPrep = @()
                [System.Collections.ArrayList]$ChocolateyPathsFinal = @()
                $env:ChocolateyPSProviderPath = "C:\Chocolatey"

                if (Test-Path $env:ChocolateyPSProviderPath) {
                    if (Test-Path "$env:ChocolateyPSProviderPath\lib") {
                        $OtherChocolateyPathsToAdd = $(Get-ChildItem "$env:ChocolateyPSProviderPath\lib" -Directory | foreach {
                            Get-ChildItem $_.FullName -Recurse -File
                        } | foreach {
                            if ($_.Extension -eq ".exe") {
                                $_.Directory.FullName
                            }
                        }) | foreach {
                            $null = $ChocolateyPathsPrep.Add($_)
                        }
                    }
                    if (Test-Path "$env:ChocolateyPSProviderPath\bin") {
                        $OtherChocolateyPathsToAdd = $(Get-ChildItem "$env:ChocolateyPSProviderPath\bin" -Directory | foreach {
                            Get-ChildItem $_.FullName -Recurse -File
                        } | foreach {
                            if ($_.Extension -eq ".exe") {
                                $_.Directory.FullName
                            }
                        }) | foreach {
                            $null = $ChocolateyPathsPrep.Add($_)
                        }
                    }
                }
                
                if ($ChocolateyPathsPrep) {
                    foreach ($ChocoPath in $ChocolateyPathsPrep) {
                        if ($(Test-Path $ChocoPath) -and $OriginalEnvPathArray -notcontains $ChocoPath) {
                            $null = $ChocolateyPathsFinal.Add($ChocoPath)
                        }
                    }
                }
            
                try {
                    $ChocolateyPathsFinal = $ChocolateyPathsFinal | Sort-Object | Get-Unique
                }
                catch {
                    [System.Collections.ArrayList]$ChocolateyPathsFinal = @($ChocolateyPathsFinal)
                }
                if ($ChocolateyPathsFinal.Count -ne 0) {
                    $ChocolateyPathsAsString = $ChocolateyPathsFinal -join ";"
                }

                foreach ($ChocPath in $ChocolateyPathsFinal) {
                    if ($($env:Path -split ";") -notcontains $ChocPath) {
                        if ($env:Path[-1] -eq ";") {
                            $env:Path = "$env:Path$ChocPath"
                        }
                        else {
                            $env:Path = "$env:Path;$ChocPath"
                        }
                    }
                }

                Write-Host "Updated `$env:Path is:`n$env:Path"

                if ($InstallNuGetCmdLine) {
                    # Next, install the NuGet CLI using the Chocolatey Repo
                    try {
                        Write-Host "Trying to find Chocolatey Package Nuget.CommandLine..."
                        while (!$(Find-Package Nuget.CommandLine)) {
                            Write-Host "Trying to find Chocolatey Package Nuget.CommandLine..."
                            Start-Sleep -Seconds 2
                        }
                        
                        Get-Package NuGet.CommandLine -ErrorAction SilentlyContinue
                        if (!$?) {
                            throw
                        }
                    } 
                    catch {
                        Install-Package Nuget.CommandLine -Source chocolatey -Force
                    }
                    
                    # Ensure there's a symlink from C:\Chocolatey\bin to the real NuGet.exe under C:\Chocolatey\lib
                    $NuGetSymlinkTest = Get-ChildItem "C:\Chocolatey\bin" | Where-Object {$_.Name -eq "NuGet.exe" -and $_.LinkType -eq "SymbolicLink"}
                    $RealNuGetPath = $(Resolve-Path "C:\Chocolatey\lib\*\*\NuGet.exe").Path
                    $TestRealNuGetPath = Test-Path $RealNuGetPath
                    if (!$NuGetSymlinkTest -and $TestRealNuGetPath) {
                        New-Item -Path "C:\Chocolatey\bin\NuGet.exe" -ItemType SymbolicLink -Value $RealNuGetPath
                    }
                }
            }
        }
        if ($PSVersionTable.PSEdition -eq "Core" -and $PSVersionTable.Platform -eq "Win32NT") {
            # Install the Chocolatey Command line
            if (!$(Get-Command choco -ErrorAction SilentlyContinue)) {
                # Suppressing all errors for Chocolatey cmdline install. They will only be a problem if
                # there is a Web Proxy between you and the Internet
                $env:chocolateyUseWindowsCompression = 'true'
                $null = Invoke-Expression $([System.Net.WebClient]::new()).DownloadString("https://chocolatey.org/install.ps1") -ErrorVariable ChocolateyInstallProblems 2>&1 6>&1
                $DateStamp = Get-Date -Format yyyyMMddTHHmmss
                $ChocolateyInstallLogFile = Get-NativePath -PathAsStringArray @($(Get-Location).Path, "ChocolateyInstallLog_$DateStamp.txt")
                $ChocolateyInstallProblems | Out-File $ChocolateyInstallLogFile
            }

            if ($InstallNuGetCmdLine) {
                if (!$(Get-Command choco -ErrorAction SilentlyContinue)) {
                    Write-Error "Unable to find chocolatey.exe, however, it should be installed. Please check your System PATH and `$env:Path and try again. Halting!"
                    $global:FunctionResult = "1"
                    return
                }
                else {
                    # 'choco update' aka 'cup' will update if already installed or install if not installed 
                    Start-Process "cup" -ArgumentList "nuget.commandline -y" -Wait -NoNewWindow
                }
                # NOTE: The chocolatey install should take care of setting $env:Path and System PATH so that
                # choco binaries and packages installed via chocolatey can be found here:
                # C:\ProgramData\chocolatey\bin
            }
        }
    }
    # Next, set the PSGallery PowerShellGet PackageProvider Source to Trusted
    if ($(Get-PackageSource | Where-Object {$_.Name -eq "PSGallery"}).IsTrusted -eq $False) {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    }

    # Next, update PackageManagement and PowerShellGet where possible
    [version]$MinimumVer = "1.0.0.1"
    $PackageManagementLatestVersion = $(Find-Module PackageManagement).Version
    $PowerShellGetLatestVersion = $(Find-Module PowerShellGet).Version
    Write-Host "PackageManagement Latest Version is: $PackageManagementLatestVersion"
    Write-Host "PowerShellGetLatestVersion Latest Version is: $PowerShellGetLatestVersion"

    if ($PackageManagementLatestVersion -gt $PackageManagementLatestLocallyAvailableVersion -and $PackageManagementLatestVersion -gt $MinimumVer) {
        if ($PSVersionTable.PSVersion.Major -lt 5) {
            Write-Host "`nUnable to update the PackageManagement Module beyond $($MinimumVer.ToString()) on PowerShell versions lower than 5."
        }
        if ($PSVersionTable.PSVersion.Major -ge 5) {
            #Install-Module -Name "PackageManagement" -Scope CurrentUser -Repository PSGallery -RequiredVersion $PowerShellGetLatestVersion -Force -WarningAction "SilentlyContinue"
            #Install-Module -Name "PackageManagement" -Scope CurrentUser -Repository PSGallery -RequiredVersion $PackageManagementLatestVersion -Force
            Write-Host "Installing latest version of PackageManagement..."
            Install-Module -Name "PackageManagement" -Force
            $PackageManagementUpdated = $True
        }
    }
    if ($PowerShellGetLatestVersion -gt $PowerShellGetLatestLocallyAvailableVersion -and $PowerShellGetLatestVersion -gt $MinimumVer) {
        # Unless the force parameter is used, Install-Module will halt with a warning saying the 1.0.0.1 is already installed
        # and it will not update it.
        Write-Host "Installing latest version of PowerShellGet..."
        #Install-Module -Name "PowerShellGet" -Scope CurrentUser -Repository PSGallery -RequiredVersion $PowerShellGetLatestVersion -Force -WarningAction "SilentlyContinue"
        #Install-Module -Name "PowerShellGet" -RequiredVersion $PowerShellGetLatestVersion -Force
        Install-Module -Name "PowerShellGet" -Force
        $PowerShellGetUpdated = $True
    }

    # Reset the LatestLocallyAvailable variables, and then load them into the current session
    $PackageManagementLatestLocallyAvailableVersion = $($(Get-Module -ListAvailable | Where-Object {$_.Name -eq "PackageManagement"}).Version | Measure-Object -Maximum).Maximum
    $PowerShellGetLatestLocallyAvailableVersion = $($(Get-Module -ListAvailable | Where-Object {$_.Name -eq "PowerShellGet"}).Version | Measure-Object -Maximum).Maximum
    Write-Host "Latest locally available PackageManagement version is $PackageManagementLatestLocallyAvailableVersion"
    Write-Host "Latest locally available PowerShellGet version is $PowerShellGetLatestLocallyAvailableVersion"

    $CurrentlyLoadedPackageManagementVersion = $(Get-Module | Where-Object {$_.Name -eq 'PackageManagement'}).Version
    $CurrentlyLoadedPowerShellGetVersion = $(Get-Module | Where-Object {$_.Name -eq 'PowerShellGet'}).Version
    Write-Host "Currently loaded PackageManagement version is $CurrentlyLoadedPackageManagementVersion"
    Write-Host "Currently loaded PowerShellGet version is $CurrentlyLoadedPowerShellGetVersion"

    if ($CurrentlyLoadedPackageManagementVersion -lt $PackageManagementLatestLocallyAvailableVersion) {
        # Need to remove PowerShellGet first since it depends on PackageManagement
        Write-Host "Removing Module PowerShellGet $CurrentlyLoadedPowerShellGetVersion ..."
        Remove-Module -Name "PowerShellGet"
        Write-Host "Removing Module PackageManagement $CurrentlyLoadedPackageManagementVersion ..."
        Remove-Module -Name "PackageManagement"
    
        if ($(Get-Host).Name -ne "Package Manager Host") {
            Write-Host "We are NOT in the Visual Studio Package Management Console. Continuing..."
            
            # Need to Import PackageManagement first since it's a dependency for PowerShellGet
            # Need to use -RequiredVersion parameter because older versions are still intalled side-by-side with new
            Write-Host "Importing PackageManagement Version $PackageManagementLatestLocallyAvailableVersion ..."
            $null = Import-Module "PackageManagement" -RequiredVersion $PackageManagementLatestLocallyAvailableVersion -ErrorVariable ImportPackManProblems 2>&1 6>&1
            Write-Host "Importing PowerShellGet Version $PowerShellGetLatestLocallyAvailableVersion ..."
            $null = Import-Module "PowerShellGet" -RequiredVersion $PowerShellGetLatestLocallyAvailableVersion -ErrorVariable ImportPSGetProblems 2>&1 6>&1
        }
        if ($(Get-Host).Name -eq "Package Manager Host") {
            Write-Host "We ARE in the Visual Studio Package Management Console. Continuing..."
    
            # Need to Import PackageManagement first since it's a dependency for PowerShellGet
            # Need to use -RequiredVersion parameter because older versions are still intalled side-by-side with new
            Write-Host "Importing PackageManagement Version $PackageManagementLatestLocallyAvailableVersion`nNOTE: Module Members will have with Prefix 'PackMan' - Example: Get-PackManPackage"
            $null = Import-Module "PackageManagement" -RequiredVersion $PackageManagementLatestLocallyAvailableVersion -Prefix PackMan -ErrorVariable ImportPackManProblems 2>&1 6>&1
            Write-Host "Importing PowerShellGet Version $PowerShellGetLatestLocallyAvailableVersion`nNOTE: Module Members will have with Prefix 'PSGet' - Example: Find-PSGetModule"
            $null = Import-Module "PowerShellGet" -RequiredVersion $PowerShellGetLatestLocallyAvailableVersion -Prefix PSGet -ErrorVariable ImportPSGetProblems 2>&1 6>&1
        }
    }
    
    # Reset CurrentlyLoaded Variables
    $CurrentlyLoadedPackageManagementVersion = $(Get-Module | Where-Object {$_.Name -eq 'PackageManagement'}).Version
    $CurrentlyLoadedPowerShellGetVersion = $(Get-Module | Where-Object {$_.Name -eq 'PowerShellGet'}).Version
    Write-Host "Currently loaded PackageManagement version is $CurrentlyLoadedPackageManagementVersion"
    Write-Host "Currently loaded PowerShellGet version is $CurrentlyLoadedPowerShellGetVersion"
    
    if ($CurrentlyLoadedPowerShellGetVersion -lt $PowerShellGetLatestLocallyAvailableVersion) {
        if (!$ImportPSGetProblems) {
            Write-Host "Removing Module PowerShellGet $CurrentlyLoadedPowerShellGetVersion ..."
        }
        Remove-Module -Name "PowerShellGet"
    
        if ($(Get-Host).Name -ne "Package Manager Host") {
            Write-Host "We are NOT in the Visual Studio Package Management Console. Continuing..."
            
            # Need to use -RequiredVersion parameter because older versions are still intalled side-by-side with new
            Write-Host "Importing PowerShellGet Version $PowerShellGetLatestLocallyAvailableVersion ..."
            Import-Module "PowerShellGet" -RequiredVersion $PowerShellGetLatestLocallyAvailableVersion
        }
        if ($(Get-Host).Name -eq "Package Manager Host") {
            Write-Host "We ARE in the Visual Studio Package Management Console. Continuing..."
    
            # Need to use -RequiredVersion parameter because older versions are still intalled side-by-side with new
            Write-Host "Importing PowerShellGet Version $PowerShellGetLatestLocallyAvailableVersion`nNOTE: Module Members will have with Prefix 'PSGet' - Example: Find-PSGetModule"
            Import-Module "PowerShellGet" -RequiredVersion $PowerShellGetLatestLocallyAvailableVersion -Prefix PSGet
        }
    }

    # Make sure all Repos Are Trusted
    if ($UseChocolatey -and $($PSVersionTable.PSEdition -eq "Desktop" -or $PSVersionTable.PSVersion.Major -le 5)) {
        $BaselineRepoNames = @("Chocolatey","nuget.org","PSGallery")
    }
    else {
        $BaselineRepoNames = @("nuget.org","PSGallery")
    }
    if ($(Get-Module -Name PackageManagement).ExportedCommands.Count -gt 0) {
        $RepoObjectsForTrustCheck = Get-PackageSource | Where-Object {$_.Name -match "$($BaselineRepoNames -join "|")"}
    
        foreach ($RepoObject in $RepoObjectsForTrustCheck) {
            if ($RepoObject.IsTrusted -ne $true) {
                Set-PackageSource -Name $RepoObject.Name -Trusted
            }
        }
    }

    # Reset CurrentlyLoaded Variables
    $CurrentlyLoadedPackageManagementVersion = $(Get-Module | Where-Object {$_.Name -eq 'PackageManagement'}).Version
    $CurrentlyLoadedPowerShellGetVersion = $(Get-Module | Where-Object {$_.Name -eq 'PowerShellGet'}).Version
    Write-Host "The FINAL loaded PackageManagement version is $CurrentlyLoadedPackageManagementVersion"
    Write-Host "The FINAL loaded PowerShellGet version is $CurrentlyLoadedPowerShellGetVersion"

    #$ErrorsArrayReversed = $($Error.Count-1)..$($Error.Count-4) | foreach {$Error[$_]}
    #$CheckForError = try {$ErrorsArrayReversed[0].ToString()} catch {$null}
    if ($($ImportPackManProblems | Out-String) -match "Assembly with same name is already loaded" -or 
        $CurrentlyLoadedPackageManagementVersion -lt $PackageManagementLatestVersion -or
        $(Get-Module -Name PackageManagement).ExportedCommands.Count -eq 0
    ) {
        Write-Warning "The PackageManagement Module has been updated and requires and brand new PowerShell Session. Please close this session, start a new one, and run the function again."
        $NewPSSessionRequired = $true
    }

    $Result = [pscustomobject][ordered]@{
        PackageManagementUpdated  = if ($PackageManagementUpdated) {$true} else {$false}
        PowerShellGetUpdated      = if ($PowerShellGetUpdated) {$true} else {$false}
        NewPSSessionRequired      = if ($NewPSSessionRequired) {$true} else {$false}
    }

    $Result
}

Function Check-InstalledPrograms {

    [CmdletBinding(
        PositionalBinding=$True,
        DefaultParameterSetName='Default Param Set'
    )]
    Param(
        [Parameter(
            Mandatory=$False,
            ParameterSetName='Default Param Set'
        )]
        [string]$ProgramTitleSearchTerm,

        [Parameter(
            Mandatory=$False,
            ParameterSetName='Default Param Set'
        )]
        [string[]]$HostName = $env:COMPUTERNAME,

        [Parameter(
            Mandatory=$False,
            ParameterSetName='Secondary Param Set'
        )]
        [switch]$AllADWindowsComputers

    )

    ##### BEGIN Variable/Parameter Transforms and PreRun Prep #####

    $uninstallWow6432Path = "\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    $uninstallPath = "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"

    $RegPaths = @(
        "HKLM:$uninstallWow6432Path",
        "HKLM:$uninstallPath",
        "HKCU:$uninstallWow6432Path",
        "HKCU:$uninstallPath"
    )
    
    ##### END Variable/Parameter Transforms and PreRun Prep #####

    ##### BEGIN Main Body #####
    # Get a list of Windows Computers from AD
    if ($AllADWindowsComputers) {
        $ComputersArray = $(Get-ADComputer -Filter * -Property * | Where-Object {$_.OperatingSystem -like "*Windows*"}).Name
    }
    else {
        $ComputersArray = $env:COMPUTERNAME
    }

    foreach ($computer in $ComputersArray) {
        if ($computer -eq $env:COMPUTERNAME -or $computer.Split("\.")[0] -eq $env:COMPUTERNAME) {
            try {
                $InstalledPrograms = foreach ($regpath in $RegPaths) {if (Test-Path $regpath) {Get-ItemProperty $regpath}}
                if (!$?) {
                    throw
                }
            }
            catch {
                Write-Warning "Unable to find registry path(s) on $computer. Skipping..."
                continue
            }
        }
        else {
            try {
                $InstalledPrograms = Invoke-Command -ComputerName $computer -ScriptBlock {
                    foreach ($regpath in $RegPaths) {
                        if (Test-Path $regpath) {
                            Get-ItemProperty $regpath
                        }
                    }
                } -ErrorAction SilentlyContinue
                if (!$?) {
                    throw
                }
            }
            catch {
                Write-Warning "Unable to connect to $computer. Skipping..."
                continue
            }
        }

        if ($ProgramTitleSearchTerm) {
            $InstalledPrograms | Where-Object {$_.DisplayName -like "*$ProgramTitleSearchTerm*"}
        }
        else {
            $InstalledPrograms
        }
    }

    ##### END Main Body #####

}

function Set-WindowStyle {
    <#
    .LINK
    https://gist.github.com/jakeballard/11240204
    #>

    [CmdletBinding(DefaultParameterSetName = 'InputObject')]
    param(
        [Parameter(Position = 0, Mandatory = $True, ValueFromPipeline = $True, ValueFromPipelinebyPropertyName = $True)]
        [Object[]] $InputObject,
        [Parameter(Position = 1)]
        [ValidateSet('FORCEMINIMIZE', 'HIDE', 'MAXIMIZE', 'MINIMIZE', 'RESTORE', 'SHOW', 'SHOWDEFAULT', 'SHOWMAXIMIZED', 'SHOWMINIMIZED', 'SHOWMINNOACTIVE', 'SHOWNA', 'SHOWNOACTIVATE', 'SHOWNORMAL')]
        [string] $Style = 'SHOW'
    )

    BEGIN {
        $WindowStates = @{
            'FORCEMINIMIZE'   = 11
            'HIDE'            = 0
            'MAXIMIZE'        = 3
            'MINIMIZE'        = 6
            'RESTORE'         = 9
            'SHOW'            = 5
            'SHOWDEFAULT'     = 10
            'SHOWMAXIMIZED'   = 3
            'SHOWMINIMIZED'   = 2
            'SHOWMINNOACTIVE' = 7
            'SHOWNA'          = 8
            'SHOWNOACTIVATE'  = 4
            'SHOWNORMAL'      = 1
        }

$Win32ShowWindowAsync = Add-Type -MemberDefinition @'
[DllImport("user32.dll")] 
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
'@ -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru
    
    }

    PROCESS {
        foreach ($process in $InputObject) {
            $Win32ShowWindowAsync::ShowWindowAsync($process.MainWindowHandle, $WindowStates[$Style]) | Out-Null
            Write-Verbose ("Set Window Style '{1} on '{0}'" -f $MainWindowHandle, $Style)
        }
    }
}

<#
.Synopsis
    Refactored From: https://gallery.technet.microsoft.com/scriptcenter/PowerShell-Credentials-d44c3cde

    Provides access to Windows CredMan basic functionality for client scripts

    ****************** IMPORTANT ******************
    *
    * If you use this script from the PS console, you 
    * should ALWAYS pass the Target, User and Password
    * parameters using single quotes:
    * 
    *   .\CredMan.ps1 -AddCred -Target 'http://server' -User 'JoeSchmuckatelli' -Pass 'P@55w0rd!'
    * 
    * to prevent PS misinterpreting special characters 
    * you might use as PS reserved characters
    * 
    ****************** IMPORTANT ******************

.Description
    Provides the following API when dot-sourced
    Del-Cred
    Enum-Creds
    Read-Cred
    Write-Cred

    Supports the following cmd-line actions
    AddCred (requires -User, -Pass; -Target is optional)
    DelCred (requires -Target)
    GetCred (requires -Target)
    RunTests (no cmd-line opts)
    ShoCred (optional -All parameter to dump cred objects to console)

.INPUTS
    See function-level notes

.OUTPUTS
      Cmd-line usage: console output relative to success or failure state
      Dot-sourced usage:
      ** Successful Action **
      * Del-Cred   : Int = 0
      * Enum-Cred  : PsUtils.CredMan+Credential[]
      * Read-Cred  : PsUtils.CredMan+Credential
      * Write-Cred : Int = 0
      ** Failure **
      * All API    : Management.Automation.ErrorRecord

.NOTES
    Author: Jim Harrison (jim@isatools.org)
    Date  : 2012/05/20
    Vers  : 1.5

    Updates:
    2012/10/13
            - Fixed a bug where the script would only read, write or delete GENERIC 
            credentials types. 
                - Added #region blocks to clarify internal functionality
                - Added 'CredType' param to specify what sort of credential is to be read, 
                created or deleted (not used for -ShoCred or Enum-Creds)
                - Added 'CredPersist' param to specify how the credential is to be stored;
                only used in Write-Cred
                - Added 'All' param for -ShoCreds to differentiate between creds summary
                list and detailed creds dump
                - Added CRED_FLAGS enum to make the credential struct flags values clearer
                - Improved parameter validation
                - Expanded internal help (used with Get-Help cmdlet)
                - Cmd-line functions better illustrate how to interpret the results when 
                dot-sourcing the script

.PARAMETER AddCred
    Specifies that you wish to add a new credential or update an existing credentials
    -Target, -User and -Pass parameters are required for this action

.PARAMETER Comment
    Specifies the information you wish to place in the credentials comment field

.PARAMETER CredPersist
    Specifies the credentials storage persistence you wish to use
    Valid values are: "SESSION", "LOCAL_MACHINE", "ENTERPRISE"
    NOTE: if not specified, defaults to "ENTERPRISE"
    
.PARAMETER CredType
    Specifies the type of credential object you want to store
    Valid values are: "GENERIC", "DOMAIN_PASSWORD", "DOMAIN_CERTIFICATE",
    "DOMAIN_VISIBLE_PASSWORD", "GENERIC_CERTIFICATE", "DOMAIN_EXTENDED",
    "MAXIMUM", "MAXIMUM_EX"
    NOTE: if not specified, defaults to "GENERIC"
    ****************** IMPORTANT ******************
    *
    * I STRONGLY recommend that you become familiar 
    * with http://msdn.microsoft.com/en-us/library/windows/desktop/aa374788(v=vs.85).aspx
    * before you create new credentials with -CredType other than "GENERIC"
    * 
    ****************** IMPORTANT ******************

.PARAMETER DelCred
    Specifies that you wish to remove an existing credential
    -CredType may be required to remove the correct credential if more than one is
    specified for a target

.PARAMETER GetCred
    Specifies that you wish to retrieve an existing credential
    -CredType may be required to access the correct credential if more than one is
    specified for a target

.PARAMETER Pass
    Specifies the credentials password

.PARAMETER RunTests
    Specifies that you wish to run built-in Win32 CredMan functionality tests

.PARAMETER ShoCred
    Specifies that you wish to retrieve all credential stored for the interactive user
    -All parameter may be used to indicate that you wish to view all credentials properties
    (default display is a summary list)

.PARAMETER Target
    Specifies the authentication target for the specified credentials
    If not specified, the -User information is used

.PARAMETER User
    Specifies the credentials username
    

.LINK
    http://msdn.microsoft.com/en-us/library/windows/desktop/aa374788(v=vs.85).aspx
    http://stackoverflow.com/questions/7162604/get-cached-credentials-in-powershell-from-windows-7-credential-manager
    http://msdn.microsoft.com/en-us/library/windows/desktop/aa374788(v=vs.85).aspx
    http://blogs.msdn.com/b/peerchan/archive/2005/11/01/487834.aspx

.EXAMPLE
    .\CredMan.ps1 -AddCred -Target 'http://aserver' -User 'UserName' -Password 'P@55w0rd!' -Comment 'cuziwanna'
    Stores the credential for 'UserName' with a password of 'P@55w0rd!' for authentication against 'http://aserver' and adds a comment of 'cuziwanna'

.EXAMPLE
    .\CredMan.ps1 -DelCred -Target 'http://aserver' -CredType 'DOMAIN_PASSWORD'
    Removes the credential used for the target 'http://aserver' as credentials type 'DOMAIN_PASSWORD'

.EXAMPLE
    .\CredMan.ps1 -GetCred -Target 'http://aserver'
    Retreives the credential used for the target 'http://aserver'

.EXAMPLE
    .\CredMan.ps1 -ShoCred
    Retrieves a summary list of all credentials stored for the interactive user

.EXAMPLE
    .\CredMan.ps1 -ShoCred -All
    Retrieves a detailed list of all credentials stored for the interactive user

#>
#requires -version 2
function Manage-StoredCredentials {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false)]
        [Switch] $AddCred,

        [Parameter(Mandatory=$false)]
        [Switch]$DelCred,
        
        [Parameter(Mandatory=$false)]
        [Switch]$GetCred,
        
        [Parameter(Mandatory=$false)]
        [Switch]$ShoCred,

        [Parameter(Mandatory=$false)]
        [Switch]$RunTests,
        
        [Parameter(Mandatory=$false)]
        [ValidateLength(1,32767) <# CRED_MAX_GENERIC_TARGET_NAME_LENGTH #>]
        [String]$Target,

        [Parameter(Mandatory=$false)]
        [ValidateLength(1,512) <# CRED_MAX_USERNAME_LENGTH #>]
        [String]$User,

        [Parameter(Mandatory=$false)]
        [ValidateLength(1,512) <# CRED_MAX_CREDENTIAL_BLOB_SIZE #>]
        [String]$Pass,

        [Parameter(Mandatory=$false)]
        [ValidateLength(1,256) <# CRED_MAX_STRING_LENGTH #>]
        [String]$Comment,

        [Parameter(Mandatory=$false)]
        [Switch]$All,

        [Parameter(Mandatory=$false)]
        [ValidateSet("GENERIC","DOMAIN_PASSWORD","DOMAIN_CERTIFICATE","DOMAIN_VISIBLE_PASSWORD",
        "GENERIC_CERTIFICATE","DOMAIN_EXTENDED","MAXIMUM","MAXIMUM_EX")]
        [String]$CredType = "GENERIC",

        [Parameter(Mandatory=$false)]
        [ValidateSet("SESSION","LOCAL_MACHINE","ENTERPRISE")]
        [String]$CredPersist = "ENTERPRISE"
    )

    #region Pinvoke
    #region Inline C#
    [String] $PsCredmanUtils = @"
    using System;
    using System.Runtime.InteropServices;

    namespace PsUtils
    {
        public class CredMan
        {
            #region Imports
            // DllImport derives from System.Runtime.InteropServices
            [DllImport("Advapi32.dll", SetLastError = true, EntryPoint = "CredDeleteW", CharSet = CharSet.Unicode)]
            private static extern bool CredDeleteW([In] string target, [In] CRED_TYPE type, [In] int reservedFlag);

            [DllImport("Advapi32.dll", SetLastError = true, EntryPoint = "CredEnumerateW", CharSet = CharSet.Unicode)]
            private static extern bool CredEnumerateW([In] string Filter, [In] int Flags, out int Count, out IntPtr CredentialPtr);

            [DllImport("Advapi32.dll", SetLastError = true, EntryPoint = "CredFree")]
            private static extern void CredFree([In] IntPtr cred);

            [DllImport("Advapi32.dll", SetLastError = true, EntryPoint = "CredReadW", CharSet = CharSet.Unicode)]
            private static extern bool CredReadW([In] string target, [In] CRED_TYPE type, [In] int reservedFlag, out IntPtr CredentialPtr);

            [DllImport("Advapi32.dll", SetLastError = true, EntryPoint = "CredWriteW", CharSet = CharSet.Unicode)]
            private static extern bool CredWriteW([In] ref Credential userCredential, [In] UInt32 flags);
            #endregion

            #region Fields
            public enum CRED_FLAGS : uint
            {
                NONE = 0x0,
                PROMPT_NOW = 0x2,
                USERNAME_TARGET = 0x4
            }

            public enum CRED_ERRORS : uint
            {
                ERROR_SUCCESS = 0x0,
                ERROR_INVALID_PARAMETER = 0x80070057,
                ERROR_INVALID_FLAGS = 0x800703EC,
                ERROR_NOT_FOUND = 0x80070490,
                ERROR_NO_SUCH_LOGON_SESSION = 0x80070520,
                ERROR_BAD_USERNAME = 0x8007089A
            }

            public enum CRED_PERSIST : uint
            {
                SESSION = 1,
                LOCAL_MACHINE = 2,
                ENTERPRISE = 3
            }

            public enum CRED_TYPE : uint
            {
                GENERIC = 1,
                DOMAIN_PASSWORD = 2,
                DOMAIN_CERTIFICATE = 3,
                DOMAIN_VISIBLE_PASSWORD = 4,
                GENERIC_CERTIFICATE = 5,
                DOMAIN_EXTENDED = 6,
                MAXIMUM = 7,      // Maximum supported cred type
                MAXIMUM_EX = (MAXIMUM + 1000),  // Allow new applications to run on old OSes
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
            public struct Credential
            {
                public CRED_FLAGS Flags;
                public CRED_TYPE Type;
                public string TargetName;
                public string Comment;
                public DateTime LastWritten;
                public UInt32 CredentialBlobSize;
                public string CredentialBlob;
                public CRED_PERSIST Persist;
                public UInt32 AttributeCount;
                public IntPtr Attributes;
                public string TargetAlias;
                public string UserName;
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
            private struct NativeCredential
            {
                public CRED_FLAGS Flags;
                public CRED_TYPE Type;
                public IntPtr TargetName;
                public IntPtr Comment;
                public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
                public UInt32 CredentialBlobSize;
                public IntPtr CredentialBlob;
                public UInt32 Persist;
                public UInt32 AttributeCount;
                public IntPtr Attributes;
                public IntPtr TargetAlias;
                public IntPtr UserName;
            }
            #endregion

            #region Child Class
            private class CriticalCredentialHandle : Microsoft.Win32.SafeHandles.CriticalHandleZeroOrMinusOneIsInvalid
            {
                public CriticalCredentialHandle(IntPtr preexistingHandle)
                {
                    SetHandle(preexistingHandle);
                }

                private Credential XlateNativeCred(IntPtr pCred)
                {
                    NativeCredential ncred = (NativeCredential)Marshal.PtrToStructure(pCred, typeof(NativeCredential));
                    Credential cred = new Credential();
                    cred.Type = ncred.Type;
                    cred.Flags = ncred.Flags;
                    cred.Persist = (CRED_PERSIST)ncred.Persist;

                    long LastWritten = ncred.LastWritten.dwHighDateTime;
                    LastWritten = (LastWritten << 32) + ncred.LastWritten.dwLowDateTime;
                    cred.LastWritten = DateTime.FromFileTime(LastWritten);

                    cred.UserName = Marshal.PtrToStringUni(ncred.UserName);
                    cred.TargetName = Marshal.PtrToStringUni(ncred.TargetName);
                    cred.TargetAlias = Marshal.PtrToStringUni(ncred.TargetAlias);
                    cred.Comment = Marshal.PtrToStringUni(ncred.Comment);
                    cred.CredentialBlobSize = ncred.CredentialBlobSize;
                    if (0 < ncred.CredentialBlobSize)
                    {
                        cred.CredentialBlob = Marshal.PtrToStringUni(ncred.CredentialBlob, (int)ncred.CredentialBlobSize / 2);
                    }
                    return cred;
                }

                public Credential GetCredential()
                {
                    if (IsInvalid)
                    {
                        throw new InvalidOperationException("Invalid CriticalHandle!");
                    }
                    Credential cred = XlateNativeCred(handle);
                    return cred;
                }

                public Credential[] GetCredentials(int count)
                {
                    if (IsInvalid)
                    {
                        throw new InvalidOperationException("Invalid CriticalHandle!");
                    }
                    Credential[] Credentials = new Credential[count];
                    IntPtr pTemp = IntPtr.Zero;
                    for (int inx = 0; inx < count; inx++)
                    {
                        pTemp = Marshal.ReadIntPtr(handle, inx * IntPtr.Size);
                        Credential cred = XlateNativeCred(pTemp);
                        Credentials[inx] = cred;
                    }
                    return Credentials;
                }

                override protected bool ReleaseHandle()
                {
                    if (IsInvalid)
                    {
                        return false;
                    }
                    CredFree(handle);
                    SetHandleAsInvalid();
                    return true;
                }
            }
            #endregion

            #region Custom API
            public static int CredDelete(string target, CRED_TYPE type)
            {
                if (!CredDeleteW(target, type, 0))
                {
                    return Marshal.GetHRForLastWin32Error();
                }
                return 0;
            }

            public static int CredEnum(string Filter, out Credential[] Credentials)
            {
                int count = 0;
                int Flags = 0x0;
                if (string.IsNullOrEmpty(Filter) ||
                    "*" == Filter)
                {
                    Filter = null;
                    if (6 <= Environment.OSVersion.Version.Major)
                    {
                        Flags = 0x1; //CRED_ENUMERATE_ALL_CREDENTIALS; only valid is OS >= Vista
                    }
                }
                IntPtr pCredentials = IntPtr.Zero;
                if (!CredEnumerateW(Filter, Flags, out count, out pCredentials))
                {
                    Credentials = null;
                    return Marshal.GetHRForLastWin32Error(); 
                }
                CriticalCredentialHandle CredHandle = new CriticalCredentialHandle(pCredentials);
                Credentials = CredHandle.GetCredentials(count);
                return 0;
            }

            public static int CredRead(string target, CRED_TYPE type, out Credential Credential)
            {
                IntPtr pCredential = IntPtr.Zero;
                Credential = new Credential();
                if (!CredReadW(target, type, 0, out pCredential))
                {
                    return Marshal.GetHRForLastWin32Error();
                }
                CriticalCredentialHandle CredHandle = new CriticalCredentialHandle(pCredential);
                Credential = CredHandle.GetCredential();
                return 0;
            }

            public static int CredWrite(Credential userCredential)
            {
                if (!CredWriteW(ref userCredential, 0))
                {
                    return Marshal.GetHRForLastWin32Error();
                }
                return 0;
            }

            #endregion

            private static int AddCred()
            {
                Credential Cred = new Credential();
                string Password = "Password";
                Cred.Flags = 0;
                Cred.Type = CRED_TYPE.GENERIC;
                Cred.TargetName = "Target";
                Cred.UserName = "UserName";
                Cred.AttributeCount = 0;
                Cred.Persist = CRED_PERSIST.ENTERPRISE;
                Cred.CredentialBlobSize = (uint)Password.Length;
                Cred.CredentialBlob = Password;
                Cred.Comment = "Comment";
                return CredWrite(Cred);
            }

            private static bool CheckError(string TestName, CRED_ERRORS Rtn)
            {
                switch(Rtn)
                {
                    case CRED_ERRORS.ERROR_SUCCESS:
                        Console.WriteLine(string.Format("'{0}' worked", TestName));
                        return true;
                    case CRED_ERRORS.ERROR_INVALID_FLAGS:
                    case CRED_ERRORS.ERROR_INVALID_PARAMETER:
                    case CRED_ERRORS.ERROR_NO_SUCH_LOGON_SESSION:
                    case CRED_ERRORS.ERROR_NOT_FOUND:
                    case CRED_ERRORS.ERROR_BAD_USERNAME:
                        Console.WriteLine(string.Format("'{0}' failed; {1}.", TestName, Rtn));
                        break;
                    default:
                        Console.WriteLine(string.Format("'{0}' failed; 0x{1}.", TestName, Rtn.ToString("X")));
                        break;
                }
                return false;
            }

            /*
             * Note: the Main() function is primarily for debugging and testing in a Visual 
             * Studio session.  Although it will work from PowerShell, it's not very useful.
             */
            public static void Main()
            {
                Credential[] Creds = null;
                Credential Cred = new Credential();
                int Rtn = 0;

                Console.WriteLine("Testing CredWrite()");
                Rtn = AddCred();
                if (!CheckError("CredWrite", (CRED_ERRORS)Rtn))
                {
                    return;
                }
                Console.WriteLine("Testing CredEnum()");
                Rtn = CredEnum(null, out Creds);
                if (!CheckError("CredEnum", (CRED_ERRORS)Rtn))
                {
                    return;
                }
                Console.WriteLine("Testing CredRead()");
                Rtn = CredRead("Target", CRED_TYPE.GENERIC, out Cred);
                if (!CheckError("CredRead", (CRED_ERRORS)Rtn))
                {
                    return;
                }
                Console.WriteLine("Testing CredDelete()");
                Rtn = CredDelete("Target", CRED_TYPE.GENERIC);
                if (!CheckError("CredDelete", (CRED_ERRORS)Rtn))
                {
                    return;
                }
                Console.WriteLine("Testing CredRead() again");
                Rtn = CredRead("Target", CRED_TYPE.GENERIC, out Cred);
                if (!CheckError("CredRead", (CRED_ERRORS)Rtn))
                {
                    Console.WriteLine("if the error is 'ERROR_NOT_FOUND', this result is OK.");
                }
            }
        }
    }
"@
    #endregion

    $PsCredMan = $null
    try
    {
        $PsCredMan = [PsUtils.CredMan]
    }
    catch
    {
        #only remove the error we generate
        #$Error.RemoveAt($Error.Count-1)
    }
    if($null -eq $PsCredMan)
    {
        Add-Type $PsCredmanUtils
    }
    #endregion

    #region Internal Tools
    [HashTable] $ErrorCategory = @{0x80070057 = "InvalidArgument";
                                   0x800703EC = "InvalidData";
                                   0x80070490 = "ObjectNotFound";
                                   0x80070520 = "SecurityError";
                                   0x8007089A = "SecurityError"}

    function Get-CredType {
        Param (
            [Parameter(Mandatory=$true)]
            [ValidateSet("GENERIC","DOMAIN_PASSWORD","DOMAIN_CERTIFICATE","DOMAIN_VISIBLE_PASSWORD",
            "GENERIC_CERTIFICATE","DOMAIN_EXTENDED","MAXIMUM","MAXIMUM_EX")]
            [String]$CredType
        )
        
        switch($CredType) {
            "GENERIC" {return [PsUtils.CredMan+CRED_TYPE]::GENERIC}
            "DOMAIN_PASSWORD" {return [PsUtils.CredMan+CRED_TYPE]::DOMAIN_PASSWORD}
            "DOMAIN_CERTIFICATE" {return [PsUtils.CredMan+CRED_TYPE]::DOMAIN_CERTIFICATE}
            "DOMAIN_VISIBLE_PASSWORD" {return [PsUtils.CredMan+CRED_TYPE]::DOMAIN_VISIBLE_PASSWORD}
            "GENERIC_CERTIFICATE" {return [PsUtils.CredMan+CRED_TYPE]::GENERIC_CERTIFICATE}
            "DOMAIN_EXTENDED" {return [PsUtils.CredMan+CRED_TYPE]::DOMAIN_EXTENDED}
            "MAXIMUM" {return [PsUtils.CredMan+CRED_TYPE]::MAXIMUM}
            "MAXIMUM_EX" {return [PsUtils.CredMan+CRED_TYPE]::MAXIMUM_EX}
        }
    }

    function Get-CredPersist {
        Param (
            [Parameter(Mandatory=$true)]
            [ValidateSet("SESSION","LOCAL_MACHINE","ENTERPRISE")]
            [String] $CredPersist
        )
        
        switch($CredPersist) {
            "SESSION" {return [PsUtils.CredMan+CRED_PERSIST]::SESSION}
            "LOCAL_MACHINE" {return [PsUtils.CredMan+CRED_PERSIST]::LOCAL_MACHINE}
            "ENTERPRISE" {return [PsUtils.CredMan+CRED_PERSIST]::ENTERPRISE}
        }
    }
    #endregion

    #region Dot-Sourced API
    function Del-Creds {
        <#
        .Synopsis
            Deletes the specified credentials

        .Description
            Calls Win32 CredDeleteW via [PsUtils.CredMan]::CredDelete

        .INPUTS
            See function-level notes

        .OUTPUTS
            0 or non-0 according to action success
            [Management.Automation.ErrorRecord] if error encountered

        .PARAMETER Target
            Specifies the URI for which the credentials are associated
          
        .PARAMETER CredType
            Specifies the desired credentials type; defaults to 
            "CRED_TYPE_GENERIC"
        #>

        Param (
            [Parameter(Mandatory=$true)]
            [ValidateLength(1,32767)]
            [String] $Target,

            [Parameter(Mandatory=$false)]
            [ValidateSet("GENERIC","DOMAIN_PASSWORD","DOMAIN_CERTIFICATE","DOMAIN_VISIBLE_PASSWORD",
            "GENERIC_CERTIFICATE","DOMAIN_EXTENDED","MAXIMUM","MAXIMUM_EX")]
            [String] $CredType = "GENERIC"
        )
        
        [Int]$Results = 0
        try {
            $Results = [PsUtils.CredMan]::CredDelete($Target, $(Get-CredType $CredType))
        }
        catch {
            return $_
        }
        if(0 -ne $Results) {
            [String]$Msg = "Failed to delete credentials store for target '$Target'"
            [Management.ManagementException] $MgmtException = New-Object Management.ManagementException($Msg)
            [Management.Automation.ErrorRecord] $ErrRcd = New-Object Management.Automation.ErrorRecord($MgmtException, $Results.ToString("X"), $ErrorCategory[$Results], $null)
            return $ErrRcd
        }
        return $Results
    }

    function Enum-Creds {
        <#
        .Synopsis
          Enumerates stored credentials for operating user

        .Description
          Calls Win32 CredEnumerateW via [PsUtils.CredMan]::CredEnum

        .INPUTS
          
        .OUTPUTS
          [PsUtils.CredMan+Credential[]] if successful
          [Management.Automation.ErrorRecord] if unsuccessful or error encountered

        .PARAMETER Filter
          Specifies the filter to be applied to the query
          Defaults to [String]::Empty
          
        #>

        Param (
            [Parameter(Mandatory=$false)]
            [AllowEmptyString()]
            [String]$Filter = [String]::Empty
        )
        
        [PsUtils.CredMan+Credential[]]$Creds = [Array]::CreateInstance([PsUtils.CredMan+Credential], 0)
        [Int]$Results = 0
        try {
            $Results = [PsUtils.CredMan]::CredEnum($Filter, [Ref]$Creds)
        }
        catch {
            return $_
        }
        switch($Results) {
            0 {break}
            0x80070490 {break} #ERROR_NOT_FOUND
            default {
                [String]$Msg = "Failed to enumerate credentials store for user '$Env:UserName'"
                [Management.ManagementException] $MgmtException = New-Object Management.ManagementException($Msg)
                [Management.Automation.ErrorRecord] $ErrRcd = New-Object Management.Automation.ErrorRecord($MgmtException, $Results.ToString("X"), $ErrorCategory[$Results], $null)
                return $ErrRcd
            }
        }
        return $Creds
    }

    function Read-Creds {
        <#
        .Synopsis
            Reads specified credentials for operating user

        .Description
            Calls Win32 CredReadW via [PsUtils.CredMan]::CredRead

        .INPUTS

        .OUTPUTS
            [PsUtils.CredMan+Credential] if successful
            [Management.Automation.ErrorRecord] if unsuccessful or error encountered

        .PARAMETER Target
            Specifies the URI for which the credentials are associated
            If not provided, the username is used as the target
          
        .PARAMETER CredType
            Specifies the desired credentials type; defaults to 
            "CRED_TYPE_GENERIC"
        #>

        Param (
            [Parameter(Mandatory=$true)]
            [ValidateLength(1,32767)]
            [String]$Target,

            [Parameter(Mandatory=$false)]
            [ValidateSet("GENERIC","DOMAIN_PASSWORD","DOMAIN_CERTIFICATE","DOMAIN_VISIBLE_PASSWORD",
            "GENERIC_CERTIFICATE","DOMAIN_EXTENDED","MAXIMUM","MAXIMUM_EX")]
            [String]$CredType = "GENERIC"
        )
        
        #CRED_MAX_DOMAIN_TARGET_NAME_LENGTH
        if ("GENERIC" -ne $CredType -and 337 -lt $Target.Length) { 
            [String]$Msg = "Target field is longer ($($Target.Length)) than allowed (max 337 characters)"
            [Management.ManagementException]$MgmtException = New-Object Management.ManagementException($Msg)
            [Management.Automation.ErrorRecord]$ErrRcd = New-Object Management.Automation.ErrorRecord($MgmtException, 666, 'LimitsExceeded', $null)
            return $ErrRcd
        }
        [PsUtils.CredMan+Credential]$Cred = New-Object PsUtils.CredMan+Credential
        [Int]$Results = 0
        try {
            $Results = [PsUtils.CredMan]::CredRead($Target, $(Get-CredType $CredType), [Ref]$Cred)
        }
        catch {
            return $_
        }
        
        switch($Results) {
            0 {break}
            0x80070490 {return $null} #ERROR_NOT_FOUND
            default {
                [String] $Msg = "Error reading credentials for target '$Target' from '$Env:UserName' credentials store"
                [Management.ManagementException]$MgmtException = New-Object Management.ManagementException($Msg)
                [Management.Automation.ErrorRecord]$ErrRcd = New-Object Management.Automation.ErrorRecord($MgmtException, $Results.ToString("X"), $ErrorCategory[$Results], $null)
                return $ErrRcd
            }
        }
        return $Cred
    }

    function Write-Creds {
        <#
        .Synopsis
          Saves or updates specified credentials for operating user

        .Description
          Calls Win32 CredWriteW via [PsUtils.CredMan]::CredWrite

        .INPUTS

        .OUTPUTS
          [Boolean] true if successful
          [Management.Automation.ErrorRecord] if unsuccessful or error encountered

        .PARAMETER Target
          Specifies the URI for which the credentials are associated
          If not provided, the username is used as the target
          
        .PARAMETER UserName
          Specifies the name of credential to be read
          
        .PARAMETER Password
          Specifies the password of credential to be read
          
        .PARAMETER Comment
          Allows the caller to specify the comment associated with 
          these credentials
          
        .PARAMETER CredType
          Specifies the desired credentials type; defaults to 
          "CRED_TYPE_GENERIC"

        .PARAMETER CredPersist
          Specifies the desired credentials storage type;
          defaults to "CRED_PERSIST_ENTERPRISE"
        #>

        Param (
            [Parameter(Mandatory=$false)]
            [ValidateLength(0,32676)]
            [String]$Target,

            [Parameter(Mandatory=$true)]
            [ValidateLength(1,512)]
            [String]$UserName,

            [Parameter(Mandatory=$true)]
            [ValidateLength(1,512)]
            [String]$Password,

            [Parameter(Mandatory=$false)]
            [ValidateLength(0,256)]
            [String]$Comment = [String]::Empty,

            [Parameter(Mandatory=$false)]
            [ValidateSet("GENERIC","DOMAIN_PASSWORD","DOMAIN_CERTIFICATE","DOMAIN_VISIBLE_PASSWORD",
            "GENERIC_CERTIFICATE","DOMAIN_EXTENDED","MAXIMUM","MAXIMUM_EX")]
            [String]$CredType = "GENERIC",

            [Parameter(Mandatory=$false)]
            [ValidateSet("SESSION","LOCAL_MACHINE","ENTERPRISE")]
            [String]$CredPersist = "ENTERPRISE"
        )

        if ([String]::IsNullOrEmpty($Target)) {
            $Target = $UserName
        }
        #CRED_MAX_DOMAIN_TARGET_NAME_LENGTH
        if ("GENERIC" -ne $CredType -and 337 -lt $Target.Length) {
            [String] $Msg = "Target field is longer ($($Target.Length)) than allowed (max 337 characters)"
            [Management.ManagementException] $MgmtException = New-Object Management.ManagementException($Msg)
            [Management.Automation.ErrorRecord] $ErrRcd = New-Object Management.Automation.ErrorRecord($MgmtException, 666, 'LimitsExceeded', $null)
            return $ErrRcd
        }
        if ([String]::IsNullOrEmpty($Comment)) {
            $Comment = [String]::Format("Last edited by {0}\{1} on {2}",$Env:UserDomain,$Env:UserName,$Env:ComputerName)
        }
        [String]$DomainName = [Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().DomainName
        [PsUtils.CredMan+Credential]$Cred = New-Object PsUtils.CredMan+Credential
        
        switch($Target -eq $UserName -and 
        $("CRED_TYPE_DOMAIN_PASSWORD" -eq $CredType -or "CRED_TYPE_DOMAIN_CERTIFICATE" -eq $CredType)) {
            $true  {$Cred.Flags = [PsUtils.CredMan+CRED_FLAGS]::USERNAME_TARGET}
            $false  {$Cred.Flags = [PsUtils.CredMan+CRED_FLAGS]::NONE}
        }
        $Cred.Type = Get-CredType $CredType
        $Cred.TargetName = $Target
        $Cred.UserName = $UserName
        $Cred.AttributeCount = 0
        $Cred.Persist = Get-CredPersist $CredPersist
        $Cred.CredentialBlobSize = [Text.Encoding]::Unicode.GetBytes($Password).Length
        $Cred.CredentialBlob = $Password
        $Cred.Comment = $Comment

        [Int] $Results = 0
        try {
            $Results = [PsUtils.CredMan]::CredWrite($Cred)
        }
        catch {
            return $_
        }

        if(0 -ne $Results) {
            [String] $Msg = "Failed to write to credentials store for target '$Target' using '$UserName', '$Password', '$Comment'"
            [Management.ManagementException] $MgmtException = New-Object Management.ManagementException($Msg)
            [Management.Automation.ErrorRecord] $ErrRcd = New-Object Management.Automation.ErrorRecord($MgmtException, $Results.ToString("X"), $ErrorCategory[$Results], $null)
            return $ErrRcd
        }
        return $Results
    }

    #endregion

    #region Cmd-Line functionality
    function CredManMain {
    #region Adding credentials
        if ($AddCred) {
            if([String]::IsNullOrEmpty($User) -or [String]::IsNullOrEmpty($Pass)) {
                Write-Host "You must supply a user name and password (target URI is optional)."
                return
            }
            # may be [Int32] or [Management.Automation.ErrorRecord]
            [Object]$Results = Write-Creds $Target $User $Pass $Comment $CredType $CredPersist
            if (0 -eq $Results) {
                [Object]$Cred = Read-Creds $Target $CredType
                if ($null -eq $Cred) {
                    Write-Host "Credentials for '$Target', '$User' was not found."
                    return
                }
                if ($Cred -is [Management.Automation.ErrorRecord]) {
                    return $Cred
                }

                New-Variable -Name "AddedCredentialsObject" -Value $(
                    [pscustomobject][ordered]@{
                        UserName    = $($Cred.UserName)
                        Password    = $($Cred.CredentialBlob)
                        Target      = $($Cred.TargetName.Substring($Cred.TargetName.IndexOf("=")+1))
                        Updated     = "$([String]::Format('{0:yyyy-MM-dd HH:mm:ss}', $Cred.LastWritten.ToUniversalTime())) UTC"
                        Comment     = $($Cred.Comment)
                    }
                )

                return $AddedCredentialsObject
            }
            # will be a [Management.Automation.ErrorRecord]
            return $Results
        }
    #endregion  

    #region Removing credentials
        if ($DelCred) {
            if (-not $Target) {
                Write-Host "You must supply a target URI."
                return
            }
            # may be [Int32] or [Management.Automation.ErrorRecord]
            [Object]$Results = Del-Creds $Target $CredType 
            if (0 -eq $Results) {
                Write-Host "Successfully deleted credentials for '$Target'"
                return
            }
            # will be a [Management.Automation.ErrorRecord]
            return $Results
        }
    #endregion

    #region Reading selected credential
        if ($GetCred) {
            if(-not $Target) {
                Write-Host "You must supply a target URI."
                return
            }
            # may be [PsUtils.CredMan+Credential] or [Management.Automation.ErrorRecord]
            [Object]$Cred = Read-Creds $Target $CredType
            if ($null -eq $Cred) {
                Write-Host "Credential for '$Target' as '$CredType' type was not found."
                return
            }
            if ($Cred -is [Management.Automation.ErrorRecord]) {
                return $Cred
            }

            New-Variable -Name "AddedCredentialsObject" -Value $(
                [pscustomobject][ordered]@{
                    UserName    = $($Cred.UserName)
                    Password    = $($Cred.CredentialBlob)
                    Target      = $($Cred.TargetName.Substring($Cred.TargetName.IndexOf("=")+1))
                    Updated     = "$([String]::Format('{0:yyyy-MM-dd HH:mm:ss}', $Cred.LastWritten.ToUniversalTime())) UTC"
                    Comment     = $($Cred.Comment)
                }
            )

            return $AddedCredentialsObject
        }
    #endregion

    #region Reading all credentials
        if ($ShoCred) {
            # may be [PsUtils.CredMan+Credential[]] or [Management.Automation.ErrorRecord]
            [Object]$Creds = Enum-Creds
            if ($Creds -split [Array] -and 0 -eq $Creds.Length) {
                Write-Host "No Credentials found for $($Env:UserName)"
                return
            }
            if ($Creds -is [Management.Automation.ErrorRecord]) {
                return $Creds
            }

            $ArrayOfCredObjects = @()
            foreach($Cred in $Creds) {
                New-Variable -Name "AddedCredentialsObject" -Value $(
                    [pscustomobject][ordered]@{
                        UserName    = $($Cred.UserName)
                        Password    = $($Cred.CredentialBlob)
                        Target      = $($Cred.TargetName.Substring($Cred.TargetName.IndexOf("=")+1))
                        Updated     = "$([String]::Format('{0:yyyy-MM-dd HH:mm:ss}', $Cred.LastWritten.ToUniversalTime())) UTC"
                        Comment     = $($Cred.Comment)
                    }
                ) -Force

                if ($All) {
                    $AddedCredentialsObject | Add-Member -MemberType NoteProperty -Name "Alias" -Value "$($Cred.TargetAlias)"
                    $AddedCredentialsObject | Add-Member -MemberType NoteProperty -Name "AttribCnt" -Value "$($Cred.AttributeCount)"
                    $AddedCredentialsObject | Add-Member -MemberType NoteProperty -Name "Attribs" -Value "$($Cred.Attributes)"
                    $AddedCredentialsObject | Add-Member -MemberType NoteProperty -Name "Flags" -Value "$($Cred.Flags)"
                    $AddedCredentialsObject | Add-Member -MemberType NoteProperty -Name "PwdSize" -Value "$($Cred.CredentialBlobSize)"
                    $AddedCredentialsObject | Add-Member -MemberType NoteProperty -Name "Storage" -Value "$($Cred.Persist)"
                    $AddedCredentialsObject | Add-Member -MemberType NoteProperty -Name "Type" -Value "$($Cred.Type)"
                }

                $ArrayOfCredObjects +=, $AddedCredentialsObject
            }
            return $ArrayOfCredObjects
        }
    #endregion

    #region Run basic diagnostics
        if($RunTests) {
            [PsUtils.CredMan]::Main()
        }
    #endregion
    }
    #endregion

    CredManMain
}


<#
.SYNOPSIS
    Sets up the GitHub Git Shell Environment in PowerShell

.DESCRIPTION
    Sets up the proper PATH and ENV to use GitHub for Window's shell environment
    
    This is a refactored version of $env:LOCALAPPDATA\GitHub\shell.ps1 that gets installed
    with GitDesktop for Windows.

.PARAMETER SkipSSHSetup
    If true, skips calling GitHub.exe to autoset and upload ssh-keys

.EXAMPLE
    Initialize-GitEnvironment

#>
function Initialize-GitEnvironment {
    [CmdletBinding(DefaultParameterSetname='Skip AuthSetup')]
    Param(
        [Parameter(Mandatory=$False)]
        [string]$GitHubUserName = $(Read-Host -Prompt "Please enter your GitHub Username"),

        [Parameter(Mandatory=$False)]
        [string]$GitHubEmail = $(Read-Host -Prompt "Please the primary GitHub email address associated with $GitHubUserName"),

        [Parameter(Mandatory=$False)]
        [ValidateSet("https","ssh")]
        [string]$AuthMethod,

        [Parameter(
            Mandatory=$False,
            ParameterSetName='SSH Auth'
        )]
        [string]$ExistingSSHPrivateKeyPath,

        [Parameter(
            Mandatory=$False,
            ParameterSetName='SSH Auth'
        )]
        [string]$NewSSHKeyName,

        [Parameter(
            Mandatory=$False,
            ParameterSetName='SSH Auth'
        )]
        $NewSSHKeyPwd,

        [Parameter(
            Mandatory=$False,
            ParameterSetName='SSH Auth'
        )]
        [switch]$DownloadAndSetupDependencies,

        [Parameter(
            Mandatory=$False,
            ParameterSetName='HTTPS Auth'
        )]
        $PersonalAccessToken
    )

    ##### BEGIN Variable/Parameter Transforms and PreRun Prep #####

    if ($PersonalAccessToken) {
        if ($PersonalAccessToken.GetType().FullName -eq "System.Security.SecureString") {
            # Convert SecureString to PlainText
            $PersonalAccessToken = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($PersonalAccessToken))
        }
    }

    $CurrentUser = $($([System.Security.Principal.WindowsIdentity]::GetCurrent()).Name -split "\\")[-1]

    if ($ExistingSSHPrivateKeyPath -or $NewSSHKeyName -or $NewSSHKeyPwd -or $DownloadAndSetupDependencies) {
        $AuthMethod = "ssh"
    }
    if ($PersonalAccessToken) {
        $AuthMethod = "https"
    }
    if ($AuthMethod -eq "https" -and $($DownloadAndSetupDependencies -or $ExistingSSHPrivateKeyPath -or $NewSSHKeyName -or $NewSSHKeyPwd)) {
        Write-Verbose "The parameters -DownloadAndSetupDependencies, -ExistingSSHPrivateKeyPath, -NewSSHKeyName, and -NewSSHKeyPwd should only be used when -AuthMethod is `"ssh`"! Halting!"
        Write-Error "The parameters -DownloadAndSetupDependencies, -ExistingSSHPrivateKeyPath, -NewSSHKeyName, and -NewSSHKeyPwd should only be used when -AuthMethod is `"ssh`"! Halting!"
        $global:FunctionResult = "1"
        return
    }
    # NOTE: We do NOT need to force use of -ExistingSSHPrivateKeyPath or -NewSSHKeyName when -AuthMethod is "ssh"
    # because Setup-GitAuthentication function can handle things if neither are provided
    if ($AuthMethod -eq "https" -and !$PersonalAccessToken) {
        Write-Verbose "If -AuthMethod is `"https`", you must use the -PersonalAccessToken parameter! Halting!"
        Write-Error "If -AuthMethod is `"https`", you must use the -PersonalAccessToken parameter! Halting!"
        $global:FunctionResult = "1"
        return
    }
    if ($NewSSHKeyPwd) {
        if ($NewSSHKeyPwd.GetType().FullName -eq "System.Security.SecureString") {
            # Convert SecureString to PlainText
            $NewSSHKeyPwd = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($NewSSHKeyPwd))
        }
        if (!$NewSSHKeyName) {
            $NewSSHKeyName = "GitAuthFor$CurrentUser"
        }
    }

    # Check to make sure Git Desktop is Installed
    $GitDesktopVersion = [version]$(Check-InstalledPrograms -ProgramTitleSearchTerm "GitHub" | Where-Object {$_.DisplayName -notmatch "Machine-Wide Installer"}).DisplayVersion
    $GitHubDesktopVersionString = $(Check-InstalledPrograms -ProgramTitleSearchTerm "GitHub" | Where-Object {$_.DisplayName -notmatch "Machine-Wide Installer"}).DisplayVersion
    if ($GitDesktopVersion -eq $null) {
        Write-Verbose "A Git Desktop installation was not found. Please install GitHub Desktop and try again. Halting!"
        Write-Error "A Git Desktop installation was not found. Please install GitHub Desktop and try again. Halting!"
        $global:FunctionResult = "1"
        return
    }
    if ($GitDesktopVersion.Major -eq 0) {
        $GitDesktopChannel = "Beta"
    }
    if ($GitDesktopVersion.Major -gt 0) {
        $GitDesktopChannel = "Stable"
    }
    if (!$(Test-Path "$HOME\Documents\GitHub")) {
        New-Item -Type Directory -Path "$HOME\Documents\GitHub"
    }

    ##### END Variable/Parameter Transforms and PreRun Prep #####

    # Set the Git PowerShell Environment
    if ($GitDesktopChannel -eq "Beta") {
        if ($env:github_shell -eq $null) {
            while (!$(Resolve-Path "$env:LocalAppData\GitHubDesktop\app-$GitHubDesktopVersionString\resources\app\git" -ErrorAction SilentlyContinue)) {
                Write-Host "Waiting for $env:LocalAppData\GitHubDesktop\app-$GitHubDesktopVersionString\resources\app\git"
                Start-Sleep -Seconds 1
            }
            $env:github_git = $(Resolve-Path "$env:LocalAppData\GitHubDesktop\app-$GitHubDesktopVersionString\resources\app\git" -ErrorAction Continue).Path
            $env:PLINK_PROTOCOL = "ssh"
            $env:TERM = "msys"
            $env:HOME = $HOME
            $env:TMP = $env:TEMP = [system.io.path]::gettemppath()
            <#
            if ($env:EDITOR -eq $null) {
              $env:EDITOR = "GitPad"
            }
            #>

            # Setup PATH
            $pGitPath = $env:github_git
            $appPath = "$env:LocalAppData\GitHubDesktop\app-$GitHubDesktopVersion"
            while (!$appPath) {
                Write-Host "Waiting for `$appPath..."
                $appPath = "$env:LocalAppData\GitHubDesktop\app-$GitHubDesktopVersion"
                Start-Sleep -Seconds 1
            }
            $HighestNetVer = $($(Get-ChildItem "$env:SystemRoot\Microsoft.NET\Framework" | Where-Object {$_.Name -match "^v[0-9]"}).Name -replace "v","" | Measure-Object -Maximum).Maximum
            $msBuildPath = "$env:SystemRoot\Microsoft.NET\Framework\v$HighestNetVer"
            $lfsamd64Path = "$env:LocalAppData\GitHubDesktop\app-$GitHubDesktopVersion\resources\app\git\mingw64\libexec\git-core"

            if ($env:Path[-1] -eq ";") {
                $env:Path = "$env:Path$pGitPath\cmd;$pGitPath\usr\bin;$lfsamd64Path;$appPath;$msBuildPath"
            }
            else {
                $env:Path = "$env:Path;$pGitPath\cmd;$pGitPath\usr\bin;$lfsamd64Path;$appPath;$msBuildPath"
            }

            $env:git_bash_path = "$env:github_git\usr\bin"
            $env:github_shell = $true
            $env:git_install_root = $pGitPath
        }
        else {
            Write-Verbose "GitHub shell environment already setup"
        }
    }
    if ($GitDesktopChannel -eq "Stable") {
        if ($env:github_shell -eq $null) {
            $env:github_posh_git = $(Resolve-Path "$env:LocalAppData\GitHub\PoshGit_*" -ErrorAction Continue).Path
            while (!$(Resolve-Path "$env:LocalAppData\GitHub\PortableGit_*" -ErrorAction SilentlyContinue)) {
                Write-Host "Waiting for $env:LocalAppData\GitHub\PortableGit_*"
                Start-Sleep -Seconds 1
            }
            $env:github_git = $(Resolve-Path "$env:LocalAppData\GitHub\PortableGit_*" -ErrorAction Continue).Path
            $env:PLINK_PROTOCOL = "ssh"
            $env:TERM = "msys"
            $env:HOME = $HOME
            $env:TMP = $env:TEMP = [system.io.path]::gettemppath()
            if ($env:EDITOR -eq $null) {
              $env:EDITOR = "GitPad"
            }

            # Setup PATH
            $pGitPath = $env:github_git
            #$appPath = Resolve-Path "$env:LocalAppData\Apps\2.0\XE9KPQJJ.N9E\GALTN70J.73D\gith..tion_317444273a93ac29_0003.0003_5794af8169eeff14"
            $appPath = $(Get-ChildItem -Recurse -Path "$env:LocalAppData\Apps" | Where-Object {$_.Name -match "^gith..tion*" -and $_.FullName -notlike "*manifests*" -and $_.FullName -notlike "*\Data\*"}).FullName
            while (!$appPath) {
                Write-Host "Waiting for `$appPath..."
                $appPath = $(Get-ChildItem -Recurse -Path "$env:LocalAppData\Apps" | Where-Object {$_.Name -match "^gith..tion*" -and $_.FullName -notlike "*manifests*" -and $_.FullName -notlike "*\Data\*"}).FullName
                Start-Sleep -Seconds 1
            }
            $HighestNetVer = $($(Get-ChildItem "$env:SystemRoot\Microsoft.NET\Framework" | Where-Object {$_.Name -match "^v[0-9]"}).Name -replace "v","" | Measure-Object -Maximum).Maximum
            $msBuildPath = "$env:SystemRoot\Microsoft.NET\Framework\v$HighestNetVer"
            $lfsamd64Path = "$env:LocalAppData\GitHub\lfs-amd*"

            if ($env:Path[-1] -eq ";") {
                $env:Path = "$env:Path$pGitPath\cmd;$pGitPath\usr\bin;$pGitPath\usr\share\git-tfs;$lfsamd64Path;$appPath;$msBuildPath"
            }
            else {
                $env:Path = "$env:Path;$pGitPath\cmd;$pGitPath\usr\bin;$pGitPath\usr\share\git-tfs;$lfsamd64Path;$appPath;$msBuildPath"
            }

            $env:github_shell = $true
            $env:git_install_root = $pGitPath
            if ($env:github_posh_git) {
                $env:posh_git = "$env:github_posh_git\profile.example.ps1"
            }
        }
        else {
            Write-Verbose "GitHub shell environment already setup"
        }
    }

    if ($(Get-Module -ListAvailable | Where-Object {$_.Name -eq "posh-git"}) -eq $null) {
        Update-PackageManagement
        Install-Module posh-git -Scope CurrentUser
    }
    if ($(Get-Module | Where-Object {$_.Name -eq "posh-git"}) -eq $null) {
        Import-Module posh-git -Verbose
    }
    

    # Setup Authentication if requested #
    # Setup SSH
    if ($AuthMethod -eq "ssh") {
        $GitAuthParams = @{
            GitHubUserName      = $GitHubUserName
            GitHubEmail         = $GitHubEmail
            AuthMethod          = $AuthMethod
        }
        if (!$ExistingSSHPrivateKeyPath -and !$NewSSHKeyName) {
            $GitAuthParams = $GitAuthParams
        }
        if ($ExistingSSHPrivateKeyPath) {
            $GitAuthParams = $GitAuthParams.Add("ExistingSSHPrivateKeyPath",$ExistingSSHPrivateKeyPath)
        }
        if ($NewSSHKeyName) {
            if (!$NewSSHKeyPwd) {
                $GitAuthParams = $GitAuthParams.Add("NewSSHKeyName",$NewSSHKeyName)
            }
            else {
                $GitAuthParams = $GitAuthParams.Add("NewSSHKeyName",$NewSSHKeyName)
                $GitAuthParams = $GitAuthParams.Add("NewSSHKeyPwd",$NewSSHKeyPwd)
            }
        }
        if ($DownloadAndSetupDependencies) {
            $global:FunctionResult = "0"
            Setup-GitAuthentication @GitAuthParams -DownloadAndSetupDependencies
            if ($global:FunctionResult -eq "1") {
                Write-Verbose "The Setup-GitAuthentication function failed. Halting!"
                Write-Error "The Setup-GitAuthentication function failed. Halting!"
                $global:FunctionResult = "1"
                return
            }
        }
        else {
            $global:FunctionResult = "0"
            Setup-GitAuthentication @GitAuthParams
            if ($global:FunctionResult -eq "1") {
                Write-Verbose "The Setup-GitAuthentication function failed. Halting!"
                Write-Error "The Setup-GitAuthentication function failed. Halting!"
                $global:FunctionResult = "1"
                return
            }
        }
    }
    # Setup https
    if ($AuthMethod -eq "https") {
        $GitAuthParams = @{
            GitHubUserName = $GitHubUserName
            GitHubEmail = $GitHubEmail
            AuthMethod = $AuthMethod
            PersonalAccessToken = $PersonalAccessToken
        }
        $global:FunctionResult = "0"
        Setup-GitAuthentication @GitAuthParams
        if ($global:FunctionResult -eq "1") {
            Write-Verbose "The Setup-GitAuthentication function failed. Halting!"
            Write-Error "The Setup-GitAuthentication function failed. Halting!"
            $global:FunctionResult = "1"
            return
        }
    }
}


<#
.SYNOPSIS
    Configures Git installed on Windows via GitDesktop to authenticate via https or ssh.
    Optionally, clone all repos from the GitHub User you authenticate as.

.DESCRIPTION
    See Synopsis.

.EXAMPLE
    $GitAuthParams = @{
        GitHubUserName = "pldmgg"
        GitHubEmail = "pldmgg@genericemailprovider.com"
        AuthMethod = "https"
        PersonalAccessToken = "234567ujhgfw456734567890okfd3456"
    }

    Setup-GitAuthentication @GitAuthParams

.EXAMPLE
    $GitAuthParams = @{
        GitHubUserName = "pldmgg"
        GitHubEmail = "pldmgg@genericemailprovider.com"
        AuthMethod = "ssh"
        NewSSHKeyName "gitauth_rsa"
    }

    Setup-GitAuthentication @GitAuthParams

.EXAMPLE
    $GitAuthParams = @{
        GitHubUserName = "pldmgg"
        GitHubEmail = "pldmgg@genericemailprovider.com"
        AuthMethod = "ssh"
        ExistingSSHPrivateKeyPath = "$HOME\.ssh\github_rsa" 
    }
    
    Setup-GitAuthentication @GitAuthParams

#>
function Setup-GitAuthentication {
    [CmdletBinding(DefaultParameterSetname='AuthSetup')]
    Param(
        [Parameter(Mandatory=$False)]
        [string]$GitHubUserName = $(Read-Host -Prompt "Please enter your GitHub Username"),

        [Parameter(Mandatory=$False)]
        [string]$GitHubEmail = $(Read-Host -Prompt "Please the primary GitHub email address associated with $GitHubUserName"),

        [Parameter(Mandatory=$False)]
        [ValidateSet("https","ssh")]
        [string]$AuthMethod = $(Read-Host -Prompt "Please select the Authentication Method you would like to use. [https/ssh]"),

        [Parameter(
            Mandatory=$False,
            ParameterSetName='SSH Auth'
        )]
        [string]$ExistingSSHPrivateKeyPath,

        [Parameter(
            Mandatory=$False,
            ParameterSetName='SSH Auth'
        )]
        [string]$NewSSHKeyName,

        [Parameter(
            Mandatory=$False,
            ParameterSetName='SSH Auth'
        )]
        $NewSSHKeyPwd,

        [Parameter(
            Mandatory=$False,
            ParameterSetName='SSH Auth'
        )]
        [switch]$DownloadAndSetupDependencies,

        [Parameter(
            Mandatory=$False,
            ParameterSetName='HTTPS Auth'
        )]
        $PersonalAccessToken
    )

    ##### BEGIN Variable/Parameter Transforms and PreRun Prep #####

    if ($PersonalAccessToken) {
        if ($PersonalAccessToken.GetType().FullName -eq "System.Security.SecureString") {
            # Convert SecureString to PlainText
            $PersonalAccessToken = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($PersonalAccessToken))
        }
    }

    $CurrentUser = $($([System.Security.Principal.WindowsIdentity]::GetCurrent()).Name -split "\\")[-1]

    if ($ExistingSSHPrivateKeyPath -or $NewSSHKeyName -or $NewSSHKeyPwd -or $DownloadAndSetupDependencies) {
        $AuthMethod = "ssh"
    }
    if ($PersonalAccessToken) {
        $AuthMethod = "https"
    }
    if ($AuthMethod -eq "https" -and $($DownloadAndSetupDependencies -or $ExistingSSHPrivateKeyPath -or $NewSSHKeyName -or $NewSSHKeyPwd)) {
        Write-Verbose "The parameters -DownloadAndSetupDependencies, -ExistingSSHPrivateKeyPath, -NewSSHKeyName, and -NewSSHKeyPwd should only be used when -AuthMethod is `"ssh`"! Halting!"
        Write-Error "The parameters -DownloadAndSetupDependencies, -ExistingSSHPrivateKeyPath, -NewSSHKeyName, and -NewSSHKeyPwd should only be used when -AuthMethod is `"ssh`"! Halting!"
        $global:FunctionResult = "1"
        return
    }
    # NOTE: We do NOT need to force use of -ExistingSSHPrivateKeyPath or -NewSSHKeyName when -AuthMethod is "ssh"
    # because Setup-GitAuthentication function can handle things if neither are provided
    if ($AuthMethod -eq "https" -and !$PersonalAccessToken) {
        $PersonalAccessToken = Read-Host -Prompt "Please enter the GitHub Personal Access Token you would like to use for https authentication." -AsSecureString
    }
    if ($ExistingSSHPrivateKeyPath) {
        $ExistingSSHPrivateKeyPath = $(Resolve-Path $ExistingSSHPrivateKeyPath -ErrorAction SilentlyContinue).Path
        if (!$(Test-Path "$ExistingSSHPrivateKeyPath")) {
            Write-Verbose "Unable to find $ExistingSSHPrivateKeyPath! Halting!"
            Write-Error "Unable to find $ExistingSSHPrivateKeyPath! Halting!"
            $global:FunctionResult = "1"
            return
        }
    }
    if ($NewSSHKeyPwd) {
        if ($NewSSHKeyPwd.GetType().FullName -eq "System.Security.SecureString") {
            # Convert SecureString to PlainText
            $NewSSHKeyPwd = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($NewSSHKeyPwd))
        }
        if (!$NewSSHKeyName) {
            $NewSSHKeyName = "GitAuthFor$CurrentUser"
        }
    }

    $GitDesktopVersion = [version]$(Check-InstalledPrograms -ProgramTitleSearchTerm "GitHub" | Where-Object {$_.DisplayName -notmatch "Machine-Wide Installer"}).DisplayVersion
    if ($GitDesktopVersion -eq $null) {
        Write-Verbose "A Git Desktop installation was not found. Please install GitHub Desktop and try again. Halting!"
        Write-Error "A Git Desktop installation was not found. Please install GitHub Desktop and try again. Halting!"
        $global:FunctionResult = "1"
        return
    }
    if ($GitDesktopVersion.Major -eq 0) {
        $GitDesktopChannel = "Beta"
    }
    if ($GitDesktopVersion.Major -gt 0) {
        $GitDesktopChannel = "Stable"
    }
    if (!$(Test-Path "$HOME\Documents\GitHub")) {
        New-Item -Type Directory -Path "$HOME\Documents\GitHub"
    }

    if ($AuthMethod -eq "ssh") {
        # ssh Utilities could come from Git or From previously installed Windows OpenSSH. We want to make sure we use
        # Windows OpenSSH
        $Potential64ArchLocation = "C:\Program Files\OpenSSH-Win64"
        $Potential32ArchLocation = "C:\Program Files (x86)\OpenSSH-Win32"
        $Potential64ArchLocationRegex = $Potential64ArchLocation -replace "\\","\\"
        $Potential32ArchLocationRegex = $($($Potential32ArchLocation -replace "\\","\\") -replace "(","\(") -replace ")","\)"
        if (!$(Get-Command "ssh-keygen" -ErrorAction SilentlyContinue) -and !$DownloadAndSetupDependencies) {
            Write-Verbose "The Setup-GitAuthentication function depends on ssh-keygen.exe from OpenSSH-Win64 and it is currently not available. Run the Setup-GitAuthentication function again with the -DownloadAndSetupDependencies switch to add the dependency! Halting!"
            Write-Error "The Setup-GitAuthentication function depends on ssh-keygen.exe from OpenSSH-Win64 and it is currently not available. Run the Setup-GitAuthentication function again with the -DownloadAndSetupDependencies switch to add the dependency! Halting!"
            $global:FunctionResult = "1"
            return
        }
        $NeedWinOpenSSHScenario1 = !$(Get-Command "ssh-keygen" -ErrorAction SilentlyContinue) -and $DownloadAndSetupDependencies
        $NeedWinOpenSSHScenario2 = $(Get-Command "ssh-keygen" -ErrorAction SilentlyContinue) -and $(Get-Command "ssh-keygen" -All | Where-Object {
            $_.Source -match "$Potential32ArchLocation\ssh.exe|$Potential64ArchLocation\ssh.exe"
        }).Source -eq $null -and $DownloadAndSetupDependencies
        if ($NeedWinOpenSSHScenario1 -or $NeedWinOpenSSHScenario2) {
            $url = 'https://github.com/PowerShell/Win32-OpenSSH/releases/latest/'
            $request = [System.Net.WebRequest]::Create($url)
            $request.AllowAutoRedirect = $false
            $response = $request.GetResponse()
            $Win64OpenSSHDLLink = $([String]$response.GetResponseHeader("Location")).Replace('tag','download') + '/OpenSSH-Win64.zip'
            Invoke-WebRequest -Uri $Win64OpenSSHDLLink -OutFile "$HOME\Downloads\OpenSSH-Win64.zip"
            #if (!$(Test-Path "$HOME\Downloads\OpenSSH-Win64")) {
            #    New-Item -Type Directory -Path "$HOME\Downloads\OpenSSH-Win64"
            #}
            # NOTE: OpenSSH-Win64.zip contains a folder OpenSSH-Win64, so no need to create one before extraction
            Unzip-File -PathToZip "$HOME\Downloads\OpenSSH-Win64.zip" -TargetDir "$HOME\Downloads"
            $OpenSSHWin64Path = "$HOME\Downloads\OpenSSH-Win64"
            $env:Path = "$OpenSSHWin64Path;$env:Path"
        }
        # Make sure we're using the Windows OpenSSH Version of ssh-keygen.exe by moving the Git SSH Utilities Path
        # to the end of $env:Path
        $RegexMatches = $($env:Path | Select-String -Pattern "$Potential64ArchLocationRegex|$Potential32ArchLocationRegex|C:\\Users\\testadmin\\AppData\\Local\\GitHub\\PortableGit_.*\\usr\\bin" -AllMatches).Matches
        if ($RegexMatches[0].Value -match "Git") {
            $RegexValueToReplace = $RegexMatches[0].Value -replace "\\","\\"
            $env:Path = $($env:Path -replace "$RegexValueToReplace","") + ";" + $RegexMatches[0].Value
        }

        if (!$ExistingSSHPrivateKeyPath -and !$NewSSHKeyName) {
            if ($GitDesktopChannel -eq "Stable") {
                if (!$(Test-Path "$HOME\.ssh\github_rsa")) {
                    GitHub.exe --set-up-ssh
                    for ($i=0; $i -lt $(0..4).Count; $i++) {
                        Write-Host "Waiting $($(0..4).Count - $i) seconds for GitHub.exe --set-up-ssh to create $HOME\.ssh\github_rsa..."
                        if (Test-Path "$HOME\.ssh\github_rsa") {
                            Write-Host "GitHub.exe --set-up-ssh successfully created $HOME\.ssh\github_rsa. Continuing..."
                            break
                        }
                        Start-Sleep -Seconds 1
                    }
                    if (!$(Test-Path "$HOME\.ssh\github_rsa")) {
                        $NewSSHKeyName = "github_rsa"
                    }
                    else {
                        $ExistingSSHPrivateKeyPath = "$HOME\.ssh\github_rsa"
                    }
                }
                else {
                    $ExistingSSHPrivateKeyPath = "$HOME\.ssh\github_rsa"
                }
            }
            if ($GitDesktopChannel -eq "Beta") {
                $NewSSHKeyName = "GitAuthFor$CurrentUser"
            }
        }
        
        if ($NewSSHKeyName) {
            # Create new public/private keypair #
            if (!$(Test-Path "$HOME\.ssh")) {
                New-Item -Type Directory -Path "$HOME\.ssh"
            }

            if ($NewSSHKeyPwd) {
                ssh-keygen.exe -t rsa -b 2048 -f "$HOME\.ssh\$NewSSHKeyName" -q -N "$NewSSHKeyPwd" -C "GitAuthFor$CurrentUser"
            }
            else {
                 # Need PowerShell Await Module (Windows version of Linux Expect) for ssh-keygen with null password
                if ($(Get-Module -ListAvailable).Name -notcontains "Await" -and $DownloadAndSetupDependencies) {
                    # Install-Module "Await" -Scope CurrentUser
                    # Clone PoshAwait repo to .zip
                    Invoke-WebRequest -Uri "https://github.com/pldmgg/PoshAwait/archive/master.zip" -OutFile "$HOME\PoshAwait.zip"
                    $tempDirectory = [IO.Path]::Combine([IO.Path]::GetTempPath(), [IO.Path]::GetRandomFileName())
                    [IO.Directory]::CreateDirectory($tempDirectory)
                    Unzip-File -PathToZip "$HOME\PoshAwait.zip" -TargetDir "$tempDirectory"
                    if (!$(Test-Path "$HOME\Documents\WindowsPowerShell\Modules\Await")) {
                        New-Item -Type Directory "$HOME\Documents\WindowsPowerShell\Modules\Await"
                    }
                    Copy-Item -Recurse -Path "$tempDirectory\PoshAwait-master\*" -Destination "$HOME\Documents\WindowsPowerShell\Modules\Await"
                    Remove-Item -Recurse -Path $tempDirectory -Force
                }

                # Make private key password $null
                Import-Module Await
                if (!$?) {
                    Write-Verbose "Unable to load the Await Module! Halting!"
                    Write-Error "Unable to load the Await Module! Halting!"
                    $global:FunctionResult = "1"
                    return
                }

                Start-AwaitSession
                Start-Sleep -Seconds 1
                Send-AwaitCommand '$host.ui.RawUI.WindowTitle = "PSAwaitSession"'
                $PSAwaitProcess = $($(Get-Process | ? {$_.Name -eq "powershell"}) | Sort-Object -Property StartTime -Descending)[0]
                Start-Sleep -Seconds 1
                Send-AwaitCommand "`$env:Path = '$env:Path'"
                Start-Sleep -Seconds 1
                Send-AwaitCommand "ssh-keygen.exe -t rsa -b 2048 -f `"$HOME\.ssh\$NewSSHKeyName`" -C `"GitAuthFor$CurrentUser`""
                Start-Sleep -Seconds 1
                Send-AwaitCommand ""
                Start-Sleep -Seconds 1
                Send-AwaitCommand ""
                Start-Sleep -Seconds 1
                $SSHKeyGenConsoleOutput = Receive-AwaitResponse
                Write-hOst ""
                Write-Host "##### BEGIN ssh-keygen Console Output From PSAwaitSession #####"
                Write-Host "$SSHKeyGenConsoleOutput"
                Write-Host "##### END ssh-keygen Console Output From PSAwaitSession #####"
                Write-Host ""
                # If Stop-AwaitSession errors for any reason, it doesn't return control, so we need to handle in try/catch block
                try {
                    Stop-AwaitSession
                }
                catch {
                    if ($PSAwaitProcess.Id -eq $PID) {
                        Write-Verbose "The PSAwaitSession never spawned! Halting!"
                        Write-Error "The PSAwaitSession never spawned! Halting!"
                        $global:FunctionResult = "1"
                        return
                    }
                    else {
                        Stop-Process -Id $PSAwaitProcess.Id
                    }
                }
            }

            if (!$(Test-Path "$HOME\.ssh\$NewSSHKeyName")) {
                Write-Verbose "ssh-keygen did not successfully create the public/private keypair! Halting!"
                Write-Error "ssh-keygen did not successfully create the public/private keypair! Halting!"
                $global:FunctionResult = "1"
                return
            }
            else {
                $ExistingSSHPrivateKeyPath = "$HOME\.ssh\$NewSSHKeyName"
            }
        }

        # At this point, $ExistingSSHPrivateKeyPath should exist
        # Validate private key format
        $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
        $ProcessInfo.FileName = "ssh-keygen.exe"
        $ProcessInfo.RedirectStandardError = $true
        $ProcessInfo.RedirectStandardOutput = $true
        $ProcessInfo.UseShellExecute = $false
        $ProcessInfo.Arguments = "-lf $HOME\.ssh\github_rsa"
        $Process = New-Object System.Diagnostics.Process
        $Process.StartInfo = $ProcessInfo
        $Process.Start() | Out-Null
        $Process.WaitForExit()
        $stdout = $Process.StandardOutput.ReadToEnd()
        $stderr = $Process.StandardError.ReadToEnd()
        $AllOutput = $stdout + $stderr
        if ($AllOutput -match "is not") {
            Write-Verbose "ssh-keygen reports that the private key is not in valid format. Please check manually. Halting!"
            Write-Error "ssh-keygen reports that the private key is not in valid format. Please check manually. Halting!"
            $global:FunctionResult = "1"
            return
        }
    }
    
    Push-Location "$HOME\Documents\GitHub"

    if (!$(Get-Command git -ErrorAction SilentlyContinue)) {
        $global:FunctionResult = "0"
        Initialize-GitEnvironment -GitHubUserName $GitHubUserName -GitHubEmail $GitHubEmail
        if ($global:FunctionResult -eq "1") {
            Write-Verbose "The Initialize-GitEnvironment function failed! Halting!"
            Write-Error "The Initialize-GitEnvironment function failed! Halting!"
            Pop-Location
            $global:FunctionResult = "1"
            return
        }
    }

    ##### END Variable/Parameter Transforms and PreRun Prep #####


    ##### BEGIN Main Body #####

    git config --global user.name "$GitHubUserName"
    git config --global user.email "$GitHubEmail"


    if ($AuthMethod -eq "https") {
        git config --global credential.helper wincred

        # Alternate Stored Credentials Format
        <#
        $ManageStoredCredsParams = @{
            Target  = "git:https://$PersonalAccessToken@github.com"
            User    = $PersonalAccessToken
            Pass    = 'x-oauth-basic'
            Comment = "Saved By Manage-StoredCredentials.ps1"
        }
        #>
        $ManageStoredCredsParams = @{
            Target  = "git:https://$GitHubUserName@github.com"
            User    = $GitHubUserName
            Pass    = $PersonalAccessToken
            Comment = "Saved By Manage-StoredCredentials.ps1"
        }
        Manage-StoredCredentials -AddCred @ManageStoredCredsParams

        # Test https OAuth2 authentication
        # More info here: https://channel9.msdn.com/Blogs/trevor-powershell/Automating-the-GitHub-REST-API-Using-PowerShell
        $Token = "$GitHubUserName`:$PersonalAccessToken"
        $Base64Token = [System.Convert]::ToBase64String([char[]]$Token)
        $Headers = @{
            Authorization = "Basic {0}" -f $Base64Token
        }
        $PublicAndPrivateRepos = $(Invoke-RestMethod -Headers $Headers -Uri "https://api.github.com/user/repos?access_token=$PersonalAccessToken").Name
        Write-Host "Writing Public and Private Repos to demonstrate https authentication success..."
        Write-Host "$($PublicAndPrivateRepos -join ", ")"
    }

    if ($AuthMethod -eq "ssh") {
        # Start the ssh agent and add your ExistingSSHPrivateKeyPath to it. We need to do this regardless...
        if (!$(Get-Module -List -Name posh-git)) {
            if ($PSVersionTable.PSVersion.Major -ge 5) {
                Update-PackageManagement
                Install-Module posh-git -Scope CurrentUser
                Import-Module posh-git -Verbose
            }
            if ($PSVersionTable.PSVersion.Major -lt 5) {
                Update-PackageManagement
                Install-Module posh-git -Scope CurrentUser
                Import-Module posh-git -Verbose
            }
        }
        Start-SshAgent
        Add-SshKey $ExistingSSHPrivateKeyPath

        # Check To Make Sure Online GitHub Account is aware of Existing Public Key
        $PubSSHKeys = Invoke-Restmethod -Uri "https://api.github.com/users/$GitHubUserName/keys"
        $tempfileLocations = @()
        foreach ($PubKeyObject in $PubSSHKeys) {
            $tmpFile = [IO.Path]::GetTempFileName()
            $PubKeyObject.key | Out-File $tmpFile -Encoding ASCII

            $tempfileLocations +=, $tmpFile
        }
        $SSHPubKeyFingerPrintsFromGitHub = foreach ($TempPubSSHKeyFile in $tempfileLocations) {
            $PubKeyFingerPrintPrep = ssh-keygen -E md5 -lf $TempPubSSHKeyFile
            $PubKeyFingerPrint = $($PubKeyFingerPrintPrep -split " ")[1] -replace "MD5:",""
            $PubKeyFingerPrint
        }
        # Cleanup Temp Files
        foreach ($TempPubSSHKeyFile in $tempfileLocations) {
            Remove-Item $TempPubSSHKeyFile
        }

        $GitHubOnlineIsAware = @()
        foreach ($fingerprint in $SSHPubKeyFingerPrintsFromGitHub) {
            $ExistingSSHPubKeyPath = "$ExistingSSHPrivateKeyPath.pub"
            $LocalPubKeyFingerPrintPrep = ssh-keygen -E md5 -lf $ExistingSSHPubKeyPath
            $LocalPubKeyFingerPrint = $($LocalPubKeyFingerPrintPrep -split " ")[1] -replace "MD5:",""
            if ($fingerprint -eq $LocalPubKeyFingerPrint) {
                $GitHubOnlineIsAware +=, $fingerprint
            }
        }

        if ($GitHubOnlineIsAware.Count -gt 0) {
            Write-Host "GitHub Online Account is aware of existing public key $ExistingSSHPubKeyPath. Testing the connection..."

            # Test the connection
            $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
            $ProcessInfo.FileName = "ssh.exe"
            $ProcessInfo.RedirectStandardError = $true
            $ProcessInfo.RedirectStandardOutput = $true
            $ProcessInfo.UseShellExecute = $false
            $ProcessInfo.Arguments = "-T git@github.com"
            $Process = New-Object System.Diagnostics.Process
            $Process.StartInfo = $ProcessInfo
            $Process.Start() | Out-Null
            $Process.WaitForExit()
            $stdout = $Process.StandardOutput.ReadToEnd()
            $stderr = $Process.StandardError.ReadToEnd()
            $AllOutput = $stdout + $stderr

            if ($AllOutput -match $GitHubUserName) {
                Write-Host "GitHub Authentication for $GitHubUserName using SSH was successful."
            }
            else {
                Write-Warning "GitHub Authentication for $GitHubUserName using SSH was NOT successful. Please check your connection and/or keys."
            }
        }
        if ($GitHubOnlineIsAware.Count -eq 0 -or $NewSSHKeyName) {
            Write-Host ""
            Write-Host "GitHub Authentication was successfully configured on the client machine, however, the GitHub Online Account is not aware of the local public SSH key $ExistingSSHPrivateKeyPath.pub"
            Write-Host "Please add $HOME\.ssh\$ExistingSSHPrivateKeyPath.pub to your GitHub Account via Web Browser by:"
            Write-Host "    1) Navigating to Settings"
            Write-Host "    2) In the user settings sidebar, click SSH and GPG keys."
            Write-Host "    3) Add SSH Key"
            Write-Host "    4) Enter a descriptive Title like: SSH Key for Paul-MacBookPro auth"
            Write-Host "    5) Paste your key into the Key field."
            Write-Host "    6) Click Add SSH key."
        }
    }

    Pop-Location

    ##### END Main Body #####

}

function Install-GitDesktop {
    [CmdletBinding(DefaultParameterSetname='AuthSetup')]
    Param(
        [Parameter(Mandatory=$False)]
        [string]$GitHubUserName = $(Read-Host -Prompt "Please enter your GitHub Username"),

        [Parameter(Mandatory=$False)]
        [string]$GitHubEmail = $(Read-Host -Prompt "Please the primary GitHub email address associated with $GitHubUserName"),

        [Parameter(Mandatory=$False)]
        [ValidateSet("https","ssh")]
        [string]$AuthMethod = $(Read-Host -Prompt "Please select the Authentication Method you would like to use. [https/ssh]"),

        [Parameter(
            Mandatory=$False,
            ParameterSetName='SSH Auth'
        )]
        [string]$ExistingSSHPrivateKeyPath,

        [Parameter(
            Mandatory=$False,
            ParameterSetName='SSH Auth'
        )]
        [string]$NewSSHKeyName,

        [Parameter(
            Mandatory=$False,
            ParameterSetName='SSH Auth'
        )]
        $NewSSHKeyPwd,

        [Parameter(
            Mandatory=$False,
            ParameterSetName='SSH Auth'
        )]
        [switch]$DownloadAndSetupDependencies,

        [Parameter(
            Mandatory=$False,
            ParameterSetName='HTTPS Auth'
        )]
        $PersonalAccessToken,

        [Parameter(Mandatory=$False)]
        [ValidateSet("Stable", "Beta")]
        [string]$Channel = "Stable",

        [Parameter(Mandatory=$False)]
        [ValidateSet("exe", "msi")]
        [string]$InstallerType
    )

    ##### BEGIN Variable/Parameter Transforms and PreRun Prep #####

    if ($PersonalAccessToken) {
        if ($PersonalAccessToken.GetType().FullName -eq "System.Security.SecureString") {
            # Convert SecureString to PlainText
            $PersonalAccessToken = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($PersonalAccessToken))
        }
    }

    $CurrentUser = $($([System.Security.Principal.WindowsIdentity]::GetCurrent()).Name -split "\\")[-1]

    if ($ExistingSSHPrivateKeyPath -or $NewSSHKeyName -or $NewSSHKeyPwd -or $DownloadAndSetupDependencies) {
        $AuthMethod = "ssh"
    }
    if ($PersonalAccessToken) {
        $AuthMethod = "https"
    }
    if ($AuthMethod -eq "https" -and $($DownloadAndSetupDependencies -or $ExistingSSHPrivateKeyPath -or $NewSSHKeyName -or $NewSSHKeyPwd)) {
        Write-Verbose "The parameters -DownloadAndSetupDependencies, -ExistingSSHPrivateKeyPath, -NewSSHKeyName, and -NewSSHKeyPwd should only be used when -AuthMethod is `"ssh`"! Halting!"
        Write-Error "The parameters -DownloadAndSetupDependencies, -ExistingSSHPrivateKeyPath, -NewSSHKeyName, and -NewSSHKeyPwd should only be used when -AuthMethod is `"ssh`"! Halting!"
        $global:FunctionResult = "1"
        return
    }
    # NOTE: We do NOT need to force use of -ExistingSSHPrivateKeyPath or -NewSSHKeyName when -AuthMethod is "ssh"
    # because Setup-GitAuthentication function can handle things if neither are provided
    if ($AuthMethod -eq "https" -and !$PersonalAccessToken) {
        $PersonalAccessToken = Read-Host -Prompt "Please enter the GitHub Personal Access Token you would like to use for https authentication." -AsSecureString
    }
    if ($ExistingSSHPrivateKeyPath) {
        $ExistingSSHPrivateKeyPath = $(Resolve-Path $ExistingSSHPrivateKeyPath -ErrorAction SilentlyContinue).Path
        if (!$(Test-Path "$ExistingSSHPrivateKeyPath")) {
            Write-Verbose "Unable to find $ExistingSSHPrivateKeyPath! Halting!"
            Write-Error "Unable to find $ExistingSSHPrivateKeyPath! Halting!"
            $global:FunctionResult = "1"
            return
        }
    }
    if ($NewSSHKeyPwd) {
        if ($NewSSHKeyPwd.GetType().FullName -eq "System.Security.SecureString") {
            # Convert SecureString to PlainText
            $NewSSHKeyPwd = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($NewSSHKeyPwd))
        }
        if (!$NewSSHKeyName) {
            $NewSSHKeyName = "GitAuthFor$CurrentUser"
        }
    }

    if ($Channel -eq "Stable" -and $InstallerType -eq "msi") {
        Write-Host "Currently, there is no .msi installer available for GitHub Desktop on Windows. Installing using .exe..."
        $InstallerType = "exe"
    }

    if (!$(Check-Elevation)) {
        $UserName = $($([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).split("\"))[1]
        $Psswd = Read-Host -Prompt "Please enter the password for $UserName" -AsSecureString
        $Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName, $Psswd
    }

    ##### END Variable/Parameter Transforms and PreRun Prep #####


    ##### BEGIN Main Body #####

    # For more info on SendKeys method, see: https://msdn.microsoft.com/en-us/library/office/aa202943(v=office.10).aspx
    # https://desktop.githubusercontent.com/releases/0.5.8-e55db469/GitHubDesktopSetup.exe
    if ($Channel -eq "Beta") {
        $SpecificReleaseNumber = "0.5.8-e55db469"
        # For latest release, use the below URLs
        # Invoke-WebRequest -Uri "https://central.github.com/deployments/desktop/desktop/latest/win32?format=msi" -OutFile "$HOME\Downloads\GitHubDesktopSetup.msi"
        # Invoke-WebRequest -Uri "https://central.github.com/deployments/desktop/desktop/latest/win32" -OutFile "$HOME\Downloads\GitHubDesktopSetup.exe"
        
        if ($InstallerType -eq "exe") {
            Invoke-WebRequest -Uri "https://desktop.githubusercontent.com/releases/$SpecificReleaseNumber/GitHubDesktopSetup.exe" -OutFile "$HOME\Downloads\GitHubDesktopSetup.exe"
            if (!$?) {
                Write-Verbose "Unable to download file! Halting!"
                Write-Error "Unable to download file! Halting!"
                $global:FunctionResult = "1"
                return
            }
            $GitHubDesktopVersion = $(Get-ChildItem $HOME\Downloads\GitHubDesktopSetup.exe).VersionInfo.ProductVersion
            & "$HOME\Downloads\GitHubDesktopSetup.exe"
        }
        if ($InstallerType -eq "msi") {
            Invoke-WebRequest -Uri "https://desktop.githubusercontent.com/releases/$SpecificReleaseNumber/GitHubDesktopSetup.msi" -OutFile "$HOME\Downloads\GitHubDesktopSetup.msi"
            if (!$?) {
                Write-Verbose "Unable to download file! Halting!"
                Write-Error "Unable to download file! Halting!"
                $global:FunctionResult = "1"
                return
            }

            # We're going to need Elevated privileges for to install the .msi if PowerShell wasn't Run As Administrator, Start-SudoSession
            if (!$(Check-Elevation)) {
                if (!$global:ElevatedPSSession) {
                    try {
                        $global:ElevatedPSSession = New-PSSession -Name "TempElevatedSession "-Authentication CredSSP -Credential $Credentials -ErrorAction SilentlyContinue
                        if (!$ElevatedPSSession) {
                            throw
                        }
                        $CredSSPAlreadyConfigured = $true
                    }
                    catch {
                        $SudoSession = New-SudoSession -Credentials $Credentials
                        $global:ElevatedPSSession = $SudoSession.ElevatedPSSession
                        $NeedToRevertAdminChangesIfAny = $true
                    }
                }
            }

            $DataStamp = Get-Date -Format yyyyMMddTHHmmss
            $MSIFullPath = "$HOME\Downloads\GitHubDesktopSetup.msi"
            $MSIParentDir = $MSIFullPath | Split-Path -Parent
            $MSIFileName = $MSIFullPath | Split-Path -Leaf
            $MSIFileNameOnly = $MSIFileName -replace "\.msi",""
            $logFile = "$HOME\$MSIFileNameOnly$DataStamp.log"
            $MSIArguments = @(
                "/i"
                $MSIFullPath
                "/qn"
                "/norestart"
                "/L*v"
                $logFile
            )

            if ($ElevatedPSSession) {
                $MSIExecOutput = Invoke-Command -Session $ElevatedPSSession -Scriptblock {
                    $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
                    $ProcessInfo.FileName = "msiexec.exe"
                    $ProcessInfo.RedirectStandardError = $true
                    $ProcessInfo.RedirectStandardOutput = $true
                    $ProcessInfo.UseShellExecute = $false
                    $ProcessInfo.Arguments = $using:MSIArguments
                    $Process = New-Object System.Diagnostics.Process
                    $Process.StartInfo = $ProcessInfo
                    $Process.Start() | Out-Null
                    $Process.WaitForExit()
                    $stdout = $Process.StandardOutput.ReadToEnd()
                    $stderr = $Process.StandardError.ReadToEnd()
                    $AllOutput = $stdout + $stderr
                    $AllOutput
                }
            }
            else {
                $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
                $ProcessInfo.FileName = "msiexec.exe"
                $ProcessInfo.RedirectStandardError = $true
                $ProcessInfo.RedirectStandardOutput = $true
                $ProcessInfo.UseShellExecute = $false
                $ProcessInfo.Arguments = $MSIArguments
                $Process = New-Object System.Diagnostics.Process
                $Process.StartInfo = $ProcessInfo
                $Process.Start() | Out-Null
                $Process.WaitForExit()
                $stdout = $Process.StandardOutput.ReadToEnd()
                $stderr = $Process.StandardError.ReadToEnd()
                $AllOutput = $stdout + $stderr
                $MSIExecOutput = $AllOutput
            }

            while ($(Test-Path "C:\Program Files (x86)\GitHub Desktop Installer\GitHubDesktop.exe").VersionInfo.ProductVersion) {
                Write-Host "Waiting for C:\Program Files (x86)\GitHub Desktop Installer\GitHubDesktop.exe to be extracted from .msi..."
                Start-Sleep -Seconds 1
            }
            $GitHubDesktopVersion = $(Get-ChildItem "C:\Program Files (x86)\GitHub Desktop Installer\GitHubDesktop.exe").VersionInfo.ProductVersion

            # Wait for 5 seconds to see if the Process with ProcessName "GitHubDesktop" spawns.
            # NOTE: The process "GitHubDesktop" (i.e. no space between characters) is the intial install process, and
            # the processes "GitHub Desktop" (i.e. with a space) are the Electron processes running when the program launches
             for ($i=0; $i -lt $(0..4).Count; $i++) {
                 Write-Host "Waiting $($(0..4).Count - $i) seconds for GitHubDesktop initial install process to start..."
                 Start-Sleep -Seconds 1
                 if (Get-Process | ? {$_.ProcessName -eq "GitHubDesktop"}) {
                    Write-Host "GitHubDesktop initial install process started...Continuing..."
                    $GitHubInitialInstallStarted = $true
                    break
                 }
             }
             if ($GitHubInitialInstallStarted -ne $true) {
                Write-Host "Starting C:\Program Files (x86)\GitHub Desktop Installer\GitHubDesktop.exe manually..."
                & "C:\Program Files (x86)\GitHub Desktop Installer\GitHubDesktop.exe"
             }
        }

        $GitHubDesktopLaunchWindow = Get-Process | Where-Object {$_.ProcessName -eq "GitHub Desktop" -and $_.MainWindowTitle -eq "GitHub Desktop"}
        while ($GitHubDesktopLaunchWindow -eq $null) {
            Write-Host "Waiting for GitHub Desktop to launch..."
            $GitHubDesktopLaunchWindow = Get-Process | Where-Object {$_.ProcessName -eq "GitHub Desktop" -and $_.MainWindowTitle -eq "GitHub Desktop"}
            Start-Sleep -Seconds 2
        }

        Write-Host "Finished Installing GitHub Desktop"

        while (!$(Resolve-Path "$env:LocalAppData\GitHubDesktop\app-$GitHubDesktopVersion\resources\app\git\cmd\git.exe" -ErrorAction SilentlyContinue)) {
            Write-Host "Waiting for $env:LocalAppData\GitHubDesktop\app-$GitHubDesktopVersion\resources\app\git\cmd\git.exe"
            Write-Host "This could take up to 1 minute..."
            Start-Sleep -Seconds 2
        }
        if (Test-Path $(Resolve-Path "$env:LocalAppData\GitHubDesktop\app-$GitHubDesktopVersion\resources\app\git\cmd\git.exe" -ErrorAction SilentlyContinue).Path) {
            Write-Host "Local App Data GitHub Directory is ready."
            Write-Host "Closing GitDesktop..."
        }

        # Redefine $GitHubDesktopLaunchWindow again because Electron makes PID of active window slippery
        $GitHubDesktopLaunchWindow = Get-Process | Where-Object {$_.ProcessName -eq "GitHub Desktop" -and $_.MainWindowTitle -eq "GitHub Desktop"}
        Stop-Process -Id $GitHubDesktopLaunchWindow.Id
    }
    if ($Channel -eq "Stable") {
        Invoke-WebRequest -Uri "https://github-windows.s3.amazonaws.com/GitHubSetup.exe" -OutFile "$HOME\Downloads\GitHubSetup.exe"
        if (!$?) {
            Write-Verbose "Unable to download file! Halting!"
            Write-Error "Unable to download file! Halting!"
            $global:FunctionResult = "1"
            return
        }
        & "$HOME\Downloads\GitHubSetup.exe"

        # Setup Potential GUI menus we may/may not have to step through
        $AppInstallSecWarnWindow = $(Get-Process | Where-Object {$_.MainWindowTitle -like "*Install - Security Warning*"}).MainWindowTitle
        $OpenFileWarning = $(Get-Process | Where-Object {$_.MainWindowTitle -like "*File - Security Warning*"}).MainWindowTitle
        $GitHubDesktop = Get-Process | Where-Object {$_.MainWindowTitle -eq "GitHub" -and $_.ProcessName -eq "GitHub"}

        while (!$AppInstallSecWarnWindow -and !$OpenFileWarning) {
            Write-Host "Waiting For Download to finish..."
            Start-Sleep -Seconds 2
            $AppInstallSecWarnWindow = $(Get-Process | Where-Object {$_.MainWindowTitle -like "*Install - Security Warning*"}).MainWindowTitle
            $OpenFileWarning = $(Get-Process | Where-Object {$_.MainWindowTitle -like "*File - Security Warning*"}).MainWindowTitle
        }
        if ($AppInstallSecWarnWindow -or $OpenFileWarning) {
            Write-Host "Download finished. Installing..."
        }

        if ($AppInstallSecWarnWindow) {
            $wshell = New-Object -ComObject wscript.shell
            $wshell.AppActivate("$AppInstallSecWarnWindow") | Out-Null
            #1..4 | foreach {$wshell.SendKeys('{TAB}')}
            $wshell.SendKeys('{i}')
        }

        if ($OpenFileWarning) {
            $wshell = New-Object -ComObject wscript.shell
            $wshell.AppActivate("$OpenFileWarning") | Out-Null
            #1..4 | foreach {$wshell.SendKeys('{TAB}')}
            $wshell.SendKeys('{r}')
        }
        
        while (!$GitHubDesktop) {
            Write-Host "Waiting For GitDesktop to launch..."
            Start-Sleep -Seconds 2
            $GitHubDesktop = Get-Process | Where-Object {$_.MainWindowTitle -eq "GitHub" -and $_.ProcessName -eq "GitHub"}
        }
        if ($GitHubDesktop) {
            Write-Host "GitDesktop launched."
            Start-Sleep -Seconds 2
            $GitHubDesktop | Set-WindowStyle -Style MINIMIZE
        }

        while (!$(Resolve-Path "$env:LocalAppData\GitHub\PortableGit_*\cmd\git.exe" -ErrorAction SilentlyContinue)) {
            Write-Host "Waiting for $env:LocalAppData\GitHub\PortableGit_*\cmd\git.exe"
            Write-Host "This could take up to 1 minute..."
            Start-Sleep -Seconds 2
        }
        if (Test-Path $(Resolve-Path "$env:LocalAppData\GitHub\PortableGit_*\cmd\git.exe" -ErrorAction SilentlyContinue).Path) {
            Write-Host "Local App Data GitHub Directory is ready."
            Write-Host "Closing GitDesktop..."
        }

        Start-Sleep -Seconds 5

        $GitHubDesktopPID = $(Get-Process | Where-Object {$_.MainWindowTitle -eq "GitHub" -and $_.ProcessName -eq "GitHub"}).Id
        Stop-Process -Id $GitHubDesktopPID

    }
    if (!$(Test-Path "$HOME\Documents\GitHub")) {
        New-Item -Type Directory -Path "$HOME\Documents\GitHub"
    }

    if (!$(Test-Path "$env:LocalAppData\GitHub\PoshGit*")) {
        if (!$(Get-Module -List -Name posh-git)) {
            if (!$(Check-Elevation)) {
                Write-Host "Updating PackageManagement. Please wait..."
                Update-PackageManagement -Credentials $Credentials
            }
            else {
                Write-Host "Updating PackageManagement. Please wait..."
                Update-PackageManagement
            }
            Install-Module posh-git -Scope CurrentUser
        }
    }

    # Check all PSModule Paths for posh-git
    <#
    $PSModulePathsArray = $env:PSModulePath -split ";"
    $PotentialPoshGitModulePaths = foreach ($potpath in $PSModulePathsArray) {
        "$potpath\posh-git"
    }
    while ($TestPotentialPoshGitModulePaths -notcontains $true) {
        $TestPotentialPoshGitModulePaths = foreach ($poshgitpath in $PotentialPoshGitModulePaths) {
            Test-Path $poshgitpath
        }
        Write-Host "Waiting for posh-git PowerShell module to be ready in $HOME\Documents\WindowsPowerShell\Modules ..."
        Start-Sleep -Seconds 2
    }
    if ($TestPotentialPoshGitModulePaths -contains $true) {
        Write-Host "posh-git PowerShell module is ready. Setting up GitHub Authentication using $AuthMethod..."
    }
    #>

    Write-Host "posh-git PowerShell module is ready. Setting up GitHub Authentication using $AuthMethod..."

    # Set the Git PowerShell Environment
    if (!$(Get-Command git -ErrorAction SilentlyContinue)) {
        $global:FunctionResult = "0"
        Initialize-GitEnvironment -GitHubUserName $GitHubUserName -GitHubEmail $GitHubEmail
        if ($global:FunctionResult -eq "1") {
            Write-Warning "GitHub Desktop was successfully installed, but the Git Environment could not be initialized"
            Write-Verbose "The Initialize-GitEnvironment function failed! Halting!"
            Write-Error "The Initialize-GitEnvironment function failed! Halting!"
            $global:FunctionResult = "1"
            return
        }
    }

    if ($AuthMethod -eq "ssh") {
        $GitAuthParams = @{
            GitHubUserName      = $GitHubUserName
            GitHubEmail         = $GitHubEmail
            AuthMethod          = $AuthMethod
        }
        if (!$ExistingSSHPrivateKeyPath -and !$NewSSHKeyName) {
            $GitAuthParams = $GitAuthParams
        }
        if ($ExistingSSHPrivateKeyPath) {
            $GitAuthParams = $GitAuthParams.Add("ExistingSSHPrivateKeyPath",$ExistingSSHPrivateKeyPath)
        }
        if ($NewSSHKeyName) {
            if (!$NewSSHKeyPwd) {
                $GitAuthParams = $GitAuthParams.Add("NewSSHKeyName",$NewSSHKeyName)
            }
            else {
                $GitAuthParams = $GitAuthParams.Add("NewSSHKeyName",$NewSSHKeyName)
                $GitAuthParams = $GitAuthParams.Add("NewSSHKeyPwd",$NewSSHKeyPwd)
            }
        }
        if ($DownloadAndSetupDependencies) {
            $global:FunctionResult = "0"
            Setup-GitAuthentication @GitAuthParams -DownloadAndSetupDependencies
            if ($global:FunctionResult -eq "1") {
                Write-Verbose "The Setup-GitAuthentication function failed. Halting!"
                Write-Error "The Setup-GitAuthentication function failed. Halting!"
                $global:FunctionResult = "1"
                return
            }
        }
        else {
            $global:FunctionResult = "0"
            Setup-GitAuthentication @GitAuthParams
            if ($global:FunctionResult -eq "1") {
                Write-Verbose "The Setup-GitAuthentication function failed. Halting!"
                Write-Error "The Setup-GitAuthentication function failed. Halting!"
                $global:FunctionResult = "1"
                return
            }
        }
    }
    # Setup https
    if ($AuthMethod -eq "https") {
        $GitAuthParams = @{
            GitHubUserName = $GitHubUserName
            GitHubEmail = $GitHubEmail
            AuthMethod = $AuthMethod
            PersonalAccessToken = $PersonalAccessToken
        }
        $global:FunctionResult = "0"
        Setup-GitAuthentication @GitAuthParams
        if ($global:FunctionResult -eq "1") {
            Write-Verbose "The Setup-GitAuthentication function failed. Halting!"
            Write-Error "The Setup-GitAuthentication function failed. Halting!"
            $global:FunctionResult = "1"
            return
        }
    }
    if (!$AuthMethod) {
        Write-Host "GitHub Authentication still needs to be setup. Use the Setup-GitAuthentication function in the GitEnv Module."
    }

    Write-Host "Git Environment is ready."

    # Write-Host "See the following site for next steps:"
    # Write-Host "https://help.github.com/articles/set-up-git/"
}

function Clone-GitRepo {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False)]
        $GitRepoParentDirectory = $(Read-Host -Prompt "Please enter the full path to the directory that will contain the cloned Git repository."),

        [Parameter(Mandatory=$False)]
        [string]$GitHubUserName = $(Read-Host -Prompt "Please enter the GitHub UserName associated with the repo you would like to clone"),

        [Parameter(Mandatory=$False)]
        [string]$GitHubEmail,

        [Parameter(
            Mandatory=$False,
            ParameterSetName='PrivateRepos'
        )]
        $PersonalAccessToken,

        [Parameter(Mandatory=$False)]
        $RemoteGitRepoName,

        [Parameter(Mandatory=$False)]
        [switch]$CloneAllPublicRepos,

        [Parameter(
            Mandatory=$False,
            ParameterSetName='PrivateRepos'
        )]
        [switch]$CloneAllPrivateRepos,

        [Parameter(Mandatory=$False)]
        [switch]$CloneAllRepos
    )

    ##### BEGIN Variable/Parameter Transforms and PreRun Prep #####
    if ($PersonalAccessToken) {
        if ($PersonalAccessToken.GetType().FullName -eq "System.Security.SecureString") {
            # Convert SecureString to PlainText
            $PersonalAccessToken = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($PersonalAccessToken))
        }
    }
    
    # Make sure we have access to the git command
    if ($env:github_shell -ne $true -or !$(Get-Command git -ErrorAction SilentlyContinue)) {
        if (!$GitHubUserName) {
            $GitHubUserName = Read-Host -Prompt "Please enter your GitHub UserName"
        }
        if (!$GitHubEmail) {
            $GitHubEmail = Read-Host -Prompt "Please enter the GitHub Email address associated with $GitHubuserName"
        }
        $global:FunctionResult = "0"
        Initialize-GitEnvironment -GitHubUserName $GitHubUserName -GitHubEmail $GitHubEmail
        if ($global:FunctionResult -eq "1") {
            Write-Verbose "The Initialize-GitEnvironment function failed! Halting!"
            Write-Error "The Initialize-GitEnvironment function failed! Halting!"
            $global:FunctionResult = "1"
            return
        }
    }

    if (!$(Test-Path $GitRepoParentDirectory)) {
        Write-Verbose "The path $GitRepoParentDirectory was not found! Halting!"
        Write-Error "The path $GitRepoParentDirectory was not found! Halting!"
        $global:FunctionResult = "1"
        return
    }

    if ($CloneAllRepos -and !$PersonalAccessToken) {
        Write-Host "Please note that if you would like to clone both Public AND Private repos, you must use the -PersonalAccessToken parameter with the -CloneAllRepos switch."
    }

    $BoundParamsArrayOfKVP = $PSBoundParameters.GetEnumerator() | foreach {$_}

    $PrivateReposParamSetCheck = $($BoundParamsArrayOfKVP.Key -join "") -match "PersonalAccessToken|CloneAllPrivateRepos|CloneAllRepos"
    $NoPrivateReposParamSetCheck = $($BoundParamsArrayOfKVP.Key -join "") -match "CloneAllPublicRepos"
    if ($RemoteGitRepoName -and !$PersonalAccessToken) {
        $NoPrivateReposParamSetCheck = $true
    }

    # For Params that are part of the PrivateRepos Parameter Set...
    if ($PrivateReposParamSetCheck -eq $true) {
        if ($($CloneAllPrivateRepos -and $CloneAllRepos) -or 
        $($CloneAllPrivateRepos -and $RemoteGitRepoName) -or
        $($CloneAllPrivateRepos -and $CloneAllPublicRepos) -or 
        $($CloneAllRepos -and $RemoteGitRepoName) -or
        $($CloneAllRepos -and $CloneAllPublicRepos) -or
        $($CloneAllPublicRepos -and $RemoteGitRepoName) )  {
            Write-Verbose "Please use *either* -CloneAllRepos *or* -CloneAllPrivateRepos *or* -RemoteGitRepoName *or* -CloneAllPublicRepos! Halting!"
            Write-Error "Please use *either* -CloneAllRepos *or* -CloneAllPrivateRepos *or* -RemoteGitRepoName *or* -CloneAllPublicRepos! Halting!"
            $global:FunctionResult = "1"
            return
        }
    }
    # For Params that are part of the NoPrivateRepos Parameter Set...
    if ($NoPrivateReposParamSetCheck -eq $true) {
        if ($CloneAllPublicRepos -and $RemoteGitRepoName) {
            Write-Verbose "Please use *either* -CloneAllPublicRepos *or* -RemoteGitRepoName! Halting!"
            Write-Error "Please use *either* -CloneAllPublicRepos *or* -RemoteGitRepoName! Halting!"
            $global:FunctionResult = "1"
            return
        }
    }

    ##### END Variable/Parameter Transforms and PreRun Prep #####


    ##### BEGIN Main Body #####

    Push-Location $GitRepoParentDirectory

    if ($PrivateReposParamSetCheck -eq $true) {
        if ($PersonalAccessToken) {
            $PublicAndPrivateRepoObjects = Invoke-RestMethod -Uri "https://api.github.com/user/repos?access_token=$PersonalAccessToken"
            $PrivateRepoObjects = $PublicAndPrivateRepoObjects | Where-Object {$_.private -eq $true}
            $PublicRepoObjects = $PublicAndPrivateRepoObjects | Where-Object {$_.private -eq $false}
        }
        else {
            $PublicRepoObjects = Invoke-RestMethod -Uri "https://api.github.com/users/$GitHubUserName/repos"
        }
        if ($PublicRepoObject.Count -lt 1) {
            if ($RemoteGitRepo) {
                Write-Verbose "$RemoteGitRepo is either private or does not exist..."
            }
            else {
                Write-Warning "No public repositories were found!"
            }
        }
        if ($PrivateRepoObjects.Count -lt 1) {
            Write-Verbose "No private repositories were found!"
        }
        if ($($PublicRepoObjects + $PrivateRepoObjects).Count -lt 1) {
            Write-Verbose "No public or private repositories were found! Halting!"
            Write-Error "No public or private repositories were found! Halting!"
            Pop-Location
            $global:FunctionResult = "1"
            return
        }
        if ($RemoteGitRepoName) {
            if ($PrivateRepoObjects.Name -contains $RemoteGitRepoName) {
                $CloningOneOrMorePrivateRepos = $true
            }
        }
        if ($CloneAllPrivateRepos -or $($CloneAllRepos -and $PrivateRepoObjects -ne $null)) {
            $CloningOneOrMorePrivateRepos = $true
        }
        # If we're cloning a private repo, we're going to need Windows Credential Caching to avoid prompts
        if ($CloningOneOrMorePrivateRepos) {
            # Check the Windows Credential Store to see if we have appropriate credentials available already
            # If not, add them to the Windows Credential Store
            $FindCachedCredentials = Manage-StoredCredentials -ShoCred | Where-Object {
                $_.UserName -eq $GitHubUserName -and
                $_.Target -match "git"
            }
            if ($FindCachedCredentials.Count -gt 1) {
                Write-Warning "More than one set of stored credentials matches the UserName $GitHubUserName and contains the string 'git' in the Target property."
                Write-Host "Options are as follows:"
                # We do NOT want the Password for any creds displayed in STDOUT...
                # ...And it's possible that the GitHub PersonalAccessToken could be found in EITHER the Target Property OR the
                # Password Property
                $FindCachedCredentialsSansPassword = $FindCachedCredentials | foreach {
                    $PotentialPersonalAccessToken = $($_.Target | Select-String -Pattern "https://.*?@git").Matches.Value -replace "https://","" -replace "@git",""
                    if ($PotentialPersonalAccessToken -notmatch $GitHubUserName) {
                        $_.Target = $_.Target -replace $PotentialPersonalAccessToken,"<redacted>"
                        $_.PSObject.Properties.Remove('Password')
                        $_
                    }
                }
                for ($i=0; $i -lt $FindCachedCredentialsSansPassword.Count; $i++) {
                    "`nOption $i)"
                    $($($FindCachedCredentialsSansPassword[$i] | fl *) | Out-String).Trim()
                }
                $CachedCredentialChoice = Read-Host -Prompt "Please enter the Option Number that corresponds with the credentials you would like to use [0..$($FindCachedCredentials.Count-1)]"
                if ($(0..$($FindCachedCredentials.Count-1)) -notcontains $CachedCredentialChoice) {
                    Write-Verbose "Option Number $CachedCredentialChoice is not a valid Option Number! Halting!"
                    Write-Error "Option Number $CachedCredentialChoice is not a valid Option Number! Halting!"
                    Pop-Location
                    $global:FunctionResult = "1"
                    return
                }
                
                if (!$PersonalAccessToken) {
                    if ($FindCachedCredentials[$CachedCredentialChoice].Password -notmatch "oauth") {
                        $PersonalAccessToken = $FindCachedCredentials[$CachedCredentialChoice].Password
                    }
                    else {
                        $PersonalAccessToken = $($FindCachedCredentials[$CachedCredentialChoice].Target | Select-String -Pattern "https://.*?@git").Matches.Value -replace "https://","" -replace "@git",""
                    }
                }
            }
            if ($FindCachedCredentials.Count -eq $null -and $FindCachedCredentials -ne $null) {
                if (!$PersonalAccessToken) {
                    if ($FindCachedCredentials.Password -notmatch "oauth") {
                        $PersonalAccessToken = $FindCachedCredentials[$CachedCredentialChoice].Password
                    }
                    else {
                        $PersonalAccessToken = $($FindCachedCredentials.Target | Select-String -Pattern "https://.*?@git").Matches.Value -replace "https://","" -replace "@git",""
                    }
                }
            }
            if ($FindCachedCredentials -eq $null) {
                $CurrentGitConfig = git config --list
                if ($CurrentGitConfig -notcontains "credential.helper=wincred") {
                    git config --global credential.helper wincred
                }
                if (!$PersonalAccessToken) {
                    $PersonalAccessToken = Read-Host -Prompt "Please enter your GitHub Personal Access Token." -AsSecureString
                }

                # Alternate Params for GitHub https auth
                <#
                $ManageStoredCredsParams = @{
                    Target  = "git:https://$PersonalAccessToken@github.com"
                    User    = $PersonalAccessToken
                    Pass    = 'x-oauth-basic'
                    Comment = "Saved By Manage-StoredCredentials.ps1"
                }
                #>
                $ManageStoredCredsParams = @{
                    Target  = "git:https://$GitHubUserName@github.com"
                    User    = $GitHubUserName
                    Pass    = $PersonalAccessToken
                    Comment = "Saved By Manage-StoredCredentials.ps1"
                }
                Manage-StoredCredentials -AddCred @ManageStoredCredsParams
            }
        }

        if ($CloneAllPrivateRepos) {
            foreach ($RepoObject in $PrivateRepoObjects) {
                if (!$(Test-Path "$GitRepoParentDirectory\$($RepoObject.Name)")) {
                    if ($CloningOneOrMorePrivateRepos) {
                        $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
                        $ProcessInfo.WorkingDirectory = $GitRepoParentDirectory
                        $ProcessInfo.FileName = "git"
                        $ProcessInfo.RedirectStandardError = $true
                        $ProcessInfo.RedirectStandardOutput = $true
                        $ProcessInfo.UseShellExecute = $false
                        $ProcessInfo.Arguments = "clone $($RepoObject.html_url)"
                        $Process = New-Object System.Diagnostics.Process
                        $Process.StartInfo = $ProcessInfo
                        $Process.Start() | Out-Null
                        # Below $FinishedInAlottedTime returns boolean true/false
                        $FinishedInAlottedTime = $Process.WaitForExit(15000)
                        if (!$FinishedInAlottedTime) {
                            $Process.Kill()
                            Write-Verbose "git is prompting for UserName and Password, which means Credential Caching is not configured correctly! Halting!"
                            Write-Error "git is prompting for UserName and Password, which means Credential Caching is not configured correctly! Halting!"
                            Pop-Location
                            $global:FunctionResult = "1"
                            return
                        }
                        $stdout = $Process.StandardOutput.ReadToEnd()
                        $stderr = $Process.StandardError.ReadToEnd()
                        $AllOutput = $stdout + $stderr
                        Write-Host "##### BEGIN git clone Console Output #####"
                        Write-Host "$AllOutput"
                        Write-Host "##### END git clone Console Output #####"
                        
                    }
                }
                else {
                    Write-Verbose "The RemoteGitRepo $RemoteGitRepoName already exists under $GitRepoParentDirectory\$RemoteGitRepoName! Skipping!"
                    Write-Error "The RemoteGitRepo $RemoteGitRepoName already exists under $GitRepoParentDirectory\$RemoteGitRepoName! Skipping!"
                    $global:FunctionResult = "1"
                    break
                }
            }
        }
        if ($CloneAllPublicRepos) {
            foreach ($RepoObject in $PublicRepoObjects) {
                if (!$(Test-Path "$GitRepoParentDirectory\$($RepoObject.Name)")) {
                    git clone $RepoObject.html_url
                }
                else {
                    Write-Verbose "The RemoteGitRepo $RemoteGitRepoName already exists under $GitRepoParentDirectory\$RemoteGitRepoName! Skipping!"
                    Write-Error "The RemoteGitRepo $RemoteGitRepoName already exists under $GitRepoParentDirectory\$RemoteGitRepoName! Skipping!"
                    $global:FunctionResult = "1"
                    break
                }
            }
        }
        if ($CloneAllRepos) {
            foreach ($RepoObject in $($PublicRepoObjects + $PrivateRepoObjects)) {
                if (!$(Test-Path "$GitRepoParentDirectory\$($RepoObject.Name)")) {
                    if ($CloningOneOrMorePrivateRepos) {
                        $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
                        $ProcessInfo.WorkingDirectory = $GitRepoParentDirectory
                        $ProcessInfo.FileName = "git"
                        $ProcessInfo.RedirectStandardError = $true
                        $ProcessInfo.RedirectStandardOutput = $true
                        $ProcessInfo.UseShellExecute = $false
                        $ProcessInfo.Arguments = "clone $($RepoObject.html_url)"
                        $Process = New-Object System.Diagnostics.Process
                        $Process.StartInfo = $ProcessInfo
                        $Process.Start() | Out-Null
                        # Below $FinishedInAlottedTime returns boolean true/false
                        $FinishedInAlottedTime = $Process.WaitForExit(15000)
                        if (!$FinishedInAlottedTime) {
                            $Process.Kill()
                            Write-Verbose "git is prompting for UserName and Password, which means Credential Caching is not configured correctly! Halting!"
                            Write-Error "git is prompting for UserName and Password, which means Credential Caching is not configured correctly! Halting!"
                            Pop-Location
                            $global:FunctionResult = "1"
                            return
                        }
                        $stdout = $Process.StandardOutput.ReadToEnd()
                        $stderr = $Process.StandardError.ReadToEnd()
                        $AllOutput = $stdout + $stderr
                        Write-Host "##### BEGIN git clone Console Output #####"
                        Write-Host "$AllOutput"
                        Write-Host "##### END git clone Console Output #####"
                        
                    }
                    else {
                        git clone $RepoObject.html_url
                    }
                }
                else {
                    Write-Verbose "The RemoteGitRepo $RemoteGitRepoName already exists under $GitRepoParentDirectory\$RemoteGitRepoName! Skipping!"
                    Write-Error "The RemoteGitRepo $RemoteGitRepoName already exists under $GitRepoParentDirectory\$RemoteGitRepoName! Skipping!"
                    Pop-Location
                    $global:FunctionResult = "1"
                    break
                }
            }
        }
        if ($RemoteGitRepoName) {
            $RemoteGitRepoObject = $($PublicRepoObjects + $PrivateRepoObjects) | Where-Object {$_.Name -eq $RemoteGitRepoName}
            if ($RemoteGitRepoObject -eq $null) {
                Write-Verbose "Unable to find a public or private repository with the name $RemoteGitRepoName! Halting!"
                Write-Error "Unable to find a public or private repository with the name $RemoteGitRepoName! Halting!"
                Pop-Location
                $global:FunctionResult = "1"
                return
            }
            if (!$(Test-Path "$GitRepoParentDirectory\$($RemoteGitRepoObject.Name)")) {
                if ($CloningOneOrMorePrivateRepos) {
                    $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
                    $ProcessInfo.WorkingDirectory = $GitRepoParentDirectory
                    $ProcessInfo.FileName = "git"
                    $ProcessInfo.RedirectStandardError = $true
                    $ProcessInfo.RedirectStandardOutput = $true
                    $ProcessInfo.UseShellExecute = $false
                    $ProcessInfo.Arguments = "clone $($RemoteGitRepoObject.html_url)"
                    $Process = New-Object System.Diagnostics.Process
                    $Process.StartInfo = $ProcessInfo
                    $Process.Start() | Out-Null
                    # Below $FinishedInAlottedTime returns boolean true/false
                    $FinishedInAlottedTime = $Process.WaitForExit(15000)
                    if (!$FinishedInAlottedTime) {
                        $Process.Kill()
                        Write-Verbose "git is prompting for UserName and Password, which means Credential Caching is not configured correctly! Halting!"
                        Write-Error "git is prompting for UserName and Password, which means Credential Caching is not configured correctly! Halting!"
                        Pop-Location
                        $global:FunctionResult = "1"
                        return
                    }
                    $stdout = $Process.StandardOutput.ReadToEnd()
                    $stderr = $Process.StandardError.ReadToEnd()
                    $AllOutput = $stdout + $stderr
                    Write-Host "##### BEGIN git clone Console Output #####"
                    Write-Host "$AllOutput"
                    Write-Host "##### END git clone Console Output #####"
                    
                }
                else {
                    git clone $RemoteGitRepoObject.html_url
                }
            }
            else {
                Write-Verbose "The RemoteGitRepo $RemoteGitRepoName already exists under $GitRepoParentDirectory\$RemoteGitRepoName! Halting!"
                Write-Error "The RemoteGitRepo $RemoteGitRepoName already exists under $GitRepoParentDirectory\$RemoteGitRepoName! Halting!"
                Pop-Location
                $global:FunctionResult = "1"
                return
            }
        }
    }
    if ($NoPrivateReposParamSetCheck -eq $true) {
        $PublicRepoObjects = Invoke-RestMethod -Uri "https://api.github.com/users/$GitHubUserName/repos"
        if ($PublicRepoObjects.Count -lt 1) {
            Write-Verbose "No public repositories were found! Halting!"
            Write-Error "No public repositories were found! Halting!"
            Pop-Location
            $global:FunctionResult = "1"
            return
        }

        if ($CloneAllPublicRepos -or $CloneAllRepos) {
            foreach ($RepoObject in $PublicRepoObjects) {
                if (!$(Test-Path "$GitRepoParentDirectory\$($RepoObject.Name)")) {
                    git clone $RepoObject.html_url
                }
                else {
                    Write-Verbose "The RemoteGitRepo $RemoteGitRepoName already exists under $GitRepoParentDirectory\$RemoteGitRepoName! Skipping!"
                    Write-Error "The RemoteGitRepo $RemoteGitRepoName already exists under $GitRepoParentDirectory\$RemoteGitRepoName! Skipping!"
                    Pop-Location
                    $global:FunctionResult = "1"
                    break
                }
            }
        }
        if ($RemoteGitRepoName) {
            $RemoteGitRepoObject = $PublicRepoObjects | Where-Object {$_.Name -eq $RemoteGitRepoName}
            if ($RemoteGitRepoObject -eq $null) {
                Write-Verbose "Unable to find a public repository with the name $RemoteGitRepoName! Is it private? If so, use the -PersonalAccessToken parameter. Halting!"
                Write-Error "Unable to find a public repository with the name $RemoteGitRepoName! Is it private? If so, use the -PersonalAccessToken parameter. Halting!"
                Pop-Location
                $global:FunctionResult = "1"
                return
            }
            if (!$(Test-Path "$GitRepoParentDirectory\$($RemoteGitRepoObject.Name)")) {
                git clone $RemoteGitRepoObject.html_url
            }
            else {
                Write-Verbose "The RemoteGitRepo $RemoteGitRepoName already exists under $GitRepoParentDirectory\$RemoteGitRepoName! Halting!"
                Write-Error "The RemoteGitRepo $RemoteGitRepoName already exists under $GitRepoParentDirectory\$RemoteGitRepoName! Halting!"
                Pop-Location
                $global:FunctionResult = "1"
                return
            }
        }
    }

    Pop-Location

    ##### END Main Body #####

}


<#
.SYNOPSIS
    Copy script from working directory to local github repository. Optionally commit and push to GitHub.
.DESCRIPTION
    If your workflow involves using a working directory that is NOT an initialized git repo for first
    drafts of scripts/functions, this function will assist with "publishing" your script/function from
    your working directory to the appropriate local git repo. Additional parameters will commit all
    changes to the local git repo and push these deltas to the appropriate repo on GitHub.

.NOTES
    IMPORTANT NOTES

    1) Using the $gitpush switch runs the following git commands which effectively 
    commmit and push ALL changes made to the local git repo since the last commit.
    
    git -C $DestinationLocalGitRepoDir add -A
    git -C $DestinationLocalGitRepoDir commit -a -m "$gitmessage"
    git -C $DestinationLocalGitRepoDir push

    The only change made by this script is the copy/paste operation from working directory to
    the specified local git repo. However, other changes outside the scope of this function may
    have occurred since the last commit. EVERYTHING will be committed and pushed if the $gitpush
    switch is used.

    DEPENDENCEIES
        None

.PARAMETER SourceFilePath
    This parameter is MANDATORY.

    This parameter takes a string that represents a file path to the script/function that you
    would like to publish.

.PARAMETER DestinationLocalGitRepoName
    This parameter is MANDATORY.

    This parameter takes a string that represents the name of the Local Git Repository that
    your script/function will be copied to. This parameter is NOT a file path. It is just
    the name of the Local Git Repository.

.PARAMETER SigningCertFilePath
    This parameter is OPTIONAL.

    This parameter takes a string that represents a file path to a certificate that can be used
    to digitally sign your script/function.

.PARAMETER gitpush
    This parameter is OPTIONAL.

    This parameter is a switch. If it is provided in the command line, then the function will 
    not only copy the source script/function from the working directory to the Local Git Repo,
    it will also commit changes to the Local Git Repo and push updates the corresponding repo
    on GitHub.

.PARAMETER gitmessage
    This parameter is OPTIONAL.

    If the $gitpush parameter is used, this parameter is MANDATORY.

    This parameter takes a string that represents a message that accompanies a git commit
    operation. The message should very briefly describe the changes that were made to the
    Git Repository.

.EXAMPLE
    Publish-MyGitRepo -SourceFilePath "V:\powershell\testscript.ps1" `
    -DestinationLocalGitRepo "misc-powershell" `
    -SigningCertFilePath "R:\zero\ZeroCode.pfx" `
    -gitpush `
    -gitmessage "Initial commit for testscript.ps1" -Confirm
#>

function Publish-MyGitRepo {

    [CmdletBinding(
        DefaultParameterSetName='Parameter Set 1', 
        SupportsShouldProcess=$true,
        PositionalBinding=$true,
        ConfirmImpact='Medium'
    )]
    [Alias('pubgitrepo')]
    Param(
        [Parameter(Mandatory=$False)]
        [Alias("source")]
        [string]$SourceFilePath = $(Read-Host -Prompt "Please enter the full file path to the script that you would like to publish to your LOCAL GitHub Project Repository."),

        [Parameter(Mandatory=$False)]
        [Alias("dest")]
        [string]$DestinationLocalGitRepoName = $(Read-Host -Prompt "Please enter the name of the LOCAL Git Repo to which the script/function will be published."),

        [Parameter(Mandatory=$False)]
        [string]$GitHubUserName,

        [Parameter(Mandatory=$False)]
        [string]$GitHubEmail,

        [Parameter(Mandatory=$False)]
        [Alias("cert")]
        [string]$SigningCertFilePath,

        [Parameter(Mandatory=$False)]
        [Alias("push")]
        [switch]$gitpush,

        [Parameter(Mandatory=$False)]
        [Alias("message")]
        [string]$gitmessage
    )

    ##### BEGIN Parameter Validation #####
    # Make sure we have access to the git command
    if ($env:github_shell -ne $true -or !$(Get-Command git -ErrorAction SilentlyContinue)){
        if (!$GitHubUserName) {
            $GitHubUserName = Read-Host -Prompt "Please enter your GitHub UserName"
        }
        if (!$GitHubEmail) {
            $GitHubEmail = Read-Host -Prompt "Please enter the GitHub Email address associated with $GitHubuserName"
        }
        $global:FunctionResult = "0"
        Initialize-GitEnvironment -GitHubUserName $GitHubUserName -GitHubEmail $GitHubEmail
        if ($global:FunctionResult -eq "1") {
            Write-Verbose "The Initialize-GitEnvironment function failed! Halting!"
            Write-Error "The Initialize-GitEnvironment function failed! Halting!"
            $global:FunctionResult = "1"
            return
        }
    }

    # Valdate Git Repo Parent Directory $GitRepoParentDir
    if (! $GitRepoParentDir) {
        [string]$GitRepoParentDir = Read-Host -Prompt "Please enter the parent directory of your local gitrepo"
    }
    if (! $(Test-Path $GitRepoParentDir)) {
        Write-Warning "The path $env:GitHubParent was not found!"
        [string]$GitRepoParentDir = Read-Host -Prompt "Please enter the parent directory of your local gitrepo"
        if (! $(Test-Path $GitRepoParentDir)) {
            Write-Host "The path $env:GitHubParent was not found! Halting!"
            Write-Error "The path $env:GitHubParent was not found! Halting!"
            $global:FunctionResult = "1"
            return
        }
    }

    # Validate $SigningCertFilePath
    if ($SigningCertFilePath) {
        if (! $(Test-Path $SigningCertFilePath)) {
            Write-Warning "The path $SigningCertFilePath was not found!"
            [string]$SigningCertFilePath = Read-Host -Prompt "Please enter the file path for the certificate you would like to use to sign the script/function"
            if (! $(Test-Path $SigningCertFilePath)) {
                Write-Host "The path $SigningCertFilePath was not found! Halting!"
                Write-Error "The path $SigningCertFilePath was not found! Halting!"
                $global:FunctionResult = "1"
                return
            }
        }
    }

    # Validate $gitpush
    if ($gitpush) {
        if (! $gitmessage) {
            [string]$gitmessage = Read-Host -Prompt "Please enter a message to publish on git for this push"
        }
    }
    ##### END Parameter Validation #####


    ##### BEGIN Variable/Parameter Transforms #####
    $ScriptFileName = $SourceFilePath | Split-Path -Leaf
    $DestinationLocalGitRepoDir = "$GitRepoParentDir\$DestinationLocalGitRepoName"
    $DestinationFilePath = "$DestinationLocalGitRepoDir\$ScriptFileName"

    if ($SigningCertFilePath) {
        Write-Host "The parameter `$SigningCertFilePath was provided. Getting certificate data..."
        [System.Security.Cryptography.X509Certificates.X509Certificate]$SigningCert = Get-PfxCertificate $SigningCertFilePath

        $CertCN = $($($SigningCert.Subject | Select-String -Pattern "CN=[\w]+,").Matches.Value -replace "CN=","") -replace ",",""
    }

    ##### END Variable/Parameter Transforms #####


    ##### BEGIN Main Body #####
    if ($SigningCertFilePath) {
        if ($pscmdlet.ShouldProcess($SourceFilePath,"Signiing $SourceFilePath with certificate $CertCN")) {
            Set-AuthenticodeSignature -FilePath $SourceFilePath -cert $SigningCert -Confirm:$false
        }
    }
    if ($pscmdlet.ShouldProcess($SourceFilePath,"Copy $SourceFilePath to Local Git Repo $DestinationLocalGitRepoDir")) {
        Copy-Item -Path $SourceFilePath -Destination $DestinationFilePath -Confirm:$false
    }
    if ($gitpush) {
        if ($pscmdlet.ShouldProcess($DestinationLocalGitRepoName,"Push deltas in $DestinationLocalGitRepoName to GitHub")) {
            Set-Location $DestinationLocalGitRepoDir
            
            git -C $DestinationLocalGitRepoDir add -A
            git -C $DestinationLocalGitRepoDir commit -a -m "$gitmessage"
            git -C $DestinationLocalGitRepoDir push
        }
    }

    ##### END Main Body #####

}










# SIG # Begin signature block
# MIIMiAYJKoZIhvcNAQcCoIIMeTCCDHUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUUrfpJz7fxr8IAel4cF2Ecfys
# I6Sgggn9MIIEJjCCAw6gAwIBAgITawAAAB/Nnq77QGja+wAAAAAAHzANBgkqhkiG
# 9w0BAQsFADAwMQwwCgYDVQQGEwNMQUIxDTALBgNVBAoTBFpFUk8xETAPBgNVBAMT
# CFplcm9EQzAxMB4XDTE3MDkyMDIxMDM1OFoXDTE5MDkyMDIxMTM1OFowPTETMBEG
# CgmSJomT8ixkARkWA0xBQjEUMBIGCgmSJomT8ixkARkWBFpFUk8xEDAOBgNVBAMT
# B1plcm9TQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDCwqv+ROc1
# bpJmKx+8rPUUfT3kPSUYeDxY8GXU2RrWcL5TSZ6AVJsvNpj+7d94OEmPZate7h4d
# gJnhCSyh2/3v0BHBdgPzLcveLpxPiSWpTnqSWlLUW2NMFRRojZRscdA+e+9QotOB
# aZmnLDrlePQe5W7S1CxbVu+W0H5/ukte5h6gsKa0ktNJ6X9nOPiGBMn1LcZV/Ksl
# lUyuTc7KKYydYjbSSv2rQ4qmZCQHqxyNWVub1IiEP7ClqCYqeCdsTtfw4Y3WKxDI
# JaPmWzlHNs0nkEjvnAJhsRdLFbvY5C2KJIenxR0gA79U8Xd6+cZanrBUNbUC8GCN
# wYkYp4A4Jx+9AgMBAAGjggEqMIIBJjASBgkrBgEEAYI3FQEEBQIDAQABMCMGCSsG
# AQQBgjcVAgQWBBQ/0jsn2LS8aZiDw0omqt9+KWpj3DAdBgNVHQ4EFgQUicLX4r2C
# Kn0Zf5NYut8n7bkyhf4wGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwDgYDVR0P
# AQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAUdpW6phL2RQNF
# 7AZBgQV4tgr7OE0wMQYDVR0fBCowKDAmoCSgIoYgaHR0cDovL3BraS9jZXJ0ZGF0
# YS9aZXJvREMwMS5jcmwwPAYIKwYBBQUHAQEEMDAuMCwGCCsGAQUFBzAChiBodHRw
# Oi8vcGtpL2NlcnRkYXRhL1plcm9EQzAxLmNydDANBgkqhkiG9w0BAQsFAAOCAQEA
# tyX7aHk8vUM2WTQKINtrHKJJi29HaxhPaHrNZ0c32H70YZoFFaryM0GMowEaDbj0
# a3ShBuQWfW7bD7Z4DmNc5Q6cp7JeDKSZHwe5JWFGrl7DlSFSab/+a0GQgtG05dXW
# YVQsrwgfTDRXkmpLQxvSxAbxKiGrnuS+kaYmzRVDYWSZHwHFNgxeZ/La9/8FdCir
# MXdJEAGzG+9TwO9JvJSyoGTzu7n93IQp6QteRlaYVemd5/fYqBhtskk1zDiv9edk
# mHHpRWf9Xo94ZPEy7BqmDuixm4LdmmzIcFWqGGMo51hvzz0EaE8K5HuNvNaUB/hq
# MTOIB5145K8bFOoKHO4LkTCCBc8wggS3oAMCAQICE1gAAAH5oOvjAv3166MAAQAA
# AfkwDQYJKoZIhvcNAQELBQAwPTETMBEGCgmSJomT8ixkARkWA0xBQjEUMBIGCgmS
# JomT8ixkARkWBFpFUk8xEDAOBgNVBAMTB1plcm9TQ0EwHhcNMTcwOTIwMjE0MTIy
# WhcNMTkwOTIwMjExMzU4WjBpMQswCQYDVQQGEwJVUzELMAkGA1UECBMCUEExFTAT
# BgNVBAcTDFBoaWxhZGVscGhpYTEVMBMGA1UEChMMRGlNYWdnaW8gSW5jMQswCQYD
# VQQLEwJJVDESMBAGA1UEAxMJWmVyb0NvZGUyMIIBIjANBgkqhkiG9w0BAQEFAAOC
# AQ8AMIIBCgKCAQEAxX0+4yas6xfiaNVVVZJB2aRK+gS3iEMLx8wMF3kLJYLJyR+l
# rcGF/x3gMxcvkKJQouLuChjh2+i7Ra1aO37ch3X3KDMZIoWrSzbbvqdBlwax7Gsm
# BdLH9HZimSMCVgux0IfkClvnOlrc7Wpv1jqgvseRku5YKnNm1JD+91JDp/hBWRxR
# 3Qg2OR667FJd1Q/5FWwAdrzoQbFUuvAyeVl7TNW0n1XUHRgq9+ZYawb+fxl1ruTj
# 3MoktaLVzFKWqeHPKvgUTTnXvEbLh9RzX1eApZfTJmnUjBcl1tCQbSzLYkfJlJO6
# eRUHZwojUK+TkidfklU2SpgvyJm2DhCtssFWiQIDAQABo4ICmjCCApYwDgYDVR0P
# AQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBS5d2bhatXq
# eUDFo9KltQWHthbPKzAfBgNVHSMEGDAWgBSJwtfivYIqfRl/k1i63yftuTKF/jCB
# 6QYDVR0fBIHhMIHeMIHboIHYoIHVhoGubGRhcDovLy9DTj1aZXJvU0NBKDEpLENO
# PVplcm9TQ0EsQ049Q0RQLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENOPVNl
# cnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9emVybyxEQz1sYWI/Y2VydGlmaWNh
# dGVSZXZvY2F0aW9uTGlzdD9iYXNlP29iamVjdENsYXNzPWNSTERpc3RyaWJ1dGlv
# blBvaW50hiJodHRwOi8vcGtpL2NlcnRkYXRhL1plcm9TQ0EoMSkuY3JsMIHmBggr
# BgEFBQcBAQSB2TCB1jCBowYIKwYBBQUHMAKGgZZsZGFwOi8vL0NOPVplcm9TQ0Es
# Q049QUlBLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENOPVNlcnZpY2VzLENO
# PUNvbmZpZ3VyYXRpb24sREM9emVybyxEQz1sYWI/Y0FDZXJ0aWZpY2F0ZT9iYXNl
# P29iamVjdENsYXNzPWNlcnRpZmljYXRpb25BdXRob3JpdHkwLgYIKwYBBQUHMAKG
# Imh0dHA6Ly9wa2kvY2VydGRhdGEvWmVyb1NDQSgxKS5jcnQwPQYJKwYBBAGCNxUH
# BDAwLgYmKwYBBAGCNxUIg7j0P4Sb8nmD8Y84g7C3MobRzXiBJ6HzzB+P2VUCAWQC
# AQUwGwYJKwYBBAGCNxUKBA4wDDAKBggrBgEFBQcDAzANBgkqhkiG9w0BAQsFAAOC
# AQEAszRRF+YTPhd9UbkJZy/pZQIqTjpXLpbhxWzs1ECTwtIbJPiI4dhAVAjrzkGj
# DyXYWmpnNsyk19qE82AX75G9FLESfHbtesUXnrhbnsov4/D/qmXk/1KD9CE0lQHF
# Lu2DvOsdf2mp2pjdeBgKMRuy4cZ0VCc/myO7uy7dq0CvVdXRsQC6Fqtr7yob9NbE
# OdUYDBAGrt5ZAkw5YeL8H9E3JLGXtE7ir3ksT6Ki1mont2epJfHkO5JkmOI6XVtg
# anuOGbo62885BOiXLu5+H2Fg+8ueTP40zFhfLh3e3Kj6Lm/NdovqqTBAsk04tFW9
# Hp4gWfVc0gTDwok3rHOrfIY35TGCAfUwggHxAgEBMFQwPTETMBEGCgmSJomT8ixk
# ARkWA0xBQjEUMBIGCgmSJomT8ixkARkWBFpFUk8xEDAOBgNVBAMTB1plcm9TQ0EC
# E1gAAAH5oOvjAv3166MAAQAAAfkwCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwx
# CjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGC
# NwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFFVjgoVpA5hMQMqg
# 61N5fDI/kVG3MA0GCSqGSIb3DQEBAQUABIIBAEzS2lspsfWFxK9N+xwqZcKogk5f
# 3ERMdAcXZd0t993/i+3AcfZIUA4D51X1HaX1/IIfqdQcpbPAX8uWLuN7OZvIRGO8
# a7s5pfXJJTG2h2ECbN1D59Uqton8Euq+fwf7pZymhNDTi6JInv+nvP11Uq3/ZV4a
# l6OeAkqnzbyYB3ov8tUPcxMS9rj5brSCgwPHPxSqRzSm7sbu46KvczrwAWsH+xjW
# eynOclhZsnV2rRR0GX+tl0vuER7s+yXbnTWidjvo3J11RfHTmb2mHgPma55vpkkp
# 9gBECPf0ejt9Gk63k5rYfgyKJdaF7IJMsqfYB4PW5NkfwC4gKYjPr2YLR5c=
# SIG # End signature block

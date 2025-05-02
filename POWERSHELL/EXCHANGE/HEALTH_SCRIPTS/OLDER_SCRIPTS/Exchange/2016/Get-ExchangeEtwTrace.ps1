# Script to work with ETW files and sessions.
# Copyright (c) 2011 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#
# .SYNOPSIS
# Script to work with ETW files and sessions.
# 
# .DESCRIPTION
# Script to work with ETW files and sessions.
[CmdletBinding()]
Param
(
    [Parameter(Mandatory=$true, ParameterSetName = "Convert", HelpMessage="Set this switch to do file converstion.")]
    [Switch]
    $Convert,

    [Parameter(Mandatory=$true, ParameterSetName = "Flush", HelpMessage="Set this switch to flush the ETW session.")]
    [Switch]
    $Flush,

    [Parameter(Mandatory=$true, ParameterSetName = "Stop", HelpMessage="Set this switch to stop the ETW session.")]
    [Switch]
    $Stop,

    [Parameter(Mandatory=$true, ParameterSetName = "StopAll", HelpMessage="Set this switch to start all ETW sessions.")]
    [Switch]
    $StopAll,

    [Parameter(Mandatory=$true, ParameterSetName = "StartAll", HelpMessage="Set this switch to stop all ETW sessions.")]
    [Switch]
    $StartAll,

    [Parameter(Mandatory=$true, ParameterSetName = "Flush")]
    [Parameter(Mandatory=$true, ParameterSetName = "Stop")]
    [Parameter(Mandatory=$true, ParameterSetName = "Convert", HelpMessage="Log name from the trace config.")]
    [String]
    $LogName = "QueryRequestLog",

    [Parameter(ParameterSetName = "Convert", HelpMessage="Full file path to folder containing ETL trace files")]
    [String]
    $TraceFolderPath = "D:\ExchangeLogs\Query",

    [Parameter(ParameterSetName = "Convert", HelpMessage="File name search pattern")]
    [String]
    $FileNamePattern = $LogName + "*.etl",

    [Parameter(ParameterSetName = "Convert", HelpMessage="Start time filter.")]
    [String]
    $StartTime = $null,

    [Parameter(ParameterSetName = "Convert", HelpMessage="End time filter.")]
    [String]
    $EndTime = $null,

    [Parameter(ParameterSetName = "Convert", HelpMessage="Whether to output to the file..")]
    [Switch]
    $ToFile = $false,

    [Parameter(Mandatory=$true, ParameterSetName = "Flush")]
    [Parameter(Mandatory=$true, ParameterSetName = "Stop")]
    [Parameter(ParameterSetName = "Convert", HelpMessage="Build to get dependencies from.")]
    [String]
    $Build
)

# Gets the directory where excuting script is located.
function Get-ScriptDirectory
{
  $Invocation = (Get-Variable MyInvocation -Scope 1).Value
  Split-Path $Invocation.MyCommand.Path
}

# Goes through possible module location trying to import the module
# If module is not found it's copied and put close to the script.
function Load-Module([string]$moduleName)
{
    $moduleDll = $moduleName+".dll";
    $currentDir = Split-Path $script:MyInvocation.MyCommand.Path
    $fullName = [System.IO.Path]::Combine($currentDir, $moduleDll);
    if([System.IO.File]::Exists($fullName))
    {
        Import-Module $fullName;
        return;
    }

    $currentLocation = Get-Location;
    $fullName = [System.IO.Path]::Combine($currentLocation.Path, $moduleDll);
    if([System.IO.File]::Exists($fullName))
    {
        Import-Module $fullName;
        return;
    }

    [string]$paths = $env:Path;
    foreach($dir in $paths.Split(@(";"), [System.StringSplitOptions]::RemoveEmptyEntries))
    {
        $fullName = [System.IO.Path]::Combine($dir, $moduleDll);
        if([System.IO.File]::Exists($fullName))
        {
            Import-Module $fullName;
            return;
        }
    }

    if([String]::IsNullOrEmpty($Build))
    {
        $Build = Get-Build
    }

    $loadPath = [System.String]::Format("\\redmond\exchange\Build\E16\{0}\sources\distrib\PRIVATE\BIN\debug\amd64", $Build);
    Write-Host "Loading the dependency from $loadPath";
    $fullName = [System.IO.Path]::Combine($loadPath, $moduleDll);
    if(![System.IO.File]::Exists($fullName))
    {
        throw "File is not found on the machine or release shares";
    }

    $copyTarget = [System.IO.Path]::Combine($currentDir, $moduleDll);
    copy $fullName $copyTarget;
    Import-Module $copyTarget;
}

function Get-Build()
{
    Write-Host "Getting the machine build."
    $command = Get-Command -Name "Get-ExchangeServer" -ErrorAction SilentlyContinue;
    if($command -eq $null)
    {
        Write-Host "Get-ExchangeServer command not found. Build is being set to LATEST."
        return "LATEST";
    }

    Write-Host "Getting the exchange version";
    $server = Get-ExchangeServer $env:COMPUTERNAME;
    $match = [System.Text.RegularExpressions.Regex]::Match($server.AdminDisplayVersion, "\(Build (\d+)");
    if(!$match.Success -or $match.Groups.Count -lt 2)
    {
        Write-Host "Couldn't get the version from the server version. Build is being set to latest." 
        return "LATEST";
    }

    $version = "15.01.0" + $match.Groups[1].Value + ".000";
    Write-Host "Build is being set to $version"; 
    return $version;
}

# Loads binaries needed for this cmdlet to work
function Load-Dependencies()
{
    Write-Host "Loading conversion dependencies"
    $modules = 
    @(
        "Microsoft.Exchange.Data.Common",
        "Microsoft.Exchange.Diagnostics", 
        "Microsoft.Office.BigData.DataLoader.Interface", 
        "Microsoft.Office.BigData.DataLoader.LogTransformerCommon",
        "Microsoft.Office.BigData.DataLoader.EventTracingTransformer",
        "Microsoft.Bond.Interfaces"
        "Microsoft.Exchange.Query.Core"
     );

     foreach($module in $modules)
     {
        $moduleLoaded = $false;
        foreach($loadedModule in Get-Module)
        {
            if($loadedModule.Name -eq $module)
            {
                $moduleLoaded = $true;
                break;
            }
        }

        if(!$moduleLoaded)
        {
            Load-Module $module;
        }
     }
}

# Gets ETL files with a specific search pattern.
function Get-Files
{
    if(![System.IO.Directory]::Exists($TraceFolderPath))
    {
        throw "Directory '$TraceFolderPath' doesn't exist.";
    }

    return [System.IO.Directory]::GetFiles($TraceFolderPath, $FileNamePattern);
}

# Decodes ETW files
function Decode-Files()
{
    Write-Host "Decoding files."
    Load-Dependencies;
    $regexString = '("(?:[^"]|"")*"|[^",]*),'
    $regex = New-Object System.Text.RegularExpressions.Regex -ArgumentList $regexString, Compiled;
    $files = Get-Files
    $columnList = @('InternalSequence', 'InternalTimestamp');

    # calculate columns
    foreach($entry in [Microsoft.Exchange.Management.DataMining.TraceDefinitions]::Entries.Values)
    {
        foreach($traceRecord in $entry)
        {
            if($traceRecord.FilePrefix.CompareTo($LogName) -eq 0)
            {
                foreach($column in $traceRecord.Columns)
                {
                    $columnList+=$column.Name
                }

                $columns = [System.String]::Join(',', $columnList);
            }
        }
    }

    # For each file we will run the decode to get a list of CSV strings from the conversion.
    # Later we will split those and convert to PS objects to filter.
    foreach($file in $files)
    {
        $targetPath = [System.IO.Path]::GetDirectoryName($file);
        $targetName = [System.IO.Path]::GetFileNameWithoutExtension($file) + ".log";
        $target = [System.IO.Path]::Combine($targetPath, $targetName);

        if($ToFile)
        {
            $columns > $target;
        }

        $start = [DateTime]::MinValue;
        $end = [DateTime]::MaxValue;
        $dateColumnIndex = 15;
        $shoulFilter = $false;
        if(![string]::IsNullOrWhiteSpace($StartTime))
        {
            $start = [DateTime]::Parse($StartTime);
            $shouldFilter = $true;
        }

        if(![string]::IsNullOrWhiteSpace($EndTime))
        {
            $end = [DateTime]::Parse($EndTime);
            $shouldFilter = $true;
        }

        $decoder = New-Object Microsoft.Office.BigData.DataLoader.EventTraceLineReaderV2 -ArgumentList $LogName, $file
        $return = @();
        while(!$decoder.Eof)
        {
            $parts = $decoder.Read();

            # add a comma here to help the parser.
            $split = $parts[0] + ",";
            $matches = $regex.Matches($split);
            $values = @{};
            for($i=0;$i -lt $columnList.Length;$i++)
            {
                $value = $matches[$i].ToString();
                $values.Add($columnList[$i], $value.Trim(',').Trim('"'));
            }

            $obj = New-Object PsObject -Property $values;

            if($shouldFilter)
            {
                $time = [DateTime]::Parse($obj.StartTime);
                if($time -lt $start -or $time -gt $end)
                {
                    continue;
                }
            }

            if($ToFile)
            {
                $parts[0] >> $target;
            }
            else
            {
                $obj
            }                
        }
    }
}

# Imports some C# types to work with ETW sessions.
function Import-TraceProperties()
{
    Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public static class PInvokes
    {
        public static EventTraceProperties CreateProperties(string sessionName)
        {
            EventTraceProperties props = new EventTraceProperties();
            props.loggerName = sessionName;
            props.etp.wnode.bufferSize = (uint)Marshal.SizeOf(props);
            props.etp.logFileNameOffset = (uint)Marshal.OffsetOf(typeof(EventTraceProperties), "logFileName");
            props.etp.loggerNameOffset = (uint)Marshal.OffsetOf(typeof(EventTraceProperties), "loggerName");
            return props;
        }

        [DllImport("advapi32.dll", EntryPoint = "FlushTraceW", CharSet = CharSet.Unicode)]
        public static extern uint FlushTrace(
            [In] Int64 sessionHandle,
            [In] string sessionName,
            [In, Out] ref EventTraceProperties properties);

        [DllImport("advapi32.dll", EntryPoint = "StopTraceW", CharSet = CharSet.Unicode)]
        public static extern uint StopTrace(
            [In] int sessionHandle,
            [In] string sessionName,
            [In, Out] ref EventTraceProperties properties);

        [StructLayout(LayoutKind.Explicit, CharSet = CharSet.Unicode)]
        public struct EventTraceProperties
        {
            [FieldOffset(0)]
            public EventTraceProperties_Inner etp;

            /// <summary>buffer for the logger name,
            /// offset above should point to this buffer</summary>
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 1024)]
            [FieldOffset(120)]
            public string loggerName;

            /// <summary>buffer for the logfile name,
            /// offset above should point to this buffer</summary>
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 1024)]
            [FieldOffset(2168)]
            public string logFileName;
        }

        /// <summary>
        /// The EVENT_TRACE_PROPERTIES structure used for many P/Invoked tracing
        /// API calls.
        /// </summary>
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct EventTraceProperties_Inner
        {
            public WNodeHeader wnode;
            public uint bufferSize;
            public uint minimumBuffers;
            public uint maximumBuffers;
            public uint maximumFileSize;
            public uint logFileMode;
            public uint flushTimer;
            public uint enableFlags;
            public int ageLimit;
            public uint numberOfBuffers;
            public uint freeBuffers;
            public uint eventsLost;
            public uint buffersWritten;
            public uint logBuffersLost;
            public uint realTimeBuffersLost;
            public IntPtr loggerThreadId;
            public uint logFileNameOffset;
            public uint loggerNameOffset;
        }
    
        /// <summary>
        /// This struct is defined in wmistr.h. It contains a couple unions and
        /// therefore has to be defined as LayoutKind.Explicit.
        /// </summary>
        [StructLayout(LayoutKind.Explicit)]
        public struct WNodeHeader
        {
            [FieldOffset(0)]
            public uint bufferSize;

            [FieldOffset(4)]
            public uint providerId;

            /// <summary>
            /// "version" and "linkage" are an embedded struct in a union.
            /// But all members of the Union are 64-bit
            /// </summary>
            [FieldOffset(8)]
            public uint version;

            /// <summary>
            /// "version" and "linkage" are an embedded struct in a union.
            /// But all members of the Union are 64-bit
            /// </summary>
            [FieldOffset(12)]
            public uint linkage;

            /// <summary>
            /// kernelHandle is of type HANDLE = pvoid. This is 32 bits on x86
            /// and 64 bits on x64. In addition, it is part of a union where the
            /// largest member is a 64-bit LARGE_INTEGER: TimeStamp (which we are
            /// not interested in, and therefore is not defined). kernelHandle
            /// is defined as a *VARIABLE* size IntPtr, but the offset of the next
            /// member is actually a *FIXED* 64-bits, because the LARGE_INTEGER
            /// union member is fixed.
            /// </summary>
            [FieldOffset(16)]
            public IntPtr kernelHandle;

            [FieldOffset(24)]
            public Guid guid;

            [FieldOffset(40)]
            public uint clientContext;

            [FieldOffset(44)]
            public uint flags;
        }
    }
"@
}

# Checks windows error code for success.
function Check-Error($errorCode)
{
    # Error 4201 happens if the trace session is not running. This error doesn't mean much when we 
    # are converting since a) conversion can happen on dev machine b) it's OK for the session not to be running
    if(($Convert -and $errorCode -ne 4201) -or (-not $Convert -and $errorCode -ne 0))
    {
        throw $errorCode;
    }

    Write-Host "Success";
}

function Stop-AllTraces
{
    [Microsoft.Exchange.Query.Core.Diagnostics.Logging.LoggerManager]::SessionManager.StopAll();
}

function Start-AllTraces
{
    [Microsoft.Exchange.Query.Core.Diagnostics.Logging.LoggerManager]::SessionManager.StartAll();
}

# Flushes a trace session.
function Flush-Trace()
{
    Write-Host "Flushing";
    Import-TraceProperties;
    $etp = [PInvokes]::CreateProperties($LogName);
    $errorCode = [PInvokes]::FlushTrace(0, $LogName, [ref] $etp);
    Check-Error $errorCode;
}

# Stops a trace session.
function Stop-Trace()
{
    Write-Host "Stopping trace session.";
    Import-TraceProperties;
    $etp = [PInvokes]::CreateProperties($LogName);
    $errorCode = [PInvokes]::StopTrace(0, $LogName, [ref] $etp);
    Check-Error $errorCode;
}

$appDomainSetup = New-Object "System.AppDomainSetup";
$appDomainSetup.ShadowCopyFiles = "true";
$appDomain = [System.AppDomain]::CreateDomain("tempDomain", $null, $appDomainSetup);

if($Convert)
{
    Decode-Files;
    return;
}

if($Flush)
{
    Flush-Trace;
    return;
}

if($Stop)
{
    Stop-Trace;
    return;
}

if($StopAll)
{
    Load-Dependencies
    Stop-AllTraces;
    return;
}

if($StartAll)
{
    Load-Dependencies
    Start-AllTraces;
    return;
}

[System.AppDomain]::Unload($appDomain);

# SIG # Begin signature block
# MIIdsAYJKoZIhvcNAQcCoIIdoTCCHZ0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU6hjh8WCf6/I+haJox4vj4fHX
# Fs+gghhkMIIEwzCCA6ugAwIBAgITMwAAAJzu/hRVqV01UAAAAAAAnDANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTMw
# WhcNMTcwNjMwMTkyMTMwWjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OjU4NDctRjc2MS00RjcwMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAzCWGyX6IegSP
# ++SVT16lMsBpvrGtTUZZ0+2uLReVdcIwd3bT3UQH3dR9/wYxrSxJ/vzq0xTU3jz4
# zbfSbJKIPYuHCpM4f5a2tzu/nnkDrh+0eAHdNzsu7K96u4mJZTuIYjXlUTt3rilc
# LCYVmzgr0xu9s8G0Eq67vqDyuXuMbanyjuUSP9/bOHNm3FVbRdOcsKDbLfjOJxyf
# iJ67vyfbEc96bBVulRm/6FNvX57B6PN4wzCJRE0zihAsp0dEOoNxxpZ05T6JBuGB
# SyGFbN2aXCetF9s+9LR7OKPXMATgae+My0bFEsDy3sJ8z8nUVbuS2805OEV2+plV
# EVhsxCyJiQIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFD1fOIkoA1OIvleYxmn+9gVc
# lksuMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAFb2avJYCtNDBNG3nxss1ZqZEsphEErtXj+MVS/RHeO3TbsT
# CBRhr8sRayldNpxO7Dp95B/86/rwFG6S0ODh4svuwwEWX6hK4rvitPj6tUYO3dkv
# iWKRofIuh+JsWeXEIdr3z3cG/AhCurw47JP6PaXl/u16xqLa+uFLuSs7ct7sf4Og
# kz5u9lz3/0r5bJUWkepj3Beo0tMFfSuqXX2RZ3PDdY0fOS6LzqDybDVPh7PTtOwk
# QeorOkQC//yPm8gmyv6H4enX1R1RwM+0TGJdckqghwsUtjFMtnZrEvDG4VLA6rDO
# lI08byxadhQa6k9MFsTfubxQ4cLbGbuIWH5d6O4wggYHMIID76ADAgECAgphFmg0
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
# bWrJUnMTDXpQzTGCBLYwggSyAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCByjAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUzCp5eRhK1wOscac+JUtOmfvr5EswagYKKwYB
# BAGCNwIBDDFcMFqgMoAwAEcAZQB0AC0ARQB4AGMAaABhAG4AZwBlAEUAdAB3AFQA
# cgBhAGMAZQAuAHAAcwAxoSSAImh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9leGNo
# YW5nZSAwDQYJKoZIhvcNAQEBBQAEggEAGGvKWxWBbfzFBRK9tzrq8hK9+nP6IJjk
# S3i5TWdW2yVg5WNYl5cN0nRS+IbOkp/B90GvJt/9CUNGF3UN1yhZv1chyHcHclTO
# zc9L79qYFR6svQgXIFTCt4gYSjhQe+5HRSsLjBpUSfd1oQBB8P9nNzTpH7P7WOen
# mf9HzKOO9mZ+1Dovom06lUXcibu/ypxbvNaWlTUh4Z8YkM8ClLQlekpVLhfGvk0R
# XL9Eb6suPlJfwyKe9o4ZzFIBj0PEt5K/eKTULz/ORVGBq2NtU4SdlO34QQ4uqMDp
# 4I5AnYk+gG+Io5M9ZoMNa2PeMOfcL6ShFKJofBdV+YGtn4dEIwPjx6GCAigwggIk
# BgkqhkiG9w0BCQYxggIVMIICEQIBATCBjjB3MQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQ
# Q0ECEzMAAACc7v4UValdNVAAAAAAAJwwCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJ
# AzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTE2MDkwMzE4NDMyMFowIwYJ
# KoZIhvcNAQkEMRYEFMnLDYT+OtklLwrF+MgNEoEz4WsfMA0GCSqGSIb3DQEBBQUA
# BIIBAB73kF6Z0MBH1eQK8cm5Qmt0BwhS5ajTtEpLhZUL8yCVMn5kkTTDLl0C0B25
# RYWwONheGZYsm38Jncpv4JCsUZ5SeazL7dCKVyjzvSTPEPwv+t61HfIcD2VToLWA
# eyO0CqgkpHZFMW4tM0UwTVs9THQ6UDdEnOVOG6qhWferOP/MUXGxqjh6l8A5Dvbw
# c1yWcp6lutKZgs9qEZ/bxxU+6bEivm3LoihRvG0db4cGZncZiYcYsl/5U6Gbd2Hw
# Cu5/aZwtTkw0rKVppTaBQuHxsoo0IGWR9bBFTAgO8FMZ7sd2lhYBqgdUdQpvGGp+
# G+81vUVZuCgIeERpSkw5KGwuBt8=
# SIG # End signature block

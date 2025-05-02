# .SYNOPSIS
# This script is used for hybrid to support site mailbox created in Office 365 gets
# synced into on-premise environment.
#
# Copyright (c) 2012 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#
# .DESCRIPTION
# This script is used for hybrid to support site mailbox created in Office 365 gets
# synced into on-premise environment.
# It runs two steps to pull changes from Office 365 to on-premise environment.
# Step 1: run Import-SyncSiteMailbox.ps1 to pull changes from Office 365 to a
# local cache csv file.
# Step 2: run Export-SyncSiteMailbox.ps1 to export site mailbox changes to on-premise
# active directory.
#
# How to configure schedule task against this script?
# We expect customer to run this script daily to ensure seamless site mailbox
# functions when site mailboxes are created in Office 365. Customer needs to use "Task
# Schedule" to set up repeated task.
#
# Prerequisites:
# - An Exchange 2013 Back End server.
# - A windows account with permission to change local active directory.
# - Microsoft Online Service Module is installed on that server.
#    It could be installed from here:
#    http://onlinehelp.microsoft.com/en-us/office365-enterprises/ff652560.aspx
# - An account which has organization and recipient read-only permissions to Exchange online
#    The account can run Get-AcceptedDomain, Get-Recipients, Get-SiteMailbox.
# - The same account has read-only permission to Microsoft Online service
#    The account can run Get-MsolUser.
#
# Configurations:
# - Use windows account logon the 2013 Back End server.
# - Create a folder on that machine to host log files and cached site mailbox csv file.
# - Go to control panel, find credential manager; Add a generic credential for Exchange online
#    account, you could use your tenant name as key, like contoso.
# - Go to control panel, create a basic task.
#    Input the name, like "sync site mailbox" and choose dialy schedule;
#    In Program/Script box find the PowerShell.exe, like
#     C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
#    In arguments box, type in .\SyncSiteMailbox.ps1 -WorkingFolder <folder> -TenantCredentialKey <contoso>
#    In start in box, type the path where the script is located, like
#     C:\Program Files\Microsoft\Exchange Server\V15\Scripts
#
# Troubleshooting:
# Sync log will indicate when the sync is triggered; Export and import logs will indicate how
# each site mailbox is synced.
#
# .PARAMETER WorkingFolder
# Specify where to keep local cache file and sync logs.
#
# TenantCredentialKey: 
# Specify the key to retrieve credential from windows generic credential manager. You
# need manually add the credential into credential manager first.
#
# .NOTES
# This is a temporty solution to enable on-premise users access site mailbox created
# in Office 365. When DirSync with user object creation feature enabled in future
# release, this script will be abandoned.
#
# This script will call two major parts, import and export.
#
# Import-SyncSiteMailbox.ps1 loads delta changes of site mailbox created Office 365
# and persists into a local cache file "SyncSiteMailboxes.csv".
# The delta changes is pulled by using Get-Recipient with WhenChangedUtc filter. After
# getting changed site mailboxes, we need to calcuate its external email address so as
# to convert it as mail user. This requires we have hybrid configured correctly
# since we use coexistent domain to do calculation. Before we allow to create this mail
# user in on-premise active directory, we need to ensure MSO already syncs this user
# object so that the user object created in on-premise can soft-match with it.
# Import-SyncSiteMailbox.ps1 also checks if a site mailbox has been deleted from Exchange
# online according to interval of $DeletionCheckInternval in days in SyncSiteMailboxLibrary.ps1
# 
# Export-SyncSiteMailbox.ps1 loads changes from cache file "SyncSiteMailboxes.csv" and
# commit changes into local active directory by using Set/New-SyncMailUser, Remove-MailUser.
#

PARAM
(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullorEmpty()]
    [string] $WorkingFolder,
    
    [Parameter(Mandatory = $true)]
    [ValidateNotNullorEmpty()]
    [string] $TenantCredentialKey
)

######################################################################
# Import common functions
######################################################################
# Stop the script when any error happens
$ErrorActionPreference = "Stop"

# Import library file
. ".\SyncSiteMailboxLibrary.ps1"

######################################################################
# Prepare log functions and log file
######################################################################
# Verify working foler exists so as to create log file
if ( -not (Test-Path -Path $WorkingFolder))
{
    Write-Error "$WorkingFolder doesn't exist, please create it first."
}

$logFile = CreateLogFile "Sync" $WorkingFolder
if([string]::IsNullOrEmpty($logFile))
{
    Write-Error "$logFile cannot be created, please ensure you have write permission of $WorkingFolder."
}

#
# Write an information log entry
#
function WriteInfoLog([string] $logData)
{
    WriteLog $logData "INFO" $logFile
}

#
# Write an error log entry
#
function WriteErrorLog([string] $logData, [bool] $throwError = $true)
{
    WriteLog $logData "ERRR" $logFile
    if ($throwError)
    {
        Write-Error $logData
    }
} 

######################################################################
# API to load credential from generic credential store
######################################################################
$CredManager = @"
using System;
using System.Net;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace SyncSiteMailbox
{
    /// <summary>
    /// </summary>
    public class CredManager
    {
        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "CredReadW")]
        public static extern bool CredRead([MarshalAs(UnmanagedType.LPWStr)] string target, [MarshalAs(UnmanagedType.I4)] CRED_TYPE type, UInt32 flags, [MarshalAs(UnmanagedType.CustomMarshaler, MarshalTypeRef = typeof(CredentialMarshaler))] out Credential cred);
        
        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Auto, EntryPoint = "CredFree")]
        public static extern void CredFree(IntPtr buffer);

        /// <summary>
        /// </summary>
        public enum CRED_TYPE : uint
        {
            /// <summary>
            /// </summary>
            CRED_TYPE_GENERIC = 1,

            /// <summary>
            /// </summary>
            CRED_TYPE_DOMAIN_PASSWORD = 2,

            /// <summary>
            /// </summary>
            CRED_TYPE_DOMAIN_CERTIFICATE = 3,

            /// <summary>
            /// </summary>
            CRED_TYPE_DOMAIN_VISIBLE_PASSWORD = 4,

            /// <summary>
            /// </summary>
            CRED_TYPE_MAXIMUM = 5, // Maximum supported cred type
        }
        
        /// <summary>
        /// </summary>
        public enum CRED_PERSIST : uint
        {
            /// <summary>
            /// </summary>
            CRED_PERSIST_SESSION = 1,

            /// <summary>
            /// </summary>
            CRED_PERSIST_LOCAL_MACHINE = 2,

            /// <summary>
            /// </summary>
            CRED_PERSIST_ENTERPRISE = 3
        }
        
        /// <summary>
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        internal struct CREDENTIAL
        {
            internal UInt32 flags;
            internal CRED_TYPE type;
            [MarshalAs(UnmanagedType.LPWStr)]
            internal string targetName;
            [MarshalAs(UnmanagedType.LPWStr)]
            internal string comment;
            internal System.Runtime.InteropServices.ComTypes.FILETIME lastWritten;
            internal UInt32 credentialBlobSize;
            internal IntPtr credentialBlob;
            internal CRED_PERSIST persist;
            internal UInt32 attributeCount;
            internal IntPtr credAttribute;
            [MarshalAs(UnmanagedType.LPWStr)]
            internal string targetAlias;
            [MarshalAs(UnmanagedType.LPWStr)]
            internal string userName;
        }
        
        /// <summary>
        /// Credential
        /// </summary>
        public class Credential
        {
            private SecureString secureString = null;

            /// <summary>
            /// </summary>
            internal Credential(CREDENTIAL cred)
            {
                this.credential = cred;
                unsafe
                {
                    this.secureString = new SecureString((char*)this.credential.credentialBlob.ToPointer(), (int)this.credential.credentialBlobSize / sizeof(char));
                }                
            }

            /// <summary>
            /// </summary>
            public string UserName
            {
                get { return this.credential.userName; }
            }

            /// <summary>
            /// </summary>
            public SecureString Password
            {
                get
                {
                    return this.secureString;
                }
            }

            /// <summary>
            /// </summary>
            internal CREDENTIAL Struct
            {
                get { return this.credential; }
            }

            private CREDENTIAL credential;
        }

        internal class CredentialMarshaler : ICustomMarshaler
        {
            public void CleanUpManagedData(object ManagedObj)
            {
                // Nothing to do since all data can be garbage collected.
            }

            public void CleanUpNativeData(IntPtr pNativeData)
            {
                if (pNativeData == IntPtr.Zero)
                {
                    return;
                }
                CredFree(pNativeData);
            }

            public int GetNativeDataSize()
            {
                return Marshal.SizeOf(typeof(CREDENTIAL));
            }

            public IntPtr MarshalManagedToNative(object obj)
            {
                Credential cred = (Credential)obj;
                if (cred == null)
                {
                    return IntPtr.Zero;
                }

                IntPtr nativeData = Marshal.AllocCoTaskMem(this.GetNativeDataSize());
                Marshal.StructureToPtr(cred.Struct, nativeData, false);

                return nativeData;
            }

            public object MarshalNativeToManaged(IntPtr pNativeData)
            {
                if (pNativeData == IntPtr.Zero)
                {
                    return null;
                }
                CREDENTIAL cred = (CREDENTIAL)Marshal.PtrToStructure(pNativeData, typeof(CREDENTIAL));
                return new Credential(cred);
            }

            public static ICustomMarshaler GetInstance(string cookie)
            {
                return new CredentialMarshaler();
            }
        }    
        

        /// <summary>
        /// ReadCredential
        /// </summary>
        /// <param name="credentialKey"></param>
        /// <returns></returns>
        public static NetworkCredential ReadCredential(string credentialKey)
        {
            Credential credential;
            CredRead(credentialKey, CRED_TYPE.CRED_TYPE_GENERIC, 0, out credential);
            return credential != null ? new NetworkCredential(credential.UserName, credential.Password) : null;
        }
    }
}
"@

######################################################################
# Load credential APIs
######################################################################
$CredManagerType = $null
try
{
    $CredManagerType = [SyncSiteMailbox.CredManager]
}
catch [Exception]
{
}

if($null -eq $CredManagerType)
{
    $compilerParameters = New-Object -TypeName System.CodeDom.Compiler.CompilerParameters
    $compilerParameters.CompilerOptions = "/unsafe"
    [void]$compilerParameters.ReferencedAssemblies.Add("System.dll")
    Add-Type $CredManager -CompilerParameters $compilerParameters
    $CredManagerType = [SyncSiteMailbox.CredManager]
}


######################################################################
# Load tenant credential from generic credential store
######################################################################
$TenantCredential = $null

WriteInfoLog "Load tenant credential is from generic credential store."
try
{
    $credential = $CredManagerType::ReadCredential($TenantCredentialKey)
    if ($null -ne $credential)
    {
        $TenantCredential = New-Object System.Management.Automation.PSCredential ($credential.UserName, $credential.SecurePassword);
    }
}
catch [Exception]
{
    $TenantCredential = $null
    $errorMessage = $_.Exception.Message
    WriteErrorLog "Tenant credential cannot be loaded correctly: $errorMessage."
}

if ($null -eq $TenantCredential)
{
    WriteErrorLog "Tenant credential cannot be loaded please ensure you have configured in credential manager correctly."
}
try
{
    ######################################################################
    # Import site mailbox changes from Exchange online
    ######################################################################
    WriteInfoLog "Import site mailbox changes from Exchange online."
    .\Import-SyncSiteMailbox.ps1 -WorkingFolder $WorkingFolder -TenantCredential $TenantCredential

    ######################################################################
    # Export sync site mailbox changes to On-Premise active directory
    ######################################################################
    WriteInfoLog "Export sync site mailbox changes to local active directory."
    .\Export-SyncSiteMailbox.ps1 -WorkingFolder $WorkingFolder
}
catch [Exception]
{
    $errorMessage = $_.Exception.Message
    WriteErrorLog "Failed to sync site mailbox changes because of: $errorMessage." $false
}
finally
{
    WriteInfoLog "This sync cycle completed."
}

# SIG # Begin signature block
# MIIdpgYJKoZIhvcNAQcCoIIdlzCCHZMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU0PatOPVjwRmxf6y2jD13uxVB
# Z1ygghhkMIIEwzCCA6ugAwIBAgITMwAAAJgEWMt/IwmwngAAAAAAmDANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI3
# WhcNMTcwNjMwMTkyMTI3WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OjdBRkEtRTQxQy1FMTQyMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA1jclqAQB7jVZ
# CvOuH5jFixrRTGFtwMHws1sEZaA3ciobVIdWIejc5fBu3XdwRLfxjsmyou3JeTaa
# 8lqA929q2AyZ9A3ZBfxf8VqTxbu06wBj4o4g5YCsz0C/81N2ESsQZbjDxbW5sKzD
# hhT0nTzr82aepe1drjT5dvyU/AvEkCzaEDU0dZTq2Bm6NIif11GzA+OkD0bdZG+u
# 4EJRylQ4ijStGgXUpAapb0y2RtlAYvZSpLYzeFFcA/yRXacCnoD++h9r66he/Scv
# Gfd/J/5hPRCtgsbNr3vFJzBWgV9zVqmWOvZBPGpLhCLglTh0stPa/ZxZjTS/nKJL
# a7MZId131QIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFPPCI5/SvSWNvaj1nBvoSHO7
# 6ZPBMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAD+xPVIhFl30XEe39rlgUqCCr2fXR9o0aL0Oioap6LAUMXLK
# 4B+/L2c+BgV32joU6vMChTaA+7XEw7pXCRN+uD8ul4ifHrdAOEEqOTBD7N5203u2
# LN667/WY71purP2ezNB1y+YAgjawEt6VjjQcSGZ9bTPRtS2JPS5BS868paym355u
# 16HMxwmhlv1klX6nfVOs1DYK5cZUrPAblCZEWzGab8j9d2ZIGLQmTEmStdslOq79
# vujEI0nqDnJBusUGi28Kh3Hz1QIHM5UZg/F5sWgt0EobFGHmk4KH2vreGZArtCIB
# amDc5cIJ48na9GfA2jqJLWsbvNcwC486g5cauwkwggYHMIID76ADAgECAgphFmg0
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
# bWrJUnMTDXpQzTGCBKwwggSoAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCBwDAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUpB5KS5DasRFSgHvaFpIGFF8rN1wwYAYKKwYB
# BAGCNwIBDDFSMFCgKIAmAFMAeQBuAGMAUwBpAHQAZQBNAGEAaQBsAGIAbwB4AC4A
# cABzADGhJIAiaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdlIDANBgkq
# hkiG9w0BAQEFAASCAQAC8PFbLlp7F4ZaLAieBbajbjUCKiTSUNNvUvD3wRUvzRpD
# Fg0/nQe9R3fV8RPQlm9t++qvBcyKkIBbvUgHgw+bkMvC3oI8iqEJp4i+/BUsQFYt
# 28jZHS0DUVPVPIKhgfuzXyXYlL8CPIAe4e7IKv6PLy/glIeHMMqJckzi43xIsTwM
# D2dxsJ/c8R4lHRcuSGWmpKRf3N72mro7sHPBfOdzQl46Nkr6Br1+qwFQfXx/siid
# udZ8rZXzUfi9xY8uS1UMl2tfucPND+h8sohTv95KCHK2NnW0E8+Gebc/td+QA7Pu
# zwHKBxeHvy/44bsrcEKKvrQUaGi4g05LMCLMzY7PoYICKDCCAiQGCSqGSIb3DQEJ
# BjGCAhUwggIRAgEBMIGOMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5n
# dG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
# YXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQQITMwAAAJgE
# WMt/IwmwngAAAAAAmDAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3
# DQEHATAcBgkqhkiG9w0BCQUxDxcNMTYwOTAzMTg0NDMxWjAjBgkqhkiG9w0BCQQx
# FgQUVBahHws0FhQd5xSNABpUz9ET4XMwDQYJKoZIhvcNAQEFBQAEggEAZqVBgARK
# QiGRWNfOrrWAPdOTxRFFWLcoJlNnlZ28F28CYG0w9EM6JGGRtGyfmKTqGoTEhcnz
# MiV7rGf2yB5iLH5H9iCLSozKUcTEpbh21N4+cE55oXw1EkPzADa85BnSkfNCRVgm
# b2OjEVPJxMOhKiK2CXLkd2k0Cx2y0gQMinchKLuOwfqXNRE5nJ5+v0bftdo6+KwI
# xOX1c7tYb+OSMv+AJPEhEwmdUO9NSZW88bzQEX6CSv2sBdo7uA3y6U5CU/ox6x1a
# JEi2iurOqhTz1gAvADLFfboag3R6epT6GO/WHePJW/woLh3fMyaqrMO2d/GBfFcR
# hiQG1NAwf8pTdQ==
# SIG # End signature block

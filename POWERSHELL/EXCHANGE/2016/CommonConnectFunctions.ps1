# Copyright (c) 2010 Microsoft Corporation. All rights reserved. 
# This script contains all the common Powershell functions required to connect to remote PS, both in Enterprise and Datacenter

#load hashtable of localized string
Import-LocalizedData -BindingVariable CommonConnectFunctions_LocalizedStrings -FileName CommonConnectFunctions.strings.psd1

## PROMPT ####################################################################

## PowerShell can support very rich prompts, this simple one prints the current
## working directory and updates the console window title to show the machine 
## name and directory.  

function prompt 
{ 
	$cwd = (get-location).Path
	$host.UI.RawUI.WindowTitle = ($CommonConnectFunctions_LocalizedStrings.res_0004 -f $global:connectedFqdn)
	$host.UI.Write("Yellow", $host.UI.RawUI.BackGroundColor, "[PS]")
	" $cwd>" 
}

## generates a tip of the day 

function get-tip
{
    param($local:number=$null) 

    if( ($global:exrandom -eq $null) -or ($exrandom -isnot [System.Random]))
    {
        $global:exrandom = new-object System.Random
    }

    $exchculture = [System.Threading.Thread]::CurrentThread.CurrentUICulture
    $foundculture = $false
    while($exchculture -ne [Globalization.CultureInfo]::InvariantCulture -and !$foundculture)
    {
    	if ( test-path "$($global:exbin)\$($exchculture.Name)\extips.xml" )
    	{
    		$foundculture = $true
    	}
    	else
    	{
    		$exchculture = $exchculture.Parent
    	}
    }

	if($foundculture -eq $true)
    {
	$exchculture = $exchculture.Name 
    }
    else 
    {
	$exchculture = 'en'
    } 

    if (test-path "$($global:exbin)\$exchculture\extips.xml")
    {

        $local:tips = [xml](get-content $global:exbin\$exchculture\extips.xml)
        if($local:number -eq $null)
        {
            $local:temp = $global:exrandom.Next( 0, $local:tips.topic.developerConceptualDocument.introduction.table.row.Count )
        }
        else
        {
            $local:temp = $local:number
        }
        $local:nav = $tips.topic.developerConceptualDocument.introduction.table.row[$local:temp].entry.CreateNavigator()
        write-host -fore Yellow ( $CommonConnectFunctions_LocalizedStrings.res_0000 -f $local:temp )
        [void] $nav.MoveToFirstChild()
        do
        {
             write-host $nav.Value
        }
        while( $nav.MoveToNext() )
        ""
    }
    else
    {
        ($CommonConnectFunctions_LocalizedStrings.res_0005 -f "$($global:exbin)\$exchculture\extips.xml")
    }

    trap
    {
	continue
    }
}

function ImportPSSession ([bool]$ClearCache, [switch]$AllowClobber)
{
	if (!($global:remoteSession -eq $null))
	{
	    # do not display all the commands - turn off verbose output
	    set-variable VerbosePreference -value SilentlyContinue

	    $serverName = $global:remoteSession.ComputerName
	    $modulePath = "$env:APPDATA\Microsoft\Exchange\RemotePowerShell\$serverName"
	    $remotePSSettinsPath = "HKCU:Software\Microsoft\ExchangeServer\v15\RemotePowerShell\$serverName"

	    if ($ClearCache)
	    {
	        Write-Host -fore Yellow $CommonConnectFunctions_LocalizedStrings.res_0001
	        ClearRegistryEntryAndModule $remotePSSettinsPath $modulePath
	    }
	    
	    try
	    {
	        if ($AllowClobber)
	        {
	            ExportPSSessionAndImportModule $remotePSSettinsPath $modulePath -AllowClobber
	        }
	        else
	        {
	            ExportPSSessionAndImportModule $remotePSSettinsPath $modulePath
	        }
	    }
	    catch
	    {
	        Write-Warning $CommonConnectFunctions_LocalizedStrings.res_0002
	    }
	    $global:importResults = Get-Module
	    if ($global:importResults -eq $null)
	    {
	        if ($AllowClobber)
	        {
	                $global:importResults=Import-PSSession $global:remoteSession -WarningAction SilentlyContinue -DisableNameChecking -AllowClobber
	        }
	        else
	        {
	                $global:importResults=Import-PSSession $global:remoteSession -WarningAction SilentlyContinue -DisableNameChecking
	        }
	    }
	    set-variable VerbosePreference -value Continue

	    Write-Verbose ($CommonConnectFunctions_LocalizedStrings.res_0003 -f $global:connectedFqdn)
	}
}


function ExportPSSessionAndImportModule ($remotePSSettinsPath, $modulePath, [switch]$AllowClobber)
{
	$hashValue = $global:remoteSession.ApplicationPrivateData.ImplicitRemoting.Hash
	$CurrentUserRemotePSSettings = Get-ItemProperty -path $remotePSSettinsPath -ErrorAction SilentlyContinue

	# PS3.0, Get-ItemProperty will return DWORD data as UInt32, instead of Int32 in PS2.0.
	# If $hashValue is negative, (CurrentUserRemotePSSettings.Hash -ne $hashValue) will always be $true
	# We use bitwise xor operation to work around
	if (($CurrentUserRemotePSSettings -eq $null) `
		-or ($CurrentUserRemotePSSettings.Hash -eq $null) `
		-or (-not ($CurrentUserRemotePSSettings.ModulePath)) `
		-or (($hashValue -bxor $CurrentUserRemotePSSettings.Hash) -ne 0))
	{
		# Redo Everything, when:
		# 1. No registry entry found, or
		# 2. Registry entry exists, but hash value or ModulePath is empty (which is very unlikely) or
		# 3. Hash value of the saved module didn't match with the hash value of the current session
		if ($AllowClobber)
		{
			CreateRegistryEntryAndImportModule $remotePSSettinsPath $hashValue $modulePath -AllowClobber
		}
		else
		{
			CreateRegistryEntryAndImportModule $remotePSSettinsPath $hashValue $modulePath
		}
	}
	else
	{
		$modulePath = $CurrentUserRemotePSSettings.ModulePath
		$module =  Get-ChildItem $modulePath -ErrorAction SilentlyContinue | where-object{$_.Extension -eq ".psm1" -or $_.Extension -eq ".psd1" -or $_.Extension -eq ".ps1xml"}
		if (($module -eq $null) -or ($module.Count -lt 3))
		{
			# If the module folder exists, but any module file is missing, then we should just export the session and import module
			if ($AllowClobber)
			{
				Export-PSSession -Session $global:remoteSession -OutputModule $modulePath -force -AllowClobber | out-null
			}
			else
			{
				Export-PSSession -Session $global:remoteSession -OutputModule $modulePath -force | out-null
			}
			Import-Module -Name $modulePath -ArgumentList $global:remoteSession -DisableNameChecking
		}
		else
		{
			if  ((New-TimeSpan -Start $module[0].LastWriteTime -End (get-date)).TotalHours -gt 72)
			{
				# If the module is expired, then we should redo everything
				if ($AllowClobber)
				{
					CreateRegistryEntryAndImportModule $remotePSSettinsPath $hashValue $modulePath -AllowClobber
				}
				else
				{
					CreateRegistryEntryAndImportModule $remotePSSettinsPath $hashValue $modulePath
				}
			}
			else
			{
				Import-Module -Name $modulePath -ArgumentList $global:remoteSession -DisableNameChecking
			}
		}
	}
}

function CreateRegistryEntryAndImportModule ($remotePSSettinsPath, $hashValue, $modulePath, [switch]$AllowClobber )
{
	ClearRegistryEntryAndModule $remotePSSettinsPath $modulePath
	new-item $remotePSSettinsPath -force | out-null
	New-ItemProperty -Path $remotePSSettinsPath -Name Hash -Value $hashValue -PropertyType DWord -force | out-null
	New-ItemProperty -Path $remotePSSettinsPath -Name ModulePath -value $modulePath -PropertyType ExpandString -force | out-null
	if ($AllowClobber)
	{
		Export-PSSession -Session $global:remoteSession -OutputModule $modulePath -force -AllowClobber | out-null
	}
	else
	{
		Export-PSSession -Session $global:remoteSession -OutputModule $modulePath -force | out-null
	}
	Import-Module -Name $modulePath -ArgumentList $global:remoteSession -DisableNameChecking
}

function ClearRegistryEntryAndModule ($remotePSSettinsPath, $modulePath)
{
	clear-item $remotePSSettinsPath -force -ErrorAction SilentlyContinue
	Get-ChildItem $modulePath -ErrorAction SilentlyContinue | Remove-Item -force -recurse -ErrorAction SilentlyContinue
}

##########################################################################################################################
#
#  Contains cmdlets which create a PSCredential object by consuming user
#  input in the form of a liveID (like user@hotmail.com) and password.
#
#  This module depends on SSPIPromptForCredentials win32 API for generating Nego2/LiveSSP token
#  PSHostUserInterface method PromptForCredential will not work for Nego2
#
  
######### PInvoke code which retrieves PSCredential using SSPIPromptForCredentials #########
$getLiveIDCredCode = @"
using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace Microsoft.PowerShell.Commands
{
    [StructLayout(LayoutKind.Sequential)]
    public struct CREDUI_INFO
    {
        public int cbSize;
        public IntPtr hwndParent;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string pszMessageText;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string pszCaptionText;
        public IntPtr hbmBanner;
    }

    public class LiveIDCredential
    {
        private const string NEGOSSP_NAME_W = "Negotiate";
        private const string NEGOSSP_2_W = "Nego2";

        public static PSCredential GetLiveIDCredential(string captionText,
            string messageText,
            string connectionUri)
        {
            // This will throw if the connectionUri is not a valid uri string.
            Uri connUri = new UriBuilder(connectionUri).Uri;
            WSManConnectionInfo cInfo = new WSManConnectionInfo(
                connUri, (string)null, (System.Management.Automation.PSCredential)null);
            cInfo.AuthenticationMechanism = AuthenticationMechanism.Negotiate;
            Runspace remoteRunspace = RunspaceFactory.CreateRunspace(cInfo);
            try
            {
                remoteRunspace.Open();
            }
            catch (Exception)
            {
            }

            CREDUI_INFO credUiInfo = new CREDUI_INFO();
            credUiInfo.pszCaptionText = captionText;
            credUiInfo.pszMessageText = messageText;
            credUiInfo.hwndParent = IntPtr.Zero;
            credUiInfo.hbmBanner = IntPtr.Zero;
            credUiInfo.cbSize = Marshal.SizeOf(credUiInfo);

            IntPtr ppAuthIdentity = IntPtr.Zero;
            bool fSave = false;

            string targetName = connUri.Host;
            int result = SspiPromptForCredentials(targetName,
                ref credUiInfo,
                0,
                NEGOSSP_NAME_W,
                IntPtr.Zero,
                ref ppAuthIdentity,
                ref fSave,
                1);

            if (0 != result)
            {
                throw new System.InvalidOperationException(
                    string.Format(CultureInfo.InvariantCulture,
                    "SspiPromptForCredentials failed with error {0}", result));
            }


            StringBuilder pszUserName = new StringBuilder();
            StringBuilder pszDomainName = new StringBuilder();
            StringBuilder pszPassword = new StringBuilder();

            result = SspiEncodeAuthIdentityAsStrings(ppAuthIdentity,
                ref pszUserName,
                ref pszDomainName,
                ref pszPassword);

            if (0 != result)
            {
                throw new System.InvalidOperationException(
                    string.Format(CultureInfo.InvariantCulture,
                    "SspiEncodeAuthIdentityAsStrings failed with error {0}", result));
            }

            System.Security.SecureString pwd = new System.Security.SecureString();
            for (int i = 0; i < pszPassword.Length; i++)
            {
                pwd.AppendChar(pszPassword[i]);
            }

            string userName = null;
            if (pszDomainName != null)
            {
                userName = string.Format(CultureInfo.InvariantCulture,
                    "{0}\\{1}", pszDomainName.ToString(), pszUserName.ToString());
            }
            else
            {
                userName = pszUserName.ToString();
            }

            PSCredential credToReturn = new PSCredential(userName, pwd);
            return credToReturn;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pszTargetName">
        /// A pointer to a null-terminated string that indicates the 
        /// service principal name (SPN) or the security context of the 
        /// destination server.
        /// </param>
        /// <param name="pUiInfo">
        /// A pointer to a CREDUI_INFO structure that contains information 
        /// for customizing the appearance of the dialog box.
        /// </param>
        /// <param name="dwAuthError">
        /// A Windows error code, defined in WinError.h, that is displayed in the dialog box.
        /// If credentials previously collected were not valid, the caller uses this parameter
        /// to pass the error message from the API that collected the credentials (for example, Winlogon)
        /// to this function. The corresponding error message is formatted and displayed in the dialog box.
        /// Set the value of this parameter to zero to display no error message.
        /// </param>
        /// <param name="pszPackage">
        /// contains the name of authentication package
        /// </param>
        /// <param name="pInputAuthIdentity"></param>
        /// <param name="ppAuthIdentity"></param>
        /// <param name="pfSave">
        /// A pointer to a Boolean value that, on input, specifies whether the Save check box is 
        /// selected in the dialog box that this function displays. On output, the value of this 
        /// parameter specifies whether the Save check box was selected when the user clicks the
        /// Submit button in the dialog box. Set this parameter to NULL to ignore the Save check box.
        /// 
        /// This parameter is ignored if the CREDUIWIN_CHECKBOX flag is not set in the dwFlags parameter.
        /// </param>
        /// <param name="dwFlags">
        /// 1. SSPIPFC_CHECKBOX  0x1 -- If applications need to show the 
        /// checkbox to save credentials, the SSPIPFC_CHECKBOX should be set.
        /// 2. SSPIPRFC_SAVE_CRED_BY_CALLER 0x2 
        /// By default, the credential provier() saves the credentials to Credman/KeyRing 
        /// with CRED_PERSIST_ENTERPRISE persistence(http://msdn2.microsoft.com/en-us/library/aa374788(VS.85).aspx),
        /// but the caller can overrides this behavior by supplying SSPIPFC_SAVE_CRED_BY_CALLER.
        /// When this flag is set, the credential provider does not save the credentials to credman/keyring. 
        /// </param>
        /// <returns></returns>
        [DllImport("credui", SetLastError = false, CharSet = CharSet.Unicode)]
        private static extern int SspiPromptForCredentials(
            string pszTargetName,
            ref CREDUI_INFO pUiInfo,
            int dwAuthError,
            string pszPackage,
            IntPtr pInputAuthIdentity,
            ref IntPtr ppAuthIdentity,
            ref bool pfSave,
            int dwFlags);

        [DllImport("sspicli", SetLastError = false, CharSet = CharSet.Unicode)]
        private static extern int SspiEncodeAuthIdentityAsStrings(
            IntPtr pAuthIdentity,
            ref StringBuilder pszUserName,
            ref StringBuilder pszDomainName,
            ref StringBuilder pszPackedCredentialsString);
    }
}
"@

######### END: PInvoke #####################################################################

function EnsureLiveIDCredentialTypeIsLoaded
{
    if (!('Microsoft.PowerShell.Commands.GetLiveIDCredential' -as [type]))
    {
        # LiveIDCredential Type is not loaded. So load it using Add-Type
        try
        {
			add-type -TypeDefinition $getLiveIDCredCode
		}
		catch
		{
			# We can ignore any TYPE_ALREADY_EXISTS exception
			if ($error[0].FullyQualifiedErrorId -notmatch "TYPE_ALREADY_EXISTS")
			{
				throw
			}
		}
    }
}

function Get-ExLiveIDCredential
{
   [CmdletBinding()]
   param(
     [Parameter(Position=0)]
     [string]$connectionUri
   )
   
   begin
   {
        EnsureLiveIDCredentialTypeIsLoaded
        $captionText = "LiveID Credential";
        $messageText = "Supply credentials for connecting to {0}" -f $connectionUri
        [Microsoft.PowerShell.Commands.LiveIDCredential]::GetLiveIDCredential($captionText,$messageText,$connectionUri)           
   }
}
##########################################################################################################################


# SIG # Begin signature block
# MIIdtAYJKoZIhvcNAQcCoIIdpTCCHaECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUZtWie4cb/LSCKBSq55hefD06
# pE+gghhkMIIEwzCCA6ugAwIBAgITMwAAAJgEWMt/IwmwngAAAAAAmDANBgkqhkiG
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
# bWrJUnMTDXpQzTGCBLowggS2AgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCBzjAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUTxeh04Euqz83aMeDWQbt9rqUTsgwbgYKKwYB
# BAGCNwIBDDFgMF6gNoA0AEMAbwBtAG0AbwBuAEMAbwBuAG4AZQBjAHQARgB1AG4A
# YwB0AGkAbwBuAHMALgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20v
# ZXhjaGFuZ2UgMA0GCSqGSIb3DQEBAQUABIIBAF4K4U6Jok2XvRogg8wVzkz5BeCE
# wez6jiw2Ny5eeeUROKjm1A1rWWqmNG9XvFP5ED3bgZOnBTFPEDLgXH+gmHrcyF9c
# D6fpQW0iglYyvDX3ky67b6wkahDz6P9zcr1E1Ac9hoFIU4kltjAqCiVTZyPbbN6m
# /sLFXhHo+gaFWAkiKMwxgSG7wkAA593VNNFFZxCCatlQ5+5pAfJSsfO4VjL56Lkz
# LS7kViAHV4eXzsHPHQtTD/XGg3xqbVylCaQctpuus9l5ikco9XGP+JkSGhXDce2p
# nWByf8AOfbj454v+Dw66MSZvymHRNmE6egBz9fBqk9RtGkf2qqyYEyWrAkihggIo
# MIICJAYJKoZIhvcNAQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3Rh
# bXAgUENBAhMzAAAAmARYy38jCbCeAAAAAACYMAkGBSsOAwIaBQCgXTAYBgkqhkiG
# 9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNjA5MDMxODQ0MzVa
# MCMGCSqGSIb3DQEJBDEWBBRCTH2MwEhZymI9KFWZBGWTSY8UCzANBgkqhkiG9w0B
# AQUFAASCAQC8fRGWiI5Fqjzp35Z4Tso6tEVjVyadSjYrzwf7gJg/BudYNoDT+Mtv
# 7Lf201MqObaLyjRk2CLdAXfDcwi1b+RuMQx2wtKl1rACzJ+FTgeMz48PCSQGepvT
# wacuMY1GQt066x9qdcA6VyITicePj8HekSqGVDswfqgPyCy3Qtn4caV3/e3Wa8rJ
# ULchCNZ6eMecseMFDqtmBCNK0UOmmzfPRoLbkOZCfcxp6ROEFLFo0ZKWEWJEfGUO
# NK357Dvj6XajmkfjNHMEA6OacewKbJ75S+ruXKFYEm6xSMIx5DiW1IK2/412Lzhv
# 6exVFaFKWtfnikn7hkDBbqscYh0ey3zo
# SIG # End signature block

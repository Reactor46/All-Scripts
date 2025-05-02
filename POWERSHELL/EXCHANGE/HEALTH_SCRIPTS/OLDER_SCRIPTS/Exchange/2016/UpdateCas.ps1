###############################################################################
# This script handles OWA/ECP updates.
###############################################################################

trap {
	Log ("Error updating OWA/ECP: " + $_)
	Exit
}

$WarningPreference = 'SilentlyContinue'
$ConfirmPreference = 'None'

$script:logDir = "$env:SYSTEMDRIVE\ExchangeSetupLogs"

# Log( $entry )
#	Append a string to a well known text file with a time stamp
# Params:
#	Args[0] - Entry to write to log
# Returns:
#	void
function Log
{
	$entry = $Args[0]

	$line = "[{0}] {1}" -F $(get-date).ToString("HH:mm:ss"), $entry
	write-output($line)
	add-content -Path "$logDir\UpdateCas.log" -Value $line
}

# If log file folder doesn't exist, create it
if (!(Test-Path $logDir)){
	New-Item $logDir -type directory	
}

# Load the Exchange PS snap-in
add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010

Log "***********************************************"
Log ("* UpdateCas.ps1: {0}" -F $(get-date))

# If Mailbox isn't installed on this server, exit without doing anything
if ((Get-ExchangeServer $([Environment]::MachineName)).ServerRole -notmatch "Mailbox") {
		Log "Warning: Mailbox role is not installed on server $([Environment]::MachineName)"
}
Log "Updating OWA/ECP on server $([Environment]::MachineName)"

# get the path to \owa on the filesystem
Log "Finding ClientAccess role install path on the filesystem"
$caspath = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath + "ClientAccess\"

# GetVersionFromDll
#	Gets the version information from a specified dll
#	appName - the friendly name of the web application
#	webApp  - the alias of the web application, should be the folder name of the physical path
#	dllName  - the name of the assembly that indicate the version of the web application (version folder)
# Returns:
#	String value of the version information on the specified dll.
function GetVersionFromDll($appName, $webApp, $dllName)
{
	$apppath = $caspath + $webApp + "\"
	$dllpath = $apppath + "bin\" + $dllName
	# figure out which version of web application (OWA/ECP) we are moving to
	if (! (test-path ($dllpath))) {
		Log "Could not find '${dllpath}'.  Aborting."
		return $null
	}
	$version = ([diagnostics.fileversioninfo]::getversioninfo($dllpath)).fileversion -replace '0*(\d+)','$1'
	return $version
}

# UpdateWebApp
#	Update a web application in current CAS server.
# Params:
#	appName - the friendly name of the web application
#	webApp  - the alias of the web application, should be the folder name of the physical path
#	version - The version to use. 
#	sourceFolder - The name of the webApp folder to copy files from
#			e.g.: "Current2" for OWA2, "Current" is used for OWA Basic, ECP.
#	destinationFolder - The optional name of the folder to copy to
# Returns:
#	void
# For example, UpdateWebApp "OWA" "owa" "15.0.815" "Current" "prem"
function UpdateWebApp($appName, $webApp, $version, $sourceFolder, $destinationFolder)
{
	if ($version -eq $null) {
		Log "Could not determine version. Aborting."
		return
	}

	$apppath = $caspath + $webApp + "\"
	Log "Updating ${appName} to version $version"

	# filesystem path to the new version directory
	if ($destinationFolder -eq $null){
		$versionpath = $apppath + $version
	}
	else {
		$versionpath = $apppath + $destinationFolder + "\" + $version
	}

	Log "Copying files from '${apppath}${sourceFolder}' to '$versionpath'"
    New-Item $versionpath -Type Directory -ErrorAction SilentlyContinue
	copy-item -recurse -force ($apppath + $sourceFolder + "\*") $versionpath
	
	Log "Update ${appName} done."
}

# Upgrade from CU5 to CU6 leaves some files missing from unversioned OWA folder
# that is updated & replicated during MSP updates
function FixUnversionedFolderAfterUpgrade
{
	try
	{
		$setupRegistry = Get-Item -Path HKLM:\Software\Microsoft\ExchangeServer\v15\Setup\ -ea SilentlyContinue
		if (!$setupRegistry) { Log "FixUnversionedFolderAfterUpgrade: No setupRegistry"; return }

		# C:\Program Files\Microsoft\Exchange Server\V15\ClientAccess
		$installPath = $setupRegistry.GetValue('MsiInstallPath')
		if (!$installPath) { Log "FixUnversionedFolderAfterUpgrade: No installPath"; return }

		# 15.0.995.32
		$installedFork = (@('MsiProductMajor','MsiProductMinor','MsiBuildMajor') | %{ $setupRegistry.GetValue($_) }) -join '.'
		$srcVersions = @((get-item "$installPath\ClientAccess\Owa\prem\$($installedFork).*").Name | Sort { [System.Version] $_ })
		if (!$srcVersions) { Log "FixUnversionedFolderAfterUpgrade: No srcVersions $($installedFork).*"; return }
		Log "FixUnversionedFolderAfterUpgrade: Found source versions: $srcVersions"

		$srcRoot = (Get-Item "$installPath\ClientAccess\Owa\prem\$($srcVersions[0])").FullName
		$destRoot = (Get-Item "$installPath\ClientAccess\Owa\Current2\version").FullName
		Log "FixUnversionedFolderAfterUpgrade: Recovering files from '$srcRoot' to '$destRoot' where necessary"
		foreach ($srcPath in (Get-ChildItem -File -Recurse $srcRoot).FullName)
		{
			$subPath = $srcPath.Substring($srcRoot.Length+1)
			$destPath = "$destRoot\$subPath"
			if (!(Get-Item $destPath -ea SilentlyContinue))
			{
				Log "Copy-Item '$srcPath' '$destPath'"
				$destParent = Split-Path $destPath
				if ($destParent -and !(Test-Path $destParent))
				{
					$null = New-Item $destParent -type Directory
				}
				Copy-Item -Force $srcPath $destPath
			}
		}
		Log "FixUnversionedFolderAfterUpgrade success"
	}
	catch
	{
		Log "FixUnversionedFolderAfterUpgrade failed: $_.Exception.Message"
	}
}
FixUnversionedFolderAfterUpgrade

# Add an attribute to a given XML Element
# <param name="xmlDocument">Document where the attribute will be added</param>
# <param name="xmlElement">Element where the attribute will be added</param>
# <param name="attributeName">Name of the attribute</param>
# <param name="attributeValue">Value of the attribute</param>
function AddXmlAttribute
{
    param ([System.Xml.XmlDocument] $xmlDocument, [System.Xml.XmlElement] $xmlElement, [string] $attributeName, [string] $attributeValue);


    $attribute = $xmlDocument.CreateAttribute($attributeName);
    $attribute.set_Value($attributeValue) | Out-Null
    $xmlElement.SetAttributeNode($attribute) | Out-Null
}

# Add an assembly to the owa web.config.
# <param name="assemblyName">The assembly name to add.</param>
# <param name="version">The version of the assembly.</param>
function AddOwaWebConfigAssembly($assemblyName, $version)
{
    $assembliesNodeName = "configuration//system.web/compilation/assemblies"
	$owaWebConfigFolder = $caspath  + "owa\"
    $owaWebConfigPath = $owaWebConfigFolder + "web.config"
		
	$owaWebConfigPathCheck = (Get-Item $owaWebConfigPath).FullName
	if (!$owaWebConfigPathCheck) { Log "no OWA web.config is found"; return }

    $xmlDocument = New-Object System.Xml.XmlDocument;
    $xmlDocument.Load($owaWebConfigPath);
    $xmlNode = $xmlDocument.SelectSingleNode($assembliesNodeName);

    if ($xmlNode -eq $null) { Log "$assembliesNodeName is not found in web.config"; return }

    $xmlKeyNode = $assembliesNodeName + "/add[starts-with(@assembly, '" + $assemblyName + ",')]";
    $xmlKeyNodeValue = $xmlDocument.SelectSingleNode($xmlKeyNode);

    if ($xmlKeyNodeValue -eq $null)
    {
       [System.Xml.XmlNode] $xmlKeyNodeValue = $xmlDocument.CreateNode([System.Xml.XmlNodeType]::Element, "add", $null);
            
        AddXmlAttribute $xmlDocument $xmlKeyNodeValue "assembly" "$assemblyName,Version=$version,Culture=neutral,publicKeyToken=31bf3856ad364e35";

        $xmlNode.PrependChild($xmlKeyNodeValue) | Out-Null
        $xmlDocument.Save($owaWebConfigPath) | Out-Null
    }
}

AddOwaWebConfigAssembly "Microsoft.Exchange.VariantConfiguration.Core" "15.0.0.0"

# Update OWA
$owaBasicVersion = (get-itemproperty -Path HKLM:\Software\Microsoft\ExchangeServer\v15\Setup\ -Name "OwaBasicVersion" -ea SilentlyContinue).OwaBasicVersion
$owaVersion = (get-itemproperty -Path HKLM:\Software\Microsoft\ExchangeServer\v15\Setup\ -Name "OwaVersion" -ea SilentlyContinue).OwaVersion

UpdateWebApp "OWA" "owa" $owaBasicVersion "Current" $null
UpdateWebApp "OWA" "owa" $owaVersion "Current2\version" "prem"

# Update ECP
# Anonymous access has been enabled on ECP root folder by default, so it isn't necessary to enable anonymous access on the version folder explicitly
$ecpVersion = GetVersionFromDll "ECP" "ecp" "Microsoft.Exchange.Management.ControlPanel.dll"
UpdateWebApp "ECP" "ecp" $ecpVersion "Current" $null

# Remove the Exchange PS snap-in
remove-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010

# SIG # Begin signature block
# MIIdmgYJKoZIhvcNAQcCoIIdizCCHYcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUHi3/gK2a4ZSxfYdrnLWnpJUZ
# hqCgghhkMIIEwzCCA6ugAwIBAgITMwAAAJ1CaO4xHNdWvQAAAAAAnTANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTMw
# WhcNMTcwNjMwMTkyMTMwWjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OjE0OEMtQzRCOS0yMDY2MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAy8PvNqh/8yl1
# MrZGvO1190vNqP7QS1rpo+Hg9+f2VOf/LWTsQoG0FDOwsQKDBCyrNu5TVc4+A4Zu
# vqN+7up2ZIr3FtVQsAf1K6TJSBp2JWunjswVBu47UAfP49PDIBLoDt1Y4aXzI+9N
# JbiaTwXjos6zYDKQ+v63NO6YEyfHfOpebr79gqbNghPv1hi9thBtvHMbXwkUZRmk
# ravqvD8DKiFGmBMOg/IuN8G/MPEhdImnlkYFBdnW4P0K9RFzvrABWmH3w2GEunax
# cOAmob9xbZZR8VftrfYCNkfHTFYGnaNNgRqV1rEFt866re8uexyNjOVfmR9+JBKU
# FbA0ELMPlQIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFGTqT/M8KvKECWB0BhVGDK52
# +fM6MB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAD9dHEh+Ry/aDJ1YARzBsTGeptnRBO73F/P7wF8dC7nTPNFU
# qtZhOyakS8NA/Zww74n4gvm1AWfHGjN1Ao8NiL3J6wFmmON/PEUdXA2zWFYhgeRe
# CPmATbwNN043ecHiGjWO+SeMYpvl1G4ma0NIUJau9DmTkfaMvNMK+/rNljr3MR8b
# xsSOZxx2iUiatN0ceMmIP5gS9vUpDxTZkxVsMfA5n63j18TOd4MJz+G0I62yqIvt
# Yy7GTx38SF56454wqMngiYcqM2Bjv6xu1GyHTUH7v/l21JBceIt03gmsIhlLNo8z
# Ii26X6D1sGCBEZV1YUyQC9IV2H625rVUyFZk8f4wggYHMIID76ADAgECAgphFmg0
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
# bWrJUnMTDXpQzTGCBKAwggScAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCBtDAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUFUS69o0NVqoi3Ae28rsheHT1LVQwVAYKKwYB
# BAGCNwIBDDFGMESgHIAaAFUAcABkAGEAdABlAEMAYQBzAC4AcABzADGhJIAiaHR0
# cDovL3d3dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASC
# AQAgP1BT1mXK9qypA+6kbNXDz9OHEyLEiN5x7HsNNR690MFiUKGSH8swSqTfgJOt
# 45NC4pdiP7VC859YmJQzVWuE43jvgYC2LyBSwE1aQAVxmfg7BX+ydBDftmbl2Kgk
# mwIVO7GxVf/piJy1y1ackDHrq+EdCqfHOsrNoZd9PFZjdlN4teGugFdEKIVRnIuR
# g8Bh0aqtgrPQt8dqaMKAKV2yfNY+QNxOkytZ8xStgL0lusdXypmRRJcxRyVEXF3Z
# xOwsnw9LOvdxHRgIfpTD1luVwGyIwkiHRTqgy/RZDnCxpeIH3kPD7etR+EtBoMB6
# VizI4WCTxoAP5OiO1zvswuojoYICKDCCAiQGCSqGSIb3DQEJBjGCAhUwggIRAgEB
# MIGOMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNV
# BAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQQITMwAAAJ1CaO4xHNdWvQAAAAAA
# nTAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMTYwOTAzMTg0NTAxWjAjBgkqhkiG9w0BCQQxFgQUX5CVW5qfClrj
# 950+hGDd1jC4QAAwDQYJKoZIhvcNAQEFBQAEggEAUbAvhNBNRgvdc5k9CBBbFw/W
# z/JIvdN0v2nl+IApKrPuNJ2Zd4rI90we20ahKo6g5/E/i7G1OqrmC8gBfrZhi8/F
# p6f/7ffqFvt06gfnH7Rp9gsxBSJdUABbVwOrp6ps8EEzxvtBViqAKtpLxvY5lxP8
# 4Wlx0ftFh8Z77RypJjn31iXO9ekwbg0Rx1XJq0SfdV86sbA9EwvI9TmX7S9BlqJZ
# qAe8tk0UL/wk+QjMgTao/3z37kr2dBAO/7PeJbNNNBRJs6e3HtV9/hUAcAHRJU/g
# rD/GoF2zMWl6h94cnGVJS/aeKIU8ABOr9+r/VQ/dhHaYLUmilB+C3Us2fQP4jw==
# SIG # End signature block

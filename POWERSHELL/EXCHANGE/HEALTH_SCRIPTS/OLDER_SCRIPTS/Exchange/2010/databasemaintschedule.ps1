# DatabaseMaintSchedule.ps1
#
# Copyright (c) Microsoft Corporation. All rights reserved.
#
# Random generate maintenance and quota notificaiton schedule time based on the 
# specified criteria, current it suppose the following criterias specified by parameters
#	Start Time
#	Maint window length
#	Occurrences
# NOTE: We have hardcoded the following factors
#	Maintenance schedule length: 3 hours
#	Quota notification schedule: 15 mins
#

#==============================================================================
# Parameters section
#==============================================================================
param
(
	# Parameter -DagOrServerName
	#	A qualified dag or mailbox server name like "MBX001", "DAG001"
	[string]$DagOrServerName=$(throw "Please specify dag or mailbox server name by parameter -DagOrServerName <name>"),
	
	# Parameter -StartTime
	#	The start time for the maintenance window, with format "HH:mm tt", qualifed examples like "10:00 PM", "5:00 AM"
	#	Default value is "10:00 PM"
	[string]$StartTime="10:00 PM",
	
	# Parameter -TimeWindowLength
	#	The length of the maintenance window, in hours
	#	Default value is 7 hours
	[int]$TimeWindowLength=7,
	
	# Parameter -OccurrencePerWeek
	#	The number of times the maintenance wil run
	#	Default value is 3 times
	[int]$OccurrencePerWeek=3
)

#==============================================================================
# Location function section
#==============================================================================
function local:Generate-MaintenanceSchedule($windowLengthInMinutes)
{
	PROCESS
	{
		$random = new-object System.Random;

		# 7/12/2009 is a 'made up date' that happens to be a Sunday. 
		$dateObj= new-object System.DateTime -ArgumentList (2009, 7, 12, $random.Next(0,24), ($random.Next(0,4)*15), 0, 0, 'Utc'); 
		$stringSchedule = "";

		for($i=0; $i -lt 7;$i++) 
		{
			if($stringSchedule -ne "")
			{	
				$stringSchedule += ", "
			}

			$stringSchedule += $dateObj.AddDays($i).ToString("ddd.HH:mm tt") + "-" + $dateObj.AddDays($i).AddMinutes($windowLengthInMinutes).ToString("ddd.HH:mm tt"); 
		}

		return $stringSchedule
	}
}

#------------------------------------------------------------------------------
# Function: Get-TimeWindowStart
# Purpose:
#	Giving a stringized time format "HH:mm tt", we generate a DateTime which the day
#	is Sunday and keep the specified time, the reason we need a Sunday based DateTime
#	instance is the maintenance schedule is weekly based, and the schedule starts from Sunday
# Parameters:
#	$strtime	stringized time format in "HH:mm tt"
# Return:
#	DateTime of 10/10/2010 with time specified on input parameter
#------------------------------------------------------------------------------
function local:Get-TimeWindowStart($strtime)
{
	[DateTime]$timestart = [DateTime]::ParseExact($strtime, "h:m tt", $null)

	# 10/10/2010 is "Sunday"
	$ret = new-object System.DateTime -ArgumentList (2010, 10, 10, $timestart.Hour, $timestart.Minute, 0)
	return $ret
}

#------------------------------------------------------------------------------
# Function: Get-MaintenanceWeekdays
# Purpose:
#	Random select weekdays based on the occurrences, this is done by first
#	populate a full list (day 0 - 6 which representing Sunday - Satarday), 
#	then do (7-$occurrences) remove operation. We are doing this becuase
#	the so called random number might same value and we can not add the same
#	weekday more than once
# Parameters:
#	$ocurrences:	Number of occurrence in one week
# Return:
#	Array representation of weekdays which the maintenance task can be kicked off
#------------------------------------------------------------------------------
function local:Get-MaintenanceWeekdays([int]$occurrences)
{
	$random = new-object System.Random
	$days = New-Object System.Collections.ArrayList

	# Populate a full list of weekdays, 0-Sun. 1-Mon. etc
	[void]$days.Add(0)
	[void]$days.Add(1)
	[void]$days.Add(2)
	[void]$days.Add(3)
	[void]$days.Add(4)
	[void]$days.Add(5)
	[void]$days.Add(6)

	# Calculate the number of days need to be removed from list randomly
	$daystoremove = 7 - ($occurrences % 8)
	while ($daystoremove -gt 0)
	{
		$days.RemoveAt($random.Next(0, $days.Count))
		$daystoremove -= 1
	}
	return $days
}

#------------------------------------------------------------------------------
# Function: Get-MaintenanceStartTimeOfDay
# Purpose:
#	Giving the maintenance window start time and size (in minutes), generate a
#	qualified start time for that day based on required length (in minutes)
# Parameters
#	WindowStart		DateTime of start time for the maint window
#	windowLength	the maint window length in minutes
#	requiredLength	required length in minutes
# Return:
#	DateTime		datetime for the actualy start time on that day (or the next day)
#------------------------------------------------------------------------------
function local:Get-MaintenanceStartTimeOfDay([DateTime]$windowStart, [int]$windowLength, [int]$requiredLength)
{
	$random = new-object System.Random
	$availableLength = $windowLength - $requiredLength
	if ($availableLength -le 0)
	{
		Write-Error "No available window left to satisfy the requested maintenance task"
	}
	
	# Granularity: 15 minutes
	$granularity = 15

	# Random start time for that day
	return $windowStart.AddMinutes($granularity * $random.Next(0, $availableLength / $granularity))
}

#------------------------------------------------------------------------------
# Function: Get-WeeklyMaintenanceSchedule
# Purpose:
#	Generate a stringized weekly schedule, like the following line
#	"Mon.02:15 AM-Mon.03:15 AM,Thu.02:15 AM-Thu.03:15 AM,Fri.02:15 AM-Fri.03:15 AM"
# Parameters
#	windowStartSunday	window start time (represented by absoluate time of Sunday)
#	windowLength		window length in minutes
#	requiredLength		Caller requested length in minutes for its maint task
#	occurrences			Occurrences per week
# Return:
#	Stringized weekly schedule
#------------------------------------------------------------------------------
function local:Get-WeeklyMaintenanceSchedule([DateTime]$windowStartSunday, [int]$windowLength, [int]$requiredLength, [int]$occurrences)
{
	$strSchedule = ""

	# First get weekdays this task will be run 
	$weekdays = Get-MaintenanceWeekdays -occurrences $occurrences
	foreach($weekday in $weekdays)
	{
		# For each weekday, generate a qualified start time
		$starttimeSunday = Get-MaintenanceStartTimeOfDay -windowStart $windowStartSunday -windowLength $windowLength -requiredLength $requiredLength
		$starttimeInstance = $starttimeSunday.AddDays($weekday)
		$endtimeInstance = $starttimeInstance.AddMinutes($requiredLength)
		
		if ($strSchedule -ne "")
		{
			$strSchedule += ","
		}
		
		# Convert to the format can be accepted by Set-mailboxdatabase
		$strSchedule += $starttimeInstance.ToString("ddd.hh:mm tt") + "-" + $endtimeInstance.ToString("ddd.hh:mm tt")
	}
	return $strSchedule
}

#==============================================================================
# Main
#==============================================================================
$TimeWindowStart = Get-TimeWindowStart -strtime $StartTime

# Validate window length in hours
if ($TimeWindowLength -ge 12)
{
	$TimeWindowLength = 12
}
else
{
	if ($TimeWindowLength -le 0)
	{
		$TimeWindowLength = 4
	}
}

# Window length in minutes representation
$windowLengthInMinutes = $TimeWindowLength * 60

# Maint schedule window length in minutes
$MaintrequiredLength = 3 * 60

# Quota schedule window length in minutes
$QuotaRequiredLength = 15

# Get database list need to be configured
$alldatabases = Get-MailboxDatabase | where { $_.MasterServerOrAvailabilityGroup -ieq $DagOrServerName }

# Re-arrange the maintenance schedule and quota notification schedule (ranmon start time) for each database in the specified dag or server
$alldatabases | foreach {
	$maintschedule = Get-WeeklyMaintenanceSchedule -windowStartSunday $TimeWindowStart -windowLength $windowLengthInMinutes -requiredLength $MaintrequiredLength -occurrences $OccurrencePerWeek
	$quotaschedule = Get-WeeklyMaintenanceSchedule -windowStartSunday $TimeWindowStart -windowLength $windowLengthInMinutes -requiredLength $QuotaRequiredLength -occurrences $OccurrencePerWeek
	$_ | set-mailboxdatabase -MaintenanceSchedule ($maintschedule) -QuotaNotificationSchedule $quotaschedule
}

# SIG # Begin signature block
# MIIXOAYJKoZIhvcNAQcCoIIXKTCCFyUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUyl42+YRRZMtHXOuSeHa2D7Gf
# RZygghIxMIIEYDCCA0ygAwIBAgIKLqsR3FD/XJ3LwDAJBgUrDgMCHQUAMHAxKzAp
# BgNVBAsTIkNvcHlyaWdodCAoYykgMTk5NyBNaWNyb3NvZnQgQ29ycC4xHjAcBgNV
# BAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFJv
# b3QgQXV0aG9yaXR5MB4XDTA3MDgyMjIyMzEwMloXDTEyMDgyNTA3MDAwMFoweTEL
# MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
# bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEjMCEGA1UEAxMaTWlj
# cm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAw
# ggEKAoIBAQC3eX3WXbNFOag0rDHa+SU1SXfA+x+ex0Vx79FG6NSMw2tMUmL0mQLD
# TdhJbC8kPmW/ziO3C0i3f3XdRb2qjw5QxSUr8qDnDSMf0UEk+mKZzxlFpZNKH5nN
# sy8iw0otfG/ZFR47jDkQOd29KfRmOy0BMv/+J0imtWwBh5z7urJjf4L5XKCBhIWO
# sPK4lKPPOKZQhRcnh07dMPYAPfTG+T2BvobtbDmnLjT2tC6vCn1ikXhmnJhzDYav
# 8sTzILlPEo1jyyzZMkUZ7rtKljtQUxjOZlF5qq2HyFY+n4JQiG4FsTXBeyS9UmY9
# mU7MK34zboRHBtGe0EqGAm6GAKTAh99TAgMBAAGjgfowgfcwEwYDVR0lBAwwCgYI
# KwYBBQUHAwMwgaIGA1UdAQSBmjCBl4AQW9Bw72lyniNRfhSyTY7/y6FyMHAxKzAp
# BgNVBAsTIkNvcHlyaWdodCAoYykgMTk5NyBNaWNyb3NvZnQgQ29ycC4xHjAcBgNV
# BAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFJv
# b3QgQXV0aG9yaXR5gg8AwQCLPDyIEdE+9mPs30AwDwYDVR0TAQH/BAUwAwEB/zAd
# BgNVHQ4EFgQUzB3OdgBwW6/x2sROmlFELqNEY/AwCwYDVR0PBAQDAgGGMAkGBSsO
# AwIdBQADggEBAHurrn5KJvLOvE50olgndCp1s4b9q0yUeABN6crrGNxpxQ6ifPMC
# Q8bKh8z4U8zCn71Wb/BjRKlEAO6WyJrVHLgLnxkNlNfaHq0pfe/tpnOsj945jj2Y
# arw4bdKIryP93+nWaQmRiL3+4QC7NPP3fPkQEi4F6ymWk0JrKHG3OI/gBw3JXWjN
# vYBBa2aou7e7jjTK8gMQfHr10uBC33v+4eGs/vbf1Q2zcNaS40+2OKJ8LdQ92zQL
# YjcCn4FqI4n2XGOPsFq7OddgjFWEGjP1O5igggyiX4uzLLehpcur2iC2vzAZhSAU
# DSq8UvRB4F4w45IoaYfBcOLzp6vOgEJydg4wggR6MIIDYqADAgECAgphAc8+AAAA
# AAAPMA0GCSqGSIb3DQEBBQUAMHkxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNo
# aW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24xIzAhBgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBMB4X
# DTA5MTIwNzIyNDAyOVoXDTExMDMwNzIyNDAyOVowgYMxCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xDTALBgNVBAsTBE1PUFIxHjAcBgNVBAMTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoC
# ggEBAL0wiftFcqhTa56JTwAjwL7UHT2xWUA483ORgibmlhIAU9kcgg48zh27vfdC
# jZfU/Dga5Ln54+zTYQO/oNPWdU1cRqntXvDS4mlbGnPqsxyNBM0pRKBkWS8emF1u
# x6sYOYJlxKe8q3WIGeqHlxQms38mdqTUODmE47Mm1Rj5K+nSyRZaVCHyl42Hhin+
# 9Eks5ov4BD99zc2WkoYNcQPi0P4MQjX/17g/3Y5FCn3210utW/B2ch13I32JNcQc
# XbJQA0tHbQenVYiYBoCmga1UTtiB1vq/QsAxvlUNmdVTSRIw6+WliHxexHpaFIcI
# tDdpoOsyJIwI6/nUFLrg/M3qpBUCAwEAAaOB+DCB9TATBgNVHSUEDDAKBggrBgEF
# BQcDAzAdBgNVHQ4EFgQUOHgFc8gbMptfkoZVr4m6xpmxdI4wDgYDVR0PAQH/BAQD
# AgeAMB8GA1UdIwQYMBaAFMwdznYAcFuv8drETppRRC6jRGPwMEQGA1UdHwQ9MDsw
# OaA3oDWGM2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L0NTUENBLmNybDBIBggrBgEFBQcBAQQ8MDowOAYIKwYBBQUHMAKGLGh0dHA6Ly93
# d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvQ1NQQ0EuY3J0MA0GCSqGSIb3DQEB
# BQUAA4IBAQAoA4OrEw3x/hFvQnbYQ6u4xkL6JGU0rnWwcTCOlraD+nLk/s035zpw
# mFEgDrpxtMn01eSK6Chwc9Xt4cJ77vFWNFqAgusQ13ID6hJO7N0NjHtETwxhdf1u
# nAfdzitaGA05sGD92v8EqgKj2PqOo7u9wWpUOzFf5Cxuy/FNKhzyHnZ7hq/Gh9ax
# Ku2szA7A5J9/6M/ohHJUghhuKsielbQaHsFvk9tipsFLa63J7T2S47Ivq3p3OWdq
# AHuCnQMQ3mOVSSZC3pUiwYcL0BSjX9xzMnI2Hv/IKcUkUk9mSMkdHanFJjlpUE6n
# 7MIGbC5yG/Eh2m0kSICBFUB8n8x0OEpPMIIEnTCCA4WgAwIBAgIQaguZT8AAJasR
# 20UfWHpnojANBgkqhkiG9w0BAQUFADBwMSswKQYDVQQLEyJDb3B5cmlnaHQgKGMp
# IDE5OTcgTWljcm9zb2Z0IENvcnAuMR4wHAYDVQQLExVNaWNyb3NvZnQgQ29ycG9y
# YXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBSb290IEF1dGhvcml0eTAeFw0wNjA5
# MTYwMTA0NDdaFw0xOTA5MTUwNzAwMDBaMHkxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xIzAhBgNVBAMTGk1pY3Jvc29mdCBUaW1lc3RhbXBpbmcg
# UENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA3Ddu+6/IQkpxGMjO
# SD5TwPqrFLosMrsST1LIg+0+M9lJMZIotpFk4B9QhLrCS9F/Bfjvdb6Lx6jVrmlw
# ZngnZui2t++Fuc3uqv0SpAtZIikvz0DZVgQbdrVtZG1KVNvd8d6/n4PHgN9/TAI3
# lPXAnghWHmhHzdnAdlwvfbYlBLRWW2ocY/+AfDzu1QQlTTl3dAddwlzYhjcsdckO
# 6h45CXx2/p1sbnrg7D6Pl55xDl8qTxhiYDKe0oNOKyJcaEWL3i+EEFCy+bUajWzu
# JZsT+MsQ14UO9IJ2czbGlXqizGAG7AWwhjO3+JRbhEGEWIWUbrAfLEjMb5xD4Gro
# fyaOawIDAQABo4IBKDCCASQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwgaIGA1UdAQSB
# mjCBl4AQW9Bw72lyniNRfhSyTY7/y6FyMHAxKzApBgNVBAsTIkNvcHlyaWdodCAo
# YykgMTk5NyBNaWNyb3NvZnQgQ29ycC4xHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFJvb3QgQXV0aG9yaXR5gg8AwQCL
# PDyIEdE+9mPs30AwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFG/oTj+XuTSr
# S4aPvJzqrDtBQ8bQMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQE
# AwIBhjAPBgNVHRMBAf8EBTADAQH/MA0GCSqGSIb3DQEBBQUAA4IBAQCUTRExwnxQ
# uxGOoWEHAQ6McEWN73NUvT8JBS3/uFFThRztOZG3o1YL3oy2OxvR+6ynybexUSEb
# bwhpfmsDoiJG7Wy0bXwiuEbThPOND74HijbB637pcF1Fn5LSzM7djsDhvyrNfOzJ
# rjLVh7nLY8Q20Rghv3beO5qzG3OeIYjYtLQSVIz0nMJlSpooJpxgig87xxNleEi7
# z62DOk+wYljeMOnpOR3jifLaOYH5EyGMZIBjBgSW8poCQy97Roi6/wLZZflK3toD
# dJOzBW4MzJ3cKGF8SPEXnBEhOAIch6wGxZYyuOVAxlM9vamJ3uhmN430IpaczLB3
# VFE61nJEsiP2MIIEqjCCA5KgAwIBAgIKYQWiMAAAAAAACDANBgkqhkiG9w0BAQUF
# ADB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSMwIQYDVQQD
# ExpNaWNyb3NvZnQgVGltZXN0YW1waW5nIFBDQTAeFw0wODA3MjUxOTAxMTVaFw0x
# MzA3MjUxOTExMTVaMIGzMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMQ0wCwYDVQQLEwRNT1BSMScwJQYDVQQLEx5uQ2lwaGVyIERTRSBFU046ODVE
# My0zMDVDLTVCQ0YxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZp
# Y2UwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDwBC2ylsAagWclsSZi
# sxNLzjC6wBI4/IFlNAfENrIkaPYHBMAHl/S38XseYixG2UukUTS302ztWju0g6FH
# PREILjVrRebCPIwCZgKpGGnrSu0nLO48d1uk1HCZS1eEENCvLfiJHebqKbTnz54G
# YqdyVMI7xs8/uOGwWBBs5aXXw8J1N730heGB6CjYG/HyrvGCo9bXA6KfFYT7Pfqr
# 4bYyyKACZPPm/xomcQhTihUC8oMndkmCcafvrTJ4xtdsFk8iZZdiTUYv/yOvheym
# cL0Dy9rYMgXFK5BAtp7VLIZst8sTMn2Nxn6uFy8y/Ga7HbBFVfit+i1ng2cpk4TS
# WqEjAgMBAAGjgfgwgfUwHQYDVR0OBBYEFOiX9vfvjPHmaeNZaE73mIp63ZsuMB8G
# A1UdIwQYMBaAFG/oTj+XuTSrS4aPvJzqrDtBQ8bQMEQGA1UdHwQ9MDswOaA3oDWG
# M2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL3RzcGNh
# LmNybDBIBggrBgEFBQcBAQQ8MDowOAYIKwYBBQUHMAKGLGh0dHA6Ly93d3cubWlj
# cm9zb2Z0LmNvbS9wa2kvY2VydHMvdHNwY2EuY3J0MBMGA1UdJQQMMAoGCCsGAQUF
# BwMIMA4GA1UdDwEB/wQEAwIGwDANBgkqhkiG9w0BAQUFAAOCAQEADT93X5E8vqU1
# pNsFBYQfVvLvmabHCI0vs80/cdWGfHcf3esXsr184/mZ8gpFSK0Uu2ks8j5nYlTy
# 7n8nEZI57M7Zh06I92BHI3snFUAIn78NMQSC2DW2DJwA04uqeGHFtYhBnT423Fik
# J5s62r0GXRSmsg9MwY48i/Jimfhm7dXzHCiwMtvKMQm8+yJoRkz603Mi5ymOIgD7
# Vr8GroGgFbo0+SiOH0piBaGJ9YFH6Q2RCNdYO48eawlpqcBIfFWCP18AOEOcBsw/
# 2C+/T3MJPf26XvTH7DfCZGGgTdQ9cMxbsBOBwdSjMRq9ZNaW0no/KltGUwk8zQP5
# P1kAzIlTYTGCBHEwggRtAgEBMIGHMHkxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xIzAhBgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENB
# AgphAc8+AAAAAAAPMAkGBSsOAwIaBQCggZwwGQYJKoZIhvcNAQkDMQwGCisGAQQB
# gjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkE
# MRYEFGHjZagmTi8CmMEimAqWXmwm9gmuMDwGCisGAQQBgjcCAQwxLjAsoBKAEABD
# AG8AZABlAFMAaQBnAG6hFoAUaHR0cDovL0NvZGVTaWduSW5mbyAwDQYJKoZIhvcN
# AQEBBQAEggEALnjdzoAukMUbAKIn7mnayADMudUvJhxTMZfmJtFm6QA/OLOkjJcO
# 4eysHtDIG0yGUfgsXe7NKQ3CEsVSooISW4Vxwj6J1YeNSBa1O4sJbPlTtGiSf2Xc
# sW7jbWvCKt1ddXKOQ3uXoJ1ky34TOYvpEYMwQyTJtkA4/wvzqhcHFoE3tG+Ey4mQ
# kh7rd6OZc2lGpWhAhkiOezRUra2/fSpOi/4dEWA2up2G1W0fXWbJKL8rQhrCGLmm
# JzXM1nVcrJHySNjv76+oJkavnzfy2GOEpbj9bHkRIUtKXPUjonYgHs6CzXB3wtmc
# lEIQBbJwHGeaiU/B9q8Ojx4f3c7/GVtBLKGCAh8wggIbBgkqhkiG9w0BCQYxggIM
# MIICCAIBATCBhzB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSMwIQYDVQQDExpNaWNyb3NvZnQgVGltZXN0YW1waW5nIFBDQQIKYQWiMAAAAAAA
# CDAHBgUrDgMCGqBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcN
# AQkFMQ8XDTEwMTAyNjE1MjUxMVowIwYJKoZIhvcNAQkEMRYEFDYpzh1NXSu40KFj
# ktcv8r6a0okeMA0GCSqGSIb3DQEBBQUABIIBALd3zFvrCKlwczv2YzD21cB3ZzLO
# s6X4fDTRo8UAOaAyJ44F3eC7pdFxx6f1SnId68ugsUAXraCro2O8aaN3yAVP0Zp0
# Jq6NZLPi7m+Ayn7uvFcVP/+GMr8RaaMYjEfez1cdmTdnMmSmXpI9sIeQyaJTbMWo
# xFYXz/wGsJB7VpAQ0apUWBFb9YO6AW2em60/SfloqTkEgtTR++WQyYuw+fu15xAl
# sn5R1Sq3rDMrdSpUMTYHAOAOgZYTfAhPyDkUKXaCB2f8BtqA7gN/OQWOV8m/aQuk
# 8pinHYErOgAcjzZj9tLzFJTzWv20yKj+74bxZuPQTWFlUZYbc9Nmoj3e6gU=
# SIG # End signature block

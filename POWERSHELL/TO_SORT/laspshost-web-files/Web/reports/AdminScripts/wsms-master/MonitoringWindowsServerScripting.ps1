#########################################################
#                                                       #
# Monitoring Windows Server Scripting                   #
#                                                       #
#########################################################

# Thanks to: https://gallery.technet.microsoft.com/scriptcenter/Windows-Updates-and-684c355c

#########################################################
# Configurations, see "configurations.txt"
#########################################################
#Get-Content ".\Configurations.txt" | foreach-object -begin {$h=@{}} -process { $k = [regex]::split($_,'='); if(($k[0].CompareTo("") -ne 0) -and ($k[0].StartsWith("[") -ne $True)) { $h.Add($k[0], $k[1]) } }
$Server = Get-Content ".\configs\mtsrvs.txt" #$h.Get_Item("Server")
$mailto = "john.battista@creditone.com" #$h.Get_Item("mailto")
$emailFrom = "ServerReporter@creditone.com" #$h.Get_Item("emailFrom")
$smtpServer = "mailgateway.Contoso.corp" #$h.Get_Item("smtpServer")

#########################################################
# Formatting Result
#########################################################

# who is this machine
$tableFragment = Get-CimInstance Win32_OperatingSystem | Select-Object Caption, InstallDate, ServicePackMajorVersion, OSArchitecture, BuildNumber, CSName | ConvertTo-HTML -fragment

# Windows Update Summary Object
$WindowsUpdateSummary = try {
        $service = Get-WmiObject Win32_Service -Filter 'Name="wuauserv"' -ComputerName $Server -Ea 0
        $WUStartMode = $service.StartMode
        $WUState = $service.State
        $WUStatus = $service.Status

        $UpdateSession = New-Object -ComObject Microsoft.Update.Session
        $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
        $SearchResult = $UpdateSearcher.Search("IsAssigned=1 and IsHidden=0 and IsInstalled=0")
        $Critical = $SearchResult.updates | where { $_.MsrcSeverity -eq "Critical" }
        $important = $SearchResult.updates | where { $_.MsrcSeverity -eq "Important" }
        $other = $SearchResult.updates | where { $_.MsrcSeverity -eq $null }
        # Get windows updates counters
        $totalUpdates = $($SearchResult.updates.count)
        $totalCriticalUp = $($Critical.count)
        $totalImportantUp = $($Important.count)

        if($totalUpdates -gt 0) {
                $updatesToInstall = $true
        } else {
                $updatesToInstall = $false
        }

        # Querying WMI for build version
        $WMI_OS = Get-WmiObject -Class Win32_OperatingSystem -Property BuildNumber, CSName -ComputerName $Server -Authentication PacketPrivacy -Impersonation Impersonate

        # Making registry connection to the local/remote computer
        $RegCon = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]"LocalMachine",$Server)

        # If Vista/2008 & Above query the CBS Reg Key
        If ($WMI_OS.BuildNumber -ge 6001) {
        $RegSubKeysCBS = $RegCon.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\").GetSubKeyNames()
        $CBSRebootPend = $RegSubKeysCBS -contains "RebootPending"
        } else {
                $CBSRebootPend = $false
        }

        # Query WUAU from the registry
        $RegWUAU = $RegCon.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\")
        $RegSubKeysWUAU = $RegWUAU.GetSubKeyNames()
        $WUAURebootReq = $RegSubKeysWUAU -contains "RebootRequired"

        If ($CBSRebootPend -OR $WUAURebootReq) {
                $machineNeedsRestart = $true
        } else {
                $machineNeedsRestart = $false
        }

        # Closing registry connection
        $RegCon.Close()

        if ($machineNeedsRestart -or $updatesToInstall -or ($WUStartMode -eq "Manual") -or ($totalUpdates -eq "nd")) {
                New-Object PSObject -Property @{
                        Computer = $WMI_OS.CSName
                        WindowsUpdateStatus = $WUStartMode + "/" + $WUState + "/" + $WUStatus
                        UpdatesToInstall = $updatesToInstall
                        TotalOfUpdates = $totalUpdates
                        TotalOfCriticalUpdates = $totalCriticalUp
                        TotalOfImportantUpdates = $totalImportantUp
                        RebootPending = $machineNeedsRestart
                }
        }
} Catch {
    Write-Warning "$Server`: $_"
}
$tableFragment += $WindowsUpdateSummary | ConvertTo-HTML -fragment

# list last 10 hotfixes
$Session = New-Object -ComObject "Microsoft.Update.Session"
$Searcher = $Session.CreateUpdateSearcher()
$historyCount = $Searcher.GetTotalHistoryCount()
$UpdateHistory = $Searcher.QueryHistory(0, $historyCount)
$hotfixesToShow = $UpdateHistory | Select-Object Date,Title | Select-Object -First 20
$hotfixesToShow += new-object psobject -property @{
        "Date" = ""
        "Title" = "(top 20 hot-fixes only)"
}
$tableFragment += $hotfixesToShow | ConvertTo-HTML -fragment

# security auditing
$group = [ADSI]("WinNT://localhost/Administrators,group")
$admins = $group.PSBase.Invoke('Members') | % { $_.GetType().InvokeMember('Name', 'GetProperty', $null, $_, $null) }

$tableFragment += $admins | foreach { $array = (net user $_); new-object psobject -property @{
        Account = $_
        "Full Name" = ($array[1] -split "\s+")[1]
        "In Use" = ($array[5] -split "\s+")[1]
        "Last Login" = ($array[18] -split "\s+")[1]
        "Last Password Modified" = ($array[8] -split "\s+")[1]
        "Password Expired" = ($array[9] -split "\s+")[1]
} } | ConvertTo-HTML -fragment

function get-securityauditing {
        Param (
                [string]$Computer = (Read-Host Remote computer name),
                [int]$Days = 10
        )
        $Result = @()
        $ELogs = Get-EventLog System -Source Microsoft-Windows-WinLogon -After (Get-Date).AddDays(-$Days) -ComputerName $Computer
        If ($ELogs) {
                ForEach ($Log in $ELogs) {
                        If ($Log.InstanceId -eq 7001) {
                                $ET = "Logon"
                        } ElseIf ($Log.InstanceId -eq 7002) {
                                $ET = "Logoff"
                        } Else {
                                Continue
                        }
                        $Result += New-Object PSObject -Property @{
                                Time = $Log.TimeWritten
                                'Event Type' = $ET
                                User = (New-Object System.Security.Principal.SecurityIdentifier $Log.ReplacementStrings[1]).Translate([System.Security.Principal.NTAccount])
                                Admin = 'System'
                        }
                }
        }
        $ELogs = Get-EventLog Security -After (Get-Date).AddDays(-$Days) -ComputerName $Computer
        If ($ELogs) {
                ForEach ($Log in $ELogs) {
                        If ($Log.InstanceId -eq 4720) {
                                $ET = "Creation"
                                $U = $Log.ReplacementStrings[0]
                                $R = $Log.ReplacementStrings[4]
                        } ElseIf ($Log.InstanceId -eq 4722) {
                                $ET = "Enabled"
                                $U = $Log.ReplacementStrings[0]
                                $R = $Log.ReplacementStrings[4]
                        } ElseIf ($Log.InstanceId -eq 4723) {
                                $ET = "Password Changed"
                                $U = $Log.ReplacementStrings[0]
                                $R = 'Security'
                        } ElseIf ($Log.InstanceId -eq 4724) {
                                $ET = "Password Reset"
                                $U = $Log.ReplacementStrings[0]
                                $R = $Log.ReplacementStrings[4]
                        } ElseIf ($Log.InstanceId -eq 4725) {
                                $ET = "Disabled"
                                $U = $Log.ReplacementStrings[0]
                                $R = 'Security'
                        } ElseIf ($Log.InstanceId -eq 4726) {
                                $ET = "Deleted"
                                $U = $Log.ReplacementStrings[0]
                                $R = $Log.ReplacementStrings[4]
                        } ElseIf ($Log.InstanceId -eq 4738) {
                                $ET = "Changed"
                                $U = $Log.ReplacementStrings[1]
                                $R = $Log.ReplacementStrings[5]
                        } ElseIf ($Log.InstanceId -eq 4740) {
                                $ET = "Locked out"
                                $U = $Log.ReplacementStrings
                                $R = 'Security'
                        } ElseIf ($Log.InstanceId -eq 4767) {
                                $ET = "Unlocked"
                                $U = $Log.ReplacementStrings[0]
                                $R = 'Security'
                        } ElseIf ($Log.InstanceId -eq 4780) {
                                $ET = "Set as administrators"
                                $U = $Log.ReplacementStrings[0]
                                $R = 'Security'
                        } ElseIf ($Log.InstanceId -eq 4781) {
                                $ET = "Name changed"
                                $U = $Log.ReplacementStrings[0]
                                $R = 'Security'
                        } ElseIf ($Log.InstanceId -eq 4719) {
                                $ET = "Policy changed"
                                $U = "Policy:" + $Log.ReplacementStrings[5]
                                $R = $Log.ReplacementStrings[1]
                        } Else {
                                Continue
                        }
                        $Result += New-Object PSObject -Property @{
                                Time = $Log.TimeWritten
                                'Event Type' = $ET
                                User = $U
                                Admin = $R
                        }
                }
        }
        return $Result | Select Time,"Event Type",User,Admin| Sort-Object Time -descending
}
$tableFragment += get-securityauditing $Server 360 | ConvertTo-HTML -fragment

# auditing audit policy
$tableFragment += auditpol /get /category:* /r | ConvertFrom-Csv | Select Subcategory,"Inclusion Setting" | ConvertTo-HTML -fragment

# networking and NTP
$tableFragment += Get-NetIPAddress | Select-Object IPAddress,InterfaceAlias | ConvertTo-HTML -fragment
$ntps = w32tm /query /configuration | ?{$_ -match 'ntpserver:'} | %{($_ -split ":\s\b")[1]}
$tableFragment += new-object psobject -property @{
    Server = $Server
    NTPSource = $ntps
} | Select-Object Server,NTPSource | ConvertTo-HTML -fragment

# HTML Format for Output
$HTMLmessage = @"
<font color=""black"" face=""Arial"" size=""3"">
<h1 style='font-family:arial;'><b>Windows Server Monitoring Report for $Server</b></h1>
<p style='font: .8em ""Lucida Grande"", Tahoma, Arial, Helvetica, sans-serif;'>This report includes System Information, Last 20 Windows Updates, and Audits Information.</p>
<br><br>
<style type=""text/css"">body{font: .8em ""Lucida Grande"", Tahoma, Arial, Helvetica, sans-serif;}
ol{margin:0;}
table{width:80%;}
thead{}
thead th{font-size:120%;text-align:left;}
th{border-bottom:2px solid rgb(79,129,189);border-top:2px solid rgb(79,129,189);padding-bottom:10px;padding-top:10px;}
tr{padding:10px 10px 10px 10px;border:none;}
#middle{background-color:#900;}
</style>
<body BGCOLOR=""white"">
$tableFragment
</body>
"@

#########################################################
# Validation and sending email
#########################################################
# Regular expression to get what's inside of <td>'s
$regexsubject = $HTMLmessage
$regex = [regex] '(?im)<td>'

# If you have data between <td>'s then you need to send the email
if ($regex.IsMatch($regexsubject)) {
     $smtp = New-Object Net.Mail.SmtpClient -ArgumentList $smtpServer
     $msg = New-Object Net.Mail.MailMessage
     $msg.From = $emailFrom
     $msg.To.Add($mailto)
     $msg.Subject = "Monitoring Windows Server for $Server"
     $msg.IsBodyHTML = $true
     $msg.Body = $HTMLmessage
     $smtp.Send($msg)
}

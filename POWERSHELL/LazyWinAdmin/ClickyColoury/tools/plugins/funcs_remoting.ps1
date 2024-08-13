# api: multitool
# version: 0.6
# title: Remoting funcs
# description: Start-SchedTask, WMI ScheduledJob, PSRemoting
# type: init
# category: functions
# status: TESTING
# hidden: 0
# config: -
# 
# Fail:
#  - psexec does not allow interactive apps in foreground even with -i and session id (from WMI/win32_process)
#  - Win32_ScheduledJob cannot be run immediately; requires latency timeing and timestamp crafting
#
# Working approach:
#  ❏ schtasks.exe /Create /S A0004915 /TN TeamViewer /RU DOMAIN\TestAccount
#    /IT /TR "'C:\Program Files\TeamViewerQS\Team ViewerQS-idcm7ct8bq.exe'"
#    /SC ONCE /SD 05/05/2022 /ST 23:59 /F
#  ❏ schtasks /run /S A0004915 /TN TeamViewer
#


#-- run as scheduled job
#
function Start-SchedTask {
    Param(
        $machine = "TESTHOST",
        $cmd = "'cmd.exe'",           # Note the double string context quoting "'...'" for paths with spaces!
        $taskname = "psonce"
    )

    #-- ping
    if (!(Test-Connection $machine -Quiet -ErrorAction SilentlyContinue)) {
        Write-Host -f Red " $machine offline"
        return
    }

    #-- current user
    $user = GWMI win32_ComputerSystem -ComputerName $machine | select -expand username
    Write-Host -f Yellow "❏ $user on $machine"

    #-- sched task
    Write-Host -f Yellow "❏ scheduling $cmd"
    schtasks.exe /Create /S $machine /TN $taskname /RU "$($cfg.domain)\$username" /IT /TR "$cmd" /SC ONCE /SD 05/05/2022 /ST 23:59 /F

    Write-Host -f Yellow "❏ starting task..."
    schtasks.exe /run /S $machine /TN $taskname
}

    
#-- check for session id (`QUSER` doesn't work)
#
function Get-RemoteSessionID {
    Param(
        $machine
    )
    $sid = Get-WmiObject Win32_Process -ComputerName $machine -Filter 'Name="explorer.exe"' | Select -First 1 -Expand SessionId
    if ($sid) {
        return $sid
    }
    else {
        Write-Host -f Red "✘ No user desktop session found; assuming default 1"
        return 1
    }
}


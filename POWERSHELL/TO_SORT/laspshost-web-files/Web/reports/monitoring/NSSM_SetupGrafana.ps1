#utility script 
#just highlight text and run with F8 to run selected text
#get details on steps to create the service already existing
<#
        #GET GUI POPUP TO INSTALL
        &$NSSMEXE install 
        # open executable with no args gives popup box to help understand options 
        ii $NSSMEXE 
                
        #dump details on installed version
        &$NSSMEXE dump $exebasename
        #edit existing service
        &$NSSMEXE stop $ExeBaseName
        &$NSSMEXE edit $exebasename 
        &$NSSMEXE stop $ExeBaseName
        &$NSSMEXE start $ExeBaseName
        get-eventlog -LogName Application -Newest 10 | out-gridview
#>

$VerbosePreference      = 'Continue'
$InformationPreference  = 'Continue'
$DebugPreference        = 'Continue'
$ErrorActionPreference  = 'Stop'


$NSSMEXE = 'C:\Scripts\Repository\jbattista\Web\reports\monitoring\bin\nssm\nssm.exe'

clear-host
[void]$error.clear()
#setup statements to automate install
$EXE = 'C:\Scripts\Repository\jbattista\Web\reports\monitoring\bin\grafana\bin\grafana-server.exe'
$config = "C:\Scripts\Repository\jbattista\Web\reports\monitoring\bin\grafana\conf\custom.ini"
$ExeBaseName = [io.path]::GetFileNameWithoutExtension($Exe)
$ParentDirectory = [io.directory]::GetParent($Exe)
if(!(test-path variable:creds)) {[PSCredential]$creds = (Get-Credential -UserName "$env:USERDOMAIN\$env:Username" -Message "Enter service account credentials") }
$ServiceAccount = $creds.GetNetworkCredential().Domain + "\" + $creds.GetNetworkCredential().UserName
$ServiceAccountPassword = $creds.GetNetworkCredential().Password
$MBFileSize  = 25


return


#uninstall
if(get-service -name $ExeBaseName -ea SilentlyContinue){ 
        stop-service -name $ExeBaseName -Force -ea SilentlyContinue
        &$NSSMEXE remove $ExeBaseName confirm
}
start-sleep -Seconds 1 -Verbose


return 

[Collections.Generic.List[string]]$Statements = @()
[void]$Statements.Add(('install {1} "{0}" -config "{2}"' -f $exe, $ExeBaseName,$Config))
#[void]$Statements.Add('set {0} AppParameters' -f $ExeBaseName) #C:\telegraf\
[void]$Statements.Add("set $ExeBaseName AppNoConsole 1")
[void]$Statements.Add("set $ExeBaseName AppRestartDelay 60000")
[void]$Statements.Add(('set {1} AppDirectory "{0}"' -f $ParentDirectory,$exeBasename))  #THIS REPLACE ALL ENVIROMENTAL VARIABLES
[void]$Statements.Add("set $ExeBaseName AppExit Default Restart")
[void]$Statements.Add("set $ExeBaseName DisplayName $ExeBaseName")
[void]$Statements.Add('set {0} Description ""nssm installed {0}""' -f $exeBasename)
[void]$Statements.Add("set $ExeBaseName ObjectName $ServiceAccount $ServiceAccountPassword")

[void]$Statements.Add("set $ExeBaseName Start SERVICE_AUTO_START")




[void]$Statements.Add("set $ExeBaseName AppStdIn `"$($ParentDirectory)\stdin.log`"")
[void]$Statements.Add("set $ExeBaseName AppStdInCreationDisposition 2")
[void]$Statements.Add("set $ExeBaseName AppStderr `"$($ParentDirectory)\stderr.log`"")
[void]$Statements.Add("set $ExeBaseName AppStderrCreationDisposition 2")

[void]$Statements.Add("set $ExeBaseName AppStdout `"$($ParentDirectory)\stdout.log`"")
[void]$Statements.Add("set $ExeBaseName AppStdoutCreationDisposition 2")
[void]$Statements.Add("set $ExeBaseName AppRotateFiles 0")
[void]$Statements.Add("set $ExeBaseName AppTimestampLog 0")


#[void]$Statements.Add("set $ExeBaseName Type SERVICE_WIN32_OWN_PROCESS")
#[void]$Statements.Add("set $ExeBaseName AppPriority NORMAL_PRIORITY_CLASS")
#[void]$Statements.Add("set $ExeBaseName AppRotateOnline 0")
#[void]$Statements.Add(("set $ExeBaseName AppRotateBytes {0}" -f ([int]$MBFileSize * 1024 * 1024)))
#[void]$Statements.Add("set $ExeBaseName AppEnvironmentExtra TELEGRAF_CONFIG_PATH=`"$ParentDirectory`" CONFIG=`"$ParentDirectory\telegraf.conf`" HOME=`"$ParentDirectory`"")
#[void]$Statements.Add("set $ExeBaseName AppEvents Start/Pre $EXE")
#[void]$Statements.Add("set $ExeBaseName AppEvents Start/Post $EXE")
#[void]$Statements.Add("set $ExeBaseName ObjectName LocalSystem")
#[void]$Statements.Add("set $ExeBaseName Type SERVICE_WIN32_OWN_PROCESS")  #if installing as localsystem: SERVICE_WIN32_OWN_PROCESS, if installed as shared service using sc.exe this came up: SERVICE_WIN32_SHARE_PROCESS

$DebugPreference = 'continue'
write-debug "`n---- NSSM Arguments ---`n$($Statements | format-list | Out-String )"
write-debug "NSSM Executable: $NSSMEXE"
@($Statements).ForEach({write-debug("`n`n$(echoargs $_)")})


foreach($statement in $statements)
{ 
        try { 
                write-debug (($(echoargs $NSSMEXE $Statement) | out-string))
                $returned = Invoke-Executable -sExe $NSSMEXE -arguments $Statement -debug:$false -sVerb "runas"
                if($returned.ExitCode -ne 0){throw "Exit Code: $($Returned.ExitCode)"}

        }
        catch { 
                #catch error and run echoargs to verify what was passed that caused an issue
                write-warning ($(echoargs $NSSMEXE $Statement) | out-string)
                write-warning $_.Exception.Message


                stop-service -name $ExeBaseName -Force -ea SilentlyContinue
                &$NSSMEXE remove $ExeBaseName confirm
                break
        }
}

try {
    start-sleep -Seconds 1 -Verbose
    get-service -name $ExeBaseName | Start-Service -Verbose
    get-service -name $ExeBaseName | select DisplayName, Status
}
catch 
{

    get-eventlog -LogName Application -Newest 10 | out-gridview
    
}
return 

<#
                &$NSSMEXE stop $exebasename
                &$NSSMEXE start $exebasename
                &$NSSMEXE continue $exebasename
#>


<#
#### RUNNING THIS WAY DOES NOT WORK AS IT NEEDS HOMEPATH VARIABLE SET. NSSM MAKES IT EASIER, OTHERWISE NEED BAT LAUNCH ####
&sc.exe delete $exebasename
$SplatMe       = @{
    Name           = $ExeBaseName
    BinaryPathName = "$exe -config `"$config`""
    DisplayName    = $ExeBaseName
    Description    = 'powershell created service' 
    StartupType    = 'Auto' 
    Credential     = $creds
    Verbose        = $true
}
new-service @SplatMe
try {
    get-service -name $ExeBaseName | Start-Service -Verbose
    get-service -name $ExeBaseName | select DisplayName, Status
}
catch 
{
    get-eventlog -LogName Application -Newest 10 | out-gridview
    
}
# invoke-executable -sExeFile $exe -cArgs "-config `"$Config`""
#>  
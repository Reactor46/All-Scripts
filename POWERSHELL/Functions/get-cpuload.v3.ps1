#Requires -RunAsAdministrator

$ComputerList = @(
    'vdcmaddc01.corp.vegas.com',
'stg-theoapp02.corp.vegas.com',
'VVCNETAP02.corp.vegas.com',
'VDCNETAP01.corp.vegas.com',
'vvcexcwa01.corp.vegas.com',
'VVCEXCMB01.corp.vegas.com',
'VVCEXCWA02.corp.vegas.com',
'VVCEXCMB02.corp.vegas.com',
'VVCPRTAP01.corp.vegas.com',
'VVDBHRAP01.corp.vegas.com',
'vvsbhrap02.corp.vegas.com',
'VVSVDCDB01.corp.vegas.com'
    )
$IntervalSeconds = 1
$DataPointCount = 5

$ScriptBlock = {
    $Readings = foreach ($DPC_Item in 1..$Using:DataPointCount)
        {
        (Get-CimInstance -ClassName Win32_Processor).LoadPercentage
        Start-Sleep -Seconds $Using:IntervalSeconds
        }

    $AverageCPU_Load = ($Readings |
        Measure-Object -Average).Average

    [PSCustomObject]@{
        AverageCPU_Load = $AverageCPU_Load
        }
    }

$Non_ReachableSystems = [System.Collections.Generic.List[PSCustomObject]]@{}
$ReachableSystems = foreach ($CL_Item in $ComputerList)
    {
    if ((Test-Connection -ComputerName $CL_Item -Count 1 -Quiet) -and
        ([bool](Test-WSMan -ComputerName $CL_Item -ErrorAction SilentlyContinue)))
        {
        $CL_Item
        }
        else
        {
        $Non_ReachableSystems.Add([PSCustomObject]@{
            AverageCPU_Load = '-- No Response --'
            PSComputerName = $CL_Item
            })
        }
    }

$RS_Results = Invoke-Command -ComputerName $ReachableSystems -ScriptBlock $ScriptBlock

$Results = $RS_Results + $Non_ReachableSystems
$Results |
    Sort-Object -Property AverageCPU_Load -Descending |
    Select-Object -Property AverageCPU_Load, PSComputerName
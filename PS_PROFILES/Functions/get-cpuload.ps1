. .\Functions\Get-Server.ps1

$ComputerList = Get-Server -CORP

$InvokeCommandScriptBlock = {
    Get-WmiObject win32_processor | 
        Measure-Object -property LoadPercentage -Average | 
        Select-Object @{e={[math]::Round($_.Average,1)};n="CPU(%)"}
}

$InvokeCommandArgs = @{
    ComputerName = $ComputerList
    ScriptBlock  = $InvokeCommandScriptBlock
    ErrorAction  = "SilentlyContinue"
}

Invoke-Command @InvokeCommandArgs  | 
    Sort-Object "CPU(%)" -Descending | 
    Select-Object "CPU(%)",PSComputerName
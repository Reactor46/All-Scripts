function Get-ScriptsErrorReport
{
    Param
    ()

    Begin
    {
        $ReportArray = @()
    }
    Process
    {
        try
        {
            $ScriptExceptions = $global:Error

            foreach($Exception in $ScriptExceptions)
            {
                $InvocationInfo = $Exception.InvocationInfo

                $Properties = @{Script = [system.io.path]::GetFileName($InvocationInfo.ScriptName)
                                Path = $InvocationInfo.ScriptName
                                "Error Message" = $Exception.Exception.Message
                                "Error Details" = $InvocationInfo.PositionMessage
                                Line = $InvocationInfo.Line
                                Status = "Warning"}
                                    
                $Obj = New-Object psobject -Property $Properties
                $ReportArray += $Obj
            }  

        }

        catch
        {
            Write-Error $_.Exception.Message
        }

    }
    End
    {
        $ReportArray
    }
}

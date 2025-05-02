function Get-IISWebsides
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true, 
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        $ComputerName
    )
    
    Begin
    {
        $ReportArray = @()

        if(!$ComputerName)
        {
            $ComputerName = $env:COMPUTERNAME
        }
    }
    Process
    {
        $Sites = gwmi -namespace "root\webadministration" -Class site -ComputerName $ComputerName -Authentication PacketPrivacy -Impersonation Impersonate

        foreach ($Site in $Sites)
        {
            try
            {
                if($Site.GetState().returnvalue -eq 0)
                {
                    $State = "Starting"
                }
                if($Site.GetState().returnvalue -eq 1)
                {
                    $State = "Started"
                }
                if($Site.GetState().returnvalue -eq 2)
                {
                    $State = "Stopping"
                }
                if($Site.GetState().returnvalue -eq 3)
                {
                    $State = "Stopped"
                }
                if($Site.GetState().returnvalue -eq 4)
                {
                    $State = "Unknown"
                }
            }
            catch
            {
                #GetState method will always return exeption when called on FTP Website on IIS, it's  know bug. This if-else block is a workaround for that bug
                if($_.Exception.Message -eq 'Exception calling "GetState" : ""')
                {
                    $State = "Unknown / FTP Website"
                }
                else
                {
                    Write-Error $_.Exception.Message
                    $State = "Exception getting state!"
                }
            }

            $Properties = @{"Server Name" = $ComputerName
                            Name = $Site.Name
                            State = $State
                            Status = ""}

            $Obj = New-Object psobject -Property $Properties


            if ($Obj.State -eq "Exception getting state!")
            {
                $Obj.Status = "Error"
            }
            if ($Obj.State -eq "Stopped" -or $Obj.State -eq "Stopping")
            {
                $Obj.Status = "Error"
            }
            if ($Obj.State -eq "Starting" -or $Obj.State -eq "Unknown")
            {
                $Obj.Status = "Warning"
            }
            if ($Obj.State -eq "Started")
            {
                $Obj.Status = "OK"
            }
            if ($Obj.State -eq "Unknown / FTP Website")
            {
                $Obj.Status = "OK"
            }

            $ReportArray += $Obj
        }
    }
    End
    {
        $ReportArray
    }
}

function Get-IISApplicationPool
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true, 
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [string]$ComputerName
    )
    
    Begin
    {
        $ReportArray = @()

        if(!$ComputerName)
        {
            $ComputerName = $env:COMPUTERNAME
        }
    }
    Process
    {
        $AppPools = gwmi -namespace "root\webadministration" -Class applicationpool -ComputerName $ComputerName -Authentication PacketPrivacy -Impersonation Impersonate
 
        foreach($AppPool in $AppPools )
		{
            if($AppPool.GetState().returnvalue -eq 0)
            {
                $State = "Starting"
            }
            if($AppPool.GetState().returnvalue -eq 1)
            {
                $State = "Started"
            }
            if($AppPool.GetState().returnvalue -eq 2)
            {
                $State = "Stopping"
            }
            if($AppPool.GetState().returnvalue -eq 3)
            {
                $State = "Stopped"
            }
            if($AppPool.GetState().returnvalue -eq 4)
            {
                $State = "Unknown"
            }	

            $Properties = @{"Server Name" = $ComputerName
                            Name = $AppPool.Name
                            State = $State
                            Status = ""}

            $Obj = New-Object psobject -Property $Properties

            if ($Obj.State -eq "Stopped" -or $Obj.State -eq "Stopping")
            {
                $Obj.Status = "Error"
            }
            if ($Obj.State -eq "Starting" -or $Obj.State -eq "Unknown")
            {
                $Obj.Status = "Warning"
            }
            if ($Obj.State -eq "Started")
            {
                $Obj.Status = "OK"
            }

            $ReportArray += $Obj
        }
    }
    End
    {
        $ReportArray
    }
}
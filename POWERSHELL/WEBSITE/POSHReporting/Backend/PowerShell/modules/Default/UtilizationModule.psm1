function Get-CPUUtilization
{
    [CmdletBinding()]
    Param
        (
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]$ComputerName,

        [Parameter(Mandatory=$true,
                   Position=1)]
        [int32]$WarningLevel,

        [Parameter(Mandatory=$true,
                   Position=2)]
        [int32]$ErrorLevel
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
        try
        {
            Write-Verbose "Getting processor average utilization for $ComputerName."
            #Get current CPU utilization (average form all cores)
            $ProcessorUtilization = (Get-WmiObject win32_processor -ComputerName $ComputerName -ErrorAction Stop |  Measure-Object -property LoadPercentage -Average).Average
	        $ProcessorUtilization = [Math]::Round($ProcessorUtilization, 1)

            if ($ProcessorUtilization -gt $ErrorLevel)
            {
                $Status = "Error"
            }
            elseif ($ProcessorUtilization -gt $WarningLevel)
            {
                $Status = "Warning"
            }
            else
            {
                $Status = "OK"
            }

            $Properties = @{ "Server Name" = $ComputerName
                             "CPU Usage" = $ProcessorUtilization.ToString() + "%"
                             "Status" = $Status
                           }

            $ReportObj = New-Object psobject -Property $Properties

            Write-Verbose "Adding information from $ComputerName to array."

            $ReportArray += $ReportObj

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

function Get-RAMUtilization
{
    [CmdletBinding()]
    Param
        (
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]$ComputerName,

        [Parameter(Mandatory=$true,
                   Position=1)]
        [int32]$WarningLevel,

        [Parameter(Mandatory=$true,
                   Position=2)]
        [int32]$ErrorLevel
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
        try
        {
            #Get Ram info
            Write-Verbose "Getting ram utilization for $ComputerName."
            $OSinfo = Get-WmiObject -Class win32_operatingsystem  -ComputerName $ComputerName -ErrorAction Stop
            $RamTotal = $OSinfo.TotalVisibleMemorySize
            $RamFree = $OSinfo.FreePhysicalMemory
            $RamUsed = $RamTotal - $RamFree
            $RamPercentage = ([Math]::Round( ( $RamUsed * 100 / $RamTotal), 0 ))

            if($RamPercentage -gt $ErrorLevel)
            {
                $Status = "Error"
            }
            elseif($RamPercentage -gt $WarningLevel)
            {
                $Status = "Warning"
            }
            else
            {
                $Status = "OK"
            }

            #Create string with ram info
            $RamUtilization =  ([Math]::Round(($RamTotal / 1MB ), 2)).ToString() + " / " + ([Math]::Round(($RamUsed / 1MB ), 2)).ToString()

            $Properties = @{ "Server Name" = $ComputerName
                             "Memory (Gb)" = $RamUtilization
                             "Usage" = $RamPercentage.toString() + "%"
                             "Status" = $Status
                           }

            $ReportObj = New-Object psobject -Property $Properties

            Write-Verbose "Adding information from $ComputerName to array."
            $ReportArray += $ReportObj
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


function Get-DiskSpaceReport
{
    [CmdletBinding()]
    Param
        (
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]$ComputerName,

                [Parameter(Mandatory=$true,
                   Position=1)]
        [int32]$WarningLevel,

        [Parameter(Mandatory=$true,
                   Position=2)]
        [int32]$ErrorLevel
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
        try
        {
            Write-Verbose "Getting logical drives for $ComputerName."

            $Drives = Get-WmiObject -ComputerName $ComputerName Win32_LogicalDisk -ErrorAction Stop | Where-Object {$_.DriveType -eq 3}

            Foreach($Drive in $Drives)
            {
                $Used = ([math]::round((($Drive.Size - $Drive.FreeSpace) / 1GB), 2))
                $Total = ([math]::round(($Drive.Size / 1GB), 2))
                $UsagePercentage = ([math]::round(($Used / $Total) * 100, 2))

                if($UsagePercentage -gt $ErrorLevel)
                {
                    $Status = "Error"
                }
                elseif($UsagePercentage -gt $WarningLevel)
                {
                    $Status = "Warning"
                }
                else
                {
                    $Status = "OK"
                }

                $SpaceString = $Used.ToString() + " / " + $Total.ToString()

                $Properties = @{
                                "Server Name" = $ComputerName
                                "Drive" = $Drive.DeviceId
                                "Space (GB)" =  $SpaceString
                                Usage = $UsagePercentage.toString() + "%"
                                Status = $Status
                                }

                $DriveObj = New-Object psobject -Property $Properties

                $ReportArray += $DriveObj
            }

            Write-Verbose "Adding logical drives for $ComputerName to array."
            
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
function Get-SPServices
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
        [string[]]$Service
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
            foreach($ServiceName in $Service)
            {
                Write-Verbose "Checking '$Service' for '$ComputerName'."

                $StartUpMode = (Get-WmiObject -ComputerName $ComputerName -Class Win32_Service -Property StartMode -Filter "Name='$ServiceName'").StartMode
                $ServiceObj = Get-Service -Name $ServiceName -ComputerName $ComputerName
                
                $Properties = @{"Server Name" = $ComputerName
                                "Display Name" = $ServiceObj.DisplayName
                                Name = $ServiceObj.Name
                                State = $ServiceObj.Status
                                "Startup Type" = $StartUpMode
                                Status = ""}
                

                $Obj = New-Object psobject -Property $Properties

                Write-Verbose "Validating Status."

                if ($Obj.State -eq "Stopped" -and $Obj."Startup Type" -eq "Auto")
                {
                    $Obj.Status = "Error"
                }
                elseif ($Obj.State -eq "Running" -and $Obj."Startup Type" -eq "Manual")
                {
                    $Obj.Status = "Warning"
                }
                else
                {
                    $Obj.Status = "OK"
                }

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

function Get-EventLogEntries
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true, 
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true,
                    Position = 0)]
        [string]$ComputerName,

        [Parameter(Mandatory=$true,
                    Position = 1)]
        [string[]]$LogName,

        [Parameter(Mandatory=$true, 
                    Position = 2)]
        [int32]$Newest,

        [Parameter(Mandatory=$true, 
                    Position = 3)]
        [string[]]$EntryType
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
        foreach($Log in $LogName)
        {
            $Entries = Get-EventLog -ComputerName $ComputerName -LogName $Log -Newest $Newest -EntryType $EntryType

                foreach($Entry in $Entries )
                {
                    
                    $Properties = @{"Server Name" = $ComputerName
                                    "Log Name" = $Log
                                    Level = $Entry.EntryType
                                    Date = $Entry.TimeWritten
                                    Source = $Entry.Source
                                    "Event ID" = $Entry.EventID
                                    Message = ($Entry.Message | Repair-IllegalCharacters)}

                    $Obj = New-Object psobject -Property $Properties

                    $ReportArray += $Obj
                }
        }
       
    }
    End
    {
        $ReportArray
    }
}


#Helper Function

function Repair-IllegalCharacters
{
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true,
    ValueFromPipeline = $true,
    Position=0)]
    $String
  )

  process
  {
    
     # Match all characters that does NOT belong in an XML document
      $rPattern = "[^\x09\x0A\x0D\x20-\xD7FF\xE000-\xFFFD\x10000\x10FFFF]"

      # Replace said characters with [String]::Empty and return
      [System.Text.RegularExpressions.Regex]::Replace($String,$rPattern,"")
  }
     
}





function Get-PingTest
{
    [CmdletBinding()]
    Param
        (
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]$ComputerName
        )

    begin
    {
        $ReportArray = @()
    }
    process
    {
        try
        {
            $Ping = Test-Connection -ComputerName $ComputerName -ErrorAction Stop

            $Properties = @{"Server Name" = $ComputerName
                            IP = $Ping[0].IPV4Address
                            "Ping success" = $true
                            "Error Message" = "None"
                            Status = "OK"}
        }
        catch
        {
            Write-Error $_.Exception.Message

            $Properties = @{"Server Name" = $ComputerName
                            IP = ""
                            "Ping success" = $false
                            "Error Message" = $_.Exception.Message
                            Status = "Error"}
        }

    
        $Obj = New-Object psobject -Property $Properties

        $ReportArray += $Obj
    }
    end
    {
        $ReportArray 
    }
}
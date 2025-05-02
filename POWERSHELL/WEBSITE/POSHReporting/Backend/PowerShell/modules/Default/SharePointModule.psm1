function Get-HealthAnalyserEntries
{
    [CmdletBinding()]
    Param()

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
            $Entries = ([Microsoft.SharePoint.Administration.Health.SPHealthReportsList]::Local).Items
            
            foreach($Entry in $Entries)
            {
                $Properties = @{ "Category" = $Entry["Category"]
                                 "Title" = $Entry["Title"]
                                 "Failing Servers" = $Entry["Failing Servers"]
                                 "Failing Services" = $Entry["Failing Services"]
                                 "Date" = $Entry["Modified"]
                                 "Severity" = $Entry["Severity"]
                }

                $ReportObj = New-Object psobject -Property $Properties

                $ReportArray += $ReportObj
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

function Get-SearchCrawls
{
    [CmdletBinding()]
    Param
    ()
    
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
            $SSA = Get-SPEnterpriseSearchServiceApplication
            $ContentSources = Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $SSA
        
            foreach($ContentSource in $ContentSources)
            {
                $Properties = @{"Name" = $ContentSource.Name
                                "Type" = $ContentSource.Type
                                "Success Count" = $ContentSource.SuccessCount
                                "Warning Count" = $ContentSource.WarningCount
                                "Error Count" = $ContentSource.ErrorCount
                                "Crawl Started" = $ContentSource.CrawlStarted
                                "Crawl Completed" = $ContentSource.CrawlCompleted
                                "Crawl Duration" = if($ContentSource.CrawlCompleted -eq ""){""}else{$ContentSource.CrawlCompleted - $ContentSource.CrawlStarted}
                                "Crawl State" = $ContentSource.CrawlState
                                }

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

function Get-SPULSEvents
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline = $true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]$ComputerName,

        [Parameter(Mandatory=$true,
                   Position=1)]
        [String]$MinimumLevel,

        [Parameter(Mandatory=$true,
                   Position=2)]
        [System.DateTime]$StartTime
    )

    begin
    {
        $ReportArray = @()
    }

    Process
    {
        

        foreach($Server in $ComputerName)
        {
            try
            {
                $ReportArray += Invoke-Command -ComputerName $Server -ScriptBlock {Add-pssnapin Microsoft.SharePoint.PowerShell -erroraction 0; Get-SPLogEvent -MinimumLevel $args[0] -StartTime $args[1]} -ArgumentList $MinimumLevel, $StartTime | Select @{N="Server Name";E={$Server}}, *
            }
            catch
            {
                Write-Error $_.Exception.Message
            }
        }
    }
    End
    {
        $ReportArray
    }
}

#Helper Function
function ConvertFrom-IllegalCharacters
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
        $String.Replace('"', "&quot;").Replace("&", "&amp;").Replace("<", "&lt;").Replace("<", "&lt;")
    }
     
}

function Get-SPSolutionReport
{
    [CmdletBinding()]
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
            $Solutions = Get-SPSolution

            foreach($Solution in $Solutions)
            {
                if($Solution.LastOperationResult -eq "DeploymentFailedCallout")
                {
                    $Status = "Error"
                }
                elseif($Solution.LastOperationResult -ne "DeploymentSucceeded")
                {
                    $Status = "Warning"
                }
                else
                {
                    $Status = "OK"
                }

                $Properties = @{Name = $Solution.Name
                                ID = $solution.ID
                                Deployed = $Solution.Deployed
                                "Last Operation Result" = $Solution.LastOperationResult
                                "Last Operation End" = $solution.LastOperationEndTime
                                Status = $Status}

                $Obj = New-Object psobject -Property $Properties

                $ReportArray += $obj
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

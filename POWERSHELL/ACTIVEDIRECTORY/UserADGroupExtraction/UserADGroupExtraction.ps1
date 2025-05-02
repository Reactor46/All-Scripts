<#	
    .NOTES
    ===========================================================================
    Created with: 	ISE
    Created on:   	01/01/2019 1:46 PM
    Created by:   	Vikas Sukhija
    Organization: 	
    Filename:     	UserADGroupExtraction.ps1
    ===========================================================================
    .DESCRIPTION
    Extract User Groups in CSV Format, Delimiter for groups is semicolon
    Input form Users.txt file
#>

param (
    [string]$userlist
)
function Write-Log
{
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory = $true,ParameterSetName = 'Create')]
    [array]$Name,
    [Parameter(Mandatory = $true,ParameterSetName = 'Create')]
    [string]$Ext,
    [Parameter(Mandatory = $true,ParameterSetName = 'Create')]
    [string]$folder,
    
    [Parameter(ParameterSetName = 'Create',Position = 0)][switch]$Create,
    
    [Parameter(Mandatory = $true,ParameterSetName = 'Message')]
    [String]$Message,
    [Parameter(Mandatory = $true,ParameterSetName = 'Message')]
    [String]$path,
    [Parameter(Mandatory = $false,ParameterSetName = 'Message')]
    [ValidateSet('Information','Warning','Error')]
    [string]$Severity = 'Information',
    
    [Parameter(ParameterSetName = 'Message',Position = 0)][Switch]$MSG
  )
  switch ($PsCmdlet.ParameterSetName) {
    "Create"
    {
      $log = @()
      $date1 = Get-Date -Format d
      $date1 = $date1.ToString().Replace("/", "-")
      $time = Get-Date -Format t
	
      $time = $time.ToString().Replace(":", "-")
      $time = $time.ToString().Replace(" ", "")
	
      foreach ($n in $Name)
      {$log += (Get-Location).Path + "\" + $folder + "\" + $n + "_" + $date1 + "_" + $time + "_.$Ext"}
      return $log
    }
    "Message"
    {
      $date = Get-Date
      $concatmessage = "|$date" + "|   |" + $Message +"|  |" + "$Severity|"
      switch($Severity){
        "Information"{Write-Host -Object $concatmessage -ForegroundColor Green}
        "Warning"{Write-Host -Object $concatmessage -ForegroundColor Yellow}
        "Error"{Write-Host -Object $concatmessage -ForegroundColor Red}
      }
      
      Add-Content -Path $path -Value $concatmessage
    }
  }
}
function ProgressBar
{
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory = $true)]
    $Title,
    [Parameter(Mandatory = $true)]
    [int]$Timer
  )
	
  For ($i = 1; $i -le $Timer; $i++)
  {
    Start-Sleep -Seconds 1;
    Write-Progress -Activity $Title -Status "$i" -PercentComplete ($i /10 * 100)
  }
}

#################Check if logs folder is created##################
$logpath  = (Get-Location).path + "\logs" 
$testlogpath = Test-Path -Path $logpath
if($testlogpath -eq $false)
{
  ProgressBar -Title "Creating logs folder" -Timer 10
  New-Item -Path (Get-Location).path -Name Logs -Type directory
}

$reportpath = (Get-Location).path + "\Report"
$testReportpath = Test-Path -Path $reportpath
if($testReportpath -eq $false)
{
  ProgressBar -Title "Creating Report folder" -Timer 10
  New-Item -Path (Get-Location).path -Name Report -Type directory
}

##########################Load variables & Logs####################
$log = Write-Log -Name "log_UserGroupExtraction" -folder logs -Ext log
$Report = Write-Log -Name "Report_UserGroupExtraction" -folder Report -Ext csv
$collection = @()
try{
  Import-Module ActiveDirectory
  Write-log -Message "Script Started" -path $log -Severity Information
  Write-log -Message "AD Module Loaded" -path $log -Severity Information
}
catch{
  write-host "$($_.Exception.Message)" -foregroundcolor red
  Write-log -Message "$($_.Exception.Message)" -path $log -Severity Error
  break
}

$users = get-content $userlist
if($users.count -gt "0"){
  $users | ForEach-Object{
    $mcoll = "" | Select Name,memberof
    try{
      $getaduser = get-aduser -Identity $_ -Properties memberof
    
      $Name = $getaduser.Name
      $memberof = $getaduser.memberof
      Write-log -Message "Extracting groups from user $Name" -path $log -Severity Information
      $mcoll.Name = $Name
      $mcoll.memberof = $memberof
      $collection += $mcoll
    }
    catch{
      write-host "$($_.Exception.Message)" -foregroundcolor red
      Write-log -Message "$($_.Exception.Message)" -path $log -Severity Error
    }
   }
}

$collection | Select Name, @{n='memberof';e={$_.memberof -join ";"}} | Export-Csv $Report -NoTypeInformation
Write-log -Message "Script Finished" -path $log -Severity Information

#######################################################################################
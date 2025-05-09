﻿<#
.SYNOPSIS
   Gets the DAC Packages registered on target server
	
.DESCRIPTION
   Writes the registered Dac Packages out to the "21 - DacPackages" folder
      
.EXAMPLE
    21_Dac_Packages.ps1 localhost
	
.EXAMPLE
    21_Dac_Packages.ps1 server01 sa password

.Inputs
    ServerName\instance, [SQLUser], [SQLPassword]

.Outputs

	
.NOTES  
    DacFX Download
    https://www.microsoft.com/en-us/download/details.aspx?id=100297 - 18.3.1
    This Install loads the DLL into [C:\Program Files\Microsoft SQL Server\150]

    Check the Registrations results here:
    select * from msdb.dbo.sysdac_instances ORDER BY instance_name

    February 2020
    The DacFX code has branched off from SSMS/SSDT and has its own dev cadence
    
    Try installing the NuGet Package in an Elevated Console
    Install-Package Microsoft.SqlServer.DacFx.x86 -ProviderName NuGet

.LINK
	https://github.com/gwalkey
	
	
#>

[CmdletBinding()]
Param(
  [string]$SQLInstance="localhost",
  [string]$myuser,
  [string]$mypass,
  [int]$registerDAC=1,
  [int]$ExportBacPac=1
)

# Load Common Modules and .NET Assemblies
try
{
    Import-Module ".\SQLTranscriptase.psm1" -ErrorAction Stop
}
catch
{
    Throw('SQLTranscriptase.psm1 not found')
}

LoadSQLSMO
LoadDacFx

# Init
Set-StrictMode -Version latest;
[string]$BaseFolder = (Get-Item -Path ".\" -Verbose).FullName
Write-Host  -f Yellow -b Black "21 - DAC Packages"
Write-Output("Server: [{0}]" -f $SQLInstance)

# Server connection check
$SQLCMD1 = "select serverproperty('productversion') as 'Version'"
try
{
    if ($mypass.Length -ge 1 -and $myuser.Length -ge 1) 
    {
        Write-Output "Testing SQL Auth"        
        $myver = ConnectSQLAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $SQLCMD1 -User $myuser -Password $mypass -ErrorAction Stop| select -ExpandProperty Version
        $serverauth="sql"
    }
    else
    {
        Write-Output "Testing Windows Auth"
		$myver = ConnectWinAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $SQLCMD1 -ErrorAction Stop | select -ExpandProperty Version
        $serverauth = "win"
    }

    if($myver -ne $null)
    {
        Write-Output ("SQL Version: {0}" -f $myver)
    }

}
catch
{
    Write-Host -f red "$SQLInstance appears offline."
    Set-Location $BaseFolder
	exit
}

[int]$ver = GetSQLNumericalVersion $myver

# Skip if server not 2008 R2+
if ($ver -lt 10)
{
    Write-Output "Dac Packages only supported on SQL Server 2008 or higher"
    set-location $BaseFolder
    exit
}


# New UP SQL SMO Object
if ($serverauth -eq "win")
{
    $srv = New-Object "Microsoft.SqlServer.Management.SMO.Server" $SQLInstance
}
else
{
    $srv = New-Object "Microsoft.SqlServer.Management.SMO.Server" $SQLInstance
    $srv.ConnectionContext.LoginSecure=$false
    $srv.ConnectionContext.set_Login($myuser)
    $srv.ConnectionContext.set_Password($mypass)
}


# Create Output Folder
$Output_path  = "$BaseFolder\$SQLInstance\21 - DAC Packages\"
if(!(test-path -path $Output_path))
{
    mkdir $Output_path | Out-Null
}

# Drift Reports
$DriftOutput_path  = "$BaseFolder\$SQLInstance\21 - DAC Packages\DriftReports\"
if(!(test-path -path $DriftOutput_path))
{
    mkdir $DriftOutput_path | Out-Null
}


# Check for existence of SqlPackage.exe and get latest version
Write-Output("Check for existence of SqlPackage.exe and get latest version")

$pkgver = $null;

$pkgexe = "C:\Program Files (x86)\Microsoft SQL Server\100\DAC\bin\sqlpackage.exe"
if((test-path -path $pkgexe))
{
    $pkgver = $pkgexe
    Write-output('SQLPackage v2008 found')
}

$pkgexe = "C:\Program Files (x86)\Microsoft SQL Server\110\DAC\bin\sqlpackage.exe"
if((test-path -path $pkgexe))
{
    $pkgver = $pkgexe
    Write-output('SQLPackage v2012 found')
}

$pkgexe = "C:\Program Files (x86)\Microsoft SQL Server\120\DAC\bin\sqlpackage.exe"
if((test-path -path $pkgexe))
{
    $pkgver = $pkgexe
    Write-output('SQLPackage v2014 found')
}

$pkgexe = "C:\Program Files (x86)\Microsoft SQL Server\130\DAC\bin\sqlpackage.exe"
if((test-path -path $pkgexe))
{
    $pkgver = $pkgexe
    Write-output('SQLPackage v2016 found')
}

$pkgexe = "C:\Program Files (x86)\Microsoft SQL Server\140\DAC\bin\sqlpackage.exe"
if((test-path -path $pkgexe))
{
    $pkgver = $pkgexe
    Write-output('SQLPackage v2017 found')
}

$pkgexe = "C:\Program Files\Microsoft SQL Server\150\DAC\bin\sqlpackage.exe"
if((test-path -path $pkgexe))
{
    $pkgver = $pkgexe
    Write-output('SQLPackage v2019 found')
}

If (!($pkgver))
{
    Write-Output "SQLPackage.exe not found, exiting"
    exit
}

Write-Output("`r`nCreating Dac Export Batch file")

# Create Batch file to run below
$myoutstring = "@ECHO OFF`n" | out-file -FilePath "$Output_path\DacExtract.cmd" -Force -Encoding ascii

foreach($sqlDatabase in $srv.databases)
{
    Write-Output("Database: [{0}] " -f $sqlDatabase.Name)

    # Skip System Databases
    if ($sqlDatabase.Name -in 'Master','Model','MSDB','TempDB','SSISDB','ReportServer','ReportServerTempDB') {continue}

    # Strip brackets from DBname
    $db = $sqlDatabase
    $myDB = $db.name
    $myServer = $SQLInstance   
    $fixedDBName = $db.name.replace('[','')
    $fixedDBName = $fixedDBName.replace(']','')

    # Skip Offline Databases (SMO still enumerates them, but cant retrieve the objects)
    if ($sqlDatabase.Status -ne 'Normal')     
    {
        Write-Output ("Skipping Offline: {0}" -f $sqlDatabase.Name)
        continue
    }

    # One Output folder per DB
    if(!(test-path -path $output_path))
    {
        mkdir $output_path | Out-Null
    }

    set-location $Output_path
       
    # ------------------
    # Script out DACPACs
    # ------------------
    if ($serverauth -eq "win")
    {
        $myoutstring = [char]34+$pkgver + [char]34+ " /action:extract /sourcedatabasename:$myDB /sourceservername:$MyServer /targetfile:$MyDB.dacpac `n"
    }
    else
    {
        $myoutstring = [char]34+$pkgver + [char]34+ " /action:extract /sourcedatabasename:$myDB /sourceservername:$MyServer /targetfile:$MyDB.dacpac /sourceuser:$myuser /sourcepassword:$mypass `n"
    }
    $myoutstring | out-file -FilePath "$Output_path\DacExtract.cmd" -Encoding ascii -append

    # 
    # Register the Database as a Data Tier Application
    #
    if ($registerDAC -eq 1)
    {
        # Specify the DAC metadata before Registration
        $applicationname = $fixedDBName
        [system.version]$version = '1.0.0.0'
        $description = "Registered during DacPac Script-Out pass on "+(Get-Date).ToString()
        # Register as 1.0.0.0    
        try
        {
            if ($serverauth -eq "win")
            {
                $dac = new-object Microsoft.SqlServer.Dac.DacServices "data source=$sqlinstance;Integrated Security=SSPI;Application Name=SQLTranscriptase"
            }
            else
            {
                $dac = new-object Microsoft.SqlServer.Dac.DacServices "data source=$sqlinstance;User ID=$myUser;Password=$myPass;Application Name=SQLTranscriptase"
            }
            $dac.register($myDB, $myDB, $version, $description)
			Write-Output ("Registered Database [{0}]  v[{1}] as [{2}]" -f $myDB, $version, $description)
        }
        catch
        {
            Write-Output('Dac Register of [{0}] failed, Error:[{1}]' -f $mydb, $error[0])
        }
        $dac = $null;
    }
    
    # -------------------------------
    # Create Drift Report batch file
    # -------------------------------
    $myDriftFileName = $DriftOutput_path+"\"+$myDB+"_DriftReport.cmd"
    $myDriftReportName = $myDB+"_DriftReport.xml"

    # SQLPackage.EXE needs DMZ username and password parameters passed in
    if ($serverauth -eq "win")
    {
        [char]34 + $pkgver + [char]34 + " /A:DriftReport /tsn:$myServer /tdn:$myDB /op:$myDriftReportName `n $myDriftReportName `n" | out-file -FilePath $myDriftFileName -Force -Encoding ascii
    }
    else
    {
        [char]34 + $pkgver + [char]34 + " /A:DriftReport /tsn:$myServer /tdn:$myDB /tu:$myuser /tp:$mypass /op:$myDriftReportName `n $myDriftReportName `n" | out-file -FilePath $myDriftFileName -Force -Encoding ascii
    }

    # ---------------------
    # Script out BACPACs
    # ---------------------
    if ($serverauth -eq "win")
    {
        $myoutstring = [char]34+$pkgver + [char]34+ " /action:export /sourcedatabasename:$myDB /sourceservername:$MyServer /targetfile:$MyDB.bacpac `n"
    }
    else
    {
        $myoutstring = [char]34+$pkgver + [char]34+ " /action:export /sourcedatabasename:$myDB /sourceservername:$MyServer /targetfile:$MyDB.bacpac /sourceuser:$myuser /sourcepassword:$mypass `n"
    }
    $myoutstring | out-file -FilePath "$Output_path\BacExport.cmd" -Encoding ascii -append

}

# Run the SQLPACKAGE batch files
Write-Output("`r`nExporting Dac Packages")
invoke-expression ".\DacExtract.cmd"

if ($ExportBacPac -eq 1)
{
    invoke-expression ".\BacExport.cmd"
}

# Remember to run the Drift Report batch files in the DriftReports folder
remove-item -Path "$Output_path\DacExtract.cmd" -Force -ErrorAction SilentlyContinue
remove-item -Path "$Output_path\BacExport.cmd" -Force -ErrorAction SilentlyContinue

# Return to Base
set-location $BaseFolder




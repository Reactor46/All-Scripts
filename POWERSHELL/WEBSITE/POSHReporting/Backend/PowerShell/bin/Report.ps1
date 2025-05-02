#Get enviromental data
if ($Host.Version.Major -gt 2)
{
    $ProjectRoot = $PSScriptRoot | Split-Path
}
else
{
    $ProjectRoot = Split-Path $MyInvocation.MyCommand.Path -Parent | Split-Path
}

Start-Transcript $ProjectRoot\logs\log.txt

#Load Assemblies
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction 0
[Reflection.Assembly]::LoadWithPartialName('System.Xml.Linq') | Out-Null
Add-Type -AssemblyName System.Web
Import-Module $ProjectRoot\bin\ReportModule.psm1


#Get application configuration
$LogPath = "$ProjectRoot\logs\log.txt"
$Modules = Get-ChildItem $ProjectRoot\modules -Recurse | Where-Object {$_.Name -like "*.psm1"}
$HTMLTemplate = Get-Content $ProjectRoot\templates\ReportTemplate.html | Out-String
[XML]$XMLMainConfig = Get-Content $ProjectRoot\config\config.xml | Out-String
$ReportConifg = $XMLMainConfig.Configuration

#Import Framework functions
$Modules | ForEach-Object { Write-Host "Importing module $($_.FullName)."; Import-Module $_.FullName }

#Deserialize
$General = $ReportConifg.General
$Servers = $ReportConifg.Servers | select -ExpandProperty Server | Select Type, @{N="Name";E={$_.'#text'}}
$Scripts = $ReportConifg.Scripts | select -ExpandProperty script | select @{N="Title";E={$_.'#text'}}, Path, Enabled
$Other = $ReportConifg.Other
$Email = $ReportConifg.Email

#Quick Variables
$ComputerName = $Servers | Select -expand "Name" 
$FileName = $General.FileName + ".html"
$ReportPath = "$ProjectRoot\reports\$FileName"
$ReportName = $General.ReportName


function Start-SPReport
{
    [CmdletBinding()]
    param()

    #Sets Verbose interherence
    if (-not $PSBoundParameters.ContainsKey('Verbose'))
    {
        $VerbosePreference = $PSCmdlet.GetVariableValue('VerbosePreference')
    }

    cd $ProjectRoot

    # Create HTML Tables from powershell scripts
    $HTMLTables = @()

    foreach($Script in $Scripts)
    {
        if($Script.Enabled -eq $true)
        {
            $ScritpPath = $Script.Path | Convert-Path

            Write-Host "Executing $ScritpPath"

            $ScriptObject = $null

            $ScriptObject = & $ScritpPath

            if($ScriptObject -ne $null)
            {
                $Properties = @{ InputObject = $ScriptObject
                                 TableID = $Script.Title.ToLower().replace(" ","_")
                                 PreContent = "<h2>$($Script.Title)</h2>"}

                $HTMLTable = ConvertTo-CustomHTMLTable @Properties
                $HTMLTables += $HTMLTable
            }
        }
    }

    New-SPReport -HTMLTemplate $HTMLTemplate -HTMLTables $HTMLTables -ReportName $ReportName -Path $ReportPath

    if($Other.Enabled -eq $true)
    {
        & ($Other.Script.Path | Convert-Path)
    }

    if($Email.Enabled -eq $true)
    {
        if($Host.Version.Major -gt 2)
        {
            $MailProperties = $Email | Convert-XMLToHashTable
            $MailProperties.Add("Attachment","$ReportPath")
            $MailProperties.Add("BodyAsHtml", $null)

            Write-Host "Sending mail to with following properties:"

            $MailProperties

            Send-MailMessage @MailProperties
        }
        else
        {
            $MailProperties = $Email | Convert-XMLToHashTable
            $MailProperties.Add("Attachment","$ReportPath")

            Write-Host "Sending mail to with following properties:"

            $MailProperties

            Send-MailMessage2.0 @MailProperties
        }
    }
}


Start-SPReport

Stop-Transcript -ErrorAction SilentlyContinue
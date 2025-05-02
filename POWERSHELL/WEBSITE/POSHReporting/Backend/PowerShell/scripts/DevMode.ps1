function Enable-DevMode
{
    [CmdletBinding()]

    param(
            
            [Parameter(
            Mandatory = $true,
            Position = 0)]
            [String][ValidateNotNullOrEmpty()]$ProjectRoot
         )

    begin
    {
        #Setting Verbose interherence
        if (-not $PSBoundParameters.ContainsKey('Verbose'))
        {
            $VerbosePreference = $PSCmdlet.GetVariableValue('VerbosePreference')
        }
    }

    process
    {
        #Load Assemblies
        Write-Verbose "Loading Microsoft.SharePoint.PowerShell snappin." 
        Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction 0

        Write-Verbose "Loading 'System.Xml.Linq' assembly."
        [Reflection.Assembly]::LoadWithPartialName('System.Xml.Linq') | Out-Null

        Write-Verbose "Loading 'System.Web' assembly."
        Add-Type -AssemblyName System.Web

        Import-Module $ProjectRoot\bin\ReportModule.psm1

        #Global variables
        $VariableDescription =  "PowerShell Reporting Framework Variable"

        Set-Variable -Scope Global -Name ProjectRoot -Value $ProjectRoot -Description $VariableDescription

        Set-Variable -Scope Global -Name LogPath -Value "$ProjectRoot\logs\log.txt" -Description $VariableDescription

        Set-Variable -Scope Global -Name Modules -Value (Get-ChildItem $ProjectRoot\modules -Recurse | Where-Object {$_.Name -like "*.psm1"}) -Description $VariableDescription

        Set-Variable -Scope Global -Name XMLMainConfig -Value ([XML](Get-Content $ProjectRoot\config\config.xml | Out-String)) -Description $VariableDescription

        Set-Variable -Scope Global -Name ReportConifg -Value ($XMLMainConfig.Configuration) -Description $VariableDescription

        #Import framework functions
        $Modules | Select -ExpandProperty FullName | Import-Module

        #Deserialize

        Set-Variable -Scope Global -Name General -Value ($ReportConifg.General) -Description $VariableDescription

        Set-Variable -Scope Global -Name Servers -Value ($ReportConifg.Servers | select -ExpandProperty Server -ErrorAction 0 | Select Type, @{N="Name";E={$_.'#text'}}) -Description $VariableDescription
 
        Set-Variable -Scope Global -Name Scripts -Value ($ReportConifg.Scripts | select -ExpandProperty Script -ErrorAction 0 | Select @{N="Title";E={$_.'#text'}}, Path, Enabled) -Description $VariableDescription

        Set-Variable -Scope Global -Name Other -Value ($ReportConifg.Other) -Description $VariableDescription

        Set-Variable -Scope Global -Name Email -Value ($ReportConifg.Email) -Description $VariableDescription

        #Quick Variables
        Set-Variable -Scope Global -Name ComputerName -Value ( $Servers | Select -expand "Name") -Description $VariableDescription

        Set-Variable -Scope Global -Name FileName -Value ($General.FileName + ".html") -Description $VariableDescription

        Set-Variable -Scope Global -Name ReportPath -Value ("$ProjectRoot\reports\$FileName") -Description $VariableDescription

        Set-Variable -Scope Global -Name ReportName -Value ($General.ReportName) -Description $VariableDescription
        
        Write-Host "Following Cmdlets are now avaliable:" -ForegroundColor Yellow

        $global:Cmdlets = @()

        $global:Cmdlets += Get-command -Module ReportModule

        $global:Cmdlets += ($Modules | Select -expand BaseName) | ForEach-Object { Get-Command -Module $_}

        $Cmdlets

        Write-Host ""
        Write-Host "Following variables are now avaliable:" -ForegroundColor Yellow

        $global:Variables = Get-Variable | Where-Object {$_.Description -eq $VariableDescription} |  Format-Table Name, @{N="Type";E={$_.value.gettype().Fullname}}, Value -AutoSize
        $Variables | Out-String
        
        Write-Host 'You can check again list of avaliable functions by typeing "$Cmdlets" in console anytime.' -ForegroundColor Green
        Write-Host 'You can check again list of avaliable variables by typeing "$Variables" in console anytime.' -ForegroundColor Green
    }
    

}



#Get enviromental data
if ($Host.Version.Major -gt 2)
{
    $ProjectRoot = $PSScriptRoot | Split-Path
}
else
{
    $ProjectRoot = Split-Path $MyInvocation.MyCommand.Path -Parent | Split-Path
}

Enable-DevMode -ProjectRoot $ProjectRoot
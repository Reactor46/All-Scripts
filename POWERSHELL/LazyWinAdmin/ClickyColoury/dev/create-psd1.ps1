# api: ps
# type: cli
# title: create .psd1
# description: module manifest from script meta header
# version: 0.1
#
# Creates a .psd1 file for a .psm1/.ps1
# Contains an even cruder PMD extraction scheme.

[CmdletBinding()]
Param($fn=(ReadHost ".psm1 filename:"));

$target = $fn -replace "\.\w+$",".psd1"
$base = $fn -replace "^.+[\\\\/]",""

# plugin meta block
$src = (Get-Content $fn | Out-String)
$meta = @{}
$src | Select-String "(?m)^#\s*(\w+):\s*([^\n]+)" -AllMatches |
     % { $_.Matches } | % { ,($_.Groups | % { $_.Value }) } |
     % { $meta[$_[1]] = $_[2] }

# create .psd1 file
$psd = @{
    Path = $target
    RootModule = $base
    ModuleToProcess = $base
    Author = $meta.author
    CompanyName = "-"
    Description = $meta.description
    ModuleVersion = $meta.version
    #ProjectUri = "$meta.url"
    Guid = [guid]::newguid()
    PowerShellVersion = "2.0"
    #FormatsToProcess = "$fn.ps1xml"
    ProcessorArchitecture = 'amd64'
    FunctionsToExport = "*"
    AliasesToExport = "*"
    VariablesToExport = ""
    CmdletsToExport = ""
    PassThru = $true
}
New-ModuleManifest @psd | Out-File $target -Encoding UTF8

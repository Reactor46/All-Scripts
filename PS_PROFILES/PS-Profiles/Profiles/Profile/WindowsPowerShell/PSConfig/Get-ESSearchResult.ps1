# source: https://github.com/DBremen/PowerShellScripts
filter Get-FileSize {
	"{0:N2} {1}" -f $(
	if ($_ -lt 1kb) { $_, 'Bytes' }
	elseif ($_ -lt 1mb) { ($_/1kb), 'KB' }
	elseif ($_ -lt 1gb) { ($_/1mb), 'MB' }
	elseif ($_ -lt 1tb) { ($_/1gb), 'GB' }
	elseif ($_ -lt 1pb) { ($_/1tb), 'TB' }
	else { ($_/1pb), 'PB' }
	)
}

function Get-ESItemProperty {
    Param
        (
         $record
        )
    if ((Test-Path $record -ErrorAction SilentlyContinue) -eq $true) {
        return $record | Get-ItemProperty;
    }
    $result = New-Object System.Object
    $result | Add-Member -type NoteProperty -name Name -value (Split-Path "$record" -leaf)
    $result | Add-Member -type NoteProperty -name DirectoryName -value (Split-Path "$record")
    $result | Add-Member -type NoteProperty -name LastWriteTime -value "?"
    return $result
}

function Get-ESSearchResult {
    [CmdletBinding()]
    Param
    (
        #searchterm
        [Parameter(Mandatory=$true, Position=0)]
        $SearchTerm,
        #openitem
        [switch]$OpenItem,
        [switch]$CopyFullPath,
        [switch]$OpenFolder,
        [switch]$AsObject
    )
    $esPath = 'C:\Program Files*\es\es.exe'
    if (!(Test-Path (Resolve-Path $esPath).Path)){
        Write-Warning "Everything commandline es.exe could not be found on the system please download and install via http://www.voidtools.com/es.zip"
        exit
    }
	$result = & (Resolve-Path $esPath).Path $SearchTerm
    if($result.Count -gt 1){
        #$result = $result | Out-GridView -PassThru
    }
    foreach($record in $result){
        switch ($PSBoundParameters){
	        { $_.ContainsKey("CopyFullPath") } { $record | clip }
	        { $_.ContainsKey("OpenItem") }     { if (Test-Path $record -PathType Leaf) {  & "$record" } }
	        { $_.ContainsKey("OpenFolder") }   {  & "explorer.exe" /select,"$(Split-Path $record)" }
	        { $_.ContainsKey("AsObject") }     { Get-ESItemProperty($record) }
	        default                            { Get-ESItemProperty($record) |
                                                    select Name,DirectoryName,@{Name="Size";Expression={$_.Length | Get-FileSize }},LastWriteTime
                                               }
        }
    }
}

Set-Alias -Name es -Value Get-ESSearchResult -Description "Find files using Voidtool's everything"

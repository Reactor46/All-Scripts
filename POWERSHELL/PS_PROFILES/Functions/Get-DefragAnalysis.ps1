#requires -version 2.0


Function Get-DefragAnalysis {

<#
.Synopsis
Run a defrag analysis.
.Description
This command uses WMI to to run a defrag analysis on selected volumes on
local or remote computers. You will get a custom object for each volume like
this:

AverageFileSize               : 64
AverageFragmentsPerFile       : 1
AverageFreeSpacePerExtent     : 17002496
ClusterSize                   : 4096
ExcessFolderFragments         : 0
FilePercentFragmentation      : 0
FragmentedFolders             : 0
FreeSpace                     : 161816576
FreeSpacePercent              : 77
FreeSpacePercentFragmentation : 29
LargestFreeSpaceExtent        : 113500160
MFTPercentInUse               : 100
MFTRecordCount                : 511
PageFileSize                  : 0
TotalExcessFragments          : 0
TotalFiles                    : 182
TotalFolders                  : 11
TotalFragmentedFiles          : 0
TotalFreeSpaceExtents         : 8
TotalMFTFragments             : 1
TotalMFTSize                  : 524288
TotalPageFileFragments        : 0
TotalPercentFragmentation     : 0
TotalUnmovableFiles           : 4
UsedSpace                     : 47894528
VolumeName                    : 
VolumeSize                    : 209711104
Driveletter                   : E:
DefragRecommended             : False
Computername                  : NOVO8

The default drive is C: on the local computer.
.Example
PS C:\> Get-DefragAnalysis
Run a defrag analysis on C: on the local computer
.Example
PS C:\> Get-DefragAnalysis -drive "C:" -computername $servers
Run a defrag analysis for drive C: on a previously defined collection of server names.
.Example
PS C:\> $data = Get-WmiObject Win32_volume -filter "driveletter like '%' AND drivetype=3" -ComputerName Novo8 | Get-DefragAnalysis
PS C:\> $data | Sort Driveletter | Select Computername,DriveLetter,DefragRecommended

Computername                    Driveletter                     DefragRecommended
------------                    -----------                     -----------------
NOVO8                           C:                                          False
NOVO8                           D:                                          True
NOVO8                           E:                                          False

Get all volumes on a remote computer that are fixed but have a drive letter,
this should eliminate CD/DVD drives, and run a defrag analysis on each one.
The results are saved to a variable, $data.
.Notes
Last Updated: 12/5/2012
Author      : Jeffery Hicks (http://jdhitsolutions.com/blog)
Version     : 0.9

.Link
Get-WMIObject
Invoke-WMIMethod

#>

[cmdletbinding(SupportsShouldProcess=$True)]

Param(
[Parameter(Position=0,ValueFromPipelineByPropertyName=$True)]
[ValidateNotNullorEmpty()]
[Alias("drive")]
[string]$Driveletter="C:",
[Parameter(Position=1,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
[ValidateNotNullorEmpty()]
[Alias("PSComputername","SystemName")]
[string[]]$Computername=$env:computername
)

Begin {
    Write-Verbose -Message "$(Get-Date) Starting $($MyInvocation.Mycommand)"   
} #close Begin

Process {
    #strip off any extra spaces on the drive letter just in case
    Write-Verbose "$(Get-Date) Processing $Driveletter"
    $Driveletter=$Driveletter.Trim()
    if ($Driveletter.length -gt 2) {
        Write-Verbose "$(Get-Date) Scrubbing drive parameter value"
        $Driveletter=$Driveletter.Substring(0,2)
    }
    #add a colon if not included
    if ($Driveletter -match "^\w$") {
        Write-Verbose "$(Get-Date) Modifying drive parameter value"
        $Driveletter="$($Driveletter):"
    }

    Write-Verbose "$(Get-Date) Analyzing drive $Driveletter"
        
    Foreach ($computer in $computername) {
        Write-Verbose "$(Get-Date) Examining $computer"
        Try {
            $volume=Get-WmiObject -Class Win32_Volume -filter "DriveLetter='$Driveletter'" -computername $computer -errorAction "Stop"
        }
        Catch {
            Write-Warning ("Failed to get volume {0} from  {1}. {2}" -f $driveletter,$computer,$_.Exception.Message)
        }
        if ($volume) {
            Write-Verbose "$(Get-Date) Running defrag analysis"
            $analysis = $volume | Invoke-WMIMethod -name DefragAnalysis
        
            #get properties for DefragAnalysis so we can filter out system properties
            $analysis.DefragAnalysis.Properties | 
            Foreach -begin {$Prop=@()} -process { $Prop+=$_.Name }
        
            Write-Verbose "$(Get-Date) Retrieving results"
            $analysis | Select @{Name="Results";Expression={$_.DefragAnalysis | 
            Select-Object -Property $Prop |
            Foreach-Object { 
              #Add on some additional property values
              $_ | Add-member -MemberType Noteproperty -Name Driveletter -value $DriveLetter
              $_ | Add-member -MemberType Noteproperty -Name DefragRecommended -value $analysis.DefragRecommended 
              $_ | Add-member -MemberType Noteproperty -Name Computername -value $volume.__SERVER -passthru
             } #foreach-object
            }}  | Select -expand Results 
            
            #clean up variables so there are no accidental leftovers
            Remove-Variable "volume","analysis"
        } #close if volume
     } #close Foreach computer
 } #close Process
 
End {
    Write-Verbose "$(Get-Date) Defrag analysis complete"
} #close End

} #close function
﻿#Requires -Version 2.0

##########################################################################################################################################
# Data Sections
##########################################################################################################################################
$Script:Html_Header = Data {
@'
<!doctype html><html><head><title>{0}</title>
<style type="text/css">
h1, h2, h3, h4, h5, h6, p, a, ul, li, ol, td, label, input, span, div {{font-weight:normal !important; font-family:Tahoma, Arial, Helvetica, sans-serif;}}
.nochange {{ color: gray; font-style: italic;}}
.computeradded {{color: green; font-weight: bold;}}
.computerremoved {{color: red; font-weight: bold;}}
.shareremoved {{color: red; font-weight: bold;}}
.shareadded {{color: green; font-weight: bold;}}
.shareremoved {{color: red; font-weight: bold;}}
.permissionremoved {{color: red; font-weight: bold;}}
.permissionchanged {{color: orange; font-weight: bold;}}
.permissionadded {{ color: green; font-weight: bold;}}
.typelabel {{ color: gray; font-weight: heavy;}}
.h2 {{ color: black; font-weight: heavy;}}
.h3 {{ color: black; font-weight: heavy; font-style: italic;}}
</style>
</head>
<body>
<h1>{0}</h1>
'@
}

$Script:Html_Footer = Data {
@'
</body></html>
'@
}

$Script:Html_ComputerName = Data {
@'
<h2>Differences on Computer {0}</h2>
'@
}

$Script:Html_ShareName = Data {
@'
<h3>Differences in Share {0}</h3>
'@
}

$Script:Html_DifferenceLine = Data {
@'
<span class='typelabel'>{0}:&nbsp;</span><span class='{1}'>{2}</span><br>
'@
}

##########################################################################################################################################
# Main CmdLets
##########################################################################################################################################


<#
.SYNOPSIS
	Creates a list of Share, File and Folder ACLs for the specified shares/computers.

.DESCRIPTION 
	Produces an array of [ACLReportTools.Permission] objects for the computers provided. Specific shares can be specified or excluded using the Include/Exclude parameters.

	The report can be stored for use as a comparison in either a variable or as a file using the Export-ACLReport cmdlet (found in this module). For example:

	New-ACLShareReport -ComputerName CLIENT01 -Include MyShare,OtherShare | Export-ACLReport -path c:\ACLReports\CLIENT01_2014_11_14.acl
     
.PARAMETER ComputerName
	This is the computer(s) to create the ACL Share report for. The Computer names can also be passed in via the pipeline.

.PARAMETER Include
	This is a list of shares to include from the report. If this parameter is not set it will default to including all shares. This parameter can't be set if the Exclude parameter is set.

.PARAMETER Exclude
	This is a list of shares to exclude from the report. If this parameter is not set it will default to excluding no shares. This parameter can't be set if the Include parameter is set.

.PARAMETER IncludeInherited
	Setting this switch will cause the non inherited file/folder ACLs to be pulled recursively.

.EXAMPLE 
	New-ACLShareReport -ComputerName CLIENT01
	Creates a report of all the Share and file/folder ACLs on the CLIENT01 machine.

.EXAMPLE 
	New-ACLShareReport -ComputerName CLIENT01 -Include MyShare,OtherShare
	Creates a report of all the Share and file/folder ACLs on the CLIENT01 machine that are in shares named either MyShare or OtherShare.

.EXAMPLE 
	New-ACLShareReport -ComputerName CLIENT01 -Exclude SysVol
	Creates a report of all the Share and file/folder ACLs on the CLIENT01 machine that are in shares not named SysVol.
#>
Function New-ACLShareReport
{
    [CmdLetBinding()]
    param (
        [Parameter(
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)]
        [String[]]$ComputerName=$env:computername,

        [String[]]$Include,

        [String[]]$Exclude,
        
        [Switch]$IncludeInherited
    ) # param
    begin
    {
        [ACLReportTools.Permission[]]$acls = $null
		$null = $PSBoundParameters.Remove('includeinherited')
    } # Begin
    process
    {
        $Shares = Get-ACLShare @PSBoundParameters
        $acls += $Shares | Get-ACLShareACL
        if ($IncludeInherited)
        {
            $acls += $Shares | Get-ACLShareFileACL -Recurse
        }
        else
        {
            $acls += $Shares | Get-ACLShareFileACL -Recurse -IncludeInherited            
        }
    } # Process
    end
    {
        return $acls
    } # End
} # Function New-ACLShareReport


<#
.SYNOPSIS
	Creates a list of File and Folder ACLs for the provided path(s).

.DESCRIPTION 
	Produces an array of [ACLReportTools.Permission] objects for the list of paths provided.

	The report can be stored for use as a comparison in either a variable or as a file using the Export-ACLReport cmdlet (found in this module). For example:

	New-ACLPathFileReport -Path e:\public | Export-ACLReport -path c:\ACLReports\Public_2015-04-04.acl
     
.PARAMETER Path
	This is the path(s) to create the ACL PathFile report for.

.PARAMETER IncludeInherited
	Setting this switch will cause the non inherited file/folder ACLs to be pulled recursively.

.EXAMPLE 
	New-ACLPathFileReport -Path e:\public
	Creates a report of all the file/folder ACLs in the e:\public folder on this machine.
#>
Function New-ACLPathFileReport
{
    [CmdLetBinding()]
    param
    (
        [Parameter(
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)]
        [String[]]$Path=(Convert-path -Path .),
        
        [Switch]$IncludeInherited
    ) # param
    begin
    {
        [ACLReportTools.Permission[]]$acls = $null
		$null = $PSBoundParameters.Remove('path')
    } # Begin
    process
    {
        Foreach ($p in $Path) {
            $acls += Get-ACLPathFileACL -Path $p -Recurse @PSBoundParameters
		}
    } # Process
    end
    {
        return $acls
    } # End
} # Function New-ACLPathFileReport


<#
.SYNOPSIS
	Export an ACL Report as a file.

.DESCRIPTION 
	This Cmdlet will save whatever ACL Report that is in the pipeline to a file.

	This cmdlet just calls Export-ACLPermission although at some point will add additional functionality.
     
.PARAMETER Path
	This is the path to the ACL Permission Report output file. This parameter is required.

.PARAMETER InputObject
	Specifies the Permissions objects to export to the file. Enter a variable that contains the objects or type a command or expression that gets the objects. You can also pipe ACLReportTools.Permission objects to this cmdlet.

.PARAMETER Force
	Causes the file to be overwritten if it exists.

.EXAMPLE 
	 New-ACLShareReport -ComputerName CLIENT01 -Include MyShare,OtherShare | Export-ACLReport -path c:\ACLReports\CLIENT01_2014_11_14.acl
	 Creates a new ACL Share Report for Computer Client01 for the MyShare and OtherShares and exports it to the file C:\ACLReports\CLIENT01_2014_11_14.acl.

.EXAMPLE 
	 Export-ACLReport -Path C:\ACLReports\server01.acl -InputObject $ShareReport
	 Saves the ACLs in the $ShareReport variable to the file C:\ACLReports\server01.acl.

.EXAMPLE 
	 Export-ACLReport -Path C:\ACLReports\server01.acl -InputObject (New-ACLShareReport -ComputerName SERVER01) -Force
	 Saves the file ACLs for all shares on the compuer SERVER01 to the file C:\ACLReports\server01.acl. If the file exists it will be overwritten.

.EXAMPLE 
	New-ACLShareReport -ComputerName SERVER01 | Export-ACLReport -Path C:\ACLReports\server01.acl -Force
	Saves the file ACLs for all shares on the compuer SERVER01 to the file C:\ACLReports\server01.acl. If the file exists it will be overwritten.
#>    
function Export-ACLReport {
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$Path,

        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)]
        [ValidateScript({$_.GetType().FullName -ne 'ACLReportTools.Permission[]'})]
        [ACLReportTools.Permission[]]$InputObject,

        [Switch]$Force

    ) # param
    begin
    {
		[ACLReportTools.Permission[]]$InputObjectNew = $Null
	}
	process
    {
		foreach ($I in $InputObject)
        {
			$InputObjectNew += $I
		}
	}
	end
    {
		[Void]$PSBoundParameters.Remove('InputObject')
		$InputObjectNew | Export-ACLPermission @PSBoundParameters
	}
} # Function Export-ACLReport


<#
.SYNOPSIS
	Export an ACL Permission Diff Report as a file.

.DESCRIPTION 
	This Cmdlet will save whatever ACL Permission Diff Report that is in the pipeline to a file.

	This cmdlet just calls Export-ACLPermissionDiff although at some point will add additional functionality.
     
.PARAMETER Path
	This is the path to the ACL Permission Diff Report output file. This parameter is required.

.PARAMETER InputObject
	Specifies the Permissions objects to export to the file. Enter a variable that contains the objects or type a command or expression that gets the objects. You can also pipe ACLReportTools.PermissionDiff objects to Export-ACLReport.

.PARAMETER Force
	Causes the file to be overwritten if it exists.

.EXAMPLE 
	Compare-ACLReports -Baseline (Import-ACLReports -Path c:\ACLReports\CLIENT01_2014_11_14.acl) -With (Get-ACLReport -ComputerName CLIENT01) | Export-ACLDiffReport -Path "$HOME\Documents\Compare.acr"
	This will perform a comparison of the current share ACL report from computer CLIENT01 with the stored share ACL report in file c:\ACLReports\CLIENT01_2014_11_14.acl and then export the report file
	to $HOME\Documents\Compare.acr
#>    
function Export-ACLDiffReport
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$Path,

        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)]
        [ValidateScript({$_.GetType().FullName -ne 'ACLReportTools.PermissionDiff[]'})]
        [ACLReportTools.PermissionDiff[]]$InputObject,

        [Switch]$Force

    ) # param
    begin
    {
		[ACLReportTools.PermissionDiff[]]$InputObjectNew = $Null
	}
	process
    {
		foreach ($I in $InputObject)
        {
			$InputObjectNew += $I
		}
	}
	end
    {
		[Void]$PSBoundParameters.Remove('InputObject')
		$InputObjectNew | Export-ACLPermissionDiff @PSBoundParameters
	}
} # Function Export-ACLDiffReport


<#
.SYNOPSIS
	Import the ACL Report that is in a file.

.DESCRIPTION 
	This Cmdlet will import all the ACL Report (ACLReportTools.Permission) objects from a specified file into the pipeline.

	This cmdlet just calls Import-ACLPermission although at some point will add additional functionality.
     
.PARAMETER Path
	This is the path to the ACL Permission Report file to import. This parameter is required.

.EXAMPLE 
	Import-ACLReport -Path C:\ACLReports\server01.acl
	Imports the ACL Share Report from the file C:\ACLReports\server01.acl and puts it into the pipeline
#>    
function Import-ACLReport
{
    [CmdLetBinding()]
#	[OutputType([ACLReportTools.Permission])]
    param
    (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$Path
    ) # param
   
    Import-ACLPermission @PSBoundParameters

} # Function Import-ACLReport


<#
.SYNOPSIS
	Import the ACL Difference Report that is in a file.

.DESCRIPTION 
	This Cmdlet will import all the ACL Difference Report (ACLReportTools.PermissionDiff) objects from a specified file into the pipeline.

	This cmdlet just calls Import-ACLPermissionDiff although at some point will add additional functionality.
     
.PARAMETER Path
	This is the path to the ACL Permission Report file to import. This parameter is required.

.EXAMPLE 
	Import-ACLDiffReport -Path C:\ACLReports\server01.acr
	Imports the ACL Share Report from the file C:\ACLReports\server01Permission and puts it into the pipeline
#>    
function Import-ACLDiffReport
{
    [CmdLetBinding()]
#	[OutputType([ACLReportTools.PermissionDiff])]
    param
    (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$Path
    ) # param
   
    Import-ACLPermissionDiff @PSBoundParameters

} # Function Import-ACLDiffReport


<#
.SYNOPSIS
	Compares two ACL reports and produces an ACL Difference report.

.DESCRIPTION 
	This cmdlets compares two ACL Share reports and produces a difference list in the pipeline that can then be reported on.

	A baseline report (usually from importing a previous ACL Share Report) must be provided. The second ACL Share report (called the current ACL Share report) will be compared against the baseline report.
	The current ACL report will be either generated by the New-ACLShareReport or New-ACLPathFileReport cmdlets (depending on parameters) or it can be passed in via the With variable.
   
.PARAMETER Baseline
	This is the baseline report data the comparison will focus on. It will usually be pulled in from a previously saved Share ACL report via the Import-ACLReports 

.PARAMETER ComputerName
	This is the computer(s) to generate the current list of Share ACLs for to perform the comparison with the baseline. The Computer names can also be passed in via the pipeline.

	This parameter should not be used if the With Parameter is provided.

.PARAMETER Include
	This is a list of shares to include from the comparison. If this parameter is not set it will default to including all shares. This parameter can't be set if the Exclude parameter is set.

	This parameter should not be used if the With Parameter is provided.

.PARAMETER Exclude
	This is a list of shares to exclude from the comparison. If this parameter is not set it will default to excluding no shares. This parameter can't be set if the Include parameter is set.

	This parameter should not be used if the With Parameter is provided.

.PARAMETER With
	This parameter provides an ACL Share report to compare with the Baseline ACL Share report.

	This parameter should not be used if the ComputerName Parameter is provided.

.PARAMETER ReportNoChange
	Setting this switch will cause a 'No Change' report item to be shown when a share is identical in both the baseline and current reports.

.PARAMETER IncludeInherited
	Setting this switch will cause the non inherited file/folder ACLs to be pulled recursively.

.EXAMPLE
	 Compare-ACLReports -Baseline (Import-ACLReports -Path c:\ACLReports\CLIENT01_2014_11_14.acl) -With (Get-ACLReport -ComputerName CLIENT01)
	 This will perform a comparison of the current share ACL report from computer CLIENT01 with the stored share ACL report in file c:\ACLReports\CLIENT01_2014_11_14.acl

.EXAMPLE
	 Compare-ACLReports -Baseline (Import-ACLReports -Path c:\ACLReports\CLIENT01_2014_11_14.acl) -ComputerName CLIENT01
	 This will perform a comparison of the current share ACL report from computer CLIENT01 with the stored share ACL report in file c:\ACLReports\CLIENT01_2014_11_14.acl

.EXAMPLE
	 Compare-ACLReports -Baseline (Import-ACLReports -Path c:\ACLReports\CLIENT01_2014_11_14_SHARE01_ONLY.acl) -ComputerName CLIENT01 -Include SHARE01
	 This will perform a comparison of the current share ACL report from computer CLIENT01 for only SHARE01 with the stored share ACL report in file c:\ACLReports\CLIENT01_2014_11_14_SHARE01_ONLY.acl

.EXAMPLE
	 "CLIENT01" | Compare-ACLReports -Baseline (Import-ACLReports -Path c:\ACLReports\CLIENT01_2014_11_14.acl)
	 This will perform a comparison of the current share ACL report from computer CLIENT01 with the stored share ACL report in file c:\ACLReports\CLIENT01_2014_11_14.acl

.EXAMPLE
	 Compare-ACLReports -Baseline (Import-ACLReports -Path c:\ACLReports\CLIENT01_2014_11_14.acl) -With (Import-ACLReports -Path c:\ACLReports\CLIENT01_2014_06_01.acl)
	 This will perform a comparison of the share ACL report in file c:\ACLReports\CLIENT01_2014_06_01.acl with the stored share ACL report in file c:\ACLReports\CLIENT01_2014_11_14.acl
#>
Function Compare-ACLReports
{
    [CmdLetBinding()]
    param
    (
        [Parameter(
            Mandatory=$true)]
        [ValidateScript( { ($_.GetType() -ne 'ACLReportTools.Permission') -and ($_.GetType() -ne 'Deserialized.ACLReportTools.Permission') } )]
        [Object[]]$Baseline,

        [Parameter(
            ParameterSetName='CompareToCurrentShares',
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)]
        [String[]]$ComputerName=$env:computername,

        [Parameter(
            ParameterSetName='CompareToCurrentShares')]
        [String[]]$Include,

        [Parameter(
            ParameterSetName='CompareToCurrentShares')]
        [String[]]$Exclude,

        [Parameter(
            ParameterSetName='CompareToCurrentFiles')]
        [String[]]$Path,

        [Parameter(
            ParameterSetName='CompareToOther')]
        [ACLReportTools.Permission[]]$With,

        [Switch]$ReportNoChange,
        
        [Switch]$IncludeInherited
    ) # param
    begin
    {
        [ACLReportTools.PermissionDiff[]]$Comparison = $Null
    } # Begin
    process
    {
        switch ($PSCmdLet.ParameterSetName)
        {
			'CompareToCurrentShares' {
				# A report to compare to wasn't specified so we need to generate
				# the current report using the other parameters passed.
				$null = $PSBoundParameters.Remove('Baseline') 
				$null = $PSBoundParameters.Remove('With')
				$null = $PSBoundParameters.Remove('ReportNoChange')
				$null = $PSBoundParameters.Remove('Path')
				Write-Verbose -Message "Assembling current ACL Share report for comparison."
				[ACLReportTools.Permission[]]$With += New-ACLShareReport @PSBoundParameters
				Break
			}
			'CompareToCurrentFiles' {
				# A report to compare to wasn't specified so we need to generate
				# the current report using the other parameters passed.
				$null = $PSBoundParameters.Remove('Baseline')
				$null = $PSBoundParameters.Remove('With')
				$null = $PSBoundParameters.Remove('ReportNoChange')
				$null = $PSBoundParameters.Remove('ComputerName')
				$null = $PSBoundParameters.Remove('Include')
				$null = $PSBoundParameters.Remove('Exclude')
				Write-Verbose -Message "Assembling current ACL Path File report for comparison."
				[ACLReportTools.Permission[]]$With += New-ACLPathFileReport @PSBoundParameters
				Break
			}
		}
    } # Process
    end
    {
        # The actual comparions is performed now
        # Get list of shares and computers we are going to compare the ACLs from
        $Current_Computers = $With | Select-Object -ExpandProperty ComputerName -Unique
        if ($Current_Computers.Length -eq 0)
        {
            Write-Error "No accessible shares were found on the computers specified."
            Return
        }
        else
        {
            $Baseline_Computers = $Baseline | Select-Object -ExpandProperty ComputerName -Unique
            foreach ($Current_Computer in $current_Computers)
            {
                # Perform a share comparison on each of the current computers
                If ($baseline_Computers -contains $Current_Computer)
                {
                    Write-Verbose -Message "Performing share comparison of computer $Current_Computer."
                    # Assemble the list of shares for the computer
                    $Current_Shares = $With | Where-Object -Property ComputerName -eq $Current_Computer | Select-Object -ExpandProperty Share -Unique
                    $Baseline_Shares = $Baseline | Where-Object -Property ComputerName -eq $Current_Computer | Select-Object -ExpandProperty Share -Unique
                    foreach ($current_share in $current_shares)
                    {
                         if ($baseline_shares -contains $current_share)
                         {
                            Write-Verbose -Message "Performing share comparison of share $Current_Share on computer $Current_Computer."

                            # Assemble list of ACLS for share/computer
                            $Filter = [ScriptBlock]::Create({ ($_.ComputerName -eq $Current_Computer) `
                                -and ($_.Share -eq $Current_Share) -and ($_.Type -eq [ACLReportTools.PermissionTypeEnum]::Share) })
                            $Current_Share_Acls = $With | Where-Object -FilterScript $Filter
                            $Baseline_Share_Acls = $Baseline | Where-Object -FilterScript $Filter
                            [boolean]$changes = $false

                            # Now compare the current share ACLs wth the baseline share ACLs
                            foreach ($current_share_acl in $current_share_acls)
                            {
                                [string]$c_accesscontroltype = $current_share_acl.Access.AccessControlType
                                [string]$c_filesystemrights = $current_share_acl.Access.AccessRights
                                [string]$c_identityreference = $current_share_acl.Access.Account
                                [boolean]$acl_found = $false
                                foreach ($baseline_share_acl in $baseline_share_acls)
                                {
                                    [string]$b_accesscontroltype = $baseline_share_acl.Access.AccessControlType
                                    [string]$b_filesystemrights = $baseline_share_acl.Access.AccessRights
                                    [string]$b_identityreference = $baseline_share_acl.Access.Account
                                    if ($c_identityreference -eq $b_identityreference)
                                    {
                                        $acl_found = $true
                                        break
                                    } # If

                                } # Foreach

                                if ($acl_found)
                                {
                                    # The IdentityReference (user) exists in both the Baseline and the Current ACLs
                                    # Check it's the same though
                                    if ($c_filesystemrights -ne $b_filesystemrights)
                                    {
                                        # The Permission rights are different
                                        $DiffMessage = "Share permission rights changed from '$b_filesystemrights' to '$c_filesystemrights' for '$c_identityreference'."                                    
                                        Write-Verbose -Message $DiffMessage
                                        $Comparison += New-PermissionDiffObject `
                                            -Type ([ACLReportTools.PermissionTypeEnum]::Share) `
                                            -DiffType ([ACLReportTools.PermissionDiffEnum]::'Permission Rights Changed') `
                                            -ComputerName $Current_Computer -Share $Current_Share `
                                            -Difference $DiffMessage                                        
                                        $changes = $true

                                    }
                                    elseif ($c_accesscontroltype -ne $b_accesscontroltype)
                                    {
                                        # The Permission access control type is different
                                        $DiffMessage = "Share permission '$c_filesystemrights $c_accesscontroltype' for '$c_identityreference' added."                                    
                                        Write-Verbose -Message $DiffMessage
                                        $Comparison += New-PermissionDiffObject `
                                            -Type ([ACLReportTools.PermissionTypeEnum]::Share) `
                                            -DiffType ([ACLReportTools.PermissionDiffEnum]::'Permission Access Control Changed') `
                                            -ComputerName $Current_Computer -Share $Current_Share `
                                            -Difference $DiffMessage
                                        $changes = $true
                                    } # If
                                }
                                else
                                {
                                    # The ACL wasn't found in the baseline so it must be newly added
                                    $DiffMessage = "Share permission '$c_filesystemrights $c_accesscontroltype' for '$c_identityreference' added."                                    
                                    Write-Verbose -Message $DiffMessage
                                    $Comparison += New-PermissionDiffObject `
                                        -Type ([ACLReportTools.PermissionTypeEnum]::Share) `
                                        -DiffType ([ACLReportTools.PermissionDiffEnum]::'Permission Added') `
                                        -ComputerName $Current_Computer -Share $Current_Share `
                                        -Difference $DiffMessage
                                    $changes = $true
                                } # If
                            } # Foreach
            
                            # Now compare the baseline share ACLs wth the current share ACLs
                            # We only need to check if a ACL has been removed from the baseline
                            foreach ($baseline_share_acl in $baseline_share_acls)
                            {
                                [string]$b_accesscontroltype = $baseline_share_acl.Access.AccessControlType
                                [string]$b_filesystemrights = $baseline_share_acl.Access.AccessRights
                                [string]$b_identityreference = $baseline_share_acl.Access.Account
                                [boolean]$acl_found = $false
                                foreach ($current_share_acl in $current_share_acls)
                                {
                                    [string]$c_accesscontroltype = $current_share_acl.Access.AccessControlType
                                    [string]$c_filesystemrights = $current_share_acl.Access.AccessRights
                                    [string]$c_identityreference = $current_share_acl.Access.Account
                                    if ($c_identityreference -eq $b_identityreference)
                                    {
                                        $acl_found = $true
                                        break
                                    } # If
                                } # Foreach

                                if (-not $acl_found)
                                {
                                    # The IdentityReference (user) exists in the Baseline but not in the Current
                                    $DiffMessage = "Share permission '$b_filesystemrights $b_accesscontroltype' for '$b_identityreference' removed."
                                    Write-Verbose -Message $DiffMessage
                                    $Comparison += New-PermissionDiffObject `
                                        -Type ([ACLReportTools.PermissionTypeEnum]::Share) `
                                        -DiffType ([ACLReportTools.PermissionDiffEnum]::'Permission Removed') `
                                        -ComputerName $Current_Computer -Share $Current_Share `
                                        -Difference $DiffMessage
                                    $changes = $true
                                } # If
                            } # Foreach

                            # Perform the baseline to current file/folder ACL comparison
                            $Filter = [ScriptBlock]::Create({ ($_.ComputerName -eq $Current_Computer) `
                                -and ($_.Share -eq $Current_Share) `
                                -and (($_.Type -eq [ACLReportTools.PermissionTypeEnum]::File) -or ($_.Type -eq [ACLReportTools.PermissionTypeEnum]::Folder)) })
                            $Current_file_Acls = $With | Where-Object -FilterScript $Filter
                            $Baseline_file_Acls = $Baseline | Where-Object -FilterScript $Filter
                            [string]$last_path = '.'

                            foreach ($current_file_acl in $current_file_acls)
                            {
                                # Put all the Current File ACL props into variables for easy access.
                                [string]$c_path = $current_file_acl.Path
                                [string]$c_owner = $current_file_acl.Owner
                                [string]$c_inherited = $current_file_acl.Inherited
                                $c_access = $current_file_acl.Access
                                [string]$c_accesscontroltype = $c_access.AccessControlType
                                [string]$c_filesystemrights = $c_access.AccessRights
                                [string]$c_identityreference = $c_access.Account
                                [string]$c_appliesto=Convert-FileSystemAppliesToString -InheritanceFlags $c_access.InheritanceFlags -PropagationFlags $c_access.PropagationFlags
                                [boolean]$acl_found = $false
                                foreach ($baseline_file_acl in $baseline_file_acls)
                                {
                                    [string]$b_path = $baseline_file_acl.Path
                                    [string]$b_owner = $baseline_file_acl.Owner
                                    [string]$b_inherited = $baseline_file_acl.Inherited
                                    $b_access = $baseline_file_acl.Access
                                    [string]$b_accesscontroltype = $b_access.AccessControlType
                                    [string]$b_filesystemrights = $b_access.AccessRights
                                    [string]$b_identityreference = $b_access.Account
                                    [string]$b_appliesto=Convert-FileSystemAppliesToString -InheritanceFlags $b_access.InheritanceFlags -PropagationFlags $b_access.PropagationFlags
                                    if ($c_path -eq $b_path)
                                    {
                                        # Perform an owner check on each file/folder only once
                                        # If we've already checked this path, don't bother checking the owner again.
                                        if ($last_path -ne $c_path)
                                        {
                                            if ($c_owner -ne $b_owner)
                                            {
                                                # The Permission Owner are different
                                                $DiffMessage = "$([ACLReportTools.PermissionTypeEnum]$current_file_acl.Type) $c_path owner changed from '$b_owner' to '$c_owner'."
                                                Write-Verbose -Message $DiffMessage
                                                $Comparison += New-PermissionDiffObject `
                                                    -Type ($current_file_acl.Type) `
                                                    -Path $c_path `
                                                    -DiffType ([ACLReportTools.PermissionDiffEnum]::'Owner Changed') `
                                                    -ComputerName $Current_Computer -Share $Current_Share `
                                                    -Difference $DiffMessage
                                                $changes = $true
                                            } # If
                                            
                                            if ($c_inherited -ne $b_inherited)
                                            {
                                                # The inheritance is different
                                                $DiffMessage = "$([ACLReportTools.PermissionTypeEnum]$current_file_acl.Type) $c_path inheritance changed from '$b_inheritance' to '$c_inheritance'."
                                                Write-Verbose -Message $DiffMessage
                                                $Comparison += New-PermissionDiffObject `
                                                    -Type ($current_file_acl.Type) `
                                                    -Path $c_path `
                                                    -DiffType ([ACLReportTools.PermissionDiffEnum]::'Inheritance Changed') `
                                                    -ComputerName $Current_Computer -Share $Current_Share `
                                                    -Difference $DiffMessage
                                                $changes = $true                                                
                                            }

                                            $last_path = $c_path
                                        } # If
                                        # Check that the Identity Reference (user) is the same one
                                        # And that the Applies To is the same
                                        if (($c_identityreference -eq $b_identityreference) -and ($c_appliesto -eq $b_appliesto))
                                        {
                                            $acl_found = $true
                                            break
                                        }
                                    } # If
                                } # Foreach
                                
                                if ($acl_found)
                                {
                                    # The IdentityReference (user) and path exists in both the Baseline and the Current ACLs
                                    # Check it's the same though
                                    if ($c_filesystemrights -ne $b_filesystemrights)
                                    {
                                        # The Permission rights are different
                                        $DiffMessage = "$([ACLReportTools.PermissionTypeEnum]$current_file_acl.Type) $c_path permission rights changed from '$b_filesystemrights' to '$c_filesystemrights' for '$c_identityreference'."
                                        Write-Verbose -Message $DiffMessage
                                        $Comparison += New-PermissionDiffObject `
                                            -Type ($current_file_acl.Type) `
                                            -Path $c_path `
                                            -DiffType ([ACLReportTools.PermissionDiffEnum]::'Permission Rights Changed') `
                                            -ComputerName $Current_Computer -Share $Current_Share `
                                            -Difference $DiffMessage                                        
                                        $changes = $true
                                    } # If
									if ($c_accesscontroltype -ne $b_accesscontroltype)
                                    {
                                        # The Permission access control type is different
                                        $DiffMessage = "$([ACLReportTools.PermissionTypeEnum]$current_file_acl.Type) $c_path permission access control type changed from '$b_accesscontroltype' to '$c_accesscontroltype' for '$c_identityreference'."
                                        Write-Verbose -Message $DiffMessage
                                        $Comparison += New-PermissionDiffObject `
                                            -Type ($current_file_acl.Type) `
                                            -Path $c_path `
                                            -DiffType ([ACLReportTools.PermissionDiffEnum]::'Permission Access Control Changed') `
                                            -ComputerName $Current_Computer -Share $Current_Share `
                                            -Difference $DiffMessage
                                        $changes = $true
                                    } # If
                                }
                                else
                                {
                                    # The Permission was not found in the baseline so it must have been added
                                    $DiffMessage = "$([ACLReportTools.PermissionTypeEnum]$current_file_acl.Type) $c_path permission '$c_filesystemrights, $c_accesscontroltype, $c_appliesto' added for '$c_identityreference'."
                                    Write-Verbose -Message $DiffMessage
                                    $Comparison += New-PermissionDiffObject `
                                        -Type ($current_file_acl.Type) `
                                        -Path $c_path `
                                        -DiffType ([ACLReportTools.PermissionDiffEnum]::'Permission Added') `
                                        -ComputerName $Current_Computer -Share $Current_Share `
                                        -Difference $DiffMessage
                                    $changes = $true
                                } # If
                            } # Foreach

                            # Now compare the baseline file ACLs wth the current file ACLs
                            # We only need to check if a ACL has been removed from the baseline
                            foreach ($baseline_file_acl in $baseline_file_acls)
                            {
                                [string]$b_path = $baseline_file_acl.Path
                                [string]$b_owner = $baseline_file_acl.Owner
                                $b_access = $baseline_file_acl.Access
                                [string]$b_accesscontroltype = $b_access.AccessControlType
                                [string]$b_filesystemrights = $b_access.AccessRights
                                [string]$b_identityreference = $b_access.Account
                                [string]$b_appliesto = Convert-FileSystemAppliesToString -InheritanceFlags $b_access.InheritanceFlags -PropagationFlags $b_access.PropagationFlags
                                [boolean]$acl_found = $false
                                foreach ($current_file_acl in $current_file_acls)
                                {
                                    [string]$c_path = $current_file_acl.Path
                                    [string]$c_owner = $current_file_acl.Owner
                                    $c_access = $current_file_acl.Access
                                    [string]$c_accesscontroltype = $c_access.AccessControlType
                                    [string]$c_filesystemrights = $c_access.AccessRights
                                    [string]$c_identityreference = $c_access.Account
                                    [string]$c_appliesto = Convert-FileSystemAppliesToString -InheritanceFlags $c_access.InheritanceFlags -PropagationFlags $c_access.PropagationFlags
                                    if (($c_path -eq $b_path) -and ($c_identityreference -eq $b_identityreference) -and ($c_appliesto -eq $b_appliesto))
                                    {
                                        $acl_found = $true
                                        break
                                    } # If
                                } # Foreach
                                if (-not $acl_found)
                                {
                                    # The IdentityReference (user) and path exists in the Baseline but not in the Current
                                    $DiffMessage = "$([ACLReportTools.PermissionTypeEnum]$baseline_file_acl.Type) $b_path permission '$b_filesystemrights, $b_accesscontroltype, $b_appliesto' removed for '$b_identityreference'."
                                    Write-Verbose -Message $DiffMessage
                                    $Comparison += New-PermissionDiffObject `
                                        -Type ($baseline_file_acl.Type) `
                                        -Path $b_path `
                                        -DiffType ([ACLReportTools.PermissionDiffEnum]::'Permission Removed') `
                                        -ComputerName $Current_Computer -Share $Current_Share `
                                        -Difference $DiffMessage
                                    $changes = $true
                                } # If
                            } # Foreach

                            # If no changes have been made to any of the Share or File/Folder ACLs then say so
                            if (-not $changes)
                            {
                                $DiffMessage = "The share, file and folder permissions for the share $Current_Share on $Current_Computer have not changed."
                                Write-Verbose -Message $DiffMessage
                                if ($ReportNoChange)
                                {
                                    $Comparison += New-PermissionDiffObject `
                                        -Type ([ACLReportTools.PermissionTypeEnum]::Share) `
                                        -DiffType ([ACLReportTools.PermissionDiffEnum]::'No Change') `
                                        -ComputerName $Current_Computer -Share $Current_Share `
                                        -Difference $DiffMessage
                                } # If ($ReportNoChange)
                            } # If
                         }
                         else
                         {
                            # The Share exists in the Current but not in the Baseline
                            $DiffMessage = "The share $Current_Share on computer $Current_Computer has been added."
                            Write-Verbose -Message $DiffMessage
                            $Comparison += New-PermissionDiffObject `
                                -Type ([ACLReportTools.PermissionTypeEnum]::Share) `
                                -DiffType ([ACLReportTools.PermissionDiffEnum]::'Share Added') `
                                -ComputerName $Current_Computer -Share $Current_Share `
                                -Difference $DiffMessage
 
                            # Get the Current File/Folder ACLs to an Array
                            $Filter = [ScriptBlock]::Create({ ($_.ComputerName -eq $Current_Computer) -and ($_.Share -eq $Current_Share) -and (($_.Type -eq [ACLReportTools.PermissionTypeEnum]::File) -or ($_.Type -eq [ACLReportTools.PermissionTypeEnum]::Folder)) })
                            $Current_file_Acls = $With | Where-Object -FilterScript $Filter

                            # Output all the current share ACLs into the report as the share is new all permissions must also be new
                            foreach ($current_file_acl in $current_file_acls)
                            {
                                [string]$c_path = $current_file_acl.Path
                                [string]$c_owner = $current_file_acl.Owner
                                $c_access = $current_file_acl.Access
                                [string]$c_accesscontroltype = $c_access.AccessControlType
                                [string]$c_filesystemrights = $c_access.AccessRights
                                [string]$c_identityreference = $c_access.Account

                                # Because this is a new share, the permission has always been added
                                $DiffMessage = "$($current_file_acl.Type) $c_path permission '$c_filesystemrights, $c_accesscontroltype, $c_appliesto' added for '$c_identityreference'."
                                Write-Verbose -Message $DiffMessage
                                $Comparison += New-PermissionDiffObject `
                                    -Type ($current_file_acl.Type) `
                                    -Path $c_path `
                                    -DiffType ([ACLReportTools.PermissionDiffEnum]::'Permission Added') `
                                    -ComputerName $Current_Computer -Share $Current_Share `
                                    -Difference $DiffMessage
                            } # Foreach ($current_file_acl in $current_file_acls)
                         } # If ($baseline_shares -contains $current_share)
                    } # Foreach ($current_share in $current_shares)

                    # Check for any removed shares
                    foreach ($baseline_share in $baseline_shares)
                    {
                        if ($current_shares -notcontains $baseline_share)
                        {
                            # Baseline Share does not exist in Current Shares (Share removed)
                            $DiffMessage = "The share $baseline_share on computer $Current_Computer has been removed."
                            Write-Verbose -Message $DiffMessage
                            $Comparison += New-PermissionDiffObject `
                                -Type ([ACLReportTools.PermissionTypeEnum]::Share) `
                                -DiffType ([ACLReportTools.PermissionDiffEnum]::'Share Removed') `
                                -ComputerName $Current_Computer -Share $baseline_share `
                                -Difference $DiffMessage
                        } # If ($current_shares -notcontains $baseline_share)
                    } # Foreach ($baseline_share in $baseline_shares)
                }
                else
                {
                    # The Computer exists in the Current but not in the Baseline
                    $DiffMessage = "Skiping share comparison of computer $Current_Computer because it was not found in the baseline report."
                    Write-Verbose -Message $DiffMessage
                    $Comparison += New-PermissionDiffObject `
                        -DiffType ([ACLReportTools.PermissionDiffEnum]::'Computer Added') `
                        -ComputerName $Current_Computer `
                        -Difference $DiffMessage
                } # If ($baseline_Computers.Contains($Current_Computer))
            } # Foreach ($Current_Computer in $current_Computers) 

            # Check for any removed computers
            foreach ($baseline_computer in $baseline_computers)
            {
                if ($current_computers -notcontains $baseline_computer)
                {
                    # Baseline computer does not exist in Current computer (Computer removed)
                    $DiffMessage = "The computer $Current_Computer has been removed."
                    Write-Verbose -Message $DiffMessage
                    $Comparison += New-PermissionDiffObject `
                        -DiffType ([ACLReportTools.PermissionDiffEnum]::'Computer Removed') `
                        -ComputerName $Current_Computer `
                        -Difference $DiffMessage

                } # If
            } # Foreach ($baseline_computer in $baseline_computers)
        } # If
        # Push the comparison result objects into the pipeline
        $Comparison
    } # End
} # Function Compare-ACLReports


##########################################################################################################################################
# Support CmdLets
##########################################################################################################################################

<#
.SYNOPSIS
	Gets a list of the Shares on a specified computer(s) with specified inclusions or exclusions.

.DESCRIPTION 
	This function will pull a list of shares that are set up on the specified computer. Shares can also be included or excluded from the share list by setting the Include or Exclude properties.

	The Cmdlet returns an array of ACLReportTools.Share objects.
     
.PARAMETER ComputerName
	This is the computer to get the shares from. If this parameter is not set it will default to the current machine.

.PARAMETER Include
	This is a list of shares to include from the computer. If this parameter is not set it will default to including all shares. This parameter can't be set if the Exclude parameter is set.

.PARAMETER Exclude
	This is a list of shares to exclude from the computer. If this parameter is not set it will default to excluding no shares. This parameter can't be set if the Include parameter is set.

.EXAMPLE 
	 Get-ACLShare -ComputerName CLIENT01
	 Returns a list of all shares set up on the CLIENT01 machine.

.EXAMPLE 
	 Get-ACLShare -ComputerName CLIENT01 -Include MyShare,OtherShare
	 Returns a list of shares that are set up on the CLIENT01 machine that are named either MyShare or OtherShare.

.EXAMPLE 
	 Get-ACLShare -ComputerName CLIENT01 -Exclude SysVol
	 Returns a list of shares that are set up on the CLIENT01 machine that are not called SysVol.

.EXAMPLE 
	 Get-ACLShare -ComputerName CLIENT01,CLIENT02
	 Returns a list of shares that are set up on the CLIENT01 and CLIENT02 machines.

.EXAMPLE 
	 Get-ACLShare -ComputerName CLIENT01,CLIENT02 -Exclude SysVol
	 Returns a list of shares that are set up on the CLIENT01 and CLIENT02 machines that are not called SysVol.
#>
Function Get-ACLShare {
    [CmdLetBinding()]
    param
    (
        [Parameter(
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)]
        [String[]]$ComputerName=$env:computername,

        [String[]]$Include,

        [String[]]$Exclude
    ) # param
    begin
    {
        [ACLReportTools.Share[]]$SelectedShares = $null
    } # Begin
    process
    {
        foreach ($Computer in $ComputerName)
        {
            Write-Verbose -Message "Getting shares list on computer $Computer"
            [Array]$AllShares = Get-WMIObject -Class win32_share -ComputerName $Computer |
                Where-Object { $_.Name -notlike "*$" } |
                select -ExpandProperty Name
            foreach ($Share in $AllShares)
            {
                if ($Include.Count -gt 0)
                {
                    if ($Share -in $Include)
                    {
                        Write-Verbose -Message "$Share on computer $Computer Included"
                        $SelectedShares += New-ShareObject -ComputerName $Computer -ShareName $Share 
                    }
                    else
                    {
                        Write-Verbose -Message "$Share on computer $Computer Not Included"
                    }
                }
                elseif ($Exclude.Count -gt 0)
                {
                    if ($Share -in $Exclude)
                    {
                        Write-Verbose -Message "$Share on computer $Computer Excluded"
                    }
                    else
                    {
                        Write-Verbose -Message "$Share on computer $Computer Not Excluded"
                        $SelectedShares += New-ShareObject -ComputerName $Computer -ShareName $Share
                    }
                }
                else
                {
                    Write-Verbose -Message "$Share on computer $Computer Included"
                    $SelectedShares += New-ShareObject -ComputerName $Computer -ShareName $Share
                } # If
            } # Foreach ($Share in $AllShares)
        } # Foreach ($Computer In $ComputerName)
    } # Process
    end
    {
        Return $SelectedShares
    } # End
} # Function Get-ACLShare


<#
.SYNOPSIS
	Gets the ACLs for a specified Share.

.DESCRIPTION 
	This function will return the share ACLs for the specified share.
     
.PARAMETER ComputerName
	This is the computer to get the share ACLs from. If this parameter is not set it will default to the current machine.

.PARAMETER ShareName
	This is the share name to pull the share ACLs for.

.PARAMETER Shares
	This is a pipeline parameter that should be used for passing in a list of shares and computers to pull ACLs for. This parameter expects an array of [ACLReportTools.Share] objects.

	This parameter is usually used with the Get-ACLShare CmdLet.

	For example:

	Get-ACLShare -ComputerName CLIENT01,CLIENT02 -Exclude SYSVOL | Get-ACLShareACL 

.EXAMPLE 
	Get-ACLShareACL -ComputerName CLIENT01 -ShareName MyShre
	Returns the share ACLs for the MyShare Share on the CLIENT01 machine.
#>
function Get-ACLShareACL
{
    [CmdLetBinding()]
    param
    (
        [Parameter(
            ParameterSetName='ByParameters')]
        [String]$ComputerName=$env:computername,
        
        [Parameter(
            ParameterSetName='ByParameters')]
        [String]$ShareName,

        [Parameter(
            ParameterSetName='ByPipeline',
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)]
        [ACLReportTools.Share[]]$Shares
    ) # param

    begin
    {
        # Create an empty array to store all the Share ACLs.
        [ACLReportTools.Permission[]]$share_acls = $null
    } # Begin
    process
    {
        if ($PsCmdlet.ParameterSetName -eq 'ByPipeline')
        {
            $ComputerName = $_.ComputerName
            $ShareName = $_.Name
        }
        $objShareSec = Get-WMIObject -Class Win32_LogicalShareSecuritySetting -Filter "name='$ShareName'"  -ComputerName $ComputerName 
        try
        {  
            $SD = $objShareSec.GetSecurityDescriptor().Descriptor
            foreach ($ace in $SD.DACL)
            {   
                $UserName = $ace.Trustee.Name
                if ($ace.Trustee.Domain -ne $Null) { $UserName = "$($ace.Trustee.Domain)\$UserName" }    
                if ($ace.Trustee.Name -eq $Null) { $UserName = $ace.Trustee.SIDString }      
                $fs_rule = New-Object Security.AccessControl.FileSystemAccessRule($UserName, $ace.AccessMask, $ace.AceType)
                $type = [ACLReportTools.PermissionTypeEnum]::Share
                $acl_object =  New-PermissionObject `
                    -Type $type `
                    -ComputerName $ComputerName `
                    -Share $ShareName `
                    -Access $fs_rule
                $share_acls += $acl_object
           } # Foreach           
        }
        catch
        { 
            Write-Error "Unable to obtain share ACLs for $ShareName"
        } # Try
    } # Process
    end
    {
        Return $share_acls
    } # End
} # function Get-ACLShareACL


<#
.SYNOPSIS
	Gets all the file/folder ACLs definited within a specified Share.

.DESCRIPTION 
	This function will return a list of file/folder ACLs for the specified share. If the Recurse switch is used then files/folder ACLs will be scanned recursively. If the IncludeInherited switch is set then inherited file/folder permissions will also be returned, otherwise only non-inherited permissions will be returned. 
     
.PARAMETER ComputerName
	This is the computer to get the share ACLs from. If this parameter is not set it will default to the current machine.

.PARAMETER ShareName
	This is the share name to pull the file/folder ACLs for.

.PARAMETER Recurse
	Setting this switch will cause the file/folder ACLs to be pulled recursively.

.PARAMETER IncludeInherited
	Setting this switch will cause the non inherited file/folder ACLs to be pulled recursively.

.EXAMPLE 
	Get-ACLShareFileACL -ComputerName CLIENT01 -ShareName MyShare
	Returns the file/folder ACLs for the root of MyShare Share on the CLIENT01 machine.

.EXAMPLE 
	Get-ACLShareFileACL -ComputerName CLIENT01 -ShareName MyShare -Recurse
	Returns the file/folder ACLs for all files/folders recursively inside the MyShare Share on the CLIENT01 machine.
#>    
function Get-ACLShareFileACL
{
    [CmdLetBinding()]
    param
    (
        [Parameter(
            ParameterSetName='ByParameters')]
        [String]$ComputerName=$env:computername,
        
        [Parameter(
            ParameterSetName='ByParameters')]
        [String]$ShareName,

        [Parameter(
            ParameterSetName='ByPipeline',
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)]
        [ACLReportTools.Share[]]$Shares,

        [Switch]$Recurse,
        
        [Switch]$IncludeInherited
    ) # param

    begin
    {
        # Create an empty array to store all the non inherited file/folder ACLs.
        [ACLReportTools.Permission[]]$file_acls = $null
    } # Begin
    process
    {
        if ($PsCmdlet.ParameterSetName -eq 'ByPipeline')
        {
            $ComputerName = $_.ComputerName
            $ShareName = $_.Name
        }
        # Now generate the root file/folder ACLs
		$Path = "\\$ComputerName\$ShareName"   
		[Security2.FileSystemAccessRule2[]]$root_file_acl = Get-NTFSAccess -Path $path
		[String]$owner = (Get-NTFSOwner -Path $path).Owner.AccountName
        foreach ($access in $root_file_acl)
        {
			# Write each non-inherited ACL from the root into the array of ACL's 
            if ($access.IsInherited)
            {
                [String] $Inherited = "Inherited from $($access.InheritedFrom)" 
            }
            else
            {
                [String] $Inherited = 'Not-inherited'                    
            }
			$file_acls += New-PermissionObject `
                -ComputerName $ComputerName `
                -Type ([ACLReportTools.PermissionTypeEnum]::Folder) `
                -Path $Path `
                -Owner $owner `
                -Access $access `
                -Share $ShareName `
                -Inherited $Inherited
			Write-Verbose -Message "Get-ACLShareFileACL: Root ACL for $ShareName path $Path owner $Owner`n$(Convert-AccessToString($Access))"
        } # Foreach
        if ($Recurse)
        {
            # Generate all file/folder ACLs for subfolders and/or files containined within the share recursively
            $node_file_acls = Get-childitem -Path $Path -recurse |
                Get-NTFSAccess              
	        if (! $IncludeInherited)
            {
                # Generate any non-inferited file/folder ACLs for subfolders and/or files containined within the share recursively
                $node_file_acls = $node_file_acls | Where-Object -Property IsInherited -eq $False    
            }
	        $lastPath = ''
			foreach ($access in $node_file_acls)
            {
				# Write each non-inherited ACL from the file/folder into the array of ACL's 
				$Path = $access.FullName
				if ($lastPath -ne $Path)
                {
					try
                    {
						[Boolean]$IsFolder = ((Get-Item -Path $Path -ErrorAction "Stop") -is [System.IO.DirectoryInfo])
					}
                    catch
                    {
						Write-Warning "Get-ACLPathFileACL: Access Denied to $Path"
						$IsFolder = $True
					}
					if ($IsFolder)
                    {
						$type = [ACLReportTools.PermissionTypeEnum]::Folder
					}
                    else
                    {
						$type = [ACLReportTools.PermissionTypeEnum]::File
					}
					[String]$Owner = (Get-NTFSOwner -Path $Path).Owner.AccountName
					$lastPath = $access.FullName
				}            
	            if ($IncludeInherited -and $access.IsInherited)
                {
                    [String] $Inherited = "Inherited from $($access.InheritedFrom)" 
                }
                else
                {
                    [String] $Inherited = 'Not-inherited'                    
                }
                $file_acls += New-PermissionObject `
                    -ComputerName $ComputerName `
                    -Type $type `
                    -Path $Path `
                    -Owner $owner `
                    -Access $access `
                    -Share $ShareName `
                    -Inherited $Inherited
				Write-Verbose -Message "Get-ACLShareFileACL: $Inherited ACL for $ShareName path $Path owner $Owner`n$(Convert-AccessToString($Access))"
			} # Foreach
        } # If
    } # Process
    end
    {
        Return $file_acls
    } # End
} # Function Get-ACLShareFileACL


<#
.SYNOPSIS
	Gets all the file/folder ACLs defined within a specified Path.

.DESCRIPTION 
	This function will return a list of file/folder ACLs for the specified share. If the Recurse switch is used then files/folder ACLs will be scanned recursively. If the IncludeInherited switch is set then inherited file/folder permissions will also be returned, otherwise only non-inherited permissions will be returned. 
     
.PARAMETER Path
	This is the path to pull the file/folder ACLs for.

.PARAMETER Recurse
	Setting this switch will cause the file/folder ACLs to be pulled recursively.

.PARAMETER IncludeInherited
	Setting this switch will cause the non inherited file/folder ACLs to be pulled recursively.

.EXAMPLE 
	Get-ACLPathFileACL -Path C:\Users
	Returns the file/folder ACLs for the root of C:\Users folder.

.EXAMPLE 
	Get-ACLPathFileACL -Path C:\Users -Recurse
	Returns the file/folder ACLs for all files/folders recursively inside the C:\Users folder.
#>    
function Get-ACLPathFileACL
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$Path,

        [Switch]$Recurse,
        
        [Switch]$IncludeInherited
    ) # param

    # Create an empty array to store all the non inherited file/folder ACLs.
    [ACLReportTools.Permission[]]$file_acls = $null
	[String]$ComputerName = $ENV:ComputerName

    # Now generate the root file/folder ACLs
	[Security2.FileSystemAccessRule2[]]$root_file_acl = Get-NTFSAccess -Path $path
    [String]$owner = (Get-NTFSOwner -Path $path).Owner.AccountName
    foreach ($access in $root_file_acl)
    {
        # Write each non-inherited ACL from the root into the array of ACL's 
        if ($access.IsInherited)
        {
            $Inherited = "Inherited from $($access.InheritedFrom)" 
        }
        else
        {
            $Inherited = 'Not-inherited'                    
        }
        $file_acls += New-PermissionObject `
            -ComputerName $ComputerName `
            -Type ([ACLReportTools.PermissionTypeEnum]::Folder) `
            -Path $Path `
            -Owner $owner `
            -Access $access `
            -Inherited $Inherited
		Write-Verbose -Message "Get-ACLPathFileACL: Root ACL for $Path owner $Owner`n$(Convert-AccessToString($Access))"
    } # Foreach
    if ($Recurse)
    {
        # Generate all file/folder ACLs for subfolders and/or files containined within the share recursively
        $node_file_acls = Get-childitem -Path $Path -recurse |
            Get-NTFSAccess              
        if (! $IncludeInherited)
        {
            # Generate any non-inferited file/folder ACLs for subfolders and/or files containined within the share recursively
            $node_file_acls = $node_file_acls | Where-Object -Property IsInherited -eq $False    
        }
        $LastPath = ''
		foreach ($access in $node_file_acls)
        {
			# Write each non-inherited ACL from the file/folder into the array of ACL's 
	        $Path = $access.FullName
			if ($LastPath -ne $Path)
            {
				try
                {
					[Boolean]$IsFolder = ((Get-Item -Path $Path -ErrorAction "Stop") -is [System.IO.DirectoryInfo])
				}
                catch
                {
					Write-Warning "Get-ACLPathFileACL: Access Denied to $Path"
					$IsFolder = $True
				}
				if ($IsFolder)
                {
					$type = [ACLReportTools.PermissionTypeEnum]::Folder
				}
                else
                {
					$type = [ACLReportTools.PermissionTypeEnum]::File
				}
				[String] $Owner = (Get-NTFSOwner -Path $Path).Owner.AccountName
				$LastPath = $Path
			}            
            if ($IncludeInherited -and $access.IsInherited)
            {
                $Inherited = "Inherited from $($access.InheritedFrom)" 
            }
            else
            {
                $Inherited = 'Not-inherited'                    
            }
			$file_acls += New-PermissionObject `
                -ComputerName $ComputerName `
                -Type $type `
                -Path $Path `
                -Owner $Owner `
                -Access $Access `
                -Inherited $Inherited
			Write-Verbose -Message "Get-ACLPathFileACL: $Inherited ACL for $Path owner $Owner`n$(Convert-AccessToString($Access))"
        } # Foreach
    } # If
    return $file_acls
} # Function Get-ACLPathFileACL


<#
.SYNOPSIS
	Export the ACL Permissions objects that are provided as a file.

.DESCRIPTION 
	This Cmdlet will save what ever ACLs (ACLReportTools.Permission) to a file.
     
.PARAMETER Path
	This is the path to the ACL Permissions file output file. This parameter is required.

.PARAMETER InputObject
	Specifies the ACL Permissions objects to export to the file. Enter a variable that contains the objects or type a command or expression that gets the objects. You can also pipe ACLReportTools.Permission objects to cmdlet.

.PARAMETER Force
	Causes the file to be overwritten if it exists.

.EXAMPLE 
	New-ACLPathFileReport -Path e:\Shares | Export-ACLPermission -Path C:\ACLReports\server01.acl
	Creates a new ACL Permission report for e:\Shares and saves it to the file C:\ACLReports\server01.acl.

.EXAMPLE 
	Export-ACLPermission -Path C:\ACLReports\server01.acl -InputObject $Acls
	Saves the ACL Permissions in the $Acls variable to the file C:\ACLReports\server01.acl.

.EXAMPLE 
	Export-ACLPermission -Path C:\ACLReports\server01.acl -InputObject (Get-ACLShare -ComputerName SERVER01 | Get-ACLShareFileACL -Recurse)
	Saves the file ACLs for all shares on the compuer SERVER01 to the file C:\ACLReports\server01.acl.
#>    
function Export-ACLPermission
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$Path,

        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)]
        [ValidateScript({$_.GetType().FullName -ne 'ACLReportTools.Permission[]' })]
        [ACLReportTools.Permission[]] $InputObject,

        [Switch]$Force
    ) # param

    begin
    {
        if ((Test-Path -Path $Path -PathType Leaf) -and ($force -eq $false))
        {
            Write-Error "The file $Path already exists. Use Force to overwrite it."
            return
        }
        [array]$Output = $null
    } # Begin
    process
    {
        foreach ($Permission in $InputObject)
        {
            $Output += $Permission
        } # Foreach
    } # Process
    end
    {
        try
        {
            $Output | Export-Clixml -Path $Path -Force
        }
        catch
        {
            Write-Error "Unable to export the ACL Permissions file $Path."
        }
    } # End
} # Function Export-ACLPermission


<#
.SYNOPSIS
	Export the ACL Difference Objects that are provided as a file.

.DESCRIPTION 
	This Cmdlet will export an array of provided Permission Difference [ACLReportTools.PermissionDiff] records to a file.
     
.PARAMETER Path
	This is the path to the ACL Permission Diff file. This parameter is required.

.PARAMETER InputObject
	Specifies the Permissions objects to export to th file. Enter a variable that contains the objects or type a command or expression that gets the objects. You can also pipe ACLReportTools.PermissionDiff objects to this cmdlet.

.PARAMETER Force
	Causes the file to be overwritten if it exists.

.EXAMPLE 
	Export-ACLPermissionDiff -Path C:\ACLReports\server01.acr -InputObject $DiffReport
	Saves the ACL Difference objects in the $DiffReport variable to the file C:\ACLReports\server01.acr.  If the file exists it will be overwritten if the Force switch is set.
#>    
function Export-ACLPermissionDiff
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$Path,

        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)]
        [ValidateScript({$_.GetType().FullName -ne 'ACLReportTools.PermissionDiff[]'})]
        [ACLReportTools.PermissionDiff[]]$InputObject,

        [Switch]$Force
    ) # param

    begin
    {
        if ((Test-Path -Path $Path -PathType Leaf) -and ($force -eq $false))
        {
            Write-Error "The file $Path already exists. Use Force to overwrite it."
            return
        }
        [array]$Output = $null
    } # Begin
    process
    {
        foreach ($PermissionDiff in $InputObject)
        {
            $Output += $PermissionDiff
        } # Foreach
    } # Process
    end
    {
        try
        {
            $Output | Export-Clixml -Path $Path -Force
        }
        catch
        {
            Write-Error "Unable to export the ACL Permission Diff $Path."
        }
    } # End
} # Function Export-ACLPermissionDiff


<#
.SYNOPSIS
	Import the a File containing serialized ACL Permission objects that are in a file back into the pipeline.

.DESCRIPTION
	This Cmdlet will load all the ACLs (ACLReportTools.Permission) records from a specified file.
     
.PARAMETER Path
	This is the path to the file containing ACL Permission objects. This parameter is required.

.EXAMPLE 
	Import-ACLPermission -Path C:\ACLReports\server01.acl
	Loads the ACLs in the file C:\ACLReports\server01.acl.
#>    
function Import-ACLPermission
{
    [CmdLetBinding()]
#	[OutputType([ACLReportTools.Permission])]
    param
    (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$Path
    ) # param

    if ((Test-Path -Path $Path -PathType Leaf) -eq $false)
    {
        Write-Error "The file $Path does not exist."
        return
    }

    try
    {
        Import-Clixml -Path $Path
    }
    catch
    { 
        Write-Error "Unable to import the ACL file $Path."
    } # Try
} # Function Import-ACLPermission


<#
.SYNOPSIS
	Import the a File containing serialized ACL Permission Diff objects that are in a file back into the pipeline.

.DESCRIPTION
	This Cmdlet will load all the ACLs (ACLReportTools.PermissionDiff) records from a specified file.
     
.PARAMETER Path
	This is the path to the file containing ACL Permission Diff objects. This parameter is required.

.EXAMPLE 
	Import-ACLPermissionDiff -Path C:\ACLReports\server01.acr
	Loads the ACL Permission Diff objects in the file C:\ACLReports\server01.acr.
#>    
function Import-ACLPermissionDiff
{
    [CmdLetBinding()]
#	[OutputType([ACLReportTools.PermissionDiff])]
    param
    (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$Path
    ) # param

    if ((Test-Path -Path $Path -PathType Leaf) -eq $false)
    {
        Write-Error "The file $Path does not exist."
        return
    }

    try
    {
        Import-Clixml -Path $Path
    }
    catch
    { 
        Write-Error "Unable to import the ACL file $Path."
    } # Try
} # Function Import-ACLPermissionDiff


<#
.SYNOPSIS
	Export the ACL Difference Objects that are provided as an HTML file.

.DESCRIPTION 
	This Cmdlet will export an array of provided Permission Difference [ACLReportTools.PermissionDiff] records to an HTML file for easy viewing and reporting.
     
.PARAMETER Path
	This is the path to the HTML output file. This parameter is required.

.PARAMETER InputObject
	Specifies the Permissions DIff objects to export to the as HTML. Enter a variable that contains the objects or type a command or expression that gets the objects. You can also pipe ACLReportTools.PermissionDiff objects to this cmdlet.

.PARAMETER Force
	Causes the file to be overwritten if it exists.

.PARAMETER Title
	Optional Title text to write into the report.

.EXAMPLE 
	Compare-ACLReports -Baseline (Import-ACLReports -Path c:\ACLReports\server01.acl) -With (Get-ACLReport -ComputerName Server01) | Export-ACLPermissionDiffHTML -Path C:\ACLReports\server01.htm

	Performs a comparison using the Baseline file c:\ACLReports\Server01.acl and the shares on Server01 and outputs ACL Difference Report as an HTML file.
#>    
function Export-ACLPermissionDiffHTML
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$Path,

        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)]
        [ValidateScript({$_.GetType().FullName -ne 'ACLReportTools.PermissionDiff[]'})]
        [ACLReportTools.PermissionDiff[]]$InputObject,

        [Switch]$Force,

		[String]$Title = 'ACL Difference Report'
    ) # param

    begin
    {
        if ((Test-Path -Path $Path -PathType Leaf) -and ($force -eq $false))
        {
            Write-Error "The file $Path already exists. Use Force to overwrite it."
            return
        }
		Set-Content -Path $Path -Value ( CreateHTMLReportHeader -Title $Title ) -Force 
		[String]$LastComputer = ''
		[String]$LastShare = ''
    } # Begin
    process
    {
        foreach ($PermissionDiff in $InputObject)
        {
			if (($ComputerName -ne '') -and ($PermissionDiff.ComputerName -ne $LastComputer))
            {
				$LastComputer = $PermissionDiff.ComputerName
				Add-Content -Path $Path -Value ( CreateHTMLComputerNameLine -ComputerName $PermissionDiff.ComputerName ) -Force 			
			}
			if (($PermissionDiff.Share -ne '') -and ($PermissionDiff.Share -ne $LastShare ))
            {
				$LastShare = $PermissionDiff.Share
				Add-Content -Path $Path -Value ( CreateHTMLShareNameLine -ShareName $PermissionDiff.Share ) -Force 			
			}
			Add-Content -Path $Path -Value ( CreateHTMLPermissionDiffLine -PermissionDiff $PermissionDiff ) -Force 
        } # Foreach
    } # Process
    end
    {
		Add-Content -Path $Path -Value ( CreateHTMLReportFooter ) -Force 
    } # End
} # Function Export-ACLPermissionDiffHTML


##########################################################################################################################################
# Hidden Support CmdLets
##########################################################################################################################################


<#
.SYNOPSIS
	This function creates the a support module containing classes and enums via reflection. It also checks for and loads the
	File System Security PowerShell Module Module (https://gallery.technet.microsoft.com/scriptcenter/1abd77a5-9c0b-4a2b-acef-90dbb2b84e85)

.DESCRIPTION 
	This function creates a .net dynamic module via reflection and adds classes and enums to it that are then used by other functions in this module.
#>
function Initialize-Module
{
    [CmdLetBinding()]
    param (
        [String]$ModuleName = 'ACLReportTools'
    ) # Param

    # Do we need to install the NTFSSecurity Module?
    if ( @( Get-Module -Name NTFSSecurity -ListAvailable ).Count -eq 0)
    {
		try
        {
            Write-Verbose -Message "NTFSSecrity Module needs to be installed."
            Get-PackageProvider -Name NuGet -ForceBootstrap -Force
            Install-Module -Name NTFSSecurity -MinimumVersion 4.0.0 -Force 
            Write-Verbose -Message "NTFSSecrity Module was installed."
        }
        catch
        {
            Throw 'NTFSSecurity Module is not available and could not be installed automatically.  Please download it from https://gallery.technet.microsoft.com/scriptcenter/1abd77a5-9c0b-4a2b-acef-90dbb2b84e85'
        }
	} # If
	Import-Module -Name NTFSSecurity

    $Domain = [AppDomain]::CurrentDomain

    if (($Domain.GetAssemblies() | Where-Object -FilterScript { $_.FullName -eq "$ModuleName, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null" } | Measure-Object).Count-eq 0)
    {
        # Define the module
        $DynAssembly = New-Object Reflection.AssemblyName($ModuleName)
        $AssemblyBuilder = $Domain.DefineDynamicAssembly($DynAssembly, 'Run')
        $ModuleBuilder = $AssemblyBuilder.DefineDynamicModule($ModuleName, $False)

        # Define Permission Difference Enumeration
        $EnumBuilder = $ModuleBuilder.DefineEnum('ACLReportTools.PermissionTypeEnum', 'Public', [Int])
        # Define values of the enum
        $EnumBuilder.DefineLiteral('Not Applicable', [Int]0)
        $EnumBuilder.DefineLiteral('Share', [Int]1)
        $EnumBuilder.DefineLiteral('Folder', [Int]2)
        $EnumBuilder.DefineLiteral('File', [Int]3)
        $PermissionTypeEnumType = $EnumBuilder.CreateType()

        # Define the ACLReportTools.Permission Class
        $Attributes = 'AutoLayout, AnsiClass, Class, Public'
        $TypeBuilder  = $ModuleBuilder.DefineType('ACLReportTools.Permission',$Attributes,[System.Object])
        $TypeBuilder.DefineField('ComputerName', [string], 'Public') | Out-Null
        $TypeBuilder.DefineField('Type', $PermissionTypeEnumType, 'Public') | Out-Null
        $TypeBuilder.DefineField('Share', [string], 'Public') | Out-Null
        $TypeBuilder.DefineField('Path', [string], 'Public') | Out-Null
        $TypeBuilder.DefineField('Owner', [string], 'Public') | Out-Null
        $TypeBuilder.DefineField('Access', [Security2.FileSystemAccessRule2], 'Public') | Out-Null
        $TypeBuilder.DefineField('Inherited', [string], 'Public') | Out-Null
        $TypeBuilder.CreateType() | Out-Null

        # Define the ACLReportTools.Share Class
        $Attributes = 'AutoLayout, AnsiClass, Class, Public'
        $TypeBuilder  = $ModuleBuilder.DefineType('ACLReportTools.Share',$Attributes,[System.Object])
        $TypeBuilder.DefineField('ComputerName', [string], 'Public') | Out-Null
        $TypeBuilder.DefineField('Name', [string], 'Public') | Out-Null
        $TypeBuilder.CreateType() | Out-Null

        # Define Permission Difference Enumeration
        $EnumBuilder = $ModuleBuilder.DefineEnum('ACLReportTools.PermissionDiffEnum', 'Public', [Int])
        # Define values of the enum
        $EnumBuilder.DefineLiteral('No Change', [Int]0)
        $EnumBuilder.DefineLiteral('Computer Added', [Int]1)
        $EnumBuilder.DefineLiteral('Computer Removed', [Int]2)
        $EnumBuilder.DefineLiteral('Share Removed', [Int]3)
        $EnumBuilder.DefineLiteral('Share Added', [Int]4)
        $EnumBuilder.DefineLiteral('Permission Removed', [Int]5)
        $EnumBuilder.DefineLiteral('Permission Added', [Int]6)
        $EnumBuilder.DefineLiteral('Permission Rights Changed', [Int]7)
        $EnumBuilder.DefineLiteral('Permission Access Control Changed', [Int]8)
        $EnumBuilder.DefineLiteral('Owner Changed', [Int]9)
        $EnumBuilder.DefineLiteral('Inheritance Changed', [Int]10)
        $PermissionDiffEnumType = $EnumBuilder.CreateType()

        # Define the ACLReportTools.PermissionDiff Class
        $Attributes = 'AutoLayout, AnsiClass, Class, Public'
        $TypeBuilder  = $ModuleBuilder.DefineType('ACLReportTools.PermissionDiff',$Attributes,[System.Object])
        $TypeBuilder.DefineField('ComputerName', [string], 'Public') | Out-Null
        $TypeBuilder.DefineField('Type', $PermissionTypeEnumType, 'Public') | Out-Null
        $TypeBuilder.DefineField('Share', [string], 'Public') | Out-Null
        $TypeBuilder.DefineField('Path', [string], 'Public') | Out-Null
        $TypeBuilder.DefineField('DiffType', $PermissionDiffEnumType, 'Public') | Out-Null
        $TypeBuilder.DefineField('Difference', [String], 'Public') | Out-Null
        $TypeBuilder.CreateType() | Out-Null
    } # If
} # Function Initialize-Module


<#
.SYNOPSIS
	This function creates an ACLReportTools.Share object and populates it.

.DESCRIPTION 
	This function creates an ACLReportTools.Share object from the class definition in the dynamic module ACLREportsModule and assigns the function parameters to the field values of the object.
#>
function New-ShareObject
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$ShareName
    ) # Param

    $share_object = New-Object -TypeName 'ACLReportTools.Share'
    $share_object.ComputerName = $ComputerName
    $share_object.Name = $ShareName
    return $share_object
} # function New-ShareObject


<#
.SYNOPSIS
	This function creates an ACLReportTools.Permission object and populates it.

.DESCRIPTION 
	This function creates an ACLReportTools.Permission object from the class definition in the dynamic module ACLREportsModule and assigns the function parameters to the field values of the object.
#>
function New-PermissionObject
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [ACLReportTools.PermissionTypeEnum]$Type,
        
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$ComputerName,

        [String]$Path='',
        
        [String]$Share='',

        [String]$Owner='',
        
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [Security2.FileSystemAccessRule2]$Access,
        
        [String]$Inherited=''
    ) # Param

    # Need to correct the $Access objects to ensure the FileSystemRights values correctly converted to string
    # When the "Generic Rights" bits are set: http://msdn.microsoft.com/en-us/library/aa374896%28v=vs.85%29.aspx
    $permission_object = New-Object -TypeName 'ACLReportTools.Permission'
    $permission_object.Type = $Type
    $permission_object.ComputerName = $ComputerName
    $permission_object.Path = $Path
    $permission_object.Share = $Share
    $permission_object.Owner = $Owner
    $permission_object.Access = $Access
    $permission_object.Inherited = $Inherited
    return $permission_object
} # function New-PermissionObject


<#
.SYNOPSIS
	This function creates an ACLReportTools.PermissionDiff object and populates it.

.DESCRIPTION 
	This function creates an ACLReportTools.PermissionDiff object from the class definition in the dynamic module ACLREportsModule and assigns the function parameters to the field values of the object.
#>
function New-PermissionDiffObject
{
    [CmdLetBinding()]
    param
    (
        [ACLReportTools.PermissionTypeEnum]$Type=([ACLReportTools.PermissionTypeEnum]::'Not Applicable'),
        
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$ComputerName,

        [String]$Path='',
        
        [String]$Share='',

        [ACLReportTools.PermissionDiffEnum]$DiffType=([ACLReportTools.PermissionDiffEnum]::'No Change'),
        
        [String]$Difference=''
    ) # Param

    # Need to correct the $Access objects to ensure the FileSystemRights values correctly converted to string
    # When the "Generic Rights" bits are set: http://msdn.microsoft.com/en-us/library/aa374896%28v=vs.85%29.aspx
    $permissiondiff_object = New-Object -TypeName 'ACLReportTools.PermissionDiff'
    $permissiondiff_object.Type = $Type
    $permissiondiff_object.ComputerName = $ComputerName
    $permissiondiff_object.Path = $Path
    $permissiondiff_object.Share = $Share
    $permissiondiff_object.DiffType = $DiffType
    $permissiondiff_object.Difference = $Difference
    return $permissiondiff_object
} # function New-PermissionDiffObject


<#
.SYNOPSIS

.DESCRIPTION 
#>
function Convert-FileSystemAppliesToString
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [String]$InheritanceFlags,
        [Parameter(Mandatory=$true)]
        [String]$PropagationFlags
    ) # Param
    if ($PropagationFlags -eq 'None')
    {
        switch ($InheritanceFlags)
        {
            'None' { return 'This folder only'; break }
            'ContainerInherit, ObjectInherit' { return 'This folder, subfolders and files'; break }
            'ContainerInherit' { return 'This folder and subfolders'; break }
            'ObjectInherit' { return 'This folder and files'; break }
        } # Switch
    }
    else
    {
        switch ($InheritanceFlags)
        {
            'ContainerInherit, ObjectInherit' { return 'Subfolders and files only'; break }
            'ContainerInherit' { return 'Subfolders only'; break }
            'ObjectInherit' { return 'Files only'; break }
        } # Switch
    } # If
    return "Unknown"
} # function Convert-FileSystemAppliesToString


<#
.SYNOPSIS

.DESCRIPTION 
#>
function Convert-AccessToString
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [Object]$Access
    ) # Param
    [string]$rights=$Access.AccessRights
    [string]$controltype=$Access.AccessControlType
    [string]$IdentityReference=$Access.IdentityReference
    [string]$IsInherited=$Access.IsInherited
    [string]$AppliesTo=Convert-FileSystemAppliesToString -InheritanceFlags $Access.InheritanceFlags -PropagationFlags $Access.PropagationFlags
    Return "AccessRights  : $rights`nAccessControlType : $controltype`nIdentityReference : $IdentityReference`nIsInherited       : $IsInherited`nAppliesTo         : $AppliesTo`n"
} # function Convert-AccessToString


<#
.SYNOPSIS

.DESCRIPTION 
#>
function Convert-ACEToString
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [Object]$ACE
    ) # Param
    [string]$path=$ACE.path
    [string]$owner=$ACE.owner
    [string]$acccessstring=Convert-ACEToString($ACE.access)
    Return "Path              : $path`nOwner             : $owner`n$acccessstring"
} # function Convert-ACEToString


Function CreateHTMLReportHeader
{
    param
    (
        [Parameter(Mandatory=$true)]
        [String]$Title
    ) # Param
    return $Script:Html_Header -f $Title
} # Function CreateHTMLReportHeader


Function CreateHTMLReportFooter
{
    return $Script:Html_Footer
} # Function CreateHTMLReportFooter


Function CreateHTMLComputerNameLine
{
    param
    (
        [Parameter(Mandatory=$true)]
        [String]$ComputerName
    ) # Param
	return $Script:Html_ComputerName -f $ComputerName
} # Function CreateHTMLComputerNameLine


Function CreateHTMLShareNameLine
{
    param
    (
        [Parameter(Mandatory=$true)]
        [String]$ShareName
    ) # Param
	return $Script:Html_ShareName -f $ShareName
} # Function CreateHTMLShareNameLine


Function CreateHTMLPermissionDiffLine
{
    param
    (
        [Parameter(Mandatory=$true)]
        [ACLReportTools.PermissionDiff]$PermissionDiff
    ) # Param

    # This function takes a Permission Diff object and formats it as HTML for a report.
	[string]$label = $PermissionDiff.Type.ToString()
	[string]$class = $PermissionDiff.DiffType.ToString().ToLower() -replace " ",""
	[string]$html = $PermissionDiff.Difference
	return $Script:Html_DifferenceLine -f $Label,$Class,$Html
} # Function CreateHTMLPermissionDiffLine


# Ensure all the custom classes are loaded in available
Initialize-Module


# Export the Module Cmdlets
Export-ModuleMember -Function `
    New-ACLShareReport,`
    New-ACLPathFileReport, `
	Import-ACLReport,`
    Export-ACLReport,`
    Import-ACLDiffReport,`
    Export-ACLDiffReport,`
	Compare-ACLReports,`
	Get-ACLShare,`
	Get-ACLShareACL,`
    Get-ACLPathFileACL,`
    Get-ACLShareFileACL,`
	Import-ACLPermission,`
    Export-ACLPermission,`
    Import-ACLPermissionDiff,`
    Export-ACLPermissionDiff,`
	Export-ACLPermissionDiffHTML

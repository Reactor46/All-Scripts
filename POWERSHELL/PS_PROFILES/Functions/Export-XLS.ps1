function Export-Xls 
{ 
 
<# 
.SYNOPSIS 
Export to Excel file. 
 
.DESCRIPTION 
Export to Excel file. Since Excel files can have multiple worksheets, you can specify the name of the Excel file and worksheet. Exports to a worksheet named "Sheet" by default. 
 
.PARAMETER Path 
Specifies the path to the Excel file to export. 
Note: The path must contain an extension for spreadsheets, such as .xls, .xlsx, .xlsm, .xml, and .ods 
 
.PARAMETER Worksheet 
Specifies the name of the worksheet where the data is exported. The default is "Sheet". 
Note: If a worksheet already exists with the given name, no error occurs. The name will be appended with (2), or (3), or (4), etc. 
 
.PARAMETER InputObject 
Specifies the objects to export. You can also pipe objects to Export-Xls. 
 
.PARAMETER Append 
Append the exported data to a new worksheet in the excel file. 
If you Append to a spreadsheet that does not allow more than one worksheet, the new data will not be saved. 
Note: For this function, -Append is not considered clobbering the file, but modifying the file, so -Append and -NoClobber do not conflict with each other. 
 
.PARAMETER NoClobber 
Do not overwrite the file. 
Use -Append if you want to add a worksheet to the excel file, but leave the others intact. 
Note: For this function, -Append is not considered clobbering the file, but modifying the file, so -Append and -NoClobber do not conflict with each other. 
 
.PARAMETER NoTypeInformation 
Omits the type information. 
 
.INPUTS 
System.Management.Automation.PSObject 
 
.OUTPUTS 
System.String 
This is a CSV list, which is then exported to a csv file, which is then converted to an Excel file. 
 
.EXAMPLE 
Get-Process | Export-Xls ".\export.xlsx" -Worksheet "Sheet1" 
Export the output of Get-Process to Worksheet "Sheet1" of export.xlsx 
Note: export.xlsx is overwritten. 
 
.EXAMPLE 
Get-Process | Export-Xls ".\export.xlsx" -Worksheet "Sheet2" -NoTypeInformation 
Export the output of Get-Process to Worksheet "Sheet2" of export.xlsx with no type information 
Note: export.xlsx is overwritten. 
 
.EXAMPLE 
Get-Process | Export-Xls ".\export.xlsx" -Worksheet "Sheet3" -Append 
Export output of Get-Process to Worksheet "Sheet3" and Append it to export.xlsx 
Note: export.xlsx is modified. 
 
.EXAMPLE 
Get-Process | Export-Xls ".\export.xlsx" -Worksheet "Sheet4" -NoClobber 
Export output of Get-Process to Worksheet "Sheet4" and create export.xlsx if it doesn't exist. 
Note: export.xlsx is created. If export.xlsx already exist, the function terminates with an error. 
 
.EXAMPLE 
(Get-Alias s*), (Get-Alias g*) | Export-Xls ".\export.xlsx" -Worksheet "Alias" 
Export Aliases that start with s and g to Worksheet "Alias" of export.xlsx 
Note: See next example for possible problems when doing something like this 
 
.EXAMPLE 
(Get-Alias), (Get-Process) | Export-Xls ".\export.xlsx" -Worksheet "Alias and Process" 
Export the result of Get-Command and Get-Process to Worksheet "Alias and Process" of export.xlsx 
Note: Since Get-Alias and Get-Process do not return objects with the same properties, not all information is recorded. 
 
.LINK 
Export-Xls 
http://gallery.technet.microsoft.com/scriptcenter/d41565f1-37ef-43cb-9462-a08cd5a610e2 
Import-Xls 
http://gallery.technet.microsoft.com/scriptcenter/17bcabe7-322a-43d3-9a27-f3f96618c74b 
Import-Csv 
Export-Csv 
 
.NOTES 
Author: Francis de la Cerna 
Created: 2011-03-27 
Modified: 2011-04-09 
#Requires –Version 2.0 
#> 
 
    [CmdletBinding(SupportsShouldProcess=$true)] 
 
    Param( 
        [parameter(mandatory=$true, position=1)] 
        $Path, 
     
        [parameter(mandatory=$false, position=2)] 
        $Worksheet = "Sheet", 
     
        [parameter( 
            mandatory=$true,  
            ValueFromPipeline=$true, 
            ValueFromPipelineByPropertyName=$true)] 
        [psobject[]] 
        $InputObject, 
     
        [parameter(mandatory=$false)] 
        [switch] 
        $Append, 
     
        [parameter(mandatory=$false)] 
        [switch] 
        $NoClobber, 
 
        [parameter(mandatory=$false)] 
        [switch] 
        $NoTypeInformation, 
 
        [parameter(mandatory=$false)] 
        [switch] 
        $Force 
    ) 
     
    Begin 
    { 
        # WhatIf, Confirm, Verbose 
        # Probably not the way to do it, but this function runs all or nothing 
        # so, exit each block (Begin, Process, End) if shouldProcesss is false. 
        # Disabled confirmations on operations on temporary files, but enabled 
        # verbose messages. 
        #  
        $shouldProcess = $Force -or $psCmdlet.ShouldProcess($Path); 
         
        if (-not $shouldProcess) { return; } 
         
        function GetTempFileName($extension) 
        { 
            $temp = [io.path]::GetTempFileName(); 
            $params = @{ 
                Path = $temp; 
                Destination = $temp + $extension; 
                Confirm = $false; 
                Verbose = $VerbosePreference; 
            } 
            Move-Item @params; 
            $temp += $extension; 
            return $temp; 
        } 
         
        # check extension of $Path to see what excel format to export to 
        # since an extension like .xls can have multiple formats, this 
        # will need to be changed 
        # 
        $xlFileFormats = @{ 
            # single worksheet formats 
            '.csv'  = 6;        # 6, 22, 23, 24 
            '.dbf'  = 11;       # 7, 8, 11 
            '.dif'  = 9;        #  
            '.prn'  = 36;       #  
            '.slk'  = 2;        # 2, 10 
            '.wk1'  = 31;       # 5, 30, 31 
            '.wk3'  = 32;       # 15, 32 
            '.wk4'  = 38;       #  
            '.wks'  = 4;        #  
            '.xlw'  = 35;       #  
             
            # multiple worksheet formats 
            '.xls'  = -4143;    # -4143, 1, 16, 18, 29, 33, 39, 43 
            '.xlsb' = 50;       # 
            '.xlsm' = 52;       # 
            '.xlsx' = 51;       # 
            '.xml'  = 46;       # 
            '.ods'  = 60;       # 
        } 
         
        $ext = [io.path]::GetExtension($Path).toLower(); 
        if ($xlFileFormats.Keys -notcontains $ext) { 
            $msg = "Error: $Path has unknown extension. Try "; 
            foreach ($extension in ($xlFileFormats.Keys | sort)) { 
                $msg += "$extension "; 
            } 
            Throw "$msg"; 
        } 
         
        # get full path 
        # 
        if (-not [io.path]::IsPathRooted($Path)) { 
            $fswd = $psCmdlet.CurrentProviderLocation("FileSystem"); 
            $Path = Join-Path -Path $fswd -ChildPath $Path; 
        } 
         
        $Path = [io.path]::GetFullPath($Path); 
 
        $obj = New-Object System.Collections.ArrayList; 
    } 
 
    Process 
    { 
        if (-not $shouldProcess) { return; } 
 
        $InputObject | ForEach-Object{ $obj.Add($_) | Out-Null; } 
    } 
 
    End 
    { 
        if (-not $shouldProcess) { return; } 
         
        $xl = New-Object -ComObject Excel.Application; 
        $xl.DisplayAlerts = $false; 
        $xl.Visible = $false; 
         
        # create temporary .csv file from all $InputObject 
        # 
        $csvTemp = GetTempFileName(".csv"); 
        $obj | Export-Csv -Path $csvTemp -Force -NoType:$NoTypeInformation -Confirm:$false; 
         
        # create a temporary excel file from the temporary .csv file 
        # 
        $xlsTemp = GetTempFileName($ext); 
        $wb = $xl.Workbooks.Add($csvTemp); 
        $ws = $wb.Worksheets.Item(1); 
        $ws.Name = $Worksheet; 
        $wb.SaveAs($xlsTemp, $xlFileFormats[$ext]); 
        $xlsTempSaved = $?; 
        $wb.Close(); 
        Remove-Variable -Name ('ws', 'wb') -Confirm:$false; 
         
        if ($xlsTempSaved) { 
            # decide how to export based on switches and $Path 
            # 
            $fileExist = Test-Path $Path; 
            $createFile = -not $fileExist; 
            $appendFile = $fileExist -and $Append; 
            $clobberFile = $fileExist -and (-not $appendFile) -and (-not $NoClobber); 
            $needNewFile = $fileExist -and (-not $appendFile) -and $NoClobber; 
         
            if ($appendFile) { 
                $wbDst = $xl.Workbooks.Open($Path); 
                $wbSrc = $xl.Workbooks.Open($xlsTemp); 
                $wsDst = $wbDst.Worksheets.Item($wbDst.Worksheets.Count); 
                $wsSrc = $wbSrc.Worksheets.Item(1); 
                $wsSrc.Name = $Worksheet; 
                $wsSrc.Copy($wsDst); 
                $wsDst.Move($wbDst.Worksheets.Item($wbDst.Worksheets.Count-1)); 
                $wbDst.Worksheets.Item(1).Select(); 
                $wbSrc.Close($false); 
                $wbDst.Close($true); 
                Remove-Variable -Name ('wsSrc', 'wbSrc') -Confirm:$false; 
                Remove-Variable -Name ('wsDst', 'wbDst') -Confirm:$false; 
            } elseif ($createFile -or $clobberFile) { 
                Copy-Item $xlsTemp -Destination $Path -Force -Confirm:$false; 
            } elseif ($needNewFile) { 
                Write-Error "The file '$Path' already exists." -Category ResourceExists; 
            } else { 
                Write-Error "Something was wrong with my logic."; 
            } 
        } 
         
        # clean up 
        # 
        $xl.Quit(); 
        Remove-Variable -name xl -Confirm:$false; 
        Remove-Item $xlsTemp -Confirm:$false -Verbose:$VerbosePreference; 
        Remove-Item $csvTemp -Confirm:$false -Verbose:$VerbosePreference; 
        [gc]::Collect(); 
    } 
} 
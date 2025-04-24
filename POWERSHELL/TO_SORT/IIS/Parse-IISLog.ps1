Function Parse-IISLog {

    <#    
    .SYNOPSIS    
        Parse IIS Logs (IIS 6, 7, 7.5) 
    .DESCRIPTION  
        Parses through each entry of an IIS log and separates each value into an data tabele that can then be sorted, selected, or exported to csv.
    .PARAMETER LogFile 
        Path to the log file to be parsed.
    .NOTES    
        Name: Parse-IISLog 
        Author: Michael Introini
        DateCreated: 02/27/2015        
    .EXAMPLE    
        Parse-IISLog -Log "C:\Temp\Log.txt"
        
        date            : 2015-02-27
        time            : 01:40:11
        s-sitename      : W3SVC1
        s-ip            : 10.10.10.10
        cs-method       : GET
        cs-uri-stem     : /url/
        cs-uri-query    : /query/
        s-port          : 443
        cs-username     : -
        c-ip            : 10.5.5.5
        cs-version      : HTTP/1.1
        cs(User-Agent)  : -
        cs(Referer)     : -
        sc-status       : 200
        sc-substatus    : 0
        sc-win32-status : 3  
      
    Description  
    ------------  
    Returns a data table with all of entries from the IIS Log.
   
    .EXAMPLE
        Parse-IISLog -Log "C:\Temp\Log.txt"  | Select Sc-Status
     
        sc-status                                                                                                                                                                               
        ---------                                                                                                                                                                               
        200      
      
    Description  
    ------------  
    Returns a data table with all of entries from the IIS Log.

    #> 

	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$True,
		ValueFromPipeline=$True)]
		[string[]]$Path
	)
	BEGIN {
        $table = New-Object system.data.datatable
        $fieldsString = gc $Path| select -First 5 | ? {$_ -Like "#[F]*"}
        $fields = $FieldsString.substring(9).split(' ')
        $fieldsCount = $fields.count - 1
    }
	PROCESS {
        for($i=0;$i -lt $fieldsCount;$i++) {
            $table.Columns.Add($fields[$i]) | Out-Null   
        }
        $content = gc $Path | where {$_ -notLike "#[D,S,V,F]*" } | % {
            $row = $table.NewRow()
            for($i=0;$i -lt $fieldsCount;$i++) {
                $row[$i] = $_.split(' ')[$i]
            }
            $table.rows.add($row)
        }
    }
    END {
        Return $table
    }
}
## Begin Out-DataTable
Function Out-DataTable {
    [CmdletBinding()]
    param([Parameter(Position=0, Mandatory=$true, ValueFromPipeline = $true)] [PSObject[]]$InputObject)

    Begin
    {
    Function Get-Type {
        param($type)

        $types = @(
        'System.Boolean',
        'System.Byte[]',
        'System.Byte',
        'System.Char',
        'System.Datetime',
        'System.Decimal',
        'System.Double',
        'System.Guid',
        'System.Int16',
        'System.Int32',
        'System.Int64',
        'System.Single',
        'System.UInt16',
        'System.UInt32',
        'System.UInt64')

        if ( $types -contains $type ) {
            Write-Output "$type"
        }
        else {
            Write-Output 'System.String'
        
        }
    } #Get-Type
        $dt = new-object Data.datatable  
        $First = $true 
    }
    Process
    {
        foreach ($object in $InputObject)
        {
            $DR = $DT.NewRow()  
            foreach($property in $object.PsObject.get_properties())
            {  
                if ($first)
                {  
                    $Col =  new-object Data.DataColumn  
                    $Col.ColumnName = $property.Name.ToString()  
                    if ($property.value)
                    {
                        if ($property.value -isnot [System.DBNull]) {
                            $Col.DataType = [System.Type]::GetType("$(Get-Type $property.TypeNameOfValue)")
                         }
                    }
                    $DT.Columns.Add($Col)
                }  
                if ($property.Gettype().IsArray) {
                    $DR.Item($property.Name) =$property.value | ConvertTo-XML -AS String -NoTypeInformation -Depth 1
                }  
               else {
                    If ($Property.Value) {
                        $DR.Item($Property.Name) = $Property.Value
                    } Else {
                        $DR.Item($Property.Name)=[DBNull]::Value
                    }
                }
            }  
            $DT.Rows.Add($DR)  
            $First = $false
        }
    } 
     
    End
    {
        Write-Output @(,($dt))
    }

}
## End Out-DataTable
## Begin Out-Excel
Function Out-Excel{
<#
	.SYNOPSIS
	.DESCRIPTION
	.PARAMETER Property
	.PARAMETER Raw
	.NOTES
    	Original Script: http://pathologicalscripter.wordpress.com/out-excel/
	
		TODO:
			Parameter to change color of header
			Parameter to activate background color on Odd unit
			Add TRY/CATCH
			Validate Excel first is present
#>
	[CmdletBinding()]
	PARAM ([string[]]$property, [switch]$raw)
	
	BEGIN
	{
		# start Excel and open a new workbook
		$Excel = New-Object -Com Excel.Application
		$Excel.visible = $True
		$Excel = $Excel.Workbooks.Add()
		$Sheet = $Excel.Worksheets.Item(1)
		# initialize our row counter and create an empty hashtable
		# which will hold our column headers
		$Row = 1
		$HeaderHash = @{ }
	}
	
	PROCESS
	{
		if ($_ -eq $null) { return }
		if ($Row -eq 1)
		{
			# when we see the first object, we need to build our header table
			if (-not $property)
			{
				# if we haven’t been provided a list of properties,
				# we’ll build one from the object’s properties
				$property = @()
				if ($raw)
				{
					$_.properties.PropertyNames | %{ $property += @($_) }
				}
				else
				{
					$_.PsObject.get_properties() | % { $property += @($_.Name.ToString()) }
				}
			}
			$Column = 1
			foreach ($header in $property)
			{
				# iterate through the property list and load the headers into the first row
				# also build a hash table so we can retrieve the correct column number
				# when we process each object
				$HeaderHash[$header] = $Column
				$Sheet.Cells.Item($Row, $Column) = $header.toupper()
				$Column++
			}
			# set some formatting values for the first row
			$WorkBook = $Sheet.UsedRange
			$WorkBook.Interior.ColorIndex = 19
			$WorkBook.Font.ColorIndex = 11
			$WorkBook.Font.Bold = $True
			$WorkBook.HorizontalAlignment = -4108
		}
		$Row++
		foreach ($header in $property)
		{
			# now for each object we can just enumerate the headers, find the matching property
			# and load the data into the correct cell in the current row.
			# this way we don’t have to worry about missing properties
			# or the “ordering” of the properties
			if ($thisColumn = $HeaderHash[$header])
			{
				if ($raw)
				{
					$Sheet.Cells.Item($Row, $thisColumn) = [string]$_.properties.$header
				}
				else
				{
					$Sheet.Cells.Item($Row, $thisColumn) = [string]$_.$header
				}
			}
		}
	}
	
	end
	{
		# now just resize the columns and we’re finished
		if ($Row -gt 1) { [void]$WorkBook.EntireColumn.AutoFit() }
	}
}
## End Out-Excel
## Begin WriteWordLine
Function WriteWordLine{
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
#updated 27-Mar-2014 to include font name, font size, italics and bold options
#update 5-May-2016 by Michael B. Smith
	Param(
		[int] $style       = 0, 
		[int] $tabs        = 0, 
		[string] $name     = '', 
		[string] $value    = '', 
		[string] $fontName = $null,
		[int] $fontSize    = 0,
		[bool] $italics    = $false,
		[bool] $boldface   = $false,
		[Switch] $nonewline
	)
	
	#Build output style
	[string]$output = ''
	Switch ($style)
	{
		0 {$Script:Selection.Style = $myHash.Word_NoSpacing}
		1 {$Script:Selection.Style = $myHash.Word_Heading1}
		2 {$Script:Selection.Style = $myHash.Word_Heading2}
		3 {$Script:Selection.Style = $myHash.Word_Heading3}
		4 {$Script:Selection.Style = $myHash.Word_Heading4}
		Default {$Script:Selection.Style = $myHash.Word_NoSpacing}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t" 
		$tabs-- 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Script:Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Script:Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Script:Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Script:Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Script:Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If( !$nonewline )
	{
		$Script:Selection.TypeParagraph()
	}
}
## End WriteWordLine
## Begin SetWordCellFormat
Function SetWordCellFormat{
	[CmdletBinding(DefaultParameterSetName='Collection')]
	Param (
		# Word COM object cell collection reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName='Collection', Position=0)] [ValidateNotNullOrEmpty()] $Collection,
		# Word COM object individual cell reference
		[Parameter(Mandatory=$true, ParameterSetName='Cell', Position=0)] [ValidateNotNullOrEmpty()] $Cell,
		# Hashtable of cell co-ordinates
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
		# Word COM object table reference
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=1)] [ValidateNotNullOrEmpty()] $Table,
		# Font name
		[Parameter()] [AllowNull()] [string] $Font = $null,
		# Font color
		[Parameter()] [AllowNull()] $Color = $null,
		# Font size
		[Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
		# Cell background color
		[Parameter()] [AllowNull()] $BackgroundColor = $null,
		# Force solid background color
		[Switch] $Solid,
		[Switch] $Bold,
		[Switch] $Italic,
		[Switch] $Underline
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
	}

	Process 
	{
		Switch ($PSCmdlet.ParameterSetName) 
		{
			'Collection' 
			{
				ForEach($Cell in $Collection) 
				{
					If($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
					If($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end ForEach
			} # end Collection
			'Cell' 
			{
				If($Bold) { $Cell.Range.Font.Bold = $true; }
				If($Italic) { $Cell.Range.Font.Italic = $true; }
				If($Underline) { $Cell.Range.Font.Underline = 1; }
				If($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
				If($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
				If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
				If($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
				If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
			} # end Cell
			'Hashtable' 
			{
				ForEach($Coordinate in $Coordinates) 
				{
					$Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column);
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
					If($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				}
			} # end Hashtable
		} # end Switch
	} # end process
}
## End SetWordCellFormat
## Begin SetWordHashTable
Function SetWordHashTable{
	Param([string]$CultureCode)

	#optimized by Michael B. SMith
	
	# DE and FR translations for Word 2010 by Vladimir Radojevic
	# Vladimir.Radojevic@Commerzreal.com

	# DA translations for Word 2010 by Thomas Daugaard
	# Citrix Infrastructure Specialist at edgemo A/S

	# CA translations by Javier Sanchez 
	# CEO & Founder 101 Consulting

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese
	
	[string]$toc = $(
		Switch ($CultureCode)
		{
			'ca-'	{ 'Taula automática 2'; Break }
			'da-'	{ 'Automatisk tabel 2'; Break }
			'de-'	{ 'Automatische Tabelle 2'; Break }
			'en-'	{ 'Automatic Table 2'; Break }
			'es-'	{ 'Tabla automática 2'; Break }
			'fi-'	{ 'Automaattinen taulukko 2'; Break }
			'fr-'	{ 'Table automatique 2'; Break } #changed 13-feb-2017 david roquier and samuel legrand
			'nb-'	{ 'Automatisk tabell 2'; Break }
			'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
			'pt-'	{ 'Sumário Automático 2'; Break }
			'sv-'	{ 'Automatisk innehållsförteckning2'; Break }
			'zh-'	{ '???? 2'; Break }
		}
	)

	$Script:myHash                      = @{}
	$Script:myHash.Word_TableOfContents = $toc
	$Script:myHash.Word_NoSpacing       = $wdStyleNoSpacing
	$Script:myHash.Word_Heading1        = $wdStyleheading1
	$Script:myHash.Word_Heading2        = $wdStyleheading2
	$Script:myHash.Word_Heading3        = $wdStyleheading3
	$Script:myHash.Word_Heading4        = $wdStyleheading4
	$Script:myHash.Word_TableGrid       = $wdTableGrid
}
## End SetWordHashTable
## Begin SetWordTableAlternateRowColor
Function SetWordTableAlternateRowColor{
	[CmdletBinding()]
	Param (
		# Word COM object table reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)] [ValidateNotNullOrEmpty()] $Table,
		# Alternate row background color
		[Parameter(Mandatory=$true, Position=1)] [ValidateNotNull()] [int] $BackgroundColor,
		# Alternate row starting seed
		[Parameter(ValueFromPipelineByPropertyName=$true, Position=2)] [ValidateSet('First','Second')] [string] $Seed = 'First'
	)

	Process 
	{
		$StartDateTime = Get-Date;
		Write-Debug ("{0}: `t`tSetting alternate table row colors.." -f $StartDateTime);

		## Determine the row seed (only really need to check for 'Second' and default to 'First' otherwise
		If($Seed.ToLower() -eq 'second') 
		{ 
			$StartRowIndex = 2; 
		}
		Else 
		{ 
			$StartRowIndex = 1; 
		}

		For($AlternateRowIndex = $StartRowIndex; $AlternateRowIndex -lt $Table.Rows.Count; $AlternateRowIndex += 2) 
		{ 
			$Table.Rows.Item($AlternateRowIndex).Shading.BackgroundPatternColor = $BackgroundColor;
		}

		## I've put verbose calls in here we can see how expensive this Functionality actually is.
		$EndDateTime = Get-Date;
		$ExecutionTime = New-TimeSpan -Start $StartDateTime -End $EndDateTime;
		Write-Debug ("{0}: `t`tDone setting alternate row style color in '{1}' seconds" -f $EndDateTime, $ExecutionTime.TotalSeconds);
	}
}
## End SetWordTableAlternateRowColor
## Begin Stop-WinWord
Function Stop-WinWord{
	Write-Debug "***Enter Stop-WinWord"
	
	## determine our login session
	$proc = Get-Process -PID $PID
	If( $null -eq $proc )
	{
		throw "Stop-WinWord: Cannot find process $PID"
	}
	
	$SessionID = $proc.SessionId
	If( $null -eq $SessionID )
	{
		Write-Debug "Stop-WinWord: SessionId on $PID is null"
		throw "Can't find a session for pid $PID"
	}

	If( 0 -eq $SessionID )
	{
		Write-Debug "Stop-WinWord: SessionId is 0 -- that is a bug"
		throw "SessionId is zero for pid $PID"
	}
	
	#Find out if winword is running in our session
	try 
	{
		$wordProc = Get-Process 'WinWord' -ErrorAction SilentlyContinue
	}
	catch
	{
		Write-Debug "***Exit Stop-WinWord: no WinWord tasks are running #1"
		Return ## not running
	}

	If( !$wordproc )
	{
		Write-Debug "***Exit Stop-WinWord: no WinWord tasks are running #2"
		Return ## WinWord is not running in ANY session
	}
	
	$wordrunning = $wordProc |? { $_.SessionId -eq $SessionID }
	If( !$wordrunning )
	{
		Write-Debug "***Exit Stop-WinWord: wordRunning eq null"
		Return ## not running in the current session
	}
	If( $wordrunning -is [Array] )
	{
		Write-Debug "***Exit Stop-WinWord: wordRunning is an array, elements=$($wordrunning.Count)"
		throw "Multiple Word processes are running in session $SessionID"
	}

	## it is possible for the below to throw a fault if Winword stops before it is executed.
	Stop-Process -Id $wordrunning.Id -ErrorAction SilentlyContinue
	Write-Debug "***Exit Stop-WinWord: sent Stop-Process to $($wordrunning.Id)"
}
## End Stop-WinWord
## Begin AddWordTable
Function AddWordTable{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string[]] $Columns = $null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string[]] $Headers = $null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines=$false,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$true)] [int] $Format = '-231'
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Columns -eq $null) -and ($Headers -ne $null)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $null;
		}
		ElseIf(($Columns -ne $null) -and ($Headers -ne $null)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end ElseIf
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
        [System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($Columns -eq $null) 
				{
					## Build the available columns from all availble PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) 
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Headers -ne $null) 
					{
                        $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Object in $CustomObject) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
                    $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach
				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Columns -eq $null) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Headers -ne $null) 
					{ 
                        $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
                        $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Hash in $Hashtable) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
                    $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach

				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end Switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date): `t`tBuilding table range");
		$WordRange = $Script:Doc.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If($Format -ge 0) 
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $true);
			$ConvertToTableArguments.Add("ApplyShading", $true);
			$ConvertToTableArguments.Add("ApplyFont", $true);
			$ConvertToTableArguments.Add("ApplyColor", $true);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $true); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $true);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $true);
			$ConvertToTableArguments.Add("ApplyLastColumn", $true);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$null,                                          # Modifiers
			$null,                                          # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting)
		If($Format -lt 0) 
		{
			Write-Debug ("$(Get-Date): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		#the next line causes the heading row to flow across page breaks
		$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}
## End AddWordTable
## Begin Export-Xls
Function Export-Xls{ 
 
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
Note: For this Function, -Append is not considered clobbering the file, but modifying the file, so -Append and -NoClobber do not conflict with each other. 
 
.PARAMETER NoClobber 
Do not overwrite the file. 
Use -Append if you want to add a worksheet to the excel file, but leave the others intact. 
Note: For this Function, -Append is not considered clobbering the file, but modifying the file, so -Append and -NoClobber do not conflict with each other. 
 
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
Note: export.xlsx is created. If export.xlsx already exist, the Function terminates with an error. 
 
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
        # Probably not the way to do it, but this Function runs all or nothing 
        # so, exit each block (Begin, Process, End) if shouldProcesss is false. 
        # Disabled confirmations on operations on temporary files, but enabled 
        # verbose messages. 
        #  
        $shouldProcess = $Force -or $psCmdlet.ShouldProcess($Path); 
         
        if (-not $shouldProcess) { return; } 
         
        Function GetTempFileName($extension) 
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
## End Export-Xls
## Begin FindWordDocumentEnd
Function FindWordDocumentEnd{
	#Return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}
## End FindWordDocumentEnd
## Begin MergeCSV
Function MergeCSV {
  $Date = Get-Date -Format "d.MMM.yyyy"
  $path = "C:\LazyWinAdmin\Logs\Server-Apps\CSV\*"
  $csvs = Get-ChildItem $path -Include *.csv
  $y = $csvs.Count
  Write-Host "Detected the following CSV files: ($y)"
  foreach ($csv in $csvs) {
    Write-Host " "$csv.Name
  }
  $outputfilename = "Final Registry Results"
  Write-Host Creating: $outputfilename
  $excelapp = New-Object -ComObject Excel.Application
  $excelapp.SheetsInNewWorkbook = $csvs.Count
  $xlsx = $excelapp.Workbooks.Add()
  $sheet = 1
  foreach ($csv in $csvs) {
    $row = 1
    $column = 1
    $worksheet = $xlsx.Worksheets.Item($sheet)
    $worksheet.Name = $csv.Name
    $file = (Get-Content $csv)
    foreach ($line in $file) {
      $linecontents = $line -split ',(?!\s*\w+")'
      foreach ($cell in $linecontents) {
        $worksheet.Cells.Item($row,$column) = $cell
        $column++
      }
      $column = 1
      $row++
    }
    $sheet++
  }
  $output = "C:\LazyWinAdmin\Logs\Server-Apps\$Date\Results.Xlsx"
  $xlsx.SaveAs($output)
  $excelapp.Quit()
}
## End MergeCSV
## Begin Export-Xlsx
Function Export-Xlsx {
<#
.SYNOPSIS
Exports data to an Excel workbook
.DESCRIPTION
Exports data to an Excel workbook and applies cosmetics. 
Optionally add a title, autofilter, autofit and a chart.
Allows for export to .xls and .xlsx format. If .xlsx is
specified but not available (Excel 2003) the data will
be exported to .xls.
.NOTES
Author:  Gilbert van Griensven
Based on
https://www.lucd.info/2010/05/29/beyond-export-csv-export-xls/
.PARAMETER InputData
The data to be exported to Excel
.PARAMETER Path
The path of the Excel file. 
Defaults to %HomeDrive%\Export.xlsx.
.PARAMETER WorksheetName
The name of the worksheet. Defaults to filename
in $Path without extension.
.PARAMETER ChartType
Name of an Excel chart to be added.
.PARAMETER Title
Adds a title to the worksheet.
.PARAMETER SheetPosition
Adds the worksheet either to the 'begin' or 'end' of
the Excel file. This parameter is ignored when creating
a new Excel file.
.PARAMETER ChartOnNewSheet
Adds a chart to a new worksheet instead of to the
worksheet containing data. The Chart will be placed after
the sheet containing data. Only works when parameter
ChartType is used.
.PARAMETER AppendWorksheet
Appends a worksheet to an existing Excel file.
This parameter is ignored when creating a new Excel file.
.PARAMETER Borders
Adds borders to all cells. Defaults to True.
.PARAMETER HeaderColor
Applies background color to the header row. 
Defaults to True.
.PARAMETER AutoFit
Apply autofit to columns. Defaults to True.
.PARAMETER AutoFilter
Apply autofilter. Defaults to True.
.PARAMETER PassThrough
When enabled returns file object of the generated file.
.PARAMETER Force
Overwrites existing Excel sheet. When this switch is
not used but the Excel file already exists, a new file
with datestamp will be generated. This switch is ignored
when using the AppendWorksheet switch.
.EXAMPLE
Get-Process | Export-Xlsx D:\Data\ProcessList.xlsx
.EXAMPLE
Get-ADuser -Filter {enabled -ne $True} | 
Select-Object Name,Surname,GivenName,DistinguishedName | 
Export-Xlsx -Path 'D:\Data\Disabled Users.xlsx' -Title 'Disabled users of Contoso.com'
.EXAMPLE
Get-Process | Sort-Object CPU -Descending | 
Export-Xlsx -Path D:\Data\Processes_by_CPU.xlsx
.EXAMPLE
Export-Xlsx (Get-Process) -AutoFilter:$False -PassThrough |
Invoke-Item
#>
[CmdletBinding()]
Param (
[Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$True)]
[ValidateNotNullOrEmpty()]
$InputData,
[Parameter(Position=1)]
[ValidateScript({
$ReqExt = [System.IO.Path]::GetExtension($_)
(          $ReqExt -eq ".xls") -or
(          $ReqExt -eq ".xlsx")
})]
## End Export-Xlsx
$Path = (Join-Path $env:HomeDrive "Export.xlsx"),
[Parameter(Position=2)] $WorksheetName = [System.IO.Path]::GetFileNameWithoutExtension($Path),
[Parameter(Position=3)]
[ValidateSet("xl3DArea","xl3DAreaStacked","xl3DAreaStacked100","xl3DBarClustered",
"xl3DBarStacked","xl3DBarStacked100","xl3DColumn","xl3DColumnClustered",
"xl3DColumnStacked","xl3DColumnStacked100","xl3DLine","xl3DPie",
"xl3DPieExploded","xlArea","xlAreaStacked","xlAreaStacked100",
"xlBarClustered","xlBarOfPie","xlBarStacked","xlBarStacked100",
"xlBubble","xlBubble3DEffect","xlColumnClustered","xlColumnStacked",
"xlColumnStacked100","xlConeBarClustered","xlConeBarStacked","xlConeBarStacked100",
"xlConeCol","xlConeColClustered","xlConeColStacked","xlConeColStacked100",
"xlCylinderBarClustered","xlCylinderBarStacked","xlCylinderBarStacked100","xlCylinderCol",
"xlCylinderColClustered","xlCylinderColStacked","xlCylinderColStacked100","xlDoughnut",
"xlDoughnutExploded","xlLine","xlLineMarkers","xlLineMarkersStacked",
"xlLineMarkersStacked100","xlLineStacked","xlLineStacked100","xlPie",
"xlPieExploded","xlPieOfPie","xlPyramidBarClustered","xlPyramidBarStacked",
"xlPyramidBarStacked100","xlPyramidCol","xlPyramidColClustered","xlPyramidColStacked",
"xlPyramidColStacked100","xlRadar","xlRadarFilled","xlRadarMarkers",
"xlStockHLC","xlStockOHLC","xlStockVHLC","xlStockVOHLC",
"xlSurface","xlSurfaceTopView","xlSurfaceTopViewWireframe","xlSurfaceWireframe",
"xlXYScatter","xlXYScatterLines","xlXYScatterLinesNoMarkers","xlXYScatterSmooth",
"xlXYScatterSmoothNoMarkers")]
[PSObject] $ChartType,
[Parameter(Position=4)] $Title,
[Parameter(Position=5)] [ValidateSet("begin","end")] $SheetPosition = "begin",
[Switch] $ChartOnNewSheet,
[Switch] $AppendWorksheet,
[Switch] $Borders = $True,
[Switch] $HeaderColor = $True,
[Switch] $AutoFit = $True,
[Switch] $AutoFilter = $True,
[Switch] $PassThrough,
[Switch] $Force
)
Begin {
## Begin Convert-NumberToA1
Function Convert-NumberToA1 {
Param([parameter(Mandatory=$true)] [int]$number)
$a1Value = $null
While ($number -gt 0) {
$multiplier = [int][system.math]::Floor(($number / 26))
$charNumber = $number - ($multiplier * 26)
If ($charNumber -eq 0) { $multiplier-- ; $charNumber = 26 }
$a1Value = [char]($charNumber + 96) + $a1Value
$number = $multiplier
}
## End Convert-NumberToA1
Return $a1Value
}
## End Convert-NumberToA1
$Script:WorkingData = @()
}
## End Convert-NumberToA1
Process {
$Script:WorkingData += $InputData
}
## End Convert-NumberToA1
End {
$Props = $Script:WorkingData[0].PSObject.properties | % { $_.Name }
$Rows = $Script:WorkingData.Count+1
$Cols = $Props.Count
$A1Cols = Convert-NumberToA1 $Cols
$Array = New-Object 'object[,]' $Rows,$Cols
$Col = 0
$Props | % {
$Array[0,$Col] = $_.ToString()
$Col++
}
## End Convert-NumberToA1
$Row = 1
$Script:WorkingData | % {
$Item = $_
$Col = 0
$Props | % {
If ($Item.($_) -eq $Null) {
$Array[$Row,$Col] = ""
} Else {
## End Convert-NumberToA1
$Array[$Row,$Col] = $Item.($_).ToString()
}
## End Convert-NumberToA1
$Col++
}
## End Convert-NumberToA1
$Row++
}
## End Convert-NumberToA1
$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $False
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XLFileFormat]::xlWorkbookNormal
If ([System.IO.Path]::GetExtension($Path) -eq '.xlsx') {
If ($xl.Version -lt 12) {
$Path = $Path.Replace(".xlsx",".xls")
} Else {
## End Convert-NumberToA1
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XLFileFormat]::xlWorkbookDefault
}
## End Convert-NumberToA1
}
## End Convert-NumberToA1
If (Test-Path -Path $Path -PathType "Leaf") {
If ($AppendWorkSheet) {
$wb = $xl.Workbooks.Open($Path)
If ($SheetPosition -eq "end") {
$wb.Worksheets.Add([System.Reflection.Missing]::Value,$wb.Sheets.Item($wb.Sheets.Count)) | Out-Null
} Else {
## End Convert-NumberToA1
$wb.Worksheets.Add($wb.Worksheets.Item(1)) | Out-Null
}
## End Convert-NumberToA1
} Else {
## End Convert-NumberToA1
If (!($Force)) {
$Path = $Path.Insert($Path.LastIndexOf(".")," - $(Get-Date -Format "ddMMyyyy-HHmm")")
}
## End Convert-NumberToA1
$wb = $xl.Workbooks.Add()
While ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item(1).Delete() }
}
## End Convert-NumberToA1
} Else {
## End Convert-NumberToA1
$wb = $xl.Workbooks.Add()
While ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item(1).Delete() }
}
## End Convert-NumberToA1
$ws = $wb.ActiveSheet
Try { $ws.Name = $WorksheetName }
Catch { }
If ($Title) {
$ws.Cells.Item(1,1) = $Title
$TitleRange = $ws.Range("a1","$($A1Cols)2")
$TitleRange.Font.Size = 18
$TitleRange.Font.Bold=$True
$TitleRange.Font.Name = "Cambria"
$TitleRange.Font.ThemeFont = 1
$TitleRange.Font.ThemeColor = 4
$TitleRange.Font.ColorIndex = 55
$TitleRange.Font.Color = 8210719
$TitleRange.Merge()
$TitleRange.VerticalAlignment = -4160
$usedRange = $ws.Range("a3","$($A1Cols)$($Rows + 2)")
If ($HeaderColor) {
$ws.Range("a3","$($A1Cols)3").Interior.ColorIndex = 48
$ws.Range("a3","$($A1Cols)3").Font.Bold = $True
}
## End Convert-NumberToA1
} Else {
## End Convert-NumberToA1
$usedRange = $ws.Range("a1","$($A1Cols)$($Rows)")
If ($HeaderColor) {
$ws.Range("a1","$($A1Cols)1").Interior.ColorIndex = 48
$ws.Range("a1","$($A1Cols)1").Font.Bold = $True
}
## End Convert-NumberToA1
}
## End Convert-NumberToA1
$usedRange.Value2 = $Array
If ($Borders) {
$usedRange.Borders.LineStyle = 1
$usedRange.Borders.Weight = 2
}
## End Convert-NumberToA1
If ($AutoFilter) { $usedRange.AutoFilter() | Out-Null }
If ($AutoFit) { $ws.UsedRange.EntireColumn.AutoFit() | Out-Null }
If ($ChartType) {
[Microsoft.Office.Interop.Excel.XlChartType]$ChartType = $ChartType
If ($ChartOnNewSheet) {
$wb.Charts.Add().ChartType = $ChartType
$wb.ActiveChart.setSourceData($usedRange)
Try { $wb.ActiveChart.Name = "$($WorksheetName) - Chart" }
Catch { }
$wb.ActiveChart.Move([System.Reflection.Missing]::Value,$wb.Sheets.Item($ws.Name))
} Else {
## End Convert-NumberToA1
$ws.Shapes.AddChart($ChartType).Chart.setSourceData($usedRange) | Out-Null
}
## End Convert-NumberToA1
}
## End Convert-NumberToA1
$wb.SaveAs($Path,$xlFixedFormat)
$wb.Close()
$xl.Quit()
While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($usedRange)) {}
While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)) {}
If ($Title) { While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($TitleRange)) {} }
While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)) {}
While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)) {}
[GC]::Collect()
If ($PassThrough) { Return Get-Item $Path }
}

}
## End Export-Xlsx
## Begin Convert-ChassisType
Function Convert-ChassisType{
    Param ([int[]]$ChassisType)
    $List = New-Object System.Collections.ArrayList
    Switch ($ChassisType) {
        0x0001  {[void]$List.Add('Other')}
        0x0002  {[void]$List.Add('Unknown')}
        0x0003  {[void]$List.Add('Desktop')}
        0x0004  {[void]$List.Add('Low Profile Desktop')}
        0x0005  {[void]$List.Add('Pizza Box')}
        0x0006  {[void]$List.Add('Mini Tower')}
        0x0007  {[void]$List.Add('Tower')}
        0x0008  {[void]$List.Add('Portable')}
        0x0009  {[void]$List.Add('Laptop')}
        0x000A  {[void]$List.Add('Notebook')}
        0x000B  {[void]$List.Add('Hand Held')}
        0x000C  {[void]$List.Add('Docking Station')}
        0x000D  {[void]$List.Add('All in One')}
        0x000E  {[void]$List.Add('Sub Notebook')}
        0x000F  {[void]$List.Add('Space-Saving')}
        0x0010  {[void]$List.Add('Lunch Box')}
        0x0011  {[void]$List.Add('Main System Chassis')}
        0x0012  {[void]$List.Add('Expansion Chassis')}
        0x0013  {[void]$List.Add('Subchassis')}
        0x0014  {[void]$List.Add('Bus Expansion Chassis')}
        0x0015  {[void]$List.Add('Peripheral Chassis')}
        0x0016  {[void]$List.Add('Storage Chassis')}
        0x0017  {[void]$List.Add('Rack Mount Chassis')}
        0x0018  {[void]$List.Add('Sealed-Case PC')}
    }
    $List -join ', '
}
## End Convert-ChassisType
## Begin ConvertFrom-Base64
Function ConvertFrom-Base64{
	<#
	.SYNOPSIS
		Converts the specified string, which encodes binary data as base-64 digits, to an equivalent 8-bit unsigned integer array.
	
	.DESCRIPTION
		Converts the specified string, which encodes binary data as base-64 digits, to an equivalent 8-bit unsigned integer array.
	
	.PARAMETER String
		Specifies the String to Convert
		
	.EXAMPLE
		ConvertFrom-Base64 -String $ImageBase64 |Out-File ImageTest.png
	
	.NOTES
		Francois-Xavier Cat
		@lazywinadm
		www.lazywinadmin.com
		github.com/lazywinadmin
#>
	[CmdletBinding()]
	PARAM (
		[parameter(Mandatory = $true, ValueFromPipeline)]
		[String]$String
	)
	TRY
	{
		Write-Verbose -Message "[ConvertFrom-Base64] Converting String"
		[System.Text.Encoding]::Default.GetString(
		[System.Convert]::FromBase64String($String)
		)
	}
	CATCH
	{
		Write-Error -Message "[ConvertFrom-Base64] Something wrong happened"
		$Error[0].Exception.Message
	}
}
## End ConvertFrom-Base64
## Begin ConvertTo-Base64
Function ConvertTo-Base64{
<#
	.SYNOPSIS
		Function to convert an image to Base64
	
	.DESCRIPTION
		Function to convert an image to Base64
	
	.PARAMETER Path
		Specifies the path of the file
	
	.EXAMPLE
		ConvertTo-Base64 -Path "C:\images\PowerShellLogo.png"
	
	.NOTES
		Francois-Xavier Cat
		@lazywinadm
		www.lazywinadmin.com
		github.com/lazywinadmin
#>
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path -Path $_ })]
		[String]$Path
	)
	Write-Verbose -Message "[ConvertTo-Base64] Converting image to Base64 $Path"
	[System.convert]::ToBase64String((Get-Content -Path $path -Encoding Byte))
}
## End ConvertTo-Base64
## Begin ConvertTo-StringList
Function ConvertTo-StringList{
<#
	.SYNOPSIS
		Function to convert an array into a string list with a delimiter.
	
	.DESCRIPTION
		Function to convert an array into a string list with a delimiter.
	
	.PARAMETER Array
		Specifies the array to process.
	
	.PARAMETER Delimiter
		Separator between value, default is ","
	
	.EXAMPLE
		$Computers = "Computer1","Computer2"
		ConvertTo-StringList -Array $Computers
	
		Output: 
		Computer1,Computer2
	
	.EXAMPLE
		$Computers = "Computer1","Computer2"
		ConvertTo-StringList -Array $Computers -Delimiter "__"
	
		Output: 
		Computer1__Computer2
	
	.EXAMPLE
		$Computers = "Computer1"
		ConvertTo-StringList -Array $Computers -Delimiter "__"
	
		Output: 
		Computer1
		
	.NOTES
		Francois-Xavier Cat
		www.lazywinadmin.com
		@lazywinadm
	
		I used this Function in System Center Orchestrator (SCORCH).
		This is sometime easier to pass data between activities
#>
	
	[CmdletBinding()]
	[OutputType([string])]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true)]
		[System.Array]$Array,
		
		[system.string]$Delimiter = ","
	)
	
	BEGIN { $StringList = "" }
	PROCESS
	{
		Write-Verbose -Message "Array: $Array"
		foreach ($item in $Array)
		{
			# Adding the current object to the list
			$StringList += "$item$Delimiter"
		}
		Write-Verbose "StringList: $StringList"
	}
	END
	{
		TRY
		{
			IF ($StringList)
			{
				$lenght = $StringList.Length
				Write-Verbose -Message "StringList Lenght: $lenght"
				
				# Output Info without the last delimiter
				$StringList.Substring(0, ($lenght - $($Delimiter.length)))
			}
		}# TRY
		CATCH
		{
			Write-Warning -Message "[END] Something wrong happening when output the result"
			$Error[0].Exception.Message
		}
		FINALLY
		{
			# Reset Variable
			$StringList = ""
		}
	}
}
## End ConvertTo-StringList
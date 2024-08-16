$mrMapiPath = "c:\mrmapi.exe"
$rptCollection = @()

function ParseBitValue($String){
	$numItempattern = '(?=\().*(?=bytes)'
	$matchedItemsNumber = [regex]::matches($String, $numItempattern) 
	$bytes = [INT64]$matchedItemsNumber[0].Value.Replace("(","").Replace(",","") /1MB
	return [System.Math]::Round($bytes,2)
}

$Encode = new-object System.Text.UnicodeEncoding
##Check for Office 2013
$RootKey = "Software\Microsoft\Office\15.0\Outlook\Profiles\"
$pkProfileskey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($RootKey, $true)

if($pkProfileskey -eq $null){
	$RootKey = "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
	$pkProfileskey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($RootKey, $true)
	$defProf = $pkProfileskey.GetValue("DefaultProfile")
}
else{
	$OutDefault = "Software\Microsoft\Office\15.0\Outlook\"
	$pkProfileskey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($OutDefault, $true)
	$defProf = $pkProfileskey.GetValue("DefaultProfile")
}
$defProf
$pkSubProfilekey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey(($RootKey + "\\" + $defProf), $true)
foreach($Valuekey in $pkSubProfilekey.getSubKeyNames()){
	$pkSubValueKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey(($RootKey + "\\" + $defProf + "\\" + $Valuekey ), $true)
	$pstVal = $pkSubValueKey.GetValue("001f6700")
	if($pstVal -ne $null)
	{
		$pstPath = $Encode.GetString($pstVal) 
		$fnFileName = $pstPath.Replace([System.Convert]::ToChar(0x0).ToString().Trim(), "")		
		if(Test-Path $fnFileName){
			$rptObj = "" | Select Name,Type,FilePath,FileSize,FreeSpace,PrecentFree
			$fiInfo = ([System.IO.FileInfo]$fnFileName) 
			$rptObj.Name = $fiInfo.Name
			$rptObj.FilePath = $fiInfo.DirectoryName
			$rptObj.Type = "PST"			
			iex ($mrMapiPath + " -pst -i '$pstPath'") | foreach-object{
				if($_.Contains("=")){
					$splitArray = $_.Split("=")				
					switch ($splitArray[0].Trim()){					
						"File Size" {
								$rptObj.FileSize =  ParseBitValue($splitArray[1].Trim()) 
							}
						"Free Space" {
								$rptObj.FreeSpace = ParseBitValue($splitArray[1].Trim()) 
							}
						"Percent free" {
								$rptObj.PrecentFree = $splitArray[1].Trim()
							}
					}
					
				}
			}
			$rptCollection += $rptObj
		}
	}
	## Check OSTs
	$ostVal = $pkSubValueKey.GetValue("001f6610")
	if($ostVal -ne $null)
	{
		$ostPath = $Encode.GetString($ostVal) 
		$fnFileName = $ostPath.Replace([System.Convert]::ToChar(0x0).ToString().Trim(), "")		
		if(Test-Path $fnFileName){
			$rptObj = "" | Select Name,Type,FilePath,FileSize,FreeSpace,PrecentFree
			$fiInfo = ([System.IO.FileInfo]$fnFileName) 
			$rptObj.Name = $fiInfo.Name
			$rptObj.FilePath = $fiInfo.DirectoryName
			$rptObj.Type = "OST"
			iex ($mrMapiPath + " -pst -i `"$fnFileName`"") | foreach-object{
				if($_.Contains("=")){
					$splitArray = $_.Split("=")				
					
					switch ($splitArray[0].Trim()){					
						"File Size" {
								$rptObj.FileSize =  ParseBitValue($splitArray[1].Trim()) 
							}
						"Free Space" {
								$rptObj.FreeSpace = ParseBitValue($splitArray[1].Trim()) 
							}
						"Percent free" {
								$rptObj.PrecentFree = $splitArray[1].Trim()
							}
					}
					
				}
			}
			$rptCollection += $rptObj
		}
	}
}


$tableStyle = @"
<style>
BODY{background-color:white;}
TABLE{border-width: 1px;
border-style: solid;
border-color: black;
border-collapse: collapse;
}
TH{border-width: 1px;
padding: 10px;
border-style: solid;
border-color: black;
background-color:thistle
}
TD{border-width: 1px;
padding: 2px;
border-style: solid;
border-color: black;
background-color:white
}
</style>
"@

$body = @"
<p style="font-size:25px;family:calibri;color:#ff9100">
$TableHeader
</p>
"@

$rptCollection | ConvertTo-HTML -head $tableStyle –body $body |Out-File c:\temp\ProfileReport-$defProf.htm






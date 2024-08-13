<#
.NOTES
	Name: PF-Find-Clean.ps1
	Original Author: Chris Heilman
	Requires: Exchange Management Shell (Exchange Server 2010) and administrator rights on the Exchange server and Public Folders.
	Version: 1.4 -- 09/25/2017

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
	
	
.SYNOPSIS
	Checks the target Exchange server for illegal characters and leading or trailing spaces in Mail-enabled Public Folder aliases.

.DESCRIPTION
	Exchange 2000 and 2003 would end up allowing illegal characters (as of Exchange 2007 and up) in the Alias attribute. When migrating legacy public folders from 2007/2010 to Exchange 2013/2016 we encounter these issues.
	
.PARAMETER Repair
	This required parameter is a $True / $False switch. When set to $False this will just run a scan of the public folders and output what it finds. When set to $True this will scan and then fix the public folders that have illegal characters or leading and trailing spaces.
	
.PARAMETER Output
	This non-required parameter is another $true / $false switch. This will produce the output file while running repair $true. This can increase the time on the script if you have a large Public Folder environment
	
.EXAMPLE
	.\PF-Clean-Fix.ps1 -Repair $false
	Run a scan against the mail-enabled public folders
	
.EXAMPLE
	.\PF-Clean-Fix.ps1 -Repair $true
	Run a scan against the mail-enabled public folders and proceed to fix them by changing the offending characters to a hyphen "-"
#>
#------------------------------------------
# Setting Parameter

Param(
  [Parameter(Mandatory=$True, Position=1)]
   [string]$Repair,

   [string]$Output
   )

#------------------------------------------
# Importing AD module:

#Import-Module ActiveDirectory
   
#------------------------------------------
# Setting our main variables

# Our list of objections in Aliases
$MPF = Get-MailPublicFolder -resultSize Unlimited | where {$_.Alias.ToCharArray() -contains ' ' -or $_.Alias.ToCharArray() -contains '@' -or $_.Alias.ToCharArray() -contains ',' -or $_.Alias.ToCharArray() -contains ':' -or $_.Alias.ToCharArray() -contains ';' -or $_.Alias.ToCharArray() -contains '(' -or $_.Alias.ToCharArray() -contains ')' -or $_.Alias.ToCharArray() -contains '\'}

# Our list of objections in DisplayName
$MPFName = Get-MailPublicFolder -resultSize Unlimited | where {$_.DisplayName.ToCharArray() -contains ' ' -or $_.DisplayName.ToCharArray() -contains '@' -or $_.DisplayName.ToCharArray() -contains ',' -or $_.DisplayName.ToCharArray() -contains ':' -or $_.DisplayName.ToCharArray() -contains ';' -or $_.DisplayName.ToCharArray() -contains '(' -or $_.DisplayName.ToCharArray() -contains ')' -or $_.DisplayName.ToCharArray() -contains '\'}

# The count of objections
$MPFCount = ($MPF | measure).Count

$MPFNameCount = ($MPFName | measure).count

#------------------------------------------
#newLine shortcut
$script:nl = "`r`n"
$nl

Clear-Host


#------------------------------------------
# Starting Transcript
function transcriptStart

{
Start-Transcript PublicFolder-Find-Clean.txt -append
$nl
}


#------------------------------------------
# Stopping Transcript
function transcriptStop

{
$nl
Stop-Transcript
}

#------------------------------------------
# Checking Exchange Versions and Server Name
function getVersion 
{
Write-Host "Script starting at:" -foregroundcolor White
Get-Date
	$nl
        Write-Host "----------------------------" -foregroundcolor Green

	Write-Host "Checking Exchange Version..." -foregroundcolor White
    $nl
    $script:serverName = Hostname
	Write-Host "Server Name:" -foregroundcolor Yellow
		Write-Host "		$serverName" -foregroundcolor Green
			Write-Host "----------------------------" -foregroundcolor Green
	$nl
        $nl
        
	$script:exVer = (get-exchangeserver $serverName).admindisplayversion
		$exVerMajor = $exVer.major
		$exVerMinor = $exVer.minor

	switch ($exVerMajor) {
        "08" {
	        $script:exVer = "2007"
        }
        "14" {
	        $script:exVer = "2010"
        }	
		
    default {
		write-host "This script is only for Exchange 2007 and 2010 servers." -foregroundcolor red $nl
		    do
			{
				Stop-Transcript
                Write-Host
				$continue = Read-Host "Press <Enter> key to exit..." -foregroundcolor Yellow
			}
			While ($continue -notmatch $null)
		    exit }
			}
}

#------------------------------------------
function beforeOutput
{
Get-PublicFolder \ -recurse -ResultSize Unlimited | where{$_.MailEnabled -eq "True"} | Get-MailPublicFolder -resultSize Unlimited | ft Alias, Identity -AutoSize| Out-File Mail-Public-Folders-Before.txt
}


#------------------------------------------
function afterOutput
{
Get-PublicFolder \ -recurse -ResultSize Unlimited | where{$_.MailEnabled -eq "True"} | Get-MailPublicFolder -resultSize Unlimited | ft Alias, Identity -AutoSize | Out-File Mail-Public-Folders-After.txt
}


#------------------------------------------
# Grabbing and Sorting through the Mail-enabled Public Folders

function sortMailPF
{
If ($MPFCount -ne $null){$MPFCount = $MPFCount}

Else {$MPFCount = 0}


Write-Host "----------------------------" -foregroundcolor Green

$nl
Write-Host "Found $MPFCount Mail-enabled Public Folders with Spaces or Bad Characters." -foregroundcolor Yellow

}

#------------------------------------------
# Grabbing and Sorting again for recheck

function sortMailPFagain
{

$MPFCount = ($MPF2).count

If ($MPFCount -eq $null){$MPFCount = 0}

Write-Host "----------------------------" -foregroundcolor Green

$nl
Write-Host "Found $MPFCount Mail-enabled Public Folders with Spaces or Bad Characters." -foregroundcolor Yellow

}


#------------------------------------------
# Previewing our replacement of Special Characters with an Hypens "-"

function previewReplaceAlias
{

foreach($pf in $MPF){
            $newAlias = $pf.alias
if($newAlias -ne $null){
            $newAlias = $newAlias.Trim()
            $newAlias = $newAlias.Replace(' ', '')
		    $newAlias = $newAlias.Replace(',', '-')
	        $newAlias = $newAlias.Replace('@', '-')
	        $newAlias = $newAlias.Replace('(', '-')
	        $newAlias = $newAlias.Replace(')', '-')
	        $newAlias = $newAlias.Replace(':', '-')
	        $newAlias = $newAlias.Replace(';', '-')
	        $newAlias = $newAlias.Replace('\', '-')
		    $newAlias = $newAlias.Trim()
			Write-Host("New Alias would be: {0}" -f $newAlias) -foregroundcolor Cyan 
}
else
{
	Write-host("Public Folder Aliases are empty") -foregroundcolor Green
	Write-Host "----------------------------" -foregroundcolor Green
        } 
    }
}


#------------------------------------------
# Replacement of Special Characters with a Hypens "-"

function replaceCharactersAlias
{

foreach($pf in $MPF){
            $newAlias = $pf.alias
	if($newAlias -ne $null){
	        $newAlias = $newAlias.Trim()
	        $newAlias = $newAlias.Replace(' ', '')
		    $newAlias = $newAlias.Replace(',', '-')
	        $newAlias = $newAlias.Replace('@', '-')
	        $newAlias = $newAlias.Replace('(', '-')
	        $newAlias = $newAlias.Replace(')', '-')
	        $newAlias = $newAlias.Replace(':', '-')
            $newAlias = $newAlias.Replace(';', '-')
            $newAlias = $newAlias.Replace('\', '-')
		    $newAlias = $newAlias.Trim()
			Write-Host("New Alias is now: {0}" -f $newAlias) -foregroundcolor Cyan 
			
		Set-MailpublicFolder -Identity $pf.identity -Alias $newAlias
		Start-Sleep -s 1
}
else{
	Write-host("Public Folder Aliases are empty") -foregroundcolor Green
	$nl
	Write-Host "----------------------------" -foregroundcolor Green
	$nl
        }
}


$nl

Write-Host "----------------------------" -foregroundcolor Green
$nl

}



function replaceCharactersName {

foreach($pf in $MPFName){
            $newName = $pf.displayName
	if($newName -ne $null){
	        $newName = $newName.Trim()
		    $newName = $newName.Replace(',', '-')
	        $newName = $newName.Replace('@', '-')
	        $newName = $newName.Replace('(', '-')
	        $newName = $newName.Replace(')', '-')
	        $newName = $newName.Replace(':', '-')
            $newName = $newName.Replace(';', '-')
            $newName = $newName.Replace('\', '-')
		    $newName = $newName.Trim()
			Write-Host("New displayName is now: {0}" -f $newName) -foregroundcolor Cyan 
			
		Set-MailpublicFolder -Identity $pf.identity -DisplayName $newName
		Start-Sleep -s 1
}
else{
	Write-host("Public Folder Display Names are clean") -foregroundcolor Green
	$nl
	Write-Host "----------------------------" -foregroundcolor Green
	$nl
        }

	}
}



#################################################################################################################
#------------------------------------------
#--------------------------------------
# Body of Script
#--------------------------------------

function Main {

#------------------------------------------

If ($Repair -eq $false)
	{
	transcriptStart
    getVersion
    If($MPF -ne $null -or $MPFName -ne $null)
	{
		Write-Host "=====================================================================================" -foregroundcolor Green $nl
		Write-Host "Sorting through the Mail-enabled Public Folders to find Spaces and Special Characters" -foregroundcolor White $nl
		Write-Host "=====================================================================================" -foregroundcolor Green 

	sortMailPF
	$nl
		Write-Host "----------------------------" -foregroundcolor Green
	previewReplace
		Write-Host "Found" $MPFNameCount "bad displayNames" -foregroundcolor White
	    Write-Host "----------------------------" -foregroundcolor Green		
        Write-host "Please re-run script with -Repair $true to fix the found items" -foregroundcolor Red
		$nl}
    Else {
	Write-Host "No Bad objects or spaces found!" $nl$nl "Ending Script" -foregroundcolor Green}
	transcriptStop
	}

#------------------------------------------

Elseif ($Repair -eq $true)
	{
	transcriptStart
	$nl
	getVersion
	If($MPF -ne $null -or $MPFName -ne $null)
	{
	If($Output -eq $true)
	{$nl
		Write-host "Outputting Mail-enabled Public Folders before changes [Mail-Public-Folders-Before.txt]" -foregroundcolor DarkYellow
    beforeOutput
          }
	Else {Write-Host "No before Output chosen" -foregroundcolor DarkYellow}
	$nl	
		Write-Host "============================" -foregroundcolor Green $nl
		Write-Host "Sorting through the Mail-enabled Public Folders to find Spaces and Special Characters" -foregroundcolor White $nl
		Write-Host "============================" -foregroundcolor Green 

	sortMailPF
	$nl
		write-host "Preview of the changes:" -foregroundcolor White
	$nl
	replaceCharactersAlias
		Write-Host "Found" $MPFNameCount "bad displayNames ---- Fixing the displayName attribute" -foregroundcolor White
	replaceCharactersName
	$nl
		Write-Host "Checking Mail Public Folders again" -foregroundcolor White
    $MPF2 = Get-MailPublicFolder -resultSize Unlimited | where {$_.Alias.ToCharArray() -contains ' ' -or $_.Alias.ToCharArray() -contains '@' -or $_.Alias.ToCharArray() -contains ',' -or $_.Alias.ToCharArray() -contains ':' -or $_.Alias.ToCharArray() -contains ';' -or $_.Alias.ToCharArray() -contains '(' -or $_.Alias.ToCharArray() -contains ')' -or $_.Alias.ToCharArray() -contains '\'}
	sortMailPFagain
	If($Output -eq $true)
	{$nl
		Write-Host "Outputting Mail-enabled Public Folders after changes [Mail-Public-Folders-After.txt]" -foregroundcolor DarkYellow
    afterOutput
          }
	Else {Write-Host "No After Output chosen" -foregroundcolor DarkYellow}
  	$nl}
	Else {
	Write-Host "No Bad objects or spaces found!" $nl $nl "Ending Script" -foregroundcolor Green}
	transcriptStop
	}
}

Main

#--------------------------------------

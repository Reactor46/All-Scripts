<#

*** THIS SCRIPT IS PROVIDED WITHOUT WARRANTY, USE AT YOUR OWN RISK ***

.DESCRIPTION
	Searches all OUs in Active Directory, or a subset of OUs and looks for the 
    ProtectedFromAccidentalDeletion property is set to $false.
    
    The code at the bottom of the script will find the OUs and set them to
    Protectedfromaccidental $true

.NOTES
	File Name: 
	Author: David Hall
	Contact Info: 
		Website: www.signalwarrant.com
		Twitter: @signalwarrant
		Facebook: facebook.com/signalwarrant/
		Google +: plus.google.com/113307879414407675617
		YouTube Subscribe link: https://www.youtube.com/c/SignalWarrant1?sub_confirmation=1
	Requires: Appropriate AD permissions
	Tested: PowerShell Version 5, Windows 10 and Windows Server 2012 R2

.PARAMETER 
    
		 
.EXAMPLE
     Run either OPTION 1 or OPTION 2

#>

# OPTION 1
# Find all OUs that are not protected from accidental deletion
Get-ADObject -Filter * -Properties CanonicalName,ProtectedFromAccidentalDeletion |
    Where-Object {$_.ProtectedFromAccidentalDeletion -eq $false -and $_.ObjectClass -eq "organizationalUnit"} | 
    Select-Object CanonicalName,ProtectedFromAccidentalDeletion |
    Out-GridView


# OPTION 2
# Find a smaller subset of OUs that are not protected from accidental deletion
Get-ADObject -Filter * -Properties CanonicalName,ProtectedFromAccidentalDeletion -SearchBase "DC=USON,DC=LOCAL" |
    Where-Object {$_.ProtectedFromAccidentalDeletion -eq $false -and $_.ObjectClass -eq "organizationalUnit"} |
    Set-ADObject -ProtectedFromAccidentalDeletion $True
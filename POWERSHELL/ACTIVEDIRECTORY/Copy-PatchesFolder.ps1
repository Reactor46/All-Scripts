<#
	** THIS SCRIPT IS PROVIDED WITHOUT WARRANTY, USE AT YOUR OWN RISK **	

    .SYNOPSIS
	    Copy Multiple files/folder to multiple computers.

	.DESCRIPTION
	    Copy a directory and all files within the directory to multiple computers.

    .REQUIREMENTS
        1.	The appropriate rights to ping and copy on the remote machine.
		2.  A computers.txt file with a list of computer names

    .NOTES
        Tested with Windows 7, Windows Vista, Windows Server 2003, Windows Server 2K8 and 2K8 R2

	.AUTHOR
		David Hall | http://www.signalwarrant.com/

	.LINK
		http://www.signalwarrant.com/2012/10/04/copy-a-folder-and-files-to-multiple-computers-powershell/

#>

# This is the file that contains the list of computers you want 
# to copy the folder and files to. Change this path IAW your folder structure.
$computers = gc "C:\scripts\computers.txt"

# This is the directory you want to copy to the computer (IE. c:\folder_to_be_copied)
$source = "c:\files"

# On the desination computer, where do you want the folder to be copied?
$dest = "c$"

foreach ($computer in $computers) {
    if (test-Connection -Cn $computer -quiet) {
        Copy-Item $source -Destination \\$computer\$dest -Recurse
    } else {
        "$computer is not online"
    }

}
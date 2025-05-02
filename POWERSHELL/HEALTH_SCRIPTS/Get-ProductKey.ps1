<#
!! Before you use !! 
Before you execute this script, you must run the following command
set-executionpolicy remotesigned
Written by Anthony Manongsong
Publishing date: 4/20/2017
Purpose: To retrieve your product key on Windows 10. This method has yet to be tested on Windows 7, Windows 8 and Windows 8.1. If
an error occurs on these operating systems, feel free to notify me on GitHub.
#>

clear
write-host "Finding Windows Product Key, please wait..." -ForegroundColor White
write-host
get-ciminstance SoftwareLicensingService | where-object {
    write-host "Your product key is:" 
    write-host $_.OA3xOriginalProductKey -ForegroundColor white
}
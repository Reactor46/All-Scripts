#purpose: deleting creditoneapp.biz users from a csv file
$csv="\\Contosocorp\share\shared\it\SupportServices\Help Desk Support\User_Admin\User_Request_Lists\Remove\Remove_AD_User_CSV\Del_AD_User_CSV.csv"
$source = Import-Csv $csv
$counter= 0
$counterDeleted=0
$counterExists=0
$misscounter=0
$TodaysDate = get-date
$global:MonthDayYear = $TodaysDate.ToString("yyyMMdd")
$global:HourMinute = $TodaysDate.ToString("hhmm")
$SuccessFileName = "Success.txt"
$FailedFileName = "Failed.txt"
$FinalFileName = "DelAdUser" + "-" + "$MonthDayYear" + "-" + "$HourMinute" + ".txt"
$InitialMsg = "Names in this file based off \\Contosocorp\share\shared\IT\SupportServices\Help Desk Support\User_Admin\User_Request_Lists\Remove\Remove_AD_User_CSV\Del_AD_User_CSV.csv"
$InitialMsg | out-file "\\Contosocorp\share\shared\IT\SupportServices\Help Desk Support\User_Admin\User_Request_Lists\Remove\Remove_AD_User_CSV\Logs\$SuccessFileName" -Append
$InitialMsg | out-file "\\Contosocorp\share\shared\IT\SupportServices\Help Desk Support\User_Admin\User_Request_Lists\Remove\Remove_AD_User_CSV\Logs\$FailedFileName" -Append

foreach ($line in $source)
{
[string] $domain=$line.Domain
[string] $EMPID=$line.EMPID

$DelUsr = Get-ADUser -Filter {SamAccountName -like $EMPID} -Server "creditoneapp.biz" -SearchBase "DC=creditoneapp,DC=biz"

if ($DelUsr -eq $Null)
   {
    $EMPID1 = $EMPID + " "
    $DelUsr1 = Get-ADUser -Filter {SamAccountName -like $EMPID1} -Server "creditoneapp.biz" -SearchBase "DC=creditoneapp,DC=biz"
      if ($Delusr1 -eq $Null)
        {
         $MissingMsg = "Unable to find Employee ID $EMPID1 on $domain, Please check for Typos in the number."
         $MissingMsg | out-file "\\Contosocorp\share\shared\IT\SupportServices\Help Desk Support\User_Admin\User_Request_Lists\Remove\Remove_AD_User_CSV\Logs\$FailedFileName" -Append
         $misscounter++
        }
      else
        {
         $Global:fn = $DelUsr.GivenName
         $Global:ln = $DelUsr.Surname
         $FoundMsg = "Found $ln, $fn...Attempting to Remove"
         write-host "Found $ln, $fn...Attempting to Remove"
         $foundMsg | out-file "\\Contosocorp\share\shared\IT\SupportServices\Help Desk Support\User_Admin\User_Request_Lists\Remove\Remove_AD_User_CSV\Logs\$SuccessFileName" -Append
         Remove-ADUser -Identity $EMPID1 -Server "$domain" -Confirm:$false
         $CheckingMsg = "Checking for deletion Success..."
         $CheckingMsg | out-file "\\Contosocorp\share\shared\IT\SupportServices\Help Desk Support\User_Admin\User_Request_Lists\Remove\Remove_AD_User_CSV\Logs\$SuccessFileName" -Append
         $CheckUser = Get-ADUser -Filter {SamAccountName -like $EMPID1} -Server "creditoneapp.biz" -SearchBase "DC=creditoneapp,DC=biz"
         if ($CheckUser -eq $Null)
            {
             $SuccessChk = "$ln, $fn, $EMPID1 - Deletion SUCCESS"
             $SuccessChk | out-file "\\Contosocorp\share\shared\IT\SupportServices\Help Desk Support\User_Admin\User_Request_Lists\Remove\Remove_AD_User_CSV\Logs\$SuccessFileName" -Append
             $Counter++
            }
         else
             {
              $FailChk = "$ln, $fn, $EMPID1 - Deletion FAILED"
              $FailChk | out-file "\\Contosocorp\share\shared\IT\SupportServices\Help Desk Support\User_Admin\User_Request_Lists\Remove\Remove_AD_User_CSV\Logs\$FailedFileName" -Append
              $misscounter++
             }
        }
   }
else
   {
    $Global:fn = $DelUsr.GivenName
    $Global:ln = $DelUsr.Surname
    $FoundMsg = "Found $ln, $fn...Attempting to Remove"
    write-host "Found $ln, $fn...Attempting to Remove"
    $foundMsg | out-file "\\Contosocorp\share\shared\IT\SupportServices\Help Desk Support\User_Admin\User_Request_Lists\Remove\Remove_AD_User_CSV\Logs\$SuccessFileName" -Append
    Remove-ADUser -Identity $EMPID -Server "$domain" -Confirm:$false
    $CheckingMsg = "Checking for deletion Success..."
    $CheckingMsg | out-file "\\Contosocorp\share\shared\IT\SupportServices\Help Desk Support\User_Admin\User_Request_Lists\Remove\Remove_AD_User_CSV\Logs\$SuccessFileName" -Append
    $CheckUser = Get-ADUser -Filter {SamAccountName -like $EMPID} -Server "creditoneapp.biz" -SearchBase "DC=creditoneapp,DC=biz"
    if ($CheckUser -eq $Null)
       {
        $SuccessChk = "$ln, $fn, $EMPID - Deletion SUCCESS"
        $SuccessChk | out-file "\\Contosocorp\share\shared\IT\SupportServices\Help Desk Support\User_Admin\User_Request_Lists\Remove\Remove_AD_User_CSV\Logs\$SuccessFileName" -Append
        $Counter++
       }
    else
        {
         $FailChk = "$ln, $fn, $EMPID - Deletion FAILED"
         $FailChk | out-file "\\Contosocorp\share\shared\IT\SupportServices\Help Desk Support\User_Admin\User_Request_Lists\Remove\Remove_AD_User_CSV\Logs\$FailedFileName" -Append
         $misscounter++
        }    
   }
 }
$FinalMsgSuccess = "$Counter Users Successfully Deleted"
$FinalMsgFailed = "$misscounter Users either not found or failed to delete"
$FinalMsgSuccess | out-file "\\Contosocorp\share\shared\IT\SupportServices\Help Desk Support\User_Admin\User_Request_Lists\Remove\Remove_AD_User_CSV\Logs\$SuccessFileName" -Append
$FinalMsgFailed | out-file "\\Contosocorp\share\shared\IT\SupportServices\Help Desk Support\User_Admin\User_Request_Lists\Remove\Remove_AD_User_CSV\Logs\$FailedFileName" -Append

get-content "\\Contosocorp\share\shared\IT\SupportServices\Help Desk Support\User_Admin\User_Request_Lists\Remove\Remove_AD_User_CSV\Logs\$SuccessFileName" > "\\Contosocorp\share\shared\IT\SupportServices\Help Desk Support\User_Admin\User_Request_Lists\Remove\Remove_AD_User_CSV\Logs\$FinalFileName"
get-content "\\Contosocorp\share\shared\IT\SupportServices\Help Desk Support\User_Admin\User_Request_Lists\Remove\Remove_AD_User_CSV\Logs\$FailedFileName" >> "\\Contosocorp\share\shared\IT\SupportServices\Help Desk Support\User_Admin\User_Request_Lists\Remove\Remove_AD_User_CSV\Logs\$FinalFileName"
remove-item "\\Contosocorp\share\shared\IT\SupportServices\Help Desk Support\User_Admin\User_Request_Lists\Remove\Remove_AD_User_CSV\Logs\$SuccessFileName"
remove-item "\\Contosocorp\share\shared\IT\SupportServices\Help Desk Support\User_Admin\User_Request_Lists\Remove\Remove_AD_User_CSV\Logs\$FailedFileName"

$web = get-spweb http://www2.kelsey-seybold.com
#$web = get-spweb http://www2.kelseycareadvantage.com

foreach ($user in $web.AllUsers) 
{  
    if ($user.ToString().Contains("|adfs20|"))
    {
        $oldID = $user.ToString()
        
        #strip the username from the account identifier
        $login = $user.ToString().Split("|")[2]
        
        #recreate the identity correctly using the login name of the user
        $newID = "i:0" + [char]0x01f5 + ".t|adfs20|" + $login
        
        try
        {
            #migrate the user account
            move-spuser -Identity $user  -NewAlias $newID  -ignoresid -confirm:$false
        }
        catch
        {        
            #use this to suppress a useless error
        }
        Write-Host $oldID "successfully migrated to" $newID
    }

}
$web.Dispose()
Write-Host "Script completed"
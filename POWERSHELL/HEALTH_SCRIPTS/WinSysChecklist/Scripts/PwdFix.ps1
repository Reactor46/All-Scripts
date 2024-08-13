Clear-Host

$UserInfo = jbattista #Get-Content 'C:\temp\pwlist.txt'


#This script will change the password last set date on an AD account to the current day, basicly reseting the Password experation start date without changing the password.  
#This should be used when a password refresh is the only thing required. This type of script is designed for users who are only setup to recive email and not 
#actualy log into the domain thus can not change their password.


    foreach($UserName in $UserInfo){

        $ADUser = Get-ADUser -Identity $UserName -Properties *
        Write-Host $ADUser.Name
        #Currently the Last Set date for the account is:
        Write-Host "PwdLastSet on "  -NoNewline -ForegroundColor Yellow
        Write-Host ([datetime]::FromFileTime($ADUser.pwdLastSet)) -ForegroundColor Yellow
        #First set the PwdLastSet to 0 - Will not work unless this is done first. (Ask Microsoft not me.)
        
        $aduser.PwdLastSet = 0
        Set-ADUser -Instance $aduser
        #Now Set the LastSet date to today
        
        $aduser.PwdLastSet = -1
        Set-ADUser -Instance $aduser
        #Now the Last Set Date is Changed.

        


    }





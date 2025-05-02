function Remove-UserProfiles {

<#
.SYNOPSIS
    Written by: JBear 1/31/2017
	
    Remove user profiles from a specified system.

.DESCRIPTION
    Remove user profiles from a specified system with the use of DelProf2.exe.

.EXAMPLE
    Remove-UserProfiles Computer123456

        Note: Follow instructions and prompts to completetion.

#>

    param(
        [parameter(mandatory=$true)]
        [string[]]$computername
    )

　
    function UseDelProf2 { 
               
        #Set parameters for remote computer and -WhatIf (/l)
        $WhatIf = @(

            "/l",
            "/c:$computer" 
        )
           
        #Runs DelProf2.exe with the /l parameter (or -WhatIf) to list potential User Profiles tagged for potential deletion
        & "C:\LazyWinAdmin\Win32-Tools\DelProf2.exe" $WhatIf

        #Display instructions on console
        Write-Host "`n`nPLEASE ENSURE YOU FULLY UNDERSTAND THIS COMMAND BEFORE USE `nTHIS WILL DELETE ALL USER PROFILE INFORMATION FOR SPECIFIED USER(S) ON THE SPECIFIED WORKSTATION!`n"

        #Prompt User for input
        $DeleteUsers = Read-Host -Prompt "To delete User Profiles, please use the following syntax ; Wildcards (*) are accepted. `nExample: /id:user1 /id:smith* /id:*john*`n `nEnter proper syntax to remove specific users" 

        #If only whitespace or a $null entry is entered, command is not run
        if([string]::IsNullOrWhiteSpace($DeleteUsers)) {

            Write-Host "`nImproper value entered, excluding all users from deletion. You will need to re-run the command on $computer, if you wish to try again...`n"

        }

        #If Read-Host contains proper syntax (Starts with /id:) run command to delete specified user; DelProf will give a confirmation prompt
        elseif($DeleteUsers -like "/id:*") {

            #Set parameters for remote computer
            $UserArgs = @(

                "/c:$computer"
            )

            #Split $DeleteUsers entries and add to $UserArgs array
            $UserArgs += $DeleteUsers.Split("")

            #Runs DelProf2.exe with $UserArgs parameters (i.e. & "C:\DelProf2.exe" /c:Computer1 /id:User1* /id:User7)
            & "C:\LazyWinAdmin\Win32-Tools\DelProf2.exe" $UserArgs
        }

        #If Read-Host doesn't begin with the input /id:, command is not run
        else {

            Write-Host "`nImproper value entered, excluding all users from deletion. You will need to re-run the command on $computer, if you wish to try again...`n"
        }
    }

    foreach($computer in $computername) {
        if(Test-Connection -Quiet -Count 1 -Computer $Computer) { 

            UseDelProf2 
        }

        else {
            
            Write-Host "`nUnable to connect to $computer. Please try again..." -ForegroundColor Red
        }

    }
}#End Remove-UserProfiles
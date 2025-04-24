#This script asks you for your DOB and then informs you of your star sign.

#A function is defined. This an efficient way of referencing and running a block of code. The code will serve as the logic behind what displays after the user enters their DOB.
function Get-StarSign {
    param(
    [string]$dob
    )

    $birthdate = Get-Date $dob
    $month = $birthdate.Month
    $day = $birthdate.Day

    if (($month -eq 3 -and $day -ge 21) -or ($month -eq 4 -and $day -le 19)) {
        "Aries"
    }
    elseif (($month -eq 4 -and $day -ge 20) -or ($month -eq 5 -and $day -le 20)) {
        "Taurus"
    }
    elseif (($month -eq 5 -and $day -ge 21) -or ($month -eq 6 -and $day -le 20)) {
        "Gemini"
    }
    elseif (($month -eq 6 -and $day -ge 21) -or ($month -eq 7 -and $day -le 22)) {
        "Cancer"
    }
    elseif (($month -eq 7 -and $day -ge 23) -or ($month -eq 8 -and $day -le 22)) {
        "Leo"
    }
    elseif (($month -eq 8 -and $day -ge 23) -or ($month -eq 9 -and $day -le 22)) {
        "Virgo"
    }
    elseif (($month -eq 9 -and $day -ge 23) -or ($month -eq 10 -and $day -le 22)) {
        "Libra"
    }
        elseif (($month -eq 10 -and $day -ge 23) -or ($month -eq 11 -and $day -le 21)) {
        "Scorpio"
    }
    elseif (($month -eq 11 -and $day -ge 22) -or ($month -eq 12 -and $day -le 21)) {
        "Sagittarius"
    }
    elseif (($month -eq 12 -and $day -ge 22) -or ($month -eq 1 -and $day -le 19)) {
        "Capricorn"
    }
    elseif (($month -eq 1 -and $day -ge 20) -or ($month -eq 2 -and $day -le 18)) {
        "Aquarius"
    }
    elseif (($month -eq 2 -and $day -ge 19) -or ($month -eq 3 -and $day -le 20)) {
        "Pisces"
    }
    else {
        "Unable to determine star sign."
    }

}

# Prompt the user for their date of birth
$dob = Read-Host "Enter your date of birth (DD/MM/YYYY)"

# Call the function to get the star sign. Notice that the user's input was stored in the $dob variable previously, so this is how it is referenced.
$starSign = Get-StarSign -dob $dob

# Display the result
Write-Host "Your star sign is: $starSign"
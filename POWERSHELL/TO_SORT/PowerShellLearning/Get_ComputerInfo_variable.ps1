#Obtaining basic facts about my computer like BIOS, OS and Hardware ops and storing this info as a variable
$computer_info = Get-ComputerInfo

#Checking individual items of info
$computer_info.TimeZone
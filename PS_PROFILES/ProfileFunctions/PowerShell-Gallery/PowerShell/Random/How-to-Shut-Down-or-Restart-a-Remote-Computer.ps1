#Shut down a computer
Stop-Computer -ComputerName REMOTE_COMPUTER_NAME -Force

# Shut down multiple computers
Stop-Computer -ComputerName "Server01", "Server02"

#Restart a computer
Restart-Computer -ComputerName REMOTE_COMPUTER_NAME -Force

#Restart mutiple computers from a list
Restart-Computer (Get-Content C:computerlist.txt)


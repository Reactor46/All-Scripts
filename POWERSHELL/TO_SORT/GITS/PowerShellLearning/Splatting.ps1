#Creating a hastable to make creating a text file easier. This is called 'splatting' and entails establishing multiple parameters.

$parameters = @{

Path = 'C:\Users\silad\Documents\Powershell_stuff\PowerShellLearning\file.txt'
Encoding = 'UTF8'

}

#Having now established @parameters as being a hashtable with the path for the resulting txt file and encoding, a command can then be used that invokes @parameters to add text to the new file. It is shown below.

Set-Content @parameters -Value "Hi"

#Ultimately, what we have produced is a text file called 'file.txt' that reads 'Hi' in the directory 'C:\Users\silad\Documents\Powershell_stuff\PowerShellLearning'
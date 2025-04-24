#Invoke-WebRequest is the cmdlet responsible for the web-scarping operation
#-Uri is the clause for the url
#-OutFile is the clause for the file location of the image I'm downloading to my local

Invoke-WebRequest -Uri https://en.wikipedia.org/wiki/The_Wizarding_World_of_Harry_Potter#/media/File:Wizarding_World_of_Harry_Potter_Castle.jpg -OutFile C:\Users\silad\Documents\Powershell_stuff\PowerShellLearning\Castle.jpg -UseBasicParsing
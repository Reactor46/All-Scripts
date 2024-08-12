@ECHO OFF
rem Written by Justin Marks
rem Uses dsquery to display all enabled users in the domain, pipes to C:\ActiveUsers.txt
rem Currently outputs user display name, email address, department, and title.
rem If you want to see other parameters for dsget user output, go to: 
rem http://technet.microsoft.com/en-us/library/cc732535(WS.10).aspx
ECHO Running Query
dsquery * -filter "(&(sAMAccountType=805306368)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))" -limit 0 | dsget user -display -email -dept -title > C:\ActiveUsers.txt
ECHO Results saved to C:\ActiveUsers.txt
PAUSE
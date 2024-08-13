Note: 

Add the list of your servers to the text file Servers.txt and and paste it in a directory on your server or desktop.

To get the script to work, please edit the script and amend the following:

1. The script is designed to report only servers with 10% or less free space. If you wish to report all free disk spaces, please comment the following lines in the swcript: #Where-Object {   ($_.freespace/$_.size) -le '0.1'} - This is found directly beneath the Get-WmiObject win32_logicaldisk command (Around line 50).

2. Edit line 20, to change number of days that reports are left on the server

3. Edit line 25 to the actual path for your disk report

4. Edit line 37 to the right path where you placed your Servers.txt file 

5. Edit line 76 the actual path for your disk report

6. If you want to send report, you will require the exchange PowerShell module, Exchange.ps1. Edit line 83 to the ctual path of your exchange module

7. Edit $messageParameters as reuired 

8. You might need to configure your antivirus software (McAfee ePo, etc) to allow this server to send email. 



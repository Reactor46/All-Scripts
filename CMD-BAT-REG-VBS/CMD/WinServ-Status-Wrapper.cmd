@ECHO OFF
###################################################################################################################################
#Command Line Switch #	Mandatory # Description 														    # Example                 #
###################################################################################################################################
# -list              # Yes       #	 The text file with the servers you wish to check.                   # C:\scripts\Servers.txt  #
###################################################################################################################################
# -o         		   # Yes       # 	The location to store the HTML file.	                               # C:\scripts              # 
###################################################################################################################################
# -cpualert	         # No        #	The minimum percentage of usage that alert status should be raised.	 # 95                      #  
###################################################################################################################################
# -diskalert         # No        #	The minimum percentage of usage that alert status should be raised.	 # 85                      #
###################################################################################################################################
# -memalert	         # No        #	The minimum percentage of usage that alert status should be raised.	 # 90                      #
###################################################################################################################################
# -refresh	          # No        #	The number of seconds the script should wait before refreshing       #                         #
#                    #           # the status.                                                          #                         # 
#			         #		      #	The webpage auto refreshes every 30 seconds, so the lowest this      #                         #
#			         #		      #	value  should be set is 30.                                          # 120                     # 
#                    #           # I recommend for 20+ servers that this number shouldn’t be lower      #                         #
#                    #           # than 60 seconds.                                                     #                         #
###################################################################################################################################
# -sendto            #	No	     #  The email address to send the report to.	                            #  me@contoso.com         #
###################################################################################################################################
# -from              # No*	     #  The email address that the report should be sent from.              #                          #
#                    #          #  *This switch isn’t mandatory but is required if you wish to         #                          #
#                    #          #  email the report.                                                   # Server-Status@contoso.com#
###################################################################################################################################
# -smtp	             # No*	     #  SMTP server address to use for the email functionality.             #   mail01.contoso.com OR  #
#                    #          #  *This switch isn’t mandatory but is required if you wish to         #   smtp.live.com      OR  #
#					 #			#  email the report.                                                   #   smtp.office365.com     #
###################################################################################################################################
# -user	             # No*	     # The username of the account to use for SMTP authentication.          #                          #
#	             #		 # *This switch isn’t mandatory but may be required depending on the    # example@contoso.com      #
#                    #          # configuration of the SMTP server.                                    #                          #
###################################################################################################################################
# -pwd	              # No*      #	The location of the file containing the encrypted password of the   #                          #
#	              #          #	account to use for SMTP authentication.                             # c:\scripts\              #
#	              #          #																		#      ps-script-pwd.txt   # 
#	              #          # 	*This switch isn’t mandatory but may be required depending          #                          #
#	              #          #	on your SMTP server.                                                #                          #
###################################################################################################################################
# -usessl	           # No*	     # Add this option if you wish to use SSL with the configured           #                          # 
#			 # SMTP server.                                                         #                          #
#                               # Tip: If you wish to send email to outlook.com or office365.com       #                          #
#                               # you will need this.                                                  #                          # 
#								# *This switch isn’t mandatory but may be required depending on the    #                          # 
#								# configuration of the SMTP server.                                    #                          #
###################################################################################################################################

SET PSCRIPT=C:\LazyWinAdmin\Servers\WinServ-Status\WinServ-Status.ps1
SET PSLIST=C:\LazyWinAdmin\Servers\WinServ-Status\Servers.txt
SET HTML=\\laspshost\scripts\Repository\jbattista\Web\reports\
SET CPUALERT=90
SET DISKALERT=90
SET MEMALERT=90
SET REFRESH=30
SET SEND=john.battista@Creditone.com
SET FROM=ServerStatus@Creditone.com
SET SMTP=mailgateway.Contoso.corp

PowerShell.exe %PSCRIPT% -list %PSLIST% -o %HTML% -cpualert %CPUALERT% -diskalert %DISKALERT% -memalert %MEMALERT% -refresh %REFRESH% -sendto %SEND% -from %FROM% -smtp %SMTP%
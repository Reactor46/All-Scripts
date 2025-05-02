#==================| Satnaam Waheguru Ji |===============================
#           
#            Author  :  Aman Dhally 
#            E-Mail  :  amandhally@gmail.com 
#            website :  www.amandhally.net 
#            twitter :   @AmanDhally 
#            blog    : http://newdelhipowershellusergroup.blogspot.in/
#            facebook: http://www.facebook.com/groups/254997707860848/ 
#            Linkedin: http://www.linkedin.com/profile/view?id=23651495 
# 
#            Creation Date    : 05-08-2013 
#            File    : 
#            Purpose : 
#            Version : 1 
#
#            My Pet Spider :          /^(o.o)^\  
#========================================================================


# Notes :- 
#             --->   Please run this script as "Administrator"
#        Tested on : 64 Bit Client of Mirage.
#                  : 64 Bit Windows 7 Professional.  

# testing for file first

$testWanovaFile = Test-Path 'C:\Program Files\Wanova\Mirage Service\Wanova.Desktop.Service.exe.config'

# 
if ( $testWanovaFile -eq $true ) 

	{

		# Mapping confifuration file as XML
		[xml] $wanovaXmlFile =  get-content 'C:\Program Files\Wanova\Mirage Service\Wanova.Desktop.Service.exe.config'

		# Finding the useSSLTransport Attribute
		$useSslKey = $wanovaXmlFile.configuration.appSettings.add | where { $_.Key -match "useSslTransport" }

		# Processing

		# If the UseSSL is not set to true
		if ( $useSslKey.value  -ne $true ) 
			{
				
				# Stopping Wanova Mirage Desktop Service
				Write-Warning "Stopping Wanova Desktop Service."
		        Stop-Service "Wanova Mirage Desktop Service" -Force
		        
				# Setting SSLkey value to True
				Write-Warning "Setting useSslTransport value to true"
				$useSslKey.value = "true"
		        
				# Saving the file with changes
				Write-Warning "Saving the Configuration file."
		        $wanovaXmlFile.Save('C:\Program Files\Wanova\Mirage Service\Wanova.Desktop.Service.exe.config')
		        
				# Staring the Service again
				Write-Host "Starting Wanova Desktop Service." -ForegroundColor 'Green'
				Start-Service "Wanova Mirage Desktop Service"

			}

		else {

				Write-Host "Value already set to true " -ForegroundColor 'Green'

			}
	
	}

else 
	{
	
	Write-Warning "Mirage Doesn't seem installed on this $env:COMPUTERNAME laptop"
	
	}	

###---- end of the script
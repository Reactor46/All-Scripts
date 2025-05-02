This script will completely build a UCS system from the ground up.
Once you have physically installed, cabled and done the initial steps via the serial port to assign base access information running this script will remove all ucs default pools, policies and templates.  It will then build all pools, policies and templates for your UCSM based on the information configured in an Answer File.

Prerequisites:
	Place this script, your data file and custom file into the same directory
	Make sure your system is running PowerShell v3.0 or later and PowerTool v1.0 or later
	Make sure you have a valid login into the UCS Manager
	Make sure the client running this script has IP access to UCS Manager

NOTE:
	Many of the settings are automatic and are based on my personal best practices.
		By default I turn on VLAN Compression
		I set many policies with full range of options
	You can always change the answer file or this script to adjust those parameters or names

The Data-BLANK answer file is very well documented to help you setup your own answer files
I have also included a Data-Sample answer file that you can run on the UCSM Emulator to see how the script sets up a sample system
	To use the Data-Sample file you will need to edit the IP Address information to run on your system
Your answer files must be named Data-<something>.ps1 to be found by this script.

There is also a sample file:Custom UCS Settings - BLANK.ps1 included.
	There is an option to run a custom script at the end of your UCS build to add any custom capabilities
	This is an advanced feature required deep knowledge of PowerShell and PowerTool.
		An example of how I have used this is a customer who had 1Gb on their northbound LAN so I sent commands
		to the uplinks to change the port speeds from 10Gb down to 1Gb
		A sample custom file has been included

Once the script has completed you should be ready to create Service Profiles from the created Service Profile Templates

Included files:
	Add-UcsDomain.ps1
	Data - BLANK.ps1
	Custom UCS Settings - BLANK.ps1
	Data - Sample.ps1
	Custom UCS Settings - Sample.ps1
	This readme
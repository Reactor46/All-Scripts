Automation with Cisco UCS PowerTool for UCSM

Following the best practices outlined in the whitepaper, Citrix XenDesktop: Best Practices with Cisco UCS (http://support.citrix.com/article/CTX135305), a Windows PowerShell script was written to automate the creation and configuration of many of the pools, policies, and templates necessary in a Citrix XenDesktop environment, culminating in the creation of a UCS Service Profile that can be used to create 1 or more Service Profiles that will ultimately be applied to physical servers used to host XenDesktop created hosted virtual desktops.
Designed primarily for Proof Of Concept environments, there is an assumption that a relatively clean UCS environment is being used, with very little configuration already completed (there are no checks in the script for existing configurations, and if there is a conflict, the script will issue an error).
The script takes the name of an Excel Configuration file containing information needed for configuration, and, optionally, a switch to output logs to the console, as input.

	Example: /UCS_Config_Excel.ps1 -ExcelFile XenDesktopBP.xlsx -ToConsole
Inside the script (starting around line 350), there is a commented-out section that could output all variables from the worksheet and derived variables (added for troubleshooting purposes).  Starting around line 505, there is another commented-out section that	can be used to clean up a UCS Platform Emulator environment prior to creation of the all the pools, policies, and templates.
The script creates the following using input from the Excel Configuration worksheet:
•	Management IP Pool
•	UUID Pool
•	8 MAC Pools (1 each for the management network, virtual machine data network, motion network (i.e.., vMotion), and IP storage network; A and B fabrics)
•	QoS System Class settings
•	4 QoS Policies (1 each for the management network, virtual machine data network, motion network (i.e.., vMotion), and IP storage network)
•	Network Control Policy
•	4 VLANs (1 each for the management network, virtual machine data network, motion network (i.e.., vMotion), and IP storage network)
•	8 vNIC Templates (1 each for the management network, virtual machine data network, motion network (i.e.., vMotion), and IP storage network; A and B fabrics)
•	WWNN Pool
•	4 WWPN Pools (for 4 vHBA's - only two are used by the template created in this script)
•	2 VSANs (A and B fabrics)
•	4 vHBA Templates (for 4 vHBA's - only two are used by the template created in this script)
•	SAN Boot Policy
•	Server Pool
•	Server Pool Policy
•	BIOS Policy
•	Host Firmware Package Policy
•	Management Firmware Package Policy
•	Maintenance Policy
•	Local Disk Configuration Policy
•	Service Profile Template
There are two parameters that can be used at script launch.  The first is –ExcelFile, followed by the name of the Excel file. This file must be in the same directory as this script.  The second is -ToConsole. If specified, the output will be printed to the console.  If the –ToConsole switch is not used, then only the name of the Excel Configuration file needs to be specified (as shown in the second example below).
Examples:
./UCS_Config_CitrixXDBP.ps1 -ExcelFile XenDesktopBP.xlsx -ToConsole
This reads the XenDesktopBP.xlsx file and outputs the status of configuration to the Powershell console.

./UCS_Config_CitrixXDBP.ps1 XenDesktopBP.xlsx
This reads XenDesktopBP.xlsx file. An output file, UCSM_Configuration_Script_Log.txt, is created.
Before using the script, a few prerequisites must be met.
•	Cisco UCS PowerTool Installation
o	You must first download and install the Cisco UCS PowerTool from http://developer.cisco.com.
o	Click on the Technologies menu item, and then select Unified Computing.  From here, click on PowerShell Downloads, and then CiscoUcs-PowerTool-0.9.10.1.zip.
o	Within a Microsoft Windows environment (this environment must also include the installation of Microsoft Excel), extract the files, and then run the CiscoUcs-PowerTool-0.9.10.1 application.  By default, a desktop icon for the UCS PowerTool will be installed to the desktop.
•	Edit the Excel Configuration file
o	Save the UCS_Config_CitrixXDBP.ps1 script and XenDesktopBP.xlsx files to a location accessible from the Windows host where the UCS PowerTool was installed.
o	Open the XenDesktopBP.xlsx file and edit the following values:
•	UCS Manager IP Address
•	Management IP Pool starting and ending IP addresses, subnet mask, and default gateway (these IP addresses are assigned to individual server blades for management purposes (1 per blade) and must reside in the same subnet as the UCS Manager IP Address
•	Site Information
•	Service Profile Template Name (single word) and Description
•	VLAN information for Management network, Virtual Desktop Data network, VM Motion network (ie., vMotion), and IP-based storage network
•	SAN Boot Target Information (if SAN boot won’t be used, the defaults in the spreadsheet can be used)
o	Save this file in the same directory as the script is located
•	Launch the UCS PowerTool
o	Using the desktop icon created during installation, launch a PowerShell window.
•	By default, PowerShell will be launched with an ExecutionPolicy of RemoteSigned.  The included UCS_Config_CitrixXDBP.ps1 script has been signed, so the script should run using the default policy.  However, if any changes are made to the script, you may need to change the ExecutionPolicy to Unrestricted to allow the script to run.
o	Change to the directory where the script and Excel Configuration file are located (ie., cd .\CitrixXDBP)
o	Execute the script, supplying the name of the Excel Configuration file and, optionally, a switch to output messages to the console.  Example:
./UCS_Config_CitrixXDBP.ps1 -ExcelFile XenDesktopBP.xlsx –ToConsole
•	During the execution of the script, you will be prompted for the username and password used for UCS Manager
•	Modify pool and policy values
o	BIOS Policy
•	The BIOS Policy created with this script assumes that performance is more important than power usage.  If this is not the case, modify the CitrixHD_Host BIOS Policy with desired values.
o	Boot Policy
•	A SAN_boot Policy is created using values contained within the Excel Configuration file.  If SAN Boot will not be utilized, an alternative boot policy should be created, and the CitrixXDTemplate Service Profile Template should be modified to utilize the new boot policy.
o	Host and Management Firmware Packages
•	Host and Management Firmware Packages were created and added to the CitrixXDTemplate Service Profile Template, however, no firmware has been added to these packages.  Please open and modify CitrixXD Host and Management Firmware packages as necessary to ensure the proper firmware is used in the environment.
o	Server Pool
•	A Server Pool named CitrixXD_Server_Pool was created and added to the CitrixXDTemplate Service Profile Template, however, no servers have been added to the pool.  Please open and modify the pool as necessary to include the server blades that will be used to host Citrix XenDesktop Hosted Virtual Desktop machines.
•	Create Service Profiles from Template
o	From the Servers tab within UCS Manager, click on the Service Profile Template named CitrixXDTemplate
o	In the Actions pane, click on Create Service Profiles From Template
o	Assign a naming prefix to be used for the new Service Profiles, and indicate the number of new Service Profiles you’d like to be created.
That’s it! You should now have some number of Service Profiles created and ready to be used in a Citrix XenDesktop environment.

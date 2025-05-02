# Get-ExchangeHealth
PowerShell Script to Get Exchange Server and Services Health Status

For full description, usage information, and sample output, please follow this link:
http://www.lazyexchangeadmin.com/2015/03/database-backup-and-disk-space-report.html

	.NOTES
	===========================================================================
	 Created on:   	8-Aug-2014 9:10 AM
	 Created by:   	Tito D. Castillote Jr.
					june.castillote@gmail.com
	 Filename:     	Get-ExchangeHealth.ps1
	 Version:		5.2 (22-Mar-2017)
	===========================================================================

	.LINK
		http://www.lazyexchangeadmin.com/2015/03/database-backup-and-disk-space-report.html

	.SYNOPSIS
		Use Get-ExchangeHealth.ps1 for gathering and reporting the overall Exchange Server health.

	.DESCRIPTION
		Test and report include:
		* Server Health (Up Time, Server Roles Services, Mail flow,...)
		* Mailbox Database Status (Mounted, Backup, Size and Space, Mailbox Count, Paths,...)
		* Public Folder Database Status (Mount, Backup, Size and Space,...)
		* Database Copy Status
		* Database Replication Status
		* Mail Queue
		* Disk Space
		* Server Components (for Exchange 2013/2016)
		
	.PARAMETER configFile
		Required switch to specify the file name of the configuration XML file to use (e.g. config.xml)

	.PARAMETER enableDebug
		Optional switch to enable logging
	
	.EXAMPLE
		.\Get-ExchangeHealth.ps1 -configFile .\config.xml

		This will read the configuration from config.xml and perform the enabled tests, create report, send via email - if enabled.

1) Dismount database
2) Suspend copies on all database copies
3) Physically move the files to the new location
4) Run the following command
	Move-DatabasePath "Test" -LogFolderPath:F:\Exchange\Logs\Test\ -EdbFilePath:D:\Exchange\Mailbox\Test\test.edb -configurationonly
5) Mount the database
6) Resume all database copies
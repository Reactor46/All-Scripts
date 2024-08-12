This Tool provided is intended to make it simple and customizable, with little amount of PowerShell knowledge.
The tool consists of several files and directories
.\Main.ps1				
.\Template.doc
.\Includes\ADsettings.ps1
.\Includes\Application.ps1
.\Includes\Exchange.ps1
.\Includes\WMICollection.ps1
.\Includes\ExchangeContent.xml
.\Includes\wmiContent.XML
.\Includes\Formats.ps1

The ps1 files contains all the logic and functions used. The .csv files contains the content being used.

So what is required?
•	PowerShell 2.0 or higher
•	The machine collecting information requires the Active PowerShell modules
•	The Document generation machine “if different requires” Word 2007 or higher.
•	Your script Execution policy might be required to be changed using Set-ExecutionPolicy

Running it (“Collecting information and building document”)
Note: All files are required for both processes to complete successfully
1.	Execute main.ps1 with PowerShell in Administrators Mode
	"A form windows  will appear"
2.	\Main.ps1 -CollectInformation -ExchangeServer <ServerName> 
	This will collect all information necessary to build the document. The script will create a directory .\Data\ if not already created and populate it with the extracted information.
3.	Once this process is complete, you can Exit or run ".\Main.ps1 -GenerateReport". IF the machine collecting the information does not have word installed. 
4.	Copy the "Data" directory over to the machine where you want the document generated and into the root where the main.ps1 and template.doc files are located. 
5.	Run the main.ps1 once more.
6.	".\Main.ps1 -GenerateReport".
7.	Once it done, a new file would have been generated in the root directory with the a time stamped name. 
8.	“It’s done, your document is build”.

How to modify the content.
The Content can easily be modified by simply editing the xml files. If you do not require section in the ford, simply remove from the xml file. 

Here how the xml files work.
ExchangeContent.xml
The objects contains 8 properties
1.	Index	- Index provides the order in which data will be collected and written to the document
2.	Heading – Contains the text at the top of the paragraph or section
3.	CallFunc – This is the PowerShell function it executes to collect information. If you only want to add text and execute no function indicate it with “[]” as empty
4.	HeadingFormat – This is the formatting in word for the heading. For Example “Heading 1, Heading 2, Heading 3”
5.	TextPara -  This is the text that will be in the paragraph section above the tables with collected information
6.	HeaderDirection – This is the direction of the Table Headers in word. Options available is 
		“0” – Horizontal
		“1” – From top Down
		“2” – From bottom up
7.	HeaderHeight – This is the height of the header column of the Table
8.	Export – The filename to which the data should be export. This is a mandatory field where code is executed.

WMIContent.xml
This objects contains 9 properties
1.	Index	- Index provides the order in which data will be collected and written to the document
2.	Class	- This section contains the WMI query for the Class you want to collect data from. For example “Select * from win32_Bios”. If you only want to add text and execute no function indicate it with “[]” as empty
3.	Property – This field indicates the properties you want to select from the class.	
4.	Heading – Contains the text at the top of the paragraph or section
5.	TextPara - This is the text that will be in the paragraph section above the tables with collected information
6.	HeadingFormat	- This is the formatting in word for the heading. For Example “Heading 1, Heading 2, Heading 3”
7.	HeaderDirection – This is the direction of the Table Headers in word. Options available is 
		“0” – Horizontal
		“1” – From top Down
		“2” – From bottom up	
8.	HeaderHeight	– This is the height of the header column of the Table
9.	Export - The filename to which the data should be export. This is a mandatory field where code executed.

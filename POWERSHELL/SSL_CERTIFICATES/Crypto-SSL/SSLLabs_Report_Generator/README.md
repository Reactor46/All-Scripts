# SSLLabs_Report_Generator
A repos to parse the information from a range of SSL Server Test results and present them in a CSV format

//IDENTIFY DOMAINS TO BE SCANNED
Create domains.txt file – This is a list of domains separated by newlines that you want to SSLLabs to scan for Cipher Suite information

//GET SSL LABS TO CACHE RESULTS FOR DOMAINS
Get “Check-SLLConfig.ps1” (from Github (https://github.com/damosan)
Put in same folder as domains.txt
Use command “$creds = get-credential”, select your account
Run .\Check-SSLConfig.ps1 .\domains.txt -Cache $true -Publish $true  , this command will take a while, approx. 2 mins per domain the list.

*if server returns error 407: Proxy Authentication Required. Visit one of the domains in a browser and re-run the command*

When this is finished it generates a simple .csv report in the same folder that is picked up and used by the next script
NOTE: Check this CSV files for any errors and document domains that may fail, as they may be omitted in the final report by the next scripts.

//COLLECT THE HTML DATA
Use Script to scrape results web pages (forEachDomainGetCipherDetails.ps1)
Scans will be created in the current directory.

//PARSE DATA
Copy the scans to a new directory in Linux, create and run the bash script generateReportFromRawScans.sh

//REVIEW REPORT
This will generate the file report.txt, which can be saved as a csv file and viewed in excel.

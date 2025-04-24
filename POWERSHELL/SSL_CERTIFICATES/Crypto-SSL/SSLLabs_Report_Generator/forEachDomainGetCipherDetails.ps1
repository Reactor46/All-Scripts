#This file pulled Cached scans .html files in Qualys SSLLABS,  
#parses the scans for the Cipher Suite information
#outputs each raw scan html to a file
#it sets the domain name as the first line of the file, the grade as the second
#then appends the raw scan

$tempFileName = "currentScan.txt"

#read in each domain to be scanned
$hostlist = Get-Content .\domains.txt
$scanFilenameNumber = 1

##open csv from SSLLAB Scans
$DATE_STRING = (get-date -Format "yyyy-MM-dd").ToString()
$csvResultsFilepath = ".\$DATE_STRING SSL Scan Results.csv"
$openCSV = Import-Csv $csvResultsFilepath


foreach ($domain in $hostlist)
{
    
        echo "$scanFilenameNumber : '$domain'"
    
        $scanFilename = "scanNumber-$scanFilenameNumber.txt"

        $uri = "https://www.ssllabs.com/ssltest/analyze.html?d=$domain"
        $scanRes = 0
        $scanRes = Invoke-WebRequest $uri -UseDefaultCredentials
        



        ##add URL and Grade to top of cipher suite output

        echo $domain > $scanFilename
             
        $domainGrade = 0

        ##get grade
        foreach ($row in $openCSV)
        {
            if($row.Domain -eq $domain)
                {
                    $domainGrade = $row.Grade 
                 }
        }


        

        #get part of html with cipher suite info
        $scanRes.ParsedHtml.body | Out-File $scanFilename -Encoding utf8 ##ensure utf8

        #create scan in utf format
        
        @("$domain", "$domainGrade") +  (Get-Content $scanFilename) | Set-Content $scanFilename

        $scanFilenameNumber++
        $count++
        
        ##delay to prevent cumulative error ...???
        $delay = 4
        $i=1
        <#
        Write-Host "Waiting.." -NoNewline
        while($i -le $delay)
            {
                Write-Host "$i" -NoNewline
                Start-Sleep -Seconds 1
                $i++
            }
        #>
        Write-Host "`n"
}

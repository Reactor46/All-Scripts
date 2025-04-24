# Script name:   	check_ms_sharepoint_health.ps1
# Version:		v1.01.151110
# Created on:    	17/03/2014
# Purpose:       	Checks Microsoft SharePoint Heath. 
# On Github:		https://github.com/2Dman/check-ms-sharepoint-health/			
# Copyright:
#	This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published
#	by the Free Software Foundation, either version 3 of the License, or (at your option) any later version. This program is distributed 
#	in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A 
#	PARTICULAR PURPOSE. See the GNU General Public License for more details. You should have received a copy of the GNU General Public 
#	License along with this program.  If not, see <http://www.gnu.org/licenses/>.

if ($PSVersionTable) {$Host.Runspace.ThreadOptions = 'ReuseThread'}

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

$Status = 3
$ReportsList = [Microsoft.SharePoint.Administration.Health.SPHealthReportsList]::Local
$FormUrl = '{0}{1}?id=' -f $ReportList.ParentWeb.Url, $ReportsList.Forms.List.DefaultDisplayFormUrl

$ReportProblems = $ReportsList.Items | Where-Object {$_['Severity'] -ne '4 - Success'} | ForEach-Object {
    New-Object PSObject -Property @{
        Url = "<a href='$FormUrl$($_.ID)'>$($_['Title'])</a>"
        Severity = $_['Category']
        Explanation = $_['Explanation']
        Modified = $_['Modified']
        FailingServers = $_['Failing Servers']
        FailingServices = $_['Failing Services']
        Remedy = $_['Remedy']
    }
} 

if ($ReportProblems.count -gt "0") {
	$ErrorMessage = "SharePoint Health Analyzer detected problems:"
	foreach($ReportProblem in $ReportProblems) {
		if ($ReportProblem.FailingServers) {
			$ServerString = " on " + $ReportProblem.FailingServers -replace "(?m)[`n`r]+",""
		}
		else {
			$ServerString = ""
		}
		$ErrorMessage += "`n Service $($ReportProblem.FailingServices)$ServerString, modified $($ReportProblem.Modified), error message $($($ReportProblem.Explanation).substring(0,150))"
	}
	Write-Host $ErrorMessage
	$Status = 2
}
else {
	Write-Host "No SharePoint health problems detected!"
	$Status = 0
}

exit $Status
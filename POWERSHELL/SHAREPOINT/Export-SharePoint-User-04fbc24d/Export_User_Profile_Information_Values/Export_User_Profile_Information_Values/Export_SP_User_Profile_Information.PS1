# -----------------------------------------------------------------------------
# Script	: To Export User Profile Information Value in Excel and HTML
# Author	: Chendrayan Venkatesan
# Company	: Tata Consultancy Services
# Date		: October 26 2012
# Version	: 1.0
# -----------------------------------------------------------------------------


#. Use Internal name of the user profile property.
#. Any custom user fields will have "SPS-".
#. Remove comment "#" in line 43 and do the same in line 48 if you need HTML output.
#. Provide location after out-file to save the HTML file.
#. In Line 59 use the same location as Out-File.
#. Append if you have any custom fields in mysite user profile property
#. For HTML look and feel please copy and paste the style.css in C:\.
#. Use any style.css as per organization standard and policy.

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

$siteUrl = "Your Mysite Host Name"
$outputFile = "Save the Output File in your desired Location"

$serviceContext = Get-SPServiceContext -Site $siteUrl
$profileManager = New-Object Microsoft.Office.Server.UserProfiles.UserProfileManager($serviceContext);
$profiles = $profileManager.GetEnumerator()

$collection = @()
foreach ($profile in $profiles) {
 
   $profileData = "" | 
   select "AccountName", "PreferredName" , "Department" , "Manager" , "Office" , "Location" , "WorkEmail" , "Assistant" , "AboutMe" , "Language" , "PictureURL" , "Role"
   
   $profileData.AccountName = $profile["AccountName"]
   $profileData.PreferredName = $profile["PreferredName"]
   $profileData.Manager = $profile["Manager"]
   $profileData.Department = $profile["Department"]
   $profileData.Office = $profile["Office"]
   $profileData.Location = $profile["Location"]
   $profileData.WorkEmail = $profile["WorkEmail"]
   $profileData.Assistant = $profile["Assistant"]
   $profileData.AboutMe = $profile["AboutMe"].Value
   $profileData.Language = $profile["Language"]
   $profileData.PictureURL = $profile["PictureURL"]
   $profileData.Role = $profile["Role"]
   
   #$collection += $profileData | ConvertTo-Html -Fragment
   $collection += $profileData
}

#ConvertTo-Html -Body "$collection" -CssUri C:\style.CSS | Out-File "Location to save"

$collection | Export-Csv $outputFile -NoTypeInformation

#Send Mail to Box for easy access
$SMTP = "Name of the SMTP Server"
$From = "Email Address"
$To = "Email Address" , "Email Address"
$Subject = "User Profile Information Report"
Send-MailMessage -From $From -To $To -SmtpServer $SMTP -Subject $Subject -Attachments "Same as Out-File Location"
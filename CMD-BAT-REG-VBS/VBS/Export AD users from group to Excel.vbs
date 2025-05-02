' Extract a list of users from a specific group in AD into Excel.
Dim objDistList, objExcel, ExcelRow, strUser, strDistListName, strOU
' This line specifies the group name, OU, and AD domain name, edit to suit your system.
Set objDistList = GetObject("LDAP://CN=All Employees,OU=Distribution Groups,DC=corp,DC=your_company,DC=com")

Set objExcel = CreateObject("Excel.Application")
With objExcel
 .SheetsInNewWorkbook = 1
 .Workbooks.Add
 .Visible = True
 .Worksheets.Item(1).Name = mid(objDistList.Name, _
 instr(1,objDistList.Name,"=") + 1 )
 ExcelRow = 1
 ' This section sets the Excel header row names, these can be changed to anything more human readable if using this script to simply extract a list.
 ' If the header names are left as is, the resulting Excel file can be edited, saved as CSV, and used by an AD import tool to do bulk updates.
 ' Note if using this to do a bulk update, format every cell as text. Also Excel does weird things with phone numbers if you re-open the saved CSV file with Excel.
 ' http://www.wisesoft.co.uk/software/bulkadusers/default.aspx Free bulk AD update tool (download link top right).
 ' Outlook uses these fields in the address book and contact properties. If all this stuff is filled in it makes the Outlook address book a very handy tool.
 ' Android and iOS Exchange email clients will also read this information into their addressbooks.
 ' http://www.wisesoft.co.uk/scripts/activedirectoryschema.aspx Clickable interface to see what all the LDAP attribute names relate to in the user properties fields.
 ' Edit / remove / change the order as you please, make sure it matches up with the next section.
 .Cells(ExcelRow, 1) = "sn" ' Last name - General tab.
 .Cells(ExcelRow, 2) = "givenName" ' First name - General tab.
 .Cells(ExcelRow, 3) = "description" ' Description - General tab.
 .Cells(ExcelRow, 4) = "physicalDeliveryOfficeName" ' Office - General tab.
 .Cells(ExcelRow, 5) = "telephoneNumber" ' Telephone number - General tab.
 .Cells(ExcelRow, 6) = "mail" ' E-mail - General tab.
 .Cells(ExcelRow, 7) = "wWWHomePage" ' Web page - General tab.
 .Cells(ExcelRow, 8) = "homePhone" ' Home number - Telephones tab.
 .Cells(ExcelRow, 9) = "mobile" ' Mobile number - Telephones tab.
 .Cells(ExcelRow, 10) = "facsimileTelephoneNumber" ' Fax number - Telephones tab.
 .Cells(ExcelRow, 11) = "ipPhone" ' IP Phone - Telephones tab (Some phone systems that integrate with AD use this).
 .Cells(ExcelRow, 12) = "title" ' Job title - Organization tab (Outlook shows this field in the address book and contact properties).
 .Cells(ExcelRow, 13) = "department" ' Department - Organization tab.
 .Cells(ExcelRow, 14) = "manager" ' Manager - Organization tab (Outlook uses this to show an employees manager, and the managers direct reports).
 .Cells(ExcelRow, 15) = "company" ' Company - Organization tab.
 .Cells(ExcelRow, 16) = "streetAddress" ' Street - Address tab.
 .Cells(ExcelRow, 17) = "l" ' City - Address tab.
 .Cells(ExcelRow, 18) = "st" ' State / province - Address tab.
 .Cells(ExcelRow, 19) = "postalCode" ' Zip / postal code - Address tab.
 .Cells(ExcelRow, 20) = "co" ' Country - Address tab.
 .Cells(ExcelRow, 21) = "sAMAccountName" ' User login name - Account tab. This field is often used by AD import tools to identify the account to update.
 
 .Rows(1).Font.Bold = True
 
 ExcelRow = ExcelRow + 1
 For Each strUser in objDistList.Member
 Set objUser =  GetObject("LDAP://" & strUser)
 ' LDAP attribute names read from Active Directory.
 .Cells(ExcelRow, 1) = objUser.sn
 .Cells(ExcelRow, 2) = objUser.givenName
 .Cells(ExcelRow, 3) = objUser.description
 .Cells(ExcelRow, 4) = objUser.physicalDeliveryOfficeName
 .Cells(ExcelRow, 5) = objUser.telephoneNumber
 .Cells(ExcelRow, 6) = objUser.mail
 .Cells(ExcelRow, 7) = objUser.wWWHomePage
 .Cells(ExcelRow, 8) = objUser.homePhone
 .Cells(ExcelRow, 9) = objUser.mobile
 .Cells(ExcelRow, 10) = objUser.facsimileTelephoneNumber
 .Cells(ExcelRow, 11) = objUser.title
 .Cells(ExcelRow, 12) = objUser.ipPhone
 .Cells(ExcelRow, 13) = objUser.department
 .Cells(ExcelRow, 14) = objUser.manager
 .Cells(ExcelRow, 15) = objUser.company
 .Cells(ExcelRow, 16) = objUser.streetAddress
 .Cells(ExcelRow, 17) = objUser.l
 .Cells(ExcelRow, 18) = objUser.st 
 .Cells(ExcelRow, 19) = objUser.postalCode
 .Cells(ExcelRow, 20) = objUser.co
 .Cells(ExcelRow, 21) = objUser.sAMAccountName
 
  ExcelRow = ExcelRow + 1
 Next
 ' Auto fit the columns.
 .Columns(1).entirecolumn.autofit
 .Columns(2).entirecolumn.autofit
 .Columns(3).entirecolumn.autofit
 .Columns(4).entirecolumn.autofit
 .Columns(5).entirecolumn.autofit
 .Columns(6).entirecolumn.autofit
 .Columns(7).entirecolumn.autofit
 .Columns(8).entirecolumn.autofit
 .Columns(9).entirecolumn.autofit
 .Columns(10).entirecolumn.autofit
 .Columns(12).entirecolumn.autofit
 .Columns(13).entirecolumn.autofit
 .Columns(14).entirecolumn.autofit
 .Columns(15).entirecolumn.autofit
 .Columns(16).entirecolumn.autofit
 .Columns(17).entirecolumn.autofit
 .Columns(18).entirecolumn.autofit
 .Columns(19).entirecolumn.autofit
 .Columns(20).entirecolumn.autofit
 .Columns(21).entirecolumn.autofit
 End With

Set objExcel = Nothing
Set objDistList = Nothing
Wscript.Quit
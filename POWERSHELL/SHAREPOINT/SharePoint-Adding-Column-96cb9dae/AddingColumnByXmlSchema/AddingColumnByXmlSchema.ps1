# Check to ensure Microsoft.SharePoint.PowerShell is loaded
if((Get-PSSnapin | Where {$_.Name -eq "Microsoft.SharePoint.PowerShell"}) -eq $null) {
    Add-PSSnapin Microsoft.SharePoint.PowerShell;
}

# Configuration
$siteUrl = "<Your Site>"
$groupName = "<Your Group Name>"

# Connection of the site
$web = get-spweb $siteUrl

# Adding Column to the site

# Single line of text
$fieldXML = '<Field Type="Text" Name="TextCol" DisplayName="TextCol" Group="' + $groupName + '" Hidden="FALSE" Required="FALSE" ShowInDisplayForm="TRUE" ShowInEditForm="TRUE" ShowInListSettings="TRUE" ShowInNewForm="TRUE"></Field>'   
$web.Fields.AddFieldAsXml($fieldXML)

# Multiple line of test - Plain Text
$fieldXML = '<Field Type="Note" Name="PlainTextCol" DisplayName="PlainTextCol" NumLines="6" RichText="FALSE" Group="' + $groupName + '" Hidden="FALSE" Required="FALSE" ShowInDisplayForm="TRUE" ShowInEditForm="TRUE" ShowInListSettings="TRUE" ShowInNewForm="TRUE"></Field>'
$web.Fields.AddFieldAsXml($fieldXML)

# Multiple line of test - Rich Text
$fieldXML = '<Field Type="Note" Name="RichTextCol" DisplayName="RichTextCol" NumLines="6" RichText="TRUE" RichTextMode="Compatible" Group="' + $groupName + '" Hidden="FALSE" Required="FALSE" ShowInDisplayForm="TRUE" ShowInEditForm="TRUE" ShowInListSettings="TRUE" ShowInNewForm="TRUE"></Field>'
$web.Fields.AddFieldAsXml($fieldXML)

# Multiple line of test - Enhanced Rich Text
$fieldXML = '<Field Type="Note" Name="EnhancedRichTextCol" DisplayName="EnhancedRichTextCol" NumLines="6" RichText="TRUE" RichTextMode="FullHtml" IsolateStyles="TRUE" Group="' + $groupName + '" Hidden="FALSE" Required="FALSE" ShowInDisplayForm="TRUE" ShowInEditForm="TRUE" ShowInListSettings="TRUE" ShowInNewForm="TRUE"></Field>'
$web.Fields.AddFieldAsXml($fieldXML)

# Choice - Dropdown
$fieldXML = '<Field Type="Choice" Name="ChoiceDropDownCol" DisplayName="ChoiceDropDownCol" Format="Dropdown" FillInChoice="FALSE" Group="' + $groupName + '" Hidden="FALSE" Required="FALSE" ShowInDisplayForm="TRUE" ShowInEditForm="TRUE" ShowInListSettings="TRUE" ShowInNewForm="TRUE"><Default>Enter Choice 1</Default><CHOICES><CHOICE>Enter Choice 1</CHOICE><CHOICE>Enter Choice 2</CHOICE></CHOICES></Field>'   
$web.Fields.AddFieldAsXml($fieldXML)

# Choice - Radio Buttons
$fieldXML = '<Field Type="Choice" Name="ChoiceRadioButtonsCol" DisplayName="ChoiceRadioButtonsCol" Format="RadioButtons" FillInChoice="FALSE" Group="' + $groupName + '" Hidden="FALSE" Required="FALSE" ShowInDisplayForm="TRUE" ShowInEditForm="TRUE" ShowInListSettings="TRUE" ShowInNewForm="TRUE"><Default>Enter Choice 1</Default><CHOICES><CHOICE>Enter Choice 1</CHOICE><CHOICE>Enter Choice 2</CHOICE></CHOICES></Field>'   
$web.Fields.AddFieldAsXml($fieldXML)

# Choice - Checkboxes
$fieldXML = '<Field Type="MultiChoice" Name="ChoiceCheckboxesCol" DisplayName="ChoiceCheckboxesCol" FillInChoice="FALSE" Group="' + $groupName + '" Hidden="FALSE" Required="FALSE" ShowInDisplayForm="TRUE" ShowInEditForm="TRUE" ShowInListSettings="TRUE" ShowInNewForm="TRUE"><Default>Enter Choice 1</Default><CHOICES><CHOICE>Enter Choice 1</CHOICE><CHOICE>Enter Choice 2</CHOICE></CHOICES></Field>'   
$web.Fields.AddFieldAsXml($fieldXML)

# Number
$fieldXML = '<Field Type="Number" Name="NumberCol" DisplayName="NumberCol" Decimals="2" Min="0" Max="1000" Group="' + $groupName + '" Hidden="FALSE" Required="FALSE" ShowInDisplayForm="TRUE" ShowInEditForm="TRUE" ShowInListSettings="TRUE" ShowInNewForm="TRUE"></Field>'   
$web.Fields.AddFieldAsXml($fieldXML)

# Currency
$fieldXML = '<Field Type="Currency" Name="CurrencyCol" DisplayName="CurrencyCol" Decimals="2" LCID="1033" Group="' + $groupName + '" Hidden="FALSE" Required="FALSE" ShowInDisplayForm="TRUE" ShowInEditForm="TRUE" ShowInListSettings="TRUE" ShowInNewForm="TRUE"></Field>'   
$web.Fields.AddFieldAsXml($fieldXML)


# Date - Date Only
$fieldXml = '<Field Type="DateTime" Name="DateOnlyCol" DisplayName="DateOnlyCol" Format= "DateOnly" Group="' + $groupName + '" Hidden="FALSE" Required="FALSE" ShowInDisplayForm="TRUE" ShowInEditForm="TRUE" ShowInListSettings="TRUE" ShowInNewForm="TRUE"></Field>'
$web.Fields.AddFieldAsXml($fieldXML)

# Date - Date and Time
$fieldXml = '<Field Type="DateTime" Name="DateTimeCol" DisplayName="DateTimeCol" Format= "DateOTime" Group="' + $groupName + '" Hidden="FALSE" Required="FALSE" ShowInDisplayForm="TRUE" ShowInEditForm="TRUE" ShowInListSettings="TRUE" ShowInNewForm="TRUE"></Field>'
$web.Fields.AddFieldAsXml($fieldXML)

# Lookup
$urlListLookup = $siteUrl + "/Shared%20Documents/"
$list = $web.GetList($urlListLookup)
$listID = $list.Id
$webID = $web.Id
$fieldXml = '<Field Type="Lookup" Name="LookupCol" DisplayName="LookupCol" Group="' + $groupName + '" Hidden="FALSE" Required="FALSE" EnforceUniqueValues="FALSE" List="' + $listID + '" WebId="'+ $webID +'" ShowField="Title" ShowInDisplayForm="TRUE" ShowInEditForm="TRUE" ShowInListSettings="TRUE" ShowInNewForm="TRUE" />'
$web.Fields.AddFieldAsXml($fieldXML)

# Yes/No
$fieldXml = '<Field Type="Boolean" Name="YesNoCol" DisplayName="YesNoCol" Group="' + $groupName + '" Hidden="FALSE" Required="FALSE" ShowInDisplayForm="TRUE" ShowInEditForm="TRUE" ShowInListSettings="TRUE" ShowInNewForm="TRUE"><Default>1</Default></Field>'
$web.Fields.AddFieldAsXml($fieldXML)

# Simple user
$fieldXml = '<Field Type="User" Name="SimpleUserCol" DisplayName="SimpleUserCol" UserSelectionMode="1" Group="' + $groupName + '" Hidden="FALSE" Required="FALSE" ShowInDisplayForm="TRUE" ShowInEditForm="TRUE" ShowInListSettings="TRUE" ShowInNewForm="TRUE"></Field>'
$web.Fields.AddFieldAsXml($fieldXML)

# Multiple user
$fieldXml = '<Field Type="UserMulti" Name="MultipleUserCol" DisplayName="MultipleUserCol" Mult="TRUE" Group="' + $groupName + '" Hidden="FALSE" Required="FALSE" ShowInDisplayForm="TRUE" ShowInEditForm="TRUE" ShowInListSettings="TRUE" ShowInNewForm="TRUE"></Field>'
$web.Fields.AddFieldAsXml($fieldXML)

# Hyperlink text
$fieldXml = '<Field Type="URL" Name="HyperlinkCol" DisplayName="HyperlinkCol" Format="Hyperlink" Group="' + $groupName + '" Hidden="FALSE" Required="FALSE" ShowInDisplayForm="TRUE" ShowInEditForm="TRUE" ShowInListSettings="TRUE" ShowInNewForm="TRUE"></Field>'
$web.Fields.AddFieldAsXml($fieldXML)

# Hyperlink image
$fieldXml = '<Field Type="URL" Name="HyperlinkImgCol" DisplayName="HyperlinkImgCol" Format="Image" Group="' + $groupName + '" Hidden="FALSE" Required="FALSE" ShowInDisplayForm="TRUE" ShowInEditForm="TRUE" ShowInListSettings="TRUE" ShowInNewForm="TRUE"></Field>'
$web.Fields.AddFieldAsXml($fieldXML)

# Calculated
$fieldXml = '<Field Type="URL" Name="CalculatedCol" DisplayName="CalculatedCol" ResultType="Number" Group="' + $groupName + '" Hidden="FALSE" Required="FALSE" ShowInDisplayForm="TRUE" ShowInEditForm="TRUE" ShowInListSettings="TRUE" ShowInNewForm="TRUE"><Formula>=NumberCol*5</Formula><FieldRefs><FieldRef Name="NumberCol"/></FieldRefs></Field>'
$web.Fields.AddFieldAsXml($fieldXML)

$web.Update()
$web.Dispose()
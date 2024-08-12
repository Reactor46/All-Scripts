<%@ Page Debug="true" %>
<%@Import namespace="System.Data"%>
<%@Import namespace="System.Net"%>
<%@Assembly name="System.DirectoryServices, Version=1.0.5000.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, Custom=null"%>
<%@Import namespace="System.DirectoryServices"%>
<%@Import namespace="System.Security.Cryptography.X509Certificates"%>
<%@Import namespace="System.Web.UI.WebControls"%> 
<script language="vb" runat="server">

Public Class AcceptAllCertificatePolicy
      Implements ICertificatePolicy
      Public Overridable Function CheckValidationResult(ByVal srvPoint As ServicePoint, ByVal certificate As X509Certificate, ByVal request As WebRequest, ByVal problem As Integer) As Boolean Implements ICertificatePolicy.CheckValidationResult
           Return True 'this accepts all certificates
      End Function
End Class


Sub Page_Load(sender As Object, e As EventArgs)
	System.Net.ServicePointManager.CertificatePolicy = New AcceptAllCertificatePolicy
        Dim Request As System.Net.HttpWebRequest
        Dim Response As System.Net.HttpWebResponse
        Dim strRootURI As String
        Dim strQuery As String
        Dim bytes() As Byte
        Dim workrow As System.Data.DataRow
        Dim resrow As System.Data.DataRow
        Dim impersonationContext As System.Security.Principal.WindowsImpersonationContext
        Dim currentWindowsIdentity As System.Security.Principal.WindowsIdentity
        currentWindowsIdentity = CType(User.Identity, System.Security.Principal.WindowsIdentity)
        impersonationContext = currentWindowsIdentity.Impersonate()
	Dim MyCredentialCache As System.Net.CredentialCache
        Dim resdataset As New System.Data.DataSet
        Dim RequestStream As System.IO.Stream
        Dim ResponseStream As System.IO.Stream
        Dim ResponseXmlDoc As System.Xml.XmlDocument
        Dim hrefNodes,SubjectNodes,FromnameNodes,x003D001ENodes,x0070001ENodes,datereceivedNodes As System.Xml.XmlNodeList
	Dim objsearch As New System.DirectoryServices.DirectorySearcher
        Dim strrootdse As String = objsearch.SearchRoot.Path
        Dim objdirentry As New system.DirectoryServices.DirectoryEntry(strrootdse)
        Dim objresult As system.DirectoryServices.SearchResult
        Dim stremailaddress As String
        Dim strhomeserver As String
	Dim ResDsSet as DataSet = New DataSet
        objsearch.Filter = "(&(&(&(& (mailnickname=*) (| (&(objectCategory=person)(objectClass=user)(|(homeMDB=*)" _
        & "(msExchHomeServerName=*))) )))(objectCategory=user)(userPrincipalName=*)(mailNickname=" & System.Environment.UserName & ")))"
        objsearch.SearchScope = DirectoryServices.SearchScope.Subtree
        objsearch.PropertiesToLoad.Add("mail")
        objsearch.PropertiesToLoad.Add("msExchHomeServerName")
        objsearch.Sort.Direction = DirectoryServices.SortDirection.Ascending
        objsearch.Sort.PropertyName = "mail"
        Dim colresults As DirectoryServices.SearchResultCollection = objsearch.FindAll()
        For Each objresult In colresults
            stremailaddress = objresult.GetDirectoryEntry().Properties("mail").Value
            strhomeserver = objresult.GetDirectoryEntry().Properties("msExchHomeServerName").Value
        Next
        Dim emailNameNodes As System.Xml.XmlNodeList
	strhomeserver = Right(strhomeserver, Len(strhomeserver) - (InStr(strhomeserver, "cn=Servers/cn=") + 13))
        strRootURI = "https://" & strhomeserver & "/exchange/" & stremailaddress & "/Inbox/"
        strQuery = "<?xml version=""1.0""?>" & _
            "<D:searchrequest xmlns:D = ""DAV:"" >" & _
            "<D:sql>SELECT ""urn:schemas:mailheader:subject"", ""urn:schemas:httpmail:fromname"", " & _
     " ""urn:schemas:httpmail:datereceived"" , ""http://schemas.microsoft.com/mapi/proptag/x003D001E"", " & _
            " ""http://schemas.microsoft.com/mapi/proptag/x0070001E"" FROM """ & strRootURI & """" & _
            "WHERE ""DAV:ishidden"" = false AND ""DAV:isfolder"" = false AND ""DAV:contentclass"" = 'urn:content-classes:message'" & _
             "</D:sql></D:searchrequest>"
        Request = CType(System.Net.WebRequest.Create(strRootURI), _
        System.Net.HttpWebRequest)
	Request.Credentials = System.Net.CredentialCache.DefaultCredentials
        Request.Method = "SEARCH"
        bytes = System.Text.Encoding.UTF8.GetBytes(strQuery)
        Request.ContentLength = bytes.Length
        RequestStream = Request.GetRequestStream()
        RequestStream.Write(bytes, 0, bytes.Length)
        RequestStream.Close()
        Request.ContentType = "text/xml"
	Request.AddRange("rows", 0,1000)
	Request.Headers.Add("Translate", "F")
        Response = CType(Request.GetResponse(), System.Net.HttpWebResponse)
        ResponseStream = Response.GetResponseStream()
        ResponseXmlDoc = New System.Xml.XmlDocument
        ResponseXmlDoc.Load(ResponseStream)
        ResDsSet = New System.Data.DataSet
	Dim resultstable As DataTable
	resultstable=new DataTable()
	resultstable.TableName = "queryres"
        Dim autoid As Data.DataColumn
        autoid = resultstable.Columns.Add("autoid", GetType(Int32))
        autoid.AutoIncrement = True
        autoid.AutoIncrementSeed = 1
        autoid.AutoIncrementStep = 1
	resultstable.Columns.Add("href")
	resultstable.Columns.Add("fromname")
	resultstable.Columns.Add("subject")
	resultstable.Columns.Add("x0070001E")
	resultstable.Columns.Add("x003D001E")
	resultstable.Columns.Add("Daterecieved")
	hrefNodes = ResponseXmlDoc.GetElementsByTagName("a:href")
	SubjectNodes = ResponseXmlDoc.GetElementsByTagName("d:subject")
        FromnameNodes = ResponseXmlDoc.GetElementsByTagName("e:fromname")
        x003D001ENodes = ResponseXmlDoc.GetElementsByTagName("f:x003D001E")
        x0070001ENodes = ResponseXmlDoc.GetElementsByTagName("f:x0070001E")
        datereceivedNodes = ResponseXmlDoc.GetElementsByTagName("e:datereceived")
        If SubjectNodes.Count > 0 Then
                Dim i As Integer
                For i = 0 To datereceivedNodes.Count - 1
                    AddRow(resultstable, hrefNodes(i).InnerText, FromnameNodes(i).InnerText, SubjectNodes(i).InnerText,x0070001ENodes(i).InnerText,x003D001ENodes(i).InnerText,datereceivedNodes(i).InnerText)
                Next
        End If
	ResDsSet.Tables.Add(resultstable)
        Dim relcol1, relcol2 As Data.DataColumn
        relcol1 = ResDsSet.Tables("queryres").Columns("subject")
        relcol2 = ResDsSet.Tables("queryres").Columns("x0070001E")
        Dim grprelations As Data.DataRelation
        grprelations = New Data.DataRelation("GroupEmails", relcol1, relcol2, False)
        ResDsSet.Relations.Add(grprelations)
        Dim fldView As DataView = New DataView(ResDsSet.Tables("queryres"), "x003D001E = ''", "", DataViewRowState.CurrentRows)
        Dim mailView As DataView    
        ConversationRepeater.DataSource = fldView
        ConversationRepeater.DataBind()



End Sub

Sub AddRow(resultstable As DataTable,href as string, FromName As String, Subject As String, x0070001E as string,x003D001E as string, Daterecieved as string )
	Dim row As DataRow
	row=resultstable.NewRow()
	row("href")=href
	row("fromname")=FromName
	row("Subject")=Subject
	row("x0070001E")=x0070001E
	row("x003D001E")=x003D001E
        Dim Daterecieved1 As System.DateTime
        Daterecieved1 = CDate(Daterecieved)
	row("Daterecieved")=Daterecieved1.ToShortDateString & " " & Daterecieved1.ToShortTimeString
	resultstable.Rows.Add(row)
End Sub
</script>

<script language="JavaScript">
  function ToggleDisplay(id)
  {
    var elem = document.getElementById('d' + id);
    if (elem) 
    {
      if (elem.style.display != 'block') 
      {
        elem.style.display = 'block';
        elem.style.visibility = 'visible';
      } 
      else
      {
        elem.style.display = 'none';
        elem.style.visibility = 'hidden';
      }
    }
  }
</script>

<style>
    .header {  }
    .details { display:none; visibility:hidden; background-color:#CCFFFF; 
               font-family: Verdana; }
</style>

<asp:Repeater id="ConversationRepeater" runat="server">
     <HeaderTemplate>
        <table width="100%" border="2" style="border-collapse: collapse; font-family: v; color: #FFFFFF; font-size: 10pt; font-weight: bold" bordercolor="#111111" width="100%">
          <tr style="background-color:0000FF">
            <th width="20%">
             Started By
            </th>
            <th>
              Conversation Subject </font>
            </th>
            <th width="13%">
             Started On
            </th>
	</table>
      </HeaderTemplate>
   <ItemTemplate>
  <table width="100%" style="border-collapse: collapse; font-family: v; color: #FFFFFF; font-size: 10pt; font-weight: bold" bordercolor="#111111" width="100%">
  <tr style="background-color:0000FF">
   <td width="20%">
       <%# ctype(Container.DataItem, DataRowView)("fromname") %>
   </td>
   <td>
     <div id='h<%# ctype(Container.DataItem, DataRowView)("autoid") %>' class="header"
          onclick='ToggleDisplay(<%# ctype(Container.DataItem, DataRowView)("autoid") %>);'>
       <%# ctype(Container.DataItem, DataRowView)("subject") %>
     </div>
    </td>
   <td width="13%">
       <%# ctype(Container.DataItem, DataRowView)("Daterecieved") %>
   </td>
   </table>
     <div id='d<%# ctype(Container.DataItem, DataRowView)("autoid") %>' class="details">
<table border="1" style="font: 8pt verdana" width="100%">
<tr style="background-color:#CCFFFF">
<asp:Repeater ID="DetailRepeater" DataSource='<%# ctype(Container.DataItem, DataRowView).CreateChildView("GroupEmails")%>' Runat="server">        
      <ItemTemplate>	
	<tr>
        <td width="20%"><b>From:</b>
       <%# ctype(Container.DataItem, DataRowView)("fromname") %>
        <td>
	<a href="<%# ctype(Container.DataItem, DataRowView)("href") %>"TARGET="_blank"><%# ctype(Container.DataItem, DataRowView)("subject") %></a>
	</tr>
      <td width="13%">
         <%# ctype(Container.DataItem, DataRowView)("Daterecieved") %>
     </td>
      </ItemTemplate>
</asp:Repeater>
</table>
	</div>
   </ItemTemplate>
</asp:Repeater>




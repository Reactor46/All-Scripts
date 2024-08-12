<%@ Page %>
<%@Import namespace="System.Data"%>
<%@Import namespace="System.Net"%>
<%@Import namespace="System.Xml"%>
<%@Import namespace="System.Xml.XPath"%>
<%@Import namespace="System.Xml.Xsl"%>
<%@Import namespace="System.Web"%>
<%@Assembly name="System.DirectoryServices, Version=1.0.5000.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, Custom=null"%>
<%@Import namespace="System.DirectoryServices"%>
<%@Import namespace="System.Security.Cryptography.X509Certificates"%>
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
	Dim objsearch As New System.DirectoryServices.DirectorySearcher
        Dim strrootdse As String = objsearch.SearchRoot.Path
        Dim objdirentry As New system.DirectoryServices.DirectoryEntry(strrootdse)
        Dim objresult As system.DirectoryServices.SearchResult
        Dim stremailaddress As String
        Dim strhomeserver As String
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
	strhomeserver = Right(strhomeserver, Len(strhomeserver) - (InStr(strhomeserver, "cn=Servers/cn=") + 13))
        strRootURI = "https://" & strhomeserver & "/exchange/" & stremailaddress & "/contacts/"
        strQuery = "<?xml version=""1.0""?>" & _
           "<D:searchrequest xmlns:D = ""DAV:"" >" & _
           "<D:sql>SELECT ""urn:schemas:contacts:cn"", ""http://schemas.microsoft.com/mapi/email1emailaddress"" " & _
           "FROM """ & strRootURI & """" & _
           "WHERE ""DAV:ishidden"" = false AND ""DAV:isfolder"" = false AND ""DAV:contentclass"" = 'urn:content-classes:person'" & _
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
        Request.Headers.Add("Translate", "F")
        Response = CType(Request.GetResponse(), System.Net.HttpWebResponse)
        ResponseStream = Response.GetResponseStream()
        ResponseXmlDoc = New System.Xml.XmlDocument
        ResponseXmlDoc.Load(ResponseStream)
	Dim myXslTransform As XslTransform
	myXslTransform = New XslTransform()
        myXslTransform.Load(Server.MapPath("showcontacts.xsl"))
	myXslTransform.Transform(ResponseXmlDoc,Nothing,Page.Response.Output)
End Sub
</script>
<HTML>
	<body>
	</body>
</HTML>


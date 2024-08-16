using System;
using System.Web;
using System.Web.Services;
using System.Web.Services.Protocols;
using System.DirectoryServices;
using System.Xml;
using System.IO;

[WebService(Namespace = "http://msgdev.mvps.org/ADhCardv2/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
public class Service : System.Web.Services.WebService
{
    public Service () {

        //Uncomment the following line if using designed components 
        //InitializeComponent(); 
    }
    public SearchResult GetADAccount(string snSamaccountname)
    {

        SearchResult srSearchResult;
        DirectorySearcher dsDirectorySearcher = new DirectorySearcher();
        string roRootDSE = dsDirectorySearcher.SearchRoot.Path;
        DirectoryEntry deDirectoryEntry = new DirectoryEntry(roRootDSE);
        dsDirectorySearcher.SearchScope = SearchScope.Subtree;
        dsDirectorySearcher.Filter = "(sAMAccountName=" + snSamaccountname + ")";
        dsDirectorySearcher.PropertiesToLoad.Add("co");
        dsDirectorySearcher.PropertiesToLoad.Add("company");
        dsDirectorySearcher.PropertiesToLoad.Add("displayName");
        dsDirectorySearcher.PropertiesToLoad.Add("department");
        dsDirectorySearcher.PropertiesToLoad.Add("distinguishedName");
        dsDirectorySearcher.PropertiesToLoad.Add("description");
        dsDirectorySearcher.PropertiesToLoad.Add("facsimileTelephoneNumber");
        dsDirectorySearcher.PropertiesToLoad.Add("givenName");
        dsDirectorySearcher.PropertiesToLoad.Add("homePhone");
        dsDirectorySearcher.PropertiesToLoad.Add("info");
        dsDirectorySearcher.PropertiesToLoad.Add("initials");
        dsDirectorySearcher.PropertiesToLoad.Add("l");
        dsDirectorySearcher.PropertiesToLoad.Add("mail");
        dsDirectorySearcher.PropertiesToLoad.Add("middleName");
        dsDirectorySearcher.PropertiesToLoad.Add("mobile");
        dsDirectorySearcher.PropertiesToLoad.Add("name");
        dsDirectorySearcher.PropertiesToLoad.Add("physicalDeliveryOfficeName");
        dsDirectorySearcher.PropertiesToLoad.Add("postalCode");
        dsDirectorySearcher.PropertiesToLoad.Add("postOfficeBox");
        dsDirectorySearcher.PropertiesToLoad.Add("pager");
        dsDirectorySearcher.PropertiesToLoad.Add("sn");
        dsDirectorySearcher.PropertiesToLoad.Add("st");
        dsDirectorySearcher.PropertiesToLoad.Add("streetAddress");
        dsDirectorySearcher.PropertiesToLoad.Add("telephoneNumber");
        dsDirectorySearcher.PropertiesToLoad.Add("title");
        dsDirectorySearcher.PropertiesToLoad.Add("wWWHomePage");
        srSearchResult = dsDirectorySearcher.FindOne();
        return srSearchResult;
    }
    public XmlDocument CreatehCard(SearchResult srSearchResult)
    {
        string fnFirstname = "";
        string snSurname = "";
        string mnMiddleName = "";
        string wwURLwebsite = "";
        string emEmailAddress = "";
        string hpHomephone = "";
        string mpMobilephone = "";
        string bpBusinessphone = "";
        string saStreetAddress = "";
        string pbPostofficeBox = "";
        string stState = "";
        string sbSuburb = "";
        string cnCountry = "";
        string pcPostcode = "";
        string fnFaxNumber = "";
        string tnTitle = "";
        string cncompanyName = ""; 
        string dnDepartmentName = "";
        string noNote = "";
        string eaExtendedAddress = "";


        if (srSearchResult.Properties.Contains("givenName")) { fnFirstname = (string)srSearchResult.Properties["givenName"][0]; }
        if (srSearchResult.Properties.Contains("sn")) { snSurname = (string)srSearchResult.Properties["sn"][0]; }
        if (srSearchResult.Properties.Contains("middleName")) { mnMiddleName = (string)srSearchResult.Properties["middleName"][0]; }
        if (srSearchResult.Properties.Contains("wWWHomePage")) { wwURLwebsite = (string)srSearchResult.Properties["wWWHomePage"][0]; }
        if (srSearchResult.Properties.Contains("homePhone")) { hpHomephone = (string)srSearchResult.Properties["homePhone"][0]; }
        if (srSearchResult.Properties.Contains("mobile")) { mpMobilephone = (string)srSearchResult.Properties["mobile"][0]; }
        if (srSearchResult.Properties.Contains("telephoneNumber")) { bpBusinessphone = (string)srSearchResult.Properties["telephoneNumber"][0]; }
        if (srSearchResult.Properties.Contains("mail")) { emEmailAddress = (string)srSearchResult.Properties["mail"][0]; }
        if (srSearchResult.Properties.Contains("l")) { sbSuburb  = (string)srSearchResult.Properties["l"][0]; }
        if (srSearchResult.Properties.Contains("postalCode")) { pcPostcode = (string)srSearchResult.Properties["postalCode"][0]; }
        if (srSearchResult.Properties.Contains("streetAddress")) { saStreetAddress = (string)srSearchResult.Properties["streetAddress"][0]; }
        if (srSearchResult.Properties.Contains("co")) { cnCountry  = (string)srSearchResult.Properties["co"][0]; }
        if (srSearchResult.Properties.Contains("st")) { stState = (string)srSearchResult.Properties["st"][0]; }
        if (srSearchResult.Properties.Contains("facsimileTelephoneNumber")) { fnFaxNumber = (string)srSearchResult.Properties["facsimileTelephoneNumber"][0]; }
        if (srSearchResult.Properties.Contains("postOfficeBox")) { pbPostofficeBox  = (string)srSearchResult.Properties["postOfficeBox"][0]; }
        if (srSearchResult.Properties.Contains("title")) { tnTitle = (string)srSearchResult.Properties["title"][0]; }
        if (srSearchResult.Properties.Contains("company")) { cncompanyName = (string)srSearchResult.Properties["company"][0]; }
        if (srSearchResult.Properties.Contains("department")) { dnDepartmentName = (string)srSearchResult.Properties["department"][0]; }
        if (srSearchResult.Properties.Contains("description")) { noNote = (string)srSearchResult.Properties["description"][0]; }
        if (srSearchResult.Properties.Contains("physicalDeliveryOfficeName")) { eaExtendedAddress = (string)srSearchResult.Properties["physicalDeliveryOfficeName"][0]; }

        XmlDocument hcHcardDocuemnt = new XmlDocument();
        StringWriter xsXmlString = new StringWriter();
        XmlWriter xrXmlWritter = new XmlTextWriter(xsXmlString);
        // write Div
        xrXmlWritter.WriteStartDocument();
        xrXmlWritter.WriteStartElement("div");
        xrXmlWritter.WriteAttributeString("class", "vcard");
        
        //Name Element
        xrXmlWritter.WriteStartElement("a");
        xrXmlWritter.WriteAttributeString("class","url fn n");
        xrXmlWritter.WriteAttributeString("href", wwURLwebsite);
        xrXmlWritter.WriteStartElement("span");
        xrXmlWritter.WriteAttributeString("class", "given-name");
        xrXmlWritter.WriteValue(fnFirstname);
        xrXmlWritter.WriteEndElement();
        if (mnMiddleName != "")
        {
            xrXmlWritter.WriteStartElement("span");
            xrXmlWritter.WriteAttributeString("class", "additional-name");
            xrXmlWritter.WriteValue(mnMiddleName);
            xrXmlWritter.WriteEndElement();
        }
        xrXmlWritter.WriteStartElement("span");
        xrXmlWritter.WriteAttributeString("class", "family-name");
        xrXmlWritter.WriteValue(snSurname);
        xrXmlWritter.WriteEndElement();
        // End Name element
        // Email Element Start
        xrXmlWritter.WriteEndElement();
        xrXmlWritter.WriteStartElement("a");
        xrXmlWritter.WriteAttributeString("class", "email fn");
        xrXmlWritter.WriteAttributeString("href","mailto:" + emEmailAddress);
        xrXmlWritter.WriteValue(fnFirstname + " " + snSurname);
        xrXmlWritter.WriteEndElement();
        if (bpBusinessphone != "")
        {
            //Primary Telephone Element
            xrXmlWritter.WriteStartElement("div");
            xrXmlWritter.WriteAttributeString("class", "tel");
            xrXmlWritter.WriteStartElement("span");
            xrXmlWritter.WriteAttributeString("class", "value");
            xrXmlWritter.WriteValue(bpBusinessphone);
            xrXmlWritter.WriteEndElement();
            xrXmlWritter.WriteStartElement("abbr");
            xrXmlWritter.WriteAttributeString("class", "type");
            xrXmlWritter.WriteAttributeString("title", "work");
            xrXmlWritter.WriteValue("business");
            xrXmlWritter.WriteEndElement();
            xrXmlWritter.WriteEndElement();
        }
        //Home Telephone Elemenet
        if (hpHomephone != "")
        {
            xrXmlWritter.WriteStartElement("div");
            xrXmlWritter.WriteAttributeString("class", "tel");
            xrXmlWritter.WriteStartElement("span");
            xrXmlWritter.WriteAttributeString("class", "value");
            xrXmlWritter.WriteValue(hpHomephone);
            xrXmlWritter.WriteEndElement();
            xrXmlWritter.WriteStartElement("abbr");
            xrXmlWritter.WriteAttributeString("class", "type");
            xrXmlWritter.WriteAttributeString("title", "home");
            xrXmlWritter.WriteValue("home");
            xrXmlWritter.WriteEndElement();
            xrXmlWritter.WriteEndElement();
        }
        //Mobile Phone Elemenet
        if (mpMobilephone != "")
        {
            xrXmlWritter.WriteStartElement("div");
            xrXmlWritter.WriteAttributeString("class", "tel");
            xrXmlWritter.WriteStartElement("span");
            xrXmlWritter.WriteAttributeString("class", "value");
            xrXmlWritter.WriteValue(mpMobilephone);
            xrXmlWritter.WriteEndElement();
            xrXmlWritter.WriteStartElement("abbr");
            xrXmlWritter.WriteAttributeString("class", "type");
            xrXmlWritter.WriteAttributeString("title", "cell");
            xrXmlWritter.WriteValue("cell");
            xrXmlWritter.WriteEndElement();
            xrXmlWritter.WriteEndElement();
        }
        //Fax Elemenet
        if (fnFaxNumber != "")
        {
            xrXmlWritter.WriteStartElement("div");
            xrXmlWritter.WriteAttributeString("class", "tel");
            xrXmlWritter.WriteStartElement("span");
            xrXmlWritter.WriteAttributeString("class", "value");
            xrXmlWritter.WriteValue(fnFaxNumber);
            xrXmlWritter.WriteEndElement();
            xrXmlWritter.WriteStartElement("abbr");
            xrXmlWritter.WriteAttributeString("class", "type");
            xrXmlWritter.WriteAttributeString("title", "fax");
            xrXmlWritter.WriteValue("fax");
            xrXmlWritter.WriteEndElement();
            xrXmlWritter.WriteEndElement();
        }
        // Address Element
        xrXmlWritter.WriteStartElement("div");
        xrXmlWritter.WriteAttributeString("class", "adr");
        xrXmlWritter.WriteStartElement("span");
        xrXmlWritter.WriteAttributeString("class", "type");
        xrXmlWritter.WriteValue("work");
        xrXmlWritter.WriteEndElement();
        if (pbPostofficeBox != "")
        {
            xrXmlWritter.WriteStartElement("span");
            xrXmlWritter.WriteAttributeString("class", "post-office-box");
            xrXmlWritter.WriteString(pbPostofficeBox);
            xrXmlWritter.WriteEndElement();
        }
        if (eaExtendedAddress != "")
        {
            xrXmlWritter.WriteStartElement("span");
            xrXmlWritter.WriteAttributeString("class", "extended-address");
            xrXmlWritter.WriteString(eaExtendedAddress);
            xrXmlWritter.WriteEndElement();
        }
        if (saStreetAddress != "")
        {
            xrXmlWritter.WriteStartElement("span");
            xrXmlWritter.WriteAttributeString("class", "street-address");
            xrXmlWritter.WriteString(saStreetAddress);
            xrXmlWritter.WriteEndElement();
        }
        if (sbSuburb != "")
        {
            xrXmlWritter.WriteStartElement("span");
            xrXmlWritter.WriteAttributeString("class", "locality");
            xrXmlWritter.WriteString(sbSuburb);
            xrXmlWritter.WriteEndElement();
        }
        if (stState != "")
        {
            xrXmlWritter.WriteStartElement("span");
            xrXmlWritter.WriteAttributeString("class", "region");
            xrXmlWritter.WriteString(stState);
            xrXmlWritter.WriteEndElement();
        }
        if (cnCountry != "")
        {
            xrXmlWritter.WriteStartElement("span");
            xrXmlWritter.WriteAttributeString("class", "country-name");
            xrXmlWritter.WriteString(cnCountry);
            xrXmlWritter.WriteEndElement();
        }
        if (pcPostcode != "")
        {
            xrXmlWritter.WriteStartElement("span");
            xrXmlWritter.WriteAttributeString("class", "postal-code");
            xrXmlWritter.WriteString(pcPostcode);
            xrXmlWritter.WriteEndElement();
        }
        xrXmlWritter.WriteEndElement();
        //End Address Element
        //Title Element
        if (tnTitle != "")
        {
            xrXmlWritter.WriteStartElement("div");
            xrXmlWritter.WriteAttributeString("class", "title");
            xrXmlWritter.WriteString(tnTitle);
            xrXmlWritter.WriteEndElement();        
        }
        // OrgElement
        if (cncompanyName != "" | dnDepartmentName != "") {
            xrXmlWritter.WriteStartElement("div");
            xrXmlWritter.WriteAttributeString("class", "org");
            if (cncompanyName != "")
            {
                xrXmlWritter.WriteStartElement("span");
                xrXmlWritter.WriteAttributeString("class", "organization-name");
                xrXmlWritter.WriteString(cncompanyName);
                xrXmlWritter.WriteEndElement();
            }
            if (dnDepartmentName != "")
            {
                xrXmlWritter.WriteStartElement("span");
                xrXmlWritter.WriteAttributeString("class", "organization-unit");
                xrXmlWritter.WriteString(dnDepartmentName);
                xrXmlWritter.WriteEndElement();
            }
            xrXmlWritter.WriteEndElement();   
        }
        if (noNote != "") {
            xrXmlWritter.WriteStartElement("div");
            xrXmlWritter.WriteAttributeString("class", "note");
            xrXmlWritter.WriteString(noNote);
            xrXmlWritter.WriteEndElement();          
        }
        // Div vcard End element
        xrXmlWritter.WriteEndElement();
        xrXmlWritter.WriteEndDocument();
        // Load and Return a XML Document
        hcHcardDocuemnt.LoadXml(xsXmlString.ToString());
        return hcHcardDocuemnt;

    }

    [WebMethod]
    public XmlDocument GethCard(string snSamaccountname)
    {
        XmlDocument xrXmldocresult;
        try
        {
            SearchResult srSearchResult = this.GetADAccount(snSamaccountname);
            if (srSearchResult != null) { xrXmldocresult = this.CreatehCard(srSearchResult); }
            else
            {
                xrXmldocresult = new XmlDocument();
                XmlElement xeFirstDivElement = xrXmldocresult.CreateElement("div");
                xeFirstDivElement.InnerText = "Active Directory Account not Found";
                xrXmldocresult.AppendChild(xeFirstDivElement);
            }
        }
       catch(Exception e) {
            xrXmldocresult = new XmlDocument();
            XmlElement xeFirstDivElement = xrXmldocresult.CreateElement("ErrorOccured");
            xeFirstDivElement.InnerText = e.Message.ToString();
            xrXmldocresult.AppendChild(xeFirstDivElement);
            XmlElement xeFirstDivElement1 = xrXmldocresult.CreateElement("ErrorLineNumber");
            xeFirstDivElement1.InnerText = e.Source.ToString();
            xrXmldocresult.AppendChild(xeFirstDivElement1);
      }
        return xrXmldocresult;
        ;

    }
    
}

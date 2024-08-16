using System;
using System.Configuration;
using System.Data;
using System.Web;
using System.Net;
using System.Web.Security;
using System.Net.Security;
using System.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Security.Cryptography.X509Certificates;
using EWSUtil.EWS;


public partial class _Map : System.Web.UI.Page 
{
    protected void Page_Load(object sender, EventArgs e)
    {
        ServicePointManager.ServerCertificateValidationCallback =
        delegate(Object obj, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
            {
                //   Ignore Self Signed Certs
                return true;
            };
	ExchangeServiceBinding esb = new ExchangeServiceBinding();
        esb.RequestServerVersionValue = new RequestServerVersion();
        esb.RequestServerVersionValue.Version = ExchangeVersionType.Exchange2007_SP1;
        esb.Url = @"https://127.0.0.1/EWS/Exchange.asmx";
        esb.Credentials = CredentialCache.DefaultCredentials;
        System.Security.Principal.WindowsImpersonationContext impersonationContext;
        impersonationContext =
            ((System.Security.Principal.WindowsIdentity)User.Identity).Impersonate();


        String ewsID = ConvertOWAid(esb,Request.QueryString["id"],Request.QueryString["ea"]);
        ContactItemType ciCurrentContact = getContact(esb,ewsID);
        AddressBox.Text = ciCurrentContact.DisplayName + Environment.NewLine + ciCurrentContact.PhysicalAddresses[0].Street + Environment.NewLine + ciCurrentContact.PhysicalAddresses[0].City + Environment.NewLine 
        + ciCurrentContact.PhysicalAddresses[0].State + Environment.NewLine + ciCurrentContact.PhysicalAddresses[0].PostalCode + Environment.NewLine  + ciCurrentContact.PhysicalAddresses[0].CountryOrRegion;
        
        String asAddressString = ciCurrentContact.PhysicalAddresses[0].Street + "," + ciCurrentContact.PhysicalAddresses[0].City + "," + ciCurrentContact.PhysicalAddresses[0].State;

        string gkGoogleKey = "abc1234";

        Uri uristring = new Uri(String.Format("{0}{1}&output={2}&key={3}", "http://maps.google.com/maps/geo?q=", HttpUtility.UrlEncode(asAddressString), "csv", gkGoogleKey));
        WebClient client = new WebClient();
        string[] geocodeInfo = client.DownloadString(uristring).Split(',');
        string js = "<script type=\"text/javascript\">"
        + "function initialize() {" 
        +  "var lnglat = new GLatLng(" + geocodeInfo[2] + "," + geocodeInfo[3] + ");"
        +  "var map = new GMap2(document.getElementById(\"map_canvas\"));"
        +  "map.setCenter(lnglat, 16);"
        +  "myPOV = { yaw:90.00 };"
    	+  "panoramaOptions = { latlng:lnglat, pov:myPOV};"
        +  "myPano = new GStreetviewPanorama(document.getElementById(\"pano\"), panoramaOptions);}"
        +  "</script>";
        Response.Write(js);
        impersonationContext.Undo();

    }
   
        

        
        static String ConvertOWAid(ExchangeServiceBinding esb,String oiOWAID,String emEmailAddress) {
        String riReturnID = "";
        ConvertIdType ciConvertIDRequest = new ConvertIdType();
        ciConvertIDRequest.SourceIds = new AlternateIdType[1];
        ciConvertIDRequest.SourceIds[0] = new AlternateIdType();
        ciConvertIDRequest.SourceIds[0].Format = IdFormatType.OwaId;
        (ciConvertIDRequest.SourceIds[0] as AlternateIdType).Id = oiOWAID;
        (ciConvertIDRequest.SourceIds[0] as AlternateIdType).Mailbox = emEmailAddress;
        ciConvertIDRequest.DestinationFormat = IdFormatType.EwsId;
        ConvertIdResponseType response = esb.ConvertId(ciConvertIDRequest);
        ArrayOfResponseMessagesType aormt = response.ResponseMessages;
        ResponseMessageType[] rmta = aormt.Items;
        foreach (ConvertIdResponseMessageType resp in rmta)
        {
            if (resp.ResponseClass == ResponseClassType.Success)
            {
                ConvertIdResponseMessageType cirmt = (resp as ConvertIdResponseMessageType);
                AlternateIdType aiID = (cirmt.AlternateId as AlternateIdType);
                riReturnID = aiID.Id.ToString();

            }
            else if (resp.ResponseClass == ResponseClassType.Error)
            {
                Console.WriteLine("Error: " + resp.MessageText);
            }
        }
        return riReturnID;
    
    }
        static ContactItemType getContact(ExchangeServiceBinding esb, String iiEWSid)
        {
            ContactItemType  rcContact = null;
            GetItemType giRequest = new GetItemType();
            ItemIdType iiItemId = new ItemIdType();
            iiItemId.Id = iiEWSid;
            giRequest.ItemIds = new ItemIdType[1];
            giRequest.ItemIds[0] = iiItemId;
            giRequest.ItemShape = new ItemResponseShapeType();
            giRequest.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;
            giRequest.ItemShape.BodyTypeSpecified = true;
            GetItemResponseType giResponse = esb.GetItem(giRequest);
            if (giResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Error)
            {
                Console.WriteLine("Error Occured");
                Console.WriteLine(giResponse.ResponseMessages.Items[0].MessageText);
            }
            else
            {
                ItemInfoResponseMessageType rmResponseMessage = giResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
                rcContact = (ContactItemType)rmResponseMessage.Items.Items[0];
            }
            return rcContact;
        
        }
}

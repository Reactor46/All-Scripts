using System;
using System.Collections.Generic;
using System.Text;
using System.Net;
using System.Xml;
using System.Net.Security;
using System.IO;
using System.Reflection;
using System.DirectoryServices;
using System.Security.Cryptography.X509Certificates;
using EWSOofUtil.EWS;

namespace EWSOofUtil
{
    public class OofUtil
    {
        private string intOofStatus = "";
        public string OofStatus
        {
            get
            {
                return intOofStatus;
            }
        }
        private string intInternalMessage = "";
        public string InternalMessage
        {
            get
            {
                return this.intInternalMessage;
            }
        }
        private string intExternalMessage = "";
        public string ExternalMessage
        {
            get
            {
                return this.intExternalMessage;
            }
        }
        private Duration  intDuration = null;
        public DateTime DurationStartTime
        {
            get
            {
                return this.intDuration.StartTime;
            }
        }
        public DateTime DurationEndTime
        {
            get
            {
                return this.intDuration.EndTime;
            }
        }
        private string intInternalMessageLanguage = "";
        public string InternalMessageLanguage
        {
            get
            {
                return this.intInternalMessageLanguage;
            }
        }

        private string intExternalMessageLanguage = "";
        public string ExternalMessageLanguage
        {
            get
            {
                return this.intExternalMessageLanguage;
            }

        }
        private string intExternalAudienceSetting = "";
        public string ExternalAudienceSetting
        {
            get
            {
                return this.intExternalAudienceSetting;
            }
        }
    
        public String SetOof(string EmailAddress, string OofStatus)
        {
            String rsResult = SetOof(EmailAddress, OofStatus, String.Empty, String.Empty, DateTime.MinValue, DateTime.MinValue, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty);
           return rsResult;
        }
        public String SetOof(string EmailAddress, string OofStatus,string UserName,string Password,string Domain, string OofURL)
        {
            String rsResult = SetOof(EmailAddress, OofStatus, String.Empty, String.Empty, DateTime.MinValue, DateTime.MinValue, String.Empty, String.Empty, String.Empty, UserName, Password, Domain, OofURL);
            return rsResult;
        }
        public String SetOof(string EmailAddress, string OofStatus,string InternalMessage,string ExternalMessage, string UserName, string Password, string Domain, string OofURL)
        {
            String rsResult = SetOof(EmailAddress, OofStatus,InternalMessage,ExternalMessage, DateTime.MinValue, DateTime.MinValue, String.Empty, String.Empty, String.Empty, UserName, Password, Domain, OofURL);
            return rsResult;
        }
        public String SetOof(string EmailAddress, string OofStatus, DateTime DurationStartTime,DateTime DurationEndTime)
        {
            String rsResult = SetOof(EmailAddress, OofStatus, String.Empty, String.Empty, DurationStartTime, DurationEndTime, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty);
            return rsResult;
        }
        public String SetOof(string EmailAddress, string OofStatus, DateTime DurationStartTime, DateTime DurationEndTime,string UserName,string Password,string Domain, string OofURL)
        {
            String rsResult = SetOof(EmailAddress, OofStatus, String.Empty, String.Empty, DurationStartTime, DurationEndTime, String.Empty, String.Empty, String.Empty, UserName, Password, Domain, OofURL);
            return rsResult;
        }
        public String SetOof(string EmailAddress, string OofStatus, string InternalMessage, string ExternalMessage)
        {
            String rsResult = SetOof(EmailAddress, OofStatus, InternalMessage, ExternalMessage, DateTime.MinValue, DateTime.MinValue, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty);
            return rsResult;
        }
        public String SetOof(string EmailAddress, string OofStatus, string InternalMessage, string ExternalMessage, string ExternalAudienceSetting)
        {
            String rsResult = SetOof(EmailAddress, OofStatus, InternalMessage, ExternalMessage, DateTime.MinValue, DateTime.MinValue, String.Empty, String.Empty, ExternalAudienceSetting, String.Empty, String.Empty, String.Empty, String.Empty);
            return rsResult;
        }
        public String SetOof(string EmailAddress, string OofStatus, string InternalMessage, string ExternalMessage, string InternalMessageLanguage, string ExternalMessageLanguage, string ExternalAudienceSetting)
        {
            String rsResult = SetOof(EmailAddress, OofStatus, InternalMessage, ExternalMessage, DateTime.MinValue, DateTime.MinValue, InternalMessageLanguage, ExternalMessageLanguage, ExternalAudienceSetting, String.Empty, String.Empty, String.Empty, String.Empty);
            return rsResult;
        }
        public String SetOof(string EmailAddress, string OofStatus, string InternalMessage, string ExternalMessage, DateTime DurationStartTime,DateTime DurationEndTime, string InternalMessageLanguage, string ExternalMessageLanguage, string ExternalAudienceSetting, string UserName, string Password,string Domain,  string OofURL)
        {
            String rsResult = "";
            try
            {
                ExchangeServiceBinding ebExchangeServiceBinding = createesb(EmailAddress,  UserName, Password,Domain, OofURL);
                UserOofSettings noNewOofSetting = ewsGetOOF(ebExchangeServiceBinding, EmailAddress);
                if (OofStatus != "")
                {
                    switch (OofStatus.ToLower())
                    {
                        case "enabled": noNewOofSetting.OofState = OofState.Enabled;
                            break;
                        case "disabled": noNewOofSetting.OofState = OofState.Disabled;
                            break;
                        case "scheduled": noNewOofSetting.OofState = OofState.Scheduled;
                            break;
                        default: throw new OofSettingException("OofSetting OofStatus Invalid" + ExternalAudienceSetting);
                    }
                }
                if (InternalMessage != "")
                {
                    ReplyBody rpReplyBody = new ReplyBody();
                    rpReplyBody.Message = InternalMessage;
                    if (InternalMessageLanguage != "")
                    {
                        rpReplyBody.lang = InternalMessageLanguage;
                    }
                    noNewOofSetting.InternalReply = rpReplyBody;
                }
                if (ExternalMessage != "")
                {
                    ReplyBody erExternalReplyBody = new ReplyBody();
                    erExternalReplyBody.Message = ExternalMessage;
                    if (ExternalMessageLanguage != "")
                    {
                        erExternalReplyBody.lang = ExternalMessageLanguage;
                    }
                    noNewOofSetting.ExternalReply = erExternalReplyBody;
                }
                // MinValue is used because DateTime Values can be null
                if (DurationStartTime != DateTime.MinValue | DurationEndTime != DateTime.MinValue)
                {
                    Duration drDuration = new Duration();
                    drDuration.StartTime = DurationStartTime;
                    drDuration.EndTime = DurationEndTime;
                    noNewOofSetting.Duration = drDuration ;
                }
                if (ExternalAudienceSetting != "")
                {
                    switch (ExternalAudienceSetting.ToLower())
                    {
                        case "all": noNewOofSetting.ExternalAudience = ExternalAudience.All;
                            break;
                        case "none": noNewOofSetting.ExternalAudience = ExternalAudience.None;
                            break;
                        case "known": noNewOofSetting.ExternalAudience = ExternalAudience.Known;
                            break;
                        default: throw new OofSettingException("OffSetting External Audience Invalid" + ExternalAudienceSetting);
                    }
                }
                rsResult = ewsSetOOF(ebExchangeServiceBinding, noNewOofSetting, EmailAddress);
                ebExchangeServiceBinding = null;
                ServicePointManager.ServerCertificateValidationCallback = null;
                return rsResult;
            }
            catch (Exception exException)
            {
                Console.WriteLine(exException.ToString());
                return exException.ToString();
            }
     
        }
        public String GetOof(string EmailAddress)
        {
            String rsResult = GetOof(EmailAddress, String.Empty, String.Empty, String.Empty, String.Empty);
            return rsResult;
        }
        public String GetOof(string EmailAddress, string UserName, string Password,string Domain)
        {
            String rsResult = GetOof(EmailAddress, UserName, Password,Domain, String.Empty);
            return rsResult;
        }
        public String GetOof(string EmailAddress, string UserName, string Password,string Domain, string OofURL) {
            String rsResult = null;
            try
            {
                ExchangeServiceBinding ebExchangeServiceBinding = createesb(EmailAddress, UserName, Password,Domain,OofURL);
                UserOofSettings uoUserOofSettings = ewsGetOOF(ebExchangeServiceBinding, EmailAddress);
                this.intOofStatus = uoUserOofSettings.OofState.ToString();
                if (uoUserOofSettings.InternalReply.Message != null) { this.intInternalMessage = uoUserOofSettings.InternalReply.Message.ToString(); }
                if (uoUserOofSettings.InternalReply.lang != null) { this.intInternalMessageLanguage = uoUserOofSettings.InternalReply.lang.ToString(); }
                if (uoUserOofSettings.ExternalReply.Message != null) { this.intExternalMessage = uoUserOofSettings.ExternalReply.Message.ToString(); }
                if (uoUserOofSettings.ExternalReply.lang != null) { this.intExternalMessageLanguage = uoUserOofSettings.ExternalReply.lang.ToString(); }
                this.intExternalAudienceSetting = uoUserOofSettings.ExternalAudience.ToString();
                if (uoUserOofSettings.Duration != null) { this.intDuration = uoUserOofSettings.Duration; }                
                rsResult = "OOF Settings retrieved";
                ServicePointManager.ServerCertificateValidationCallback = null;
                ebExchangeServiceBinding = null;
                return rsResult;
            }
            catch (Exception exException)
            {
                Console.WriteLine(exException.ToString());
                return exException.ToString();
            }
        }

        private ExchangeServiceBinding createesb(String EmailAddress, string UserName, string Password, string Domain,string OofURL) {
             ServicePointManager.ServerCertificateValidationCallback =
             delegate(Object obj, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
             {
                 //   Ignore Self Signed Certs
                 return true;
             };
            ExchangeServiceBinding ebExchangeServiceBinding = new ExchangeServiceBinding();
            ebExchangeServiceBinding.RequestServerVersionValue = new RequestServerVersion();
            ebExchangeServiceBinding.RequestServerVersionValue.Version = ExchangeVersionType.Exchange2007_SP1;
            if (UserName == "")
            {
                ebExchangeServiceBinding.UseDefaultCredentials = true;
            }
            else
            {
                NetworkCredential ncNetCredential = new NetworkCredential(UserName, Password, Domain);
                ebExchangeServiceBinding.Credentials = ncNetCredential;
            }
            if (OofURL == "")
            {
                String caCasURL = DiscoverCAS();
                OofURL = DiscoverOofURL(caCasURL, EmailAddress, UserName, Password,Domain);
            }
            ebExchangeServiceBinding.Url = OofURL;
            return ebExchangeServiceBinding;
        
        }
        private String DiscoverCAS()
        {
            String ScpUrlGuidString = "77378F46-2C66-4aa9-A6A6-3E7A48B19596";
            String ScpPtrGuidString = "67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68";
            DirectoryEntry rdRootDSE = new DirectoryEntry("LDAP://RootDSE");
            DirectoryEntry cfConfigPartition = new DirectoryEntry("LDAP://" + rdRootDSE.Properties["configurationnamingcontext"].Value);
            DirectorySearcher cfConfigPartitionSearch = new DirectorySearcher(cfConfigPartition);
            cfConfigPartitionSearch.Filter = "(&(objectClass=serviceConnectionPoint)(|(keywords=" + ScpPtrGuidString + ")(keywords=" + ScpUrlGuidString + ")))";
            cfConfigPartitionSearch.SearchScope = SearchScope.Subtree;
            string CASURL = null;
            SearchResult srSearchResult = cfConfigPartitionSearch.FindOne();
            if (srSearchResult != null)
            {
                DirectoryEntry scpServiceConnectionPoint = srSearchResult.GetDirectoryEntry();
                CASURL = scpServiceConnectionPoint.Properties["serviceBindingInformation"].Value.ToString();
            }
            else
            {
                throw new ADSearchException("No SCP found");
            }
            return CASURL;
        }
        private String DiscoverOofURL(string caCASURL, string emEmailAddress,string UserName,string Password,string Domain)
        {
   
            String OofURL = null;
            String auDisXML = "<Autodiscover xmlns=\"http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006\"><Request>" +
                  "<EMailAddress>" + emEmailAddress + "</EMailAddress>" +
                  "<AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>" +
                  "</Request>" +
                  "</Autodiscover>";
            System.Net.HttpWebRequest adAutoDiscoRequest = (System.Net.HttpWebRequest)System.Net.HttpWebRequest.Create(caCASURL);
            adAutoDiscoRequest.ContentType = "text/xml";
            adAutoDiscoRequest.Headers.Add("Translate", "F");
            adAutoDiscoRequest.Method = "Post";
            if (UserName == "")
            {
                adAutoDiscoRequest.UseDefaultCredentials = true;
            }
            else
            {
                NetworkCredential ncNetCredential = new NetworkCredential(UserName, Password, Domain);
                adAutoDiscoRequest.Credentials = ncNetCredential;
            }
           
            byte[] bytes = Encoding.UTF8.GetBytes(auDisXML);
            adAutoDiscoRequest.ContentLength = bytes.Length;
            Stream rsRequestStream = adAutoDiscoRequest.GetRequestStream();
            rsRequestStream.Write(bytes, 0, bytes.Length);
            rsRequestStream.Close();
            WebResponse adResponse = adAutoDiscoRequest.GetResponse();
            Stream rsResponseStream = adResponse.GetResponseStream();
            XmlDocument reResponseDoc = new XmlDocument();
            reResponseDoc.Load(rsResponseStream);
            XmlNodeList OofNodes = reResponseDoc.GetElementsByTagName("OOFUrl");
            if (OofNodes.Count != 0)
            {
                OofURL = OofNodes[0].InnerText;
            }
            else { 
                throw new AutoDiscoveryException("Error during AutoDiscovery");            
            }
            return OofURL;

        }
        private UserOofSettings ewsGetOOF(ExchangeServiceBinding ebExchangeServiceBinding, String emEmailAddress)
        {
            GetUserOofSettingsRequest goGetUserOofSettings = new GetUserOofSettingsRequest();
            UserOofSettings ouOffSetting = null;
            EmailAddress mbMailbox = new EmailAddress();
            mbMailbox.Address = emEmailAddress;
            goGetUserOofSettings.Mailbox = mbMailbox;
            GetUserOofSettingsResponse goGetOoFResponse = ebExchangeServiceBinding.GetUserOofSettings(goGetUserOofSettings);
            if (goGetOoFResponse.ResponseMessage.ResponseClass == ResponseClassType.Success)
            {
                ouOffSetting = goGetOoFResponse.OofSettings;
            }
            else
            {
                throw  new EWSException(goGetOoFResponse.ResponseMessage.MessageText.ToString());
            }

            return ouOffSetting;
        }
        private String ewsSetOOF(ExchangeServiceBinding ebExchangeServiceBinding, UserOofSettings uoNewOoFSettings, String emEmailAddress)
        {
            SetUserOofSettingsRequest soSetUserOofSettings = new SetUserOofSettingsRequest();
            soSetUserOofSettings.UserOofSettings = uoNewOoFSettings;
            EmailAddress mbMailbox = new EmailAddress();
            mbMailbox.Address = emEmailAddress;
            soSetUserOofSettings.Mailbox = mbMailbox;
            SetUserOofSettingsResponse soSetOoFResponse = ebExchangeServiceBinding.SetUserOofSettings(soSetUserOofSettings);
            String rsResponse = "";

            if (soSetOoFResponse.ResponseMessage.ResponseClass == ResponseClassType.Success)
            {
                this.intOofStatus = uoNewOoFSettings.OofState.ToString(); 
                if (uoNewOoFSettings.InternalReply.Message != null) { this.intInternalMessage = uoNewOoFSettings.InternalReply.Message.ToString(); }
                if (uoNewOoFSettings.InternalReply.lang != null) { this.intInternalMessageLanguage = uoNewOoFSettings.InternalReply.lang.ToString(); }
                if (uoNewOoFSettings.ExternalReply.Message != null) { this.intExternalMessage = uoNewOoFSettings.ExternalReply.Message.ToString(); }
                if (uoNewOoFSettings.ExternalReply.lang != null) { this.intExternalMessageLanguage = uoNewOoFSettings.ExternalReply.lang.ToString(); }
                this.intExternalAudienceSetting = uoNewOoFSettings.ExternalAudience.ToString(); 
                if (uoNewOoFSettings.Duration != null) { this.intDuration = uoNewOoFSettings.Duration; }
                rsResponse = "Oof Setting Update Succesfully";

            }
            else
            {
                throw new EWSException(soSetOoFResponse.ResponseMessage.MessageText.ToString());
            }

            return rsResponse;
        }
    }
    class EWSException : Exception
    {
        public EWSException(string ewsError)
        {
            Console.WriteLine(ewsError);
        }
    }
    class ADSearchException : Exception
    {
        public ADSearchException(string AdSearchError)
        {
            Console.WriteLine(AdSearchError);
        }
    }
    class AutoDiscoveryException : Exception
    {
        public AutoDiscoveryException(string AutoDiscoveryError)
        {
            Console.WriteLine(AutoDiscoveryError);
        }
    }
    class OofSettingException : Exception
    {
        public OofSettingException(string OofSettingError)
        {
            Console.WriteLine(OofSettingError);
        }
    }
}

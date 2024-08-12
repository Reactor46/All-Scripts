using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Net;
using System.Xml;
using System.Net.Security;
using System.Security;
using System.Web;
using System.IO;
using System.Reflection;
using System.DirectoryServices;
using System.Web.Services;
using System.Security.Cryptography.X509Certificates;
using EWSUtil.EWS;


namespace EWSUtil
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
        public String SetOof(string EmailAddress, string OofStatus, DateTime DurationStartTime,DateTime DurationEndTime)
        {
            String rsResult = SetOof(EmailAddress, OofStatus, String.Empty, String.Empty, DurationStartTime, DurationEndTime, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty);
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
                ServicePointManager.ServerCertificateValidationCallback =
                delegate(Object obj, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
                {
                    //   Ignore Self Signed Certs
                    return true;
                };
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
                        default: throw new OofSettingException("OffSetting OofStatus Invalid" + ExternalAudienceSetting);
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
                return rsResult;
            }
            catch (Exception exException)
            {
                Console.WriteLine(exException.ToString());
                return exException.ToString();
            }
        }

        private ExchangeServiceBinding createesb(String EmailAddress, string UserName, string Password, string Domain,string OofURL) {
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
    public class CalendarUtil {
        public CalendarUtil(string emEmailAddress)
        {
            initlizeDACL(emEmailAddress, false, "", "", "", "");
        }
        public CalendarUtil(string emEmailAddress, Boolean Impersonate)
        {
            initlizeDACL(emEmailAddress, Impersonate,"","","","");
        }
        public CalendarUtil(string emEmailAddress, Boolean Impersonate, string UserName, string Password, string Domain)
        {
            initlizeDACL(emEmailAddress, Impersonate,UserName,Password,Domain,"");
        }
        public CalendarUtil(string emEmailAddress, Boolean Impersonate, string UserName, string Password, string Domain, string casURL)
        {
            initlizeDACL(emEmailAddress, Impersonate, UserName, Password, Domain, casURL);
        }
        private void initlizeDACL(string EmailAddress,Boolean Impersonate, string UserName, string Password, string Domain, string ewsURL)
        {
            try
            {
                ServicePointManager.ServerCertificateValidationCallback =
                delegate(Object obj, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
                {
                    //   Ignore Self Signed Certs
                    return true;
                };
                intEmailAddress = EmailAddress;
                ExchangeServiceBinding ebExchangeServiceBinding = createesb(EmailAddress,Impersonate, UserName, Password, Domain, ewsURL);
                CalendarPermissionSetType cdCalendarDACL = GetCalendarDACL(EmailAddress);
                for (int cpint = 0; cpint < cdCalendarDACL.CalendarPermissions.Length; cpint++)
                {
                    if (cdCalendarDACL.CalendarPermissions[cpint].CalendarPermissionLevel == CalendarPermissionLevelType.Custom)
                    {
                        intInternalCalendarDACL.Add(cdCalendarDACL.CalendarPermissions[cpint]);
                    }
                    else
                    {
                        CalendarPermissionType cfNewCalPermsionsSet = new CalendarPermissionType();
                        {
                            cfNewCalPermsionsSet.UserId = cdCalendarDACL.CalendarPermissions[cpint].UserId;
                            cfNewCalPermsionsSet.CalendarPermissionLevel = cdCalendarDACL.CalendarPermissions[cpint].CalendarPermissionLevel;
                        }
                        intInternalCalendarDACL.Add(cfNewCalPermsionsSet);
                    }

                }
                PermissionSetType fbFreeBusyDACL = GetFreeBusyDACL(EmailAddress);
                for (int fbpint = 0; fbpint < fbFreeBusyDACL.Permissions.Length; fbpint++)
                {
                    if (fbFreeBusyDACL.Permissions[fbpint].PermissionLevel == PermissionLevelType.Custom)
                    {
                        intInternalFreeBusyDACL.Add(fbFreeBusyDACL.Permissions[fbpint]);
                    }
                    else
                    {
                        PermissionType fbNewPermsion = new PermissionType();
                        {
                            fbNewPermsion.UserId = fbFreeBusyDACL.Permissions[fbpint].UserId;
                            fbNewPermsion.PermissionLevel = fbFreeBusyDACL.Permissions[fbpint].PermissionLevel;
                        }
                        intInternalFreeBusyDACL.Add(fbNewPermsion);
                    }

                }
            }
            catch (Exception exException)
            {
                Console.WriteLine(exException.ToString());
            }
        }
        private List<CalendarPermissionType> intInternalCalendarDACL =  new List<CalendarPermissionType>();
        public List<CalendarPermissionType> CalendarDACL {
            get
            {
                return this.intInternalCalendarDACL;
            }
        }
        private List<PermissionType> intInternalFreeBusyDACL = new List<PermissionType>();
        public List<PermissionType> FreeBusyDACL
        {
            get
            {
                return this.intInternalFreeBusyDACL;
            }
        }
        public String Update(){
            try
            {
                String upUpdateResult = "";
                upUpdateResult = SetCalendarDACL();
                upUpdateResult = upUpdateResult + (char)13 + SetFreeBusyDACL();
                intInternalCalendarDACL.Clear();
                CalendarPermissionSetType cdCalendarDACL = GetCalendarDACL(intEmailAddress);
                for (int cpint = 0; cpint < cdCalendarDACL.CalendarPermissions.Length; cpint++)
                {
                    if (cdCalendarDACL.CalendarPermissions[cpint].CalendarPermissionLevel == CalendarPermissionLevelType.Custom)
                    {
                        intInternalCalendarDACL.Add(cdCalendarDACL.CalendarPermissions[cpint]);
                    }
                    else
                    {
                        CalendarPermissionType cfNewCalPermsionsSet = new CalendarPermissionType();
                        {
                            cfNewCalPermsionsSet.UserId = cdCalendarDACL.CalendarPermissions[cpint].UserId;
                            cfNewCalPermsionsSet.CalendarPermissionLevel = cdCalendarDACL.CalendarPermissions[cpint].CalendarPermissionLevel;
                        }
                        intInternalCalendarDACL.Add(cfNewCalPermsionsSet);
                    }

                }
                return upUpdateResult;
            }
            catch (Exception exException)
            {
                Console.WriteLine(exException.ToString());
                return exException.ToString();
            }
        }
        public CalendarPermissionType EditorPermissions(string emEmailaddress)
        {
            CalendarPermissionType cpCalendarPerm = new CalendarPermissionType();
            UserIdType auAceUser = new UserIdType();
            if (emEmailaddress.ToLower() == "default" | emEmailaddress.ToLower() == "anonymous")
            {
                auAceUser.DistinguishedUserSpecified = true;
                if (emEmailaddress.ToLower() == "default")
                {
                    auAceUser.DistinguishedUser = DistinguishedUserType.Default;
                }
                else
                {
                    auAceUser.DistinguishedUser = DistinguishedUserType.Anonymous;
                }
            }
            else
            {
                auAceUser.PrimarySmtpAddress = emEmailaddress;
            }
            cpCalendarPerm.UserId = auAceUser;
            cpCalendarPerm.CalendarPermissionLevel = CalendarPermissionLevelType.Custom;
            cpCalendarPerm.ReadItemsSpecified = true;
            cpCalendarPerm.ReadItems = CalendarPermissionReadAccessType.FullDetails;
            cpCalendarPerm.EditItemsSpecified = true;
            cpCalendarPerm.EditItems = PermissionActionType.All;
            cpCalendarPerm.CanCreateItems = true;
            cpCalendarPerm.CanCreateItemsSpecified = true;
            cpCalendarPerm.DeleteItemsSpecified = true;
            cpCalendarPerm.DeleteItems = PermissionActionType.All;
            cpCalendarPerm.IsFolderVisibleSpecified = true;
            cpCalendarPerm.IsFolderVisible = true;
            return cpCalendarPerm;

        }
        public CalendarPermissionType PublishingEditorPermissions(string emEmailaddress)
        {
            CalendarPermissionType cpCalendarPerm = this.EditorPermissions(emEmailaddress);
            cpCalendarPerm.CanCreateSubFolders = true;
            cpCalendarPerm.CanCreateSubFoldersSpecified = true;
            return cpCalendarPerm;

        }
        public CalendarPermissionType OwnerPermissions(string emEmailaddress)
        {
            CalendarPermissionType cpCalendarPerm = this.EditorPermissions(emEmailaddress);
            cpCalendarPerm.CanCreateSubFolders = true;
            cpCalendarPerm.CanCreateSubFoldersSpecified = true;
            cpCalendarPerm.IsFolderOwner = true;
            cpCalendarPerm.IsFolderOwnerSpecified = true;

            return cpCalendarPerm;

        }
        public CalendarPermissionType AuthorPermissions(string emEmailaddress)
        {
            CalendarPermissionType cpCalendarPerm = this.EditorPermissions(emEmailaddress);
            cpCalendarPerm.EditItems = PermissionActionType.Owned;
            cpCalendarPerm.DeleteItems = PermissionActionType.Owned;
            return cpCalendarPerm;

        }
        public CalendarPermissionType NonEditingAuthorPermissions(string emEmailaddress)
        {
            CalendarPermissionType cpCalendarPerm = this.EditorPermissions(emEmailaddress);
            cpCalendarPerm.EditItems = PermissionActionType.None;
            cpCalendarPerm.DeleteItems = PermissionActionType.Owned;
            return cpCalendarPerm;

        }
        public CalendarPermissionType PublishingAuthorPermissions(string emEmailaddress)
        {
            CalendarPermissionType cpCalendarPerm = this.EditorPermissions(emEmailaddress);
            cpCalendarPerm.EditItems = PermissionActionType.Owned;
            cpCalendarPerm.DeleteItems = PermissionActionType.Owned;
            cpCalendarPerm.CanCreateSubFolders = true;
            cpCalendarPerm.CanCreateSubFoldersSpecified = true;
            return cpCalendarPerm;

        }
        public CalendarPermissionType Reviewer(string emEmailaddress)
        {
            CalendarPermissionType cpCalendarPerm = new CalendarPermissionType();
            UserIdType auAceUser = new UserIdType();
            if (emEmailaddress.ToLower() == "default" | emEmailaddress.ToLower() == "anonymous")
            {
                auAceUser.DistinguishedUserSpecified = true;
                if (emEmailaddress.ToLower() == "default")
                {
                    auAceUser.DistinguishedUser = DistinguishedUserType.Default;
                }
                else
                {
                    auAceUser.DistinguishedUser = DistinguishedUserType.Anonymous;
                }
            }
            else
            {
                auAceUser.PrimarySmtpAddress = emEmailaddress;
            }
            cpCalendarPerm.UserId = auAceUser;
            cpCalendarPerm.CalendarPermissionLevel = CalendarPermissionLevelType.Custom;
            cpCalendarPerm.ReadItemsSpecified = true;
            cpCalendarPerm.ReadItems = CalendarPermissionReadAccessType.FullDetails;
            cpCalendarPerm.EditItemsSpecified = true;
            cpCalendarPerm.EditItems = PermissionActionType.None;
            cpCalendarPerm.CanCreateItems = false;
            cpCalendarPerm.CanCreateItemsSpecified = true;
            cpCalendarPerm.DeleteItemsSpecified = true;
            cpCalendarPerm.DeleteItems = PermissionActionType.None;
            cpCalendarPerm.IsFolderVisibleSpecified = true;
            cpCalendarPerm.IsFolderVisible = true;
            return cpCalendarPerm;

            }
        public CalendarPermissionType Contributer(string emEmailaddress)
            {
                CalendarPermissionType cpCalendarPerm = new CalendarPermissionType();
                UserIdType auAceUser = new UserIdType();
                if (emEmailaddress.ToLower() == "default" | emEmailaddress.ToLower() == "anonymous")
                {
                    auAceUser.DistinguishedUserSpecified = true;
                    if (emEmailaddress.ToLower() == "default")
                    {
                        auAceUser.DistinguishedUser = DistinguishedUserType.Default;
                    }
                    else
                    {
                        auAceUser.DistinguishedUser = DistinguishedUserType.Anonymous;
                    }
                }
                else
                {
                    auAceUser.PrimarySmtpAddress = emEmailaddress;
                }
                cpCalendarPerm.UserId = auAceUser;
                cpCalendarPerm.CalendarPermissionLevel = CalendarPermissionLevelType.Custom;
                cpCalendarPerm.ReadItemsSpecified = true;
                cpCalendarPerm.ReadItems = CalendarPermissionReadAccessType.None;
                cpCalendarPerm.EditItemsSpecified = true;
                cpCalendarPerm.EditItems = PermissionActionType.None;
                cpCalendarPerm.CanCreateItems = true;
                cpCalendarPerm.CanCreateItemsSpecified = true;
                cpCalendarPerm.DeleteItemsSpecified = true;
                cpCalendarPerm.DeleteItems = PermissionActionType.None;
                cpCalendarPerm.IsFolderVisibleSpecified = true;
                cpCalendarPerm.IsFolderVisible = true;
                return cpCalendarPerm;

       }
        public CalendarPermissionType FreeBusyTimeOnly(string emEmailaddress)
        {
            CalendarPermissionType cpCalendarPerm = new CalendarPermissionType();
            UserIdType auAceUser = new UserIdType();
            if (emEmailaddress.ToLower() == "default" | emEmailaddress.ToLower() == "anonymous")
            {
                auAceUser.DistinguishedUserSpecified = true;
                if (emEmailaddress.ToLower() == "default")
                {
                    auAceUser.DistinguishedUser = DistinguishedUserType.Default;
                }
                else {
                    auAceUser.DistinguishedUser = DistinguishedUserType.Anonymous;
                }
            }
            else
            {
                auAceUser.PrimarySmtpAddress = emEmailaddress;
            }
            cpCalendarPerm.UserId = auAceUser;
            cpCalendarPerm.CalendarPermissionLevel = CalendarPermissionLevelType.FreeBusyTimeOnly;
            return cpCalendarPerm;

        }
        public CalendarPermissionType FreeBusyTimeAndSubjectAndLocation(string emEmailaddress)
        {
            CalendarPermissionType cpCalendarPerm = new CalendarPermissionType();
            UserIdType auAceUser = new UserIdType();
            if (emEmailaddress.ToLower() == "default" | emEmailaddress.ToLower() == "anonymous")
            {
                auAceUser.DistinguishedUserSpecified = true;
                if (emEmailaddress.ToLower() == "default")
                {
                    auAceUser.DistinguishedUser = DistinguishedUserType.Default;
                }
                else
                {
                    auAceUser.DistinguishedUser = DistinguishedUserType.Anonymous;
                }
            }
            else
            {
                auAceUser.PrimarySmtpAddress = emEmailaddress;
            }
            cpCalendarPerm.UserId = auAceUser;
            cpCalendarPerm.CalendarPermissionLevel = CalendarPermissionLevelType.FreeBusyTimeAndSubjectAndLocation;
            return cpCalendarPerm;

        }
        public CalendarPermissionType NonePermissions(string emEmailaddress)
        {
            CalendarPermissionType cpCalendarPerm = new CalendarPermissionType();
            UserIdType auAceUser = new UserIdType();
            if (emEmailaddress.ToLower() == "default" | emEmailaddress.ToLower() == "anonymous")
            {
                auAceUser.DistinguishedUserSpecified = true;
                if (emEmailaddress.ToLower() == "default")
                {
                    auAceUser.DistinguishedUser = DistinguishedUserType.Default;
                }
                else
                {
                    auAceUser.DistinguishedUser = DistinguishedUserType.Anonymous;
                }
            }
            else
            {
                auAceUser.PrimarySmtpAddress = emEmailaddress;
            }
            cpCalendarPerm.UserId = auAceUser;
            cpCalendarPerm.CalendarPermissionLevel = CalendarPermissionLevelType.None;
            return cpCalendarPerm;        
        }
        public PermissionType FolderEditorPermissions(string emEmailaddress)
        {
            PermissionType fpFolderPerm = new PermissionType();
            UserIdType auAceUser = new UserIdType();
            if (emEmailaddress.ToLower() == "default" | emEmailaddress.ToLower() == "anonymous")
            {
                auAceUser.DistinguishedUserSpecified = true;
                if (emEmailaddress.ToLower() == "default")
                {
                    auAceUser.DistinguishedUser = DistinguishedUserType.Default;
                }
                else
                {
                    auAceUser.DistinguishedUser = DistinguishedUserType.Anonymous;
                }
            }
            else
            {
                auAceUser.PrimarySmtpAddress = emEmailaddress;
            }
            fpFolderPerm.UserId = auAceUser;
            fpFolderPerm.PermissionLevel = PermissionLevelType.Custom;
            fpFolderPerm.ReadItemsSpecified = true;
            fpFolderPerm.ReadItems = PermissionReadAccessType.FullDetails;
            fpFolderPerm.EditItemsSpecified = true;
            fpFolderPerm.EditItems = PermissionActionType.All;
            fpFolderPerm.CanCreateItems = true;
            fpFolderPerm.CanCreateItemsSpecified = true;
            fpFolderPerm.DeleteItemsSpecified = true;
            fpFolderPerm.DeleteItems = PermissionActionType.All;
            fpFolderPerm.IsFolderVisibleSpecified = true;
            fpFolderPerm.IsFolderVisible = true;
            return fpFolderPerm;

        }
        public PermissionType FolderPublishingEditorPermissions(string emEmailaddress)
        {
            PermissionType fpFolderPerm = this.FolderEditorPermissions(emEmailaddress);
            fpFolderPerm.CanCreateSubFolders = true;
            fpFolderPerm.CanCreateSubFoldersSpecified = true;
            return fpFolderPerm;

        }
        public PermissionType FolderOwnerPermissions(string emEmailaddress)
        {
            PermissionType fpFolderPerm = this.FolderEditorPermissions(emEmailaddress);
            fpFolderPerm.CanCreateSubFolders = true;
            fpFolderPerm.CanCreateSubFoldersSpecified = true;
            fpFolderPerm.IsFolderOwner = true;
            fpFolderPerm.IsFolderOwnerSpecified = true;

            return fpFolderPerm;

        }
        public PermissionType FolderAuthorPermissions(string emEmailaddress)
        {
            PermissionType fpFolderPerm = this.FolderEditorPermissions(emEmailaddress);
            fpFolderPerm.EditItems = PermissionActionType.Owned;
            fpFolderPerm.DeleteItems = PermissionActionType.Owned;
            return fpFolderPerm;

        }
        public PermissionType FolderNonEditingAuthorPermissions(string emEmailaddress)
        {
            PermissionType fpFolderPerm = this.FolderEditorPermissions(emEmailaddress);
            fpFolderPerm.EditItems = PermissionActionType.None;
            fpFolderPerm.DeleteItems = PermissionActionType.Owned;
            return fpFolderPerm;

        }
        public PermissionType FolderPublishingAuthorPermissions(string emEmailaddress)
        {
            PermissionType fpFolderPerm = this.FolderEditorPermissions(emEmailaddress);
            fpFolderPerm.EditItems = PermissionActionType.Owned;
            fpFolderPerm.DeleteItems = PermissionActionType.Owned;
            fpFolderPerm.CanCreateSubFolders = true;
            fpFolderPerm.CanCreateSubFoldersSpecified = true;
            return fpFolderPerm;

        }
        public PermissionType FolderReviewer(string emEmailaddress)
        {
            PermissionType fpFolderPerm = new PermissionType();
            UserIdType auAceUser = new UserIdType();
            if (emEmailaddress.ToLower() == "default" | emEmailaddress.ToLower() == "anonymous")
            {
                auAceUser.DistinguishedUserSpecified = true;
                if (emEmailaddress.ToLower() == "default")
                {
                    auAceUser.DistinguishedUser = DistinguishedUserType.Default;
                }
                else
                {
                    auAceUser.DistinguishedUser = DistinguishedUserType.Anonymous;
                }
            }
            else
            {
                auAceUser.PrimarySmtpAddress = emEmailaddress;
            }
            fpFolderPerm.UserId = auAceUser;
            fpFolderPerm.PermissionLevel = PermissionLevelType.Custom;
            fpFolderPerm.ReadItemsSpecified = true;
            fpFolderPerm.ReadItems = PermissionReadAccessType.FullDetails;
            fpFolderPerm.EditItemsSpecified = true;
            fpFolderPerm.EditItems = PermissionActionType.None;
            fpFolderPerm.CanCreateItems = false;
            fpFolderPerm.CanCreateItemsSpecified = true;
            fpFolderPerm.DeleteItemsSpecified = true;
            fpFolderPerm.DeleteItems = PermissionActionType.None;
            fpFolderPerm.IsFolderVisibleSpecified = true;
            fpFolderPerm.IsFolderVisible = true;
            return fpFolderPerm;

        }
        public PermissionType FolderContributer(string emEmailaddress)
        {
            PermissionType fpFolderPerm = new PermissionType();
            UserIdType auAceUser = new UserIdType();
            if (emEmailaddress.ToLower() == "default" | emEmailaddress.ToLower() == "anonymous")
            {
                auAceUser.DistinguishedUserSpecified = true;
                if (emEmailaddress.ToLower() == "default")
                {
                    auAceUser.DistinguishedUser = DistinguishedUserType.Default;
                }
                else
                {
                    auAceUser.DistinguishedUser = DistinguishedUserType.Anonymous;
                }
            }
            else
            {
                auAceUser.PrimarySmtpAddress = emEmailaddress;
            }
            fpFolderPerm.UserId = auAceUser;
            fpFolderPerm.PermissionLevel = PermissionLevelType.Custom;
            fpFolderPerm.ReadItemsSpecified = true;
            fpFolderPerm.ReadItems = PermissionReadAccessType.None;
            fpFolderPerm.EditItemsSpecified = true;
            fpFolderPerm.EditItems = PermissionActionType.None;
            fpFolderPerm.CanCreateItems = true;
            fpFolderPerm.CanCreateItemsSpecified = true;
            fpFolderPerm.DeleteItemsSpecified = true;
            fpFolderPerm.DeleteItems = PermissionActionType.None;
            fpFolderPerm.IsFolderVisibleSpecified = true;
            fpFolderPerm.IsFolderVisible = true;
            return fpFolderPerm;

        }
        public PermissionType FolderEmpty(string emEmailaddress)
        {
            PermissionType fpFolderPerm = new PermissionType();
            UserIdType auAceUser = new UserIdType();
            if (emEmailaddress.ToLower() == "default" | emEmailaddress.ToLower() == "anonymous")
            {
                auAceUser.DistinguishedUserSpecified = true;
                if (emEmailaddress.ToLower() == "default")
                {
                    auAceUser.DistinguishedUser = DistinguishedUserType.Default;
                }
                else
                {
                    auAceUser.DistinguishedUser = DistinguishedUserType.Anonymous;
                }
            }
            else
            {
                auAceUser.PrimarySmtpAddress = emEmailaddress;
            }
            fpFolderPerm.UserId = auAceUser;
            fpFolderPerm.PermissionLevel = PermissionLevelType.None;
            return fpFolderPerm;

        }
        public List<BaseFolderType> GetFolder(BaseFolderIdType[] biArray)
        {
            try
            {
                List<BaseFolderType> rtReturnList = new List<BaseFolderType>();
                GetFolderType gfGetfolder = new GetFolderType();
                FolderResponseShapeType fsFoldershape = new FolderResponseShapeType();
                fsFoldershape.BaseShape = DefaultShapeNamesType.AllProperties;
                DistinguishedFolderIdType rfRootFolder = new DistinguishedFolderIdType();
                gfGetfolder.FolderIds = biArray;
                gfGetfolder.FolderShape = fsFoldershape;
                GetFolderResponseType fldResponse = intebExchangeServiceBinding.GetFolder(gfGetfolder);
                ArrayOfResponseMessagesType aormt = fldResponse.ResponseMessages;
                ResponseMessageType[] rmta = aormt.Items;
                foreach (ResponseMessageType rmt in rmta)
                {
                    if (rmt.ResponseClass == ResponseClassType.Success)
                    {
                        FolderInfoResponseMessageType firmt;
                        firmt = (rmt as FolderInfoResponseMessageType);
                        BaseFolderType[] rfFolders = firmt.Folders;
                        rtReturnList.Add(rfFolders[0]);
                    }
                    else
                    {
                        //Deal with Error
                    }

                }
                return rtReturnList;
            }
            catch (Exception exException)
            {
                Console.WriteLine(exException.ToString());
                throw new GetFolderException("Get Folder Exception");
            }
            
         }
        public List<BaseFolderType> FindFolder(BaseFolderIdType[] biArray, String fnFolderName)
        {
            try
            {
                List<BaseFolderType> rtReturnList = new List<BaseFolderType>();
                FindFolderType fiFindFolder = new FindFolderType();
                fiFindFolder.Traversal = FolderQueryTraversalType.Shallow;
                FolderResponseShapeType rsResponseShape = new FolderResponseShapeType();
                rsResponseShape.BaseShape = DefaultShapeNamesType.AllProperties;
                fiFindFolder.FolderShape = rsResponseShape;
                fiFindFolder.ParentFolderIds = biArray;
                RestrictionType ffRestriction = new RestrictionType();
                IsEqualToType ieToType = new IsEqualToType();
                PathToUnindexedFieldType fnFolderNameField = new PathToUnindexedFieldType();
                fnFolderNameField.FieldURI = UnindexedFieldURIType.folderDisplayName;

                FieldURIOrConstantType ciConstantType = new FieldURIOrConstantType();
                ConstantValueType cvConstantValueType = new ConstantValueType();
                cvConstantValueType.Value = fnFolderName;
                ciConstantType.Item = cvConstantValueType;
                ieToType.Item = fnFolderNameField;
                ieToType.FieldURIOrConstant = ciConstantType;
                ffRestriction.Item = ieToType;
                fiFindFolder.Restriction = ffRestriction;

                FindFolderResponseType findFolderResponse = intebExchangeServiceBinding.FindFolder(fiFindFolder);
                ResponseMessageType[] rmta = findFolderResponse.ResponseMessages.Items;
                foreach (ResponseMessageType rmt in rmta)
                {
                    if (((FindFolderResponseMessageType)rmt).ResponseClass == ResponseClassType.Success)
                    {
                        FindFolderResponseMessageType ffResponse = (FindFolderResponseMessageType)rmt;
                        if (ffResponse.RootFolder.TotalItemsInView > 0)
                        {
                            foreach (BaseFolderType suSubFolder in ffResponse.RootFolder.Folders)
                            {
                                rtReturnList.Add(suSubFolder);
                            }

                        }
                        else
                        {
                            throw new FindFolderException("Find Folder Exception No Folder");
                        }
                    }
                    else
                    {
                        throw new FindFolderException("Find Folder Exception Seach Error");
                    }

                }

                return rtReturnList;
            }
            catch (Exception exException)
            {
                Console.WriteLine(exException.ToString());
                throw new FindFolderException("Get Folder Exception");
            }
        }
        public string enumOutlookRole(CalendarPermissionType cpCalendarPermissions) {
            string orOutlookRole = "Error Cant Determine";
            if (cpCalendarPermissions.CalendarPermissionLevel == CalendarPermissionLevelType.Custom)
            {
                string hpHexperm = hexCalPerms(cpCalendarPermissions);
                switch (hpHexperm)
                {
                    case "01-00-02-02-00-04-01":
                        orOutlookRole = "Editor";
                        break;
                    case "01-01-02-02-00-04-01":
                        orOutlookRole = "PublishingEditor";
                        break;
                    case "01-01-02-02-01-04-01":
                        orOutlookRole = "Owner";
                        break;
                    case "01-00-03-03-00-04-01":
                        orOutlookRole = "Author";
                        break;
                    case "01-00-03-01-00-04-01":
                        orOutlookRole = "NonEditingAuthor";
                        break;
                    case "01-01-03-03-00-04-01":
                        orOutlookRole = "PublishingAuthor";
                        break;
                    case "00-00-01-01-00-04-01":
                        orOutlookRole = "Reviewer";
                        break;
                    case "01-00-01-01-00-01-01":
                        orOutlookRole = "Contributer";
                        break;

                }
            }
            else {
                orOutlookRole = cpCalendarPermissions.CalendarPermissionLevel.ToString();            
            }
            return orOutlookRole;
        
        }
        private String hexCalPerms(CalendarPermissionType cpCalendarPermissions){
            byte[] pmPermMask = new byte[7];
            if (cpCalendarPermissions.CanCreateItemsSpecified == true) {
                if (cpCalendarPermissions.CanCreateItems == true)
                {
                    pmPermMask[0] = 0x1;
                }
            }
            if (cpCalendarPermissions.CanCreateSubFoldersSpecified == true) {
                if (cpCalendarPermissions.CanCreateSubFolders == true)
                {
                    pmPermMask[1] = 0x1;
                }
            }
            if (cpCalendarPermissions.DeleteItemsSpecified == true) {
                switch (cpCalendarPermissions.DeleteItems ) { 
                    case PermissionActionType.None:
                        pmPermMask[2] = 0x1;
                        break;
                    case PermissionActionType.All:
                        pmPermMask[2] = 0x2;
                        break;
                    case PermissionActionType.Owned:
                        pmPermMask[2] = 0x3;
                        break;
                                        
                } 
            }
            if (cpCalendarPermissions.EditItemsSpecified == true){
                switch (cpCalendarPermissions.EditItems){
                    case PermissionActionType.None :
                        pmPermMask[3] = 0x1;
                        break;
                    case PermissionActionType.All:
                        pmPermMask[3] = 0x2;
                        break ;
                    case PermissionActionType.Owned:
                        pmPermMask[3] = 0x3;
                        break;
                     }            
            }
            if (cpCalendarPermissions.IsFolderOwnerSpecified == true){
                if (cpCalendarPermissions.IsFolderOwner == true){
                    pmPermMask[4] = 0x1;
                }
            
            }
            if (cpCalendarPermissions.ReadItemsSpecified == true) { 
                switch (cpCalendarPermissions.ReadItems){
                    case CalendarPermissionReadAccessType.None:
                        pmPermMask[5] = 0x1;
                        break;
                    case CalendarPermissionReadAccessType.TimeOnly:
                        pmPermMask[5] = 0x2;
                        break;
                    case CalendarPermissionReadAccessType.TimeAndSubjectAndLocation:
                        pmPermMask[5] = 0x3;
                        break;
                    case CalendarPermissionReadAccessType.FullDetails:
                        pmPermMask[5] = 0x4;
                        break;
                
                }
            }
            if (cpCalendarPermissions.IsFolderVisibleSpecified == true) {
                if (cpCalendarPermissions.IsFolderVisible == true) {
                    pmPermMask[6] = 0x1;
                }
            }
            return BitConverter.ToString(pmPermMask);
 
        }
        public PermissionSetType GetFreeBusyDACL(String emEmailAddress)
        {
            PermissionSetType fbPermissionSet = new PermissionSetType();
            BaseFolderIdType[] bfBaseFolderArray = new BaseFolderIdType[1];
            DistinguishedFolderIdType dfRootFolderID = new DistinguishedFolderIdType();
            if (intebExchangeServiceBinding.ExchangeImpersonation == null)
            {
                EmailAddressType mbMailbox = new EmailAddressType();
                mbMailbox.EmailAddress = emEmailAddress;
                dfRootFolderID.Mailbox = mbMailbox;
            }
            dfRootFolderID.Id = DistinguishedFolderIdNameType.msgfolderroot;
            bfBaseFolderArray[0] = dfRootFolderID;
            List<BaseFolderType>rfRootFolder = GetFolder(bfBaseFolderArray);
            BaseFolderIdType[] bfIdArray = new BaseFolderIdType[1];
            bfIdArray[0] = rfRootFolder[0].ParentFolderId;
            List<BaseFolderType> fbFolder = FindFolder(bfIdArray, "Freebusy Data");
            FolderResponseShapeType frFolderRShape = new FolderResponseShapeType();
            frFolderRShape.BaseShape = DefaultShapeNamesType.AllProperties;
            GetFolderType gfRequest = new GetFolderType();
            gfRequest.FolderIds = new BaseFolderIdType[1] { fbFolder[0].FolderId };
            intFreeBusyFolderID = fbFolder[0].FolderId;
            gfRequest.FolderShape = frFolderRShape;
            GetFolderResponseType gfGetFolderResponse = intebExchangeServiceBinding.GetFolder(gfRequest);
            FolderType cfCurrentFolder = null;
            if (gfGetFolderResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
            {
                cfCurrentFolder = (FolderType)((FolderInfoResponseMessageType)gfGetFolderResponse.ResponseMessages.Items[0]).Folders[0];
                fbPermissionSet = cfCurrentFolder.PermissionSet;
            }
            else
            {
                throw new GetFolderException("Error During Getfolder request : " + gfGetFolderResponse.ResponseMessages.Items[0].MessageText.ToString());
            }
            return fbPermissionSet;
        
        }
        private FolderIdType intCalendarFolderID = null;
        private FolderIdType intFreeBusyFolderID = null;
        private String intEmailAddress = "";
        private ExchangeServiceBinding intebExchangeServiceBinding = null;
        private ExchangeServiceBinding createesb(String EmailAddress, Boolean Impersonate, string UserName, string Password, string Domain, string EwsURL)
        {
            ExchangeServiceBinding ebExchangeServiceBinding = new ExchangeServiceBinding();
            ebExchangeServiceBinding.RequestServerVersionValue = new RequestServerVersion();
            ebExchangeServiceBinding.RequestServerVersionValue.Version = ExchangeVersionType.Exchange2007_SP1;
            if (Impersonate == true)
            {
                ConnectingSIDType csConSid = new ConnectingSIDType();
                csConSid.PrimarySmtpAddress = EmailAddress;
                ExchangeImpersonationType exImpersonate = new ExchangeImpersonationType();
                exImpersonate.ConnectingSID = csConSid;
                ebExchangeServiceBinding.ExchangeImpersonation = exImpersonate;

            }
            if (UserName == "")
            {
                ebExchangeServiceBinding.UseDefaultCredentials = true;
            }
            else
            {
                NetworkCredential ncNetCredential = new NetworkCredential(UserName, Password, Domain);
                ebExchangeServiceBinding.Credentials = ncNetCredential;
            }
            if (EwsURL == "")
            {
                String caCasURL = DiscoverCAS();
                EwsURL = DiscoverEWSURL(caCasURL, EmailAddress, UserName, Password, Domain);
            }
            ebExchangeServiceBinding.Url = EwsURL;
            intebExchangeServiceBinding = ebExchangeServiceBinding;
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
        private String DiscoverEWSURL(string caCASURL, string emEmailAddress, string UserName, string Password, string Domain)
        {

            String EWSURL = null;
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
            XmlNodeList EWSNodes = reResponseDoc.GetElementsByTagName("ASUrl");
            if (EWSNodes.Count != 0)
            {
                EWSURL = EWSNodes[0].InnerText;
            }
            else
            {
                throw new AutoDiscoveryException("Error during AutoDiscovery");
            }
            return EWSURL;

        }
        private String SetCalendarDACL()
        {
            CalendarPermissionSetType cfNewCalPermsionsSet = new CalendarPermissionSetType();
            cfNewCalPermsionsSet.CalendarPermissions = new CalendarPermissionType[CalendarDACL.Count];
            Hashtable cdCheckDuplicates = new Hashtable();
            for (int cpint1 = 0; cpint1 < CalendarDACL.Count; cpint1++)
            {
                if (CalendarDACL[cpint1].UserId.DistinguishedUserSpecified == false)
                {
                    if (cdCheckDuplicates.ContainsKey(CalendarDACL[cpint1].UserId.PrimarySmtpAddress.ToLower()))
                    {
                        cfNewCalPermsionsSet.CalendarPermissions[(int)cdCheckDuplicates[CalendarDACL[cpint1].UserId.PrimarySmtpAddress.ToLower()]] = CalendarDACL[cpint1];
                    }
                    else
                    {
                        cdCheckDuplicates.Add(CalendarDACL[cpint1].UserId.PrimarySmtpAddress.ToLower(), cpint1);
                        cfNewCalPermsionsSet.CalendarPermissions[cpint1] = CalendarDACL[cpint1];
                    }
                }
                else {
                    switch (CalendarDACL[cpint1].UserId.DistinguishedUser) {
                        case DistinguishedUserType.Default:
                            if (cdCheckDuplicates.ContainsKey("default@default"))
                            {
                                cfNewCalPermsionsSet.CalendarPermissions[(int)cdCheckDuplicates["default@default"]] = CalendarDACL[cpint1];
                            }
                            else
                            {
                                cdCheckDuplicates.Add("default@default", cpint1);
                                cfNewCalPermsionsSet.CalendarPermissions[cpint1] = CalendarDACL[cpint1];
                            }
                            break;
                        case DistinguishedUserType.Anonymous:
                            if (cdCheckDuplicates.ContainsKey("anonymous@default"))
                            {
                                cfNewCalPermsionsSet.CalendarPermissions[(int)cdCheckDuplicates["anonymous@default"]] = CalendarDACL[cpint1];
                            }
                            else
                            {
                                cdCheckDuplicates.Add("anonymous@default", cpint1);
                                cfNewCalPermsionsSet.CalendarPermissions[cpint1] = CalendarDACL[cpint1];
                            }
                            break;
                        default :
                            cfNewCalPermsionsSet.CalendarPermissions[cpint1] = CalendarDACL[cpint1];
                            break;
                    
                    
                    }
                               
                }

            
            }
            CalendarFolderType cfUpdateCalFolder = new CalendarFolderType();
            cfUpdateCalFolder.PermissionSet = cfNewCalPermsionsSet;

            UpdateFolderType upUpdateFolderRequest = new UpdateFolderType();

            FolderChangeType fcFolderchanges = new FolderChangeType();

            fcFolderchanges.Item = this.intCalendarFolderID;

            SetFolderFieldType cpCalPerms = new SetFolderFieldType();
            PathToUnindexedFieldType cpFieldURI = new PathToUnindexedFieldType();
            cpFieldURI.FieldURI = UnindexedFieldURIType.folderPermissionSet;
            cpCalPerms.Item = cpFieldURI;
            cpCalPerms.Item1 = cfUpdateCalFolder;

            fcFolderchanges.Updates = new FolderChangeDescriptionType[1] { cpCalPerms };
            upUpdateFolderRequest.FolderChanges = new FolderChangeType[1] { fcFolderchanges };
            String upres = "";
            UpdateFolderResponseType ufUpdateFolderResponse = intebExchangeServiceBinding.UpdateFolder(upUpdateFolderRequest);
            if (ufUpdateFolderResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
            {
                upres = "Permissions Updated sucessfully";
            }
            else
            {
                throw new GetFolderException("Folder update error : " + ufUpdateFolderResponse.ResponseMessages.Items[0].MessageText.ToString());  

            }
            return upres;
        }
        private String SetFreeBusyDACL()
        {
            PermissionSetType fbNewPermsionsSet = new PermissionSetType();
            fbNewPermsionsSet.Permissions = new PermissionType[FreeBusyDACL.Count];
            Hashtable cdCheckDuplicates = new Hashtable();
            for (int cpint1 = 0; cpint1 < FreeBusyDACL.Count; cpint1++)
            {
                if (FreeBusyDACL[cpint1].UserId.DistinguishedUserSpecified == false)
                {
                    if (cdCheckDuplicates.ContainsKey(FreeBusyDACL[cpint1].UserId.PrimarySmtpAddress.ToLower()))
                    {
                        fbNewPermsionsSet.Permissions[(int)cdCheckDuplicates[FreeBusyDACL[cpint1].UserId.PrimarySmtpAddress.ToLower()]] = FreeBusyDACL[cpint1];
                    }
                    else
                    {
                        cdCheckDuplicates.Add(FreeBusyDACL[cpint1].UserId.PrimarySmtpAddress.ToLower(), cpint1);
                        fbNewPermsionsSet.Permissions[cpint1] = FreeBusyDACL[cpint1];
                    }
                }
                else
                {
                    switch (FreeBusyDACL[cpint1].UserId.DistinguishedUser)
                    {
                        case DistinguishedUserType.Default:
                            if (cdCheckDuplicates.ContainsKey("default@default"))
                            {
                                fbNewPermsionsSet.Permissions[(int)cdCheckDuplicates["default@default"]] = FreeBusyDACL[cpint1];
                            }
                            else
                            {
                                cdCheckDuplicates.Add("default@default", cpint1);
                                fbNewPermsionsSet.Permissions[cpint1] = FreeBusyDACL[cpint1];
                            }
                            break;
                        case DistinguishedUserType.Anonymous:
                            if (cdCheckDuplicates.ContainsKey("anonymous@default"))
                            {
                                fbNewPermsionsSet.Permissions[(int)cdCheckDuplicates["anonymous@default"]] = FreeBusyDACL[cpint1];
                            }
                            else
                            {
                                cdCheckDuplicates.Add("anonymous@default", cpint1);
                                fbNewPermsionsSet.Permissions[cpint1] = FreeBusyDACL[cpint1];
                            }
                            break;
                        default:
                            fbNewPermsionsSet.Permissions[cpint1] = FreeBusyDACL[cpint1];
                            break;


                    }

                }


            }
            FolderType fbUpdateFolder = new FolderType();
            fbUpdateFolder.PermissionSet = fbNewPermsionsSet;

            UpdateFolderType upUpdateFolderRequest = new UpdateFolderType();

            FolderChangeType fcFolderchanges = new FolderChangeType();

            fcFolderchanges.Item = this.intFreeBusyFolderID;

            SetFolderFieldType fbFreeBusyPerms = new SetFolderFieldType();
            PathToUnindexedFieldType fbFieldURI = new PathToUnindexedFieldType();
            fbFieldURI.FieldURI = UnindexedFieldURIType.folderPermissionSet;
            fbFreeBusyPerms.Item = fbFieldURI;
            fbFreeBusyPerms.Item1 = fbUpdateFolder;

            fcFolderchanges.Updates = new FolderChangeDescriptionType[1] { fbFreeBusyPerms };
            upUpdateFolderRequest.FolderChanges = new FolderChangeType[1] { fcFolderchanges };
            String upres = "";
            UpdateFolderResponseType ufUpdateFolderResponse = intebExchangeServiceBinding.UpdateFolder(upUpdateFolderRequest);
            if (ufUpdateFolderResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
            {
                upres = "Permissions Updated sucessfully";
            }
            else
            {
                throw new GetFolderException("Folder update error : " + ufUpdateFolderResponse.ResponseMessages.Items[0].MessageText.ToString());

            }
            return upres;
        }
        private CalendarPermissionSetType GetCalendarDACL(String emEmailAddress) {

            DistinguishedFolderIdType cfCurrentCalendar = new DistinguishedFolderIdType();
            CalendarPermissionSetType cfCurrentCalPermsionsSet = null;
            cfCurrentCalendar.Id = DistinguishedFolderIdNameType.calendar;
            if (intebExchangeServiceBinding.ExchangeImpersonation == null) {
                EmailAddressType mbMailbox = new EmailAddressType();
                mbMailbox.EmailAddress = emEmailAddress;
                cfCurrentCalendar.Mailbox = mbMailbox;
            }
            FolderResponseShapeType frFolderRShape = new FolderResponseShapeType();
            frFolderRShape.BaseShape = DefaultShapeNamesType.AllProperties;

            GetFolderType gfRequest = new GetFolderType();
            gfRequest.FolderIds = new BaseFolderIdType[1] { cfCurrentCalendar };
            gfRequest.FolderShape = frFolderRShape;


            GetFolderResponseType gfGetFolderResponse = intebExchangeServiceBinding.GetFolder(gfRequest);
            CalendarFolderType cfCurrentFolder = null;
            if (gfGetFolderResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
            {
                cfCurrentFolder = (CalendarFolderType)((FolderInfoResponseMessageType)gfGetFolderResponse.ResponseMessages.Items[0]).Folders[0];
                this.intCalendarFolderID = cfCurrentFolder.FolderId;
                cfCurrentCalPermsionsSet = cfCurrentFolder.PermissionSet;
            }
            else
            {
                throw new GetFolderException("Error During Getfolder request : " + gfGetFolderResponse.ResponseMessages.Items[0].MessageText.ToString());                
            }
            return cfCurrentCalPermsionsSet;
        }
    
    }
    public class ContactUtil {
        public ContactUtil(string emEmailAddress)
        {
            initlizeContactUtil(emEmailAddress, false, "", "", "", "");
        }
        public ContactUtil(string emEmailAddress, Boolean Impersonate)
        {
            initlizeContactUtil(emEmailAddress, Impersonate, "", "", "", "");
        }
        public ContactUtil(string emEmailAddress, Boolean Impersonate, string UserName, string Password, string Domain)
        {
            initlizeContactUtil(emEmailAddress, Impersonate, UserName, Password, Domain, "");
        }
        public ContactUtil(string emEmailAddress, Boolean Impersonate, string UserName, string Password, string Domain, string casURL)
        {
            initlizeContactUtil(emEmailAddress, Impersonate, UserName, Password, Domain, casURL);
        }
        public String AddContact(Contact cnContact) {
            ContactItemType ewsContact = new ContactItemType();
            ewsContact.AssistantName = cnContact.AssistantName;
 //           ewsContact.Body.Value = cnContact.Notes;
            ewsContact.BusinessHomePage = cnContact.BusinessWebPage;
            ewsContact.CompanyName = cnContact.Company;
            //CompleteNameType cnCompleteaName = new CompleteNameType();
            //cnCompleteaName.FirstName = cnContact.FirstName;
            //cnCompleteaName.FullName = cnContact.FullName;
            //cnCompleteaName.LastName = cnContact.LastName;
            //cnCompleteaName.MiddleName = cnContact.MiddleName;
            //cnCompleteaName.Nickname = cnContact.Nickname;
            //cnCompleteaName.Suffix = cnContact.Suffix;
            //cnCompleteaName.Title = cnContact.JobTitle;
           /// ewsContact.CompleteName = cnCompleteaName;
            ewsContact.Department = cnContact.Department;
            ewsContact.GivenName = cnContact.FirstName;
            ewsContact.Surname = cnContact.LastName;
            ewsContact.OfficeLocation = cnContact.OfficeLocation;
            ewsContact.MiddleName = cnContact.MiddleName;
            ewsContact.Profession = cnContact.JobTitle;
            ewsContact.Nickname = cnContact.Nickname;
            ewsContact.FileAs = cnContact.FileAs;
            //Set Email Type and DisplayName
            ewsContact.EmailAddresses = new EmailAddressDictionaryEntryType[1];
            EmailAddressDictionaryEntryType emEmailAddress = new EmailAddressDictionaryEntryType();
            emEmailAddress.Key = EmailAddressKeyType.EmailAddress1;
            emEmailAddress.Value = cnContact.EmailAddress;
            ewsContact.EmailAddresses[0] = emEmailAddress;
            ExtendedPropertyType emEmailType = new ExtendedPropertyType();
            PathToExtendedFieldType epExPath = new PathToExtendedFieldType();
            epExPath.PropertySetId = "00062004-0000-0000-C000-000000000046";
            epExPath.PropertyId = 0x8082;
            epExPath.PropertyIdSpecified = true;
            epExPath.PropertyType = MapiPropertyTypeType.String;
            emEmailType.ExtendedFieldURI = epExPath;
            emEmailType.Item = "SMTP";
            ExtendedPropertyType emDisplayName = new ExtendedPropertyType();
            PathToExtendedFieldType epExPath1 = new PathToExtendedFieldType();
            epExPath1.PropertySetId = "00062004-0000-0000-C000-000000000046";
            epExPath1.PropertyId = 0x8080;
            epExPath1.PropertyIdSpecified = true;
            epExPath1.PropertyType = MapiPropertyTypeType.String;
            emDisplayName.ExtendedFieldURI = epExPath1;
            emDisplayName.Item = cnContact.DisplayName + "(" + cnContact.EmailAddressDisplayName + ")" ;
            //Set Outlook Card DisplayName
            ExtendedPropertyType emCardDisplayName = new ExtendedPropertyType();
            PathToExtendedFieldType epExPath2 = new PathToExtendedFieldType();
            epExPath2.PropertySetId = "00062004-0000-0000-C000-000000000046";
            epExPath2.PropertyId = 0x8084;
            epExPath2.PropertyIdSpecified = true;
            epExPath2.PropertyType = MapiPropertyTypeType.String;
            emCardDisplayName.ExtendedFieldURI = epExPath2;
            emCardDisplayName.Item = cnContact.EmailAddress;
            ewsContact.Subject = cnContact.DisplayName;
            ewsContact.ExtendedProperty = new ExtendedPropertyType[3];
            ewsContact.ExtendedProperty[0] = emEmailType;
            ewsContact.ExtendedProperty[1] = emDisplayName;
            ewsContact.ExtendedProperty[2] = emCardDisplayName;
            //Set Address Information
            ewsContact.PhysicalAddresses = new PhysicalAddressDictionaryEntryType[2];
            PhysicalAddressDictionaryEntryType hmHomeAddress = new PhysicalAddressDictionaryEntryType();
            hmHomeAddress.Key = PhysicalAddressKeyType.Home;
            hmHomeAddress.City = cnContact.HomeCity;
            hmHomeAddress.CountryOrRegion = cnContact.HomeCountry;
            hmHomeAddress.PostalCode = cnContact.HomePostalCode;
            hmHomeAddress.State = cnContact.HomeState;
            hmHomeAddress.Street = cnContact.HomeStreet;
            PhysicalAddressDictionaryEntryType bsBusinessAddress = new PhysicalAddressDictionaryEntryType();
            bsBusinessAddress.Key = PhysicalAddressKeyType.Business;
            bsBusinessAddress.City = cnContact.BusinessCity;
            bsBusinessAddress.CountryOrRegion = cnContact.BusinessCountry;
            bsBusinessAddress.PostalCode = cnContact.BusinessPostalCode;
            bsBusinessAddress.State = cnContact.BusinessState;
            bsBusinessAddress.Street = cnContact.BusinessStreet;          
            ewsContact.PhysicalAddresses[0] = hmHomeAddress;
            ewsContact.PhysicalAddresses[1] = bsBusinessAddress;
            // Telephone Numbers
            ewsContact.PhoneNumbers = new PhoneNumberDictionaryEntryType[5];
            PhoneNumberDictionaryEntryType hmHomePhone = new PhoneNumberDictionaryEntryType();
            hmHomePhone.Key = PhoneNumberKeyType.HomePhone;
            hmHomePhone.Value = cnContact.HomePhone;
            PhoneNumberDictionaryEntryType hmHomeFax = new PhoneNumberDictionaryEntryType();
            hmHomeFax.Key = PhoneNumberKeyType.HomeFax;
            hmHomeFax.Value = cnContact.HomeFax;
            PhoneNumberDictionaryEntryType bsBusinessPhone = new PhoneNumberDictionaryEntryType();
            bsBusinessPhone.Key = PhoneNumberKeyType.BusinessPhone;
            bsBusinessPhone.Value = cnContact.BusinessPhone;
            PhoneNumberDictionaryEntryType bsBusinessFax = new PhoneNumberDictionaryEntryType();
            bsBusinessFax.Key = PhoneNumberKeyType.BusinessFax;
            bsBusinessFax.Value = cnContact.BusinessFax;
            PhoneNumberDictionaryEntryType mpMobilePhone = new PhoneNumberDictionaryEntryType();
            mpMobilePhone.Key = PhoneNumberKeyType.MobilePhone;
            mpMobilePhone.Value = cnContact.MobilePhone;
            ewsContact.PhoneNumbers[0] = hmHomePhone;
            ewsContact.PhoneNumbers[1] = hmHomeFax;
            ewsContact.PhoneNumbers[2] = bsBusinessPhone;
            ewsContact.PhoneNumbers[3] = bsBusinessFax;
            ewsContact.PhoneNumbers[4] = mpMobilePhone;

            ItemIdType iiItemid = new ItemIdType();
            CreateItemType ciCreateItemRequest = new CreateItemType();
            ciCreateItemRequest.MessageDisposition = MessageDispositionType.SaveOnly;
            ciCreateItemRequest.MessageDispositionSpecified = true;
            ciCreateItemRequest.SavedItemFolderId = new TargetFolderIdType();
            DistinguishedFolderIdType cfContactsFolder = new DistinguishedFolderIdType();
            cfContactsFolder.Id = DistinguishedFolderIdNameType.contacts;
            ciCreateItemRequest.SavedItemFolderId.Item = cfContactsFolder;
            ciCreateItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            ciCreateItemRequest.Items.Items = new ItemType[1];
            ciCreateItemRequest.Items.Items[0] = ewsContact;

            CreateItemResponseType createItemResponse = intebExchangeServiceBinding.CreateItem(ciCreateItemRequest);
            if (createItemResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Error)
            {
                Console.WriteLine("Error Occured");
                Console.WriteLine(createItemResponse.ResponseMessages.Items[0].MessageText);
            }
            else
            {
                ItemInfoResponseMessageType rmResponseMessage = createItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
                Console.WriteLine("Item was created");
                Console.WriteLine("Item ID : " + rmResponseMessage.Items.Items[0].ItemId.Id.ToString());
                Console.WriteLine("ChangeKey : " + rmResponseMessage.Items.Items[0].ItemId.ChangeKey.ToString());
                iiItemid.Id = rmResponseMessage.Items.Items[0].ItemId.Id.ToString();
                iiItemid.ChangeKey = rmResponseMessage.Items.Items[0].ItemId.ChangeKey.ToString();
            }

                  


            return "done";
        
        
        
        }
        public String UpdateContact(String emEmailAddress, Contact cnContact) {
            if (cnContact.AssistantName != "")
            {
                SetItemFieldType sfAssitantName = new SetItemFieldType();
                sfAssitantName.Item = new PathToUnindexedFieldType();
                ((PathToUnindexedFieldType)sfAssitantName.Item).FieldURI = UnindexedFieldURIType.contactsAssistantName;
                sfAssitantName.Item1 = new ContactItemType();
                ((ContactItemType)sfAssitantName.Item1).AssistantName = cnContact.AssistantName;
            }

            return "Done";
        
        
        }

        private void initlizeContactUtil(string EmailAddress, Boolean Impersonate, string UserName, string Password, string Domain, string ewsURL)
        {
            try
            {
                ServicePointManager.ServerCertificateValidationCallback =
                delegate(Object obj, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
                {
                    //   Ignore Self Signed Certs
                    return true;
                };
                intEmailAddress = EmailAddress;
                ExchangeServiceBinding ebExchangeServiceBinding = createesb(EmailAddress, Impersonate, UserName, Password, Domain, ewsURL);
            }
            catch (Exception exException)
            {
                Console.WriteLine(exException.ToString());
            }
        }
        private String intEmailAddress = "";
        private ExchangeServiceBinding intebExchangeServiceBinding = null;
        private ExchangeServiceBinding createesb(String EmailAddress, Boolean Impersonate, string UserName, string Password, string Domain, string EwsURL)
        {
            ExchangeServiceBinding ebExchangeServiceBinding = new ExchangeServiceBinding();
            ebExchangeServiceBinding.RequestServerVersionValue = new RequestServerVersion();
            ebExchangeServiceBinding.RequestServerVersionValue.Version = ExchangeVersionType.Exchange2007_SP1;
            if (Impersonate == true)
            {
                ConnectingSIDType csConSid = new ConnectingSIDType();
                csConSid.PrimarySmtpAddress = EmailAddress;
                ExchangeImpersonationType exImpersonate = new ExchangeImpersonationType();
                exImpersonate.ConnectingSID = csConSid;
                ebExchangeServiceBinding.ExchangeImpersonation = exImpersonate;

            }
            if (UserName == "")
            {
                ebExchangeServiceBinding.UseDefaultCredentials = true;
            }
            else
            {
                NetworkCredential ncNetCredential = new NetworkCredential(UserName, Password, Domain);
                ebExchangeServiceBinding.Credentials = ncNetCredential;
            }
            if (EwsURL == "")
            {
                String caCasURL = DiscoverCAS();
                EwsURL = DiscoverEWSURL(caCasURL, EmailAddress, UserName, Password, Domain);
            }
            ebExchangeServiceBinding.Url = EwsURL;
            intebExchangeServiceBinding = ebExchangeServiceBinding;
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
        private String DiscoverEWSURL(string caCASURL, string emEmailAddress, string UserName, string Password, string Domain)
        {

            String EWSURL = null;
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
            XmlNodeList EWSNodes = reResponseDoc.GetElementsByTagName("ASUrl");
            if (EWSNodes.Count != 0)
            {
                EWSURL = EWSNodes[0].InnerText;
            }
            else
            {
                throw new AutoDiscoveryException("Error during AutoDiscovery");
            }
            return EWSURL;

        }
    
    }
    public class QueryFolder
    {
        public QueryFolder(EWSConnection ewc, DistinguishedFolderIdType fiFolderID, Duration duDuration, Boolean urUnreadOnly, Boolean cfCalendar)
        {
            FindItemType fiFindItemRequest = new FindItemType();
            DistinguishedFolderIdType[] faFolderIDArray = new DistinguishedFolderIdType[1];
            faFolderIDArray[0] = new DistinguishedFolderIdType();
            faFolderIDArray[0].Mailbox = new EmailAddressType();
            faFolderIDArray[0].Mailbox.EmailAddress = ewc.emEmailAddress;
            faFolderIDArray[0].Id = fiFolderID.Id;
            fiFindItemRequest.ParentFolderIds = faFolderIDArray;
            EWSFindItems(ewc.esb, fiFindItemRequest, duDuration, urUnreadOnly, cfCalendar);
          
        }
        public QueryFolder(EWSConnection ewc, FolderIdType fiFolderID, Duration duDuration, Boolean urUnreadOnly, Boolean cfCalendar) {
            FindItemType fiFindItemRequest = new FindItemType();
            FolderIdType[] faFolderIDArray = new FolderIdType[1];
            faFolderIDArray[0] = new FolderIdType();
            faFolderIDArray[0].Id = fiFolderID.Id;
            fiFindItemRequest.ParentFolderIds = faFolderIDArray;
            EWSFindItems(ewc.esb, fiFindItemRequest, duDuration, urUnreadOnly, cfCalendar);
        }
        private void EWSFindItems(ExchangeServiceBinding esb, FindItemType fiFindItemRequest, Duration duDuration, Boolean urUnreadOnly, Boolean cfCalendar)
        {
            try
            {
                intFiFolderItems = new List<ItemType>();

                fiFindItemRequest.Traversal = ItemQueryTraversalType.Shallow;

                ItemResponseShapeType ipItemProperties = new ItemResponseShapeType();
                ipItemProperties.BaseShape = DefaultShapeNamesType.AllProperties;
                fiFindItemRequest.ItemShape = ipItemProperties;

                RestrictionType ffRestriction = new RestrictionType();


                if (cfCalendar == true)
                {
                    CalendarViewType cpCalendarPageing = new CalendarViewType();
                    cpCalendarPageing.StartDate = duDuration.StartTime;
                    cpCalendarPageing.EndDate = duDuration.EndTime;
                    fiFindItemRequest.Item = cpCalendarPageing;
                }
                else
                {
                    if (urUnreadOnly == true & duDuration == null)
                    {
                        IsEqualToType ieToTypeRead = new IsEqualToType();
                        PathToUnindexedFieldType rsReadStatus = new PathToUnindexedFieldType();
                        rsReadStatus.FieldURI = UnindexedFieldURIType.messageIsRead;
                        ieToTypeRead.Item = rsReadStatus;
                        FieldURIOrConstantType constantType = new FieldURIOrConstantType();
                        ConstantValueType constantValueType = new ConstantValueType();
                        constantValueType.Value = "0";
                        constantType.Item = constantValueType;
                        ieToTypeRead.Item = rsReadStatus;
                        ieToTypeRead.FieldURIOrConstant = constantType;
                        ffRestriction.Item = ieToTypeRead;
                        fiFindItemRequest.Restriction = ffRestriction;
                    }
                    else
                    {
                        if (duDuration != null)
                        {
                            AndType raRestictionAnd = new AndType();
                            IsGreaterThanOrEqualToType igteToType = new IsGreaterThanOrEqualToType();
                            PathToUnindexedFieldType mfModifiedTime = new PathToUnindexedFieldType();
                            mfModifiedTime.FieldURI = UnindexedFieldURIType.itemLastModifiedTime;
                            igteToType.Item = mfModifiedTime;
                            igteToType.FieldURIOrConstant = new FieldURIOrConstantType();
                            igteToType.FieldURIOrConstant.Item = new ConstantValueType();
                            (igteToType.FieldURIOrConstant.Item as ConstantValueType).Value = duDuration.StartTime.ToUniversalTime().ToString("yyyy-MM-ddThh:mm:ssZ");

                            IsLessThanOrEqualToType ilteToType = new IsLessThanOrEqualToType();
                            ilteToType.Item = mfModifiedTime;
                            ilteToType.FieldURIOrConstant = new FieldURIOrConstantType();
                            ilteToType.FieldURIOrConstant.Item = new ConstantValueType();
                            (ilteToType.FieldURIOrConstant.Item as ConstantValueType).Value = duDuration.EndTime.ToUniversalTime().ToString("yyyy-MM-ddThh:mm:ssZ");

                            if (urUnreadOnly == false)
                            {
                                raRestictionAnd.Items = new SearchExpressionType[2];
                                raRestictionAnd.Items[0] = igteToType;
                                raRestictionAnd.Items[1] = ilteToType;
                            }
                            else
                            {
                                raRestictionAnd.Items = new SearchExpressionType[3];
                                IsEqualToType ieToTypeRead = new IsEqualToType();
                                PathToUnindexedFieldType rsReadStatus = new PathToUnindexedFieldType();
                                rsReadStatus.FieldURI = UnindexedFieldURIType.messageIsRead;
                                ieToTypeRead.Item = rsReadStatus;
                                FieldURIOrConstantType constantType = new FieldURIOrConstantType();
                                ConstantValueType constantValueType = new ConstantValueType();
                                constantValueType.Value = "0";
                                constantType.Item = constantValueType;
                                ieToTypeRead.Item = rsReadStatus;
                                ieToTypeRead.FieldURIOrConstant = constantType;

                                raRestictionAnd.Items[0] = igteToType;
                                raRestictionAnd.Items[1] = ilteToType;
                                raRestictionAnd.Items[2] = ieToTypeRead;

                            }
                            ffRestriction.Item = raRestictionAnd;



                        }

                    }

                }
                if (ffRestriction.Item != null)
                {
                    fiFindItemRequest.Restriction = ffRestriction;
                }
                FindItemResponseType frFindItemResponse = esb.FindItem(fiFindItemRequest);
                if (frFindItemResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
                {
                    foreach (FindItemResponseMessageType firmtMessage in frFindItemResponse.ResponseMessages.Items)
                    {
                        Console.WriteLine("Number of Items for feed : " + firmtMessage.RootFolder.TotalItemsInView);
                        if (firmtMessage.RootFolder.TotalItemsInView > 0)
                        {
                            foreach (ItemType miMailboxItem in ((ArrayOfRealItemsType)firmtMessage.RootFolder.Item).Items)
                            {
                                intFiFolderItems.Add(miMailboxItem);
                             }
                        }


                    }
                }
                else
                {
                    throw new FindItemException("Error During FindItem request : " + frFindItemResponse.ResponseMessages.Items[0].MessageText.ToString());
                }

            }
            catch (Exception exException)
            {
                Console.WriteLine(exException.ToString());
               // return exException.ToString();
            }
        }
    public List<ItemType> fiFolderItems {
            get
            {
                return intFiFolderItems;
            }
        
        }
        private List<ItemType> intFiFolderItems = null;


    }
    public class EWSConnection
    {
        public EWSConnection(String emEmailAddress, Boolean Impersonate, string UserName, string Password, string Domain, string EwsURL)
        {
            ServicePointManager.ServerCertificateValidationCallback =
            delegate(Object obj, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
            {
                //   Ignore Self Signed Certs
                return true;
            };
            ExchangeServiceBinding ebExchangeServiceBinding = new ExchangeServiceBinding();
            ebExchangeServiceBinding.RequestServerVersionValue = new RequestServerVersion();
            ebExchangeServiceBinding.RequestServerVersionValue.Version = ExchangeVersionType.Exchange2007_SP1;
            if (Impersonate == true)
            {
                ConnectingSIDType csConSid = new ConnectingSIDType();
                csConSid.PrimarySmtpAddress = emEmailAddress;
                ExchangeImpersonationType exImpersonate = new ExchangeImpersonationType();
                exImpersonate.ConnectingSID = csConSid;
                ebExchangeServiceBinding.ExchangeImpersonation = exImpersonate;

            }
            if (UserName == "")
            {
                ebExchangeServiceBinding.UseDefaultCredentials = true;
            }
            else
            {
                NetworkCredential ncNetCredential = new NetworkCredential(UserName, Password, Domain);
                ebExchangeServiceBinding.Credentials = ncNetCredential;
            }
            if (EwsURL == "")
            {
                String caCasURL = DiscoverCAS();
                EwsURL = DiscoverEWSURL(caCasURL, emEmailAddress, UserName, Password, Domain);
            }
            ebExchangeServiceBinding.Url = EwsURL;
            intemEmailAddress = emEmailAddress;
            intExchangeServiceBinding = ebExchangeServiceBinding;
         
        }
        public ExchangeServiceBinding esb {
            get
            {
                return intExchangeServiceBinding;
            }
        }
        public String emEmailAddress
        {
            get
            {
                return intemEmailAddress;
            }
        }
        private ExchangeServiceBinding intExchangeServiceBinding = null;
        private String intemEmailAddress = null;
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
        private String DiscoverEWSURL(string caCASURL, string emEmailAddress, string UserName, string Password, string Domain)
        {

            String EWSURL = null;
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
            XmlNodeList EWSNodes = reResponseDoc.GetElementsByTagName("ASUrl");
            if (EWSNodes.Count != 0)
            {
                EWSURL = EWSNodes[0].InnerText;
            }
            else
            {
                throw new AutoDiscoveryException("Error during AutoDiscovery");
            }
            return EWSURL;

        }
        public void DownloadAttachment(String fnFileName, AttachmentIdType atAttachID)
        {
            GetAttachmentType gaGetAttachment = new GetAttachmentType();
            AttachmentIdType[] atArray = new AttachmentIdType[1];
            atArray[0] = atAttachID;
            gaGetAttachment.AttachmentIds = atArray;
            GetAttachmentResponseType gaGetAttachmentResponse = esb.GetAttachment(gaGetAttachment);
            foreach (AttachmentInfoResponseMessageType atAttachmentResponse in gaGetAttachmentResponse.ResponseMessages.Items)
            {
                if (atAttachmentResponse.ResponseClass == ResponseClassType.Success)
                {
                    FileAttachmentType fiFileAttachmentType = atAttachmentResponse.Attachments[0] as FileAttachmentType;
                    if (fiFileAttachmentType != null)
                    {
                        FileStream fiFile = new FileStream(fnFileName, FileMode.Create);
                        fiFile.Write(fiFileAttachmentType.Content,0, fiFileAttachmentType.Content.Length);
                        fiFile.Close();

                    }
                }
                else
                {
                    throw new DownloadAttachmentException("Error Downloding Attachment");

                }



            }
        }
        public List<ItemType> FindMessage(BaseFolderIdType[] biArray, PathToUnindexedFieldType uifieldType, String msMessageID, Boolean Dumpster)
        {
            try
            {
                List<ItemType> rtReturnList = new List<ItemType>();
                FindItemType fiFindItemRequest = new FindItemType();
                ItemResponseShapeType ipItemProperties = new ItemResponseShapeType();
                if (Dumpster == true)
                {
                    fiFindItemRequest.Traversal = ItemQueryTraversalType.SoftDeleted;
                    ipItemProperties.BaseShape = DefaultShapeNamesType.AllProperties;
                }
                else
                {
                    fiFindItemRequest.Traversal = ItemQueryTraversalType.Shallow;
                    ipItemProperties.BaseShape = DefaultShapeNamesType.IdOnly;
                }
                fiFindItemRequest.ItemShape = ipItemProperties;
                fiFindItemRequest.ParentFolderIds = biArray;
                //Add Restriction for MessageID
                RestrictionType ffRestriction = new RestrictionType();
                IsEqualToType ieToType = new IsEqualToType();
                FieldURIOrConstantType ciConstantType = new FieldURIOrConstantType();
                ConstantValueType cvConstantValueType = new ConstantValueType();
                cvConstantValueType.Value = msMessageID;
                ciConstantType.Item = cvConstantValueType;
                ieToType.Item = uifieldType;
                ieToType.FieldURIOrConstant = ciConstantType;
                ffRestriction.Item = ieToType;
                fiFindItemRequest.Restriction = ffRestriction;
                FindItemResponseType frFindItemResponse = esb.FindItem(fiFindItemRequest);
                foreach (FindItemResponseMessageType firmtMessage in frFindItemResponse.ResponseMessages.Items)
                {
                    if (firmtMessage.RootFolder.TotalItemsInView > 0)
                    {
                        foreach (ItemType miMailboxItem in ((ArrayOfRealItemsType)firmtMessage.RootFolder.Item).Items)
                        {
                            if (Dumpster == false)
                            {
                                GetItemType giRequest = new GetItemType();
                                ItemIdType iiItemId = new ItemIdType();
                                iiItemId.Id = miMailboxItem.ItemId.Id;
                                iiItemId.ChangeKey = miMailboxItem.ItemId.ChangeKey;
                                giRequest.ItemIds = new ItemIdType[1];
                                giRequest.ItemIds[0] = iiItemId;
                                giRequest.ItemShape = new ItemResponseShapeType();
                                giRequest.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;
                                giRequest.ItemShape.BodyTypeSpecified = true;
                                giRequest.ItemShape.BodyType = BodyTypeResponseType.HTML;
                                giRequest.ItemShape.IncludeMimeContent = true;
                                giRequest.ItemShape.IncludeMimeContentSpecified = true;
                                GetItemResponseType giResponse = esb.GetItem(giRequest);
                                if (giResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Error)
                                {
                                    Console.WriteLine("Error Occured");
                                    Console.WriteLine(giResponse.ResponseMessages.Items[0].MessageText);
                                }
                                else
                                {
                                    ItemInfoResponseMessageType rmResponseMessage = giResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
                                    rtReturnList.Add(rmResponseMessage.Items.Items[0]);

                                }
                            }
                            else
                            {
                                rtReturnList.Add(miMailboxItem);

                            }
                        }
                    }



                }
                return rtReturnList;
            }
            catch (Exception exException)
            {
                Console.WriteLine(exException.ToString());
                throw new FindFolderException("Find Folder Exception");
            }
        }
        public List<ItemType> RecurseFolder(BaseFolderIdType[] biArray, PathToUnindexedFieldType uifieldType, String msMessageID, Boolean Dumpster)
        {
            try
            {
                List<ItemType> rtReturnList = new List<ItemType>();
                // start by searching root
                List<ItemType> ritRootReturnList = FindMessage(biArray, uifieldType, msMessageID, Dumpster);
                if (ritRootReturnList.Count != 0)
                {
                    foreach (ItemType iitem in ritRootReturnList)
                    {
                        rtReturnList.Add(iitem);
                    }
                }

                FindFolderType fiFindFolder = new FindFolderType();
                fiFindFolder.Traversal = FolderQueryTraversalType.Deep;
                FolderResponseShapeType rsResponseShape = new FolderResponseShapeType();
                rsResponseShape.BaseShape = DefaultShapeNamesType.IdOnly;
                fiFindFolder.FolderShape = rsResponseShape;
                fiFindFolder.ParentFolderIds = biArray;
                FindFolderResponseType findFolderResponse = esb.FindFolder(fiFindFolder);
                ResponseMessageType[] rmta = findFolderResponse.ResponseMessages.Items;
                foreach (ResponseMessageType rmt in rmta)
                {
                    if (((FindFolderResponseMessageType)rmt).ResponseClass == ResponseClassType.Success)
                    {
                        FindFolderResponseMessageType ffResponse = (FindFolderResponseMessageType)rmt;
                        if (ffResponse.RootFolder.TotalItemsInView > 0)
                        {
                            foreach (BaseFolderType suSubFolder in ffResponse.RootFolder.Folders)
                            {
                                BaseFolderIdType[] bfarray = new BaseFolderIdType[1];
                                bfarray[0] = suSubFolder.FolderId;
                                List<ItemType> ritReturnList = FindMessage(bfarray, uifieldType, msMessageID, Dumpster);
                                if (ritReturnList.Count != 0)
                                {
                                    foreach (ItemType iitem in ritReturnList)
                                    {
                                        rtReturnList.Add(iitem);
                                    }
                                }
                            }

                        }
                        else
                        { //handle no folder
                        }
                    }
                    else
                    { //handle error
                    }

                }
                return rtReturnList;

            }
            catch (Exception exException)
            {
                Console.WriteLine(exException.ToString());
                throw new FindFolderException("Find Folder Exception");
            }
        }
        public List<BaseFolderType> GetFolder(BaseFolderIdType[] biArray)
        {
            List<BaseFolderType> rtReturnList = new List<BaseFolderType>();
            GetFolderType gfGetfolder = new GetFolderType();
            FolderResponseShapeType fsFoldershape = new FolderResponseShapeType();
            fsFoldershape.BaseShape = DefaultShapeNamesType.AllProperties;
            DistinguishedFolderIdType rfRootFolder = new DistinguishedFolderIdType();
            gfGetfolder.FolderIds = biArray;
            gfGetfolder.FolderShape = fsFoldershape;
            GetFolderResponseType fldResponse = esb.GetFolder(gfGetfolder);
            ArrayOfResponseMessagesType aormt = fldResponse.ResponseMessages;
            ResponseMessageType[] rmta = aormt.Items;
            foreach (ResponseMessageType rmt in rmta)
            {
                if (rmt.ResponseClass == ResponseClassType.Success)
                {
                    FolderInfoResponseMessageType firmt;
                    firmt = (rmt as FolderInfoResponseMessageType);
                    BaseFolderType[] rfFolders = firmt.Folders;
                    rtReturnList.Add(rfFolders[0]);
                }
                else { 
                //Deal with Error
                }

            }
            return rtReturnList;
        }
        public List<BaseFolderType> FindFolder(BaseFolderIdType[] biArray, String fnFolderName)
        {
            List<BaseFolderType> rtReturnList = new List<BaseFolderType>();
            FindFolderType fiFindFolder = new FindFolderType();
            fiFindFolder.Traversal = FolderQueryTraversalType.Shallow;
            FolderResponseShapeType rsResponseShape = new FolderResponseShapeType();
            rsResponseShape.BaseShape = DefaultShapeNamesType.AllProperties;
            fiFindFolder.FolderShape = rsResponseShape;
            fiFindFolder.ParentFolderIds = biArray;
            //Add Restriction for DisplayName

            RestrictionType ffRestriction = new RestrictionType();
            IsEqualToType ieToType = new IsEqualToType();
            PathToUnindexedFieldType fnFolderNameField = new PathToUnindexedFieldType();
            fnFolderNameField.FieldURI = UnindexedFieldURIType.folderDisplayName;

            FieldURIOrConstantType ciConstantType = new FieldURIOrConstantType();
            ConstantValueType cvConstantValueType = new ConstantValueType();
            cvConstantValueType.Value = fnFolderName;
            ciConstantType.Item = cvConstantValueType;
            ieToType.Item = fnFolderNameField;
            ieToType.FieldURIOrConstant = ciConstantType;
            ffRestriction.Item = ieToType;
            fiFindFolder.Restriction = ffRestriction;

            FindFolderResponseType findFolderResponse = esb.FindFolder(fiFindFolder);
            ResponseMessageType[] rmta = findFolderResponse.ResponseMessages.Items;
            foreach (ResponseMessageType rmt in rmta)
            {
                if (((FindFolderResponseMessageType)rmt).ResponseClass == ResponseClassType.Success)
                {
                    FindFolderResponseMessageType ffResponse = (FindFolderResponseMessageType)rmt;
                    if (ffResponse.RootFolder.TotalItemsInView > 0)
                    {
                        foreach (BaseFolderType suSubFolder in ffResponse.RootFolder.Folders)
                        {
                            rtReturnList.Add(suSubFolder);
                        }

                    }
                    else
                    { //handle no folder
                    }
                }
                else
                { //handle error
                }

            }
            return rtReturnList;
        }
        public String convertid(ItemIdType iiItemId, IdFormatType dfDestinationformat)
        {
            String riReturnID = "";
            ConvertIdType ciConvertIDRequest = new ConvertIdType();
            ciConvertIDRequest.SourceIds = new AlternateIdType[1];
            ciConvertIDRequest.SourceIds[0] = new AlternateIdType();
            ciConvertIDRequest.SourceIds[0].Format = IdFormatType.EwsId;
            (ciConvertIDRequest.SourceIds[0] as AlternateIdType).Id = iiItemId.Id;
            (ciConvertIDRequest.SourceIds[0] as AlternateIdType).Mailbox = emEmailAddress;
            ciConvertIDRequest.DestinationFormat = dfDestinationformat;

            try
            {
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
            }
            catch (Exception e)
            {

                Console.WriteLine(e.Message);
            }
            return riReturnID;
        }
        
       }
    public class WriteFeed 
    {
        public WriteFeed(EWSConnection ewc,String fnFileName, String fnFeedTitle, String snServerName, List<ItemType> miMailboxItems)
        {
            try
            {
                ExchangeServiceBinding esb = ewc.esb;
                String emMailboxEmailAddress = ewc.emEmailAddress;
                String StartTime = "";
                String EndTime = "";
                String Location = "";
                String gmtStartTime = "";
                XmlWriter xrXmlWritter = new XmlTextWriter(fnFileName, null);
                xrXmlWritter.WriteStartDocument();
                xrXmlWritter.WriteStartElement("rss");
                xrXmlWritter.WriteAttributeString("version", "2.0");
                xrXmlWritter.WriteStartElement("channel");
                xrXmlWritter.WriteElementString("title", fnFeedTitle);
                xrXmlWritter.WriteElementString("link", "https://" + snServerName + "/owa/");
                xrXmlWritter.WriteElementString("description", "Exchange Feed For " + fnFeedTitle);

                foreach (ItemType itItemType in miMailboxItems)
                {
                    xrXmlWritter.WriteStartElement("item");
                    String sbSubject = "";
                    if (itItemType.Subject != null) { sbSubject = itItemType.Subject.ToString(); }
                    xrXmlWritter.WriteElementString("title", sbSubject);
                    String openType = "ae=Item";
                    if (itItemType is CalendarItemType) { openType = "ae=PreFormAction&a=Open&"; }
                    xrXmlWritter.WriteElementString("link", "https://" + snServerName + "/owa/?" + openType + "&t=" + itItemType.ItemClass.ToString() + "&id=" + convertid(esb, emMailboxEmailAddress, itItemType.ItemId));
                    if (itItemType is CalendarItemType)
                    {
                        CalendarItemType ciCalenderItem = (CalendarItemType)itItemType;
                        if (ciCalenderItem.Organizer.Item.Name != null)
                        {
                            xrXmlWritter.WriteElementString("author", ciCalenderItem.Organizer.Item.Name.ToString());
                        }
                        else
                        {
                            xrXmlWritter.WriteElementString("author", itItemType.LastModifiedName);
                        }
                        if (ciCalenderItem.Start != null)
                        {
                            gmtStartTime = ciCalenderItem.Start.ToString("r");
                            StartTime = ciCalenderItem.Start.ToLocalTime().ToString("dd/MM/yyyy hh:mm:ss tt");
                        }
                        if (ciCalenderItem.End != null)
                        {
                            EndTime = ciCalenderItem.End.ToLocalTime().ToString("dd/MM/yyyy hh:mm:ss tt");
                        }
                        if (ciCalenderItem.Location != null)
                        {
                            Location = ciCalenderItem.Location.ToString();
                        }
                        xrXmlWritter.WriteStartElement("description");
                        xrXmlWritter.WriteCData("Start : " + StartTime + "<br>" + "End : " + EndTime + "<br>" + "Location : " + Location + "<br><br>");
                    }
                    else if (itItemType is MessageType)
                    {

                        if (((MessageType)itItemType).From.Item.EmailAddress != null)
                        {
                            xrXmlWritter.WriteElementString("author", ((MessageType)itItemType).From.Item.EmailAddress.ToString());
                            xrXmlWritter.WriteStartElement("description");
                        }
                        else
                        {
                            if (((MessageType)itItemType).From.Item.Name != null)
                            {
                                xrXmlWritter.WriteElementString("author", ((MessageType)itItemType).From.Item.Name.ToString());
                            }
                            else {
                                xrXmlWritter.WriteElementString("author", "");
                            }                     
                           
                            xrXmlWritter.WriteStartElement("description");
                        }

                    }
                    else
                    {
                        xrXmlWritter.WriteElementString("author", itItemType.LastModifiedName);
                        xrXmlWritter.WriteStartElement("description");
                    }

                    GetItemType giRequest = new GetItemType();
                    ItemIdType iiItemId = new ItemIdType();
                    iiItemId.Id = itItemType.ItemId.Id;
                    iiItemId.ChangeKey = itItemType.ItemId.ChangeKey;
                    giRequest.ItemIds = new ItemIdType[1];
                    giRequest.ItemIds[0] = iiItemId;
                    giRequest.ItemShape = new ItemResponseShapeType();
                    giRequest.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;
                    giRequest.ItemShape.BodyTypeSpecified = true;
                    giRequest.ItemShape.BodyType = BodyTypeResponseType.HTML;
                    giRequest.ItemShape.IncludeMimeContent = true;
                    GetItemResponseType giResponse = esb.GetItem(giRequest);
                    if (giResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Error)
                    {
                        Console.WriteLine("Error Occured");
                        Console.WriteLine(giResponse.ResponseMessages.Items[0].MessageText);
                    }
                    else
                    {
                        ItemInfoResponseMessageType rmResponseMessage = giResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
                        if (rmResponseMessage.Items.Items[0].Body != null)
                        {
                            if (rmResponseMessage.Items.Items[0].Body.Value != null)
                            {
                                String bsBodyString = rmResponseMessage.Items.Items[0].Body.Value.ToString();
                                if (bsBodyString.IndexOf("<body>") >= 0)
                                {
                                    xrXmlWritter.WriteCData("<html>" + bsBodyString.Substring(bsBodyString.IndexOf("<body>")));
                                }
                            }
                        }
                    }
                    xrXmlWritter.WriteEndElement();

                    if (itItemType is CalendarItemType)
                    {
                        xrXmlWritter.WriteElementString("category", Location);
                        xrXmlWritter.WriteElementString("pubDate", gmtStartTime);
                    }
                    else
                    {
                        xrXmlWritter.WriteElementString("pubDate", itItemType.DateTimeCreated.ToString("r"));
                    }
                    xrXmlWritter.WriteElementString("guid", itItemType.ItemId.Id.ToString());
                    xrXmlWritter.WriteEndElement();


                }
                xrXmlWritter.WriteEndElement();
                xrXmlWritter.WriteEndElement();
                xrXmlWritter.WriteEndDocument();
                xrXmlWritter.Flush();
                xrXmlWritter.Close();
                xrXmlWritter = null;
            }
            catch (Exception rsswriteException)
            {
                Console.WriteLine(rsswriteException.ToString());
    
            }
    
        
        }
        public String convertid(ExchangeServiceBinding esb, String emMailboxEmailAddress, ItemIdType iiItemId)
        {
            String riReturnID = "";
            ConvertIdType ciConvertIDRequest = new ConvertIdType();
            ciConvertIDRequest.SourceIds = new AlternateIdType[1];
            ciConvertIDRequest.SourceIds[0] = new AlternateIdType();
            ciConvertIDRequest.SourceIds[0].Format = IdFormatType.EwsId;
            (ciConvertIDRequest.SourceIds[0] as AlternateIdType).Id = iiItemId.Id;
            (ciConvertIDRequest.SourceIds[0] as AlternateIdType).Mailbox = emMailboxEmailAddress;
            ciConvertIDRequest.DestinationFormat = IdFormatType.OwaId;

            try
            {
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
            }
            catch (Exception e)
            {

                Console.WriteLine(e.Message);
            }
            return riReturnID;
        }
    
    }
    public class List<T> : CollectionBase
    {
        public List() { }

        public T this[int index]
        {
            get { return (T)List[index]; }
            set { List[index] = value; }
        }

        public int Add(T value)
        {
            return List.Add(value);
        }
    }
    public class Contact {
        private String intFirstName = "";
        public String FirstName {
            get
            {
                return this.intFirstName;
            }
            set {
                intFirstName = value;        
            }        
        }
        private String intLastName;
        public String LastName
        {
            get
            {
                return this.intLastName;
            }
            set
            {
                intLastName = value;
            }
        }
        private String intMiddleName;
        public String MiddleName
        {
            get
            {
                return this.intMiddleName;
            }
            set
            {
                intMiddleName = value;
            }
        }
        private String intNickname;
        public String Nickname
        {
            get
            {
                return this.intNickname;
            }
            set
            {
                intNickname = value;
            }
        }
        private String intEmailAddress;
        public String EmailAddress
        {
            get
            {
                return this.intEmailAddress;
            }
            set
            {
                intEmailAddress = value;
            }
        }
        private String intEmailAddressDisplayName;
        public String EmailAddressDisplayName
        {
            get
            {
                return this.intEmailAddressDisplayName;
            }
            set
            {
                intEmailAddressDisplayName = value;
            }
        }
        private String intHomeStreet;
        public String HomeStreet
        {
            get
            {
                return this.intHomeStreet;
            }
            set
            {
                intHomeStreet = value;
            }
        }
        private String intHomeCity;
        public String HomeCity
        {
            get
            {
                return this.intHomeCity;
            }
            set
            {
                intHomeCity = value;
            }
        }
        private String intHomePostalCode;
        public String HomePostalCode
        {
            get
            {
                return this.intHomePostalCode;
            }
            set
            {
                intHomePostalCode = value;
            }
        }
        private String intHomeState;
        public String HomeState
        {
            get
            {
                return this.intHomeState;
            }
            set
            {
                intHomeState = value;
            }
        }
        private String intHomeCountry;
        public String HomeCountry
        {
            get
            {
                return this.intHomeCountry;
            }
            set
            {
                intHomeCountry = value;
            }
        }
        private String intHomePhone;
        public String HomePhone
        {
            get
            {
                return this.intHomePhone;
            }
            set
            {
                intHomePhone = value;
            }
        }
        private String intHomeFax;
        public String HomeFax
        {
            get
            {
                return this.intHomeFax;
            }
            set
            {
                intHomeFax = value;
            }
        }
        private String intMobilePhone;
        public String MobilePhone
        {
            get
            {
                return this.intMobilePhone;
            }
            set
            {
                intMobilePhone = value;
            }
        }
        private String intPersonalWebPage;
        public String PersonalWebPage
        {
            get
            {
                return this.intPersonalWebPage;
            }
            set
            {
                intPersonalWebPage = value;
            }
        }
        private String intBusinessStreet;
        public String BusinessStreet
        {
            get
            {
                return this.intBusinessStreet;
            }
            set
            {
                intBusinessStreet = value;
            }
        }
        private String intBusinessCity;
        public String BusinessCity
        {
            get
            {
                return this.intBusinessCity;
            }
            set
            {
                intBusinessCity = value;
            }
        }
        private String intBusinessPostalCode;
        public String BusinessPostalCode
        {
            get
            {
                return this.intBusinessPostalCode;
            }
            set
            {
                intBusinessPostalCode = value;
            }
        }
        private String intBusinessState;
        public String BusinessState
        {
            get
            {
                return this.intBusinessState;
            }
            set
            {
                intBusinessState = value;
            }
        }
        private String intBusinessCountry;
        public String BusinessCountry
        {
            get
            {
                return this.intBusinessCountry;
            }
            set
            {
                intBusinessCountry = value;
            }
        }
        private String intBusinessPhone;
        public String BusinessPhone
        {
            get
            {
                return this.intBusinessPhone;
            }
            set
            {
                intBusinessPhone = value;
            }
        }
        private String intBusinessFax;
        public String BusinessFax
        {
            get
            {
                return this.intBusinessFax;
            }
            set
            {
                intBusinessFax = value;
            }
        }
        private String intBusinessWebPage;
        public String BusinessWebPage
        {
            get
            {
                return this.intBusinessWebPage;
            }
            set
            {
                intBusinessWebPage = value;
            }
        }
        private String intCompany;
        public String Company
        {
            get
            {
                return this.intCompany;
            }
            set
            {
                intCompany = value;
            }
        }
        private String intJobTitle;
        public String JobTitle
        {
            get
            {
                return this.intJobTitle;
            }
            set
            {
                intJobTitle = value;
            }
        }
        private String intDepartment;
        public String Department
        {
            get
            {
                return this.intDepartment;
            }
            set
            {
                intDepartment = value;
            }
        }
        private String intOfficeLocation;
        public String OfficeLocation
        {
            get
            {
                return this.intOfficeLocation;
            }
            set
            {
                intOfficeLocation = value;
            }
        }
        private String intNotes;
        public String Notes
        {
            get
            {
                return this.intNotes;
            }
            set
            {
                intNotes = value;
            }
        }
        private String intFileAs;
        public String FileAs
        {
            get
            {
                return this.intFileAs;
            }
            set
            {
                intFileAs = value;
            }
        }
        private String intAssistantName;
        public String AssistantName
        {
            get
            {
                return this.intAssistantName;
            }
            set
            {
                intAssistantName = value;
            }
        }
        private String intDisplayName;
        public String DisplayName
        {
            get
            {
                return this.intDisplayName;
            }
            set
            {
                intDisplayName = value;
            }
        }
        private String intFullName;
        public String FullName
        {
            get
            {
                return this.intFullName;
            }
            set
            {
                intFullName = value;
            }
        }
        private String intSuffix;
        public String Suffix
        {
            get
            {
                return this.intSuffix;
            }
            set
            {
                intSuffix = value;
            }
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
    class UpdateFolderException : Exception
    {
        public UpdateFolderException(string UpdateFolderError)
        {
            Console.WriteLine(UpdateFolderError);
        }
    }
    class GetFolderException : Exception
    {
        public GetFolderException(string GetFolderError)
        {
            Console.WriteLine(GetFolderError);
        }
    }
    class OofSettingException : Exception
    {
        public OofSettingException(string OofSettingError)
        {
            Console.WriteLine(OofSettingError);
        }
    }
    class FindItemException : Exception
      {
       public FindItemException(string fiFindItemError)
        {
            Console.WriteLine(fiFindItemError);
        }
    }
    class FindFolderException : Exception
    {
        public FindFolderException(string fiFindFolderError)
        {
            Console.WriteLine(fiFindFolderError);
        }
    }
    class  rsswriteException : Exception
      {
           public rsswriteException(string rsRssError)
        {
            Console.WriteLine(rsRssError);
        }
    }
    class DownloadAttachmentException : Exception
    {
        public DownloadAttachmentException(string dlDownloadAttachment)
        {
            Console.WriteLine(dlDownloadAttachment);
        }
    }
   
}

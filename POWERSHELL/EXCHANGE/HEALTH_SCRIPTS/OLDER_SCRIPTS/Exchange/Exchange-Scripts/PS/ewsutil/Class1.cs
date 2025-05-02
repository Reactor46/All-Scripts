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
using Microsoft.Win32;
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
        private struct SYSTEMTIME
        {
            public Int16 wYear;
            public Int16 wMonth;
            public Int16 wDayOfWeek;
            public Int16 wDay;
            public Int16 wHour;
            public Int16 wMinute;
            public Int16 wSecond;
            public Int16 wMilliseconds;
            public void getSysTime(byte[] Tzival, int offset)
            {
                wYear = BitConverter.ToInt16(Tzival, offset);
                wMonth = BitConverter.ToInt16(Tzival, offset + 2);
                wDayOfWeek = BitConverter.ToInt16(Tzival, offset + 4);
                wDay = BitConverter.ToInt16(Tzival, offset + 6);
                wHour = BitConverter.ToInt16(Tzival, offset + 8);
                wMinute = BitConverter.ToInt16(Tzival, offset + 10);
                wSecond = BitConverter.ToInt16(Tzival, offset + 12);
                wMilliseconds = BitConverter.ToInt16(Tzival, offset + 14);
            }
        }
        private struct REG_TZI_FORMAT
        {
            public Int32 Bias;
            public Int32 StandardBias;
            public Int32 DaylightBias;
            public SYSTEMTIME StandardDate;
            public SYSTEMTIME DaylightDate;
            public void regget(byte[] Tzival)
            {
                Bias = BitConverter.ToInt32(Tzival, 0);
                StandardBias = BitConverter.ToInt32(Tzival, 4);
                DaylightBias = BitConverter.ToInt32(Tzival, 8);
                StandardDate = new SYSTEMTIME();
                StandardDate.getSysTime(Tzival, 12);
                DaylightDate = new SYSTEMTIME();
                DaylightDate.getSysTime(Tzival, 28);
            }

        }
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
        public ExchangeServiceBinding esb
        {
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
                        fiFile.Write(fiFileAttachmentType.Content, 0, fiFileAttachmentType.Content.Length);
                        fiFile.Close();

                    }
                }
                else
                {
                    throw new DownloadAttachmentException("Error Downloding Attachment");

                }



            }
        }
        public List<ItemType> GetNDRs(BaseFolderIdType[] biArray, Duration duDuration)
        {

            List<ItemType> intFiFolderItems = new List<ItemType>();
            try
            {

                BasePathToElementType[] beAdditionproperteis = new BasePathToElementType[5];

                PathToExtendedFieldType osOriginalSenderEmail = new PathToExtendedFieldType();
                osOriginalSenderEmail.PropertyTag = "0x0067";
                osOriginalSenderEmail.PropertyType = MapiPropertyTypeType.String;

                PathToExtendedFieldType osOriginalSenderAdrType = new PathToExtendedFieldType();
                osOriginalSenderAdrType.PropertyTag = "0x0066";
                osOriginalSenderAdrType.PropertyType = MapiPropertyTypeType.String;

                PathToExtendedFieldType osOriginalSenderName = new PathToExtendedFieldType();
                osOriginalSenderName.PropertyTag = "0x005A";
                osOriginalSenderName.PropertyType = MapiPropertyTypeType.String;

                PathToExtendedFieldType osOriginalSubject = new PathToExtendedFieldType();
                osOriginalSubject.PropertyTag = "0x0049";
                osOriginalSubject.PropertyType = MapiPropertyTypeType.String;

                PathToExtendedFieldType osOriginalRecp = new PathToExtendedFieldType();
                osOriginalRecp.PropertyTag = "0x0074";
                osOriginalRecp.PropertyType = MapiPropertyTypeType.String;



                beAdditionproperteis[0] = osOriginalSenderAdrType;
                beAdditionproperteis[1] = osOriginalSenderEmail;
                beAdditionproperteis[2] = osOriginalSenderName;
                beAdditionproperteis[3] = osOriginalRecp;
                beAdditionproperteis[4] = osOriginalSubject;


                intFiFolderItems = FindItems(biArray, duDuration, beAdditionproperteis, "REPORT.IPM.Note.NDR");
            }
            catch (Exception exException)
            {
                Console.WriteLine(exException.ToString());
                // return exException.ToString();
            }
            return intFiFolderItems;

        }
        public List<ItemType> FindItems(BaseFolderIdType[] biArray, Duration duDuration, BasePathToElementType[] exExtenededProperties, String mtMessageType)
        {
            List<ItemType> intFiFolderItems = new List<ItemType>();
            try
            {
                FindItemType fiFindItemRequest = new FindItemType();
                fiFindItemRequest.ParentFolderIds = biArray;
                fiFindItemRequest.Traversal = ItemQueryTraversalType.Shallow;
                ItemResponseShapeType ipItemProperties = new ItemResponseShapeType();
                ipItemProperties.BaseShape = DefaultShapeNamesType.AllProperties;
                ipItemProperties.AdditionalProperties = exExtenededProperties;
                fiFindItemRequest.ItemShape = ipItemProperties;
                RestrictionType ffRestriction = new RestrictionType();
                if (duDuration != null)
                {
                    AndType raRestictionAnd = new AndType();
                    IsGreaterThanOrEqualToType igteToType = new IsGreaterThanOrEqualToType();
                    PathToUnindexedFieldType rfReceivedTime = new PathToUnindexedFieldType();
                    rfReceivedTime.FieldURI = UnindexedFieldURIType.itemDateTimeReceived;
                    igteToType.Item = rfReceivedTime;
                    igteToType.FieldURIOrConstant = new FieldURIOrConstantType();
                    igteToType.FieldURIOrConstant.Item = new ConstantValueType();
                    (igteToType.FieldURIOrConstant.Item as ConstantValueType).Value = duDuration.StartTime.ToUniversalTime().ToString("yyyy-MM-ddThh:mm:ssZ");

                    IsLessThanOrEqualToType ilteToType = new IsLessThanOrEqualToType();
                    ilteToType.Item = rfReceivedTime;
                    ilteToType.FieldURIOrConstant = new FieldURIOrConstantType();
                    ilteToType.FieldURIOrConstant.Item = new ConstantValueType();
                    (ilteToType.FieldURIOrConstant.Item as ConstantValueType).Value = duDuration.EndTime.ToUniversalTime().ToString("yyyy-MM-ddThh:mm:ssZ");

                    if (mtMessageType == "")
                    {
                        raRestictionAnd.Items = new SearchExpressionType[2];
                    }
                    else
                    {
                        raRestictionAnd.Items = new SearchExpressionType[3];
                        IsEqualToType ieToTypeClass = new IsEqualToType();
                        PathToUnindexedFieldType itItemType = new PathToUnindexedFieldType();
                        itItemType.FieldURI = UnindexedFieldURIType.itemItemClass;
                        ieToTypeClass.Item = itItemType;
                        FieldURIOrConstantType constantType = new FieldURIOrConstantType();
                        ConstantValueType constantValueType = new ConstantValueType();
                        constantValueType.Value = mtMessageType;
                        constantType.Item = constantValueType;
                        ieToTypeClass.Item = itItemType;
                        ieToTypeClass.FieldURIOrConstant = constantType;
                        raRestictionAnd.Items[2] = ieToTypeClass;
                    }

                    raRestictionAnd.Items[0] = igteToType;
                    raRestictionAnd.Items[1] = ilteToType;
                    ffRestriction.Item = raRestictionAnd;
                    fiFindItemRequest.Restriction = ffRestriction;
                }
                FindItemResponseType frFindItemResponse = esb.FindItem(fiFindItemRequest);
                if (frFindItemResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
                {
                    foreach (FindItemResponseMessageType firmtMessage in frFindItemResponse.ResponseMessages.Items)
                    {
                        Console.WriteLine("Number of Items Found : " + firmtMessage.RootFolder.TotalItemsInView);
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
            return intFiFolderItems;

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
                if (frFindItemResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
                {
                    foreach (FindItemResponseMessageType firmtMessage in frFindItemResponse.ResponseMessages.Items)
                    {
                        Console.WriteLine("Number of Items Found : " + firmtMessage.RootFolder.TotalItemsInView);
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
                                    giRequest.ItemShape.BodyType = BodyTypeResponseType.Text;
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
                }
                else
                {
                    Console.WriteLine("Error Occured");
                    Console.WriteLine(frFindItemResponse.ResponseMessages.Items[0].MessageText.ToString());
                }
                return rtReturnList;
            }
            catch (Exception exException)
            {
                Console.WriteLine(exException.ToString());
                throw new FindFolderException("Find Folder Exception");
            }
        }
        public List<ItemType> FindUnread(BaseFolderIdType[] biArray, Duration duDuration, BasePathToElementType[] exExtenededProperties, String mtMessageType)
        {
            List<ItemType> intFiFolderItems = new List<ItemType>();
            try
            {
                FindItemType fiFindItemRequest = new FindItemType();
                fiFindItemRequest.ParentFolderIds = biArray;
                fiFindItemRequest.Traversal = ItemQueryTraversalType.Shallow;
                ItemResponseShapeType ipItemProperties = new ItemResponseShapeType();
                ipItemProperties.BaseShape = DefaultShapeNamesType.AllProperties;
                ipItemProperties.AdditionalProperties = exExtenededProperties;
                fiFindItemRequest.ItemShape = ipItemProperties;
                RestrictionType ffRestriction = new RestrictionType();
                if (duDuration != null)
                {
                    AndType raRestictionAnd = new AndType();
                    IsGreaterThanOrEqualToType igteToType = new IsGreaterThanOrEqualToType();
                    PathToUnindexedFieldType rfReceivedTime = new PathToUnindexedFieldType();
                    rfReceivedTime.FieldURI = UnindexedFieldURIType.itemDateTimeReceived;
                    igteToType.Item = rfReceivedTime;
                    igteToType.FieldURIOrConstant = new FieldURIOrConstantType();
                    igteToType.FieldURIOrConstant.Item = new ConstantValueType();
                    (igteToType.FieldURIOrConstant.Item as ConstantValueType).Value = duDuration.StartTime.ToUniversalTime().ToString("yyyy-MM-ddThh:mm:ssZ");

                    IsLessThanOrEqualToType ilteToType = new IsLessThanOrEqualToType();
                    ilteToType.Item = rfReceivedTime;
                    ilteToType.FieldURIOrConstant = new FieldURIOrConstantType();
                    ilteToType.FieldURIOrConstant.Item = new ConstantValueType();
                    (ilteToType.FieldURIOrConstant.Item as ConstantValueType).Value = duDuration.EndTime.ToUniversalTime().ToString("yyyy-MM-ddThh:mm:ssZ");

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

                    if (mtMessageType == "")
                    {
                        raRestictionAnd.Items = new SearchExpressionType[3];
                    }
                    else
                    {
                        raRestictionAnd.Items = new SearchExpressionType[4];
                        IsEqualToType ieToTypeClass = new IsEqualToType();
                        PathToUnindexedFieldType itItemType = new PathToUnindexedFieldType();
                        itItemType.FieldURI = UnindexedFieldURIType.itemItemClass;
                        ieToTypeClass.Item = itItemType;
                        FieldURIOrConstantType reConstantType = new FieldURIOrConstantType();
                        ConstantValueType reConstantValueType = new ConstantValueType();
                        reConstantValueType.Value = mtMessageType;
                        reConstantType.Item = reConstantValueType;
                        ieToTypeClass.Item = itItemType;
                        ieToTypeClass.FieldURIOrConstant = reConstantType;
                        raRestictionAnd.Items[3] = ieToTypeClass;
                    }

                    raRestictionAnd.Items[0] = igteToType;
                    raRestictionAnd.Items[1] = ilteToType;
                    raRestictionAnd.Items[2] = ieToTypeRead;
                    ffRestriction.Item = raRestictionAnd;
                    fiFindItemRequest.Restriction = ffRestriction;
                }
                FindItemResponseType frFindItemResponse = esb.FindItem(fiFindItemRequest);
                if (frFindItemResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
                {
                    foreach (FindItemResponseMessageType firmtMessage in frFindItemResponse.ResponseMessages.Items)
                    {
                        Console.WriteLine("Number of Items Found : " + firmtMessage.RootFolder.TotalItemsInView);
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
            return intFiFolderItems;

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
                else
                {
                    //Deal with Error
                }

            }
            return rtReturnList;
        }
        public List<BaseFolderType> GetFolder(BaseFolderIdType[] biArray,BasePathToElementType[] adAdditionalProps)
        {
            List<BaseFolderType> rtReturnList = new List<BaseFolderType>();
            GetFolderType gfGetfolder = new GetFolderType();
            FolderResponseShapeType fsFoldershape = new FolderResponseShapeType();
            fsFoldershape.BaseShape = DefaultShapeNamesType.AllProperties;
            fsFoldershape.AdditionalProperties = adAdditionalProps; 
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
                else
                {
                    //Deal with Error
                }

            }
            return rtReturnList;
        }
        public List<ItemType> GetItems(BaseItemIdType[] biArray,BasePathToElementType[] adAdditionalProps){
            List<ItemType>rtReturnList = new List<ItemType>();
            GetItemType giRequest = new GetItemType();
            giRequest.ItemIds = biArray;
            giRequest.ItemShape = new ItemResponseShapeType();
            giRequest.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;
            if (adAdditionalProps.Length != 0){
                giRequest.ItemShape.AdditionalProperties = adAdditionalProps;
            }
            giRequest.ItemShape.BodyTypeSpecified = true;
            giRequest.ItemShape.BodyType = BodyTypeResponseType.Text;
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
                foreach (ItemType miMailboxItem in rmResponseMessage.Items.Items)
                {
                    rtReturnList.Add(miMailboxItem);
                }

            }
            return rtReturnList;
    
        }
        public String SendMessage(EmailAddressType[] ToAddress, String sbSubject, String bdBody)
        {
            String MessageStatus = "";
            try
            {
                MessageType msMessage = new MessageType();
                msMessage.ToRecipients = ToAddress;
                msMessage.Subject = sbSubject;
                msMessage.Sensitivity = SensitivityChoicesType.Normal;
                msMessage.Body = new BodyType();
                msMessage.Body.BodyType1 = BodyTypeType.HTML;
                msMessage.Body.Value = bdBody;

                CreateItemType ciCreateItem = new CreateItemType();
                ciCreateItem.MessageDisposition = MessageDispositionType.SendAndSaveCopy;
                ciCreateItem.MessageDispositionSpecified = true;

                DistinguishedFolderIdType siFolder = new DistinguishedFolderIdType();
                siFolder.Id = DistinguishedFolderIdNameType.sentitems;

                TargetFolderIdType tfTargetFolder = new TargetFolderIdType();
                tfTargetFolder.Item = siFolder;
                ciCreateItem.SavedItemFolderId = tfTargetFolder;

                ciCreateItem.Items = new NonEmptyArrayOfAllItemsType();
                ciCreateItem.Items.Items = new ItemType[1] { msMessage };
                CreateItemResponseType rsResponse = esb.CreateItem(ciCreateItem);
                ResponseMessageType[] msSentResponse = rsResponse.ResponseMessages.Items;
                if (msSentResponse[0].ResponseClass == ResponseClassType.Success)
                {
                    MessageStatus = "Message Sent";
                }
                else
                {
                    Console.WriteLine("Error Occured");
                    Console.WriteLine(msSentResponse[0].MessageText.ToString());
                }
            }
            catch (Exception exException)
            {
                Console.WriteLine(exException.ToString());
                // return exException.ToString();
            }
            return MessageStatus;

        }
        public String DeleteItems(BaseItemIdType[] idItemstoDelete, DisposalType diType)
        {
            String rtReturnString = "";
            DeleteItemType diDeleteItemRequest = new DeleteItemType();
            diDeleteItemRequest.ItemIds = idItemstoDelete;
            diDeleteItemRequest.DeleteType = diType;
            diDeleteItemRequest.SendMeetingCancellationsSpecified = true;
            diDeleteItemRequest.SendMeetingCancellations = CalendarItemCreateOrDeleteOperationType.SendToNone;
            DeleteItemResponseType diResponse = esb.DeleteItem(diDeleteItemRequest);
            ResponseMessageType[] diResponseMessages = diResponse.ResponseMessages.Items;
            int DeleteSuccess = 0;
            int DeleteError = 0;
            foreach (ResponseMessageType repMessage in diResponseMessages)
            {
                if (repMessage.ResponseClass == ResponseClassType.Success)
                {
                    DeleteSuccess++;
                }
                else
                {
                    DeleteError++;
                    Console.WriteLine("Error Occured");
                    Console.WriteLine(repMessage.MessageText.ToString());
                }
            }
            rtReturnString = DeleteSuccess.ToString() + " Messages Deleted " + DeleteError + " Errors";
            return rtReturnString;
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
        public List<BaseFolderType> GetAllMailboxFolders(BaseFolderIdType[] biArray)
        {
            List<BaseFolderType> rtReturnList = new List<BaseFolderType>();
            FindFolderType fiFindFolder = new FindFolderType();
            fiFindFolder.Traversal = FolderQueryTraversalType.Deep;
            FolderResponseShapeType rsResponseShape = new FolderResponseShapeType();
            rsResponseShape.BaseShape = DefaultShapeNamesType.Default;
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
        public String convertHexid(String iiItemId, IdFormatType dfDestinationformat)
        {
            
            String riReturnID = "";
            ConvertIdType ciConvertIDRequest = new ConvertIdType();
            ciConvertIDRequest.SourceIds = new AlternateIdType[1];
            ciConvertIDRequest.SourceIds[0] = new AlternateIdType();
            ciConvertIDRequest.SourceIds[0].Format = IdFormatType.HexEntryId;
            (ciConvertIDRequest.SourceIds[0] as AlternateIdType).Id = iiItemId;
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
        public String convertHexidPublicFolder(String iiItemId, IdFormatType dfDestinationformat)
        {

            String riReturnID = "";
            ConvertIdType ciConvertIDRequest = new ConvertIdType();
            ciConvertIDRequest.SourceIds = new AlternatePublicFolderIdType[1];
            ciConvertIDRequest.SourceIds[0] = new AlternatePublicFolderIdType();
            ciConvertIDRequest.SourceIds[0].Format = IdFormatType.HexEntryId;
            (ciConvertIDRequest.SourceIds[0] as AlternatePublicFolderIdType).FolderId = iiItemId;
            (ciConvertIDRequest.SourceIds[0] as AlternatePublicFolderIdType).Format = IdFormatType.HexEntryId;
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
                        AlternatePublicFolderIdType aiID = (cirmt.AlternateId as AlternatePublicFolderIdType);
                        riReturnID = aiID.FolderId.ToString();

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
        public List<PermissionType> PermissionSetToList(PermissionSetType psPermissionSet)
        {
            List<PermissionType> plPermissionList = new List<PermissionType>();
            for (int fbpint = 0; fbpint < psPermissionSet.Permissions.Length; fbpint++)
            {
                if (psPermissionSet.Permissions[fbpint].PermissionLevel == PermissionLevelType.Custom)
                {
                    plPermissionList.Add(psPermissionSet.Permissions[fbpint]);
                }
                else
                {
                    PermissionType fbNewPermsion = new PermissionType();
                    {
                        fbNewPermsion.UserId = psPermissionSet.Permissions[fbpint].UserId;
                        fbNewPermsion.PermissionLevel = psPermissionSet.Permissions[fbpint].PermissionLevel;
                    }
                    plPermissionList.Add(fbNewPermsion);
                }

            }
            return plPermissionList;



        }
        public String UpdateFolderDACL(List<PermissionType> npPermissionsList, BaseFolderIdType FolderToUpdate)
        {
            PermissionSetType fbNewPermsionsSet = new PermissionSetType();
            fbNewPermsionsSet.Permissions = new PermissionType[npPermissionsList.Count];
            Hashtable cdCheckDuplicates = new Hashtable();
            for (int cpint1 = 0; cpint1 < npPermissionsList.Count; cpint1++)
            {
                if (npPermissionsList[cpint1].UserId.DistinguishedUserSpecified == false)
                {
                    if (cdCheckDuplicates.ContainsKey(npPermissionsList[cpint1].UserId.PrimarySmtpAddress.ToLower()))
                    {
                        fbNewPermsionsSet.Permissions[(int)cdCheckDuplicates[npPermissionsList[cpint1].UserId.PrimarySmtpAddress.ToLower()]] = npPermissionsList[cpint1];
                    }
                    else
                    {
                        cdCheckDuplicates.Add(npPermissionsList[cpint1].UserId.PrimarySmtpAddress.ToLower(), cpint1);
                        fbNewPermsionsSet.Permissions[cpint1] = npPermissionsList[cpint1];
                    }
                }
                else
                {
                    switch (npPermissionsList[cpint1].UserId.DistinguishedUser)
                    {
                        case DistinguishedUserType.Default:
                            if (cdCheckDuplicates.ContainsKey("default@default"))
                            {
                                fbNewPermsionsSet.Permissions[(int)cdCheckDuplicates["default@default"]] = npPermissionsList[cpint1];
                            }
                            else
                            {
                                cdCheckDuplicates.Add("default@default", cpint1);
                                fbNewPermsionsSet.Permissions[cpint1] = npPermissionsList[cpint1];
                            }
                            break;
                        case DistinguishedUserType.Anonymous:
                            if (cdCheckDuplicates.ContainsKey("anonymous@default"))
                            {
                                fbNewPermsionsSet.Permissions[(int)cdCheckDuplicates["anonymous@default"]] = npPermissionsList[cpint1];
                            }
                            else
                            {
                                cdCheckDuplicates.Add("anonymous@default", cpint1);
                                fbNewPermsionsSet.Permissions[cpint1] = npPermissionsList[cpint1];
                            }
                            break;
                        default:
                            fbNewPermsionsSet.Permissions[cpint1] = npPermissionsList[cpint1];
                            break;


                    }

                }
            }

            FolderType fbUpdateFolder = new FolderType();
            fbUpdateFolder.PermissionSet = fbNewPermsionsSet;
            UpdateFolderType upUpdateFolderRequest = new UpdateFolderType();
            FolderChangeType fcFolderchanges = new FolderChangeType();
            fcFolderchanges.Item = FolderToUpdate;
            SetFolderFieldType fbFolderPerms = new SetFolderFieldType();
            PathToUnindexedFieldType fbFieldURI = new PathToUnindexedFieldType();
            fbFieldURI.FieldURI = UnindexedFieldURIType.folderPermissionSet;
            fbFolderPerms.Item = fbFieldURI;
            fbFolderPerms.Item1 = fbUpdateFolder;

            fcFolderchanges.Updates = new FolderChangeDescriptionType[1] { fbFolderPerms };
            upUpdateFolderRequest.FolderChanges = new FolderChangeType[1] { fcFolderchanges };
            String upres = "";
            UpdateFolderResponseType ufUpdateFolderResponse = esb.UpdateFolder(upUpdateFolderRequest);
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
        public String UpdateCalendarFolderDACL(List<CalendarPermissionType> npPermissionsList, BaseFolderIdType FolderToUpdate)
        {
            CalendarPermissionSetType fbNewPermsionsSet = new CalendarPermissionSetType();
            fbNewPermsionsSet.CalendarPermissions = new CalendarPermissionType[npPermissionsList.Count];
            Hashtable cdCheckDuplicates = new Hashtable();
            for (int cpint1 = 0; cpint1 < npPermissionsList.Count; cpint1++)
            {
                if (npPermissionsList[cpint1].UserId.DistinguishedUserSpecified == false)
                {
                    if (cdCheckDuplicates.ContainsKey(npPermissionsList[cpint1].UserId.PrimarySmtpAddress.ToLower()))
                    {
                        fbNewPermsionsSet.CalendarPermissions[(int)cdCheckDuplicates[npPermissionsList[cpint1].UserId.PrimarySmtpAddress.ToLower()]] = npPermissionsList[cpint1];
                    }
                    else
                    {
                        cdCheckDuplicates.Add(npPermissionsList[cpint1].UserId.PrimarySmtpAddress.ToLower(), cpint1);
                        fbNewPermsionsSet.CalendarPermissions[cpint1] = npPermissionsList[cpint1];
                    }
                }
                else
                {
                    switch (npPermissionsList[cpint1].UserId.DistinguishedUser)
                    {
                        case DistinguishedUserType.Default:
                            if (cdCheckDuplicates.ContainsKey("default@default"))
                            {
                                fbNewPermsionsSet.CalendarPermissions[(int)cdCheckDuplicates["default@default"]] = npPermissionsList[cpint1];
                            }
                            else
                            {
                                cdCheckDuplicates.Add("default@default", cpint1);
                                fbNewPermsionsSet.CalendarPermissions[cpint1] = npPermissionsList[cpint1];
                            }
                            break;
                        case DistinguishedUserType.Anonymous:
                            if (cdCheckDuplicates.ContainsKey("anonymous@default"))
                            {
                                fbNewPermsionsSet.CalendarPermissions[(int)cdCheckDuplicates["anonymous@default"]] = npPermissionsList[cpint1];
                            }
                            else
                            {
                                cdCheckDuplicates.Add("anonymous@default", cpint1);
                                fbNewPermsionsSet.CalendarPermissions[cpint1] = npPermissionsList[cpint1];
                            }
                            break;
                        default:
                            fbNewPermsionsSet.CalendarPermissions[cpint1] = npPermissionsList[cpint1];
                            break;


                    }

                }
            }

            CalendarFolderType fbUpdateFolder = new CalendarFolderType();
            fbUpdateFolder.PermissionSet = fbNewPermsionsSet;
            UpdateFolderType upUpdateFolderRequest = new UpdateFolderType();
            FolderChangeType fcFolderchanges = new FolderChangeType();
            fcFolderchanges.Item = FolderToUpdate;
            SetFolderFieldType fbFolderPerms = new SetFolderFieldType();
            PathToUnindexedFieldType fbFieldURI = new PathToUnindexedFieldType();
            fbFieldURI.FieldURI = UnindexedFieldURIType.folderPermissionSet;
            fbFolderPerms.Item = fbFieldURI;
            fbFolderPerms.Item1 = fbUpdateFolder;

            fcFolderchanges.Updates = new FolderChangeDescriptionType[1] { fbFolderPerms };
            upUpdateFolderRequest.FolderChanges = new FolderChangeType[1] { fcFolderchanges };
            String upres = "";
            UpdateFolderResponseType ufUpdateFolderResponse = esb.UpdateFolder(upUpdateFolderRequest);
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
        public BaseFolderType FindSubFolder(BaseFolderType pfParentFolder, String sfChildSub)
        {
            BaseFolderType rvFolderID = null;
            FolderType dd = new FolderType();
            BaseFolderIdType bf = new FolderIdType();

            // Create the request and specify the travesal type
            FindFolderType findFolderRequest = new FindFolderType();
            findFolderRequest.Traversal = FolderQueryTraversalType.Shallow;

            // Define the properties returned in the response
            FolderResponseShapeType responseShape = new FolderResponseShapeType();
            responseShape.BaseShape = DefaultShapeNamesType.Default;
            findFolderRequest.FolderShape = responseShape;

            // Identify which folders to search
            FolderIdType[] folderIDArray = new FolderIdType[1];

            folderIDArray[0] = new FolderIdType();
            folderIDArray[0] = pfParentFolder.FolderId;

            // Add the folders to search to the request
            findFolderRequest.ParentFolderIds = folderIDArray;
            // Send the request and get the response
            //Add Restriction for DisplayName
            RestrictionType ffRestriction = new RestrictionType();
            IsEqualToType ieToType = new IsEqualToType();
            PathToUnindexedFieldType snSubjectName = new PathToUnindexedFieldType();
            snSubjectName.FieldURI = UnindexedFieldURIType.folderDisplayName;

            FieldURIOrConstantType ciConstantType = new FieldURIOrConstantType();
            ConstantValueType cvConstantValueType = new ConstantValueType();
            cvConstantValueType.Value = sfChildSub;
            ciConstantType.Item = cvConstantValueType;
            ieToType.Item = snSubjectName;
            ieToType.FieldURIOrConstant = ciConstantType;
            ffRestriction.Item = ieToType;
            findFolderRequest.Restriction = ffRestriction;
            FindFolderResponseType findFolderResponse = esb.FindFolder(findFolderRequest);

            // Get the response messages
            ResponseMessageType[] rmta = findFolderResponse.ResponseMessages.Items;

            foreach (ResponseMessageType rmt in rmta)
            {
                if (((FindFolderResponseMessageType)rmt).ResponseClass == ResponseClassType.Success)
                {
                    FindFolderResponseMessageType ffResponse = (FindFolderResponseMessageType)rmt;
                    if (ffResponse.RootFolder.TotalItemsInView > 0)
                    {
                        foreach (BaseFolderType fld in ffResponse.RootFolder.Folders)
                        {
                            rvFolderID = (BaseFolderType)fld;
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
            return rvFolderID;


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
        public List<DelegateUserResponseMessageType> getDeletgates(String emEmailAdrress)
        {
            List<DelegateUserResponseMessageType> reList = new List<DelegateUserResponseMessageType>();
            GetDelegateType gdGetDelegate = new GetDelegateType();
            gdGetDelegate.IncludePermissions = true;
            gdGetDelegate.Mailbox = new EmailAddressType();
            gdGetDelegate.Mailbox.EmailAddress = emEmailAdrress;
            GetDelegateResponseMessageType gdResponse = esb.GetDelegate(gdGetDelegate);

            if (gdResponse.ResponseClass == ResponseClassType.Success)
            {
                if (gdResponse.ResponseMessages != null)
                {
                    foreach (DelegateUserResponseMessageType drResponse in gdResponse.ResponseMessages)
                    {
                        reList.Add(drResponse);
                    }
                }

            }
            else
            {
                Console.WriteLine("Error Getting Delegates from Mailbox " + emEmailAdrress);
                if (gdResponse.MessageText != null)
                {
                    Console.WriteLine(gdResponse.MessageText.ToString());
                }
            }
            return reList;
        }
        public List<DelegateUserResponseMessageType> getDeletgatesdebug(String emEmailAdrress)
        {
            Console.WriteLine("Debugging Delegate Code");
            Console.WriteLine("MailboxAddress : " + emEmailAddress.ToString());
            List<DelegateUserResponseMessageType> reList = new List<DelegateUserResponseMessageType>();
            GetDelegateType gdGetDelegate = new GetDelegateType();
            gdGetDelegate.IncludePermissions = true;
            gdGetDelegate.Mailbox = new EmailAddressType();
            gdGetDelegate.Mailbox.EmailAddress = emEmailAdrress;
            GetDelegateResponseMessageType gdResponse = esb.GetDelegate(gdGetDelegate);
            if (gdResponse.ResponseMessages != null)
            {
                Console.WriteLine("Request Result :" + gdResponse.ResponseMessages[0].ResponseClass);
            }
            else {
                Console.WriteLine("Response Null");
            
            }
            if (gdResponse != null)
            {
                if (gdResponse.ResponseClass == ResponseClassType.Success)
                {
                    if (gdResponse.ResponseMessages != null)
                    {
                        foreach (DelegateUserResponseMessageType drResponse in gdResponse.ResponseMessages)
                        {
                            reList.Add(drResponse);
                        }
                    }

                }
                else
                {
                    Console.WriteLine("Error Getting Delegates from Mailbox " + emEmailAdrress);
                    if (gdResponse.MessageText != null)
                    {
                        Console.WriteLine(gdResponse.MessageText.ToString());
                    }
                }
            }
            else { Console.WriteLine("No Response at all from EWS"); }
            return reList;
        }     
        public string enumOutlookRole(CalendarPermissionType cpCalendarPermissions)
        {
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
            else
            {
                orOutlookRole = cpCalendarPermissions.CalendarPermissionLevel.ToString();
            }
            return orOutlookRole;

        }
        private String hexCalPerms(CalendarPermissionType cpCalendarPermissions)
        {
            byte[] pmPermMask = new byte[7];
            if (cpCalendarPermissions.CanCreateItemsSpecified == true)
            {
                if (cpCalendarPermissions.CanCreateItems == true)
                {
                    pmPermMask[0] = 0x1;
                }
            }
            if (cpCalendarPermissions.CanCreateSubFoldersSpecified == true)
            {
                if (cpCalendarPermissions.CanCreateSubFolders == true)
                {
                    pmPermMask[1] = 0x1;
                }
            }
            if (cpCalendarPermissions.DeleteItemsSpecified == true)
            {
                switch (cpCalendarPermissions.DeleteItems)
                {
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
            if (cpCalendarPermissions.EditItemsSpecified == true)
            {
                switch (cpCalendarPermissions.EditItems)
                {
                    case PermissionActionType.None:
                        pmPermMask[3] = 0x1;
                        break;
                    case PermissionActionType.All:
                        pmPermMask[3] = 0x2;
                        break;
                    case PermissionActionType.Owned:
                        pmPermMask[3] = 0x3;
                        break;
                }
            }
            if (cpCalendarPermissions.IsFolderOwnerSpecified == true)
            {
                if (cpCalendarPermissions.IsFolderOwner == true)
                {
                    pmPermMask[4] = 0x1;
                }

            }
            if (cpCalendarPermissions.ReadItemsSpecified == true)
            {
                switch (cpCalendarPermissions.ReadItems)
                {
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
            if (cpCalendarPermissions.IsFolderVisibleSpecified == true)
            {
                if (cpCalendarPermissions.IsFolderVisible == true)
                {
                    pmPermMask[6] = 0x1;
                }
            }
            return BitConverter.ToString(pmPermMask);

        }
        public Hashtable FindSyncContacts(BaseFolderIdType[] biArray, PathToExtendedFieldType uifieldType, String propval)
        {
            Hashtable rtReturnList = new Hashtable();

            FindItemType fiFindItemRequest = new FindItemType();
            ItemResponseShapeType ipItemProperties = new ItemResponseShapeType();
            fiFindItemRequest.Traversal = ItemQueryTraversalType.Shallow;
            ipItemProperties.BaseShape = DefaultShapeNamesType.AllProperties;

            PathToExtendedFieldType dnEmailDisplayName1 = new PathToExtendedFieldType();
            dnEmailDisplayName1.DistinguishedPropertySetIdSpecified = true;
            dnEmailDisplayName1.DistinguishedPropertySetId = DistinguishedPropertySetType.Address;
            dnEmailDisplayName1.PropertyId = 0x8080;
            dnEmailDisplayName1.PropertyIdSpecified = true;
            dnEmailDisplayName1.PropertyType = MapiPropertyTypeType.String;
            PathToExtendedFieldType dnEmailDisplayName2 = new PathToExtendedFieldType();
            dnEmailDisplayName2.DistinguishedPropertySetIdSpecified = true;
            dnEmailDisplayName2.DistinguishedPropertySetId = DistinguishedPropertySetType.Address;
            dnEmailDisplayName2.PropertyId = 0x8090;
            dnEmailDisplayName2.PropertyIdSpecified = true;
            dnEmailDisplayName2.PropertyType = MapiPropertyTypeType.String;
            PathToExtendedFieldType dnEmailDisplayName3 = new PathToExtendedFieldType();
            dnEmailDisplayName3.DistinguishedPropertySetIdSpecified = true;
            dnEmailDisplayName3.DistinguishedPropertySetId = DistinguishedPropertySetType.Address;
            dnEmailDisplayName3.PropertyId = 0x80A0;
            dnEmailDisplayName3.PropertyIdSpecified = true;
            dnEmailDisplayName3.PropertyType = MapiPropertyTypeType.String;
            PathToExtendedFieldType daEmailDisplayName1 = new PathToExtendedFieldType();
            daEmailDisplayName1.DistinguishedPropertySetIdSpecified = true;
            daEmailDisplayName1.DistinguishedPropertySetId = DistinguishedPropertySetType.Address;
            daEmailDisplayName1.PropertyId = 0x8084;
            daEmailDisplayName1.PropertyIdSpecified = true;
            daEmailDisplayName1.PropertyType = MapiPropertyTypeType.String;
            PathToExtendedFieldType daEmailDisplayName2 = new PathToExtendedFieldType();
            daEmailDisplayName2.DistinguishedPropertySetIdSpecified = true;
            daEmailDisplayName2.DistinguishedPropertySetId = DistinguishedPropertySetType.Address;
            daEmailDisplayName2.PropertyId = 0x8094;
            daEmailDisplayName2.PropertyIdSpecified = true;
            daEmailDisplayName2.PropertyType = MapiPropertyTypeType.String;
            PathToExtendedFieldType daEmailDisplayName3 = new PathToExtendedFieldType();
            daEmailDisplayName3.DistinguishedPropertySetIdSpecified = true;
            daEmailDisplayName3.DistinguishedPropertySetId = DistinguishedPropertySetType.Address;
            daEmailDisplayName3.PropertyId = 0x80A4;
            daEmailDisplayName3.PropertyIdSpecified = true;
            daEmailDisplayName3.PropertyType = MapiPropertyTypeType.String;

            PathToExtendedFieldType atEmailAddressType1 = new PathToExtendedFieldType();
            atEmailAddressType1.DistinguishedPropertySetIdSpecified = true;
            atEmailAddressType1.DistinguishedPropertySetId = DistinguishedPropertySetType.Address;
            atEmailAddressType1.PropertyId = 0x8082;
            atEmailAddressType1.PropertyIdSpecified = true;
            atEmailAddressType1.PropertyType = MapiPropertyTypeType.String;


            PathToExtendedFieldType atEmailAddressType2 = new PathToExtendedFieldType();
            atEmailAddressType2.DistinguishedPropertySetIdSpecified = true;
            atEmailAddressType2.DistinguishedPropertySetId = DistinguishedPropertySetType.Address;
            atEmailAddressType2.PropertyId = 0x8092;
            atEmailAddressType2.PropertyIdSpecified = true;
            atEmailAddressType2.PropertyType = MapiPropertyTypeType.String;


            PathToExtendedFieldType atEmailAddressType3 = new PathToExtendedFieldType();
            atEmailAddressType3.DistinguishedPropertySetIdSpecified = true;
            atEmailAddressType3.DistinguishedPropertySetId = DistinguishedPropertySetType.Address;
            atEmailAddressType3.PropertyId = 0x80A2;
            atEmailAddressType3.PropertyIdSpecified = true;
            atEmailAddressType3.PropertyType = MapiPropertyTypeType.String;


            ipItemProperties.AdditionalProperties = new BasePathToElementType[10] { uifieldType,dnEmailDisplayName1, dnEmailDisplayName2, dnEmailDisplayName3
                    , daEmailDisplayName1, daEmailDisplayName2, daEmailDisplayName3,atEmailAddressType1,atEmailAddressType2,atEmailAddressType3 };


            fiFindItemRequest.ItemShape = ipItemProperties;
            fiFindItemRequest.ParentFolderIds = biArray;
            //Add Restriction for CustomProp
            RestrictionType ffRestriction = new RestrictionType();
            IsEqualToType ieToType = new IsEqualToType();
            FieldURIOrConstantType ciConstantType = new FieldURIOrConstantType();
            ConstantValueType cvConstantValueType = new ConstantValueType();
            cvConstantValueType.Value = propval;
            ciConstantType.Item = cvConstantValueType;
            ieToType.Item = uifieldType;
            ieToType.FieldURIOrConstant = ciConstantType;
            ffRestriction.Item = ieToType;
            fiFindItemRequest.Restriction = ffRestriction;
            FindItemResponseType frFindItemResponse = esb.FindItem(fiFindItemRequest);
            if (frFindItemResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
            {
                foreach (FindItemResponseMessageType firmtMessage in frFindItemResponse.ResponseMessages.Items)
                {
                    Console.WriteLine("Number of Items Found : " + firmtMessage.RootFolder.TotalItemsInView);
                    if (firmtMessage.RootFolder.TotalItemsInView > 0)
                    {
                        foreach (ContactItemType miMailboxItem in ((ArrayOfRealItemsType)firmtMessage.RootFolder.Item).Items)
                        {
                            rtReturnList.Add(miMailboxItem.EmailAddresses[0].Value.ToString(), miMailboxItem);
                        }

                    }
                }
            }
            else
            {
                Console.WriteLine("Error Occured");
                Console.WriteLine(frFindItemResponse.ResponseMessages.Items[0].MessageText.ToString());
            }
            return rtReturnList;
        }
        public List<SetItemFieldType> DiffConact(ContactItemType cnContact1, ContactItemType cnContact2)
        {
            List<SetItemFieldType> rlReturnList = new List<SetItemFieldType>();
            if (cnContact1.AssistantName != null)
            {
                if (cnContact2.AssistantName != null)
                {
                    if (cnContact1.AssistantName.ToString() != cnContact2.AssistantName.ToString())
                    {
                        ContactItemType asAssistantName = new ContactItemType();
                        asAssistantName.AssistantName = cnContact1.AssistantName;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsAssistantName, asAssistantName));
                    }
                }
                else
                {
                    ContactItemType asAssistantName = new ContactItemType();
                    asAssistantName.AssistantName = cnContact1.AssistantName;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsAssistantName, asAssistantName));
                }
            }
            if (cnContact1.BusinessHomePage != null)
            {
                if (cnContact2.BusinessHomePage != null)
                {
                    if (cnContact1.BusinessHomePage.ToString() != cnContact2.BusinessHomePage.ToString())
                    {
                        ContactItemType bhBusinessHomePage = new ContactItemType();
                        bhBusinessHomePage.BusinessHomePage = cnContact1.BusinessHomePage;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsBusinessHomePage, bhBusinessHomePage));
                    }
                }
                else
                {
                    ContactItemType bhBusinessHomePage = new ContactItemType();
                    bhBusinessHomePage.BusinessHomePage = cnContact1.BusinessHomePage;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsBusinessHomePage, bhBusinessHomePage));
                }
            }
            if (cnContact1.Categories != null)
            {
                if (cnContact2.Categories != null)
                {
                    string cnContact1Cateogies = String.Join(";", cnContact1.Categories);
                    string cnContact2Cateogies = String.Join(";", cnContact2.Categories);
                    if (cnContact1Cateogies != cnContact2Cateogies)
                    {
                        ContactItemType cnCategories = new ContactItemType();
                        cnCategories.Categories = cnContact1.Categories;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.itemCategories, cnCategories));
                    }
                }
                else
                {
                    ContactItemType cnCategories = new ContactItemType();
                    cnCategories.Categories = cnContact1.Categories;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.itemCategories, cnCategories));
                }
            }
            if (cnContact1.Children != null)
            {
                if (cnContact2.Children != null)
                {
                    string cnContact1Children = String.Join(";", cnContact1.Children);
                    string cnContact2Children = String.Join(";", cnContact2.Children);
                    if (cnContact1Children != cnContact2Children)
                    {
                        ContactItemType cnChildren = new ContactItemType();
                        cnChildren.Children = cnContact1.Children;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsChildren, cnChildren));
                    }
                }
                else
                {
                    ContactItemType cnChildren = new ContactItemType();
                    cnChildren.Children = cnContact1.Children;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsChildren, cnChildren));
                }
            }
            if (cnContact1.Companies != null)
            {
                if (cnContact2.Companies != null)
                {
                    string cnContact1Companies = String.Join(";", cnContact1.Companies);
                    string cnContact2Companies = String.Join(";", cnContact2.Companies);
                    if (cnContact1Companies != cnContact2Companies)
                    {
                        ContactItemType cnCompanies = new ContactItemType();
                        cnCompanies.Companies = cnContact1.Companies;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsCompanies, cnCompanies));
                    }
                }
                else
                {
                    ContactItemType cnCompanies = new ContactItemType();
                    cnCompanies.Companies = cnContact1.Companies;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsCompanies, cnCompanies));
                }
            }

            if (cnContact1.CompanyName != null)
            {
                if (cnContact2.CompanyName != null)
                {
                    if (cnContact1.CompanyName.ToString() != cnContact2.CompanyName.ToString())
                    {
                        ContactItemType cnCompanyName = new ContactItemType();
                        cnCompanyName.CompanyName = cnContact1.CompanyName;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsCompanyName, cnCompanyName));
                    }
                }
                else
                {
                    ContactItemType cnCompanyName = new ContactItemType();
                    cnCompanyName.CompanyName = cnContact1.CompanyName;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsCompanyName, cnCompanyName));
                }
            }
            if (cnContact1.Department != null)
            {
                if (cnContact2.Department != null)
                {
                    if (cnContact1.Department.ToString() != cnContact2.Department.ToString())
                    {
                        ContactItemType cnDepartment = new ContactItemType();
                        cnDepartment.Department = cnContact1.Department;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsDepartment, cnDepartment));
                    }
                }
                else
                {
                    ContactItemType cnDepartment = new ContactItemType();
                    cnDepartment.Department = cnContact1.Department;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsDepartment, cnDepartment));
                }
            }
            if (cnContact1.DisplayName != null)
            {
                if (cnContact2.DisplayName != null)
                {
                    if (cnContact1.DisplayName.ToString() != cnContact2.DisplayName.ToString())
                    {
                        ContactItemType dnDisplayName = new ContactItemType();
                        dnDisplayName.DisplayName = cnContact1.DisplayName;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsDisplayName, dnDisplayName));
                    }
                }
                else
                {
                    ContactItemType dnDisplayName = new ContactItemType();
                    dnDisplayName.DisplayName = cnContact1.DisplayName;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsDisplayName, dnDisplayName));
                }
            }
            if (cnContact1.EmailAddresses != null)
            {
                if (cnContact2.EmailAddresses != null)
                {
                    Hashtable emAddress1 = new Hashtable();
                    Hashtable emAddress2 = new Hashtable();
                    for (int emadarr = 0; emadarr < cnContact1.EmailAddresses.Length; emadarr++)
                    {
                        emAddress1.Add(cnContact1.EmailAddresses[emadarr].Key.ToString(), cnContact1.EmailAddresses[emadarr].Value.ToString());
                    }
                    for (int emadarr = 0; emadarr < cnContact2.EmailAddresses.Length; emadarr++)
                    {
                        emAddress2.Add(cnContact2.EmailAddresses[emadarr].Key.ToString(), cnContact2.EmailAddresses[emadarr].Value.ToString());
                    }
                    foreach (String key in emAddress1.Keys)
                    {
                        Boolean upUpdateEmail = false;
                        if (emAddress2.ContainsKey(key))
                        {
                            if (emAddress2[key].ToString() != emAddress1[key].ToString()) { upUpdateEmail = true; }
                        }
                        else
                        {
                            if (emAddress1[key].ToString() != "")
                            {
                                upUpdateEmail = true;
                            }
                        }
                        if (upUpdateEmail == true)
                        {
                            switch (key)
                            {
                                case "EmailAddress1": rlReturnList.Add(AddContactDifferanceEmail("EmailAddress1", EmailAddressKeyType.EmailAddress1, emAddress1[key].ToString()));
                                    break;
                                case "EmailAddress2": rlReturnList.Add(AddContactDifferanceEmail("EmailAddress2", EmailAddressKeyType.EmailAddress2, emAddress1[key].ToString()));
                                    break;
                                case "EmailAddress3": rlReturnList.Add(AddContactDifferanceEmail("EmailAddress3", EmailAddressKeyType.EmailAddress3, emAddress1[key].ToString()));
                                    break;
                            }
                        }

                    }
                }
                else
                {
                    for (int emadarr = 0; emadarr < cnContact1.EmailAddresses.Length; emadarr++)
                    {
                        switch (cnContact1.EmailAddresses[emadarr].Key)
                        {
                            case EmailAddressKeyType.EmailAddress1: rlReturnList.Add(AddContactDifferanceEmail("EmailAddress1", EmailAddressKeyType.EmailAddress1, cnContact1.EmailAddresses[emadarr].Value));
                                break;
                            case EmailAddressKeyType.EmailAddress2: rlReturnList.Add(AddContactDifferanceEmail("EmailAddress2", EmailAddressKeyType.EmailAddress2, cnContact1.EmailAddresses[emadarr].Value));
                                break;
                            case EmailAddressKeyType.EmailAddress3: rlReturnList.Add(AddContactDifferanceEmail("EmailAddress3", EmailAddressKeyType.EmailAddress3, cnContact1.EmailAddresses[emadarr].Value));
                                break;
                        }
                    }
                }
            }
            if (cnContact1.GivenName != null)
            {
                if (cnContact2.GivenName != null)
                {
                    if (cnContact1.GivenName.ToString() != cnContact2.GivenName.ToString())
                    {
                        ContactItemType cnGivenName = new ContactItemType();
                        cnGivenName.GivenName = cnContact1.GivenName;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsGivenName, cnGivenName));
                    }
                }
                else
                {
                    ContactItemType cnGivenName = new ContactItemType();
                    cnGivenName.GivenName = cnContact1.GivenName;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsGivenName, cnGivenName));
                }
            }
            if (cnContact1.JobTitle != null)
            {
                if (cnContact2.JobTitle != null)
                {
                    if (cnContact1.JobTitle.ToString() != cnContact2.JobTitle.ToString())
                    {
                        ContactItemType jnJobTitle = new ContactItemType();
                        jnJobTitle.JobTitle = cnContact1.JobTitle;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsJobTitle, jnJobTitle));
                    }
                }
                else
                {
                    ContactItemType jnJobTitle = new ContactItemType();
                    jnJobTitle.JobTitle = cnContact1.JobTitle;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsJobTitle, jnJobTitle));
                }
            }
            if (cnContact1.Manager != null)
            {
                if (cnContact2.Manager != null)
                {
                    if (cnContact1.Manager.ToString() != cnContact2.Manager.ToString())
                    {
                        ContactItemType mnManager = new ContactItemType();
                        mnManager.Manager = cnContact1.Manager;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsManager, mnManager));
                    }
                }
                else
                {
                    ContactItemType mnManager = new ContactItemType();
                    mnManager.Manager = cnContact1.Manager;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsManager, mnManager));
                }
            }
            if (cnContact1.MiddleName != null)
            {
                if (cnContact2.MiddleName != null)
                {
                    if (cnContact1.MiddleName.ToString() != cnContact2.MiddleName.ToString())
                    {
                        ContactItemType mnMiddleName = new ContactItemType();
                        mnMiddleName.MiddleName = cnContact1.MiddleName;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsMiddleName, mnMiddleName));
                    }
                }
                else
                {
                    ContactItemType mnMiddleName = new ContactItemType();
                    mnMiddleName.MiddleName = cnContact1.MiddleName;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsMiddleName, mnMiddleName));
                }
            }
            if (cnContact1.Mileage != null)
            {
                if (cnContact2.Mileage != null)
                {
                    if (cnContact1.Mileage.ToString() != cnContact2.Mileage.ToString())
                    {
                        ContactItemType mnMileage = new ContactItemType();
                        mnMileage.Mileage = cnContact1.Mileage;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsMileage, mnMileage));
                    }
                }
                else
                {
                    ContactItemType mnMileage = new ContactItemType();
                    mnMileage.Mileage = cnContact1.Mileage;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsMileage, mnMileage));
                }
            }
            if (cnContact1.Nickname != null)
            {
                if (cnContact2.Nickname != null)
                {
                    if (cnContact1.Nickname.ToString() != cnContact2.Nickname.ToString())
                    {
                        ContactItemType nnNickName = new ContactItemType();
                        nnNickName.Nickname = cnContact1.Nickname;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsNickname, nnNickName));
                    }
                }
                else
                {
                    ContactItemType nnNickName = new ContactItemType();
                    nnNickName.Nickname = cnContact1.Nickname;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsNickname, nnNickName));
                }
            }
            if (cnContact1.OfficeLocation != null)
            {
                if (cnContact2.OfficeLocation != null)
                {
                    if (cnContact1.OfficeLocation.ToString() != cnContact2.OfficeLocation.ToString())
                    {
                        ContactItemType onOfficeName = new ContactItemType();
                        onOfficeName.OfficeLocation = cnContact1.OfficeLocation;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsOfficeLocation, onOfficeName));
                    }
                }
                else
                {
                    ContactItemType onOfficeName = new ContactItemType();
                    onOfficeName.OfficeLocation = cnContact1.OfficeLocation;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsOfficeLocation, onOfficeName));
                }
            }
            if (cnContact1.PhysicalAddresses != null)
            {
                if (cnContact2.PhysicalAddresses != null)
                {
                    Hashtable phAddress1 = new Hashtable();
                    Hashtable phAddress2 = new Hashtable();
                    for (int phadarr = 0; phadarr < cnContact1.PhysicalAddresses.Length; phadarr++)
                    {
                        Hashtable phAddress3 = new Hashtable();
                        if (cnContact1.PhysicalAddresses[phadarr].City != null) { phAddress3.Add("City", cnContact1.PhysicalAddresses[phadarr].City); }
                        if (cnContact1.PhysicalAddresses[phadarr].CountryOrRegion != null) { phAddress3.Add("CountryOrRegion", cnContact1.PhysicalAddresses[phadarr].CountryOrRegion); }
                        if (cnContact1.PhysicalAddresses[phadarr].PostalCode != null) { phAddress3.Add("PostalCode", cnContact1.PhysicalAddresses[phadarr].PostalCode); }
                        if (cnContact1.PhysicalAddresses[phadarr].State != null) { phAddress3.Add("State", cnContact1.PhysicalAddresses[phadarr].State); }
                        if (cnContact1.PhysicalAddresses[phadarr].Street != null) { phAddress3.Add("Street", cnContact1.PhysicalAddresses[phadarr].Street); }
                        phAddress1.Add(cnContact1.PhysicalAddresses[phadarr].Key.ToString(), phAddress3);
                    }
                    for (int phadarr = 0; phadarr < cnContact2.PhysicalAddresses.Length; phadarr++)
                    {
                        Hashtable phAddress4 = new Hashtable();
                        if (cnContact2.PhysicalAddresses[phadarr].City != null) { phAddress4.Add("City", cnContact2.PhysicalAddresses[phadarr].City); }
                        if (cnContact2.PhysicalAddresses[phadarr].CountryOrRegion != null) { phAddress4.Add("CountryOrRegion", cnContact2.PhysicalAddresses[phadarr].CountryOrRegion); }
                        if (cnContact2.PhysicalAddresses[phadarr].PostalCode != null) { phAddress4.Add("PostalCode", cnContact2.PhysicalAddresses[phadarr].PostalCode); }
                        if (cnContact2.PhysicalAddresses[phadarr].State != null) { phAddress4.Add("State", cnContact2.PhysicalAddresses[phadarr].State); }
                        if (cnContact2.PhysicalAddresses[phadarr].Street != null) { phAddress4.Add("Street", cnContact2.PhysicalAddresses[phadarr].Street); }
                        phAddress2.Add(cnContact2.PhysicalAddresses[phadarr].Key.ToString(), phAddress4);
                    }
                    foreach (String key in phAddress1.Keys)
                    {
                        if (phAddress2.ContainsKey(key))
                        {
                            Hashtable adhash1 = (Hashtable)phAddress1[key];
                            Hashtable adhash2 = (Hashtable)phAddress2[key];
                            foreach (String ahhashkey in adhash1.Keys)
                            {
                                Boolean upUpdateAddress = false;
                                if (adhash2.ContainsKey(ahhashkey))
                                {
                                    if (adhash1[ahhashkey].ToString() != adhash2[ahhashkey].ToString())
                                    {
                                        upUpdateAddress = true;
                                    }
                                }
                                else
                                {
                                    if (adhash1[ahhashkey].ToString() != "")
                                    {
                                        upUpdateAddress = true;
                                    }
                                }
                                if (upUpdateAddress == true)
                                {
                                    PhysicalAddressDictionaryEntryType paAddress = new PhysicalAddressDictionaryEntryType();
                                    switch (key)
                                    {
                                        case "Business": paAddress.Key = PhysicalAddressKeyType.Business;
                                            break;
                                        case "Home": paAddress.Key = PhysicalAddressKeyType.Home;
                                            break;
                                        case "Other": paAddress.Key = PhysicalAddressKeyType.Other;
                                            break;
                                    }
                                    switch (ahhashkey)
                                    {
                                        case "City": paAddress.City = adhash1[ahhashkey].ToString();
                                            rlReturnList.Add(AddContactDifferancePhyicalAddress(paAddress.Key.ToString(), DictionaryURIType.contactsPhysicalAddressCity, paAddress));
                                            break;
                                        case "CountryOrRegion": paAddress.CountryOrRegion = adhash1[ahhashkey].ToString();
                                            rlReturnList.Add(AddContactDifferancePhyicalAddress(paAddress.Key.ToString(), DictionaryURIType.contactsPhysicalAddressCountryOrRegion, paAddress));
                                            break;
                                        case "PostalCode": paAddress.PostalCode = adhash1[ahhashkey].ToString();
                                            rlReturnList.Add(AddContactDifferancePhyicalAddress(paAddress.Key.ToString(), DictionaryURIType.contactsPhysicalAddressPostalCode, paAddress));
                                            break;
                                        case "State": paAddress.State = adhash1[ahhashkey].ToString();
                                            rlReturnList.Add(AddContactDifferancePhyicalAddress(paAddress.Key.ToString(), DictionaryURIType.contactsPhysicalAddressState, paAddress));
                                            break;
                                        case "Street": paAddress.Street = adhash1[ahhashkey].ToString();
                                            rlReturnList.Add(AddContactDifferancePhyicalAddress(paAddress.Key.ToString(), DictionaryURIType.contactsPhysicalAddressStreet, paAddress));
                                            break;

                                    }
                                }
                            }
                        }

                    }
                }
                else
                {
                    for (int phadarr = 0; phadarr < cnContact1.PhysicalAddresses.Length; phadarr++)
                    {
                        PhysicalAddressDictionaryEntryType paAddress = new PhysicalAddressDictionaryEntryType();
                        paAddress.Key = cnContact1.PhysicalAddresses[phadarr].Key;
                        if (cnContact1.PhysicalAddresses[phadarr].City != null)
                        {
                            paAddress.City = cnContact1.PhysicalAddresses[phadarr].City;
                            rlReturnList.Add(AddContactDifferancePhyicalAddress(cnContact1.PhysicalAddresses[phadarr].Key.ToString(), DictionaryURIType.contactsPhysicalAddressCity, paAddress));
                        }
                        if (cnContact1.PhysicalAddresses[phadarr].CountryOrRegion != null)
                        {
                            paAddress.CountryOrRegion = cnContact1.PhysicalAddresses[phadarr].CountryOrRegion;
                            AddContactDifferancePhyicalAddress(cnContact1.PhysicalAddresses[phadarr].Key.ToString(), DictionaryURIType.contactsPhysicalAddressCountryOrRegion, paAddress);
                        }
                        if (cnContact1.PhysicalAddresses[phadarr].PostalCode != null)
                        {
                            paAddress.PostalCode = cnContact1.PhysicalAddresses[phadarr].PostalCode;
                            AddContactDifferancePhyicalAddress(cnContact1.PhysicalAddresses[phadarr].Key.ToString(), DictionaryURIType.contactsPhysicalAddressPostalCode, paAddress);
                        }
                        if (cnContact1.PhysicalAddresses[phadarr].State != null)
                        {
                            paAddress.State = cnContact1.PhysicalAddresses[phadarr].State;
                            AddContactDifferancePhyicalAddress(cnContact1.PhysicalAddresses[phadarr].Key.ToString(), DictionaryURIType.contactsPhysicalAddressState, paAddress);
                        }
                        if (cnContact1.PhysicalAddresses[phadarr].Street != null)
                        {
                            paAddress.Street = cnContact1.PhysicalAddresses[phadarr].Street;
                            AddContactDifferancePhyicalAddress(cnContact1.PhysicalAddresses[phadarr].Key.ToString(), DictionaryURIType.contactsPhysicalAddressStreet, paAddress);
                        }
                    }
                }
            }
            if (cnContact1.PhoneNumbers != null)
            {
                if (cnContact2.PhoneNumbers != null)
                {
                    Hashtable phPhone1 = new Hashtable();
                    Hashtable phPhone2 = new Hashtable();
                    for (int phPhoneint = 0; phPhoneint < cnContact1.PhoneNumbers.Length; phPhoneint++)
                    {
                        if (cnContact1.PhoneNumbers[phPhoneint].Value != null)
                        {
                            phPhone1.Add(cnContact1.PhoneNumbers[phPhoneint].Key.ToString(), cnContact1.PhoneNumbers[phPhoneint].Value);
                        }
                    }
                    for (int phPhoneint = 0; phPhoneint < cnContact2.PhoneNumbers.Length; phPhoneint++)
                    {
                        if (cnContact2.PhoneNumbers[phPhoneint].Value != null)
                        {
                            phPhone2.Add(cnContact2.PhoneNumbers[phPhoneint].Key.ToString(), cnContact2.PhoneNumbers[phPhoneint].Value);
                        }
                    }
                    foreach (String key in phPhone1.Keys)
                    {
                        Boolean upUpdatePhone = false;
                        if (phPhone2.ContainsKey(key))
                        {
                            if (phPhone2[key].ToString() != phPhone1[key].ToString()) { upUpdatePhone = true; }
                        }
                        else
                        {
                            if (phPhone1[key].ToString() != "")
                            {
                                upUpdatePhone = true;
                            }
                        }
                        if (upUpdatePhone == true)
                        {
                            switch (key)
                            {
                                case "AssistantPhone": rlReturnList.Add(AddContactDifferancePhone("AssistantPhone", PhoneNumberKeyType.AssistantPhone, phPhone1[key].ToString()));
                                    break;
                                case "BusinessFax": rlReturnList.Add(AddContactDifferancePhone("BusinessFax", PhoneNumberKeyType.BusinessFax, phPhone1[key].ToString()));
                                    break;
                                case "BusinessPhone": rlReturnList.Add(AddContactDifferancePhone("BusinessPhone", PhoneNumberKeyType.BusinessPhone, phPhone1[key].ToString()));
                                    break;
                                case "BusinessPhone2": rlReturnList.Add(AddContactDifferancePhone("BusinessPhone2", PhoneNumberKeyType.BusinessPhone2, phPhone1[key].ToString()));
                                    break;
                                case "Callback": rlReturnList.Add(AddContactDifferancePhone("Callback", PhoneNumberKeyType.Callback, phPhone1[key].ToString()));
                                    break;
                                case "CarPhone": rlReturnList.Add(AddContactDifferancePhone("CarPhone", PhoneNumberKeyType.CarPhone, phPhone1[key].ToString()));
                                    break;
                                case "CompanyMainPhone": rlReturnList.Add(AddContactDifferancePhone("CompanyMainPhone", PhoneNumberKeyType.CompanyMainPhone, phPhone1[key].ToString()));
                                    break;
                                case "HomeFax": rlReturnList.Add(AddContactDifferancePhone("HomeFax", PhoneNumberKeyType.HomeFax, phPhone1[key].ToString()));
                                    break;
                                case "HomePhone": rlReturnList.Add(AddContactDifferancePhone("HomePhone", PhoneNumberKeyType.HomePhone, phPhone1[key].ToString()));
                                    break;
                                case "HomePhone2": rlReturnList.Add(AddContactDifferancePhone("HomePhone2", PhoneNumberKeyType.HomePhone2, phPhone1[key].ToString()));
                                    break;
                                case "Isdn": rlReturnList.Add(AddContactDifferancePhone("Isdn", PhoneNumberKeyType.Isdn, phPhone1[key].ToString()));
                                    break;
                                case "MobilePhone": rlReturnList.Add(AddContactDifferancePhone("MobilePhone", PhoneNumberKeyType.MobilePhone, phPhone1[key].ToString()));
                                    break;
                                case "OtherFax": rlReturnList.Add(AddContactDifferancePhone("OtherFax", PhoneNumberKeyType.OtherFax, phPhone1[key].ToString()));
                                    break;
                                case "OtherTelephone": rlReturnList.Add(AddContactDifferancePhone("OtherTelephone", PhoneNumberKeyType.OtherTelephone, phPhone1[key].ToString()));
                                    break;
                                case "Pager": rlReturnList.Add(AddContactDifferancePhone("Pager", PhoneNumberKeyType.Pager, phPhone1[key].ToString()));
                                    break;
                                case "PrimaryPhone": rlReturnList.Add(AddContactDifferancePhone("PrimaryPhone", PhoneNumberKeyType.PrimaryPhone, phPhone1[key].ToString()));
                                    break;
                                case "RadioPhone": rlReturnList.Add(AddContactDifferancePhone("RadioPhone", PhoneNumberKeyType.RadioPhone, phPhone1[key].ToString()));
                                    break;
                                case "Telex": rlReturnList.Add(AddContactDifferancePhone("Telex", PhoneNumberKeyType.Telex, phPhone1[key].ToString()));
                                    break;
                                case "TtyTddPhone": rlReturnList.Add(AddContactDifferancePhone("TtyTddPhone", PhoneNumberKeyType.TtyTddPhone, phPhone1[key].ToString()));
                                    break;
                            }
                        }
                    }
                }
                else
                {
                    for (int phPhoneint = 0; phPhoneint < cnContact1.PhoneNumbers.Length; phPhoneint++)
                    {
                        switch (cnContact1.PhoneNumbers[phPhoneint].Key)
                        {
                            case PhoneNumberKeyType.AssistantPhone: rlReturnList.Add(AddContactDifferancePhone("AssistantPhone", PhoneNumberKeyType.AssistantPhone, cnContact1.PhoneNumbers[phPhoneint].Value));
                                break;
                            case PhoneNumberKeyType.BusinessFax: rlReturnList.Add(AddContactDifferancePhone("BusinessFax", PhoneNumberKeyType.BusinessFax, cnContact1.PhoneNumbers[phPhoneint].Value));
                                break;
                            case PhoneNumberKeyType.BusinessPhone: rlReturnList.Add(AddContactDifferancePhone("BusinessPhone", PhoneNumberKeyType.BusinessPhone, cnContact1.PhoneNumbers[phPhoneint].Value));
                                break;
                            case PhoneNumberKeyType.BusinessPhone2: rlReturnList.Add(AddContactDifferancePhone("BusinessPhone2", PhoneNumberKeyType.BusinessPhone2, cnContact1.PhoneNumbers[phPhoneint].Value));
                                break;
                            case PhoneNumberKeyType.Callback: rlReturnList.Add(AddContactDifferancePhone("Callback", PhoneNumberKeyType.Callback, cnContact1.PhoneNumbers[phPhoneint].Value));
                                break;
                            case PhoneNumberKeyType.CarPhone: rlReturnList.Add(AddContactDifferancePhone("CarPhone", PhoneNumberKeyType.CarPhone, cnContact1.PhoneNumbers[phPhoneint].Value));
                                break;
                            case PhoneNumberKeyType.CompanyMainPhone: rlReturnList.Add(AddContactDifferancePhone("CompanyMainPhone", PhoneNumberKeyType.CompanyMainPhone, cnContact1.PhoneNumbers[phPhoneint].Value));
                                break;
                            case PhoneNumberKeyType.HomeFax: rlReturnList.Add(AddContactDifferancePhone("HomeFax", PhoneNumberKeyType.HomeFax, cnContact1.PhoneNumbers[phPhoneint].Value));
                                break;
                            case PhoneNumberKeyType.HomePhone: rlReturnList.Add(AddContactDifferancePhone("HomePhone", PhoneNumberKeyType.HomePhone, cnContact1.PhoneNumbers[phPhoneint].Value));
                                break;
                            case PhoneNumberKeyType.HomePhone2: rlReturnList.Add(AddContactDifferancePhone("HomePhone2", PhoneNumberKeyType.HomePhone2, cnContact1.PhoneNumbers[phPhoneint].Value));
                                break;
                            case PhoneNumberKeyType.Isdn: rlReturnList.Add(AddContactDifferancePhone("Isdn", PhoneNumberKeyType.Isdn, cnContact1.PhoneNumbers[phPhoneint].Value));
                                break;
                            case PhoneNumberKeyType.MobilePhone: rlReturnList.Add(AddContactDifferancePhone("MobilePhone", PhoneNumberKeyType.MobilePhone, cnContact1.PhoneNumbers[phPhoneint].Value));
                                break;
                            case PhoneNumberKeyType.OtherFax: rlReturnList.Add(AddContactDifferancePhone("OtherFax", PhoneNumberKeyType.OtherFax, cnContact1.PhoneNumbers[phPhoneint].Value));
                                break;
                            case PhoneNumberKeyType.OtherTelephone: rlReturnList.Add(AddContactDifferancePhone("OtherTelephone", PhoneNumberKeyType.OtherTelephone, cnContact1.PhoneNumbers[phPhoneint].Value));
                                break;
                            case PhoneNumberKeyType.Pager: rlReturnList.Add(AddContactDifferancePhone("Pager", PhoneNumberKeyType.Pager, cnContact1.PhoneNumbers[phPhoneint].Value));
                                break;
                            case PhoneNumberKeyType.PrimaryPhone: rlReturnList.Add(AddContactDifferancePhone("PrimaryPhone", PhoneNumberKeyType.PrimaryPhone, cnContact1.PhoneNumbers[phPhoneint].Value));
                                break;
                            case PhoneNumberKeyType.RadioPhone: rlReturnList.Add(AddContactDifferancePhone("RadioPhone", PhoneNumberKeyType.RadioPhone, cnContact1.PhoneNumbers[phPhoneint].Value));
                                break;
                            case PhoneNumberKeyType.Telex: rlReturnList.Add(AddContactDifferancePhone("Telex", PhoneNumberKeyType.Telex, cnContact1.PhoneNumbers[phPhoneint].Value));
                                break;
                            case PhoneNumberKeyType.TtyTddPhone: rlReturnList.Add(AddContactDifferancePhone("TtyTddPhone", PhoneNumberKeyType.TtyTddPhone, cnContact1.PhoneNumbers[phPhoneint].Value));
                                break;
                        }
                    }
                }
            }

            if (cnContact1.Profession != null)
            {
                if (cnContact2.Profession != null)
                {
                    if (cnContact1.Profession.ToString() != cnContact2.Profession.ToString())
                    {
                        ContactItemType prProfession = new ContactItemType();
                        prProfession.Profession = cnContact1.Profession;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsProfession, prProfession));
                    }
                }
                else
                {
                    ContactItemType prProfession = new ContactItemType();
                    prProfession.Profession = cnContact1.Profession;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsProfession, prProfession));
                }
            }
            if (cnContact1.SpouseName != null)
            {
                if (cnContact2.SpouseName != null)
                {
                    if (cnContact1.SpouseName.ToString() != cnContact2.SpouseName.ToString())
                    {
                        ContactItemType snSpouseName = new ContactItemType();
                        snSpouseName.SpouseName = cnContact1.SpouseName;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsSpouseName, snSpouseName));
                    }
                }
                else
                {
                    ContactItemType snSpouseName = new ContactItemType();
                    snSpouseName.SpouseName = cnContact1.SpouseName;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsSpouseName, snSpouseName));
                }
            }
            if (cnContact1.Subject != null)
            {
                if (cnContact2.Subject != null)
                {
                    if (cnContact1.Subject.ToString() != cnContact2.Subject.ToString())
                    {
                        ContactItemType snSubject = new ContactItemType();
                        snSubject.Subject = cnContact1.Subject;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.itemSubject, snSubject));
                    }
                }
                else
                {
                    ContactItemType snSubject = new ContactItemType();
                    snSubject.Subject = cnContact1.Subject;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.itemSubject, snSubject));
                }
            }
            if (cnContact1.Surname != null)
            {
                if (cnContact2.Surname != null)
                {
                    if (cnContact1.Surname.ToString() != cnContact2.Surname.ToString())
                    {
                        ContactItemType snSurname = new ContactItemType();
                        snSurname.Surname = cnContact1.Surname;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsSurname, snSurname));
                    }
                }
                else
                {
                    ContactItemType snSurname = new ContactItemType();
                    snSurname.Surname = cnContact1.Surname;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsSurname, snSurname));
                }
            }
            if (cnContact1.WeddingAnniversary != null)
            {
                if (cnContact2.WeddingAnniversary != null)
                {
                    if (cnContact1.WeddingAnniversary.ToLongDateString() != cnContact2.WeddingAnniversary.ToLongDateString())
                    {
                        ContactItemType waWeddingAnv = new ContactItemType();
                        waWeddingAnv.WeddingAnniversarySpecified = true;
                        waWeddingAnv.WeddingAnniversary = cnContact1.WeddingAnniversary;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsWeddingAnniversary, waWeddingAnv));
                    }
                }
                else
                {
                    ContactItemType waWeddingAnv = new ContactItemType();
                    waWeddingAnv.WeddingAnniversarySpecified = true;
                    waWeddingAnv.WeddingAnniversary = cnContact1.WeddingAnniversary;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsWeddingAnniversary, waWeddingAnv));
                }
            }
            if (cnContact1.Birthday != null)
            {
                if (cnContact2.Birthday != null)
                {
                    if (cnContact1.Birthday.ToLongDateString() != cnContact2.Birthday.ToLongDateString())
                    {
                        ContactItemType biBirthday = new ContactItemType();
                        biBirthday.BirthdaySpecified = true;
                        biBirthday.Birthday = cnContact1.Birthday;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsBirthday, biBirthday));
                    }
                }
                else
                {
                    ContactItemType biBirthday = new ContactItemType();
                    biBirthday.BirthdaySpecified = true;
                    biBirthday.Birthday = cnContact1.Birthday;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsBirthday, biBirthday));
                }
            }
            if (cnContact1.Initials != null)
            {
                if (cnContact2.Initials != null)
                {
                    if (cnContact1.Initials != cnContact2.Initials)
                    {
                        ContactItemType inInitials = new ContactItemType();
                        inInitials.Initials = cnContact1.Initials;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsInitials, inInitials));
                    }
                }
                else
                {
                    ContactItemType inInitials = new ContactItemType();
                    inInitials.Initials = cnContact1.Initials;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsInitials, inInitials));
                }
            }
            if (cnContact1.PostalAddressIndexSpecified == true)
            {
                if (cnContact2.PostalAddressIndexSpecified == true)
                {
                    if (cnContact1.PostalAddressIndex != cnContact2.PostalAddressIndex)
                    {
                        ContactItemType paAddressIndex = new ContactItemType();
                        paAddressIndex.PostalAddressIndexSpecified = true;
                        paAddressIndex.PostalAddressIndex = cnContact1.PostalAddressIndex;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsPostalAddressIndex, paAddressIndex));
                    }
                }
                else
                {
                    ContactItemType paAddressIndex = new ContactItemType();
                    paAddressIndex.PostalAddressIndexSpecified = true;
                    paAddressIndex.PostalAddressIndex = cnContact1.PostalAddressIndex;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsPostalAddressIndex, paAddressIndex));
                }
            }
            if (cnContact1.Generation != null)
            {
                if (cnContact2.Generation != null)
                {
                    if (cnContact1.Generation.ToString() != cnContact2.Generation.ToString())
                    {
                        ContactItemType gnGeneration = new ContactItemType();
                        gnGeneration.Generation = cnContact1.Generation;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsGeneration, gnGeneration));
                    }
                }
                else
                {
                    ContactItemType gnGeneration = new ContactItemType();
                    gnGeneration.Generation = cnContact1.Generation;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsGeneration, gnGeneration));
                }
            }
            if (cnContact1.FileAs != null)
            {
                if (cnContact2.FileAs != null)
                {
                    if (cnContact1.FileAs.ToString() != cnContact2.FileAs.ToString())
                    {
                        ContactItemType fnFileAS = new ContactItemType();
                        fnFileAS.FileAs = cnContact1.FileAs;
                        rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsFileAs, fnFileAS));
                    }
                }
                else
                {
                    ContactItemType fnFileAS = new ContactItemType();
                    fnFileAS.FileAs = cnContact1.FileAs;
                    rlReturnList.Add(AddContactDifferance(UnindexedFieldURIType.contactsFileAs, fnFileAS));
                }
            }
            if (cnContact1.ExtendedProperty != null)
            {
                if (cnContact2.ExtendedProperty != null)
                {
                    Hashtable exProp1 = new Hashtable();
                    Hashtable exProp2 = new Hashtable();
                    for (int exPropint = 0; exPropint < cnContact1.ExtendedProperty.Length; exPropint++)
                    {
                        if (cnContact1.ExtendedProperty[exPropint].ExtendedFieldURI.PropertyIdSpecified == true)
                        {
                            exProp1.Add(cnContact1.ExtendedProperty[exPropint].ExtendedFieldURI.PropertyId.ToString(), cnContact1.ExtendedProperty[exPropint]);
                        }
                        else
                        {
                            if (cnContact1.ExtendedProperty[exPropint].ExtendedFieldURI.PropertyTag != null)
                            {
                                exProp1.Add(cnContact1.ExtendedProperty[exPropint].ExtendedFieldURI.PropertyTag, cnContact1.ExtendedProperty[exPropint]);
                            }
                            else
                            {
                                exProp1.Add(cnContact1.ExtendedProperty[exPropint].ExtendedFieldURI.PropertyName, cnContact1.ExtendedProperty[exPropint]);
                            }
                        }
                    }
                    for (int exPropint = 0; exPropint < cnContact2.ExtendedProperty.Length; exPropint++)
                    {
                        if (cnContact2.ExtendedProperty[exPropint].ExtendedFieldURI.PropertyIdSpecified == true)
                        {
                            exProp2.Add(cnContact2.ExtendedProperty[exPropint].ExtendedFieldURI.PropertyId.ToString(), cnContact2.ExtendedProperty[exPropint]);
                        }
                        else
                        {
                            if (cnContact2.ExtendedProperty[exPropint].ExtendedFieldURI.PropertyTag != null)
                            {
                                exProp2.Add(cnContact2.ExtendedProperty[exPropint].ExtendedFieldURI.PropertyTag, cnContact2.ExtendedProperty[exPropint]);
                            }
                            else
                            {
                                exProp2.Add(cnContact2.ExtendedProperty[exPropint].ExtendedFieldURI.PropertyName, cnContact2.ExtendedProperty[exPropint]);
                            }
                        }
                    }
                    foreach (String key in exProp1.Keys)
                    {
                        if (exProp2.ContainsKey(key))
                        {
                            if (((ExtendedPropertyType)exProp2[key]).Item.ToString() != ((ExtendedPropertyType)exProp1[key]).Item.ToString())
                            {
                                rlReturnList.Add(AddContactDifferanceExtendedProp((ExtendedPropertyType)exProp1[key]));
                            }
                        }
                        else
                        {
                            rlReturnList.Add(AddContactDifferanceExtendedProp((ExtendedPropertyType)exProp1[key]));
                        }


                    }

                }
                else
                {
                    for (int exPropint = 0; exPropint < cnContact1.ExtendedProperty.Length; exPropint++)
                    {
                        rlReturnList.Add(AddContactDifferanceExtendedProp(cnContact1.ExtendedProperty[exPropint]));
                    }
                }
            }
            return rlReturnList;


        }
        public SetItemFieldType AddContactDifferance(UnindexedFieldURIType ptFieldURI, ContactItemType ctContact)
        {
            SetItemFieldType setItem = new SetItemFieldType();
            PathToUnindexedFieldType epExPath = new PathToUnindexedFieldType();
            epExPath.FieldURI = ptFieldURI;
            setItem.Item = epExPath;
            setItem.Item1 = ctContact;
            return setItem;
        }
        public SetItemFieldType AddContactDifferanceEmail(String fiFieldIndex, EmailAddressKeyType ktEmailKeyType, String emEmailAddress)
        {
            SetItemFieldType setItem = new SetItemFieldType();
            PathToIndexedFieldType em1indexedField = new PathToIndexedFieldType();
            em1indexedField.FieldIndex = fiFieldIndex;
            em1indexedField.FieldURI = DictionaryURIType.contactsEmailAddress;
            EmailAddressDictionaryEntryType emailAddress = new EmailAddressDictionaryEntryType();
            emailAddress.Key = ktEmailKeyType;
            emailAddress.Value = emEmailAddress;
            ContactItemType cnContact = new ContactItemType();
            cnContact.EmailAddresses = new EmailAddressDictionaryEntryType[1];
            cnContact.EmailAddresses[0] = emailAddress;
            setItem.Item = em1indexedField;
            setItem.Item1 = cnContact;
            return setItem;
        }
        public SetItemFieldType AddContactDifferancePhone(String fiFieldIndex, PhoneNumberKeyType pkPhoneKeyType, String phPhoneNumber)
        {
            SetItemFieldType setItem = new SetItemFieldType();
            PathToIndexedFieldType phindexedField = new PathToIndexedFieldType();
            phindexedField.FieldIndex = pkPhoneKeyType.ToString();
            phindexedField.FieldURI = DictionaryURIType.contactsPhoneNumber;
            PhoneNumberDictionaryEntryType phPhoneNumber1 = new PhoneNumberDictionaryEntryType();
            phPhoneNumber1.Key = pkPhoneKeyType;
            phPhoneNumber1.Value = phPhoneNumber;
            ContactItemType cnContact = new ContactItemType();
            cnContact.PhoneNumbers = new PhoneNumberDictionaryEntryType[1];
            cnContact.PhoneNumbers[0] = phPhoneNumber1;
            setItem.Item = phindexedField;
            setItem.Item1 = cnContact;
            return setItem;
        }
        public SetItemFieldType AddContactDifferancePhyicalAddress(String fiFieldIndex, DictionaryURIType diDictionaryURI, PhysicalAddressDictionaryEntryType paAddress)
        {
            SetItemFieldType setItem = new SetItemFieldType();
            PathToIndexedFieldType pa1indexedField = new PathToIndexedFieldType();
            pa1indexedField.FieldIndex = fiFieldIndex;
            pa1indexedField.FieldURI = diDictionaryURI;

            ContactItemType cnContact = new ContactItemType();
            cnContact.PhysicalAddresses = new PhysicalAddressDictionaryEntryType[1];
            cnContact.PhysicalAddresses[0] = paAddress;
            setItem.Item = pa1indexedField;
            setItem.Item1 = cnContact;
            return setItem;
        }
        public SetItemFieldType AddContactDifferanceExtendedProp(ExtendedPropertyType exProp)
        {
            ContactItemType cnContact = new ContactItemType();
            cnContact.ExtendedProperty = new ExtendedPropertyType[1];
            cnContact.ExtendedProperty[0] = exProp;
            SetItemFieldType setItem = new SetItemFieldType();
            setItem.Item = exProp.ExtendedFieldURI;
            setItem.Item1 = cnContact;
            return setItem;
        }
        public String AddNewContact(ContactItemType cnContact, BaseFolderIdType bfFolderid)
        {
            String rsReturnString = "";
            ItemIdType iiItemid = new ItemIdType();
            CreateItemType ciCreateItemRequest = new CreateItemType();
            ciCreateItemRequest.MessageDisposition = MessageDispositionType.SaveOnly;
            ciCreateItemRequest.MessageDispositionSpecified = true;
            ciCreateItemRequest.SavedItemFolderId = new TargetFolderIdType();
            ciCreateItemRequest.SavedItemFolderId.Item = bfFolderid;
            ciCreateItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            ciCreateItemRequest.Items.Items = new ItemType[1];
            ciCreateItemRequest.Items.Items[0] = cnContact;
            CreateItemResponseType createItemResponse = esb.CreateItem(ciCreateItemRequest);
            if (createItemResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Error)
            {
                rsReturnString = "Error";
                Console.WriteLine("Error Occured");
                Console.WriteLine(createItemResponse.ResponseMessages.Items[0].MessageText);

            }
            else
            {
                rsReturnString = "Success";
                ItemInfoResponseMessageType rmResponseMessage = createItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
                Console.WriteLine("Contact was created");
                // Console.WriteLine("Item ID : " + rmResponseMessage.Items.Items[0].ItemId.Id.ToString());
                //    Console.WriteLine("ChangeKey : " + rmResponseMessage.Items.Items[0].ItemId.ChangeKey.ToString());
                iiItemid.Id = rmResponseMessage.Items.Items[0].ItemId.Id.ToString();
                iiItemid.ChangeKey = rmResponseMessage.Items.Items[0].ItemId.ChangeKey.ToString();
            }
            return rsReturnString;

        }
        public String UpdateContact(List<SetItemFieldType> cnChanges, ItemIdType cnContactID)
        {
            String upUpdateresult = "";
            UpdateItemType updateItemType = new UpdateItemType();
            updateItemType.ConflictResolution = ConflictResolutionType.AlwaysOverwrite;
            updateItemType.MessageDisposition = MessageDispositionType.SaveOnly;
            updateItemType.MessageDispositionSpecified = true;
            updateItemType.ItemChanges = new ItemChangeType[1];
            ItemChangeType changeType = new ItemChangeType();

            changeType.Item = cnContactID;
            changeType.Updates = new ItemChangeDescriptionType[cnChanges.Count];

            for (int ucount = 0; ucount < cnChanges.Count; ucount++)
            {
                changeType.Updates[ucount] = cnChanges[ucount];
            }
            updateItemType.ItemChanges[0] = changeType;
            // Send the update item request and receive the response
            UpdateItemResponseType updateItemResponse = esb.UpdateItem(updateItemType);
            if (updateItemResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
            {
                Console.WriteLine("Update Successful");
                upUpdateresult = "Success";
            }
            else
            {
                Console.WriteLine(updateItemResponse.ResponseMessages.Items[0].MessageText.ToString());
                upUpdateresult = "Error";
            }
            return upUpdateresult;
        }
        public String UpdateFolderExtendedProperty(ExtendedPropertyType exProp, BaseFolderType Folder)
        {

            FolderType cfUpdateFolder = new FolderType();
            cfUpdateFolder.ExtendedProperty = new ExtendedPropertyType[1];
            cfUpdateFolder.ExtendedProperty[0] = exProp;

            UpdateFolderType upUpdateFolderRequest = new UpdateFolderType();
            FolderChangeType fcFolderchanges = new FolderChangeType();

            FolderIdType cfFolderid = new FolderIdType();
            cfFolderid.Id = Folder.FolderId.Id;
            cfFolderid.ChangeKey = Folder.FolderId.ChangeKey;
            fcFolderchanges.Item = cfFolderid;

            SetFolderFieldType dnDisplayNameChange = new SetFolderFieldType();
            dnDisplayNameChange.Item = exProp.ExtendedFieldURI;
            dnDisplayNameChange.Item1 = cfUpdateFolder;

            fcFolderchanges.Updates = new FolderChangeDescriptionType[1] { dnDisplayNameChange };
            upUpdateFolderRequest.FolderChanges = new FolderChangeType[1] { fcFolderchanges };
            String upres = "";
            UpdateFolderResponseType ufUpdateFolderResponse = esb.UpdateFolder(upUpdateFolderRequest);
            if (ufUpdateFolderResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
            {
                Console.WriteLine("Folder Updated sucessfully");
                upres = "Success";
            }
            else
            {
                Console.WriteLine("Folder Updated Error" + ufUpdateFolderResponse.ResponseMessages.Items[0].MessageText);
                upres = "Error";
            }
            return upres;
        }
        public String UpdateItem(SetItemFieldType cnChange, ItemIdType iiItemID)
        {
            String upUpdateresult = "";
            UpdateItemType updateItemType = new UpdateItemType();
            updateItemType.ConflictResolution = ConflictResolutionType.AlwaysOverwrite;
            updateItemType.MessageDisposition = MessageDispositionType.SaveOnly;
            updateItemType.MessageDispositionSpecified = true;
            updateItemType.ItemChanges = new ItemChangeType[1];
            ItemChangeType changeType = new ItemChangeType();

            changeType.Item = iiItemID;
            changeType.Updates = new ItemChangeDescriptionType[1];
            changeType.Updates[0] = cnChange;
          
            updateItemType.ItemChanges[0] = changeType;
            // Send the update item request and receive the response
            UpdateItemResponseType updateItemResponse = esb.UpdateItem(updateItemType);
            if (updateItemResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
            {
                Console.WriteLine("Update Successful");
                upUpdateresult = "Success";
            }
            else
            {
                Console.WriteLine(updateItemResponse.ResponseMessages.Items[0].MessageText.ToString());
                upUpdateresult = "Error";
            }
            return upUpdateresult;
        }
        public Hashtable GetAvailiblity(String[] emEmailAddresses, Duration fbDuration, int FbInterval)
        {
            Hashtable retList = new Hashtable();
            int itIntevalNum = DateTime.Compare(fbDuration.StartTime, fbDuration.EndTime);
            FreeBusyViewOptionsType fbViewOptions = new FreeBusyViewOptionsType();
            fbViewOptions.TimeWindow = fbDuration;
            fbViewOptions.RequestedView = FreeBusyViewType.DetailedMerged;
            fbViewOptions.RequestedViewSpecified = true;
            fbViewOptions.MergedFreeBusyIntervalInMinutes = FbInterval;
            fbViewOptions.MergedFreeBusyIntervalInMinutesSpecified = true;

            MailboxData[] mbMailboxes = new MailboxData[emEmailAddresses.Length];
            for (int fbc = 0; fbc < emEmailAddresses.Length; fbc++) {
                mbMailboxes[fbc] = new MailboxData();
                EmailAddress eaEmailAddress = new EmailAddress();
                eaEmailAddress.Address = emEmailAddresses[fbc];
                eaEmailAddress.Name = String.Empty;
                mbMailboxes[fbc].Email = eaEmailAddress;
                mbMailboxes[fbc].ExcludeConflicts = false;
            }
            String tzString = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Time Zones\" + TimeZone.CurrentTimeZone.StandardName;
            RegistryKey TziRegKey = Registry.LocalMachine;
            TziRegKey = TziRegKey.OpenSubKey(tzString);
            byte[] Tzival = (byte[])TziRegKey.GetValue("TZI");
            REG_TZI_FORMAT rtRegTimeZone = new REG_TZI_FORMAT();
            rtRegTimeZone.regget(Tzival);
            GetUserAvailabilityRequestType fbRequest = new GetUserAvailabilityRequestType();
            fbRequest.TimeZone = new SerializableTimeZone();
            fbRequest.TimeZone.DaylightTime = new SerializableTimeZoneTime();
            fbRequest.TimeZone.StandardTime = new SerializableTimeZoneTime();
            fbRequest.TimeZone.Bias = rtRegTimeZone.Bias;
            fbRequest.TimeZone.StandardTime.Bias = rtRegTimeZone.StandardBias;
            fbRequest.TimeZone.DaylightTime.Bias = rtRegTimeZone.DaylightBias;
            if (rtRegTimeZone.StandardDate.wMonth != 0)
            {
                fbRequest.TimeZone.StandardTime.DayOfWeek = ((DayOfWeek)rtRegTimeZone.StandardDate.wDayOfWeek).ToString();
                fbRequest.TimeZone.StandardTime.DayOrder = (short)rtRegTimeZone.StandardDate.wDay;
                fbRequest.TimeZone.StandardTime.Month = rtRegTimeZone.StandardDate.wMonth;
                fbRequest.TimeZone.StandardTime.Time = String.Format("{0:0#}:{1:0#}:{2:0#}", rtRegTimeZone.StandardDate.wHour, rtRegTimeZone.StandardDate.wMinute, rtRegTimeZone.StandardDate.wSecond);
            }
            else
            {
                fbRequest.TimeZone.StandardTime.DayOfWeek = "Sunday";
                fbRequest.TimeZone.StandardTime.DayOrder = 1;
                fbRequest.TimeZone.StandardTime.Month = 1;
                fbRequest.TimeZone.StandardTime.Time = "00:00:00";

            }
            if (rtRegTimeZone.DaylightDate.wMonth != 0)
            {
                fbRequest.TimeZone.DaylightTime.DayOfWeek = ((DayOfWeek)rtRegTimeZone.DaylightDate.wDayOfWeek).ToString();
                fbRequest.TimeZone.DaylightTime.DayOrder = (short)rtRegTimeZone.DaylightDate.wDay;
                fbRequest.TimeZone.DaylightTime.Month = rtRegTimeZone.DaylightDate.wMonth;
                fbRequest.TimeZone.DaylightTime.Time = "00:00:00";
            }
            else
            {
                fbRequest.TimeZone.DaylightTime.DayOfWeek = "Sunday";
                fbRequest.TimeZone.DaylightTime.DayOrder = 5;
                fbRequest.TimeZone.DaylightTime.Month = 12;
                fbRequest.TimeZone.DaylightTime.Time = "23:59:59";

            }
            fbRequest.MailboxDataArray = mbMailboxes;
            fbRequest.FreeBusyViewOptions = fbViewOptions;
            GetUserAvailabilityResponseType fbResponse = esb.GetUserAvailability(fbRequest);
            if (fbResponse.FreeBusyResponseArray != null) {
                System.TimeSpan ftsTimeSpan = fbDuration.EndTime - fbDuration.StartTime;
                double frspan = ftsTimeSpan.TotalMinutes / FbInterval;
                int tsseg = 0;
                for (DateTime htStartTime = fbDuration.StartTime; htStartTime < fbDuration.EndTime; htStartTime = htStartTime.AddMinutes(FbInterval))
                {
                    for (int mbNumCount = 0; mbNumCount < mbMailboxes.Length; mbNumCount++)
                    {
                        FBEntry fbEnt = new FBEntry();
                        fbEnt.MailboxEmailAddress = mbMailboxes[mbNumCount].Email.Address.ToString();
                        fbEnt.FBTime = htStartTime;
                        if (fbResponse.FreeBusyResponseArray[mbNumCount].FreeBusyView.MergedFreeBusy != null)
                        {
                            fbEnt.FBStatus = fbResponse.FreeBusyResponseArray[mbNumCount].FreeBusyView.MergedFreeBusy.Substring(tsseg, 1);
                            bool getsub = false;
                            switch (fbResponse.FreeBusyResponseArray[mbNumCount].FreeBusyView.MergedFreeBusy.Substring(tsseg, 1))
                            {
                                case "0": getsub = false;
                                    break;
                                case "1": getsub = true;
                                    break;
                                case "2": getsub = true;
                                    break;
                                case "3": getsub = true;
                                    break;
                                case "4": getsub = false;
                                    break;
                            }
                            if (getsub == true)
                            {
                                foreach (CalendarEvent calevent in fbResponse.FreeBusyResponseArray[mbNumCount].FreeBusyView.CalendarEventArray)
                                {
                                    if (calevent.CalendarEventDetails != null)
                                    {
                                        if (calevent.StartTime <= htStartTime & calevent.EndTime >= htStartTime)
                                        {
                                            if (fbEnt.FBSubject == "" | fbEnt.FBSubject == null)
                                            {
                                                fbEnt.FBSubject = calevent.CalendarEventDetails.Subject;
                                            }
                                            else {
                                                fbEnt.FBSubject = fbEnt.FBSubject.ToString() + " || " + calevent.CalendarEventDetails.Subject;
                                            }
                                            if (fbEnt.FBLocation == "" | fbEnt.FBLocation == null)
                                            {
                                                fbEnt.FBLocation = calevent.CalendarEventDetails.Location;
                                            }
                                            else {
                                                fbEnt.FBLocation = fbEnt.FBLocation.ToString() + " || " + calevent.CalendarEventDetails.Location;
                                            }
                                        }
                                    }

                                }
                            }
                        }
                        else {
                            fbEnt.FBStatus = "N/A";
                        }
                        if (retList.ContainsKey(fbEnt.MailboxEmailAddress))
                        {
                            Hashtable exHash = (Hashtable)retList[fbEnt.MailboxEmailAddress];
                            exHash.Add(fbEnt.FBTime.ToString("HH:mm"),fbEnt);
                        }
                        else {
                            Hashtable exHash = new Hashtable();
                            exHash.Add(fbEnt.FBTime.ToString("HH:mm"),fbEnt);
                            retList.Add(fbEnt.MailboxEmailAddress, exHash);
                        }
                       
                    }
                    tsseg++;
                }
            
            
            
            }

         return retList;
        }
        public String CreateTwitMail(BaseFolderIdType fiFolderID, DateTime stSentTime, String twTwitEmail, String twDisplayName, String twSubject)
        { 
            String rsReturnString = "";
            CreateItemType ciCreateItemRequest = new CreateItemType();
            ciCreateItemRequest.MessageDisposition = MessageDispositionType.SaveOnly;
            ciCreateItemRequest.MessageDispositionSpecified = true;
            ciCreateItemRequest.SavedItemFolderId = new TargetFolderIdType();
            ciCreateItemRequest.SavedItemFolderId.Item = fiFolderID;
            ciCreateItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            MessageType wsMessage = new MessageType();
            wsMessage.ToRecipients = new EmailAddressType[1];
            wsMessage.ToRecipients[0] = new EmailAddressType();
            wsMessage.ToRecipients[0].EmailAddress = twTwitEmail;
            wsMessage.ToRecipients[0].Name = twDisplayName;
            wsMessage.Subject = twSubject;
            wsMessage.Body = new BodyType();
            wsMessage.Body.BodyType1 = BodyTypeType.Text;
            wsMessage.Body.Value = twSubject;
            wsMessage.From = new SingleRecipientType();
            wsMessage.From.Item = new EmailAddressType();
            wsMessage.From.Item.EmailAddress = twTwitEmail;
            wsMessage.From.Item.Name = twDisplayName;
            ExtendedPropertyType msSentFlag = new ExtendedPropertyType();
            PathToExtendedFieldType epExPathmc = new PathToExtendedFieldType();
            epExPathmc.PropertyType = MapiPropertyTypeType.Integer;
            epExPathmc.PropertyTag = "0x0E07";
            msSentFlag.ExtendedFieldURI = epExPathmc;
            msSentFlag.Item = "0";
            ExtendedPropertyType msClientSubmit = new ExtendedPropertyType();
            PathToExtendedFieldType epExPathcs = new PathToExtendedFieldType();
            epExPathcs.PropertyType = MapiPropertyTypeType.SystemTime;
            epExPathcs.PropertyTag = "0x0039";
            msClientSubmit.ExtendedFieldURI = epExPathcs;
            msClientSubmit.Item = stSentTime.ToUniversalTime().ToString("yyyy-MM-ddThh:mm:ssZ");
            ExtendedPropertyType rsRecievedTime = new ExtendedPropertyType();
            PathToExtendedFieldType epExPathrt = new PathToExtendedFieldType();
            epExPathrt.PropertyType = MapiPropertyTypeType.SystemTime;
            epExPathrt.PropertyTag = "0x0E06";
            rsRecievedTime.ExtendedFieldURI = epExPathrt;
            rsRecievedTime.Item = stSentTime.ToUniversalTime().ToString("yyyy-MM-ddThh:mm:ssZ");
            wsMessage.ExtendedProperty = new ExtendedPropertyType[3];
            wsMessage.ExtendedProperty[0] = msClientSubmit;
            wsMessage.ExtendedProperty[1] = msSentFlag;
            wsMessage.ExtendedProperty[2] = rsRecievedTime;
            ciCreateItemRequest.Items.Items = new ItemType[1];
            ciCreateItemRequest.Items.Items[0] = wsMessage;
            CreateItemResponseType crCreateItemResponse = esb.CreateItem(ciCreateItemRequest);
                if (crCreateItemResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Error)
                {
                    rsReturnString = crCreateItemResponse.ResponseMessages.Items[0].MessageText.ToString();
                }
                else
                {
                    rsReturnString = "Message Created";
                }
            return rsReturnString;
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
                                    if (rmResponseMessage.Items.Items[0] is CalendarItemType)
                                    {
                                        if (bsBodyString.IndexOf("<div class=") >= 0 | bsBodyString.Length > 50)
                                        {
                                            xrXmlWritter.WriteCData("<html>" + bsBodyString.Substring(bsBodyString.IndexOf("<body>")));
                                        }
                                    }
                                    else {
                                        xrXmlWritter.WriteCData("<html>" + bsBodyString.Substring(bsBodyString.IndexOf("<body>")));
                                    }
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
    public class FBEntry {
        public String MailboxEmailAddress {
            get
            {
                return this.intMailboxEmailAddress;
            }
            set
            {
                intMailboxEmailAddress = value;
            }
        }
        private String intMailboxEmailAddress;
        public DateTime FBTime
        {
            get
            {
                return this.intFBTime;
            }
            set
            {
                intFBTime = value;
            }
        }
        private DateTime intFBTime;
        public String FBStatus
        {
            get
            {
                return this.intFBStatus;
            }
            set
            {
                intFBStatus = value;
            }
        }
        private String intFBStatus;
        public String FBSubject
        {
            get
            {
                return this.intFBSubject;
            }
            set
            {
                intFBSubject = value;
            }
        }
        private String intFBSubject;
        public String FBLocation
        {
            get
            {
                return this.intFBLocation;
            }
            set
            {
                intFBLocation = value;
            }
        }
        private String intFBLocation;
    
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
    class rsswriteException : Exception
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

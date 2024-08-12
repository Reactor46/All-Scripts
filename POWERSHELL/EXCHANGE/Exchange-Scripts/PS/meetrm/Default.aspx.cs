using System;
using System.Globalization;
using System.Data;
using System.Drawing;
using System.Configuration;
using System.Web;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Xml;
using System.Web.Security;
using Microsoft.Win32;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Collections;
using ResourceFind;
using System.Net;
using EWS;


public partial class _Default : System.Web.UI.Page 
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

    protected void Page_Load(object sender, EventArgs e)
    {
        string rvRowValue;
        // Define the Result Set Table to bind to ASP data controls

        String unUserName = "username";
        String pnPassWord = "password";
        String dnDomain = "domain";
        String snServerName = "servername";
        String stStartTime = "08:00:00";
        String etEndTime = "18:00:00";
        
        //Deal with Self Signed Certificate Errors
        ServicePointManager.ServerCertificateValidationCallback = delegate(Object obj, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
            {
                return true;
            };

        DataTable raDataTable = new DataTable();
        raDataTable.Columns.Add("Time");
        //Define FreeBusy Connection
        ExchangeServiceBinding ewsServiceBinding = new ExchangeServiceBinding();
        ewsServiceBinding.Credentials = new NetworkCredential(unUserName, pnPassWord ,dnDomain );
        ewsServiceBinding.Url = @"http://" + snServerName + "/EWS/exchange.asmx";
        Duration fbDuration = new Duration();
        fbDuration.StartTime = DateTime.ParseExact(DateTime.Now.ToString("yyyyMMdd") + "T" + stStartTime, "yyyyMMddTHH:mm:ss", null);
        fbDuration.EndTime = DateTime.ParseExact(DateTime.Now.ToString("yyyyMMdd") + "T" + etEndTime , "yyyyMMddTHH:mm:ss", null);
        int itIntevalNum = DateTime.Compare(fbDuration.StartTime, fbDuration.EndTime);
        FreeBusyViewOptionsType fbViewOptions = new FreeBusyViewOptionsType();
        fbViewOptions.TimeWindow = fbDuration;
        fbViewOptions.RequestedView = FreeBusyViewType.DetailedMerged;
        fbViewOptions.RequestedViewSpecified = true;
        fbViewOptions.MergedFreeBusyIntervalInMinutes = 30;
        fbViewOptions.MergedFreeBusyIntervalInMinutesSpecified = true;
        //find Resource and populate FreeBusy Request
        ResourceFind.Service rfResourceFinder = new ResourceFind.Service();
        XmlNode frFindRooms = rfResourceFinder.FindRooms();
        XmlNodeList mbMailboxsNodes = frFindRooms.SelectNodes("//Mailbox");
        Hashtable mbMailboxhash = new Hashtable();
        foreach (XmlNode vlNode in mbMailboxsNodes)
        {
           mbMailboxhash.Add(vlNode.Attributes.GetNamedItem("emailaddress").Value.ToString(), "");
           Hashtable mbDetails = new Hashtable();
           foreach(XmlNode vlChildNodes in vlNode.ChildNodes){
               mbDetails.Add(vlChildNodes.Name, vlChildNodes.InnerText.ToString());
           }
           raDataTable.Columns.Add((string)mbDetails["displayName"]);
         }
       
        int mbCount = 0;
        MailboxData[] mbMailboxes = new MailboxData[mbMailboxhash.Count];
        foreach (string emMailboxEmail in mbMailboxhash.Keys) {
            mbMailboxes[mbCount] = new MailboxData();
            EmailAddress eaEmailAddress = new EmailAddress();
            eaEmailAddress.Address = emMailboxEmail;
            eaEmailAddress.Name = String.Empty;
            mbMailboxes[mbCount].Email = eaEmailAddress;
            mbMailboxes[mbCount].ExcludeConflicts = false;
            mbCount++;
        }
        //Deal with timeZone in Request
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
        GetUserAvailabilityResponseType fbResponse = ewsServiceBinding.GetUserAvailability(fbRequest);
        System.TimeSpan ftsTimeSpan = fbDuration.EndTime - fbDuration.StartTime;
        double frspan = ftsTimeSpan.TotalMinutes / 30;
        int tsseg = 0;
        for (DateTime htStartTime = fbDuration.StartTime; htStartTime < fbDuration.EndTime; htStartTime = htStartTime.AddMinutes(30))
        {
            DataRow drDataRow = raDataTable.NewRow();
            drDataRow[0] = htStartTime.ToString("HH:mm");
            for (int mbNumCount = 0; mbNumCount < mbMailboxes.Length; mbNumCount++)
            {
                rvRowValue = fbResponse.FreeBusyResponseArray[mbNumCount].FreeBusyView.MergedFreeBusy.Substring(tsseg, 1);
                if (rvRowValue != "0") {
                    foreach (CalendarEvent calevent in fbResponse.FreeBusyResponseArray[mbNumCount].FreeBusyView.CalendarEventArray) {
                        System.TimeSpan tsTimeSpan = calevent.EndTime - calevent.StartTime;
                        double rspan = tsTimeSpan.TotalMinutes / 30;
                        if (htStartTime.ToString("HH:mm") == fbDuration.StartTime.ToString("HH:mm"))
                            {
                                if (rspan >= 48)
                                {
                                    if (calevent.CalendarEventDetails != null)
                                    {
                                        if (calevent.CalendarEventDetails.Subject != "")
                                        {
                                            rvRowValue = frspan + " " + calevent.CalendarEventDetails.Subject;
                                        }
                                        else {
                                            rvRowValue = frspan + " Occupied";
                                        }

                                    }
                                    else
                                    {
                                        rvRowValue = frspan + " Occupied";
                                    }
                                    
                                }
                                else if (calevent.StartTime < fbDuration.StartTime)
                                {
                                    System.TimeSpan stTimeSpan = fbDuration.StartTime - calevent.StartTime;
                                    double stspan = stTimeSpan.TotalMinutes / 30;
                                    if (calevent.CalendarEventDetails != null)
                                    {
                                        if (calevent.CalendarEventDetails.Subject != "")
                                        {
                                            rvRowValue = (rspan - stspan) + " " + calevent.CalendarEventDetails.Subject;
                                        }
                                        else {
                                            rvRowValue = (rspan - stspan) + " Occupied";
                                        }
                                    }
                                    else
                                    {
                                        rvRowValue = (rspan - stspan) + " Occupied";
                                    }
                                    
                                }
                            }                      
            
                        if (htStartTime == calevent.StartTime)
                        {
                            if (calevent.CalendarEventDetails != null)
                            {
                                if (calevent.CalendarEventDetails.Subject != "")
                                {
                                    rvRowValue = rspan.ToString() + " " + calevent.CalendarEventDetails.Subject;
                                }
                                else {
                                    rvRowValue = rspan.ToString() + " Occupied";
                                }
                            }
                            else {
                                   rvRowValue = rspan.ToString() + " Occupied";
                            }

                        }  
               
                    }
                
                }
                drDataRow[mbNumCount + 1] = rvRowValue;
            }
            raDataTable.Rows.Add(drDataRow);
            tsseg++;
        }
        dgDataGrid.DataSource = raDataTable;
        dgDataGrid.DataBind();
   
  
         

    }
    public void dgDataGrid_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            
            e.Row.Cells[0].BackColor = Color.LightGray;
            e.Row.Cells[0].Font.Bold = true;
            for (int cnColNumber = 1; cnColNumber < e.Row.Cells.Count;cnColNumber++ ){
                e.Row.Cells[cnColNumber].Width = 150;
                if (e.Row.Cells[cnColNumber].Text.Substring(0,1) != "0") {
                    if ((cnColNumber % 2) == 0)
                    {
                        e.Row.Cells[cnColNumber].BackColor = Color.LightGreen;
                    }
                    else {
                        e.Row.Cells[cnColNumber].BackColor = Color.LightBlue;
                    }
                    if (e.Row.Cells[cnColNumber].Text.Length > 2)
                    {
                        string rspan = e.Row.Cells[cnColNumber].Text.Substring(0, 2).ToString();
                        e.Row.Cells[cnColNumber].Text = e.Row.Cells[cnColNumber].Text.Substring(2, (e.Row.Cells[cnColNumber].Text.Length - 2));                   
                        e.Row.Cells[cnColNumber].RowSpan = Convert.ToInt16(rspan);
                    }
                    else
                    {
                        e.Row.Cells[cnColNumber].Visible = false;
                        e.Row.Cells[cnColNumber].Text = " ";
                    }
                }
                else {e.Row.Cells[cnColNumber].Text = " ";
                }
            } 
        }
    }

}

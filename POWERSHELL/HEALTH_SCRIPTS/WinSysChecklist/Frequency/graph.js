
	var chart;
	var SystemData =[
    {
        "FirstTimeWritten":  "12/16/2017 3:34:11 AM",
        "Count":  11115,
        "EventID":  36874,
        "color":  "#457986",
        "LastTimeWritten":  "@{TimeWritten=12/16/2017 3:34:35 AM}",
        "Source":  "Schannel",
        "EntryType":  1,
        "Message":  "An TLS 1.0 connection request was received from a remote client application, but none of the cipher suites supported by the client application are supported by the server. The TLS connection request has failed."
    },
    {
        "FirstTimeWritten":  "12/21/2017 9:22:37 AM",
        "Count":  1119,
        "EventID":  10028,
        "color":  "#851818",
        "LastTimeWritten":  "@{TimeWritten=12/21/2017 9:22:58 AM}",
        "Source":  "DCOM",
        "EntryType":  1,
        "Message":  "The description for Event ID \u002710028\u0027 in Source \u0027DCOM\u0027 cannot be found.  The local computer may not have the necessary registry information or message DLL files to display the message, or you may not have permission to access them.  The following information is part of the event:\u00278.8.8.8\u0027, \u0027    29b4\u0027, \u0027C:\\windows\\system32\\dcdiag.exe\u0027"
    },
    {
        "FirstTimeWritten":  "12/21/2017 1:47:00 PM",
        "Count":  296,
        "EventID":  10016,
        "color":  "#851818",
        "LastTimeWritten":  "@{TimeWritten=12/21/2017 3:38:39 PM}",
        "Source":  "DCOM",
        "EntryType":  1,
        "Message":  "The description for Event ID \u002710016\u0027 in Source \u0027DCOM\u0027 cannot be found.  The local computer may not have the necessary registry information or message DLL files to display the message, or you may not have permission to access them.  The following information is part of the event:\u0027application-specific\u0027, \u0027Local\u0027, \u0027Activation\u0027, \u0027{D63B10C5-BB46-4990-A94F-E40B9D520160}\u0027, \u0027{9CA88EE3-ACB7-47C8-AFC4-AB702511C276}\u0027, \u0027NT AUTHORITY\u0027, \u0027SYSTEM\u0027, \u0027S-1-5-18\u0027, \u0027LocalHost (Using LRPC)\u0027, \u0027Unavailable\u0027, \u0027Unavailable\u0027"
    },
    {
        "FirstTimeWritten":  "12/16/2017 10:13:22 PM",
        "Count":  129,
        "EventID":  4,
        "color":  "#457986",
        "LastTimeWritten":  "@{TimeWritten=12/16/2017 10:21:31 PM}",
        "Source":  "Kerberos",
        "EntryType":  1,
        "Message":  "The Kerberos client received a KRB_AP_ERR_MODIFIED error from the server lashafile01n01$. The target name used was cifs/Lasfs02. This indicates that the target server failed to decrypt the ticket provided by the client. This can occur when the target server principal name (SPN) is registered on an account other than the account the target service is using. Ensure that the target SPN is only registered on the account used by the server. This error can also happen if the target service account password is different than what is configured on the Kerberos Key Distribution Center for that target service. Ensure that the service on the server and the KDC are both configured to use the same password. If the server name is not fully qualified, and the target domain (Contoso.CORP) is different from the client domain (Contoso.CORP), check if there are identically named server accounts in these two domains, or use the fully-qualified name to identify the server."
    },
    {
        "FirstTimeWritten":  "12/14/2017 1:59:11 PM",
        "Count":  66,
        "EventID":  1130,
        "color":  "#AE3B3B",
        "LastTimeWritten":  "@{TimeWritten=12/17/2017 1:45:26 AM}",
        "Source":  "Microsoft-Windows-GroupPolicy",
        "EntryType":  1,
        "Message":  "0 failed. \r\n\tGPO Name : Security Configuration for CASH\r\n\tGPO File System Path : \\\\Contoso.corp\\SysVol\\Contoso.corp\\Policies\\{14846A6E-2442-4DEC-BFF7-689EF944C310}\\Machine\r\n\tScript Name: add_URL_Reservation_for_Avaya_one-X.bat"
    },
    {
        "FirstTimeWritten":  "12/17/2017 1:45:25 AM",
        "Count":  62,
        "EventID":  7000,
        "color":  "#114652",
        "LastTimeWritten":  "@{TimeWritten=12/18/2017 10:38:31 PM}",
        "Source":  "Service Control Manager",
        "EntryType":  1,
        "Message":  "The CldFlt service failed to start due to the following error: \r\n%%50"
    },
    {
        "FirstTimeWritten":  "12/16/2017 11:34:02 PM",
        "Count":  38,
        "EventID":  139,
        "color":  "#275F6C",
        "LastTimeWritten":  "@{TimeWritten=12/16/2017 11:34:03 PM}",
        "Source":  "Mup",
        "EntryType":  2,
        "Message":  "{Delayed Write Failed}\r\nWindows was unable to save all the data for the file \\\\Contosocorp\\share\\user\\jbattista\\Outlook-Archive\\~archive.pst.tmp; the data has been lost.\r\nThis error may be caused by network connectivity issues. Please try to save this file elsewhere."
    },
    {
        "FirstTimeWritten":  "12/11/2017 11:16:47 AM",
        "Count":  27,
        "EventID":  14,
        "color":  "#AE743B",
        "LastTimeWritten":  "@{TimeWritten=12/12/2017 11:17:23 AM}",
        "Source":  "Kerberos",
        "EntryType":  2,
        "Message":  "The password stored in Credential Manager is invalid. This might be caused by the logged on user changing the password from this computer or a different computer. To resolve this error, open Credential Manager in Control Panel, and reenter the password for the credential ContosoCORP\\cbrown."
    },
    {
        "FirstTimeWritten":  "12/17/2017 1:45:25 AM",
        "Count":  24,
        "EventID":  5719,
        "color":  "#FFABAB",
        "LastTimeWritten":  "@{TimeWritten=12/18/2017 10:38:31 PM}",
        "Source":  "NETLOGON",
        "EntryType":  1,
        "Message":  "This computer was not able to set up a secure session with a domain\r\ncontroller in domain ContosoCORP due to the following: \r\n%We can\u0027t sign you in with this credential because your domain isn\u0027t available. Make sure your device is connected to your organization\u0027s network and try again. If you previously signed in on this device with another credential, you can sign in with that credential.\r\n\r\nThis may lead to authentication problems. Make sure that this\r\ncomputer is connected to the network. If the problem persists,\r\nplease contact your domain administrator.\r\n\r\n\r\n\r\nADDITIONAL INFO\r\n\r\nIf this computer is a domain controller for the specified domain, it\r\nsets up the secure session to the primary domain controller emulator in the specified\r\ndomain. Otherwise, this computer sets up the secure session to any domain controller\r\nin the specified domain."
    },
    {
        "FirstTimeWritten":  "12/17/2017 1:45:26 AM",
        "Count":  19,
        "EventID":  10154,
        "color":  "#275F6C",
        "LastTimeWritten":  "@{TimeWritten=12/18/2017 10:38:31 PM}",
        "Source":  "WinRM",
        "EntryType":  2,
        "Message":  "The description for Event ID \u0027468906\u0027 in Source \u0027WinRM\u0027 cannot be found.  The local computer may not have the necessary registry information or message DLL files to display the message, or you may not have permission to access them.  The following information is part of the event:\u0027WSMAN/DPC-20890.Contoso.corp\u0027, \u0027WSMAN/DPC-20890\u0027, \u00271355\u0027"
    },
    {
        "FirstTimeWritten":  "12/14/2017 1:59:11 PM",
        "Count":  19,
        "EventID":  1129,
        "color":  "#136B13",
        "LastTimeWritten":  "@{TimeWritten=12/17/2017 1:45:26 AM}",
        "Source":  "Microsoft-Windows-GroupPolicy",
        "EntryType":  1,
        "Message":  "The processing of Group Policy failed because of lack of network connectivity to a domain controller. This may be a transient condition. A success message would be generated once the machine gets connected to the domain controller and Group Policy has successfully processed. If you do not see a success message for several hours, then contact your administrator."
    },
    {
        "FirstTimeWritten":  "11/28/2017 9:30:40 AM",
        "Count":  17,
        "EventID":  7034,
        "color":  "#AE3B3B",
        "LastTimeWritten":  "@{TimeWritten=12/7/2017 10:17:42 AM}",
        "Source":  "Service Control Manager",
        "EntryType":  1,
        "Message":  "The Netwrix Account Lockout Examiner service terminated unexpectedly.  It has done this 1 time(s)."
    },
    {
        "FirstTimeWritten":  "11/3/2017 5:01:00 AM",
        "Count":  15,
        "EventID":  1085,
        "color":  "#FFD5AB",
        "LastTimeWritten":  "@{TimeWritten=11/3/2017 5:01:10 AM}",
        "Source":  "Microsoft-Windows-GroupPolicy",
        "EntryType":  2,
        "Message":  "Windows failed to apply the Group Policy Printers settings. Group Policy Printers settings might have its own log file. Please click on the \"More information\" link."
    },
    {
        "FirstTimeWritten":  "12/14/2017 1:58:45 PM",
        "Count":  15,
        "EventID":  10149,
        "color":  "#FFABAB",
        "LastTimeWritten":  "@{TimeWritten=12/17/2017 1:44:59 AM}",
        "Source":  "WinRM",
        "EntryType":  2,
        "Message":  "The description for Event ID \u0027468901\u0027 in Source \u0027WinRM\u0027 cannot be found.  The local computer may not have the necessary registry information or message DLL files to display the message, or you may not have permission to access them.  The following information is part of the event:"
    },
    {
        "FirstTimeWritten":  "12/21/2017 3:31:31 PM",
        "Count":  14,
        "EventID":  1112,
        "color":  "#2F8B2F",
        "LastTimeWritten":  "@{TimeWritten=12/21/2017 3:32:03 PM}",
        "Source":  "Microsoft-Windows-GroupPolicy",
        "EntryType":  2,
        "Message":  "The Group Policy Client Side Extension Folder Redirection was unable to apply one or more settings because the changes must be processed before system startup or user logon. The system will wait for Group Policy processing to finish completely before the next startup or logon for this user, and this may result in slow startup and boot performance."
    },
    {
        "FirstTimeWritten":  "12/2/2017 1:45:43 PM",
        "Count":  14,
        "EventID":  1014,
        "color":  "#AE3B3B",
        "LastTimeWritten":  "@{TimeWritten=12/5/2017 9:19:46 AM}",
        "Source":  "Microsoft-Windows-DNS-Client",
        "EntryType":  2,
        "Message":  "Name resolution for the name laspshost timed out after none of the configured DNS servers responded."
    },
    {
        "FirstTimeWritten":  "12/18/2017 6:19:17 AM",
        "Count":  10,
        "EventID":  129,
        "color":  "#FFABAB",
        "LastTimeWritten":  "@{TimeWritten=12/18/2017 11:01:14 AM}",
        "Source":  "Microsoft-Windows-Time-Service",
        "EntryType":  2,
        "Message":  "NtpClient was unable to set a domain peer to use as a time source because of discovery error. NtpClient will try again in 15 minutes and double the reattempt interval thereafter. The error was: The entry is not found. (0x800706E1)"
    },
    {
        "FirstTimeWritten":  "8/30/2017 9:47:23 AM",
        "Count":  10,
        "EventID":  7023,
        "color":  "#275E6C",
        "LastTimeWritten":  "@{TimeWritten=8/30/2017 9:47:24 AM}",
        "Source":  "Service Control Manager",
        "EntryType":  1,
        "Message":  "The Interactive Services Detection service terminated with the following error: \r\n%%1"
    },
    {
        "FirstTimeWritten":  "12/11/2017 10:38:56 PM",
        "Count":  9,
        "EventID":  7043,
        "color":  "#FFD5AB",
        "LastTimeWritten":  "@{TimeWritten=12/18/2017 10:37:14 PM}",
        "Source":  "Service Control Manager",
        "EntryType":  1,
        "Message":  "The Update Orchestrator Service service did not shut down properly after receiving a preshutdown control."
    },
    {
        "FirstTimeWritten":  "12/2/2017 1:10:56 PM",
        "Count":  9,
        "EventID":  6038,
        "color":  "#275F6C",
        "LastTimeWritten":  "@{TimeWritten=12/19/2017 9:29:53 PM}",
        "Source":  "LsaSrv",
        "EntryType":  2,
        "Message":  "Microsoft Windows Server has detected that NTLM authentication is presently being used between clients and this server. This event occurs once per boot of the server on the first time a client uses NTLM with this server.\r\n \r\nNTLM is a weaker authentication mechanism. Please check:\r\n \r\n      Which applications are using NTLM authentication?\r\n      Are there configuration issues preventing the use of stronger authentication such as Kerberos authentication?\r\n      If NTLM must be supported, is Extended Protection configured?\r\n \r\nDetails on how to complete these checks can be found at http://go.microsoft.com/fwlink/?LinkId=225699."
    },
    {
        "FirstTimeWritten":  "12/2/2017 4:54:51 PM",
        "Count":  8,
        "EventID":  6008,
        "color":  "#D86D6D",
        "LastTimeWritten":  "@{TimeWritten=12/2/2017 5:07:55 PM}",
        "Source":  "EventLog",
        "EntryType":  1,
        "Message":  "The previous system shutdown at 4:54:51 PM on ‎12/‎2/‎2017 was unexpected."
    },
    {
        "FirstTimeWritten":  "9/23/2017 10:06:23 AM",
        "Count":  6,
        "EventID":  7030,
        "color":  "#FFD5AB",
        "LastTimeWritten":  "@{TimeWritten=9/29/2017 4:51:11 PM}",
        "Source":  "Service Control Manager",
        "EntryType":  1,
        "Message":  "The Track-It! Workstation Manager service is marked as an interactive service.  However, the system is configured to not allow interactive services.  This service may not function properly."
    },
    {
        "FirstTimeWritten":  "12/16/2017 11:33:49 PM",
        "Count":  5,
        "EventID":  50,
        "color":  "#275F6C",
        "LastTimeWritten":  "@{TimeWritten=12/18/2017 5:55:17 AM}",
        "Source":  "Microsoft-Windows-Time-Service",
        "EntryType":  2,
        "Message":  "The time service detected a time difference of greater than 5000 milliseconds for 900 seconds. The time difference might be caused by synchronization with low-accuracy time sources or by suboptimal network conditions. The time service is no longer synchronized and cannot provide the time to other clients or update the system clock. When a valid time stamp is received from a time service provider, the time service will correct itself."
    },
    {
        "FirstTimeWritten":  "@{TimeWritten=8/8/2017 4:51:04 PM}",
        "Count":  5,
        "EventID":  51,
        "color":  "#275F6C",
        "LastTimeWritten":  "@{TimeWritten=8/8/2017 4:51:04 PM}",
        "Source":  "Disk",
        "EntryType":  2,
        "Message":  "An error was detected on device \\Device\\Harddisk1\\DR1 during a paging operation."
    },
    {
        "FirstTimeWritten":  "10/25/2017 3:51:56 PM",
        "Count":  4,
        "EventID":  2004,
        "color":  "#136B13",
        "LastTimeWritten":  "@{TimeWritten=11/3/2017 9:59:58 AM}",
        "Source":  "Microsoft-Windows-Resource-Exhaustion-Detector",
        "EntryType":  2,
        "Message":  "Windows successfully diagnosed a low virtual memory condition. The following programs consumed the most virtual memory: powershell.exe (23100) consumed 60897222656 bytes, explorer.exe (16828) consumed 837713920 bytes, and VISIO.EXE (17580) consumed 536875008 bytes."
    },
    {
        "FirstTimeWritten":  "9/24/2017 2:10:54 AM",
        "Count":  4,
        "EventID":  10006,
        "color":  "#57AD57",
        "LastTimeWritten":  "@{TimeWritten=10/13/2017 12:31:31 PM}",
        "Source":  "DCOM",
        "EntryType":  1,
        "Message":  "The description for Event ID \u002710006\u0027 in Source \u0027DCOM\u0027 cannot be found.  The local computer may not have the necessary registry information or message DLL files to display the message, or you may not have permission to access them.  The following information is part of the event:\u00272147943515\u0027, \u0027LASMT04\u0027, \u0027{8BC3F05E-D86B-11D0-A075-00C04FB68820}\u0027"
    },
    {
        "FirstTimeWritten":  "9/20/2017 8:36:17 PM",
        "Count":  4,
        "EventID":  29,
        "color":  "#FFABAB",
        "LastTimeWritten":  "@{TimeWritten=12/1/2017 6:02:12 PM}",
        "Source":  "Microsoft-Windows-Kernel-Boot",
        "EntryType":  1,
        "Message":  "The description for Event ID \u002729\u0027 in Source \u0027Microsoft-Windows-Kernel-Boot\u0027 cannot be found.  The local computer may not have the necessary registry information or message DLL files to display the message, or you may not have permission to access them.  The following information is part of the event:\u00273221225684\u0027, \u00271202832\u0027"
    },
    {
        "FirstTimeWritten":  "12/15/2017 9:41:11 AM",
        "Count":  4,
        "EventID":  5009,
        "color":  "#AE3B3B",
        "LastTimeWritten":  "@{TimeWritten=12/15/2017 9:51:11 AM}",
        "Source":  "WAS",
        "EntryType":  2,
        "Message":  "A process serving application pool \u0027DefaultAppPool\u0027 terminated unexpectedly. The process id was \u00279528\u0027. The process exit code was \u00270xffffffff\u0027."
    },
    {
        "FirstTimeWritten":  "12/2/2017 1:46:08 PM",
        "Count":  4,
        "EventID":  27,
        "color":  "#AE3B3B",
        "LastTimeWritten":  "@{TimeWritten=12/2/2017 1:46:40 PM}",
        "Source":  "e1dexpress",
        "EntryType":  2,
        "Message":  "Intel(R) Ethernet Connection (5) I219-LM\r\n\r\nNetwork link is disconnected.\r\n"
    },
    {
        "FirstTimeWritten":  "9/22/2017 1:14:59 AM",
        "Count":  4,
        "EventID":  10010,
        "color":  "#275F6C",
        "LastTimeWritten":  "@{TimeWritten=12/11/2017 10:37:58 PM}",
        "Source":  "DCOM",
        "EntryType":  1,
        "Message":  "The description for Event ID \u002710010\u0027 in Source \u0027DCOM\u0027 cannot be found.  The local computer may not have the necessary registry information or message DLL files to display the message, or you may not have permission to access them.  The following information is part of the event:\u0027{9BA05972-F6A8-11CF-A442-00A0C90A8F39}\u0027"
    },
    {
        "FirstTimeWritten":  "9/27/2017 10:27:59 AM",
        "Count":  4,
        "EventID":  4227,
        "color":  "#AE3B3B",
        "LastTimeWritten":  "@{TimeWritten=9/27/2017 10:44:17 AM}",
        "Source":  "Tcpip",
        "EntryType":  2,
        "Message":  "TCP/IP failed to establish an outgoing connection because the selected local endpoint\r\nwas recently used to connect to the same remote endpoint. This error typically occurs\r\nwhen outgoing connections are opened and closed at a high rate, causing all available\r\nlocal ports to be used and forcing TCP/IP to reuse a local port for an outgoing connection.\r\nTo minimize the risk of data corruption, the TCP/IP standard requires a minimum time period\r\nto elapse between successive connections from a given local endpoint to a given remote endpoint."
    },
    {
        "FirstTimeWritten":  "@{TimeWritten=8/1/2017 1:08:11 PM}",
        "Count":  3,
        "EventID":  98,
        "color":  "#AE3B3B",
        "LastTimeWritten":  "8/1/2017 4:27:56 PM",
        "Source":  "Microsoft-Windows-Ntfs",
        "EntryType":  2,
        "Message":  "The description for Event ID \u002798\u0027 in Source \u0027Microsoft-Windows-Ntfs\u0027 cannot be found.  The local computer may not have the necessary registry information or message DLL files to display the message, or you may not have permission to access them.  The following information is part of the event:\u0027D:\u0027, \u0027\\Device\\HarddiskVolume5\u0027, \u00271\u0027"
    },
    {
        "FirstTimeWritten":  "9/20/2017 8:36:32 PM",
        "Count":  3,
        "EventID":  8033,
        "color":  "#014801",
        "LastTimeWritten":  "@{TimeWritten=10/12/2017 5:45:15 PM}",
        "Source":  "Microsoft-Windows-DNS-Client",
        "EntryType":  2,
        "Message":  "The system failed to update and remove host (A or AAAA) resource records (RRs) for network adapter\r\nwith settings:\r\n\r\n\r\n          Adapter Name : {7F374DBC-798E-4857-B261-5AD734716601}\r\n\r\n          Host Name : DPC-20890\r\n\r\n          Primary Domain Suffix : Contoso.corp\r\n\r\n          DNS server list :\r\n\r\n            \t192.168.11.91, 192.168.11.92\r\n\r\n          Sent update to server : \u003c?\u003e\r\n\r\n          IP Address(es) :\r\n\r\n            192.168.17.151\r\n\r\n\r\n        The system could not remove these host (A or AAAA) RRs because the update request timed out while awaiting a response from the DNS server. This is probably because the DNS server authoritative for the zone where these RRs need to be updated is either not currently running or reachable on the network."
    },
    {
        "FirstTimeWritten":  "@{TimeWritten=8/1/2017 1:08:14 PM}",
        "Count":  2,
        "EventID":  219,
        "color":  "#5A0101",
        "LastTimeWritten":  "8/1/2017 4:30:11 PM",
        "Source":  "Microsoft-Windows-Kernel-PnP",
        "EntryType":  2,
        "Message":  "The driver \\Driver\\WudfRd failed to load for the device SWD\\WPDBUSENUM\\{6c6dcffd-7719-11e7-a32f-806e6f6e6963}#0000000000100000."
    },
    {
        "FirstTimeWritten":  "9/29/2017 8:17:08 AM",
        "Count":  2,
        "EventID":  15301,
        "color":  "#AE3B3B",
        "LastTimeWritten":  "@{TimeWritten=10/23/2017 7:37:36 AM}",
        "Source":  "HTTP",
        "EntryType":  2,
        "Message":  "The description for Event ID \u0027-2147468347\u0027 in Source \u0027HTTP\u0027 cannot be found.  The local computer may not have the necessary registry information or message DLL files to display the message, or you may not have permission to access them.  The following information is part of the event:\u0027\u0027, \u00270.0.0.0:4094\u0027"
    },
    {
        "FirstTimeWritten":  "8/1/2017 1:07:41 PM",
        "Count":  2,
        "EventID":  1055,
        "color":  "#014801",
        "LastTimeWritten":  "@{TimeWritten=11/1/2017 2:40:23 PM}",
        "Source":  "Microsoft-Windows-GroupPolicy",
        "EntryType":  1,
        "Message":  "The processing of Group Policy failed. Windows could not resolve the computer name. This could be caused by one of more of the following: \r\na) Name Resolution failure on the current domain controller. \r\nb) Active Directory Replication Latency (an account created on another domain controller has not replicated to the current domain controller)."
    },
    {
        "FirstTimeWritten":  "12/4/2017 8:02:07 AM",
        "Count":  2,
        "EventID":  1058,
        "color":  "#D86D6D",
        "LastTimeWritten":  "@{TimeWritten=12/4/2017 8:02:42 AM}",
        "Source":  "Microsoft-Windows-GroupPolicy",
        "EntryType":  1,
        "Message":  "The processing of Group Policy failed. Windows attempted to read the file \\\\Contoso.corp\\SysVol\\Contoso.corp\\Policies\\{EF859132-7D5D-41BD-8716-F401A6E4787A}\\gpt.ini from a domain controller and was not successful. Group Policy settings may not be applied until this event is resolved. This issue may be transient and could be caused by one or more of the following: \r\na) Name Resolution/Network Connectivity to the current domain controller. \r\nb) File Replication Service Latency (a file created on another domain controller has not replicated to the current domain controller). \r\nc) The Distributed File System (DFS) client has been disabled."
    },
    {
        "FirstTimeWritten":  "11/15/2017 11:21:52 AM",
        "Count":  2,
        "EventID":  7031,
        "color":  "#AE3B3B",
        "LastTimeWritten":  "@{TimeWritten=11/15/2017 11:21:56 AM}",
        "Source":  "Service Control Manager",
        "EntryType":  1,
        "Message":  "The BMC Client Management Agent service terminated unexpectedly.  It has done this 1 time(s).  The following corrective action will be taken in 60000 milliseconds: Restart the service."
    },
    {
        "FirstTimeWritten":  "10/12/2017 5:45:08 PM",
        "Count":  2,
        "EventID":  7016,
        "color":  "#851818",
        "LastTimeWritten":  "@{TimeWritten=11/29/2017 12:36:18 PM}",
        "Source":  "Service Control Manager",
        "EntryType":  1,
        "Message":  "The BMC Client Management Agent service has reported an invalid current state 0."
    },
    {
        "FirstTimeWritten":  "12/2/2017 1:45:47 PM",
        "Count":  2,
        "EventID":  8015,
        "color":  "#5A0101",
        "LastTimeWritten":  "@{TimeWritten=12/2/2017 1:45:59 PM}",
        "Source":  "Microsoft-Windows-DNS-Client",
        "EntryType":  2,
        "Message":  "The system failed to register host (A or AAAA) resource records (RRs) for network adapter\r\nwith settings:\r\n\r\n\r\n          Adapter Name : {7F374DBC-798E-4857-B261-5AD734716601}\r\n\r\n          Host Name : DPC-20890\r\n\r\n          Primary Domain Suffix : Contoso.corp\r\n\r\n          DNS server list :\r\n\r\n            \t192.168.11.91, 192.168.11.92\r\n\r\n          Sent update to server : \u003c?\u003e\r\n\r\n          IP Address(es) :\r\n\r\n            192.168.17.151\r\n\r\nThe reason the system could not register these RRs was because the update request it sent to the DNS server timed out. The most likely cause of this is that the DNS server authoritative for the name it was attempting to register or update is not running at this time.\r\n\r\nYou can manually retry DNS registration of the network adapter and its settings by typing \u0027ipconfig /registerdns\u0027 at the command prompt. If problems still persist, contact your DNS server or network systems administrator."
    },
    {
        "FirstTimeWritten":  "@{TimeWritten=11/1/2017 2:40:53 PM}",
        "Count":  2,
        "EventID":  8016,
        "color":  "#5A2D01",
        "LastTimeWritten":  "@{TimeWritten=11/1/2017 2:40:53 PM}",
        "Source":  "Microsoft-Windows-DNS-Client",
        "EntryType":  2,
        "Message":  "The system failed to register host (A or AAAA) resource records (RRs) for network adapter\r\nwith settings:\r\n\r\n\r\n          Adapter Name : {7F374DBC-798E-4857-B261-5AD734716601}\r\n\r\n          Host Name : DPC-20890\r\n\r\n          Primary Domain Suffix : Contoso.corp\r\n\r\n          DNS server list :\r\n\r\n            \t8.8.8.8, 8.8.4.4\r\n\r\n          Sent update to server : \u003c?\u003e\r\n\r\n          IP Address(es) :\r\n\r\n            192.168.17.151\r\n\r\nThe reason the system could not register these RRs was because the DNS server failed the update request. The most likely cause of this is that the authoritative DNS server required to process this update request has a lock in place on the zone, probably because a zone transfer is in progress.\r\n\r\nYou can manually retry DNS registration of the network adapter and its settings by typing \u0027ipconfig /registerdns\u0027 at the command prompt. If problems still persist, contact your DNS server or network systems administrator."
    },
    {
        "FirstTimeWritten":  "@{TimeWritten=10/23/2017 7:56:51 AM}",
        "Count":  1,
        "EventID":  15300,
        "color":  "#AE3B3B",
        "LastTimeWritten":  "@{TimeWritten=10/23/2017 7:56:51 AM}",
        "Source":  "HTTP",
        "EntryType":  2,
        "Message":  "The description for Event ID \u0027-2147468348\u0027 in Source \u0027HTTP\u0027 cannot be found.  The local computer may not have the necessary registry information or message DLL files to display the message, or you may not have permission to access them.  The following information is part of the event:\u0027\u0027, \u00270.0.0.0:4094\u0027"
    },
    {
        "FirstTimeWritten":  "@{TimeWritten=11/3/2017 9:59:58 AM}",
        "Count":  1,
        "EventID":  40960,
        "color":  "#AE743B",
        "LastTimeWritten":  "@{TimeWritten=11/3/2017 9:59:58 AM}",
        "Source":  "LsaSrv",
        "EntryType":  2,
        "Message":  "The Security System detected an authentication error for the server ldap/LASDC01.Contoso.corp. The failure code from authentication protocol Kerberos was \"Insufficient system resources exist to complete the API.\r\n (0xc000009a)\"."
    },
    {
        "FirstTimeWritten":  "@{TimeWritten=10/23/2017 7:24:19 AM}",
        "Count":  1,
        "EventID":  7024,
        "color":  "#022E38",
        "LastTimeWritten":  "@{TimeWritten=10/23/2017 7:24:19 AM}",
        "Source":  "Service Control Manager",
        "EntryType":  1,
        "Message":  "The Apache2.2 service terminated with the following service-specific error: \r\n%%1"
    },
    {
        "FirstTimeWritten":  "@{TimeWritten=12/2/2017 1:45:26 PM}",
        "Count":  1,
        "EventID":  10400,
        "color":  "#D86D6D",
        "LastTimeWritten":  "@{TimeWritten=12/2/2017 1:45:26 PM}",
        "Source":  "Microsoft-Windows-NDIS",
        "EntryType":  2,
        "Message":  "The network interface \"Intel(R) Ethernet Connection (5) I219-LM\" has begun resetting.  There will be a momentary disruption in network connectivity while the hardware resets.\r\nReason: 3.\r\nThis network interface has reset 1 time(s) since it was last initialized."
    },
    {
        "FirstTimeWritten":  "@{TimeWritten=12/14/2017 1:57:57 PM}",
        "Count":  1,
        "EventID":  1073,
        "color":  "#FFABAB",
        "LastTimeWritten":  "@{TimeWritten=12/14/2017 1:57:57 PM}",
        "Source":  "User32",
        "EntryType":  2,
        "Message":  "The attempt by user ContosoCORP\\jbattista to restart/shutdown computer DPC-20890 failed"
    },
    {
        "FirstTimeWritten":  "@{TimeWritten=8/14/2017 11:05:01 AM}",
        "Count":  1,
        "EventID":  20,
        "color":  "#275F6C",
        "LastTimeWritten":  "@{TimeWritten=8/14/2017 11:05:01 AM}",
        "Source":  "Microsoft-Windows-WindowsUpdateClient",
        "EntryType":  1,
        "Message":  "Installation Failure: Windows failed to install the following update with error 0x80070002: Windows Alarms \u0026 Clock."
    },
    {
        "FirstTimeWritten":  "@{TimeWritten=12/15/2017 9:39:31 AM}",
        "Count":  1,
        "EventID":  5011,
        "color":  "#014801",
        "LastTimeWritten":  "@{TimeWritten=12/15/2017 9:39:31 AM}",
        "Source":  "WAS",
        "EntryType":  2,
        "Message":  "A process serving application pool \u0027DefaultAppPool\u0027 suffered a fatal communication error with the Windows Process Activation Service. The process id was \u00272036\u0027. The data field contains the error number."
    },
    {
        "FirstTimeWritten":  "@{TimeWritten=12/18/2017 6:03:49 AM}",
        "Count":  1,
        "EventID":  24,
        "color":  "#D8A26D",
        "LastTimeWritten":  "@{TimeWritten=12/18/2017 6:03:49 AM}",
        "Source":  "Microsoft-Windows-Time-Service",
        "EntryType":  2,
        "Message":  "Time Provider NtpClient: No valid response has been received from domain controller LASDC05.Contoso.corp after 8 attempts to contact it. This domain controller will be discarded as a time source and NtpClient will attempt to discover a new domain controller from which to synchronize. The error was: The peer is unreachable. "
    },
    {
        "FirstTimeWritten":  "@{TimeWritten=8/1/2017 1:05:04 PM}",
        "Count":  1,
        "EventID":  4097,
        "color":  "#014801",
        "LastTimeWritten":  "@{TimeWritten=8/1/2017 1:05:04 PM}",
        "Source":  "NetJoin",
        "EntryType":  1,
        "Message":  "The machine DPC-20890 attempted to join the domain Contoso.corp but failed. The error code was 1332."
    }
];
	var appData =[
    {
        "FirstTimeWritten":  "12/21/2017 3:32:08 PM",
        "Count":  2360,
        "EventID":  4098,
        "color":  "#D86D6D",
        "LastTimeWritten":  "@{TimeWritten=12/21/2017 4:25:15 PM}",
        "Source":  "Group Policy Files",
        "EntryType":  2,
        "Message":  "The user \u0027backgroundDefault.jpg\u0027 preference item in the \u0027TemporaryLogonScreen {5EC51E94-AECC-434C-940B-C6187151ACD9}\u0027 Group Policy Object did not apply because it failed with error code \u00270x80070002 The system cannot find the file specified.\u0027 This error was suppressed."
    },
    {
        "FirstTimeWritten":  "12/17/2017 1:20:33 AM",
        "Count":  164,
        "EventID":  3,
        "color":  "#FFABAB",
        "LastTimeWritten":  "@{TimeWritten=12/17/2017 1:26:24 AM}",
        "Source":  "Lync",
        "EntryType":  1,
        "Message":  "Lync was unable to resolve the DNS hostname of the login server sipinternal.creditone.com.\r\n\r\n\r\n\r\nResolution:\r\n\r\nIf you are using manual configuration for Communicator, please check that the server name is typed correctly and in full.  If you are using automatic configuration, the network administrator will need to double-check the DNS A record configuration for sipinternal.creditone.com because it could not be resolved."
    },
    {
        "FirstTimeWritten":  "12/20/2017 12:19:47 AM",
        "Count":  46,
        "EventID":  1001,
        "color":  "#AE3B3B",
        "LastTimeWritten":  "@{TimeWritten=12/20/2017 12:27:39 AM}",
        "Source":  "MsiInstaller",
        "EntryType":  2,
        "Message":  "Detection of product \u0027{90160000-0051-0000-1000-0000000FF1CE}\u0027, feature \u0027VisioCore\u0027 failed during request for component \u0027{D92DAA7B-7DBB-4471-A93A-275AB0D24B20}\u0027"
    },
    {
        "FirstTimeWritten":  "12/20/2017 12:19:47 AM",
        "Count":  46,
        "EventID":  1004,
        "color":  "#FFABAB",
        "LastTimeWritten":  "@{TimeWritten=12/20/2017 12:27:39 AM}",
        "Source":  "MsiInstaller",
        "EntryType":  2,
        "Message":  "Detection of product \u0027{90160000-0051-0000-1000-0000000FF1CE}\u0027, feature \u0027ProductFiles\u0027, component \u0027{77586F20-86BA-4A4F-8A47-10B34A263C86}\u0027 failed.  The resource \u0027HKEY_CLASSES_ROOT(64)\\CLSID\\{000C0126-0000-0000-C000-000000000046}\\\u0027 does not exist."
    },
    {
        "FirstTimeWritten":  "12/20/2017 12:27:39 AM",
        "Count":  31,
        "EventID":  10023,
        "color":  "#014801",
        "LastTimeWritten":  "@{TimeWritten=12/20/2017 12:35:32 AM}",
        "Source":  "Windows Search Service",
        "EntryType":  2,
        "Message":  "The protocol host process 11492 did not respond and is being forcibly terminated {filter host process 6376}. \n"
    },
    {
        "FirstTimeWritten":  "12/20/2017 4:37:16 AM",
        "Count":  25,
        "EventID":  0,
        "color":  "#D8A26D",
        "LastTimeWritten":  "@{TimeWritten=12/21/2017 4:17:14 AM}",
        "Source":  "Office 2016 Licensing Service",
        "EntryType":  1,
        "Message":  "The description for Event ID \u00270\u0027 in Source \u0027Office 2016 Licensing Service\u0027 cannot be found.  The local computer may not have the necessary registry information or message DLL files to display the message, or you may not have permission to access them.  The following information is part of the event:\u0027Subscription licensing service failed: -1073418220\u0027"
    },
    {
        "FirstTimeWritten":  "11/29/2017 2:13:01 PM",
        "Count":  21,
        "EventID":  106,
        "color":  "#FFABAB",
        "LastTimeWritten":  "@{TimeWritten=11/29/2017 2:13:12 PM}",
        "Source":  "MSExchange Common",
        "EntryType":  2,
        "Message":  "Performance counter updating error. Counter name is Percentage of Failed Offline GLS Requests in Last Minute, category name is MSExchange Global Locator OfflineGLS Processes. Optional code: 3. Exception: System.InvalidOperationException: The requested Performance Counter is not a custom counter, it has to be initialized as ReadOnly.\r\n   at System.Diagnostics.PerformanceCounter.InitializeImpl()\r\n   at System.Diagnostics.PerformanceCounter.get_RawValue()\r\n   at Microsoft.Exchange.Diagnostics.ExPerformanceCounter.set_RawValue(Int64 value)"
    },
    {
        "FirstTimeWritten":  "12/20/2017 10:48:40 AM",
        "Count":  20,
        "EventID":  1000,
        "color":  "#5A2D01",
        "LastTimeWritten":  "@{TimeWritten=12/21/2017 9:45:17 AM}",
        "Source":  "Application Error",
        "EntryType":  1,
        "Message":  "Faulting application name: mstsc.exe, version: 10.0.15063.674, time stamp: 0x57950623\r\nFaulting module name: mstsc.exe, version: 10.0.15063.674, time stamp: 0x57950623\r\nException code: 0xc0000005\r\nFault offset: 0x000000000000c065\r\nFaulting process id: 0x25fc\r\nFaulting application start time: 0x01d37a836f980adc\r\nFaulting application path: C:\\windows\\system32\\mstsc.exe\r\nFaulting module path: C:\\windows\\system32\\mstsc.exe\r\nReport Id: 246891ab-2a6c-421f-ad3f-e77eb0798329\r\nFaulting package full name: \r\nFaulting package-relative application ID: "
    },
    {
        "FirstTimeWritten":  "12/21/2017 3:31:31 PM",
        "Count":  9,
        "EventID":  510,
        "color":  "#136B13",
        "LastTimeWritten":  "@{TimeWritten=12/21/2017 3:32:03 PM}",
        "Source":  "Microsoft-Windows-Folder Redirection",
        "EntryType":  2,
        "Message":  "The folder redirection policy hasn\u0027t been applied. It\u0027s been delayed until the next sign-in because the group policy sign-in optimization is in effect."
    },
    {
        "FirstTimeWritten":  "12/18/2017 8:30:34 PM",
        "Count":  8,
        "EventID":  36,
        "color":  "#5A0101",
        "LastTimeWritten":  "@{TimeWritten=12/20/2017 9:18:29 PM}",
        "Source":  "Outlook",
        "EntryType":  2,
        "Message":  "Search cannot complete the indexing of your Outlook data. Indexing cannot continue for U:\\Outlook-Archive\\archive.pst (error=0x81940804). If this error continues, contact Microsoft Support."
    },
    {
        "FirstTimeWritten":  "12/11/2017 9:51:16 AM",
        "Count":  8,
        "EventID":  11,
        "color":  "#AE743B",
        "LastTimeWritten":  "@{TimeWritten=12/15/2017 12:38:40 PM}",
        "Source":  "Lync",
        "EntryType":  2,
        "Message":  "A SIP request made by Lync failed in an unexpected manner (status code 0). More information is contained in the following technical data:\r\n\r\n\r\n\r\nRequestUri:   sip:james.adkins@creditone.com;opaque=user:epid:FtUHhVuejVmDLF6n1yidSwAA;gruu\r\nFrom:         sip:john.battista@creditone.com;tag=9a1c015fa0\r\nTo:           sip:james.adkins@creditone.com;tag=43d0dde4b4\r\nCall-ID:      e7cd67d8808f43e6b10dbcd4cc582525\r\nContent-type: application/sdp;call-type=im\r\n\r\nv=0\r\no=- 0 0 IN IP4 192.168.17.151\r\ns=session\r\nc=IN IP4 192.168.17.151\r\nt=0 0\r\nm=message 5060 sip null\r\na=accept-types:text/plain multipart/alternative image/gif text/rtf text/html application/x-ms-ink application/ms-imdn+xml text/x-msmsgsinvite \r\n\r\n\r\nResponse Data:\r\n\r\n101  Progress Report\r\nms-diagnostics:  13004;reason=\"Request was proxied to one or more registered endpoints\";source=\"LASSKYPE01.Contoso.CORP\";Count=\"1\";appName=\"InboundRouting\"\r\n\r\n\r\n0  (null)\r\n(null):  51004; reason=\"Action initiated by user\";OriginalPresenceState=\"3000\";CurrentPresenceState=\"3000\";MeInsideUser=\"Yes\";ConversationInitiatedBy=\"6\";SourceNetwork=\"2\";RemotePartyCanDoIM=\"Yes\"\r\n\r\n\r\n\r\n\r\nResolution:\r\n\r\nIf this error continues to occur, please contact your network administrator. The network administrator can use a tool like winerror.exe from the Windows Resource Kit or lcserror.exe from the Office Communications Server Resource Kit in order to interpret any error codes listed above."
    },
    {
        "FirstTimeWritten":  "12/18/2017 8:36:35 AM",
        "Count":  6,
        "EventID":  1002,
        "color":  "#022E38",
        "LastTimeWritten":  "@{TimeWritten=12/21/2017 8:45:01 AM}",
        "Source":  "Application Hang",
        "EntryType":  1,
        "Message":  "The program iexplore.exe version 11.0.15063.608 stopped interacting with Windows and was closed. To see if more information about the problem is available, check the problem history in the Security and Maintenance control panel.\r\n\r\nProcess ID: 371c\r\n\r\nStart Time: 01d379f74104b7e0\r\n\r\nTermination Time: 15\r\n\r\nApplication Path: C:\\Program Files (x86)\\Internet Explorer\\iexplore.exe\r\n\r\nReport Id: c0c38544-548f-46e9-aeaf-5fcf7b2b8c5a\r\n\r\nFaulting package full name: \r\n\r\nFaulting package-relative application ID: \r\n"
    },
    {
        "FirstTimeWritten":  "12/16/2017 8:08:47 PM",
        "Count":  6,
        "EventID":  10010,
        "color":  "#136B13",
        "LastTimeWritten":  "@{TimeWritten=12/18/2017 10:06:04 PM}",
        "Source":  "Microsoft-Windows-RestartManager",
        "EntryType":  2,
        "Message":  "Application \u0027C:\\Program Files (x86)\\Internet Explorer\\iexplore.exe\u0027 (pid 10596) cannot be restarted - 1."
    },
    {
        "FirstTimeWritten":  "12/15/2017 10:18:34 AM",
        "Count":  5,
        "EventID":  1008,
        "color":  "#5A0101",
        "LastTimeWritten":  "@{TimeWritten=12/17/2017 12:27:38 AM}",
        "Source":  "Perflib",
        "EntryType":  1,
        "Message":  "The description for Event ID \u0027-1073740816\u0027 in Source \u0027Perflib\u0027 cannot be found.  The local computer may not have the necessary registry information or message DLL files to display the message, or you may not have permission to access them.  The following information is part of the event:\u0027BITS\u0027, \u0027C:\\Windows\\System32\\bitsperf.dll\u0027, \u00278\u0027"
    },
    {
        "FirstTimeWritten":  "12/14/2017 1:03:37 PM",
        "Count":  5,
        "EventID":  1026,
        "color":  "#136B13",
        "LastTimeWritten":  "@{TimeWritten=12/20/2017 10:48:37 AM}",
        "Source":  ".NET Runtime",
        "EntryType":  1,
        "Message":  "Application: mmc.exe\nFramework Version: v4.0.30319\nDescription: The process was terminated due to an unhandled exception.\nException Info: exception code c0000005, exception address 00007FFF00DB34BD\nStack:\n"
    },
    {
        "FirstTimeWritten":  "@{TimeWritten=12/4/2017 8:02:07 AM}",
        "Count":  1,
        "EventID":  472,
        "color":  "#5A2D01",
        "LastTimeWritten":  "@{TimeWritten=12/4/2017 8:02:07 AM}",
        "Source":  "ESENT",
        "EntryType":  2,
        "Message":  "taskhostw (7412) WebCacheLocal: The shadow header page of file C:\\Users\\jbattista\\AppData\\Local\\Microsoft\\Windows\\WebCache\\WebCacheV01.dat was damaged. The primary header page (32768 bytes) was used instead."
    },
    {
        "FirstTimeWritten":  "@{TimeWritten=12/15/2017 5:07:48 AM}",
        "Count":  1,
        "EventID":  1534,
        "color":  "#AE3B3B",
        "LastTimeWritten":  "@{TimeWritten=12/15/2017 5:07:48 AM}",
        "Source":  "Microsoft-Windows-User Profiles Service",
        "EntryType":  2,
        "Message":  "Profile notification of event Create for component {2c86c843-77ae-4284-9722-27d65366543c} failed, error code is Not implemented\r\n. \r\n\r\n"
    },
    {
        "FirstTimeWritten":  "@{TimeWritten=12/7/2017 11:33:42 AM}",
        "Count":  1,
        "EventID":  1023,
        "color":  "#AE3B3B",
        "LastTimeWritten":  "@{TimeWritten=12/7/2017 11:33:42 AM}",
        "Source":  "Perflib",
        "EntryType":  1,
        "Message":  "The description for Event ID \u0027-1073740801\u0027 in Source \u0027Perflib\u0027 cannot be found.  The local computer may not have the necessary registry information or message DLL files to display the message, or you may not have permission to access them.  The following information is part of the event:\u0027rdyboost\u0027, \u00274\u0027"
    },
    {
        "FirstTimeWritten":  "@{TimeWritten=12/7/2017 9:48:59 AM}",
        "Count":  1,
        "EventID":  636,
        "color":  "#5A2D01",
        "LastTimeWritten":  "@{TimeWritten=12/7/2017 9:48:59 AM}",
        "Source":  "ESENT",
        "EntryType":  2,
        "Message":  "DllHost (8212) Internet_NOEDP_LEGACY_IDB: Flush map file \"C:\\Users\\jbattista\\AppData\\Local\\Microsoft\\Internet Explorer\\Indexed DB\\Internet.jfm\" will be deleted. Reason: ReadHdrFailed."
    },
    {
        "FirstTimeWritten":  "@{TimeWritten=12/7/2017 9:48:59 AM}",
        "Count":  1,
        "EventID":  640,
        "color":  "#2F8B2F",
        "LastTimeWritten":  "@{TimeWritten=12/7/2017 9:48:59 AM}",
        "Source":  "ESENT",
        "EntryType":  2,
        "Message":  "DllHost (8212) Internet_NOEDP_LEGACY_IDB: Error -1919 validating header page on flush map file \"C:\\Users\\jbattista\\AppData\\Local\\Microsoft\\Internet Explorer\\Indexed DB\\Internet.jfm\". The flush map file will be invalidated.\r\n\r\nAdditional information: [SignDbHdrFromDb:Create time:00/00/1900 00:00:00.000 Rand:0 Computer:] [SignFmHdrFromDb:Create time:00/00/1900 00:00:00.000 Rand:0 Computer:] [SignDbHdrFromFm:Create time:11/27/2017 19:26:19.555 Rand:247133080 Computer:] [SignFmHdrFromFm:Create time:11/27/2017 19:31:33.070 Rand:405386050 Computer:]"
    },
    {
        "FirstTimeWritten":  "@{TimeWritten=12/21/2017 4:03:06 PM}",
        "Count":  1,
        "EventID":  1015,
        "color":  "#851818",
        "LastTimeWritten":  "@{TimeWritten=12/21/2017 4:03:06 PM}",
        "Source":  "MsiInstaller",
        "EntryType":  2,
        "Message":  "Failed to connect to server. Error: 0x800401F0"
    }
];;
        AmCharts.ready(function () {
		var chart = AmCharts.makeChart("chartdiv",{
			"type": "serial",
			"dataProvider": SystemData,
			"categoryField": "EventID",
			"startDuration": 1,
			//axes
			"valueAxes": [ {
				"dashLength": 5,
				"title": "Frecuency of the event",
				"axisAlpha": 0,
			}],
			"gridAboveGraphs": false,
			
			"graphs": [ {
				"balloonText": "EventID [[category]]</br>Repeated: <b>[[value]]</b> times</br>Source: [[Source]]</br>[[Message]]</br>First on:<b>[[FirstTimeWritten]]</b></br>Last on:<b>[[LastTimeWritten]]</b> </br> <b class=Yellow>[[EntryType]]</b>",
				"fillAlphas": 0.8,
				"lineAlpha": 0.2,
				"type": "column",
				"valueField": "Count",
				"colorField": "color"
			}],
			"chartCursor": {
				"categoryBalloonEnabled": false,
				"cursorAlpha": 0,
				"zoomable": false
			},
			
			"categoryAxis": {
				"gridPosition": "start",
				"gridAlpha": 0,
				"fillAlpha": 1,
				"labelRotation" : 60,
				"fillColor": "#EEEEEE",
				"gridPosition": "start"
			},
			"creditsPosition" : "top-right",
			"export": {
				"enabled": true
			}
    });

		var chart2 = AmCharts.makeChart("chart2div",{
			"type": "serial",
			"dataProvider":appData,
			"categoryField": "EventID",
			"startDuration": 1,
			//axes
			"valueAxes": [ {
				"dashLength": 5,
				"title": "Frecuency of the event",
				"axisAlpha": 0,
			}],
			"gridAboveGraphs": false,
			
			"graphs": [ {
				"balloonText": "EventID [[category]]</br>Repeated: <b>[[value]]</b> times</br>Source: [[Source]]</br>[[Message]]</br>First on:<b>[[FirstTimeWritten]]</b></br>Last on:<b>[[LastTimeWritten]]</b> </br> <b class=Yellow>[[EntryType]]</b>",
				"fillAlphas": 0.8,
				"lineAlpha": 0.2,
				"type": "column",
				"valueField": "Count",
				"colorField": "color"
			}],
			"chartCursor": {
				"categoryBalloonEnabled": false,
				"cursorAlpha": 0,
				"zoomable": false
			},
			
			"categoryAxis": {
				"gridPosition": "start",
				"gridAlpha": 0,
				"fillAlpha": 1,
				"labelRotation" : 60,
				"fillColor": "#EEEEEE",
				"gridPosition": "start"
			},
			"creditsPosition" : "top-right",
			"export": {
				"enabled": true
			}
    });
    
			//Original
		/*
        // SERIAL CHART
        chart = new AmCharts.AmSerialChart();
        chart.dataProvider = SystemData;
        chart.categoryField = "EventID";
        chart.startDuration = 1;


        // AXES
        // category
        var categoryAxis = chart.categoryAxis;
        categoryAxis.labelRotation = 60; // this line makes category values to be rotated
        categoryAxis.gridAlpha = 0;
        categoryAxis.fillAlpha = 1;
        categoryAxis.fillColor = "#EEEEEE";
        categoryAxis.gridPosition = "start";

        // value
        var valueAxis = new AmCharts.ValueAxis();
        valueAxis.dashLength = 5;
        valueAxis.title = "Frecuency of the event";
        valueAxis.axisAlpha = 0;
        chart.addValueAxis(valueAxis);

        // GRAPH
        var graph = new AmCharts.AmGraph();
        graph.valueField = "Count";
        graph.colorField = "color";
        graph.balloonText = "<b>[[category]]: [[value]]</b>";
        graph.type = "column";
        graph.lineAlpha = 0;
        graph.fillAlphas = 1;
		
        chart.addGraph(graph);

        // CURSOR
        var chartCursor = new AmCharts.ChartCursor();
        chartCursor.cursorAlpha = 0;
        chartCursor.zoomable = false;
        chartCursor.categoryBalloonEnabled = false;
        chart.addChartCursor(chartCursor);

        chart.creditsPosition = "top-right";

        // WRITE
        chart.write("chartdiv");
		*/
});

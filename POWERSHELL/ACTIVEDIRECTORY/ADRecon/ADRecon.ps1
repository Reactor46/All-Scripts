﻿<#

.SYNOPSIS
    ADRecon is a tool which gathers information about the Active Directory and generates a report which can provide a holistic picture of the current state of the target AD environment.

.DESCRIPTION

    ADRecon is a tool which extracts various artifacts (as highlighted below) out of an AD environment in a specially formatted Microsoft Excel report that includes summary views with metrics to facilitate analysis.
    The report can provide a holistic picture of the current state of the target AD environment.
    The tool is useful to various classes of security professionals like auditors, DIFR, students, administrators, etc. It can also be an invaluable post-exploitation tool for a penetration tester.
    It can be run from any workstation that is connected to the environment even hosts that are not domain members. Furthermore, the tool can be executed in the context of a non-privileged (i.e. standard domain user) accounts. Fine Grained Password Policy, LAPS and BitLocker may require Privileged user accounts.
    The tool will use Microsoft Remote Server Administration Tools (RSAT) if available, otherwise it will communicate with the Domain Controller using LDAP.
    The following information is gathered by the tool:
    - Forest;
    - Domains in the Forest and other attributes such as Sites;
    - Domain Password Policy;
    - Domain Controllers and their roles;
    - Users and their attributes;
    - Service Principal Names;
    - Groups and memberships;
    - Organizational Units and their ACLs;
    - Group Policy Object details;
    - DNS Zones and Records;
    - Printers;
    - Computers and their attributes;
    - LAPS passwords (if implemented);
    - BitLocker Recovery Keys (if implemented); and
    - Domain GPO Report (requires RSAT).

    Author     : Prashant Mahajan
    Company    : https://www.senseofsecurity.com.au

.NOTES

    The following commands can be used to turn off ExecutionPolicy: (Requires Admin Privs)

    PS > $ExecPolicy = Get-ExecutionPolicy
    PS > Set-ExecutionPolicy bypass
    PS > .\ADRecon.ps1
    PS > Set-ExecutionPolicy $ExecPolicy

    OR

    Start the PowerShell as follows:
    powershell.exe -ep bypass

    OR

    Already have a PowerShell open ?
    PS > $Env:PSExecutionPolicyPreference = 'Bypass'

    OR

    powershell.exe -nologo -executionpolicy bypass -noprofile -file ADRecon.ps1

.PARAMETER Protocol
	Which protocol to use; ADWS (default) or LDAP

.PARAMETER DomainController
	Domain Controller IP Address or Domain FQDN.

.PARAMETER Credential
	Domain Credentials.

.PARAMETER GenExcel
	Path for ADRecon output folder containing the CSV files to generate the ADRecon-Report.xlsx. Use it to generate the ADRecon-Report.xlsx when Microsoft Excel is not installed on the host used to run ADRecon.

.PARAMETER OutputDir
	Path for ADRecon output folder to save the CSV files and the ADRecon-Report.xlsx. (The folder specified will be created if it doesn't exist)

.PARAMETER Collect
    What attributes to collect; Comma separated; e.g Forest,Domain (Default all)
    Valid values include: Forest, Domain, PasswordPolicy, DCs, Users, UserSPNs, Groups, GroupMembers, OUs, OUPermissions, GPOs, GPOReport, DNSZones, Printers, Computers, ComputerSPNs, LAPS, BitLocker.

.PARAMETER OutputType
    Output Type; Comma seperated; e.g CSV,STDOUT,Excel (Default STDOUT with -Collect parameter, else CSV and Excel).
    Valid values include: STDOUT, CSV, Excel.

.PARAMETER DormantTimeSpan
    Timespan for Dormant accounts. (Default 90 days)

.PARAMETER PageSize
    The PageSize to set for the LDAP searcher object.

.PARAMETER Threads
    The number of threads to use during processing objects. (Default 10)

.PARAMETER FlushCount
    The number of processed objects which will be flushed to disk. (Default -1; Flush after all objects are processed).

.PARAMETER Log
    Create ADRecon Log using Start-Transcript

.EXAMPLE

	.\ADRecon.ps1 -GenExcel C:\ADRecon-Report-<timestamp>
    [*] ADRecon <version> by Prashant Mahajan (@prashant3535) from Sense of Security.
    [*] Generating ADRecon-Report.xlsx
    [+] Excelsheet Saved to: C:\ADRecon-Report-<timestamp>\<domain>-ADRecon-Report.xlsx

.EXAMPLE

	.\ADRecon.ps1 -DomainController <IP or FQDN> -Credential <domain\username>
    [*] ADRecon <version> by Prashant Mahajan (@prashant3535) from Sense of Security.
	[*] Running on <domain>\<hostname> - Member Workstation
    <snip>

    Example output from Domain Member with Alternate Credentials.

.EXAMPLE

	.\ADRecon.ps1 -DomainController <IP or FQDN> -Credential <domain\username> -Collect DCs -OutputType Excel
    [*] ADRecon <version> by Prashant Mahajan (@prashant3535) from Sense of Security.
    [*] Running on WORKGROUP\<hostname> - Standalone Workstation
    [*] Commencing - <timestamp>
    [-] Domain Controllers
    [*] Total Execution Time (mins): <minutes>
    [*] Generating ADRecon-Report.xlsx
    [+] Excelsheet Saved to: C:\ADRecon-Report-<timestamp>\<domain>-ADRecon-Report.xlsx
    [*] Completed.
    [*] Output Directory: C:\ADRecon-Report-<timestamp>

    Example output from from a Non-Member using RSAT to only enumerate Domain Controllers.

.EXAMPLE

    .\ADRecon.ps1 -Protocol ADWS -DomainController <IP or FQDN> -Credential <domain\username>
    [*] ADRecon <version> by Prashant Mahajan (@prashant3535) from Sense of Security.
    [*] Running on WORKGROUP\<hostname> - Standalone Workstation
    [*] Commencing - <timestamp>
    [-] Domain
    [-] Forest
    [-] Default Password Policy
    [-] Fine Grained Password Policy - May need a Privileged Account
    [-] Domain Controllers
    [-] Domain Users - May take some time
    [*] Total Users: <number>
    [-] Domain User SPNs
    [-] Domain Groups - May take some time
    [*] Total Groups: <number>
    [-] Domain Group Memberships - May take some time
    [*] Total GroupMember Objects: <number>
    [-] Domain OrganizationalUnits
    [*] Total OUs: <number>
    [-] Domain OrganizationalUnits Permissions - May take some time
    [-] Domain GPOs
    [*] Total GPOs: <number>
    [-] Domain DNS Zones and Records
    [*] Total DNS Zones: <number>
    [-] Domain Printers
    [*] Total Printers: <number>
    [-] Domain Computers - May take some time
    [*] Total Computers: <number>
    [-] Domain Computer SPNs
    [-] LAPS - Needs Privileged Account
    [*] LAPS is not implemented.
    [-] BitLocker Recovery Keys - Needs Privileged Account
    [-] Domain GPO Report - May take some time
    [EXCEPTION] Current security context is not associated with an Active Directory domain or forest.
    [*] Run the tool using RUNAS.
    [*] runas /user:<Domain FQDN>\<Username> /netonly powershell.exe
    [*] Total Execution Time (mins): <minutes>
    [*] Generating ADRecon-Report.xlsx
    [+] Excelsheet Saved to: C:\ADRecon-Report-<timestamp>\<domain>-ADRecon-Report.xlsx
    [*] Completed.
    [*] Output Directory: C:\ADRecon-Report-<timestamp>

    Example output from a Non-Member using RSAT.

.EXAMPLE

    .\ADRecon.ps1 -Protocol LDAP -DomainController <IP or FQDN> -Credential <domain\username>
    [*] ADRecon <version> by Prashant Mahajan (@prashant3535) from Sense of Security.
    Standalone Workstation
    WORKGROUP
    [*] LDAP bind Successful
    [*] Commencing - <timstamp>
    [-] Domain
    [-] Forest
    [-] Default Password Policy
    [-] Fine Grained Password Policy - May need a Privileged Account
    [-] Domain Controllers
    [-] Domain Users - May take some time
    [*] Calculating if the user Cannot Change Password
    [*] Total Users: <number>
    [-] Domain User SPNs
    [-] Domain Groups - May take some time
    [*] Total Groups: <number>
    [-] Domain Group Memberships - May take some time
    [*] Total GroupMember Objects: <number>
    [-] Domain OrganizationalUnits
    [*] Total OUs: <number>
    [-] Domain OrganizationalUnits Permissions - May take some time
    [-] Domain GPOs
    [*] Total GPOs: <number>
    [-] Domain DNS Zones and Records
    [*] Total DNS Zones: <number>
    [-] Domain Printers
    [*] Total Printers: <number>
    [-] Domain Computers - May take some time
    [*] Total Computers: <number>
    [-] Domain Computer SPNs
    [-] LAPS - Needs Privileged Account
    [*] LAPS is not implemented.
    [-] BitLocker Recovery Keys - Needs Privileged Account
    [-] Domain GPO Report - May take some time
    [*] Currently, the module is only supported with ADWS.
    [*] Total Execution Time (mins): <timestamp>
    [*] Generating ADRecon-Report.xlsx
    [+] Excelsheet Saved to: C:\ADRecon-Report-<timestamp>\<domain>-ADRecon-Report.xlsx
    [*] Completed.
    [*] Output Directory: C:\ADRecon-Report-<timestamp>

    Example output from a Non-Member using LDAP.

.LINK
https://github.com/sense-of-security/ADRecon
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory = $false, HelpMessage = "Which protocol to use; ADWS (default) or LDAP.")]
    [string] $Protocol = 'ADWS',

    [Parameter(Mandatory = $false, HelpMessage = "Domain Controller IP Address or Domain FQDN.")]
    [string] $DomainController,

    [Parameter(Mandatory = $false, HelpMessage = "Domain Credentials.")]
    [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

    [Parameter(Mandatory = $false, HelpMessage = "Path for ADRecon output folder containing the CSV files to generate the ADRecon-Report.xlsx. Use it to generate the ADRecon-Report.xlsx when Microsoft Excel is not installed on the host used to run ADRecon.")]
    [string] $GenExcel,

    [Parameter(Mandatory = $false, HelpMessage = "Path for ADRecon output folder to save the CSV files and the ADRecon-Report.xlsx. (The folder specified will be created if it doesn't exist)")]
    [string] $OutputDir,

    [Parameter(Mandatory = $false, HelpMessage = "What attributes to collect; Comma separated; e.g Forest,Domain (Default all) Valid values include: Forest, Domain, PasswordPolicy, DCs, Users, UserSPNs, Groups, GroupMembers, OUs, OUPermissions, GPOs, GPOReport, DNSZones, Printers, Computers, ComputerSPNs, LAPS, BitLocker")]
    [ValidateSet('Forest', 'Domain', 'PasswordPolicy', 'DCs', 'Users', 'UserSPNs', 'Groups', 'GroupMembers', 'OUs', 'OUPermissions', 'GPOs', 'GPOReport', 'DNSZones', 'Printers', 'Computers', 'ComputerSPNs', 'LAPS', 'BitLocker', 'Default')]
    [array] $Collect = 'Default',

    [Parameter(Mandatory = $false, HelpMessage = "Output type; Comma seperated; e.g CSV,STDOUT,Excel (Default STDOUT with -Collect parameter, else CSV and Excel)")]
    [ValidateSet('STDOUT', 'CSV', 'EXCEL', 'Default')]
    [array] $OutputType = 'Default',

    [Parameter(Mandatory = $false, HelpMessage = "Timespan for Dormant accounts. Default 90 days")]
    [ValidateRange(1,1000)]
    [int] $DormantTimeSpan = 90,

    [Parameter(Mandatory = $false, HelpMessage = "The PageSize to set for the LDAP searcher object. Default 200")]
    [ValidateRange(1,10000)]
    [int] $PageSize = 200,

    [Parameter(Mandatory = $false, HelpMessage = "The number of threads to use during processing of objects. Default 10")]
    [ValidateRange(1,100)]
    [int] $Threads = 10,

    [Parameter(Mandatory = $false, HelpMessage = "The number of processed objects which will be flushed to disk. Default -1 (After all objects are processed).")]
    [ValidateRange(-1,1000000)]
    [int] $FlushCount = -1,

    [Parameter(Mandatory = $false, HelpMessage = "Create ADRecon Log using Start-Transcript")]
    [switch] $Log
)

$ADWSSource = @"
// Thanks Dennis Albuquerque for the C# multithreading code
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Security.Principal;
using System.Management.Automation;

namespace ADRecon
{
    public static class ADWSClass
    {
        private static DateTime Date1;
        private static int PassMaxAge;
        private static int DormantTimeSpan;
        private static string FilePath;
        private static bool ADRSTDOUT = false;
        private static bool ADRCSV = false;
        private static readonly HashSet<string> Groups = new HashSet<string> ( new String[] {"268435456", "268435457", "536870912", "536870913"} );
        private static readonly HashSet<string> Users = new HashSet<string> ( new String[] { "805306368" } );
        private static readonly HashSet<string> Computers = new HashSet<string> ( new String[] { "805306369" }) ;
        private static readonly HashSet<string> TrustAccounts = new HashSet<string> ( new String[] { "805306370" } );

		private static readonly Dictionary<String, String> Replacements = new Dictionary<String, String>()
        {
            //{System.Environment.NewLine, ""},
            //{",", ";"},
            {"\"", "'"}
        };

        public static String CleanString(String StringtoClean)
        {
            // Remove extra spaces and new lines
            String CleanedString = String.Join(" ", ((Convert.ToString(StringtoClean)).Split((string[]) null, StringSplitOptions.RemoveEmptyEntries)));
            foreach (String Replacement in Replacements.Keys)
            {
                CleanedString = CleanedString.Replace(Replacement, Replacements[Replacement]);
            }
            return CleanedString;
        }

        public static void UserParser(Object[] AdUsers, DateTime Date1, int PassMaxAge, string FilePath, int DormantTimeSpan, int numOfThreads, int flushCnt, String[] OutputType)
        {
            ADWSClass.Date1 = Date1;
            ADWSClass.PassMaxAge = PassMaxAge;
            ADWSClass.DormantTimeSpan = DormantTimeSpan;
            ADWSClass.FilePath = FilePath;

            foreach (string OType in OutputType)
            {
                switch (OType)
                {
                    case "STDOUT":
                        ADWSClass.ADRSTDOUT = true;
                        continue;
                    case "CSV":
                        ADWSClass.ADRCSV = true;
                        break;
                }
            }
            if (ADWSClass.ADRCSV == true)
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@FilePath))
                {
                    string HeaderRow = String.Format("Name,UserName,Enabled,Cannot Change Password,Password Never Expires,Must Change Password at Logon,Days Since Last Password Change,Password Not Changed after Max Age,Account Locked Out,Never Logged in,Days Since Last Logon,Dormant (> {0} days),Reversibly Encryped Password,Password Not Required,Trusted for Delegation,Trusted to Auth for Delegation,Does Not Require Pre Auth,Logon Workstations,AdminCount,Primary GroupID,SID,SIDHistory,Description,Password LastSet,Last Logon Date,When Created,When Changed,DistinguishedName,CanonicalName",DormantTimeSpan);
                    file.WriteLine(HeaderRow);
                }
            }
            Console.WriteLine("[*] Total Users: " + AdUsers.Length);
            runProcessor(AdUsers, numOfThreads, flushCnt, "Users", "CSV");
            ADWSClass.ADRSTDOUT = false;
            ADWSClass.ADRCSV = false;
        }

        public static int UserSPNParser(Object[] AdUsers, string FilePath, int numOfThreads, int flushCnt, String[] OutputType)
        {
            if (AdUsers.Length == 1)
            {
                return AdUsers.Length;
            }

            ADWSClass.FilePath = FilePath;

            foreach (string OType in OutputType)
            {
                switch (OType)
                {
                    case "STDOUT":
                        ADWSClass.ADRSTDOUT = true;
                        continue;
                    case "CSV":
                        ADWSClass.ADRCSV = true;
                        break;
                }
            }
            if (ADWSClass.ADRCSV == true)
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@FilePath))
                {
                    string HeaderRow = "Name,Username,Service,Host,Password Last Set,Description";
                    file.WriteLine(HeaderRow);
                }
            }
            runProcessor(AdUsers, numOfThreads, flushCnt, "UserSPNs", "CSV");
            ADWSClass.ADRSTDOUT = false;
            ADWSClass.ADRCSV = false;
            return AdUsers.Length;
        }

        public static void GroupParser(Object[] AdGroups, string FilePath, int numOfThreads, int flushCnt, String[] OutputType)
        {
            ADWSClass.FilePath = FilePath;
            foreach (string OType in OutputType)
            {
                switch (OType)
                {
                    case "STDOUT":
                        ADWSClass.ADRSTDOUT = true;
                        continue;
                    case "CSV":
                        ADWSClass.ADRCSV = true;
                        break;
                }
            }
            if (ADWSClass.ADRCSV == true)
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@FilePath))
                {
                    string HeaderRow = "Group,ManagedBy,whenCreated,whenChanged,Description,SID,DistinguishedName,CanonicalName";
                    file.WriteLine(HeaderRow);
                }
            }
            Console.WriteLine("[*] Total Groups: " + AdGroups.Length);
            runProcessor(AdGroups, numOfThreads, flushCnt, "Groups", "CSV");
            ADWSClass.ADRSTDOUT = false;
            ADWSClass.ADRCSV = false;
        }

        public static void GroupMemberParser(Object[] AdGroupMembers, string FilePath, int numOfThreads, int flushCnt, String[] OutputType)
        {
            ADWSClass.FilePath = FilePath;

            foreach (string OType in OutputType)
            {
                switch (OType)
                {
                    case "STDOUT":
                        ADWSClass.ADRSTDOUT = true;
                        continue;
                    case "CSV":
                        ADWSClass.ADRCSV = true;
                        break;
                }
            }
            if (ADWSClass.ADRCSV == true)
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@FilePath))
                {
                    string HeaderRow = "Group Name, Member UserName, Member Name, AccountType";
                    file.WriteLine(HeaderRow);
                }
            }
            Console.WriteLine("[*] Total GroupMember Objects: " + AdGroupMembers.Length);
            runProcessor(AdGroupMembers, numOfThreads, flushCnt, "GroupMembers", "CSV");
            ADWSClass.ADRSTDOUT = false;
            ADWSClass.ADRCSV = false;
        }

        public static int ComputerParser(Object[] AdComputers, DateTime Date1, string FilePath, int numOfThreads, int flushCnt, String[] OutputType)
        {
            Console.WriteLine("[*] Total Computers: " + AdComputers.Length);
            if (AdComputers.Length == 1)
            {
                return AdComputers.Length;
            }

            ADWSClass.Date1 = Date1;
            ADWSClass.FilePath = FilePath;

            foreach (string OType in OutputType)
            {
                switch (OType)
                {
                    case "STDOUT":
                        ADWSClass.ADRSTDOUT = true;
                        continue;
                    case "CSV":
                        ADWSClass.ADRCSV = true;
                        break;
                }
            }
            if (ADWSClass.ADRCSV == true)
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@FilePath))
                {
                    string HeaderRow = "Name,DNSHostName,Enabled,IPv4Address,OperatingSystem,Days Since Last Logon,Days Since Last Password Change,Trusted for Delegation,Trusted to Auth for Delegation,Username,Primary Group ID,Description,Password LastSet,Last Logon Date,whenCreated,whenChanged,Distinguished Name";
                    file.WriteLine(HeaderRow);
                }
            }
            runProcessor(AdComputers, numOfThreads, flushCnt, "Computers", "CSV");
            ADWSClass.ADRSTDOUT = false;
            ADWSClass.ADRCSV = false;
            return AdComputers.Length;
        }

        public static int ComputerSPNParser(Object[] AdComputers, string FilePath, int numOfThreads, int flushCnt, String[] OutputType)
        {
            if (AdComputers.Length == 1)
            {
                return AdComputers.Length;
            }

            ADWSClass.FilePath = FilePath;

            foreach (string OType in OutputType)
            {
                switch (OType)
                {
                    case "STDOUT":
                        ADWSClass.ADRSTDOUT = true;
                        continue;
                    case "CSV":
                        ADWSClass.ADRCSV = true;
                        break;
                }
            }
            if (ADWSClass.ADRCSV == true)
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@FilePath))
                {
                    string HeaderRow = "Name,Service,Host";
                    file.WriteLine(HeaderRow);
                }
            }
            runProcessor(AdComputers, numOfThreads, flushCnt, "ComputerSPNs", "CSV");
            ADWSClass.ADRSTDOUT = false;
            ADWSClass.ADRCSV = false;
            return AdComputers.Length;
        }

        static void runProcessor(Object[] arrayToProcess, int numOfThreads, int flushCnt, string processorType, String resultHandlerType)
        {
            int totalRecords = arrayToProcess.Length;
            //Console.WriteLine(String.Format("Running {0} records over {1} threads, flushing every {2} records",
            //    totalRecords, numOfThreads, (flushCnt < 0 ? "NEVER" : flushCnt.ToString())));
            IRecordProcessor recordProcessor = recordProcessorFactory(processorType);
            IResultsHandler resultsHandler = resultHandlerFactory(resultHandlerType, flushCnt);
            int numberOfRecordsPerThread = totalRecords / numOfThreads;
            int remainders = totalRecords % numOfThreads;

            Thread[] threads = new Thread[numOfThreads];
            for (int i = 0; i < numOfThreads; i++)
            {
                int numberOfRecordsToProcess = numberOfRecordsPerThread;
                if (i == (numOfThreads - 1))
                {
                    //last thread, do the remaining records
                    numberOfRecordsToProcess += remainders;
                }

                //split the full array into chunks to be given to different threads
                Object[] sliceToProcess = new Object[numberOfRecordsToProcess];
                Array.Copy(arrayToProcess, i * numberOfRecordsPerThread, sliceToProcess, 0, numberOfRecordsToProcess);
                ProcessorThread processorThread = new ProcessorThread(i, recordProcessor, resultsHandler, sliceToProcess);
                threads[i] = new Thread(processorThread.processThreadRecords);
                threads[i].Start();
            }
            foreach (Thread t in threads)
            {
                t.Join();
            }

            resultsHandler.finalise();
        }

        static IRecordProcessor recordProcessorFactory(String name)
        {
            switch (name)
            {
                case "Users":
                    return new UserRecordProcessor();
                case "UserSPNs":
                    return new UserSPNRecordProcessor();
                case "Groups":
                    return new GroupRecordProcessor();
                case "GroupMembers":
                    return new GroupMemberRecordProcessor();
                case "Computers":
                    return new ComputerRecordProcessor();
                case "ComputerSPNs":
                    return new ComputerSPNRecordProcessor();
            }
            throw new ArgumentException("Invalid processor type " + name);
        }

        static IResultsHandler resultHandlerFactory(String name, int flushCnt)
        {
            switch (name)
            {
                case "CSV":
                    return new CsvResultsHandler(flushCnt);
                case "TXT":
                    return new TxtResultsHandler(flushCnt);
            }
            throw new ArgumentException("Invalid processor type " + name);
        }

        class ProcessorThread
        {
            readonly int id;
            readonly IRecordProcessor recordProcessor;
            readonly IResultsHandler resultsHandler;
            readonly Object[] objectsToBeProcessed;

            public ProcessorThread(int id, IRecordProcessor recordProcessor, IResultsHandler resultsHandler, Object[] objectsToBeProcessed)
            {
                this.recordProcessor = recordProcessor;
                this.id = id;
                this.resultsHandler = resultsHandler;
                this.objectsToBeProcessed = objectsToBeProcessed;
            }

            public void processThreadRecords()
            {
                for (int i = 0; i < objectsToBeProcessed.Length; i++)
                {
                    Object[] result = recordProcessor.processRecord(objectsToBeProcessed[i]);
                    resultsHandler.processResults(result); //this is a thread safe operation
                }
            }
        }

        //The interface and implmentation class used to process a record (this implemmentation just returns a log type string)

        interface IRecordProcessor
        {
            Object[] processRecord(Object record);
        }

        class UserRecordProcessor : IRecordProcessor
        {
            public Object[] processRecord(Object record)
            {
                try
                {
                    PSObject AdUser = (PSObject) record;
                    bool? Enabled = null;
                    bool MustChangePasswordatLogon = false;
                    int DaysSinceLastPasswordChange = -1;
                    bool PasswordNotChangedafterMaxAge = false;
                    bool NeverLoggedIn = false;
                    int DaysSinceLastLogon = -1;
                    bool Dormant = false;
                    String SIDHistory = "";
                    DateTime PasswordLastSet = Convert.ToDateTime(AdUser.Members["PasswordLastSet"].Value);
                    try
                    {
                        // The Enabled field can be blank which raises an exception. This may occur when the user is not allowed to query the UserAccountControl attribute.
                        Enabled = (bool) AdUser.Members["Enabled"].Value;
                    }
                    catch //(Exception e)
                    {
                        //    Console.WriteLine("{0} Exception caught.", e);
                    }
                    if (Convert.ToString(AdUser.Members["pwdlastset"].Value) == "0")
                    {
                        MustChangePasswordatLogon = true;
                    }
                    else
                    {
                        DaysSinceLastPasswordChange = Math.Abs((Date1 - PasswordLastSet).Days);
                        if (DaysSinceLastPasswordChange > PassMaxAge)
                        {
                            PasswordNotChangedafterMaxAge = true;
                        }
                    }
                    DateTime LastLogonDate = Convert.ToDateTime(AdUser.Members["LastLogonDate"].Value);
                    if (AdUser.Members["LastLogonDate"].Value != null)
                    {
                        DaysSinceLastLogon = Math.Abs((Date1 - LastLogonDate).Days);
                        if (DaysSinceLastLogon > DormantTimeSpan)
                        {
                            Dormant = true;
                        }
                    }
                    else
                    {
                        NeverLoggedIn = true;
                    }
                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection history = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdUser.Members["SIDHistory"].Value;
                    if (history.Value is System.Security.Principal.SecurityIdentifier[])
                    {
                        string sids = "";
                        foreach (var value in (SecurityIdentifier[]) history.Value)
                        {
                            sids = sids + "," + Convert.ToString(value);
                        }
                        SIDHistory = sids.TrimStart(',');
                    }
                    else
                    {
                        SIDHistory = history != null ? Convert.ToString(history.Value) : "";
                    }
                    if (ADWSClass.ADRSTDOUT == true)
                    {
                        Console.WriteLine("\nName{0,33}{2}\nUserName{0,29}{3}\nEnabled{0,30}{4}\nCannot Change Password{0,15}{5}\nPassword Never Expires{0,15}{6}\nMust Change Password at Logon{0,8}{7}\nDays Since Last Password Change{0,6}{8}\nPassword Not Changed after Max Age{0}{9}\nAccount Locked Out{0,19}{10}\nNever Logged in{0,22}{11}\nDays Since Last Logon{0,16}{12}\nDormant (> {1} days){0,18}{13}\nReversibly Encryped Password{0,9}{14}\nPassword Not Required{0,16}{15}\nTrusted for Delegation{0,15}{16}\nTrusted to Auth for Delegation{0,7}{17}\nDoes Not Require Pre Auth{0,12}{18}\nLogon Workstations{0,19}{19}\nAdminCount{0,27}{20}\nPrimary GroupID{0,22}{21}\nSID{0,34}{22}\nSIDHistory{0,27}{23}\nDescription{0,26}{24}\nPassword LastSet{0,21}{25}\nLast Logon Date{0,22}{26}\nWhen Created{0,25}{27}\nWhen Changed{0,25}{28}\nDistinguishedName{0,20}{29}\nCanonicalName{0,24}{30}\n", " : ", DormantTimeSpan, AdUser.Members["Name"].Value, AdUser.Members["SamAccountName"].Value, Enabled, AdUser.Members["CannotChangePassword"].Value, AdUser.Members["PasswordNeverExpires"].Value, MustChangePasswordatLogon, DaysSinceLastPasswordChange, PasswordNotChangedafterMaxAge, AdUser.Members["LockedOut"].Value, NeverLoggedIn, DaysSinceLastLogon, Dormant, AdUser.Members["AllowReversiblePasswordEncryption"].Value, AdUser.Members["PasswordNotRequired"].Value, AdUser.Members["TrustedForDelegation"].Value, AdUser.Members["TrustedToAuthForDelegation"].Value, AdUser.Members["DoesNotRequirePreAuth"].Value, AdUser.Members["LogonWorkstations"].Value, AdUser.Members["AdminCount"].Value, AdUser.Members["primaryGroupID"].Value, AdUser.Members["SID"].Value, SIDHistory, AdUser.Members["Description"].Value, PasswordLastSet, LastLogonDate, AdUser.Members["whenCreated"].Value, AdUser.Members["whenChanged"].Value, AdUser.Members["DistinguishedName"].Value, AdUser.Members["CanonicalName"].Value);
                    }
                    return new Object[] { AdUser.Members["Name"].Value, AdUser.Members["SamAccountName"].Value, Enabled, AdUser.Members["CannotChangePassword"].Value, AdUser.Members["PasswordNeverExpires"].Value, MustChangePasswordatLogon, DaysSinceLastPasswordChange, PasswordNotChangedafterMaxAge, AdUser.Members["LockedOut"].Value, NeverLoggedIn, DaysSinceLastLogon, Dormant, AdUser.Members["AllowReversiblePasswordEncryption"].Value, AdUser.Members["PasswordNotRequired"].Value, AdUser.Members["TrustedForDelegation"].Value, AdUser.Members["TrustedToAuthForDelegation"].Value, AdUser.Members["DoesNotRequirePreAuth"].Value, AdUser.Members["LogonWorkstations"].Value, AdUser.Members["AdminCount"].Value, AdUser.Members["primaryGroupID"].Value, AdUser.Members["SID"].Value, SIDHistory, AdUser.Members["Description"].Value, PasswordLastSet, LastLogonDate, AdUser.Members["whenCreated"].Value, AdUser.Members["whenChanged"].Value, AdUser.Members["DistinguishedName"].Value, AdUser.Members["CanonicalName"].Value };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new Object[] { };
                }
            }
        }

        class UserSPNRecordProcessor : IRecordProcessor
        {
            public Object[] processRecord(Object record)
            {
                try
                {
                    PSObject AdUser = (PSObject) record;
                    List<Object> SPNList = new List<Object>();
                    DateTime PasswordLastSet = DateTime.FromFileTime((long)AdUser.Members["pwdLastSet"].Value);
                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection SPNs = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdUser.Members["servicePrincipalName"].Value;
                    if (SPNs.Value is System.String[])
                    {
                        foreach (String SPN in (System.String[])SPNs.Value)
                        {
                            String[] SPNArray = SPN.Split('/');
                            SPNList.Add(new Object[] { AdUser.Members["Name"].Value, AdUser.Members["SamAccountName"].Value, SPNArray[0], SPNArray[1], PasswordLastSet, AdUser.Members["Description"].Value });
                        }
                    }
                    else
                    {
                        String[] SPNArray = Convert.ToString(SPNs.Value).Split('/');
                        SPNList.Add(new Object[] { AdUser.Members["Name"].Value, AdUser.Members["SamAccountName"].Value, SPNArray[0], SPNArray[1], PasswordLastSet, AdUser.Members["Description"].Value });
                    }
                    if (ADWSClass.ADRSTDOUT == true)
                    {
                        foreach (Object[] Obj in SPNList.ToArray())
                        {
                            Console.WriteLine("Name{0,16}{1}\nUsername{0,12}{2}\nService{0,13}{3}\nHost{0,16}{4}\nPassword Last Set{0}{5}\nDescription{0,9}{6}\n", " : ", Obj[0], Obj[1], Obj[2], Obj[3], Obj[4], Obj[5]);
                        }
                    }
                    return SPNList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new Object[] { };
                }
            }
        }

        class GroupRecordProcessor : IRecordProcessor
        {
            public Object[] processRecord(Object record)
            {
                try
                {
                    PSObject AdGroup = (PSObject) record;
                    string ManagedByValue = Convert.ToString(AdGroup.Members["managedBy"].Value);
                    string ManagedBy = "";
                    if (AdGroup.Members["managedBy"].Value != null)
                    {
                        ManagedBy = (ManagedByValue.Split(',')[0]).Split('=')[1];
                    }
                    if (ADWSClass.ADRSTDOUT == true)
                    {
                        Console.WriteLine("Group{0,15}{1}\nManagedBy{0,11}{2}\nwhenCreated{0,9}{3}\nwhenChanged{0,9}{4}\nDescription{0,9}{5}\nSID{0,17}{6}\nDistinguishedName{0}{7}\nCanonicalName{0,7}{8}\n", " : ", AdGroup.Members["SamAccountName"].Value, ManagedBy, AdGroup.Members["whenCreated"].Value, AdGroup.Members["whenChanged"].Value, AdGroup.Members["Description"].Value, AdGroup.Members["sid"].Value, AdGroup.Members["DistinguishedName"].Value, AdGroup.Members["CanonicalName"].Value);
                    }
                    return new Object[] { AdGroup.Members["SamAccountName"].Value, ManagedBy, AdGroup.Members["whenCreated"].Value, AdGroup.Members["whenChanged"].Value, AdGroup.Members["Description"].Value, AdGroup.Members["sid"].Value, AdGroup.Members["DistinguishedName"].Value, AdGroup.Members["CanonicalName"].Value };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new Object[] { };
                }
            }
        }

        class GroupMemberRecordProcessor : IRecordProcessor
        {
            public Object[] processRecord(Object record)
            {
                try
                {
                    // based on https://github.com/BloodHoundAD/BloodHound/blob/master/PowerShell/BloodHound.ps1
                    PSObject AdGroup = (PSObject) record;
                    List<Object> GroupsList = new List<Object>();
                    string SamAccountType = Convert.ToString(AdGroup.Members["samaccounttype"].Value);
                    string AccountType = "";
                    string GroupName = "";
                    string MemberUserName = "-";
                    string MemberName = "";
                    if (Groups.Contains(SamAccountType))
                    {
                        AccountType = "group";
                        MemberName = ((Convert.ToString(AdGroup.Members["DistinguishedName"].Value)).Split(',')[0]).Split('=')[1];
                        Microsoft.ActiveDirectory.Management.ADPropertyValueCollection MemberGroups = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdGroup.Members["memberof"].Value;
                        if (AdGroup.Members["memberof"].Value != null)
                        {
                            if (MemberGroups.Value is System.String[])
                            {
                                foreach (String GroupMember in (System.String[])MemberGroups.Value)
                                {
                                    GroupName = ((Convert.ToString(GroupMember)).Split(',')[0]).Split('=')[1];
                                    GroupsList.Add(new Object[] { GroupName, MemberUserName, MemberName, AccountType });
                                }
                            }
                            else
                            {
                                GroupName = (Convert.ToString(MemberGroups.Value).Split(',')[0]).Split('=')[1];
                                GroupsList.Add(new Object[] { GroupName, MemberUserName, MemberName, AccountType });
                            }
                        }
                    }
                    if (Users.Contains(SamAccountType))
                    {
                        AccountType = "user";
                        MemberName = ((Convert.ToString(AdGroup.Members["DistinguishedName"].Value)).Split(',')[0]).Split('=')[1];
                        MemberUserName = Convert.ToString(AdGroup.Members["sAMAccountName"].Value);
                        Microsoft.ActiveDirectory.Management.ADPropertyValueCollection MemberGroups = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdGroup.Members["memberof"].Value;
                        if (AdGroup.Members["memberof"].Value != null)
                        {
                            if (MemberGroups.Value is System.String[])
                            {
                                foreach (String GroupMember in (System.String[])MemberGroups.Value)
                                {
                                    GroupName = ((Convert.ToString(GroupMember)).Split(',')[0]).Split('=')[1];
                                    GroupsList.Add(new Object[] { GroupName, MemberUserName, MemberName, AccountType });
                                }
                            }
                            else
                            {
                                GroupName = (Convert.ToString(MemberGroups.Value).Split(',')[0]).Split('=')[1];
                                GroupsList.Add(new Object[] { GroupName, MemberUserName, MemberName, AccountType });
                            }
                        }
                    }
                    if (Computers.Contains(SamAccountType))
                    {
                        AccountType = "computer";
                        MemberName = ((Convert.ToString(AdGroup.Members["DistinguishedName"].Value)).Split(',')[0]).Split('=')[1];
                        MemberUserName = Convert.ToString(AdGroup.Members["sAMAccountName"].Value);
                        Microsoft.ActiveDirectory.Management.ADPropertyValueCollection MemberGroups = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdGroup.Members["memberof"].Value;
                        if (AdGroup.Members["memberof"].Value != null)
                        {
                            if (MemberGroups.Value is System.String[])
                            {
                                foreach (String GroupMember in (System.String[])MemberGroups.Value)
                                {
                                    GroupName = ((Convert.ToString(GroupMember)).Split(',')[0]).Split('=')[1];
                                    GroupsList.Add(new Object[] { GroupName, MemberUserName, MemberName, AccountType });
                                }
                            }
                            else
                            {
                                GroupName = (Convert.ToString(MemberGroups.Value).Split(',')[0]).Split('=')[1];
                                GroupsList.Add(new Object[] { GroupName, MemberUserName, MemberName, AccountType });
                            }
                        }
                    }
                    if (TrustAccounts.Contains(SamAccountType))
                    {
                        // TO DO
                    }
                    if (ADWSClass.ADRSTDOUT == true)
                    {
                        foreach (Object[] Obj in GroupsList.ToArray())
                        {
                            Console.WriteLine("Group Name{0,8}{1}\nMember UserName{0}{2}\nMember Name{0,7}{3}\nAccountType{0,7}{4}\n", " : ", Obj[0], Obj[1], Obj[2], Obj[3]);
                        }
                    }
                    return GroupsList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new Object[] { };
                }
            }
        }

        class ComputerRecordProcessor : IRecordProcessor
        {
            public Object[] processRecord(Object record)
            {
                try
                {
                    PSObject AdComputer = (PSObject) record;
                    int DaysSinceLastPasswordChange = -1;
                    int DaysSinceLastLogon = -1;
                    DateTime LastLogonDate = Convert.ToDateTime(AdComputer.Members["LastLogonDate"].Value);
                    if (AdComputer.Members["LastLogonDate"].Value != null)
                    {
                        DaysSinceLastLogon = Math.Abs((Date1 - LastLogonDate).Days);
                    }
                    DateTime PasswordLastSet = Convert.ToDateTime(AdComputer.Members["PasswordLastSet"].Value);
                    if (AdComputer.Members["PasswordLastSet"].Value != null)
                    {
                        DaysSinceLastPasswordChange = Math.Abs((Date1 - PasswordLastSet).Days);
                    }
                    if (ADWSClass.ADRSTDOUT == true)
                    {
                        Console.WriteLine("Name{0,30}{1}\nDNSHostName{0,23}{2}\nEnabled{0,27}{3}\nIPv4Address{0,23}{4}\nOperatingSystem{0,19}{5}\nDays Since Last Logon{0,13}{6}\nDays Since Last Password Change{0}{7}\nTrusted for Delegation{0,12}{8}\nTrusted to Auth for Delegation{0,4}{9}\nUsername{0,26}{10}\nPrimary Group ID{0,18}{11}\nDescription{0,23}{12}\nPassword LastSet{0,18}{13}\nLast Logon Date{0,19}{14}\nwhenCreated{0,23}{15}\nwhenChanged{0,23}{16}\nDistinguished Name{0,16}{17}\n", " : ", AdComputer.Members["Name"].Value, AdComputer.Members["DNSHostName"].Value, AdComputer.Members["Enabled"].Value, AdComputer.Members["IPv4Address"].Value, (AdComputer.Members["OperatingSystem"].Value != null ? AdComputer.Members["OperatingSystem"].Value : "-"), DaysSinceLastLogon, DaysSinceLastPasswordChange, AdComputer.Members["TrustedForDelegation"].Value, AdComputer.Members["TrustedToAuthForDelegation"].Value, AdComputer.Members["SamAccountName"].Value, AdComputer.Members["primaryGroupID"].Value, AdComputer.Members["Description"].Value, PasswordLastSet, LastLogonDate, AdComputer.Members["whenCreated"].Value, AdComputer.Members["whenChanged"].Value, AdComputer.Members["DistinguishedName"].Value);
                    }
                    return new Object[] { AdComputer.Members["Name"].Value, AdComputer.Members["DNSHostName"].Value, AdComputer.Members["Enabled"].Value, AdComputer.Members["IPv4Address"].Value, (AdComputer.Members["OperatingSystem"].Value != null ? AdComputer.Members["OperatingSystem"].Value : "-"), DaysSinceLastLogon, DaysSinceLastPasswordChange, AdComputer.Members["TrustedForDelegation"].Value, AdComputer.Members["TrustedToAuthForDelegation"].Value, AdComputer.Members["SamAccountName"].Value, AdComputer.Members["primaryGroupID"].Value, AdComputer.Members["Description"].Value, PasswordLastSet, LastLogonDate, AdComputer.Members["whenCreated"].Value, AdComputer.Members["whenChanged"].Value, AdComputer.Members["DistinguishedName"].Value };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new Object[] { };
                }
            }
        }

        class ComputerSPNRecordProcessor : IRecordProcessor
        {
            public Object[] processRecord(Object record)
            {
                try
                {
                    PSObject AdComputer = (PSObject) record;
                    List<Object> SPNList = new List<Object>();
                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection SPNs = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdComputer.Members["servicePrincipalName"].Value;
                    if (SPNs.Value is System.String[])
                    {
                        foreach (String SPN in (System.String[])SPNs.Value)
                        {
                            String[] SPNArray = SPN.Split('/');
                            SPNList.Add(new Object[] { AdComputer.Members["Name"].Value, SPNArray[0], SPNArray[1] });
                        }
                    }
                    else
                    {
                        String[] SPNArray = Convert.ToString(SPNs.Value).Split('/');
                        SPNList.Add(new Object[] { AdComputer.Members["Name"].Value, SPNArray[0], SPNArray[1] });
                    }
                    if (ADWSClass.ADRSTDOUT == true)
                    {
                        foreach (Object[] Obj in SPNList.ToArray())
                        {
                            Console.WriteLine("Name{0,6}{1}\nService{0}{2}\nHost{0,6}{3}\n", " : ", Obj[0], Obj[1], Obj[2]);
                        }
                    }
                    return SPNList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new Object[] { };
                }
            }
        }

        //The interface and implmentation class used to handle the results (this implementation just writes the strings to a file)

        interface IResultsHandler
        {
            void processResults(Object[] t);

            void finalise();
        }

        abstract class SimpleResultsHandler : IResultsHandler
        {
            private Object lockObj = new Object();
            private List<String> processed = new List<String>();
            private readonly int flushCnt;

            public SimpleResultsHandler(int flushCnt)
            {
                this.flushCnt = flushCnt;
            }

            public void processResults(Object[] results)
            {
                lock (lockObj)
                {
                    if (results.Length != 0)
                    {
                        if (results[0] is System.Object[])
                        {
                            for (var i = 0; i < results.Length; i++)
                            {
                                processed.Add(convertObject((Object[])results[i]));
                            }
                        }
                        else
                        {
                            processed.Add(convertObject(results));
                        }
                        if (flushCnt > 0)
                        {
                            if (processed.Count >= flushCnt)
                            {
                                writeFile();
                            }
                        }
                    }
                }
            }

            public void finalise()
            {
                writeFile();
            }

            private void writeFile()
            {
                lock (lockObj)
                {
                    if (ADWSClass.ADRCSV == true)
                    {
                        using (StreamWriter outputFile = new StreamWriter(@ADWSClass.FilePath, true))
                        {
                            outputFile.Write(String.Join("\r\n", processed.ToArray()));
                        }
                    }
                    processed.Clear();
                }
            }

            protected abstract String convertObject(Object[] resultsObject);
        }


        class CsvResultsHandler : SimpleResultsHandler
        {
            public CsvResultsHandler(int flushCnt) : base(flushCnt)
            {
            }

            protected override String convertObject(Object[] resultsObject)
            {
                return createCsvLine(resultsObject);
            }

            static String createCsvLine(Object[] resultsObject)
            {
                try
                {
                    // No String.Join(String, Object[]) in CLR 2.0.50727 (Windows 7)
                    String[] row = new String[resultsObject.Length];
                    for (int i=0; i < resultsObject.Length; i++)
                    {
                        row[i] = CleanString(Convert.ToString(resultsObject[i]));
                    }
                    return "\"" + String.Join("\",\"", row) + "\"";
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return "";
                }
            }
        }

        class TxtResultsHandler : SimpleResultsHandler
        {
            public TxtResultsHandler(int flushCnt) : base(flushCnt)
            {
            }

            protected override String convertObject(Object[] resultsObject)
            {
                try
                {
                    // No String.Join(String, Object[]) in CLR 2.0.50727 (Windows 7)
                    String[] row = new String[resultsObject.Length];
                    for (int i=0; i < resultsObject.Length; i++)
                    {
                        row[i] = CleanString(Convert.ToString(resultsObject[i]));
                    }
                    return String.Join("\t", row);
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return "";
                }
            }
        }
    }
}
"@

$LDAPSource = @"
// Thanks Dennis Albuquerque for the C# multithreading code
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.DirectoryServices;
using System.Security.Principal;

namespace ADRecon
{
    public static class LDAPClass
    {
        private static DateTime Date1;
        private static int PassMaxAge;
        private static int DormantTimeSpan;
        private static string FilePath;
        private static bool ADRSTDOUT = false;
        private static bool ADRCSV = false;
        private static Dictionary<string, bool> CannotChangePasswordDict;
        private static readonly HashSet<string> Groups = new HashSet<string> ( new String[] {"268435456", "268435457", "536870912", "536870913"} );
        private static readonly HashSet<string> Users = new HashSet<string> ( new String[] { "805306368" } );
        private static readonly HashSet<string> Computers = new HashSet<string> ( new String[] { "805306369" }) ;
        private static readonly HashSet<string> TrustAccounts = new HashSet<string> ( new String[] { "805306370" } );

        [Flags]
        //Values taken from https://support.microsoft.com/en-au/kb/305144
        public enum UACFlags
        {
            SCRIPT = 1,        // 0x1
            ACCOUNTDISABLE = 2,        // 0x2
            HOMEDIR_REQUIRED = 8,        // 0x8
            LOCKOUT = 16,       // 0x10
            PASSWD_NOTREQD = 32,       // 0x20
            PASSWD_CANT_CHANGE = 64,       // 0x40
            ENCRYPTED_TEXT_PASSWORD_ALLOWED = 128,      // 0x80
            TEMP_DUPLICATE_ACCOUNT = 256,      // 0x100
            NORMAL_ACCOUNT = 512,      // 0x200
            INTERDOMAIN_TRUST_ACCOUNT = 2048,     // 0x800
            WORKSTATION_TRUST_ACCOUNT = 4096,     // 0x1000
            SERVER_TRUST_ACCOUNT = 8192,     // 0x2000
            DONT_EXPIRE_PASSWD = 65536,    // 0x10000
            MNS_LOGON_ACCOUNT = 131072,   // 0x20000
            SMARTCARD_REQUIRED = 262144,   // 0x40000
            TRUSTED_FOR_DELEGATION = 524288,   // 0x80000
            NOT_DELEGATED = 1048576,  // 0x100000
            USE_DES_KEY_ONLY = 2097152,  // 0x200000
            DONT_REQUIRE_PREAUTH = 4194304,  // 0x400000
            PASSWORD_EXPIRED = 8388608,  // 0x800000
            TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION = 16777216, // 0x1000000
            PARTIAL_SECRETS_ACCOUNT = 67108864 // 0x04000000
        }

		private static readonly Dictionary<String, String> Replacements = new Dictionary<String, String>()
        {
            //{System.Environment.NewLine, ""},
            //{",", ";"},
            {"\"", "'"}
        };

        public static String CleanString(String StringtoClean)
        {
            // Remove extra spaces and new lines
            String CleanedString = String.Join(" ", ((Convert.ToString(StringtoClean)).Split((string[]) null, StringSplitOptions.RemoveEmptyEntries)));
            foreach (String Replacement in Replacements.Keys)
            {
                CleanedString = CleanedString.Replace(Replacement, Replacements[Replacement]);
            }
            return CleanedString;
        }

        public static void UserParser(Object[] AdUsers, DateTime Date1, int PassMaxAge, string FilePath, Dictionary<string, bool> CannotChangePasswordDict, int DormantTimeSpan, int numOfThreads, int flushCnt, String[] OutputType)
        {
            LDAPClass.Date1 = Date1;
            LDAPClass.PassMaxAge = PassMaxAge;
            LDAPClass.DormantTimeSpan = DormantTimeSpan;
            LDAPClass.FilePath = FilePath;
            LDAPClass.CannotChangePasswordDict = CannotChangePasswordDict;

            foreach (string OType in OutputType)
            {
                switch (OType)
                {
                    case "STDOUT":
                        LDAPClass.ADRSTDOUT = true;
                        continue;
                    case "CSV":
                        LDAPClass.ADRCSV = true;
                        break;
                }
            }
            if (LDAPClass.ADRCSV == true)
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@FilePath))
                {
                    string HeaderRow = String.Format("Name,UserName,Enabled,Cannot Change Password,Password Never Expires,Must Change Password at Logon,Days Since Last Password Change,Password Not Changed after Max Age,Account Locked Out,Never Logged in,Days Since Last Logon,Dormant (> {0} days),Reversibly Encryped Password,Password Not Required,Trusted for Delegation,Trusted to Auth for Delegation,Does Not Require Pre Auth,Logon Workstations,AdminCount,Primary GroupID,SID,SIDHistory,Description,Password LastSet,Last Logon Date,When Created,When Changed,DistinguishedName,CanonicalName",DormantTimeSpan);
                    file.WriteLine(HeaderRow);
                }
            }
            Console.WriteLine("[*] Total Users: " + AdUsers.Length);
            runProcessor(AdUsers, numOfThreads, flushCnt, "Users", "CSV");
            LDAPClass.ADRSTDOUT = false;
            LDAPClass.ADRCSV = false;
        }

        public static void UserSPNParser(Object[] AdUsers, string FilePath, int numOfThreads, int flushCnt, String[] OutputType)
        {
            LDAPClass.FilePath = FilePath;

            foreach (string OType in OutputType)
            {
                switch (OType)
                {
                    case "STDOUT":
                        LDAPClass.ADRSTDOUT = true;
                        continue;
                    case "CSV":
                        LDAPClass.ADRCSV = true;
                        break;
                }
            }
            if (LDAPClass.ADRCSV == true)
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@FilePath))
                {
                    string HeaderRow = "Name,Username,Service,Host,Password Last Set,Description";
                    file.WriteLine(HeaderRow);
                }
            }
            runProcessor(AdUsers, numOfThreads, flushCnt, "UserSPNs", "CSV");
            LDAPClass.ADRSTDOUT = false;
            LDAPClass.ADRCSV = false;
        }

        public static void GroupParser(Object[] AdGroups, string FilePath, int numOfThreads, int flushCnt, String[] OutputType)
        {
            LDAPClass.FilePath = FilePath;

            foreach (string OType in OutputType)
            {
                switch (OType)
                {
                    case "STDOUT":
                        LDAPClass.ADRSTDOUT = true;
                        continue;
                    case "CSV":
                        LDAPClass.ADRCSV = true;
                        break;
                }
            }
            if (LDAPClass.ADRCSV == true)
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@FilePath))
                {
                    string HeaderRow = "Group,ManagedBy,whenCreated,whenChanged,Description,SID,DistinguishedName,CanonicalName";
                    file.WriteLine(HeaderRow);
                }
            }
            Console.WriteLine("[*] Total Groups: " + AdGroups.Length);
            runProcessor(AdGroups, numOfThreads, flushCnt, "Groups", "CSV");
            LDAPClass.ADRSTDOUT = false;
            LDAPClass.ADRCSV = false;
        }

        public static void GroupMemberParser(Object[] AdGroupMembers, string FilePath, int numOfThreads, int flushCnt, String[] OutputType)
        {
            LDAPClass.FilePath = FilePath;

            foreach (string OType in OutputType)
            {
                switch (OType)
                {
                    case "STDOUT":
                        LDAPClass.ADRSTDOUT = true;
                        continue;
                    case "CSV":
                        LDAPClass.ADRCSV = true;
                        break;
                }
            }
            if (LDAPClass.ADRCSV == true)
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@FilePath))
                {
                    string HeaderRow = "Group Name, Member UserName, Member Name, AccountType";
                    file.WriteLine(HeaderRow);
                }
            }
            Console.WriteLine("[*] Total GroupMember Objects: " + AdGroupMembers.Length);
            runProcessor(AdGroupMembers, numOfThreads, flushCnt, "GroupMembers", "CSV");
            LDAPClass.ADRSTDOUT = false;
            LDAPClass.ADRCSV = false;
        }

        public static void ComputerParser(Object[] AdComputers, DateTime Date1, string FilePath, int numOfThreads, int flushCnt, String[] OutputType)
        {
            LDAPClass.Date1 = Date1;
            LDAPClass.FilePath = FilePath;

            foreach (string OType in OutputType)
            {
                switch (OType)
                {
                    case "STDOUT":
                        LDAPClass.ADRSTDOUT = true;
                        continue;
                    case "CSV":
                        LDAPClass.ADRCSV = true;
                        break;
                }
            }
            if (LDAPClass.ADRCSV == true)
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@FilePath))
                {
                    string HeaderRow = "Name,DNSHostName,Enabled,OperatingSystem,Days Since Last Logon,Days Since Last Password Change,Trusted for Delegation,Trusted to Auth for Delegation,Username,Primary Group ID,Description,Password LastSet,Last Logon Date,whenCreated,whenChanged,Distinguished Name";
                    file.WriteLine(HeaderRow);
                }
            }
            Console.WriteLine("[*] Total Computers: " + AdComputers.Length);
            runProcessor(AdComputers, numOfThreads, flushCnt, "Computers", "CSV");
            LDAPClass.ADRSTDOUT = false;
            LDAPClass.ADRCSV = false;
        }

        public static void ComputerSPNParser(Object[] AdComputers, string FilePath, int numOfThreads, int flushCnt, String[] OutputType)
        {
            LDAPClass.FilePath = FilePath;

            foreach (string OType in OutputType)
            {
                switch (OType)
                {
                    case "STDOUT":
                        LDAPClass.ADRSTDOUT = true;
                        continue;
                    case "CSV":
                        LDAPClass.ADRCSV = true;
                        break;
                }
            }
            if (LDAPClass.ADRCSV == true)
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@FilePath))
                {
                    string HeaderRow = "Name,Service,Host";
                    file.WriteLine(HeaderRow);
                }
            }
            runProcessor(AdComputers, numOfThreads, flushCnt, "ComputerSPNs", "CSV");
            LDAPClass.ADRSTDOUT = false;
            LDAPClass.ADRCSV = false;
        }

        static void runProcessor(Object[] arrayToProcess, int numOfThreads, int flushCnt, string processorType, String resultHandlerType)
        {
            int totalRecords = arrayToProcess.Length;
            //Console.WriteLine(String.Format("Running {0} records over {1} threads, flushing every {2} records",
            //    totalRecords, numOfThreads, (flushCnt < 0 ? "NEVER" : flushCnt.ToString())));
            IRecordProcessor recordProcessor = recordProcessorFactory(processorType);
            IResultsHandler resultsHandler = resultHandlerFactory(resultHandlerType, flushCnt);
            int numberOfRecordsPerThread = totalRecords / numOfThreads;
            int remainders = totalRecords % numOfThreads;

            Thread[] threads = new Thread[numOfThreads];
            for (int i = 0; i < numOfThreads; i++)
            {
                int numberOfRecordsToProcess = numberOfRecordsPerThread;
                if (i == (numOfThreads - 1))
                {
                    //last thread, do the remaining records
                    numberOfRecordsToProcess += remainders;
                }

                //split the full array into chunks to be given to different threads
                Object[] sliceToProcess = new Object[numberOfRecordsToProcess];
                Array.Copy(arrayToProcess, i * numberOfRecordsPerThread, sliceToProcess, 0, numberOfRecordsToProcess);
                ProcessorThread processorThread = new ProcessorThread(i, recordProcessor, resultsHandler, sliceToProcess);
                threads[i] = new Thread(processorThread.processThreadRecords);
                threads[i].Start();
            }
            foreach (Thread t in threads)
            {
                t.Join();
            }

            resultsHandler.finalise();
        }

        static IRecordProcessor recordProcessorFactory(String name)
        {
            switch (name)
            {
                case "Users":
                    return new UserRecordProcessor();
                case "UserSPNs":
                    return new UserSPNRecordProcessor();
                case "Groups":
                    return new GroupRecordProcessor();
                case "GroupMembers":
                    return new GroupMemberRecordProcessor();
                case "Computers":
                    return new ComputerRecordProcessor();
                case "ComputerSPNs":
                    return new ComputerSPNRecordProcessor();
            }
            throw new ArgumentException("Invalid processor type " + name);
        }

        static IResultsHandler resultHandlerFactory(String name, int flushCnt)
        {
            switch (name)
            {
                case "CSV":
                    return new CsvResultsHandler(flushCnt);
                case "TXT":
                    return new TxtResultsHandler(flushCnt);
            }
            throw new ArgumentException("Invalid processor type " + name);
        }

        class ProcessorThread
        {
            readonly int id;
            readonly IRecordProcessor recordProcessor;
            readonly IResultsHandler resultsHandler;
            readonly Object[] objectsToBeProcessed;

            public ProcessorThread(int id, IRecordProcessor recordProcessor, IResultsHandler resultsHandler, Object[] objectsToBeProcessed)
            {
                this.recordProcessor = recordProcessor;
                this.id = id;
                this.resultsHandler = resultsHandler;
                this.objectsToBeProcessed = objectsToBeProcessed;
            }

            public void processThreadRecords()
            {
                for (int i = 0; i < objectsToBeProcessed.Length; i++)
                {
                    Object[] result = recordProcessor.processRecord(objectsToBeProcessed[i]);
                    resultsHandler.processResults(result); //this is a thread safe operation
                }
            }
        }

        //The interface and implmentation class used to process a record (this implemmentation just returns a log type string)

        interface IRecordProcessor
        {
            Object[] processRecord(Object record);
        }

        class UserRecordProcessor : IRecordProcessor
        {
            public Object[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdUser = (SearchResult) record;
                    bool? Enabled = null;
                    bool? PasswordNeverExpires = null;
                    bool? AccountLockedOut = null;
                    bool? ReversiblyEncrypedPassword = null;
                    bool? PasswordNotRequired = null;
                    bool? TrustedforDelegation = null;
                    bool? TrustedtoAuthforDelegation = null;
                    bool? DoesNotRequirePreAuth = null;
                    // When the user is not allowed to query the UserAccountControl attribute.
                    if (AdUser.Properties["useraccountcontrol"].Count != 0)
                    {
                        var userFlags = (UACFlags) AdUser.Properties["useraccountcontrol"][0];
                        Enabled = !((userFlags & UACFlags.ACCOUNTDISABLE) == UACFlags.ACCOUNTDISABLE);
                        PasswordNeverExpires = (userFlags & UACFlags.DONT_EXPIRE_PASSWD) == UACFlags.DONT_EXPIRE_PASSWD;
                        AccountLockedOut = (userFlags & UACFlags.LOCKOUT) == UACFlags.LOCKOUT;
                        ReversiblyEncrypedPassword = (userFlags & UACFlags.ENCRYPTED_TEXT_PASSWORD_ALLOWED) == UACFlags.ENCRYPTED_TEXT_PASSWORD_ALLOWED;
                        PasswordNotRequired = (userFlags & UACFlags.PASSWD_NOTREQD) == UACFlags.PASSWD_NOTREQD;
                        TrustedforDelegation = (userFlags & UACFlags.TRUSTED_FOR_DELEGATION) == UACFlags.TRUSTED_FOR_DELEGATION;
                        TrustedtoAuthforDelegation = (userFlags & UACFlags.TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION) == UACFlags.TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION;
                        DoesNotRequirePreAuth = (userFlags & UACFlags.DONT_REQUIRE_PREAUTH) == UACFlags.DONT_REQUIRE_PREAUTH;
                    }
                    bool MustChangePasswordatLogon = false;
                    int DaysSinceLastPasswordChange = -1;
                    bool PasswordNotChangedafterMaxAge = false;
                    bool NeverLoggedIn = false;
                    int DaysSinceLastLogon = -1;
                    bool Dormant = false;
                    DateTime PasswordLastSet = new DateTime();
                    DateTime LastLogonDate = new DateTime();
                    bool CannotChangePassword = CannotChangePasswordDict[Convert.ToString(AdUser.Properties["samaccountname"][0])];
                    if (AdUser.Properties["pwdLastSet"].Count != 0)
                    {
                        if (Convert.ToString(AdUser.Properties["pwdlastset"][0]) == "0")
                        {
                            MustChangePasswordatLogon = true;
                        }
                        else
                        {
                            PasswordLastSet = DateTime.FromFileTime((long)(AdUser.Properties["pwdLastSet"][0]));
                            DaysSinceLastPasswordChange = Math.Abs((Date1 - PasswordLastSet).Days);
                            if (DaysSinceLastPasswordChange > PassMaxAge)
                            {
                                PasswordNotChangedafterMaxAge = true;
                            }
                        }
                    }
                    if (AdUser.Properties["lastlogontimestamp"].Count != 0)
                    {
                        LastLogonDate = DateTime.FromFileTime((long)(AdUser.Properties["lastlogontimestamp"][0]));
                        DaysSinceLastLogon = Math.Abs((Date1 - LastLogonDate).Days);
                        if (DaysSinceLastLogon > DormantTimeSpan)
                        {
                            Dormant = true;
                        }
                    }
                    else
                    {
                        NeverLoggedIn = true;
                    }
                    string SIDHistory = "";
                    if (AdUser.Properties["sidhistory"].Count >= 1)
                    {
                        string sids = "";
                        for (int i = 0; i < AdUser.Properties["sidhistory"].Count; i++)
                        {
                            var history = AdUser.Properties["sidhistory"][i];
                            sids = sids + "," + Convert.ToString(new SecurityIdentifier((byte[])history, 0));
                        }
                        SIDHistory = sids.TrimStart(',');
                    }
                    if (LDAPClass.ADRSTDOUT == true)
                    {
                        Console.WriteLine("\nName{0,33}{2}\nUserName{0,29}{3}\nEnabled{0,30}{4}\nCannot Change Password{0,15}{5}\nPassword Never Expires{0,15}{6}\nMust Change Password at Logon{0,8}{7}\nDays Since Last Password Change{0,6}{8}\nPassword Not Changed after Max Age{0}{9}\nAccount Locked Out{0,19}{10}\nNever Logged in{0,22}{11}\nDays Since Last Logon{0,16}{12}\nDormant (> {1} days){0,18}{13}\nReversibly Encryped Password{0,9}{14}\nPassword Not Required{0,16}{15}\nTrusted for Delegation{0,15}{16}\nTrusted to Auth for Delegation{0,7}{17}\nDoes Not Require Pre Auth{0,12}{18}\nLogon Workstations{0,19}{19}\nAdminCount{0,27}{20}\nPrimary GroupID{0,22}{21}\nSID{0,34}{22}\nSIDHistory{0,27}{23}\nDescription{0,26}{24}\nPassword LastSet{0,21}{25}\nLast Logon Date{0,22}{26}\nWhen Created{0,25}{27}\nWhen Changed{0,25}{28}\nDistinguishedName{0,20}{29}\nCanonicalName{0,24}{30}\n", " : ", DormantTimeSpan, (AdUser.Properties["name"].Count != 0 ? AdUser.Properties["name"][0] : ""), (AdUser.Properties["samaccountname"].Count != 0 ? AdUser.Properties["samaccountname"][0] : ""), Enabled, CannotChangePassword, PasswordNeverExpires, MustChangePasswordatLogon, DaysSinceLastPasswordChange, PasswordNotChangedafterMaxAge, AccountLockedOut, NeverLoggedIn, DaysSinceLastLogon, Dormant, ReversiblyEncrypedPassword, PasswordNotRequired, TrustedforDelegation, TrustedtoAuthforDelegation, DoesNotRequirePreAuth, (AdUser.Properties["userworkstations"].Count != 0 ? AdUser.Properties["userworkstations"][0] : ""), (AdUser.Properties["admincount"].Count != 0 ? AdUser.Properties["admincount"][0] : ""), (AdUser.Properties["primarygroupid"].Count != 0 ? AdUser.Properties["primarygroupid"][0] : ""), Convert.ToString(new SecurityIdentifier((byte[])AdUser.Properties["objectSID"][0], 0)), SIDHistory, (AdUser.Properties["Description"].Count != 0 ? AdUser.Properties["Description"][0] : ""), PasswordLastSet, LastLogonDate, (AdUser.Properties["whencreated"].Count != 0 ? AdUser.Properties["whencreated"][0] : ""), (AdUser.Properties["whenchanged"].Count != 0 ? AdUser.Properties["whenchanged"][0] : ""), (AdUser.Properties["distinguishedname"].Count != 0 ? AdUser.Properties["distinguishedname"][0] : ""), (AdUser.Properties["canonicalname"].Count != 0 ? AdUser.Properties["canonicalname"][0] : ""));
                    }
                    return new Object[] { (AdUser.Properties["name"].Count != 0 ? AdUser.Properties["name"][0] : ""), (AdUser.Properties["samaccountname"].Count != 0 ? AdUser.Properties["samaccountname"][0] : ""), Enabled, CannotChangePassword, PasswordNeverExpires, MustChangePasswordatLogon, DaysSinceLastPasswordChange, PasswordNotChangedafterMaxAge, AccountLockedOut, NeverLoggedIn, DaysSinceLastLogon, Dormant, ReversiblyEncrypedPassword, PasswordNotRequired, TrustedforDelegation, TrustedtoAuthforDelegation, DoesNotRequirePreAuth, (AdUser.Properties["userworkstations"].Count != 0 ? AdUser.Properties["userworkstations"][0] : ""), (AdUser.Properties["admincount"].Count != 0 ? AdUser.Properties["admincount"][0] : ""), (AdUser.Properties["primarygroupid"].Count != 0 ? AdUser.Properties["primarygroupid"][0] : ""), Convert.ToString(new SecurityIdentifier((byte[])AdUser.Properties["objectSID"][0], 0)), SIDHistory, (AdUser.Properties["Description"].Count != 0 ? AdUser.Properties["Description"][0] : ""), PasswordLastSet, LastLogonDate, (AdUser.Properties["whencreated"].Count != 0 ? AdUser.Properties["whencreated"][0] : ""), (AdUser.Properties["whenchanged"].Count != 0 ? AdUser.Properties["whenchanged"][0] : ""), (AdUser.Properties["distinguishedname"].Count != 0 ? AdUser.Properties["distinguishedname"][0] : ""), (AdUser.Properties["canonicalname"].Count != 0 ? AdUser.Properties["canonicalname"][0] : "") };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new Object[] { };
                }
            }
        }

        class UserSPNRecordProcessor : IRecordProcessor
        {
            public Object[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdUser = (SearchResult) record;
                    List<Object> SPNList = new List<Object>();
                    DateTime PasswordLastSet = DateTime.FromFileTime((long)(AdUser.Properties["pwdLastSet"][0]));
                    String Description = (AdUser.Properties["Description"].Count != 0 ? Convert.ToString(AdUser.Properties["Description"][0]) : "");
                    foreach (String SPN in AdUser.Properties["serviceprincipalname"])
                    {
                        String[] SPNArray = SPN.Split('/');
                        SPNList.Add(new Object[] { AdUser.Properties["name"][0], AdUser.Properties["samaccountname"][0], SPNArray[0], SPNArray[1], PasswordLastSet, Description });
                    }
                    if (LDAPClass.ADRSTDOUT == true)
                    {
                        foreach (Object[] Obj in SPNList.ToArray())
                        {
                            Console.WriteLine("Name{0,16}{1}\nUsername{0,12}{2}\nService{0,13}{3}\nHost{0,16}{4}\nPassword Last Set{0}{5}\nDescription{0,9}{6}\n", " : ", Obj[0], Obj[1], Obj[2], Obj[3], Obj[4], Obj[5]);
                        }
                    }
                    return SPNList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new Object[] { };
                }
            }
        }

        class GroupRecordProcessor : IRecordProcessor
        {
            public Object[] processRecord(Object record)
            {
                try
                {

                    SearchResult AdGroup = (SearchResult) record;
                    string ManagedByValue = AdGroup.Properties["managedby"].Count != 0 ? Convert.ToString(AdGroup.Properties["managedby"][0]) : "";
                    string ManagedBy = "";
                    if (AdGroup.Properties["managedBy"].Count != 0)
                    {
                        ManagedBy = (ManagedByValue.Split(',')[0]).Split('=')[1];
                    }
                    if (LDAPClass.ADRSTDOUT == true)
                    {
                        Console.WriteLine("Group{0,15}{1}\nManagedBy{0,11}{2}\nwhenCreated{0,9}{3}\nwhenChanged{0,9}{4}\nDescription{0,9}{5}\nSID{0,17}{6}\nDistinguishedName{0}{7}\nCanonicalName{0,7}{8}\n", " : ", AdGroup.Properties["samaccountname"][0], ManagedBy, AdGroup.Properties["whencreated"][0], AdGroup.Properties["whenchanged"][0], (AdGroup.Properties["Description"].Count != 0 ? AdGroup.Properties["Description"][0] : ""), Convert.ToString(new SecurityIdentifier((byte[])AdGroup.Properties["objectSID"][0], 0)), AdGroup.Properties["distinguishedname"][0], AdGroup.Properties["canonicalname"][0]);
                    }
                    return new Object[] { AdGroup.Properties["samaccountname"][0], ManagedBy, AdGroup.Properties["whencreated"][0], AdGroup.Properties["whenchanged"][0], (AdGroup.Properties["Description"].Count != 0 ? AdGroup.Properties["Description"][0] : ""), Convert.ToString(new SecurityIdentifier((byte[])AdGroup.Properties["objectSID"][0], 0)), AdGroup.Properties["distinguishedname"][0], AdGroup.Properties["canonicalname"][0] };
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new Object[] { };
                }
            }
        }

        class GroupMemberRecordProcessor : IRecordProcessor
        {
            public Object[] processRecord(Object record)
            {
                try
                {
                    // https://github.com/BloodHoundAD/BloodHound/blob/master/PowerShell/BloodHound.ps1
                    SearchResult AdGroup = (SearchResult) record;
                    List<Object> GroupsList = new List<Object>();
                    string SamAccountType = AdGroup.Properties["samaccounttype"].Count != 0 ? Convert.ToString(AdGroup.Properties["samaccounttype"][0]) : "";
                    string AccountType = "";
                    string GroupName = "";
                    string MemberUserName = "-";
                    string MemberName = "";
                    if (Groups.Contains(SamAccountType))
                    {
                        AccountType = "group";
                        MemberName = ((Convert.ToString(AdGroup.Properties["DistinguishedName"][0])).Split(',')[0]).Split('=')[1];
                        foreach (String GroupMember in AdGroup.Properties["memberof"])
                        {
                            GroupName = ((Convert.ToString(GroupMember)).Split(',')[0]).Split('=')[1];
                            GroupsList.Add(new Object[] { GroupName, MemberUserName, MemberName, AccountType });
                        }
                    }
                    if (Users.Contains(SamAccountType))
                    {
                        AccountType = "user";
                        MemberName = ((Convert.ToString(AdGroup.Properties["DistinguishedName"][0])).Split(',')[0]).Split('=')[1];
                        MemberUserName = Convert.ToString(AdGroup.Properties["sAMAccountName"][0]);
                        foreach (String GroupMember in AdGroup.Properties["memberof"])
                        {
                            GroupName = ((Convert.ToString(GroupMember)).Split(',')[0]).Split('=')[1];
                            GroupsList.Add(new Object[] { GroupName, MemberUserName, MemberName, AccountType });
                        }
                    }
                    if (Computers.Contains(SamAccountType))
                    {
                        AccountType = "computer";
                        MemberName = ((Convert.ToString(AdGroup.Properties["DistinguishedName"][0])).Split(',')[0]).Split('=')[1];
                        MemberUserName = Convert.ToString(AdGroup.Properties["sAMAccountName"][0]);
                        foreach (String GroupMember in AdGroup.Properties["memberof"])
                        {
                            GroupName = ((Convert.ToString(GroupMember)).Split(',')[0]).Split('=')[1];
                            GroupsList.Add(new Object[] { GroupName, MemberUserName, MemberName, AccountType });
                        }
                    }
                    if (TrustAccounts.Contains(SamAccountType))
                    {
                        // TO DO
                    }
                    if (LDAPClass.ADRSTDOUT == true)
                    {
                        foreach (Object[] Obj in GroupsList.ToArray())
                        {
                            Console.WriteLine("Group Name{0,8}{1}\nMember UserName{0}{2}\nMember Name{0,7}{3}\nAccountType{0,7}{4}\n", " : ", Obj[0], Obj[1], Obj[2], Obj[3]);
                        }
                    }
                    return GroupsList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine(((SearchResult)record).Properties["DistinguishedName"][0]);
                    Console.WriteLine("{0} Exception caught.", e);
                    return new Object[] { };
                }
            }
        }

        class ComputerRecordProcessor : IRecordProcessor
        {
            public Object[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdComputer = (SearchResult) record;
                    bool? Enabled = null;
                    bool? TrustedforDelegation = null;
                    bool? TrustedtoAuthforDelegation = null;
                    // When the user is not allowed to query the UserAccountControl attribute.
                    if (AdComputer.Properties["useraccountcontrol"].Count != 0)
                    {
                        var userFlags = (UACFlags) AdComputer.Properties["useraccountcontrol"][0];
                        Enabled = !((userFlags & UACFlags.ACCOUNTDISABLE) == UACFlags.ACCOUNTDISABLE);
                        TrustedforDelegation = (userFlags & UACFlags.TRUSTED_FOR_DELEGATION) == UACFlags.TRUSTED_FOR_DELEGATION;
                        TrustedtoAuthforDelegation = (userFlags & UACFlags.TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION) == UACFlags.TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION;
                    }
                    int DaysSinceLastPasswordChange = -1;
                    int DaysSinceLastLogon = -1;
                    DateTime LastLogonDate = new DateTime();
                    if (AdComputer.Properties["lastlogontimestamp"].Count != 0)
                    {
                        LastLogonDate = DateTime.FromFileTime((long)(AdComputer.Properties["lastlogontimestamp"][0]));
                        DaysSinceLastLogon = Math.Abs((Date1 - LastLogonDate).Days);
                        if (DaysSinceLastLogon >= 152246)
                        {
                            DaysSinceLastLogon = -1;
                        }
                    }
                    DateTime PasswordLastSet = DateTime.FromFileTime((long)(AdComputer.Properties["pwdLastSet"][0]));
                    if (AdComputer.Properties["pwdLastSet"].Count != 0)
                    {
                        DaysSinceLastPasswordChange = Math.Abs((Date1 - PasswordLastSet).Days);
                        if (DaysSinceLastPasswordChange >= 152246)
                        {
                            DaysSinceLastPasswordChange = -1;
                        }
                    }
                    if (LDAPClass.ADRSTDOUT == true)
                    {
                        Console.WriteLine("Name{0,30}{1}\nDNSHostName{0,23}{2}\nEnabled{0,27}{3}\nOperatingSystem{0,19}{4}\nDays Since Last Logon{0,13}{5}\nDays Since Last Password Change{0}{6}\nTrusted for Delegation{0,12}{7}\nTrusted to Auth for Delegation{0,4}{8}\nUsername{0,26}{9}\nPrimary Group ID{0,18}{10}\nDescription{0,23}{11}\nPassword LastSet{0,18}{12}\nLast Logon Date{0,19}{13}\nwhenCreated{0,23}{14}\nwhenChanged{0,23}{15}\nDistinguished Name{0,16}{16}\n", " : ", (AdComputer.Properties["name"].Count != 0 ? AdComputer.Properties["name"][0] : ""), (AdComputer.Properties["dnshostname"].Count != 0 ? AdComputer.Properties["dnshostname"][0] : ""), Enabled, (AdComputer.Properties["operatingsystem"].Count != 0 ? AdComputer.Properties["operatingsystem"][0] : "-"), DaysSinceLastLogon, DaysSinceLastPasswordChange, TrustedforDelegation, TrustedtoAuthforDelegation, (AdComputer.Properties["samaccountname"].Count != 0 ? AdComputer.Properties["samaccountname"][0] : ""), (AdComputer.Properties["primarygroupid"].Count != 0 ? AdComputer.Properties["primarygroupid"][0] : ""), (AdComputer.Properties["Description"].Count != 0 ? AdComputer.Properties["Description"][0] : ""), PasswordLastSet, LastLogonDate, AdComputer.Properties["whencreated"][0], AdComputer.Properties["whenchanged"][0], AdComputer.Properties["distinguishedname"][0]);
                    }
                    return new Object[] { (AdComputer.Properties["name"].Count != 0 ? AdComputer.Properties["name"][0] : ""), (AdComputer.Properties["dnshostname"].Count != 0 ? AdComputer.Properties["dnshostname"][0] : ""), Enabled, (AdComputer.Properties["operatingsystem"].Count != 0 ? AdComputer.Properties["operatingsystem"][0] : "-"), DaysSinceLastLogon, DaysSinceLastPasswordChange, TrustedforDelegation, TrustedtoAuthforDelegation, (AdComputer.Properties["samaccountname"].Count != 0 ? AdComputer.Properties["samaccountname"][0] : ""), (AdComputer.Properties["primarygroupid"].Count != 0 ? AdComputer.Properties["primarygroupid"][0] : ""), (AdComputer.Properties["Description"].Count != 0 ? AdComputer.Properties["Description"][0] : ""), PasswordLastSet, LastLogonDate, AdComputer.Properties["whencreated"][0], AdComputer.Properties["whenchanged"][0], AdComputer.Properties["distinguishedname"][0] };
                }
                catch (Exception e)
                {
                    Console.WriteLine(((SearchResult)record).Properties["name"][0]);
                    Console.WriteLine("{0} Exception caught.", e);
                    return new Object[] { };
                }
            }
        }

        class ComputerSPNRecordProcessor : IRecordProcessor
        {
            public Object[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdComputer = (SearchResult) record;
                    List<Object> SPNList = new List<Object>();
                    foreach (String SPN in AdComputer.Properties["serviceprincipalname"])
                    {
                        String[] SPNArray = SPN.Split('/');
                        SPNList.Add(new Object[] { AdComputer.Properties["name"][0], SPNArray[0], SPNArray[1] });
                    }
                    if (LDAPClass.ADRSTDOUT == true)
                    {
                        foreach (Object[] Obj in SPNList.ToArray())
                        {
                            Console.WriteLine("Name{0,6}{1}\nService{0}{2}\nHost{0,6}{3}\n", " : ", Obj[0], Obj[1], Obj[2]);
                        }
                    }
                    return SPNList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return new Object[] { };
                }
            }
        }

        //The interface and implmentation class used to handle the results (this implementation just writes the strings to a file)

        interface IResultsHandler
        {
            void processResults(Object[] t);

            void finalise();
        }

        abstract class SimpleResultsHandler : IResultsHandler
        {
            private Object lockObj = new Object();
            private List<String> processed = new List<String>();
            private readonly int flushCnt;

            public SimpleResultsHandler(int flushCnt)
            {
                this.flushCnt = flushCnt;
            }

            public void processResults(Object[] results)
            {
                lock (lockObj)
                {
                    if (results.Length != 0)
                    {
                        if (results[0] is System.Object[])
                        {
                            for (var i = 0; i < results.Length; i++)
                            {
                                processed.Add(convertObject((Object[])results[i]));
                            }
                        }
                        else
                        {
                            processed.Add(convertObject(results));
                        }
                        if (flushCnt > 0)
                        {
                            if (processed.Count >= flushCnt)
                            {
                                writeFile();
                            }
                        }
                    }
                }
            }

            public void finalise()
            {
                writeFile();
            }

            private void writeFile()
            {
                lock (lockObj)
                {
                    if (LDAPClass.ADRCSV == true)
                    {
                        using (StreamWriter outputFile = new StreamWriter(@LDAPClass.FilePath, true))
                        {
                            outputFile.Write(String.Join("\r\n", processed.ToArray()));
                        }
                    }
                    processed.Clear();
                }
            }

            protected abstract String convertObject(Object[] resultsObject);
        }


        class CsvResultsHandler : SimpleResultsHandler
        {
            public CsvResultsHandler(int flushCnt) : base(flushCnt)
            {
            }

            protected override String convertObject(Object[] resultsObject)
            {
                return createCsvLine(resultsObject);
            }

            static String createCsvLine(Object[] resultsObject)
            {
                try
                {
                    // No String.Join(String, Object[]) in CLR 2.0.50727 (Windows 7)
                    String[] row = new String[resultsObject.Length];
                    for (int i=0; i < resultsObject.Length; i++)
                    {
                        row[i] = CleanString(Convert.ToString(resultsObject[i]));
                    }
                    return "\"" + String.Join("\",\"", row) + "\"";
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return "";
                }
            }
        }

        class TxtResultsHandler : SimpleResultsHandler
        {
            public TxtResultsHandler(int flushCnt) : base(flushCnt)
            {
            }

            protected override String convertObject(Object[] resultsObject)
            {
                try
                {
                    // No String.Join(String, Object[]) in CLR 2.0.50727 (Windows 7)
                    String[] row = new String[resultsObject.Length];
                    for (int i=0; i < resultsObject.Length; i++)
                    {
                        row[i] = CleanString(Convert.ToString(resultsObject[i]));
                    }
                    return String.Join("\t", row);
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} Exception caught.", e);
                    return "";
                }
            }
        }
    }
}
"@

#Add-Type -TypeDefinition $Source -ReferencedAssemblies ([System.String[]]@(([system.reflection.assembly]::LoadWithPartialName("Microsoft.ActiveDirectory.Management")).Location,([system.reflection.assembly]::LoadWithPartialName("System.DirectoryServices")).Location))

# modified version from https://github.com/vletoux/SmbScanner/blob/master/smbscanner.ps1
$PingCastleSMBScannerSource = @"
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Runtime.InteropServices;
using System.Management.Automation;

namespace PingCastle.Scanners
{
    public class SmbScanner
	{
        [StructLayout(LayoutKind.Explicit)]
		struct SMB_Header {
			[FieldOffset(0)]
			public UInt32 Protocol;
			[FieldOffset(4)]
			public byte Command;
			[FieldOffset(5)]
			public int Status;
			[FieldOffset(9)]
			public byte  Flags;
			[FieldOffset(10)]
			public UInt16 Flags2;
			[FieldOffset(12)]
			public UInt16 PIDHigh;
			[FieldOffset(14)]
			public UInt64 SecurityFeatures;
			[FieldOffset(22)]
			public UInt16 Reserved;
			[FieldOffset(24)]
			public UInt16 TID;
			[FieldOffset(26)]
			public UInt16 PIDLow;
			[FieldOffset(28)]
			public UInt16 UID;
			[FieldOffset(30)]
			public UInt16 MID;
		};
		// https://msdn.microsoft.com/en-us/library/cc246529.aspx
		[StructLayout(LayoutKind.Explicit)]
		struct SMB2_Header {
			[FieldOffset(0)]
			public UInt32 ProtocolId;
			[FieldOffset(4)]
			public UInt16 StructureSize;
			[FieldOffset(6)]
			public UInt16 CreditCharge;
			[FieldOffset(8)]
			public UInt32 Status; // to do SMB3
			[FieldOffset(12)]
			public UInt16 Command;
			[FieldOffset(14)]
			public UInt16 CreditRequest_Response;
			[FieldOffset(16)]
			public UInt32 Flags;
			[FieldOffset(20)]
			public UInt32 NextCommand;
			[FieldOffset(24)]
			public UInt64 MessageId;
			[FieldOffset(32)]
			public UInt32 Reserved;
			[FieldOffset(36)]
			public UInt32 TreeId;
			[FieldOffset(40)]
			public UInt64 SessionId;
			[FieldOffset(48)]
			public UInt64 Signature1;
			[FieldOffset(56)]
			public UInt64 Signature2;
		}
        [StructLayout(LayoutKind.Explicit)]
		struct SMB2_NegotiateRequest
		{
			[FieldOffset(0)]
			public UInt16 StructureSize;
			[FieldOffset(2)]
			public UInt16 DialectCount;
			[FieldOffset(4)]
			public UInt16 SecurityMode;
			[FieldOffset(6)]
			public UInt16 Reserved;
			[FieldOffset(8)]
			public UInt32 Capabilities;
			[FieldOffset(12)]
			public Guid ClientGuid;
			[FieldOffset(28)]
			public UInt64 ClientStartTime;
			[FieldOffset(36)]
			public UInt16 DialectToTest;
		}
		const int SMB_COM_NEGOTIATE	= 0x72;
		const int SMB2_NEGOTIATE = 0;
		const int SMB_FLAGS_CASE_INSENSITIVE = 0x08;
		const int SMB_FLAGS_CANONICALIZED_PATHS = 0x10;
		const int SMB_FLAGS2_LONG_NAMES					= 0x0001;
		const int SMB_FLAGS2_EAS							= 0x0002;
		const int SMB_FLAGS2_SECURITY_SIGNATURE_REQUIRED	= 0x0010	;
		const int SMB_FLAGS2_IS_LONG_NAME					= 0x0040;
		const int SMB_FLAGS2_ESS							= 0x0800;
		const int SMB_FLAGS2_NT_STATUS					= 0x4000;
		const int SMB_FLAGS2_UNICODE						= 0x8000;
		const int SMB_DB_FORMAT_DIALECT = 0x02;
		static byte[] GenerateSmbHeaderFromCommand(byte command)
		{
			SMB_Header header = new SMB_Header();
			header.Protocol = 0x424D53FF;
			header.Command = command;
			header.Status = 0;
			header.Flags = SMB_FLAGS_CASE_INSENSITIVE | SMB_FLAGS_CANONICALIZED_PATHS;
			header.Flags2 = SMB_FLAGS2_LONG_NAMES | SMB_FLAGS2_EAS | SMB_FLAGS2_SECURITY_SIGNATURE_REQUIRED | SMB_FLAGS2_IS_LONG_NAME | SMB_FLAGS2_ESS | SMB_FLAGS2_NT_STATUS | SMB_FLAGS2_UNICODE;
			header.PIDHigh = 0;
			header.SecurityFeatures = 0;
			header.Reserved = 0;
			header.TID = 0xffff;
			header.PIDLow = 0xFEFF;
			header.UID = 0;
			header.MID = 0;
			return getBytes(header);
		}
		static byte[] GenerateSmb2HeaderFromCommand(byte command)
		{
			SMB2_Header header = new SMB2_Header();
			header.ProtocolId = 0x424D53FE;
			header.Command = command;
			header.StructureSize = 64;
			header.Command = command;
			header.MessageId = 0;
			header.Reserved = 0xFEFF;
			return getBytes(header);
		}
		static byte[] getBytes(object structure)
		{
			int size = Marshal.SizeOf(structure);
			byte[] arr = new byte[size];
			IntPtr ptr = Marshal.AllocHGlobal(size);
			Marshal.StructureToPtr(structure, ptr, true);
			Marshal.Copy(ptr, arr, 0, size);
			Marshal.FreeHGlobal(ptr);
			return arr;
		}
		static byte[] getDialect(string dialect)
		{
			byte[] dialectBytes = Encoding.ASCII.GetBytes(dialect);
			byte[] output = new byte[dialectBytes.Length + 2];
			output[0] = 2;
			output[output.Length - 1] = 0;
			Array.Copy(dialectBytes, 0, output, 1, dialectBytes.Length);
			return output;
		}
		static byte[] GetNegotiateMessage(byte[] dialect)
		{
			byte[] output = new byte[dialect.Length + 3];
			output[0] = 0;
			output[1] = (byte) dialect.Length;
			output[2] = 0;
			Array.Copy(dialect, 0, output, 3, dialect.Length);
			return output;
		}
		// MS-SMB2  2.2.3 SMB2 NEGOTIATE Request
		static byte[] GetNegotiateMessageSmbv2(int DialectToTest)
		{
			SMB2_NegotiateRequest request = new SMB2_NegotiateRequest();
			request.StructureSize = 36;
			request.DialectCount = 1;
			request.SecurityMode = 1; // signing enabled
			request.ClientGuid = Guid.NewGuid();
			request.DialectToTest = (UInt16) DialectToTest;
			return getBytes(request);
		}
		static byte[] GetNegotiatePacket(byte[] header, byte[] smbPacket)
		{
			byte[] output = new byte[smbPacket.Length + header.Length + 4];
			output[0] = 0;
			output[1] = 0;
			output[2] = 0;
			output[3] = (byte)(smbPacket.Length + header.Length);
			Array.Copy(header, 0, output, 4, header.Length);
			Array.Copy(smbPacket, 0, output, 4 + header.Length, smbPacket.Length);
			return output;
		}
		public static bool DoesServerSupportDialect(string server, string dialect)
		{
			Trace.WriteLine("Checking " + server + " for SMBV1 dialect " + dialect);
			TcpClient client = new TcpClient();
			try
			{
				client.Connect(server, 445);
			}
			catch (Exception)
			{
				throw new Exception("port 445 is closed on " + server);
			}
			try
			{
				NetworkStream stream = client.GetStream();
				byte[] header = GenerateSmbHeaderFromCommand(SMB_COM_NEGOTIATE);
				byte[] dialectEncoding = getDialect(dialect);
				byte[] negotiatemessage = GetNegotiateMessage(dialectEncoding);
				byte[] packet = GetNegotiatePacket(header, negotiatemessage);
				stream.Write(packet, 0, packet.Length);
				stream.Flush();
				byte[] netbios = new byte[4];
				if (stream.Read(netbios, 0, netbios.Length) != netbios.Length)
                {
                    return false;
                }
				byte[] smbHeader = new byte[Marshal.SizeOf(typeof(SMB_Header))];
				if (stream.Read(smbHeader, 0, smbHeader.Length) != smbHeader.Length)
                {
                    return false;
                }
				byte[] negotiateresponse = new byte[3];
				if (stream.Read(negotiateresponse, 0, negotiateresponse.Length) != negotiateresponse.Length)
                {
                    return false;
                }
				if (negotiateresponse[1] == 0 && negotiateresponse[2] == 0)
				{
					Trace.WriteLine("Checking " + server + " for SMBV1 dialect " + dialect + " = Supported");
					return true;
				}
				Trace.WriteLine("Checking " + server + " for SMBV1 dialect " + dialect + " = Not supported");
				return false;
			}
			catch (Exception)
			{
				throw new ApplicationException("Smb1 is not supported on " + server);
			}
		}
		public static bool DoesServerSupportDialectWithSmbV2(string server, int dialect, bool checkSMBSigning)
		{
			Trace.WriteLine("Checking " + server + " for SMBV2 dialect 0x" + dialect.ToString("X2"));
			TcpClient client = new TcpClient();
			try
			{
				client.Connect(server, 445);
			}
			catch (Exception)
			{
				throw new Exception("port 445 is closed on " + server);
			}
			try
			{
				NetworkStream stream = client.GetStream();
				byte[] header = GenerateSmb2HeaderFromCommand(SMB2_NEGOTIATE);
				byte[] negotiatemessage = GetNegotiateMessageSmbv2(dialect);
				byte[] packet = GetNegotiatePacket(header, negotiatemessage);
				stream.Write(packet, 0, packet.Length);
				stream.Flush();
				byte[] netbios = new byte[4];
				if( stream.Read(netbios, 0, netbios.Length) != netbios.Length)
                {
                    return false;
                }
				byte[] smbHeader = new byte[Marshal.SizeOf(typeof(SMB2_Header))];
				if (stream.Read(smbHeader, 0, smbHeader.Length) != smbHeader.Length)
                {
                    return false;
                }
				if (smbHeader[8] != 0 || smbHeader[9] != 0 || smbHeader[10] != 0 || smbHeader[11] != 0)
				{
					Trace.WriteLine("Checking " + server + " for SMBV2 dialect 0x" + dialect.ToString("X2") + " = Not supported via error code");
					return false;
				}
				byte[] negotiateresponse = new byte[6];
				if (stream.Read(negotiateresponse, 0, negotiateresponse.Length) != negotiateresponse.Length)
                {
                    return false;
                }
                if (checkSMBSigning)
                {
                    // https://support.microsoft.com/en-in/help/887429/overview-of-server-message-block-signing
                    // https://msdn.microsoft.com/en-us/library/cc246561.aspx
				    if (negotiateresponse[2] == 3)
				    {
					    Trace.WriteLine("Checking " + server + " for SMBV2 SMB Signing dialect 0x" + dialect.ToString("X2") + " = Supported");
					    return true;
				    }
                    else
                    {
                        return false;
                    }
                }
				int selectedDialect = negotiateresponse[5] * 0x100 + negotiateresponse[4];
				if (selectedDialect == dialect)
				{
					Trace.WriteLine("Checking " + server + " for SMBV2 dialect 0x" + dialect.ToString("X2") + " = Supported");
					return true;
				}
				Trace.WriteLine("Checking " + server + " for SMBV2 dialect 0x" + dialect.ToString("X2") + " = Not supported via not returned dialect");
				return false;
			}
			catch (Exception)
			{
				throw new ApplicationException("Smb2 is not supported on " + server);
			}
		}
		public static bool SupportSMB1(string server)
		{
			try
			{
				return DoesServerSupportDialect(server, "NT LM 0.12");
			}
			catch (Exception)
			{
				return false;
			}
		}
		public static bool SupportSMB2(string server)
		{
			try
			{
				return (DoesServerSupportDialectWithSmbV2(server, 0x0202, false) || DoesServerSupportDialectWithSmbV2(server, 0x0210, false));
			}
			catch (Exception)
			{
				return false;
			}
		}
		public static bool SupportSMB3(string server)
		{
			try
			{
				return (DoesServerSupportDialectWithSmbV2(server, 0x0300, false) || DoesServerSupportDialectWithSmbV2(server, 0x0302, false) || DoesServerSupportDialectWithSmbV2(server, 0x0311, false));
			}
			catch (Exception)
			{
				return false;
			}
		}
		public static string Name { get { return "smb"; } }
		public static PSObject GetPSObject(string computer)
		{
            PSObject DCSMBObj = new PSObject();
            if (computer == "")
            {
                DCSMBObj.Members.Add(new PSNoteProperty("SMB Port Open", null));
                DCSMBObj.Members.Add(new PSNoteProperty("SMB1(NT LM 0.12)", null));
                DCSMBObj.Members.Add(new PSNoteProperty("SMB2(0x0202)", null));
                DCSMBObj.Members.Add(new PSNoteProperty("SMB2(0x0210)", null));
                DCSMBObj.Members.Add(new PSNoteProperty("SMB3(0x0300)", null));
                DCSMBObj.Members.Add(new PSNoteProperty("SMB3(0x0302)", null));
                DCSMBObj.Members.Add(new PSNoteProperty("SMB3(0x0311)", null));
                DCSMBObj.Members.Add(new PSNoteProperty("SMB Signing", null));
                return DCSMBObj;
            }
            bool isPortOpened = true;
			bool SMBv1 = false;
			bool SMBv2_0x0202 = false;
			bool SMBv2_0x0210 = false;
			bool SMBv3_0x0300 = false;
			bool SMBv3_0x0302 = false;
			bool SMBv3_0x0311 = false;
            bool SMBSigning = false;
			try
			{
				try
				{
					SMBv1 = DoesServerSupportDialect(computer, "NT LM 0.12");
				}
				catch (ApplicationException)
				{
				}
				try
				{
					SMBv2_0x0202 = DoesServerSupportDialectWithSmbV2(computer, 0x0202, false);
					SMBv2_0x0210 = DoesServerSupportDialectWithSmbV2(computer, 0x0210, false);
					SMBv3_0x0300 = DoesServerSupportDialectWithSmbV2(computer, 0x0300, false);
					SMBv3_0x0302 = DoesServerSupportDialectWithSmbV2(computer, 0x0302, false);
					SMBv3_0x0311 = DoesServerSupportDialectWithSmbV2(computer, 0x0311, false);
				}
				catch (ApplicationException)
				{
				}
			}
			catch (Exception)
			{
				isPortOpened = false;
			}
			if (SMBv3_0x0311)
			{
				SMBSigning = DoesServerSupportDialectWithSmbV2(computer, 0x0311, true);
			}
			else if (SMBv3_0x0302)
			{
				SMBSigning = DoesServerSupportDialectWithSmbV2(computer, 0x0302, true);
			}
			else if (SMBv3_0x0300)
			{
				SMBSigning = DoesServerSupportDialectWithSmbV2(computer, 0x0300, true);
			}
			else if (SMBv2_0x0210)
			{
				SMBSigning = DoesServerSupportDialectWithSmbV2(computer, 0x0210, true);
			}
			else if (SMBv2_0x0202)
			{
				SMBSigning = DoesServerSupportDialectWithSmbV2(computer, 0x0202, true);
			}
            DCSMBObj.Members.Add(new PSNoteProperty("SMB Port Open", isPortOpened));
            DCSMBObj.Members.Add(new PSNoteProperty("SMB1(NT LM 0.12)", SMBv1));
            DCSMBObj.Members.Add(new PSNoteProperty("SMB2(0x0202)", SMBv2_0x0202));
            DCSMBObj.Members.Add(new PSNoteProperty("SMB2(0x0210)", SMBv2_0x0210));
            DCSMBObj.Members.Add(new PSNoteProperty("SMB3(0x0300)", SMBv3_0x0300));
            DCSMBObj.Members.Add(new PSNoteProperty("SMB3(0x0302)", SMBv3_0x0302));
            DCSMBObj.Members.Add(new PSNoteProperty("SMB3(0x0311)", SMBv3_0x0311));
            DCSMBObj.Members.Add(new PSNoteProperty("SMB Signing", SMBSigning));
            return DCSMBObj;
		}
	}
}
"@

Function Get-DateDiff
{
<#
.SYNOPSIS
    Get difference between two dates.

.DESCRIPTION
    Returns the difference between two dates.

.PARAMETER Date1
    [DateTime]
    Date

.PARAMETER Date2
    [DateTime]
    Date

.OUTPUTS
    [System.ValueType.TimeSpan]
    Returns the difference between the two dates.
#>
    param (
        [Parameter(Mandatory = $true)]
        [DateTime] $Date1,

        [Parameter(Mandatory = $true)]
        [DateTime] $Date2
    )
    if ($Date2 -gt $Date1)
    {
        $DDiff = $Date2 - $Date1
    }
    Else
    {
        $DDiff = $Date1 - $Date2
    }
    Return $DDiff
}

Function Get-DNtoFQDN
{
<#
.SYNOPSIS
    Gets Domain Distinguished Name (DN) from the Fully Qualified Domain Name (FQDN).

.DESCRIPTION
    Converts Domain Distinguished Name (DN) to Fully Qualified Domain Name (FQDN).

.PARAMETER ADObjectDN
    [string]
    Domain Distinguished Name (DN)

.OUTPUTS
    [String]
    Returns the Fully Qualified Domain Name (FQDN).

.LINK
    https://adsecurity.org/?p=440
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $ADObjectDN
    )
    # Modified version from https://adsecurity.org/?p=440
    [array] $ADObjectDNArray = $ADObjectDN -Split ("DC=")
    $ADObjectDNArray | ForEach-Object {
        [array] $temp = $_ -Split (",")
        [string] $ADObjectDNArrayItemDomainName += $temp[0] + "."
    }
    $ADObjectDNDomainName = $ADObjectDNArrayItemDomainName.Substring(1, $ADObjectDNArrayItemDomainName.Length - 2)
    Return $ADObjectDNDomainName
}

Function Get-ADRExcelComObj
{
<#
.SYNOPSIS
    Creates a ComObject to interact with Microsoft Excel.

.DESCRIPTION
    Creates a ComObject to interact with Microsoft Excel if installed, else warning is raised.

.OUTPUTS
    [System.__ComObject] and [System.MarshalByRefObject]
    Creates global variables $excel and $workbook.
#>
    #Check if Excel is installed.
    Try
    {
        $global:excel = New-Object -ComObject excel.application
    }
    Catch
    {
        Write-Warning "[*] Excel is not installed. Skipping ADRecon-Report.xlsx. Use the -GenExcel parameter to generate the ADRecon-Report.xslx on a host with Microsoft Excel installed."
        Write-Output "Run Get-Help .\ADRecon.ps1 -Examples for additional information."
        return $null
    }
    $excel.visible = $true
    $global:workbook = $excel.Workbooks.Add()
    If ($workbook.Worksheets.Count -eq 3)
    {
        $workbook.WorkSheets.Item(3).Delete()
        $workbook.WorkSheets.Item(2).Delete()
    }
}

Function Get-ADRExcelWorkbook
{
<#
.SYNOPSIS
    Adds a WorkSheet to the Workbook.

.DESCRIPTION
    Adds a WorkSheet to the Workbook using the $workboook global variable and assigns it a name.

.PARAMETER name
    [string]
    Name of the WorkSheet.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $name
    )
    $workbook.Worksheets.Add() | Out-Null
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Name = $name
}

Function Get-ADRExcelImport
{
<#
.SYNOPSIS
    Helper to import CSV to the current WorkSheet.

.DESCRIPTION
    Helper to import CSV to the current WorkSheet. Supports two methods.

.PARAMETER filename
    [string]
    Filename of the CSV file to import.

.PARAMETER method
    [int]
    Method to use for the import.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $filename,

        [Parameter(Mandatory = $true)]
        [int] $method
    )
    If ($method -eq 1)
    {
        $row = 1
        $column = 1
        $worksheet = $workbook.Worksheets.Item(1)
        If (Test-Path $filename)
        {
            $ADTemp = Import-Csv -Path $filename
            $ADTemp | ForEach-Object {
                Foreach ($prop in $_.PSObject.Properties)
                {
                    $worksheet.Cells.Item($row, $column) = $prop.Name
                    $worksheet.Cells.Item($row, $column + 1) = $prop.Value
                    $row++
                }
            }
            Remove-Variable ADTemp
            $listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $worksheet.UsedRange, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, $null)
            $listObject.TableStyle = "TableStyleLight2" # Style Cheat Sheet: https://msdn.microsoft.com/en-au/library/documentformat.openxml.spreadsheet.tablestyle.aspx
            $usedRange = $worksheet.UsedRange
            $usedRange.EntireColumn.AutoFit() | Out-Null
        }
        Else
        {
            $worksheet.Cells.Item($row, $column) = "Error!"
        }
        Remove-Variable filename
    }
    Elseif ($method -eq 2)
    {
        If (Test-Path $filename)
        {
            $worksheet = $workbook.Worksheets.Item(1)
            $TxtConnector = ("TEXT;" + $filename)
            $CellRef = $worksheet.Range("A1")
            #Build, use and remove the text file connector
            $Connector = $worksheet.QueryTables.add($TxtConnector, $CellRef)

            #65001: Unicode (UTF-8)
            $worksheet.QueryTables.item($Connector.name).TextFilePlatform = 65001
            $worksheet.QueryTables.item($Connector.name).TextFileCommaDelimiter = $True
            $worksheet.QueryTables.item($Connector.name).TextFileParseType = 1
            $worksheet.QueryTables.item($Connector.name).Refresh() | Out-Null
            $worksheet.QueryTables.item($Connector.name).delete()
            $listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $worksheet.UsedRange, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, $null)
            $listObject.TableStyle = "TableStyleLight2" # Style Cheat Sheet: https://msdn.microsoft.com/en-au/library/documentformat.openxml.spreadsheet.tablestyle.aspx
            $worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null
        }
        Remove-Variable filename
    }
}

Function Get-ADRExcelChart
{
<#
.SYNOPSIS
    Helper to add charts to the current WorkSheet.

.DESCRIPTION
    Helper to add charts to the current WorkSheet.

.PARAMETER ChartType
    [int]
    Chart Type.

.PARAMETER ChartLayout
    [int]
    Chart Layout.

.PARAMETER ChartTitle
    [string]
    Title of the Chart.

.PARAMETER RangetoCover
    WorkSheet Range to be covered by the Chart.

.PARAMETER chartdata
    Data for the Chart.
#>
    param (
        [Parameter(Mandatory = $true)]
        [int] $ChartType,

        [Parameter(Mandatory = $true)]
        [int] $ChartLayout,

        [Parameter(Mandatory = $true)]
        [string] $ChartTitle,

        [Parameter(Mandatory = $true)]
        $RangetoCover,

        [Parameter(Mandatory = $false)]
        $chartdata
    )
    $worksheet = $workbook.Worksheets.Item(1)
    $xlChart=[Microsoft.Office.Interop.Excel.XLChartType]
    $chart=$worksheet.Shapes.AddChart().Chart
    $chart.chartType= $ChartType
    $chart.ApplyLayout($ChartLayout)
    $xlDirection=[Microsoft.Office.Interop.Excel.XLDirection]
    If ($null -eq $chartdata)
    {
        $start=$worksheet.range("A1")
        #get the last cell
        $Y=$worksheet.Range($start,$start.End($xlDirection::xlDown))
        $start=$worksheet.range("B1")
        #get the last cell
        $X=$worksheet.Range($start,$start.End($xlDirection::xlDown))
        $chartdata=$worksheet.Range("A$($Y.item(1).Row):A$($Y.item($Y.count).Row),B$($X.item(1).Row):B$($X.item($X.count).Row)")
    }
    $chart.SetSourceData($chartdata)
    $chart.seriesCollection(1).Select() | Out-Null
    $chart.SeriesCollection(1).ApplyDataLabels() | out-Null
    #modify the chart title
    $chart.HasTitle = $True
    $chart.ChartTitle.Text = $ChartTitle
    If ($ChartTitle -eq "Status of User Accounts")
    {
        $chart.PlotBy = 1
        $chart.axes(2).axistitle.text = "Count"
    }
    #Reposition the Chart
    $temp = $worksheet.Range($RangetoCover)
    $chartparent = $chart.parent
    # $chartparent.placement = 3
    $chartparent.top = $temp.Top
    $chartparent.left = $temp.Left
    $chartparent.width = $temp.Width
    If ($ChartTitle -ne "Privileged Groups in AD")
    {
        $chartparent.height = $temp.Height
    }
    #$chart.Legend.Delete()
}

Function Get-ADRGenExcel
{
<#
.SYNOPSIS
    Automates the generation of the ADRecon report.

.DESCRIPTION
    Automates the generation of the ADRecon report. If specific files exist, they are imported into the ADRecon report.

.PARAMETER ExcelPath
    [string]
    Path for ADRecon output folder containing the CSV files to generate the ADRecon-Report.xlsx

.OUTPUTS
    Creates the ADRecon-Report.xlsx report in the folder.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $ExcelPath
    )
    $ExcelPath = $((Convert-Path $ExcelPath).TrimEnd("\"))
    $ReportPath = -join($ExcelPath,'\','CSV-Files')
    If (!(Test-Path $ReportPath))
    {
        Write-Output "[ERROR] Could not locate the CSV-Files directory ... Exiting"
        Return $null
    }
    Get-ADRExcelComObj
    If ($excel)
    {
        Write-Output "[*] Generating ADRecon-Report.xlsx"

        $ADFileName = -join($ReportPath,'\','AboutADRecon.csv')
        If (Test-Path $ADFileName)
        {
            $worksheet= $workbook.Worksheets.Item(1)
            $worksheet.Name = "About ADRecon"
            Get-ADRExcelImport $ADFileName 1
            $worksheet.Hyperlinks.Add($worksheet.Cells.Item(3,2) , "https://github.com/sense-of-security/ADRecon", "" , "", "github.com/sense-of-security/ADRecon") | Out-Null
            $usedRange = $worksheet.UsedRange
            $usedRange.EntireColumn.AutoFit() | Out-Null
        }

        $ADFileName = -join($ReportPath,'\','Forest.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook("AD Forest Info")
            Get-ADRExcelImport $ADFileName 1
        }

        $ADFileName = -join($ReportPath,'\','Domain.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook("AD Domain Info")
            Get-ADRExcelImport $ADFileName 1
            $DomainObj = Import-CSV -Path $ADFileName
            $DomainName = -join($DomainObj.Name,"-")
            Remove-Variable DomainObj
        }

        $ADFileName = -join($ReportPath,'\','DefaultPasswordPolicy.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook("Password Policy")
            Get-ADRExcelImport $ADFileName 1
        }

        $ADFileName = -join($ReportPath,'\','DCs.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook("Domain Controllers")
            Get-ADRExcelImport $ADFileName 2
        }

        $ADFileName = -join($ReportPath,'\','GPOs.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook("Domain GPOs")
            Get-ADRExcelImport $ADFileName 2
        }

        $ADFileName = -join($ReportPath,'\','DNSNodes','.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook("DNS Records")
            Get-ADRExcelImport $ADFileName 2
        }

        $ADFileName = -join($ReportPath,'\','DNSZones.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook("DNS Zones")
            Get-ADRExcelImport $ADFileName 2
        }

        $ADFileName = -join($ReportPath,'\','Printers.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook("Printers")
            Get-ADRExcelImport $ADFileName 2
        }

        $ADFileName = -join($ReportPath,'\','BitLockerRecoveryKeys.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook("BitLocker")
            Get-ADRExcelImport $ADFileName 2
        }

        $ADFileName = -join($ReportPath,'\','LAPS.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook("LAPS")
            Get-ADRExcelImport $ADFileName 2
        }

        $ADFileName = -join($ReportPath,'\','ComputerSPNs.csv')
        If (Test-Path $ADFileName)
        {
            $CompObj = Import-CSV -Path $ADFileName
            $ADCompStat = $CompObj | Sort-Object Name,Service -Unique | Select-Object Name,Service
            Remove-Variable CompObj

            $ADFileName = -join($ReportPath,'\','ComputerSPNsStats.csv')
            $ADCompStat | Export-Csv -Path $ADFileName -NoTypeInformation
            Remove-Variable ADCompStat

            Get-ADRExcelWorkbook("Computer SPNs")
            Get-ADRExcelImport $ADFileName 2
        }

        $ADFileName = -join($ReportPath,'\','Computers.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook("Computers")
            Get-ADRExcelImport $ADFileName 2

            $worksheet= $workbook.Worksheets.Item(1)
            If ($worksheet.Cells.Item(1,4).text -eq "IPv4Address")
            {
                [void] $worksheet.Cells.Item(1,4).Addcomment("May not be current.")
            }
        }

        $ADFileName = -join($ReportPath,'\','OUPermissions.csv')
        If (Test-Path $ADFileName)
        {
            $Obj = Import-CSV -Path $ADFileName
            $TempObj = $Obj | Select-Object OrganizationalUnit,ObjectTypeName,ActiveDirectoryRights,IdentityReference,AccessControlType,isInherited
            Remove-Variable Obj

            $ADFileName = -join($ReportPath,'\','OUPermissions1.csv')
            $TempObj | Export-Csv -Path $ADFileName -NoTypeInformation
            Remove-Variable TempObj

            Get-ADRExcelWorkbook("OUPerms")
            Get-ADRExcelImport $ADFileName 2

            $worksheet= $workbook.Worksheets.Item(1)
            $worksheet.Activate();
            $worksheet.Application.ActiveWindow.FreezePanes = $isFreeze
            $worksheet.Cells.Item(1,6).Interior.ColorIndex = 5
            $worksheet.Cells.Item(1,6).font.ColorIndex = 2
            # Set Filter to Explicitly Assigned Permissions Only
            $worksheet.UsedRange.Select() | Out-Null
            $excel.Selection.AutoFilter(6,$true) | Out-Null
            $worksheet.Range("A1").Select() | Out-Null
        }

        $ADFileName = -join($ReportPath,'\','OUs.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook("OUs")
            Get-ADRExcelImport $ADFileName 2
        }

        $ADFileName = -join($ReportPath,'\','UserSPNs.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook("User SPNs")
            Get-ADRExcelImport $ADFileName 2
        }

        $ADFileName = -join($ReportPath,'\','Groups.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook("Groups")
            Get-ADRExcelImport $ADFileName 2
        }

        $ADFileName = -join($ReportPath,'\','GroupMembers.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook("Group Members")
            Get-ADRExcelImport $ADFileName 2
            $filter = "Account Operators","Administrators","Backup Operators","Cert Publishers","Crypto Operators","Dns Admins","Domain Admins","Enterprise Admins","Incoming Forest Trust Builders","Network Operators","Print Operators","Schema Admins","Server Operators","Enterprise Key Admins","Key Admins"
            $xlFilterValues = 7
            $worksheet= $workbook.Worksheets.Item(1)
            $worksheet.Cells.Item(1,1).Interior.ColorIndex = 5
            $worksheet.Cells.Item(1,1).font.ColorIndex = 2
            $worksheet.UsedRange.AutoFilter(1,$filter,$xlFilterValues) | Out-Null
        }

        $ADFileName = -join($ReportPath,'\','Users.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook("Users")
            Get-ADRExcelImport $ADFileName 2

            $worksheet= $workbook.Worksheets.Item(1)
            $worksheet.Activate();
            $worksheet.Application.ActiveWindow.FreezePanes = $isFreeze
            $worksheet.Cells.Item(1,3).Interior.ColorIndex = 5
            $worksheet.Cells.Item(1,3).font.ColorIndex = 2
            # Set Filter to Enabled Accounts only
            $worksheet.UsedRange.Select() | Out-Null
            $excel.Selection.AutoFilter(3,$true) | Out-Null
            $worksheet.Range("A1").Select() | Out-Null
        }

        $ADFileName = -join($ReportPath,'\','Computers.csv')
        If (Test-Path $ADFileName)
        {
            $CompObj = Import-CSV -Path $ADFileName
            $ADCompStat = $CompObj | Select-Object OperatingSystem | Group-Object -Property OperatingSystem | Sort-Object -Property @{Expression="Count";Descending=$true}
            Remove-Variable CompObj

            Get-ADRExcelWorkbook("Computer Stats")
            $worksheet= $workbook.Worksheets.Item(1)

            $row = 1
            $column = 1
            "Operating System","Count" | ForEach-Object {
                $worksheet.Cells.Item($row,$column)=$_
                $column++
            }
            $column = 1
            $row = 2
            $ADCompStat | ForEach-Object {
                $worksheet.Cells.Item($row,$column) = $_.Name
                $column++
                $worksheet.Cells.Item($row,$column) = $_.Count
                $column=1
                $row++
            }
            Remove-Variable ADCompStat
            $listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $worksheet.UsedRange, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, $null)
            $listObject.TableStyle = "TableStyleLight2" # Style Cheat Sheet: https://msdn.microsoft.com/en-au/library/documentformat.openxml.spreadsheet.tablestyle.aspx
            $usedRange = $worksheet.UsedRange
            $usedRange.EntireColumn.AutoFit() | Out-Null

            #Add Pie Chart
            #Get-ADRExcelChart $ChartType $ChartLayout $ChartTitle $RangetoCover $chardata
            Get-ADRExcelChart 51 10 "Operating Systems in AD" "D2:S16" $null
        }

        $ADFileName = -join($ReportPath,'\','ComputerSPNsStats.csv')
        If (Test-Path $ADFileName)
        {
            $CompObj = Import-CSV -Path $ADFileName
            $ADCompStat = $CompObj | Group-Object -Property Service | Sort-Object -Property @{Expression="Count";Descendin=$true}
            Remove-Variable CompObj

            Get-ADRExcelWorkbook("Computer Role Stats")
            $worksheet= $workbook.Worksheets.Item(1)

            $row = 1
            $column = 1
            "Computer Role","Count" | ForEach-Object {
                $worksheet.Cells.Item($row,$column)=$_
                $column++
            }
            $column = 1
            $row = 2
            $ADCompStat | ForEach-Object {
                $worksheet.Cells.Item($row,$column) = $_.Name
                $column++
                $worksheet.Cells.Item($row,$column) = $_.Count
                $column=1
                $row++
            }
            Remove-Variable ADCompStat
            $listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $worksheet.UsedRange, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, $null)
            $listObject.TableStyle = "TableStyleLight2" # Style Cheat Sheet: https://msdn.microsoft.com/en-au/library/documentformat.openxml.spreadsheet.tablestyle.aspx
            $usedRange = $worksheet.UsedRange
            $usedRange.EntireColumn.AutoFit() | Out-Null

            #Add Pie Chart
            #Get-ADRExcelChart $ChartType $ChartLayout $ChartTitle $RangetoCover $chardata
            Get-ADRExcelChart 51 10 "Computer Roles in AD" "D2:U16" $null
        }

        $ADFileName = -join($ReportPath,'\','GroupMembers.csv')
        If (Test-Path $ADFileName)
        {
            $GroupObj = Import-CSV -Path $ADFileName
            $ADGroupStat = $GroupObj | Where-Object {$_.'AccountType' -eq 'user'} | Select-Object 'Group Name' | Group-Object -Property 'Group Name' | Sort-Object -Property @{Expression="Count";Descending=$true}
            Remove-Variable GroupObj

            Get-ADRExcelWorkbook("Privileged User Group Stats")
            $worksheet= $workbook.Worksheets.Item(1)
            $row = 1
            $column = 1
            $worksheet.Cells.Item($row,$column).Interior.ColorIndex = 5
            $worksheet.Cells.Item($row,$column).font.ColorIndex = 2
            "Group Name","User Count (Not-Recursive)" | ForEach-Object {
                $worksheet.Cells.Item($row,$column)=$_
                $column++
            }
            $column = 1
            $row = 2
            $ADGroupStat | ForEach-Object {
                $worksheet.Cells.Item($row,$column) = $_.Name
                $column++
                $worksheet.Cells.Item($row,$column) = $_.Count
                $column=1
                $row++
            }
            Remove-Variable ADGroupStat

            $filter = "Account Operators","Administrators","Backup Operators","Cert Publishers","Crypto Operators","Dns Admins","Domain Admins","Enterprise Admins","Incoming Forest Trust Builders","Network Operators","Print Operators","Schema Admins","Server Operators","Enterprise Key Admins","Key Admins"
            $xlFilterValues = 7
            $worksheet= $workbook.Worksheets.Item(1)
            $worksheet.UsedRange.AutoFilter(1,$filter,$xlFilterValues) | Out-Null

            $listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $worksheet.UsedRange, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, $null)
            $listObject.TableStyle = "TableStyleLight2" # Style Cheat Sheet: https://msdn.microsoft.com/en-au/library/documentformat.openxml.spreadsheet.tablestyle.aspx
            $usedRange = $worksheet.UsedRange
            $usedRange.EntireColumn.AutoFit() | Out-Null

            #Get-ADRExcelChart $ChartType $ChartLayout $ChartTitle $RangetoCover $chardata
            Get-ADRExcelChart 51 10 "Privileged Groups in AD" "D2:P16" $null
        }

        $ADFileName = -join($ReportPath,'\','Users.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook("User Stats")
            $worksheet= $workbook.Worksheets.Item(1)

            $ADTemp = Import-Csv -Path $ADFileName

            $row = 1
            $column = 1
            $worksheet.Cells.Item($row, $column) = "User Accounts in AD"
            $worksheet.Cells.Item($row,$column).Style = "Heading 2"
            $worksheet.Cells.Item($row,$column).HorizontalAlignment = -4108
            $MergeCells = $worksheet.Range("A1:C1")
            $MergeCells.Select() | Out-Null
            $MergeCells.MergeCells = $true
            Remove-Variable MergeCells

            $row++
            $worksheet.Cells.Item($row, $column) = "Type"
            $worksheet.Cells.Item($row, $column).Font.Bold=$True
            $worksheet.Cells.Item($row, $column+1) = "Count"
            $worksheet.Cells.Item($row, $column+1).Font.Bold=$True
            $worksheet.Cells.Item($row,$column+2) = 'Percentage'
            $worksheet.Cells.Item($row, $column+2).Font.Bold=$True

            $total = ($ADTemp | Measure-Object | Select-Object -ExpandProperty Count)
            $enabled = ($ADTemp | Where-Object ({$_.Enabled -eq $true}) | Measure-Object | Select-Object -ExpandProperty Count)
            $disabled = ($ADTemp | Where-Object ({$_.Enabled -eq $false}) | Measure-Object | Select-Object -ExpandProperty Count)

            $row++
            $worksheet.Cells.Item($row, $column) = "Enabled"
            $worksheet.Cells.Item($row, $column+1) = $enabled
            $worksheet.Cells.Item($row, $column+2) = "{0:P2}" -f ($enabled/$total)

            $row++
            $worksheet.Cells.Item($row, $column) = "Disabled"
            $worksheet.Cells.Item($row, $column+1) = $disabled
            $worksheet.Cells.Item($row, $column+2) = "{0:P2}" -f ($disabled/$total)

            $row++
            $worksheet.Cells.Item($row, $column) = "Total"
            $worksheet.Cells.Item($row, $column+1) = $total
            If ($total -ne ($enabled + $disabled))
            {
                $worksheet.Cells.Item($row, $column+1).Interior.ColorIndex = 3
                $worksheet.Cells.Item($row, $column+1).font.ColorIndex = 2
                Write-Warning "Enabled + Disabled != Total Users, Try running ADRecon as another user."
            }
            $worksheet.Cells.Item($row, $column+2) = "{0:P2}" -f ($total/$total)

            #Get-ADRExcelChart $ChartType $ChartLayout $ChartTitle $RangetoCover $chardata
            Get-ADRExcelChart 5 3 "User Accounts in AD" "A14:D26" $worksheet.Range("A3:A4,B3:B4")

            $row = 1
            $column = 6
            $worksheet.Cells.Item($row, $column) = "Status of User Accounts"
            $worksheet.Cells.Item($row,$column).Style = "Heading 2"
            $worksheet.Cells.Item($row,$column).HorizontalAlignment = -4108
            $MergeCells = $worksheet.Range("F1:J1")
            $MergeCells.Select() | Out-Null
            $MergeCells.MergeCells = $true
            Remove-Variable MergeCells

            $row++
            $temp = @("Category","Enabled Count","Disabled Count","Enabled Percentage","Disabled Percentage")
            $temp | ForEach-Object {
                $worksheet.Cells.Item($row, $column) = $_
                $worksheet.Cells.Item($row, $column).Font.Bold=$True
                $column++
            }

            $column = 6
            $UserProperties = @("Cannot Change Password","Must Change Password at Logon","Password Not Changed after Max Age","Password Never Expires","Password Not Required","Reversibly Encryped Password","Does Not Require Pre Auth","Account Locked Out","Never Logged in",$(($ADTemp | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -like "Dormant*" }).Name))
            ForEach ($property in $UserProperties)
            {
                $row++
                $worksheet.Cells.Item($row, $column) = $property
                $worksheet.Cells.Item($row, $column+1) = ($ADTemp | Where-Object ({$_.$property -eq $true -and $_.Enabled -eq $true}) | Measure-Object | Select-Object -ExpandProperty Count)
                $worksheet.Cells.Item($row, $column+2) = ($ADTemp | Where-Object ({$_.$property -eq $true -and $_.Enabled -eq $false}) | Measure-Object | Select-Object -ExpandProperty Count)
                $worksheet.Cells.Item($row, $column+3) = "{0:P2}" -f (([int] $worksheet.Cells.Item($row,$column+1).text)/$total)
                $worksheet.Cells.Item($row, $column+4) = "{0:P2}" -f (([int] $worksheet.Cells.Item($row,$column+2).text)/$total)
            }

            #Get-ADRExcelChart $ChartType $ChartLayout $ChartTitle $RangetoCover $chardata
            Get-ADRExcelChart 51 5 "Status of User Accounts" "F14:J36" $worksheet.Range("F2:F12,G2:H12")

            Remove-Variable ADTemp
            $usedRange = $worksheet.UsedRange
            $usedRange.EntireColumn.AutoFit() | Out-Null
        }

        # Create Table of Contents

        Get-ADRExcelWorkbook("Table of Contents")
        $worksheet= $workbook.Worksheets.Item(1)

        # Image format and properties
        #$path = ".\c1b-logo.jpg"
        #$base64sos = [convert]::ToBase64String((Get-Content $path -Encoding byte))
        
        #$base64sos = "/9j/4AAQSkZJRgABAgEASABIAAD/7QAsUGhvdG9zaG9wIDMuMAA4QklNA+0AAAAAABAASAAAAAEAAQBIAAAAAQAB/+Fik2h0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8APD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4KPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgNC4yLjItYzA2MyA1My4zNTE3MzUsIDIwMDgvMDcvMjItMTg6MTE6MTIgICAgICAgICI+CiAgIDxyZGY6UkRGIHhtbG5zOnJkZj0iaHR0cDovL3d3dy53My5vcmcvMTk5OS8wMi8yMi1yZGYtc3ludGF4LW5zIyI+CiAgICAgIDxyZGY6RGVzY3JpcHRpb24gcmRmOmFib3V0PSIiCiAgICAgICAgICAgIHhtbG5zOmRjPSJodHRwOi8vcHVybC5vcmcvZGMvZWxlbWVudHMvMS4xLyI+CiAgICAgICAgIDxkYzpmb3JtYXQ+aW1hZ2UvanBlZzwvZGM6Zm9ybWF0PgogICAgICA8L3JkZjpEZXNjcmlwdGlvbj4KICAgICAgPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIKICAgICAgICAgICAgeG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIgogICAgICAgICAgICB4bWxuczp4bXBHSW1nPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvZy9pbWcvIj4KICAgICAgICAgPHhtcDpNZXRhZGF0YURhdGU+MjAxMy0xMC0wM1QxMToyNjoyNSsxMDowMDwveG1wOk1ldGFkYXRhRGF0ZT4KICAgICAgICAgPHhtcDpNb2RpZnlEYXRlPjIwMTMtMTAtMDNUMDE6MjY6MzBaPC94bXA6TW9kaWZ5RGF0ZT4KICAgICAgICAgPHhtcDpDcmVhdGVEYXRlPjIwMTMtMTAtMDNUMTE6MjY6MjUrMTA6MDA8L3htcDpDcmVhdGVEYXRlPgogICAgICAgICA8eG1wOkNyZWF0b3JUb29sPkFkb2JlIElsbHVzdHJhdG9yIENTNDwveG1wOkNyZWF0b3JUb29sPgogICAgICAgICA8eG1wOlRodW1ibmFpbHM+CiAgICAgICAgICAgIDxyZGY6QWx0PgogICAgICAgICAgICAgICA8cmRmOmxpIHJkZjpwYXJzZVR5cGU9IlJlc291cmNlIj4KICAgICAgICAgICAgICAgICAgPHhtcEdJbWc6d2lkdGg+MjU2PC94bXBHSW1nOndpZHRoPgogICAgICAgICAgICAgICAgICA8eG1wR0ltZzpoZWlnaHQ+OTY8L3htcEdJbWc6aGVpZ2h0PgogICAgICAgICAgICAgICAgICA8eG1wR0ltZzpmb3JtYXQ+SlBFRzwveG1wR0ltZzpmb3JtYXQ+CiAgICAgICAgICAgICAgICAgIDx4bXBHSW1nOmltYWdlPi85ai80QUFRU2taSlJnQUJBZ0VCTEFFc0FBRC83UUFzVUdodmRHOXphRzl3SURNdU1BQTRRa2xOQSswQUFBQUFBQkFCTEFBQUFBRUEmI3hBO0FRRXNBQUFBQVFBQi8rNEFEa0ZrYjJKbEFHVEFBQUFBQWYvYkFJUUFCZ1FFQkFVRUJnVUZCZ2tHQlFZSkN3Z0dCZ2dMREFvS0N3b0smI3hBO0RCQU1EQXdNREF3UURBNFBFQThPREJNVEZCUVRFeHdiR3hzY0h4OGZIeDhmSHg4Zkh3RUhCd2NOREEwWUVCQVlHaFVSRlJvZkh4OGYmI3hBO0h4OGZIeDhmSHg4Zkh4OGZIeDhmSHg4Zkh4OGZIeDhmSHg4Zkh4OGZIeDhmSHg4Zkh4OGZIeDhmSHg4Zi84QUFFUWdBWUFFQUF3RVImI3hBO0FBSVJBUU1SQWYvRUFhSUFBQUFIQVFFQkFRRUFBQUFBQUFBQUFBUUZBd0lHQVFBSENBa0tDd0VBQWdJREFRRUJBUUVBQUFBQUFBQUEmI3hBO0FRQUNBd1FGQmdjSUNRb0xFQUFDQVFNREFnUUNCZ2NEQkFJR0FuTUJBZ01SQkFBRklSSXhRVkVHRTJFaWNZRVVNcEdoQnhXeFFpUEImI3hBO1V0SGhNeFppOENSeWd2RWxRelJUa3FLeVkzUENOVVFuazZPek5oZFVaSFREMHVJSUpvTUpDaGdaaEpSRlJxUzBWdE5WS0JyeTQvUEUmI3hBOzFPVDBaWFdGbGFXMXhkWGw5V1oyaHBhbXRzYlc1dlkzUjFkbmQ0ZVhwN2ZIMStmM09FaFlhSGlJbUtpNHlOam8rQ2s1U1ZscGVZbVomI3hBO3FibkoyZW41S2pwS1dtcDZpcHFxdXNyYTZ2b1JBQUlDQVFJREJRVUVCUVlFQ0FNRGJRRUFBaEVEQkNFU01VRUZVUk5oSWdaeGdaRXkmI3hBO29iSHdGTUhSNFNOQ0ZWSmljdkV6SkRSRGdoYVNVeVdpWTdMQ0IzUFNOZUpFZ3hkVWt3Z0pDaGdaSmpaRkdpZGtkRlUzOHFPend5Z3AmI3hBOzArUHpoSlNrdE1UVTVQUmxkWVdWcGJYRjFlWDFSbFptZG9hV3ByYkcxdWIyUjFkbmQ0ZVhwN2ZIMStmM09FaFlhSGlJbUtpNHlOam8mI3hBOytEbEpXV2w1aVptcHVjblo2ZmtxT2twYWFucUttcXE2eXRycSt2L2FBQXdEQVFBQ0VRTVJBRDhBOVU0cTdGVXQ4eUN5ZlJMcUMrWGwmI3hBO1ozU3JhejlQaFM1WVFsOS81ZWZMNk1qT0hFQ085dHdacFlweG5IbkVnL0o4b1hWMzVrOHY2bmM2Y2wvZFdzOWpLOERDR2FTT2hqYmomI3hBO1VjU052aEZNNS9lSjdpK3lZOGVEVVl4TXhqSVNGN2dIbTk5L0pIemJkYTc1WGt0Nys0ZTUxSFRwVEhKTkt4ZVI0cFBpalptSkpQN1MmI3hBOzcrR2JYUlpUS0pCNWg4NzlxT3o0NmZVQ1VBSXdtT1E1V09mNkM5RXpOZWFkaXJzVmRpcnNWZGlyc1ZkaXJzVmRpcnNWZGlyc1ZkaXImI3hBO3NWZGlyc1ZkaXJzVmRpcnNWZGlyc1ZkaXJzVmRpcnNWZGlyc1ZRT3U2WU5WMFRVTk1MY0JmVzB0dnpxUVY5VkNuSUViaWxhN1lxK1kmI3hBOy9QNWZVNGRGODJGQXNtdDJnWFVWVUNpYWhaL3VMcE50dnRybW4xdVBobmZlK2wreU9zOFRUbkdlZU0vWWR4K2xNZnlROHhmb256dkQmI3hBO2F5TlMyMVZEYXVPM3FING9qL3dRNC9Ua05KazRjZzg5bkk5cU5INDJrTWg5V1AxZkRyK3Y0UHBuTjIrV094VjJLdXhWMkt1eFYyS3UmI3hBO3hWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4Vjg0YXRwdk81L01YeWF5VW4wYSsmI3hBOy93QVQ2U1ArWGE3Vld1bFgvSlFTQS9QTVRXNCtLRjl6MEhzenJQQjFjUWZwbjZUK2o3WG5GdGNUVzF4RmNRTVVtaGRaSW5IVU1ocXAmI3hBOytnak5NK3FUZ0pSTVR5TDdFOHQ2MURyZWcyR3JRMENYa0tTbFIreTVIeHIvQUxGcWpPZ3haT09JTDRycmRNY0dhV00vd212MUpsbGomI3hBO2l1SkFCSk5BT3B4VkFycm1rU1QvQUZlRzZqdUp3YU5GQWZXWmEvemlQa1ZIdTJWZU5DNkJ2M2J1UWRMa0E0akVnZWUzeXZtdnM5WDAmI3hBO3ErbnVMZXl2SWJtZTBLaTZqaGtXUm9pOWVJY0tUeEo0blk1YTQ2M1d0YTB2Uk5MdU5WMVM0VzFzTFJPYzg3MW9CMDZDcEpKMkFHNU8mI3hBO0t2QjljLzV5MGdTNmVQUTlCTTFzcG9seGR6ZW16YjlmU1JXNC93REI0cHBHZVYvK2NxYkhVZFRnc05WMENXMitzeUpGRk5hVEM0UEsmI3hBO1JncWd4c2taNm5zeCtXS0tlODRxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhnZjUmI3hBO3IzU2VWUHozOHBlWXBSVFR0Y3RtMHZVbHBWWFF1WXBDM2lGVzRqYi9BR09BaXhSWlJrUWJITU1SOHcvbGQ1eDA3Vkx5RzIwaTd1N0smI3hBO0taMXRwNFltbDV4aHZnYWlBbmRhZHMwVXNFd2FvL0o5YTBmYnVseTQ0bVdTTVpFQ3dUVkhyelpWNUM4OStjL0xlaWp5N0I1YXVyKzYmI3hBOzlaNUxYbWtxY0VlaEtsQkdTUnlxMWFqcmx1SFVUZ09FQ3orT2pxZTF1eXRMcXN2am5OR0VhMzVHNjg3N2syMXp6djU1czRqTjVyOHgmI3hBO2FYNUx0MkZmcXNhcmMzeFhyOEVBTThoMjhDTXlnTTgrWjRRNkxKUHNyVDdRalBQTHpORDlIM0Y1bjVoL1BMeXREelN3dGRRODEzbmEmI3hBOzkxKzRkYlFIL0lzb21veWV6RVpaSFNSL2lKa2ZOd01uYldUbGlqRENQNklGL3dDbTV2UGZNbjV1ZWZkZmhOcGNhazFucHZSTk0wOVImI3hBO2FXeXIvTDZjUEhrUDljbk1xTVFCUTJkVE9jcG01RWsrYjZEL0FPY1JOUE1Ya1hWcjVoUTNXcE5HcDhWaGhqb2YrQ2tiQ3dMWC9PV1YmI3hBOzllUmVXOUVzNDJaYlc1dTVIdUFLMExSUmowd2YrRFkweFFFcC93Q2NiZklYa1BYZkxtb2FocTFuQnFtcXBkR0Y0TGdCeERDRVZrSWomI3hBO08zeGt0OFZPMU94eFNYcC8vS2tQeThpOHg2YnI5aHAvNlB2Tk9tRndzVnN4V0dSa0I0Y296eVVjV293NFU2WW9aNDdwR2pTU01FUkEmI3hBO1dabU5BQU55U1RpcnpQVy8rY2l2eXcwcThhMEY3TnFEeG5qSTlsRVpJd1I0U01VVnZtcEl4VkhlVlB6MC9ManpMZXBZMm1vTmFYMHAmI3hBOzR3Mjk2aGhMazlBcjFhTWs5aHlxZTJLc244MithOUk4cTZKTnJXcnM2V01ESXNqUnFYYXNqQkYrRWU1eFZpeC9QbjhzbDh2cHJiYW0mI3hBO1Zna2tlR08xOU5qY3M4ZEN3OUlWSUZHQjVIYmZyaXFWMkgvT1MvNVgzVnlrRWs5M1pxNXA2ODl1ZlRIejlOcEcvREZYcUZwZDJ0NWEmI3hBO3hYVnBLazl0T29raG1qWU1qb3dxR1Zoc1FjVmVlZVl2K2NndnkwME8rZXhlK2t2N2lJOFp2cU1mcW9yRHFQVUpSRFQvQUNTY1ZWUEwmI3hBO0g1OS9sdjVoMUNIVHJhOWx0YjI1ZFlyYUc3aWFQMUhjOFZWWFhtbFNUUUFzTVZaN2ZYOWxwOW5MZTMwOGRyYVFLWG11SldDSWlqdXomI3hBO05RREZYbU4vL3dBNUwvbGZhWEx3UnozZDRxR25yMjl1ZlRKSGdaR2pKKzdGV1NlVFB6YjhpZWI1dnEya2FnUHI5QzMxRzRVd3pFRGMmI3hBOzhRMnowNzhDY1Zaamlyc1ZZLzUwODkrVy9KdW1McU91M0JoaWtmMDRJa1V2TEs5SzhVVWVBNms3WXFyK1V2TlZoNW8wbjlLV01GekImI3hBO2JtUm9nbDNFWVpLcUFTZUpKMjM2NHFuT0t1eFYyS3RNeXFwWmlBb0ZTVHNBQmlyejN6ZCtmbjVaZVdlY1UycXJxRjZsZjlEMDhDNGYmI3hBO2tQMlM2a1JLZlpuR0t2RmZOMy9PVzNtaTk1d2VXTk9oMG1FN0xkWEZMbTQrWVVnUkw4aXJZcHA1NzVZMXp6UDUxL012eTJtdTZsY2EmI3hBO2xKTHFkclgxNUdaVVQxbE1uQlBzSU9LblpRTVZmZCtLSGtIL0FEazk1cDFiUVBJRnQraWJ5YXh2TDYvamdhYTNkb3BQUkVVcnVBNjAmI3hBO081VlIxeFVQamlhYWFlVnBabmFXVnpWNUhKWm1KN2tuYzRzbVNlVmZ5MDg5K2FtWDlCNk5jWFVKL3dDUG9yNlZ1UDhBbnRKd2oraXQmI3hBO2NWdDdUNVIvNXhDbmJoUDVzMWdSall0WTZjT1RVOERQS0FBZmxHZm5paTN2L2sveWRvUGxIUkl0RjBPRm9MR05ta283dEl6Tys3TXomI3hBO01UdWZ1eFFsZjVvZmw5YWVlZkswdWtTeUNDN2pZVDZmZEVWRWN5Z2djcWI4V0JLdDkvYkZYeWpjNmIrWlg1VStZVnVTczJsM1ZTa2QmI3hBOzFIKzh0YmhBYThhN3h5S2FWNHR1UEFIQ2w3aitWLzhBemtmcDNtQzd0OUc4elFwcHVxVGtSd1hrWnBheXVlaXR5Sk1UTWVtNUI4UmcmI3hBO1FrMy9BRGxINS92TGI2cDVPc0pURWx4RUxyVldRMExvekZZb1NSMnFoWmgzK0hGSVN6OHB2eUg4cWF2NWFnMXp6VmVzWmI5ZlV0YkcmI3hBO0daWWdrVmZoYVE3c1dicUJ0UVlVV2tINTNmbERvZmsrM3ROWjh1WGpUNmRQTDlYdUxXU1JaWGhrS2xrWldXaDROd1BYb2UrK0tVNnUmI3hBOy9PdDk1by81eHMxSmRSbE0yb2FWZVcxbkxPNXE4a1lsamVKMkozSjR0eHIzNDRxeG44ai9BTW9iSHo3YzM5enFsNUpiNmJwcGpWNGImI3hBO2VnbWtlVU1SOGJCbFZRRTMySitXS3NyL0FEcS9JZnkzNVk4cHY1aTh2UFBHTE9TTmJ5Mm1mMVZhT1Z4R3JxU0F5a093OGExd0todnkmI3hBO2U4ejY3TitVZm5yUmJXUjJuMHUwTnhwL0dwZEV1RWtFd1FqcHg5UGtQYzdZVUY1NStWT24rUnRRODNSV25uT2MyK2xTeE9JcFBVOUcmI3hBO1A2eFVjQkxJS2NVSzh0NmplbmJGSmZTMmcva04rWHVsK1k5TDh6YUlaVkZrV2xpZzlYMTRKQ3lGVWNNMVdxcFBJVWFtQkR4Ly9uSlQmI3hBOzh3TlExVHpYTDVXdDVtVFNOSjRDYUZUUVMzUlhrelA0OEF3VlIyTlQzd3BETWZKbi9PT2ZrUmRCdHB2TTk2OXhxODZMSmNSUlhDeFImI3hBO3dsaFgwd0I4Uksxb3hKNjRvdDVkK2JYNWZ4Zmw3NWxzWjlDMUY1ckc2Qm4wK2ZtcG5obGhaZVNsa29EeDVLUTFCMTlzVXZvdlFQelUmI3hBO3RXL0tDRHp6cW81UERiRVhjVWRGTWx6SEo2SEZhOVBVa0FwNFZ3SVkwZk92NThONWJQbk5kSzBoZEVFWDF3YU94bk43OVU0OHVmS28mI3hBO1d2RDR1dGY4bnRpckR2TkhtZnpoNXk4OGVSZFgwcTMwNXJhNWFlNDh0VzF5WmlGbGhWZnJBdk9MRDRvNW9pRUtVclFZVmZSbWpOcXomI3hBO2FWYU5yQ3dwcXBpWDY2dHJ5OUFTMCtQMCtaTGNhOUs0RlJtS3V4VjJLdmhYODJ2ekY4MStZUE5tdDJkMXF0eEpvOEY5Y3cyZGdybEkmI3hBO0ZoamxaWTZ4cFJXUEVENGpVNHBEQTRvcFpaRmlpUnBKSElWRVVFc1NlZ0FIWEZMMERRUHlLOCs2bGJyZmFqQkY1ZTBvMExYMnJ5QzEmI3hBO1hpZjVZMi9lazA2ZkR2NDVHVXhFV1RUWml3enlTNFlSTWo1QzJjZVU5Qi9LN3lIckZucXlhaGUrYXRmczM1d0MyUVd0a2toQlg5dmwmI3hBO0k1RmRxYkh3ekR5YStBK25mN0hvdEw3S2FySUxuV01lZTUrUS9XOVl0UDhBbklueXE3Y2J2VDcyM1Bpb2lrQStmeG9md3lNZTBJOVEmI3hBO1cvSjdHNmdmVE9CK1kvUVhuMy9PU3ZtUFR2TS9sTHl4cStubVI5RFhVSjRMMlNuQ1JKZlRRcW5FZ2lwakRrYjVsNGN3eUN3ODdydEImI3hBO2swdVR3OGc5Vlc5UDhtZmtOK1Z2bCtLRzZ0dE5HcVhKQ3ZIZmFnUmNNYWlxc3FFTEV2V3V5Vnkxd25vNklpSXFJb1ZGQUNxQlFBRFkmI3hBO0FBWXEzaXJzVmFkMFJlVHNGV29GU2FDcE5BTi9FbkZWRFVkTjAvVXJPV3kxQzJqdTdPWWNaYmVaUTZNUGRXcU1WZkdmNTNlUnRPOG0mI3hBO2VlWDAvUzJJMCs2Z2p2YldFc1dhRlpHZERIeU81bzBaSzEzcFRDa0lMODBiN1VkUzFIUk5Udnl6VDN1aTJNaGticS9CREV6Ky9KNDImI3hBO09LaG1ubG4vQUp4cTFUekQ1ZjAvVzdMWDdRVzJvUUpPaUdLUXNoWWZFalVOT1NOVlQ3akZiVFAvQUtGTDh3LzlYKzAvNUZTLzF4VzAmI3hBO1o1ci9BQ3h1dklINUhlWWJLNnZJNzJlOXZyU2IxSWxaVkNMSkdxaWpkNjF4UWp2K2NTUCtPWjVrL3dDTTlyL3hDVEFwWjMvemtMLzUmI3hBO0ovWC9BUG8wL3dDbzJERlhtUDhBemlRQWRROHpBaW9NTnFDRC9yUzRwTElmekEvNXhpMGpWSjU5UjhxM0s2VmR5VmM2ZEtDMXF6bmYmI3hBOzRHV3J4QStGR0hnQU1VUEd2TG5uTHoxK1YzbXFYVDJsa2pXeW40YWxwRHZ5Z2xVR3BvTjFCWlRWWFg5VzJGS1hmbWU4ZHgrWXV0ejEmI3hBO0tRWGQyYmlOMkcvcFhBRXFOUWY1RGc0cUhwY1AvT0tHdVR3cE5ENWhzNUlwVkR4dXNVaERLd3FDRFhvUml0ci9BUG9VdnpEL0FOWCsmI3hBOzAvNUZTLzF4VzJYK2JmeXgxVFJmK2NmTG55ekJMOWZ2dFBZM3NwaEJVU0tMa3pPRlU3L0RFYTA4UmdRbVRmbkY1Q1g4cC9yUzZwYm0mI3hBOzlPbWZWMTByMUYrdGZXUFI5UDB6Rjl1blA5dW5HbStGV0ErU05GdjlJMW44b29MNUdpbm4vU2wxNlRiRlVuQmVPb1BTcUVOOU9LdnAmI3hBO2ZBcnNWZGlyc1ZmT1dqZjg0cEtKVzFEelRxRDZsZTNEbVdTeDA4aUNBT3g1TjZseE55a1pTVCt6SFhJekpISVcyWXhFbmM4SStmNCsmI3hBO3g2Um9YNVZEUm8vVDBOTlA4dEpUaTB0aGIvVzcxZ2V2Szl1dmkvNUo1UVlaWmRSSDNidWZqejZYSC9CTElmNlI0Ui9wWS84QUZJOVAmI3hBO3loOHBTM1AxclZtdTlidSt2cmFoY1BJZC93REpUMDFwN1V5STBVTHMyWEtQdEZxUkhoeDhPS1BkQ0lIMzJ5WFRQTG1nYVVBTk4wNjImI3hBO3M2ZnRReElqSDVzQlU1ZkREQ1BJQjFlZlc1czM5NU9VdmVTK2FQemM4dS9vVHp6ZnhvbkMydlQ5Y3RxQ2c0ekVsZ1BsSnlHYWJVWSsmI3hBO0daRDZsN1Bheng5SkVuNm8ray9EOWxLSGx1MFBtTHlkNW84bkVjN2k1dHYwbnBLOS9ybGo4ZkJQOHFXT3EvTEw5QmtxUmozdW45c2QmI3hBO0h4WTQ1aC9DYVB1UDdmdmUrL2szcnY2Yy9MRHk1Zmx1Y2dzMHQ1bTdtUzFKZ1luM0pqcm0yZlBHWXN5cXBaaUFvRlNUc0FCaW9Gdm0mI3hBO1h6aitZWG1uenByNTAvU3BKbDArV1gwZFAwK0FsRElLMFZwS1VMRnVwcnN2NDVvODJhV1ErWGMrcTltOWo2ZlE0ZVBJQnhnWEtSNmUmI3hBOzc4YnRYSDVOL21UcHNJdkliVG02Q3BXMW5VeXI4Z0NDZjlqWEdXbHlBWFN3OXBORGxQQ1pmNlliZmozcHg1cTByODA5VS9KV3p0SUkmI3hBO2J2VkxpNnZHbXZZMlBLNmp0WVQrNmpDSDk0OVpVNTkyRzNiTmxvK0xnc3ZEZTBKd2ZtaU1JQWlBT1hJbm4rS2ViYWYrYW41MjZIQismI3hBO2pvdFF2a1dJQkJIZDJ5VHlKVHR5dUlwSDlxRTVsdWtSUGx6OHIvek0vTWZ6RCtrZGJTNmh0NW1YNjdyRityUi9BdTNHRkdDbHpUWlEmI3hBO280anZURlh0WDV1ZmtqQjVrOHI2WmIrWGdsdnFlZ1FDMnNJNURSWnJaVkE5Rm43TjhOVlk3VnJYclVCRHd6UXZNbjV5Zmx1MHVuVzkmI3hBO3ZlV05zejhtczdxMk1zQmMvdFJsbEkzcDFSdDhLVTd0UFBQL0FEa0Q1dTFqVDViT0c4bFMwbmpuamdndC9xMW9XUnEvdnBLSXBVMG8mI3hBO1E3MHhROXAvUDJ5MUxVZnlxdm9MVzBrbnZaSkxWamF3SzB6MUV5RmdBZ3EzSHhwZ1Zpbi9BRGl6bzJyNlpwM21GZFNzYml5YVdhMk0mI3hBO2EzTVR4RmdGa3J4NWhhMHhWbTM1OFdWN2ZmbFJybHJaVzhsMWN5ZlZmVGdoUnBKRzQza0xHaXFDVFFBbkZYbGYvT09tbGViTkV0dk4mI3hBOzl5Tkd1VjFENm5DZE90cm1OcmNUVEw2cFZBMDNCZnRFVjN3cExHN0w4MmZ6MThwdEpZYWpEY1ROeUpFZXAycnV3TEdwNHVPREZmRDQmI3hBO2lQREZVcjAveVQrWmY1cGViWDFTL3M1b2x2SGpONXFrMFJndDQ0bEFRZW1HQURjVVhaVnFmSHh4VjZwK2QzNUUzZXRMYmExNVZpRWwmI3hBOy9hVzhkcmRXQlpWYWFLQlFrVG9UUmVhb09KQk80QXB2MUNIbDJoZm1SK2Mza20xR2pJbDFIYlFmREZaMzlveitrQjJRdW9jTDdWcDQmI3hBO1lVc3UvTDN6Uitldm1Mei9BS1ZxMTNiM2x6cGtEbU83amVMNnBaQzNsK0dRaW9qUm5VZkV2VnFnWW9mUjJzYWxIcGVrWDJweUlaSTcmI3hBO0czbHVYaldnWmxoUXVRSzl6eHlFNWNNU2U1dTAyRTVja2NZMk01QWZNMDhjL0xuVmZLdm5EejVjc2ZLV2xXVWNOckpkeHlDM2plY3omI3hBO0NXSkE3UHhWYTBjblphMTc1aDZmVlN5VG83Q25wTzJmWitHaTB3bnhHVXpNRHVIS1I1ZkR2ZXlYR2thVmMzdHRmM05sQlBmV2ZMNnAmI3hBO2RTUkk4c1BNVWIwM1lGazVEcnhPWnp5cUt4VjJLdXhWMkt1eFYyS3V4VjJLdkp2K2NoUEx2MXZRTFRXNGxyTHAwbnBUa2RmUm1vQVQmI3hBOy9xeUFmZm12MStQWVNldzlqOVp3WnBZanltTEh2SDdQdWVKZVY5YmwwUHpEcCtyUjFyWnpMSTZqcXlWcEl2OEFza0pHYTJFakVnam8mI3hBOzk1cjlLTlJnbmpQOFErM3A5cjZDL0tLMWkwZWZ6UjVhaE5iT3kxSDYvcFpIMmZxT3B4aWVIajdLNnlMOUdkQ0NDTEQ0cktKaVNEekQmI3hBO1A3dUFYRnJOYms4Uk5HMGZJZHVRSXIrT0NjYkJDY2MrR1FsM0Y4bDZKcUYvNUs4NlEzVnhiOHJ2U3Azam50MjI1QWhvM0FKSGRXUEUmI3hBOzA5ODBNSkdFcjZoOWkxV0dHdTBwakUrbkpFVWZ0SDdYMDU1Vzg3ZVhQTTlxSnRLdTFlUUNzdHE1Q3p4LzZ5SGY2UnQ3NXVzV2VNK1QmI3hBOzVWcit5OCtsbFdTTzNmMFB4VDNMblh2TXZ6US9OOVBMVXgwalIwUzQxamlEUEpKdkhiaGhVQWdmYWNnMXAwSGZ3ekIxT3I0VHd4NXYmI3hBO1ZkaGV6cDFROFhMY2NYVHZsK3g1dkQ1cC9PN1VZRHFscytwUzJwK05aWWJmOTBSMStGVlRpdytXWVBpWlR2Y25wNWFEc3JHZkRsNFkmI3hBO2w1eTMrOWszNWZmbmpxTDZsRHBQbXJneVNzSWsxRUtJM1NRbWdFeWlpOGE3VkFGTytYNE5hUWFueTczVjlyK3kwQkE1TlAwMzRlZGomI3hBO3lldjYvcnVuYURwRnhxdW95ZW5hMnk4bXA5cGlkbFJSM1pqc00yT1RJSVJzdkY2VFNUMUdRWTREMVNlQjZ4K2MvbjdYOVJOcjVmUnImI3hBO09KeVJCYTJzUW51R1h0eVlxNXIvQUtnR2FxZXJ5U05EYjNQb2VtOW1kSHA0Y1dZOFI2bVJxUDZQdFVaUE9mNTArWDVZNTlRYTlTSnkmI3hBO0FGdkxmbEU1Sm9GNUZPcDlpRGtQR3l3M0pJOS83V3lQWm5aV29CRU9DLzZNdC92ZXgrYTlkOHhhZitXdHpyRGNMRFhJN2FLVjFqQWQmI3hBO1lwWGRReWdTQmhzR3B2WE5qbHlUR0xpNVMyKzk0blFhWEJrMXd4ZlhpTWlOOXJHL2N4UDhsUFBmbXJ6THFlcFFhMWZmVzRyZUJIaFgmI3hBOzBvWTZNWG9UV0pFSjI4Y3AwbWVjNVZJOUhjZTAvWlduMHVPQnhSNFNTYjNKKzhsNkY1MDFDODAzeW5xOS9aU2VsZDJ0ckxMQkpSVzQmI3hBO3Vxa2cwWUZUOUl6TXp5TVlFam04MzJaaGpsMU9PRXhjWlNBTHpqOGx2UDNtenpKcmwvYTYxZmZXNEliWDFZazlLR09qK29xMXJHaUgmI3hBO29jd3RKbm5PZEU3VitwNmIybTdKMDJseFJsaWp3a3lybkk5UE1sUi9PZjhBTUx6ZjVjODBXdGpvMS84QVZiV1N4am5lUDBZSkt5Tk4mI3hBO0twYXNpT2VpRHZoMWVlY0pWRTlHejJhN0gwMnAwOHA1WThVaE1qbkliVkh1STcyTGo4eVB6Yjh6S2lhTXR3VXRZa1c0YXhnREY1QW8mI3hBO0RTU09FMloyQlBGYURzQm1NZFJsbnl2NE8yL2tYczNTNzVlSDFFMXhTNmR3RjlPODJ5M1FOVC9OVWZscnJHclN6M1Urc2V0RkhwTnMmI3hBO2JaWkxoUWs2Sk8zcG1ObWFvTENqQTA0azVrWXBaZkRKcytYNlhUYXZCMmQrZXg0d0lqSFI0enhWSDZTWTczN3ZmYkRMTDgyZnpQWFgmI3hBO0xhd3Y5U2VKL3JNY056YnlXdHRHNHE0VmxZR0lNcHpIT3J5OS93QmdkNWw5bit6emlNNFF2MGtnaVVqMC9yUFJ2enkxTHpoYWFYYXcmI3hBOzZFazdhZGRRWGlhMFliY1RJSWVFWS9lT1VmMGh4Wjk2ajhNemRaS1lIcDViMjh6N0xZTk5QSkk1cTQ0bUhCY3EzczhoWXZldTk0bDUmI3hBO00xTHpocCtxU3plVkVuZlVXZ1pKUmJXNHVYOUV1aGFxRkpLRGtGM3BtdHd5bUQ2T2IzZmFlRFRaTVlHcHJndnJMaDNvOWJIUzNzM2smI3hBO1RWUHpXMVhRL01hNnU5MWFhcEZEQ2RGbHViT08zL2UwbExoVmFKRmZrVlFHb05LNXNNVXNzb3l1NzZiUEQ5cTRPenNPWEQ0WERMR1MmI3hBO2VPcG1XM3AvcGJkZmVrUDVXL216NW8xRHpkRnBYbUs5RnhiM2l0RkNHaGlpS1RqNGwzalJEOFhFclE5emxPbjFjak1DUjJMc2UzdlomI3hBOy9UNDlNY21DTlNqdWR5Ykh4Sjk3M0dTU09LTnBKR0NSb0N6c2RnQUJVazV0Q2FGdkF4aVNhSE40Qm92NW4vbUQ1bDg4dzZicG1wRzMmI3hBOzArOHV6NlVJdDdkakhhaGl4K0pvMllsWWgzUFhOVEhVNUp5b0htZko5RjFQWVdpMHVrT1RKQzV4ano0cGJ5K2ZlOVV2dnpFc2JUVjMmI3hBO3NXdFhhR1BueW5CUElyQ1pCTTZJRklLeC9WNWVYSmdmaFBFSGF1VlBXQVNxdngxKzU1TEYyTk9lTGo0aFpyYjMxUUp2bWVLUFFqZmMmI3hBO2pkbHVacnBuWXE3RlhZcWdOZjBpRFdkRnZkTG4vdTd5RjRpZkFzUGhiL1ltaHl2TERqaVIzdVJwTlFjT1dPUWM0bTN4MWVXczluZHomI3hBOzJrNjhKN2VSb3BWOEhSaXJEN3htZ2ZiTWVRVGlKRGtSZnplOS9rM3FQNlF0OU92eVFiaUcxazBTK3A5cHZxN2ZXYkpqMzRyRTh5MTgmI3hBO2MyK2l5Y1VLN255ejJuMGZnNnNrZlRrOVg2L3RlczVtUFBNUjg3L2xqNWM4MktacnBEYTZrRkNwZncwNTBIUU9wMmNmUGZ3T1kyZlMmI3hBO3h5YjhpN25zdnR6UG85bytxSDgwL283bmd2bTd5RjVxOGlYMFYzNnJHMzUwdE5VdGl5ZkgxNG1ueEkzdDl4T2FyTGhsak8vemZRK3omI3hBO3UxdFAyaEF4cjFkWXkvRzRleS9rNytZVjE1cDB1ZTAxTmcycTZmeDlTVUFEMW9ucnhjZ0FEa0NLTlQyelk2UFVHWW84dzhSN1NkangmI3hBOzBtUVN4LzNjL3NQZCtwODl5YXdsMTVqYldOU2hONGsxMGJxNnRpL0QxQVg1dEh5bzFBZW5UTlZkbXkrangweGhnOExHZUdvOElQZHQmI3hBO1Z2V1Uvd0Nja2xSUWllV3dxS0FGVVhsQUFPZ0ErcjVzQjJoWDhQMi9zZVBQc1ZlNXpmN0Qvanp6RHpyNWt0Zk1ubUdmV0lMQWFjYmsmI3hBO0tab0JKNm9NaWloZmx3aisxdFhiTUhMTVNrU0JWdlY5bWFLV213akVaY2ZEeU5WdDNjeTlCL09UV3IyYnlYNU10Sm1ZUGVXaVhsMkQmI3hBO1VWa1dDSUNvUHZJL1hNblZTOUVCNVBPZXplbWdOVnFKRCtHUmlQZHhIOVFaWitRR2hXZHI1U2ZWZ2ltOTFDYVFOTlQ0aEZFZUN4MTgmI3hBO09TbHN5TkJBQ0psMWRQN1hhdVU5VDRmOE1BTnZNNzI5T2tqamxRcElnZERTcXNBUnNhalk1bkVBODNsQklnMkdKL20zL3dDUzYxdi8mI3hBO0FJeEovd0FuVXpHMXY5MGZoOTRkeDdQZjQ3ajkvd0Nndk0vK2NjUCtPMXJIL01OSC93QW5NdzlCOVo5ejFYdHAvZFkvNngrNTZ4K1kmI3hBO3YvS0NhOS96QXpmOFFPWitwL3V5OGYyTi9qbUwrdVB2ZVEvODQ1LzhwTnFuL01GL3pOVE5mb1A3ejRmcEQyZnRuL2NRL3Ivb0tILzUmI3hBO3lKLzVUV3kvN1pzWC9KK2ZEci9ySHUvVzJleHYrS3kvNFlmOXpGN0QrV2VsV21tK1JORmp0a0MvV0xXSzZtWUNoYVM0UVNNVDQvYXAmI3hBOzhobWZwb0NPTWVlN3hYYm1lV1hWNURMcEl4SHVpYVpQbDdxbnl2NXEvd0RKcjN2L0FHMXYrWnd6bjh2MXk5NSs5OWMwSC9HZEgvaFgmI3hBOzZIMGI1NS81UXJ6Qi93QnMyOC81TVBtOHpmUkwzRjh5N0sveHJGL3d5SCs2RHhML0FKeDIvd0NVMXZmKzJiTC9BTW40TTF1ZytzKzcmI3hBOzlUM2Z0bC9pc2Y4QWhnLzNNbjBSbTJmTlh6RithK2h6ZVYvekFrdTdPc1VWeTY2alpPT2l1VzVNQi9xeXFUVHdwbWoxT1BnbWE5NzYmI3hBO3I3UDZvYXJSaU10ekVjRXZkL1k5Uy9NRHo1QkorVTZhcGFzRW0xMkpMYUpBZDFhVUVUanY5aFZkZm5tYm56M2hCL25mZy9xZVQ3STcmI3hBO0pJN1I4T1hMRWIrWDAvUFlzWi81eDI4dDg1dFI4eFRMdEdCWjJoUDh6VWVVajVEaVBwT1ZhREhaTXU3OGZqM3UxOXN0YlFoZ0hYMUgmI3hBOzdoK2w2VmQrUWRNdWRRbHVtdUpsaG5ZdE5iTHdwUitabGpWeXZOWTVUTTVkYTdsajB6SmxvNG1WMmZ4K2pkNWJIMnRrakFSb1dPUjMmI3hBOzhxTmNyalFvK1RKOHkzVk94VjJLdXhWMkt2UHZNMzVKZVZkZTFTNTFTU2U3dGJ1NmJuS0lXajlNdFFBbml5TWQ2VisxbUhrMFVaRW0mI3hBO3p1OUpvZmFmVWFmR01ZRVpSajMzZjNvcnlCK1dhK1RMMjhsdGRUa3VyUzhqVlpMYVNNS1E2R3FQekRkZ1dGT1BmMnc2ZlRIR1NidHAmI3hBOzdYN2MvT3dpSlFFWlJQTUhwM2N2ZDFadm1XNkY1M2Fmbmo1TmZWcjNUNzUzczB0cG1pZ3ZDclNRektwSTVmQUN5MXAzSDA1Z3gxMEMmI3hBO1NEeWVseWV5MnFHT000VkxpRmtjaVBuelNuODF2eks4azMvbEM5MHF5dTAxRzl1d2doU0pXS29WZFc1czVBQXBUYW0rVjZyVXdsQ2gmI3hBO3VYTTdBN0UxV1BVeHlUandSanp2cnR5cEl2OEFuSEhUcm82bnEycGNDTFZJRXQrWjZHUm5EMEh5VmQvbU1yMEVUeGsrVHNQYlBOSHcmI3hBOzhjUDRydjRjbUJYdHBMNU04OXRGZDJpeng2ZGRjdnE4b3FrMXZ5cXYyZ2FoNHpzY3hESGdsUjNvdlE0c2cxdWp1TXFNNDh4MGwrd3YmI3hBO2RkTjgyL2s1ZldhWElPbDIvTVZhRzRoaGlrVTkxS3N2YjIyelpSeWFjamtQazhCbTdQN1R4eU1mM2g4d1NSOTZ0NWYxejhzUE1HdFgmI3hBO09rNlZaV2R4TmJSQ1gxZnFzYXh1SzBZUmxsQmJqVVYyNzdaTEhMRE9YQ0lqNU5lczB1djArSVpNa3BnU05mVWJIdjM2c2QvNXlEOHUmI3hBO3kzSGwvVGRVdFl2M2VsTzBVeW90QWtNd1VLZHYyVmFNRDZjcjErUFlFY2c3UDJRMWdqbW5qa2Q4Z3NlOFgrdjdFbC9KZjh6ZEUwalMmI3hBOzMwRFc1L3FpTEswdGxkT0NZNlBRdEd4RmVQeFZJSjIzeXZTYWtRSERKenZhYnNQTG15ZU5pSEZ0VWgxMjZ2UTlhL04zeUZwZG9aLzAmI3hBO25IZlNmc1c5bVJMSXgrZzhWLzJSR1pjOVpqQTUyODFwdlozV1paVndHUG5MWWZqM0tuNWpsZFUvTFRWcHJGdldpbXRCY1JPdTRhTlMmI3hBO3N2SWY3QVZ4MVhxeEVqeUxIc1VlRnI0Q2V4RTYrUEw3M2pmNUgrYk5GMER6QmVMcTA0dFliMkFSeDNEL0FHRmRHNVVjOXFqdm12MG0mI3hBO1VRblo3bnQvYWpzL0xxTU1mREhFWXk1ZFhvSDVuZm1wNVMvd3hmYVhwMTJtcFh1b1F0QXEyNTVJaXVLRjNmN093N0RmTXJVNnFCaVkmI3hBO2pjbDV6c1BzSFUvbUk1Sng0SXdONzlhNlV4SC9BSnh6L3dDVW0xVC9BSmd2K1pxWlJvUDd6NGZwRHVmYlArNGgvWC9RVVA4QTg1RS8mI3hBOzhwclpmOXMyTC9rL1BoMS8xajNmcmJQWTMvRlpmOE1QKzVpOXQ4amY4b1Y1Zi83WnRuL3lZVE5saCtpUHVEd25hdjhBaldYL0FJWlAmI3hBOy9kRk84c2NCOHIrYXYvSnIzdjhBMjF2K1p3em44djF5OTUrOTljMEgvR2RIL2hYNkgwaDV6aGxtOG42N0RFcGVXWFQ3cEkwSFVzMEQmI3hBO2dENzgzZWI2SmU0dm1IWnNoSFU0aWVReVIvM1FmUFA1TCtaZE0wRHppWjlTbVczdGJxMWt0ak85ZUtNekpJcFlqb0NZNlZ6VTZYS0kmI3hBO1RzOG4wbjJtMFdUVWFXc1k0cFJrSlY4eCtsOUNhVjUwOHI2dnFSMDNTOVJpdmJ0WVd1SFdFbDFFYXNxRWx3T05hdU5xMXpiUXp3a2EmI3hBO0JmTjlSMlpxTU9QanlRTVkzVy9mdjArREQvejQ4dGZwVHlrTlRoV3QxcEQrcVQzTUVsRmxIMGZDMzBaamE3SGNlTHVkMTdKNjd3dFQmI3hBOzRaK25JSytJNWZwSHhmUHMycjZqYzZYWjZRN2w3U3psbGt0b2hXb2FmanlIdnVtM3pPYW95MnA5SGpwNFJ5U3lnZXFRQVB3djliNnQmI3hBOzhoK1hSNWU4cDZkcFpVQ2VLSVBkVTd6U2ZISnYzb3pVSHRtOTArUGdnQStROXJhejh6cVo1T2hPM3VHd1QvTG5YT3hWMkt1eFYyS3UmI3hBO3hWMkt1eFZxUkZrUmthdkZ3VmFoSU5EdDFHK0FpMGcwYmVSNjEvempyb3R4SzBta2FuTllxYWtRVElMaFI3SzNLTmdQbnl6WHk3UEgmI3hBO1F2WmFiMnl5eEZaWUNmbUR3L3IvQUVLR2wvOEFPT0dueHlxK3A2ekpjUmpkb2JlRVFrKzNObWwvNGpnajJmM2xzeisya3lLeDR3RDMmI3hBO2szOWxENzNxK2phTHBtaTZkRHAybVc2MjFuQUtKR3RlL1VrbXBZbnVUbWZER0lpZzhmcWRUa3p6TThoNHBGSi9PZjVlK1hmTnNDalUmI3hBO1ltanVvaFNDOWhJV1ZSL0xVZ2hscjJJeXJOcDQ1T2ZOenV6ZTJNK2pQb1BwUE9KNVBPWlArY2JJeTVNZm1FckgreXJXZ1lqNWtUTFgmI3hBOzdzeGY1UDhBNlgyZnRlbGo3YW10OFgrei93Q09zdThrZms5b1BsYlVFMU5ibWU4MUNNTXNjamtSeHFISEUwUk91eC9hSnk3RG94QTImI3hBO1RaZE4ycDdSNXRYRHd5SXhnZmlmbitwblUwTU04THd6SXNzTXFsSkkzQVpXVmhRZ2c5UWN5eUFSUmRCR1JpUVFhSWVWNjkvemoxNWUmI3hBO3ZibDU5S3ZwZExEbXBnS0M0alhmOWdGbzJBK2JITUNlZ0JPeHA2M1NlMkdhRWF5UkdUenZoUDNFZllnckgvbkcvVDQ1dzE5cmt0eEMmI3hBO0NLeHd3TEN4SGNjbWViOVdBZG4vQU5MN0hJeSsya3lQUmpBUG5LLzBCNnZwZWphZnBta3dhVGF4L3dDZ3dSK2lrVWhNbFU3aGk5YTEmI3hBO3JtZERHSXg0UnllUHo2bWVYSWNrajZ5YjdubU91LzhBT08raVhsNDgrbGFsSnBzVWpGamJORUxoRnIyUTg0MkErWk9ZVXV6d1RzYUQmI3hBOzFXbDlzY3NJZ1pJQ1pIVytINTdIOUNwcHYvT1BIbHlDMm1XK3Y1N3k2a2pLUlNoUkVrYkVVNWlNRmlTTzFXcGlPengxS00zdGpubEkmI3hBO2NFWXhpRDc3OHIvWW4va0w4cTlQOG5hamNYdHRmUzNUWE1Qb3NrcXFvQTVCcWpqL0FLdVc0Tkw0Y3J1OW5XOXJkdlQxc0JDVVJIaE4mI3hBOzdMUFBuNVRhZDV3MWVIVTdtL210WkliZGJZUnhxckFoWGQ2L0YveGt4ejZYeEpYZE11eWZhQ2VpeEhIR0lsY3IzOXdINkdYNk5wcWEmI3hBO1hwRmpwa2JtU094dDRyWkpHMkxDRkFnSnA0OGN5WVI0WWdkenBkVG1PWExMSWR1T1JQek5vekpOTHpiVXZ5UjBxLzhBTTAydlBxVTYmI3hBO1RUWFgxc3doRUtodWZQalhyVE5mTFEzSW0rWmVvd2UxR1RIZ0dFUWpRanczWmVrOWRqbXdlWGVUZVlmK2NlOUZ2OVFsdXRMMUY5TWomI3hBO21ZdTFzWVJQR3BQVVIvSEVWSHNTYzE4OUFDZGpRZXcwZnRobHh3RWNrQk1qcmRINDdGT2Z5Ny9LU0R5YnFrK3BmcE5yNmVhM2EyNCsmI3hBO2lJVkNzNk9UOXVRMS9kanZsdW4wdmh5dTdjTHRuMmhPdHhqSHdjQUVyNTMwSTdoM3N0OHkzbWoyMmlYdjZYdUk3ZXhsZ2tqbWFRZ1YmI3hBO1JsS3NBRDlvMFBRWmRtbEVSUEVhMmRQb3NlV1dXUGhBbVlrS2ZOdjVRZVcvMDU1NHNsa1RsYTJIK20zSGgrNkk0QS9PUXI5R2FmVFkmI3hBOytLWUQ2ZjdSNjN3TkpLdnFuNlI4ZWYyVytwYzNyNUs3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXEmI3hBOzdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhnbXMva2Q1djFmelpxdDc2dHJiV041ZTNGeEhLN2xtOU9XVm5YNEZVNzBQUWtacVomI3hBOzZTY3BrOUxmUTlON1U2YkRwb1FxVXB4aEVWWFVBRG05VDhoZmwvcFBrN1QzZ3RXTnhlWEZEZDNyaWpPVjZBS0s4VkZUUVpuWU5PTVkmI3hBOzgza3UxdTE4bXRuY3RvamxIdS9heWpNaDFMc1ZmLy9aPC94bXBHSW1nOmltYWdlPgogICAgICAgICAgICAgICA8L3JkZjpsaT4KICAgICAgICAgICAgPC9yZGY6QWx0PgogICAgICAgICA8L3htcDpUaHVtYm5haWxzPgogICAgICA8L3JkZjpEZXNjcmlwdGlvbj4KICAgICAgPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIKICAgICAgICAgICAgeG1sbnM6eG1wTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iCiAgICAgICAgICAgIHhtbG5zOnN0UmVmPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VSZWYjIgogICAgICAgICAgICB4bWxuczpzdEV2dD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlRXZlbnQjIj4KICAgICAgICAgPHhtcE1NOkRlcml2ZWRGcm9tIHJkZjpwYXJzZVR5cGU9IlJlc291cmNlIj4KICAgICAgICAgICAgPHN0UmVmOmluc3RhbmNlSUQ+eG1wLmlpZDo0RDFEMTFCN0NBMkJFMzExQjI4QUY0ODQzMDE2MTY1OTwvc3RSZWY6aW5zdGFuY2VJRD4KICAgICAgICAgICAgPHN0UmVmOmRvY3VtZW50SUQ+eG1wLmRpZDo0RDFEMTFCN0NBMkJFMzExQjI4QUY0ODQzMDE2MTY1OTwvc3RSZWY6ZG9jdW1lbnRJRD4KICAgICAgICAgICAgPHN0UmVmOm9yaWdpbmFsRG9jdW1lbnRJRD51dWlkOkQ1MkU0NzFBRThFMERCMTE4OUQ0RUM1M0VCQ0ZGRUQ3PC9zdFJlZjpvcmlnaW5hbERvY3VtZW50SUQ+CiAgICAgICAgICAgIDxzdFJlZjpyZW5kaXRpb25DbGFzcz5wcm9vZjpwZGY8L3N0UmVmOnJlbmRpdGlvbkNsYXNzPgogICAgICAgICA8L3htcE1NOkRlcml2ZWRGcm9tPgogICAgICAgICA8eG1wTU06SW5zdGFuY2VJRD54bXAuaWlkOjRGMUQxMUI3Q0EyQkUzMTFCMjhBRjQ4NDMwMTYxNjU5PC94bXBNTTpJbnN0YW5jZUlEPgogICAgICAgICA8eG1wTU06RG9jdW1lbnRJRD54bXAuZGlkOjRGMUQxMUI3Q0EyQkUzMTFCMjhBRjQ4NDMwMTYxNjU5PC94bXBNTTpEb2N1bWVudElEPgogICAgICAgICA8eG1wTU06SGlzdG9yeT4KICAgICAgICAgICAgPHJkZjpTZXE+CiAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6YWN0aW9uPmNvbnZlcnRlZDwvc3RFdnQ6YWN0aW9uPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6cGFyYW1ldGVycz5mcm9tIGFwcGxpY2F0aW9uL3Bvc3RzY3JpcHQgdG8gYXBwbGljYXRpb24vdm5kLmFkb2JlLmlsbHVzdHJhdG9yPC9zdEV2dDpwYXJhbWV0ZXJzPgogICAgICAgICAgICAgICA8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaSByZGY6cGFyc2VUeXBlPSJSZXNvdXJjZSI+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDphY3Rpb24+c2F2ZWQ8L3N0RXZ0OmFjdGlvbj4KICAgICAgICAgICAgICAgICAgPHN0RXZ0Omluc3RhbmNlSUQ+eG1wLmlpZDpCMzI2QzE1RjcxMUVFMzExQTNBNUI2MDA1RjMzNDREMzwvc3RFdnQ6aW5zdGFuY2VJRD4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OndoZW4+MjAxMy0wOS0xNlQxMTo0MzoyMysxMDowMDwvc3RFdnQ6d2hlbj4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OnNvZnR3YXJlQWdlbnQ+QWRvYmUgSWxsdXN0cmF0b3IgQ1M0PC9zdEV2dDpzb2Z0d2FyZUFnZW50PgogICAgICAgICAgICAgICAgICA8c3RFdnQ6Y2hhbmdlZD4vPC9zdEV2dDpjaGFuZ2VkPgogICAgICAgICAgICAgICA8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaSByZGY6cGFyc2VUeXBlPSJSZXNvdXJjZSI+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDphY3Rpb24+Y29udmVydGVkPC9zdEV2dDphY3Rpb24+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDpwYXJhbWV0ZXJzPmZyb20gYXBwbGljYXRpb24vcG9zdHNjcmlwdCB0byBhcHBsaWNhdGlvbi92bmQuYWRvYmUuaWxsdXN0cmF0b3I8L3N0RXZ0OnBhcmFtZXRlcnM+CiAgICAgICAgICAgICAgIDwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpIHJkZjpwYXJzZVR5cGU9IlJlc291cmNlIj4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OmFjdGlvbj5zYXZlZDwvc3RFdnQ6YWN0aW9uPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6aW5zdGFuY2VJRD54bXAuaWlkOjM2NDMwOEE4QkMyMEUzMTE5NjgzOUExODdDMjM1OUVGPC9zdEV2dDppbnN0YW5jZUlEPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6d2hlbj4yMDEzLTA5LTE5VDA5OjQ3OjE5KzEwOjAwPC9zdEV2dDp3aGVuPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6c29mdHdhcmVBZ2VudD5BZG9iZSBJbGx1c3RyYXRvciBDUzQ8L3N0RXZ0OnNvZnR3YXJlQWdlbnQ+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDpjaGFuZ2VkPi88L3N0RXZ0OmNoYW5nZWQ+CiAgICAgICAgICAgICAgIDwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpIHJkZjpwYXJzZVR5cGU9IlJlc291cmNlIj4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OmFjdGlvbj5jb252ZXJ0ZWQ8L3N0RXZ0OmFjdGlvbj4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OnBhcmFtZXRlcnM+ZnJvbSBhcHBsaWNhdGlvbi9wb3N0c2NyaXB0IHRvIGFwcGxpY2F0aW9uL3ZuZC5hZG9iZS5pbGx1c3RyYXRvcjwvc3RFdnQ6cGFyYW1ldGVycz4KICAgICAgICAgICAgICAgPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6YWN0aW9uPnNhdmVkPC9zdEV2dDphY3Rpb24+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDppbnN0YW5jZUlEPnhtcC5paWQ6NEMxRDExQjdDQTJCRTMxMUIyOEFGNDg0MzAxNjE2NTk8L3N0RXZ0Omluc3RhbmNlSUQ+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDp3aGVuPjIwMTMtMTAtMDNUMTE6MjU6NDArMTA6MDA8L3N0RXZ0OndoZW4+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDpzb2Z0d2FyZUFnZW50PkFkb2JlIElsbHVzdHJhdG9yIENTNDwvc3RFdnQ6c29mdHdhcmVBZ2VudD4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OmNoYW5nZWQ+Lzwvc3RFdnQ6Y2hhbmdlZD4KICAgICAgICAgICAgICAgPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6YWN0aW9uPnNhdmVkPC9zdEV2dDphY3Rpb24+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDppbnN0YW5jZUlEPnhtcC5paWQ6NEQxRDExQjdDQTJCRTMxMUIyOEFGNDg0MzAxNjE2NTk8L3N0RXZ0Omluc3RhbmNlSUQ+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDp3aGVuPjIwMTMtMTAtMDNUMTE6MjU6NDgrMTA6MDA8L3N0RXZ0OndoZW4+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDpzb2Z0d2FyZUFnZW50PkFkb2JlIElsbHVzdHJhdG9yIENTNDwvc3RFdnQ6c29mdHdhcmVBZ2VudD4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OmNoYW5nZWQ+Lzwvc3RFdnQ6Y2hhbmdlZD4KICAgICAgICAgICAgICAgPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6YWN0aW9uPnNhdmVkPC9zdEV2dDphY3Rpb24+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDppbnN0YW5jZUlEPnhtcC5paWQ6NEYxRDExQjdDQTJCRTMxMUIyOEFGNDg0MzAxNjE2NTk8L3N0RXZ0Omluc3RhbmNlSUQ+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDp3aGVuPjIwMTMtMTAtMDNUMTE6MjY6MjUrMTA6MDA8L3N0RXZ0OndoZW4+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDpzb2Z0d2FyZUFnZW50PkFkb2JlIElsbHVzdHJhdG9yIENTNDwvc3RFdnQ6c29mdHdhcmVBZ2VudD4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OmNoYW5nZWQ+Lzwvc3RFdnQ6Y2hhbmdlZD4KICAgICAgICAgICAgICAgPC9yZGY6bGk+CiAgICAgICAgICAgIDwvcmRmOlNlcT4KICAgICAgICAgPC94bXBNTTpIaXN0b3J5PgogICAgICAgICA8eG1wTU06T3JpZ2luYWxEb2N1bWVudElEPnV1aWQ6RDUyRTQ3MUFFOEUwREIxMTg5RDRFQzUzRUJDRkZFRDc8L3htcE1NOk9yaWdpbmFsRG9jdW1lbnRJRD4KICAgICAgICAgPHhtcE1NOlJlbmRpdGlvbkNsYXNzPnByb29mOnBkZjwveG1wTU06UmVuZGl0aW9uQ2xhc3M+CiAgICAgIDwvcmRmOkRlc2NyaXB0aW9uPgogICAgICA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIgogICAgICAgICAgICB4bWxuczp4bXBUUGc9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC90L3BnLyIKICAgICAgICAgICAgeG1sbnM6c3REaW09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9EaW1lbnNpb25zIyIKICAgICAgICAgICAgeG1sbnM6eG1wRz0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL2cvIj4KICAgICAgICAgPHhtcFRQZzpNYXhQYWdlU2l6ZSByZGY6cGFyc2VUeXBlPSJSZXNvdXJjZSI+CiAgICAgICAgICAgIDxzdERpbTp3PjI5Ni45OTk5NTk8L3N0RGltOnc+CiAgICAgICAgICAgIDxzdERpbTpoPjIwOS45OTk5Mjk8L3N0RGltOmg+CiAgICAgICAgICAgIDxzdERpbTp1bml0Pk1pbGxpbWV0ZXJzPC9zdERpbTp1bml0PgogICAgICAgICA8L3htcFRQZzpNYXhQYWdlU2l6ZT4KICAgICAgICAgPHhtcFRQZzpOUGFnZXM+MTwveG1wVFBnOk5QYWdlcz4KICAgICAgICAgPHhtcFRQZzpIYXNWaXNpYmxlVHJhbnNwYXJlbmN5PkZhbHNlPC94bXBUUGc6SGFzVmlzaWJsZVRyYW5zcGFyZW5jeT4KICAgICAgICAgPHhtcFRQZzpIYXNWaXNpYmxlT3ZlcnByaW50PkZhbHNlPC94bXBUUGc6SGFzVmlzaWJsZU92ZXJwcmludD4KICAgICAgICAgPHhtcFRQZzpQbGF0ZU5hbWVzPgogICAgICAgICAgICA8cmRmOlNlcT4KICAgICAgICAgICAgICAgPHJkZjpsaT5NYWdlbnRhPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+WWVsbG93PC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+QmxhY2s8L3JkZjpsaT4KICAgICAgICAgICAgPC9yZGY6U2VxPgogICAgICAgICA8L3htcFRQZzpQbGF0ZU5hbWVzPgogICAgICAgICA8eG1wVFBnOlN3YXRjaEdyb3Vwcz4KICAgICAgICAgICAgPHJkZjpTZXE+CiAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICA8eG1wRzpncm91cE5hbWU+RGVmYXVsdCBTd2F0Y2ggR3JvdXA8L3htcEc6Z3JvdXBOYW1lPgogICAgICAgICAgICAgICAgICA8eG1wRzpncm91cFR5cGU+MDwveG1wRzpncm91cFR5cGU+CiAgICAgICAgICAgICAgICAgIDx4bXBHOkNvbG9yYW50cz4KICAgICAgICAgICAgICAgICAgICAgPHJkZjpTZXE+CiAgICAgICAgICAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzpzd2F0Y2hOYW1lPldoaXRlPC94bXBHOnN3YXRjaE5hbWU+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOm1vZGU+Q01ZSzwveG1wRzptb2RlPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzp0eXBlPlBST0NFU1M8L3htcEc6dHlwZT4KICAgICAgICAgICAgICAgICAgICAgICAgICAgPHhtcEc6Y3lhbj4wLjAwMDAwMDwveG1wRzpjeWFuPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzptYWdlbnRhPjAuMDAwMDAwPC94bXBHOm1hZ2VudGE+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOnllbGxvdz4wLjAwMDAwMDwveG1wRzp5ZWxsb3c+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOmJsYWNrPjAuMDAwMDAwPC94bXBHOmJsYWNrPgogICAgICAgICAgICAgICAgICAgICAgICA8L3JkZjpsaT4KICAgICAgICAgICAgICAgICAgICAgICAgPHJkZjpsaSByZGY6cGFyc2VUeXBlPSJSZXNvdXJjZSI+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOnN3YXRjaE5hbWU+QmxhY2s8L3htcEc6c3dhdGNoTmFtZT4KICAgICAgICAgICAgICAgICAgICAgICAgICAgPHhtcEc6bW9kZT5DTVlLPC94bXBHOm1vZGU+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOnR5cGU+UFJPQ0VTUzwveG1wRzp0eXBlPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzpjeWFuPjAuMDAwMDAwPC94bXBHOmN5YW4+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOm1hZ2VudGE+MC4wMDAwMDA8L3htcEc6bWFnZW50YT4KICAgICAgICAgICAgICAgICAgICAgICAgICAgPHhtcEc6eWVsbG93PjAuMDAwMDAwPC94bXBHOnllbGxvdz4KICAgICAgICAgICAgICAgICAgICAgICAgICAgPHhtcEc6YmxhY2s+MTAwLjAwMDAwMDwveG1wRzpibGFjaz4KICAgICAgICAgICAgICAgICAgICAgICAgPC9yZGY6bGk+CiAgICAgICAgICAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzpzd2F0Y2hOYW1lPlNtb2tlPC94bXBHOnN3YXRjaE5hbWU+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOm1vZGU+Q01ZSzwveG1wRzptb2RlPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzp0eXBlPlBST0NFU1M8L3htcEc6dHlwZT4KICAgICAgICAgICAgICAgICAgICAgICAgICAgPHhtcEc6Y3lhbj4wLjAwMDAwMDwveG1wRzpjeWFuPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzptYWdlbnRhPjAuMDAwMDAwPC94bXBHOm1hZ2VudGE+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOnllbGxvdz4wLjAwMDAwMDwveG1wRzp5ZWxsb3c+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOmJsYWNrPjMwLjAwMDAwMTwveG1wRzpibGFjaz4KICAgICAgICAgICAgICAgICAgICAgICAgPC9yZGY6bGk+CiAgICAgICAgICAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzpzd2F0Y2hOYW1lPlJlZDwveG1wRzpzd2F0Y2hOYW1lPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzptb2RlPkNNWUs8L3htcEc6bW9kZT4KICAgICAgICAgICAgICAgICAgICAgICAgICAgPHhtcEc6dHlwZT5QUk9DRVNTPC94bXBHOnR5cGU+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOmN5YW4+MC4wMDAwMDA8L3htcEc6Y3lhbj4KICAgICAgICAgICAgICAgICAgICAgICAgICAgPHhtcEc6bWFnZW50YT4xMDAuMDAwMDAwPC94bXBHOm1hZ2VudGE+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOnllbGxvdz4xMDAuMDAwMDAwPC94bXBHOnllbGxvdz4KICAgICAgICAgICAgICAgICAgICAgICAgICAgPHhtcEc6YmxhY2s+MC4wMDAwMDA8L3htcEc6YmxhY2s+CiAgICAgICAgICAgICAgICAgICAgICAgIDwvcmRmOmxpPgogICAgICAgICAgICAgICAgICAgICA8L3JkZjpTZXE+CiAgICAgICAgICAgICAgICAgIDwveG1wRzpDb2xvcmFudHM+CiAgICAgICAgICAgICAgIDwvcmRmOmxpPgogICAgICAgICAgICA8L3JkZjpTZXE+CiAgICAgICAgIDwveG1wVFBnOlN3YXRjaEdyb3Vwcz4KICAgICAgPC9yZGY6RGVzY3JpcHRpb24+CiAgIDwvcmRmOlJERj4KPC94OnhtcG1ldGE+CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAKPD94cGFja2V0IGVuZD0idyI/Pv/uAA5BZG9iZQBkwAAAAAH/2wCEAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQECAgICAgICAgICAgMDAwMDAwMDAwMBAQEBAQEBAgEBAgICAQICAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDA//AABEIANkCTgMBEQACEQEDEQH/xAGiAAAABgIDAQAAAAAAAAAAAAAHCAYFBAkDCgIBAAsBAAAGAwEBAQAAAAAAAAAAAAYFBAMHAggBCQAKCxAAAgEDBAEDAwIDAwMCBgl1AQIDBBEFEgYhBxMiAAgxFEEyIxUJUUIWYSQzF1JxgRhikSVDobHwJjRyChnB0TUn4VM2gvGSokRUc0VGN0djKFVWVxqywtLi8mSDdJOEZaOzw9PjKThm83UqOTpISUpYWVpnaGlqdnd4eXqFhoeIiYqUlZaXmJmapKWmp6ipqrS1tre4ubrExcbHyMnK1NXW19jZ2uTl5ufo6er09fb3+Pn6EQACAQMCBAQDBQQEBAYGBW0BAgMRBCESBTEGACITQVEHMmEUcQhCgSORFVKhYhYzCbEkwdFDcvAX4YI0JZJTGGNE8aKyJjUZVDZFZCcKc4OTRnTC0uLyVWV1VjeEhaOzw9Pj8ykalKS0xNTk9JWltcXV5fUoR1dmOHaGlqa2xtbm9md3h5ent8fX5/dIWGh4iJiouMjY6Pg5SVlpeYmZqbnJ2en5KjpKWmp6ipqqusra6vr/2gAMAwEAAhEDEQA/AN/j37r3Xvfuvde9+691737r3Xvfuvde9+691WV/Nw2L2Vuf4W7+3p09vPfOxuxOkqzH9vYzKbA3RndpZeswG2Uqqfe9DVZLb9ZRVsmOo9n5Ksyfi1Waqx0J4tcAT3GtL6flea622WWG9tSJgY3ZCVWocEqQaBCzU9VHWYn3FuZeUNn+8LtfL3Pu37buXKnMUb7XJHe20N1Es1wVazdY50dBI11HFb6qVEc8gzWnWoPtD+Z78/8AY/h/gvyr7XrfBo0f3vy9J2Fq8fi0+b+/1BuX7i/hGryatV2vfU18brbn3nG0p4W4XBp/GRJ/1cDV/PrvFv33Nvuu8yav3hyTsceqtfpYmseNeH0T2+niaaaUxT4RQz+0f57f8wjbXi/jO8euOwPHbX/e7rDb9F57fXy/3C/uRbV+dGj/AAt7Prb3b5yg/tZYJv8ATxKP+rejqG9+/u0fusbvq/d9hu+1V4fS7jO9Ps+t+s/nXqxL4e/z6uze3O+OqOpO8+reqdvbe7H3bjtmVe9dlVO6sIcJlNwCTG7eqXxm5M9uanahn3HPSQ1DSVSCGnkeTUSliNOWvd6/3Ld7fbt2t7dIZ5AhdC66S2FNGZhTVQHOASesUPfv+7I5O5E9s98569tt63u73XaLF7tbS7W2m8aOCkk6+Jbw27axbrK6BYzqdVWndUbPnueuuNnXvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3RetwdgY6fv9fjrvamx1bgO2+gtzb12XiqmkaePPwddbuxuzu9MbmJXT7ZqZcV3JsoUtMWLzxzVrBdMTn3SSOOaNoZQGidSCDwIIoQfkR0rsL682y+h3Lb5Hhv7eVJYpENGSSNg6OpGQysAwPkQD18+T5TdH5T42/InuHo7KioZ+ut8ZjC4yqqk8U+U2xLIuT2fnHju3j/j21K+irQLmyzgXP194QcwbTJse9XO0yVrBKVBPmvFG/2yFW/Pr65PZX3Hs/d32o2D3IstIXdttilkVTVY7gDw7qEHz8G5SWKvqnAdAF7J+pQ6k0dZV46spchQVM1HXUNTBWUVXTSNDUUtXTSrPT1NPMhV4poJkDIwIKsAR7srMjB0JDg1BHEEcD0zcW8F3A9rcosltKhR1YAqysCGVgcEEEgg4IPX0evhr33S/J34v9Ld4RSxSZDe+ycfLueOFY44qTe+Gabb2+aKKKJY0jp6Td2JrUh9CaoQjBVBA95t8s7uu/bDa7qPjliGr5SL2yD8nDU+XXyT/eA9sZvZv3l5h9uJFYWu27i4tyaktZygT2bkmpLNayxFsmjlhUkV6M17Peoe697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuqe/wCbr2TL8XqD4R/OcVEeO2z8ZPmf1ptvvHLVLMtBRfGn5TY/NfGztGoyKRFJZ6fAbo39tfcEKFxGa7B07MrFVt7r3VVv/CiL48rt/s3qD5M4Wh8dB2Hg6jrPe08Mb+Jd2bPD5Pa9fWStdDW53a1dPSoqkDw4McXuTjp70bN4N/bb7EOyZDE/+nTKE/NkJH2R9dzf7qT3VO6cnb97PbhJW62q5XcLMEiv011SO5RRx0Q3KJISfxXnGlANbj3CHXXLr3v3XutrH/hOr8hVr9td1/F7M14NXgK6k7j2JSSMGmkw+X+x2vv2nhLMDHSYvLQYaZEUEGXJTNwb6shPZXedcF1sEp7kImjH9E0SQfYDoP2seuJP96/7VG13jl33m2+L9C6jbarxgMCWLXc2TH1aSI3aEmnbbxrny2bPc7dcduve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917qvL+bR8fn+Un8tD5wdGUtA2Uze8fjj2TWbQxyRmV6zsDZeDm3713TqisrapN9bYxwBFyp5ANrH3Xuqkfi52Gn83v/AITtdcbmkds53T1p1VBt3OeQxz5lu8/izG2BrquclqmNcz2rs/ELWjSV/Z3MOIiSiA/nzZf37ytdWiCtwieLH664+6g+bLqT/bdZQ/c291T7P/eI5e5kuZPD2O7uf3fe1NF+mvSIS70I7YJTFcnjmEYPDrVd94a9fU91737r3R2f5dfyGPxh+Y3SXaVZkP4ftUbpg2hv+SSZoaMbE3up21uKsr1UgT0+34MguVRG9P3FBGfwPYp5K3r9w8zWu4MdNv4miT08N+1if9LXX9qjrHX71/tUPeT2B5j5Lt4vF3v6I3VkAKv9bZ/4xAiejTlDbEjOiZx59fRE95o9fKV1737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3XvfuvdcXRJEaORVdHVkdHUMjowKsrKwIZWBsQeCPfuvdaOn/CZzs9vhx/Mu/mtfygdy1/8M27ge5+x+1ehsPWNFR0ss3WW9J9gblbERyLBNWZDffUtftPKwxKl1x235ZQoVXt7r3RS/5i3x3/ANlf+Y3dfV1DQ/Y7UO55d49fxxw+GkGxd7qNyYChofURLT7ejr3xTPxqnoJOB9PeF/Ouy/uDma629BS38TXH6eG/coH+lro+1T19Wf3Tvdb/AF5fYHl3nO5l8Xe/oxa3xJq31ln/AIvO7+jTlBcgeSTLk9Ej9hXrIzr3v3XuvoO/yv8A5DN8lfhN0vvfIV3327duYP8A0Zb8keRJKpt19frFgpa+vMdkWt3HhYqLLOAFAFeLADj3mTyFvP785WtbtzW5RPCk9dcfbU/Nl0v/ALbr5Yvvl+1Q9oPvE8w8uWsfh7Fd3P7wsgAQv019WYIlc6IJjNbA5/sDk8erAPYx6xd697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3XuvnAfzrt/Vn8q3/hT38f/nhQQ1VBs7sjC9E9y77OLhYQ5DY1Xh8v8Xe9cLRxRwCnfNZDrvZVXVPHaRjV5CKoe7yg+/de62Df+FDfQ1FuTZPRXyz2mlFkoMZK/Vm7czjGirosjtjcsVVu3rjLx1tIrwy4aiyMeViE/kaN5MtAE/VdoJ96tl129rv8Q7kJhkPnparRn7A2sfa467D/AN1B7q/Sb1zD7M7hKRDdxLulkpNFE0Oi3vVUE5eWJrVwFFdFtIxqBjVU9499duOve/de62RP+E73yGO3u0O3PjNma3RjexMHB2XsuCV28ce7dnomN3NR0iB9P3Wc2tWQ1EhKn9rC/UWsZv8AZfevBv7nYpT2TJ4qf6dMMB82Qg/YnXI7+9a9qhuvJuxe8O3x1u9quTt92wGTbXRMlu7GldMNyroufiu+B4jbZ95F9cMeve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de6r970/mr/y5PjZlq3b3cnzJ6L2xubGTmlyu08ZvGm3vvDD1KzJTtT5naWwY90bkxE6ySC6VNLEwUMxGlWI3Q9e6KfRf8KLv5M9fklxUHzQx0dS8k8Ylrej/AJLY3GhqdJZJC2YyPTNLiEjZYjoYzhZWKqhZmUH1D17o0nU/82n+Wh3bWUWM66+b/wAdMhmMk0ceNwW4exsNsHcGSmlRJEpcft/sCXa+ZrqwpJcwRQPKNLXUaGt6h691YTBPBVQQ1VLNFU01TFHPT1EEiTQTwTIJIZoZoy0csUsbBlZSQwNxx7117rL7917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de60UP8Ahbz8fkyfTnwe+U9DQLHLsrsnsfoPc+SijBkrafsvbNB2DsmlrJCxZYcPN1VnmgCqF1ZCXUblB7917qzn+T/2jSfzbP8AhPztzqjcOUp812t1519nvi5uaevqhLNjuzujIsXkukM5lKuQw1E0ldtFNoZOtmcqZpZqhGd/U7EXM20Lv2w3W0mmuWI6a+Ug7oz+ThSfl1MHsD7nT+zfvJy97kRFhbbbuMbXAXLPZy1gvIwM1Z7WSVVwaMVIFQOtWGso6vHVlVj6+mmo66hqZ6OtpKmNoailq6aVoKimqIXCvContosohV1IBVgQfeEjKyMUcEODQg8QRxHX1s29xBdwJdWzrJbSoHRlIKsrAFWUjBBBBBGCD1G916e6Hz4t93ZL43/ACI6e7wxgmkbrnfWGzmTpKdtE2T208xx+7sLHIQ3jOc2tW1lHqsdInvY29nHL+6vse9W27R1/QlViB5rwdf9shYfn1GHvT7c2nu57Ub/AO3F5pA3bbZYY2bIjuAPEtZiPPwblIpaeeinX0j8NmMXuHEYrP4Sup8nhc5jaHMYjJUj+SlyGLydLFW0FdTSWGunq6SdJEP5Vgfeb8Usc0azRENE6hgRwIIqCPtHXyNbhYXm1X8+17jG0O4W0zxSxsKMkkbFHRh5MrAqR6jpy936Sde9+691737r3Xvfuvde9+690GfYvdHT/UNF/Ee1u0+u+tqIxtJHUb63nt3aqTqiliKYZvI0TVUjAWVIwzseACTb2hvd023bU17hcQwJ6yOqfs1EV/LoYcp+3vPvPlx9JyTsu7bvcVAK2dpPckV/i8GN9I9S1ABkkDotOE/mCdD9h18mJ6Dxna/yWycFUaKpbpPrLcOT2rR1IYp4q/tLeabI6gx1yOGqNwRLaxvYg+yKLnHaL1/D2dbi+kBofAiYoD85X8OEfnIOpf3H7rXuZyrbC+9z5tj5Ps2TWv733CCO5ZeNU260N5ukn2JYsfKlejL7a3Zvaqo67O7+2Vg+tNuUlBVZJny2/wCiy+4MZS0ieeeXdFJi8Idn4eKnpkkkmlps/kIYlW5ci5U/tZr2buuIVhQ8Br1P/tgq6B/tZG6hnmDbOVdsAg2PdJtyu1ajutm0FsRQ1aGSaYXLitAPFs7cnJIFBWu3Pfztv5adD8lOqfiHsr5J7Z7t+QfcHY2F6z27sfoiCq7UoMPmctVCmqa/d+/NtrP11t3G7eGqXIxSZZsjDFHIUpZDG6qs6C3Vrvv3Xuve/de6C/urubrT48dU787t7i3ZjNj9Z9a7erNz7v3Pl5hFS4/G0mhEiiT/ADtdk8nWSxUlFSQh6itrZ4qeFHlkRG917r5of81n/hQH8o/5gG5Nyde9T5zc/wAdfiYk9VjMT1ttfNS4zefZOISQpHme590YeeOoykmWRfKdvUcy4OiRkhkGQnh+/luBTr3Wv5731rr3v3Xuve/de6v0/kDfLr56bQ+ePxu+M/x47f3BN1j2n2BS4zsXqLetRkd49Uw9Z4inrt2dnbjxm0a3IQwbW3NhdlYWvraSuxM+Mqp6yCGCeaSnkkhfRpTrfX1G/dOvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3VC//AApi+Po+Qf8AJl+XFLR0LVm4+ncTtD5BbakVXkFAep934jM71rnjjR3ZV6qn3BFe6iMyh2OlWB917rU//wCEXXzLPW/y575+E+5Mp4tufJXruHsrr2kqZ5GRO2OlUrarLYvFUt/FHUbo6vzmTrayX6tHtmnX8D37r3R1v5xfx3Px9+cvZb42hNHs/uMQd0bUKRqtOG3lU1g3jRRmJEp4jR79oMmY4FAMNJLT3FmUtiJ7mbL+5ubJygpbXX66eneTrHpiQNjyBHX02/cE91v9dL7tuzreS+Jv/L+rabmp7qWqp9K5qSx1WT24Zzh5VloaqQKtPYA6zS697917rev/AJKnyH/06/B3ZOAyle9Xu/onIVfT+dE8uuofCYSGnyGwKtI3klmWgi2VkqTHRu1g82NmCgBbDLT2t3r97cqRQyGtzaEwt66VzGfs0EL9qnr5qP7w/wBqv9bX7yG47nZRCPYeZYl3SHSKKJpiyXqkgAazdxyzsBkJcR1qTU2yV1dQ4yjqMhkqylx9BRxNPV11dUQ0lHSwILvNUVNQ8cMESD6szAD3Ijukal3IVBxJNAPtPWD1tbXN5OlraRvLdSMFVEUszE8AqqCST5ACvRLu1P5kHwZ6aNTFvf5NdXmvpA33GH2hmn7GzkEoVmFNVYbr2n3PkaKpcLwk8cZAZSbKwJC+4c78p7ZUXd/b6xxVG8RvsKx6iD9oHWQnJX3SPvJe4AR+XOTt6+lf4ZbqIWELD+JZb5reN1HqjNwIFSCOq3e1P+FDfxX2uaik6r6y7Z7WrodfirchBhOu9r1Vm0x+HI5Gsz25U1AFj5MLHpBX6ksFBG4e8/L9vVdvgubhx5nTGh/Mlm/anWXHJX91T7170Fn513jY9ktmpVEM19cL61jjWG3PoNN21TXgKElBo/5zP8xj5SZabbXxE+K2Aptcv289fhtrbv7YyuDlYK8T5LdVTNt3YeDp/G41yZHHLGSy2ZbgMG19zudt/kMHLe3oPmqPMV+1zpjX/bL1PFx/d+fdO9l7Fd499+drp6LqVJbm12yOYcCI7ZRPezNUGiwTlqA1BpUGF2l8Ev5uvydjp8n8tvm5n+jdtZG75HYXWmWp/wC8clMyDyYzLYTqWXY3XK08yEIrNk8qEYF3hZlGs5tuUvcffgJOY91e0gbjHERqp6FYfDj/AONP6kHzivffvL/cS9m3az9jPbm15k3eLEd7uETeAGriSKXcxeX+oGpIFvbVFFWRQTpCvt/Of8J9f5Us9bX/ACk7p2V3P3xikeqr9q74z1T8i+4MlmYopFFJXdN7KpcjtnbT11ZSulNU7ix1HDE6gSVwCl/Yu2r215V2xhNNCbu74l7g+JU/6TEf2VUkevWM3uP9/X7xnP0LbXtu6x8s8s0KpabNGLIIlTQC5Ba8rTDBLhI2yfDFadUv/MP/AIWoZaGjq9i/y6viThdmYOij+wwvZ3ySmhrKuCgEawgYbpDrLMUeCwFRR2ZqWWo3Tk6drp5aIBWjYeIiRqEjAVAKAAUAHyA6w4urq6vrh7y9kkmu5GLO7sXdmPFmZiWYnzJJJ61NfmJ/NT/mD/Patq3+Uvyo7S7F27VTeaPrelzEey+pKIrLJLTmk6p2PT7d2B91SrJ41q5cfLXPGqiSZyL+7dMdH8/4S+dbN2P/ADtvh6ZaZqjFbBXufsnMFYTKaZds9GdipgKklqGupoVTeeQxil5fCAGtHKk5iv7r3X1+Pfuvde9+691o7/8ACvf5j5+iq/jv8Edr5KegweVwcnyN7ahpqhlGfV81m9k9UYSr8Oi9Diq7b+4MhUU0pkSaoagn0I1PGzWUefXutHj3brXWzv8AyW/+E8OY/mH9eUnyg+R2/tz9R/GzIZzJ4jYmC2TRY7/SV2+m3q2XFZ/O4nN56kymB2bs7H56mnx8dZLQZOqrqujqkSCCOOOpl0TTrfWydW/8JU/5U1ViWxsGO+Q2NrGgghGfou40kyySQtEZKpYcjtTIYIz1QjIkBojEA7aEQ6Stanr3VTHy/wD+Egu5cJisruf4OfI9t61NItTUUfU3yFocdg87X00QeWOlxna+zaCl25W5uoQCKKKt29iaJpPVJVwoTp3q9evdCT/wl7/lk91fH/5JfLDvb5R9Obq6s3109hsX0BsTCb5w60s/9495mk3h2DuHb9QfNSZCCg2njcLT0eWx809FXUG4JxBNJGzH34nr3W7R7r17r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de6DPunq3b/AHl052z0puwX2r3D1nvzq3cw8EdVfb/YO1srtLMj7aYrFUXx2Xk9DEK/0PB9+6918SP4qd2b/wD5dvz06e7okpaqi3v8VPkLi6veWBo5mWbI0ux91yYDs3ZLyh6ZmpdzbeiyeIns0ZaCqcXW9x7r3X1Pv5yvxjy3zX+M/RffXxw25Xdq7l23U4/cO24dn445HLbw6f7YwNBlRk8ZDDprsnHSVdHiqyCEJIyU9TUyKFu+qKPdfli733a4LzbYml3C2kI0qKs0bjuoOJ0sqkDNAWPr10g/u3/vAct+z/uHu/LHPu4wbbyXvtijie4k0QRXtoxMWpm7YxNDLOjOSup0hVq9tNUndXxS+UGxjJ/fP45d67VSK5abcHUu/cTTMgadBNFVVuBhppqdzSyaZEZo3CMVJAJ946XHL2/2n+5Vjdx0/ihkA/aVp5HPXc/ZPe72Z5kA/q/zby1es3lBudlK1aA0KpMWDDUtVIBFQCAT0BdXR1mPqZaOvpamiq4CFmpauCWmqYWZVdRLBMqSxlkYEXAuCD7KWVkbS4IYeRwepKguILqFbi1dJIG4MpDKfLBBIOcYPRg+iPlv8jPjHj974vobtTOda0nYyYNd4fwOkwk1VlDts5U4WSCvymKyFfh6ii/jdUPLQy00kiylXZgFAOdo5j3vYUlj2i4eBZ9OvSFqdNdOSCVpqOVIOc9RZ7mexftN7x3W3Xnubsltu8+0mb6XxmmCx/UeH4wKRyIkqv4MfbMsiqVqoBJJzw1vy2+Y+7o8GmS75+SW8JJI5lx0+S3x2bXUSESiOokjqajLR4bHU8SvaR/DTwRK3KoptsPzHzNc+EGvL659KySkfPNdIHrgAfLpuS29i/YHYjuTQ8s8obCARrWOz29HOKqCqxGWRjTtGuR2IwzEVtI+Pv8AIF+WPZTUWU7qz+zvj7tydYpZqOtqIOwt/mKUCSPw7b2xkYtuU+qL9Yqs5T1ELMA0BYOqj/ZvZ7mK+pJujxWcB8ifEk/3lTp/bICPTj1hd7p/3oXsfygJLL28tb/mndlJAdFNjZVGDW4uIzO2eBjs3RwCRIAVJsG3P8Qf5Jv8rTAU29vmx3X19X7jgpIsjSU/yF3vQZHLZgxkJUvsj497Np/41vWjdnBanOH3DNCovrtqPuVtm9q+VNqpJcRteXA85jVa/KNaLT5OH+3rnD7p/wB4x94/3F8Sy2W+t+WNkao8PbEKTlTw13spkuA4/itmtgfNOFKoPlf/AMLKfin0vg6jrT+XH8Ucr2OuJgqsbgt6dm0VF0j01h5I1cUOR231ntOOu3vunCNpTVS1bbOqQCwuNI1SJBBBbRCC2RI4V4KoCqPsAoB1g9u+87vv+4Sbtv13c3u6zNWSa4leaVz6vJIzOx+bEnrU++af/CgD+ar86DlcT2Z8ntzdc9cZRnV+oPj35+mevxRSoqy4nKvtisG997YmSRfJ4Nx5rMoslittKhXei3qmV3aRmd2Z3di7u5LM7MSWZmJJZmJuSeSffuvdcffuvde9+691t+/8IuutP70fzNe6exauk82P6t+IG9/sqrweT7PdW9+z+p8FjP35KCogp/uNsQZtfRPT1L2snki86+/de6+n37917r3v3Xuvmcf8Kq1yq/zWsocj9z9o/wAe+nGwXnYtEMUP70JN9mCTopv44tZccfvaz+bm68Ovda2/vfWuvrv/AMmzefXe+f5WnwRynWVTQ1GCxPxr6x2ZmloWhK0vYmxdu0e0O0qaqSGefxVy9jYbKNKGKuzPrKrq0ih49b6sw96691737r3Xvfuvde9+691737r3QAfI75U/HT4i7Aqez/kr3Fsfp3ZUHnSnye8MvHS1ubq6aNZpsXtTb1KtXuTeWcELaxQYmjra1k9SxEAn37r3Wtj3j/wrt+E2yctVYno7ofvTvOOknMP95MxJtvqLa2SjWZkNVh3ys2693SQNCodBW4aglJOlkXk+7aevdFSxv/Cy/Gy1sMeX/l111Djm8n3FVjflhT5WtitFIYvDQVXxvw0E+ucKraqmPShLDUQFb2nr3Vhnxw/4VZfy4O4cpR7e7exHcfxiy1ZPDTrm997XpN69d+WqdYoIzufriuz+4qO0xtNNW4Kjo4EKu84QSGPWk9e62Oeu+yevO3dnYPsPqrfO0eydhblpRXbe3nsXcWJ3VtfNUhJXz4zOYSrrcbWIrgq3jkYo4KmxBHvXXulr7917qtf+ZP8AzS/j9/K12R1pv75A7P7j3hh+091ZbaG36bp3b+ydwZKjyWGxC5qqnzMO9ewuvqWnoZKVwsbQTVEhk4KKPV72BXr3VZPTv/Cqz+W/3H2v1v1LSdf/ACv2FW9l732zsSg3l2LsXpvE7D23kN1ZekwmPyu78tg++9x5TE7dpa2tQ1dVFQ1ApodUjroViPaT17q6j5efOr4pfBHYUPYnyl7j211hh8g1TDtzEVX3mZ3pvKtpEV5qHZuyMDTZLdO45YWljWeWnpWpaPyo1TLDG2v3rj17rWl7d/4WG/Gnb+UraTo/4h9zdoY2m8sdNl+xd97O6dXITxPoEtPQ4PF9xVUOOqLF45JvHUaCuuCNiyLbT17oMto/8LJ9kVuTSLfnwB3VtvDGSIS1+0fkfiN7ZNIj5PM6YnM9LbApZJI7JpU1qh9Ruy6Rq9p691er8E/56v8ALy+fubxWw+tezsl1n3FmXjgxXTPeWNodib1z1ZJqCUO0MhTZfObH3rkpjG7R0OLy9VkzEpkemRb20QR17q4j3rr3XvfuvdFp+UnzF+Mnwq6+PaHyh7k2h1BtCWWelxU24ampqc7ufIUsK1FRitm7Qw1Lk927zy8FO4kkpcXQ1c8cR8jqqAsPde61qO6f+FfvxF2pk67HdF/GbvHuSCikqIYs7vHO7S6dweVeLyCGpxYiXsrcIxtSwSz1mOo6lVYloAVCtbT17ovGI/4WW4aatjjz38u/J43HFW8tViPlTS5utRhbSI8fW/Hfb8Einm5NStv6H37T17qwz48/8Ksf5a/bmSosF2vju6PjRk6uVIP4zv7Z1NvLYSyzMscCf3h6yyW6dw06tK1pJqvCUlLCvreUIGZdaT17rYY6g7s6f+QGycb2R0f2dsXtrYeWVfsd19fbnxG6sK8xhineinrMPV1UdFk6aOdRPST+Opp3OiWNGBA117oUPfuvdRqyso8dR1WQyFVTUNBQ009ZXV1ZPFS0dHR0sTT1NVVVM7JDT01PCjO7uwVFBJIA9+691QN8t/8AhSv/ACyPi7lcntLbm/N0/J7fGLlqKOrxXx4xGN3JtOhr4SVVKzs7cGZ23sPI0bOCGmwlZmmjtYpfj3uh691UtuH/AIWV7apslLFtP+XtnM1hxr8NduH5P0G2MlJaonWPy4nG9B7upYddKsbtatfTI7INQQSPvT17oZOp/wDhYX8Xs/X0FL3V8Su7+sqSpanirMn1/vHZXbsGNkmiiEs8tPmYOpK2px9LVO2t4o3qDAmtIGkIh9+09e62HPhv/Ms+Enz3xb1Xxh762nvnP0mPGSzXXVeazaXaO36ZXENTPlevt0U2K3McfR1J8T19NBU413I8dRIrIzaoR17o9fvXXuve/de697917r3v3Xuve/de697917oEPkZ8j+lfiZ0/u/vj5A79xHXPV+yKNarN7hy3nmeWedxBjsNhcVQw1WV3BuLMVbLBRY+ihnq6qZgsaMb2917rVS3/AP8ACxL48YndWTx/Wnw47d3xtCmlaLG7m3Z2RtLr3L5RUkkQ1D7Wx+39/wANBTSqqvHqyLylW9ccbAj3bT17rY0/l0/OfbP8xf4wbb+UWzOr+weqdrbn3FunbuIw3YbbfmrMwdoZL+CZbPbertu5XJU+S21/H4KvHxzzpR1BrMfUq0CoqSSVOOvdHo9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Qe9o9t9VdIbMynY3c/ZewepOv8IobMb47L3ft/Y20cWGSR0FfuLc+QxmIpGkSJioeYFtJsDb37r3Ws780f8AhXX/ACx/ja+U230J/f75pdgULVFOkfV9A2yOpqevpiganyfbG+aGnkrKWcP+1Wbewm4qR9J/cHF/de6+Zh8te94PlH8n/kB8k4di43rNu++3t/dv1uw8PlqnOYvbOW7D3JkN1ZrHUGVrKSgqKylGXyk7oTBCFD6VRVAUe6919s34M9dv1B8Jvh51NIjRSdX/ABZ+PvXbxOxZ432T1LtHbTIzGeqLMjYyxJlkuf7TfU+690ab37r3TBuDam192Uy0W6tt4DctGl9FJuDD47M0y6pIZW0wZGmqYl1S00bGw5aNT9VFmZre3uF0XCJIvoyhh/MH0HRpte97zscxuNlu7qzuD+KCWSJuBHxRsp4Mw48GI8z0XXc3wY+GO8Cz7h+Knx9rqh7666LqTY+PyTjQ8YV8njMLR5B1VZCVBlIVrMLMAQSz8p8sXOZtvsy3r4MYP7QoP8+pX2f7yX3g9gAXauduaY4hwQ7nePGMg4jkmdAcZOnIwcEjrXN+b3/CpP8Alv8A8uXcHY3xh+LXQu5e7O0enN37n653Ds/rzbeF6C6H2tvnZOXn2vuzb9bvLI4SXM11Tgs3j56bz4ba2TxtaaVjDWmJopnOLWztLGEW9lFHDAOCooVR+SgDqMOYuZ+Zeb9zfeua9wvdz3iT457qeW4mbNcySs7nJPE9ajvzS/4VNfzYPlumV27tDtPC/EbrfILV0o2r8Z8dWbV3XPQzCSKnfJdw5quznZ8GWgppSrzYTI4Kmlc6/tlKpoUdEfWvFufdO597Z/K7r3nuPPbu3Tnapq7N7l3Pl8hn8/mK11VHrMrmcrUVeRyFUyIAZJpHcgAX49+690w+/de697917r3v3Xuve/de65IjSMqIrO7sEREBZnZiAqqoBLMxNgByT7917r6Lv/CMb4W979HYj5p/IDvLpPsjqfGdq4f4/wC1ulsp2PsjObLm3vtvHy9q7h3zmttRbnwePr8rttp63b3hrKOVqOofVcOY42X3Xut5n37r3XvfuvdaOv8Awr9+Iueq6z41fOHbmLqa/B43CVvxy7SrKWkkkXABMxmN99U5GulhMgjx+Vrc9uOjknlWKOGpWkh1u9TGgsvp17rR692611eB/Jx/nWdt/wArTe9dtXMYmu7W+KfYGcpsn2R1THVJFn9uZXwJQzdg9T1ldVU+Nxm7lo4olraGqK4/OU9PHBM9NMlPXUuiK9b6+k58QfnH8XPnZ1tTdofGLtnbvYuFEVINwYOCb+Hb32NkqqESnCb62ZX+HPbZycTalUzw/bVQQyUss8JWVqcOvdGz9+691737r3XvfuvdU9fzh/5uHWf8rHo+lzP2OK7A+RvZUGRoekupKqtMdLU1FInjrN/7+Siq6bLUPXO2qiRFlFO0dVlaxkoqaSHVPV0mwK9e6+YP8qfl18h/mr2zmO6fkp2Znuyd8ZRpYaR8lMIMFtbDtPJPT7Z2XtulEWG2ntqieQmOjooYo2ctLJ5Jnkke/Wui2+/de6fK7bG5cZjaPM5Lb2cx+HyH2/2GWrsTX0mNrvu6d6ul+zrp6eOlqfuaWNpI9DtrjUsLgX9+690x+/de6sl/lsfzRvkt/LL7foN9dQbirs51nmMrSS9tdD5rLVcfXvZ+HUR01U1RR6aqDbu9aShW2Mz9LAa2ikVUkFRRvUUc+iK9e6+qP8NPl70586fjvsD5K9GZafIbJ31QyefF5JaaHceztzY5xTbi2VuyhpaiqhoNx7drwYpkSSSGaMxzwPLTzRSvTh1vrWL/AOFin/ZNnw4/8Tjvv/3goPdl49e60B/dutdDj8g/kr3x8rOwZO0/kR2lu7tnfjYbE7cp89u7JyV0uN29gqf7bFYPD0iiKgxGKpQzymGmiiSWqnmqZQ9RPNLJ7r3QTYfbu4NxSzQ7fwWYzs1Miy1EWHxlbk5YInbQsk0dFBO0SM3ALAAn37r3TP7917rLBPPSzw1VLNLTVNNLHPT1EEjwzwTwuJIpoZYyskUsUihlZSCpFxz7917r6JP/AAmv/nF7t+X21cv8LPk5uqp3L8geptrLuLq7sbO1j1Gf7d6sxUlNjspitz11QPJl+wevZKmmL1rySVmaxU/nmV56GtqqipHW+tjX5dfJPZnw9+M3dnyb3/FLV7X6Z2Dmd41OLpp6emrNw5OmjSk23tXH1FU6U0OT3ZuWspMbStIQgqKtL8e69e6+RH8yPmR3v86+993/ACB+QO76zcu7Ny1ky4jELNNHtjYe2I5nbDbI2RhmdqbB7awdMwREQeWpl11NS81VNNNI5w610WGjo6vI1dLj8fS1NdX11TBR0VFRwS1VXWVdVKsFNS0tNArzVFTUTOqIiKWdiAASffuvdGUn+E3zNpaCTKVPxH+TtPjIqb7yXIz9CdqxUEdJo8v3UlZJtRadKbxnVrLBdPN7e/VHXui2V1BXYusqsdk6Oqx2QoZ5KWtoK6nmpKyjqYWKTU9VS1CRz088TghkdQykWI9+6919Pr/hNN8Pv9le/lpbE3zn8X9j2J8rMzVd9biknh0VsGzspTQ4bqXE+YqrSY2TY2OhzUKEXjnzs4ufdDx631el2Z2XsPprr3efa3aO6cVsnrvrzbmV3bvPdmbmaHGYHb2EpJK3I5CpMaSzy+KCI6IokknnkKxxI8jKp117r5lH84z+e53h/Mb3Xn+qurshuDp/4a4nJTUeF67oMg9Hn+3YsdWS/Zby7eraNKeasTIBI6ml22HfF4wiPyfd1cQqzcCnXuqAPe+tdPWG23uLcbVCbewGazr0qxtVJhsVXZRqZZS4iaoWignMKymNgpa2rSbfT37r3TL7917pXbC39vjqzeW2+xOtd3bj2Fv3Z2Vps5tXeO0czX7f3Lt7MUbaqbI4fM4yemr6CriJIDxupKkqbgkH3Xuvo8/yB/54r/zBNvy/Gb5MZDC4v5ebCwLZLCbgpYYMRjvkBsrD08SZHcVHjY/HQ0HZG3ox5s1j6QJBV0xOQo4Y4Y6uCjoRTrfWzR7117r3v3Xuve/de697917r3v3Xuvntf8K5/lJuveHyx6Y+JNDlayDrfpnqzHdn5jCxyT09HlO0+zMhnKQZKvpwqQZJtvbDwtDHj5mMhpWyleiaPNKGsvXutkj+XJ/Is+B3xy+KvWW3u4fjJ0r353buvZWA3J3Dv/unrrbPZ9e+9c/iabIZnC7Rg3xjc1SbN25tmaqOPokxkNFJPFTLUVGuqklkOiT17q53qzqrrjpDr/bHVPUWy9v9edcbLoZMbtXZm1qCLF4DA0M1XU5Camx1DCBHAk1dWSzP+XkkZiSST7117oQPfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+69035bLYrA4zIZvOZPH4XDYmjqMjlcvlq2mx2MxmPpImnq67IV9ZJDS0VHSwoXklkdURQSSAPfuvdUD/ND/hTd/KW+HaZfC0vezfJ7sjGq8cewPi3R0fZtO1WQ8caV3aD5LD9O0UMFUuirSLP1OQplDEUkjgRt7r3WpJ8z/wDhZZ87e4Wym3Ph91f1v8RdpzSVcFHu7LRUfeHcUlMJBHS1cWR3ZhKDrPByVNOrPLT/AN28jJTvIFjrG8fkf3XutWHv/wCUPyO+Ve8H3/8AJTvLtTvPeB8q0+b7P3vuDeFRjKeVgzUGDhzNdVUmAxa6QEpKKOnpo1ACRqAB7917oCPfuvdCn0Z15L273b071PTpJJP2f2n1915DHCWEry713biNtxpEUs4kZ8kAtub/AE9+69196NESNFjjVURFVERFCoiKAqqqqAFVQLADgD37r3XL37r3XvfuvdcXdI0aSRlREVnd3YKiIoLMzMxAVVAuSeAPfuvdfBe727Bk7b7v7k7Wmmaom7N7V7D7BlncSq88m893ZfcbzMsyRzK0jZIsQ6qwJ5APHv3Xugq9+691737r3Xvfuvde9+691Yj8Nv5Tv8w/59VNE/xc+K/Z2/dq1kxibs/J4yDYnUFKIp0hrDJ2rvup25sWrqqBXMklFSV1TkGRT46eRrKfde621Phd/wAIn8zVHEbo+f8A8qKXFU7LR1dd1F8X6A5DIlZUM8tBk+4exMJDQUFVTHTDURUW166J2MnhrdKpK/uvdbbnwy/k4fy2PgP/AArJfG/4p9b4LfeJ0yU/bu8qGbszuBKv0meroOxt+zZ/ce3PupEDSU+Ilx9FqUaYFCqB7r3Vm/v3Xuve/de697917oHPkF0H1X8oumOxOge7Nr0m8ese0NuVe2t04OpPjkamqCk1Hk8XWqDPidwYLJQQ12OroStRQ11PFPEyyRqR7r3XzBv5rn8j/wCTn8tPdu4N3UuHy/b/AMTavLt/czvbb9A9YdvY+vqmjxW3e48XQQA7J3TTl46f71kGFycjxtSzrNI9FT3Br17qkz3vrXQrdMd6dy/HTfuK7Q6J7O3v1L2BhW/yDdew9w5HbuWEDSRyTY+rmx88KZLEVhiC1NFUrNSVMfoljdCVPuvdbiv8t/8A4ViZ2jrMF1b/ADJNpQZjFzyUOMpvkz1ZgI6LL40swhkyXafV2LRMflqW8nlnr9sx0k0EUemPEVTuXWpX0631uydWdrda939f7Y7V6g3ztjsjrjeeOTLbX3ns/L0ecwGZomd4ZHpa+ikliE9JUxPBUQPpnpqiN4ZUSVHRa9e6UG7N04HY21dzb23VkYMPtjZ+38zuncmXqiVpsXgdv46py2YyNQQCRBQ4+kklcgfpU+/de6+Ph/MY+a29/wCYH8wO3/kzvGfIRUG7M9Niut9s10qumxOpsBNPQ9f7OpoYnajp5cfhdM9e0AWOry1TV1ZHkqHJuBTrXQB/Hnobsj5Qd39X/HzqLDNnuxu294YnZu16Al46WOryU3+VZXK1KJJ9hgcBjo5q/I1TKUpKGmlmf0ofe+vdfUv/AJcH8lj4Z/y8Ov8AbEWE652n2x3/AA4+lm3t8iN/bYx2Z3jlNwsqy5A7Ggy4ysPWe14qj9ulocW0UzwRRmsnq6gNO1Ca9b6txr6ChylFVY7J0VJkcfXQSUtbQV9PDV0VZTTKUmp6qlqEkgqIJUJDI6lWBsR7117rV8/na/yBfj38iekexfkR8RuqNudQ/Kbrrb+W3o+1+s8PQ7X2X3risLTzZbO7ay2ysPSU+Cpux66khnlxWVooaWoyGQcU+RadJYp6SwPXuvm/e7da62vv+En3zWzXU3zD3b8Ndw5idut/k/tjM7g2niZ50+zxPdfWuEqNwwZCkWokWOj/AL09cYvK0lZ4h5ayoocahDCFdNWHW+rOv+Fin/ZNnw4/8Tjvv/3goPfl49e60B/dutdbjH/Cev8AkP8ARvyz6pxfzk+XUi9ideZLdOew3UvRWNyFfjsFlp9k5qswOe3L2xW0iUeRyVEc/QSRUGGo6iKnligMtbJPFN9otSfLrfW951z1d1p09tXHbG6l692R1hsrERRwYvaXX21cHs7bWPhijWKOOjwm3qHHY2nVY0A9MY4HuvXuiO/P7+Vp8Rv5h3WG6to9vdX7Tx/ZFfiMkmxu98Dt7HY/tLYO5pqQpi81TbkoUocpn8RTVscL1eHrp5cfXxRhZEDiOWPYNOvdfJh7z6e3h8e+6O1+iewaZKXe/T3Ye7+tt0RwCcUkma2bnq7A1tZjnqIoJajFV8tCZ6SbQBPTSJIvpYH3frXRmv5YnyDyvxb/AJgfxH7txmSbF0m1+79kYvdk/wByaSOfrzeuVi2P2PQzzl0ijirti7iyEWqS8aMwZgQtvejw6919Bb/hT3JuJP5QndK4TyfwyXsfouPeGhiFG3R2ht+Wl8oDrrj/AL2x4vghvVY24uKjj1vr5ffu/WutvT/hIRszoLcHyi+Tm5N9U22sj39srrDY9b0RR5uGkqsrjtr5fN7nxvcu6NqQVZY0uXxhO2qCWrp0+5iosrNEHWKomV6t1vr6DfuvXuiL/NX+W78Of5gGzK3afyU6c29uXLPR/a4Ls/CUlHt7t7ZkiWamqNq9h0VI2bpIqeZVdqGpNViqnSFqaWeO6Hdade6Oft3b2E2jt/BbU21jKTC7c2xhsXt7b+GoI/DQ4nCYWigxuKxlHFc+KkoKCmjijX+yiAe9de60gP8AhWt/MBzaZrrP+XZ15nJaLBriMT3V8hP4fNNE+XrayqrIeqtgZCSKWLXj8ZT0U+4K2kkWSKeaoxMwKvTW92UefXutIj3brXW7L/wnq/kKdT9v9V7a+dnze2TTdgYHelTLX/HzovckcjbQr9tYuumo/wDSf2XhWMX956bcGSo5Vw2Gqw2Lmxsf3tTHWR1lMsFSfIdb63iNq7Q2nsTBUG1tkbX27s3bOKiWDF7d2rhMbt7BY2BQFWGgxGIpqPH0cSqoAWONQAPdevdFG+Yf8un4c/O3ZeY2h8jukNnbrr8jQy0uM7Ix2IxuD7a2jUtFIlLktp9i0NGNxY2eimkEoppJZ8dUsirVU1RFeM7rTr3Xyxf5mPwL3x/Lf+XnYfxm3fkX3HiMUlDu3q/fLU0dCN/9Wbmaqba253oY5ZhQZKKaiqsbkoATHDlcfVJE0kIjlewNetdFp+PPe3YPxi7y6q+QfVWUbEdgdQ73wW+NtVJaUUtRV4WsSeow+VihkieswO4KDy0GRpi2iqoamWF7o7A76919lHoPuLbPyG6P6f742ZqG1O5es9kdnbfhkniqKijxe99t47cdLj6yWG0f8QxseR+3qFspSeN1IUggN9b6Jv8AzRP5lXUf8sH43VndnYVJ/e3ee4ckdqdO9UUeR/huX7I3q1M1ZLTGuWjyJwm2Nv49Gq8rk5IHipovHCoeqqqWGbYFevdaKG1flT/P+/nX9mbuX4/7+7npdq4KqSXObf6H3rN8b+hOuaOvAmx+Aze64NzbZTP1brStPSUmbzGbzc4R5IldYyUtgde6cu7tqf8ACjb+UfR0fdG/+5vkpS9ZUeZo1qt603d7fJLpely1ZUwUVKm+Nobjzm+MLt9MtMlPSxVWcw1HDVySxU8UzzFYx7B691bttD+cv8qP5rX8snvjqz40Yvd3Xn80Pqyv6mzaYD49ZnJ7fy3Y/WK9j7Sx+7+y+u2nyMNfgqWjpauSj3Bi2ragUoqYZI5XiqxDBqlD8uvdadvzapPmNjvkFuTG/PCr7VqvkdQYfa8W5v8ATNmqrPb6gwc2EpK3acFZX1VdkZPsTgqqKSmQSFVjccAk+7fZ1rrZg/l27L/4UWU/y++I2S70r/nRL8a4+3ur6vsgbx7BylfsOTqz+L46TL/x/Hy7lqI6rbrYInyxtE4aHgqfp7qadb63+/devde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3XvfuvdVLfNH+eR/K++Boy+K7w+VOxcr2JiGraaXpzqKc9vdqrlqFgk2EzG3NjnKUmxckzX0f3lq8LTsVI8t+PfuvdakXzQ/4WudnbjTL7Y+BHxewvXGOmWSnoO2vkhkI947z8EodfvMf1TsrI02z9uZam9Jjasz+46RjfXAwFj7r3Wpl8u/5lvzw+eGSnrfld8oe1u28XJXSZGl2Pkc+2C6uxFXJOlR5sF1VtWLB9dYWZHijAkpsZHKVijBY6Ft7r3RGvfuvde9+691737r3XvfuvdWlfyRuuW7T/AJuf8u7aopfvEpPlZ1RvuopikkiSUfVW4IO0a7zRxxTeSnSj2c7SKy+MxghyE1Ee6919q737r3Xvfuvde9+690Vv5x9ht1F8KfmD2us32zdYfFv5A9hrUaZH8DbL6m3buQTaIqWtlfxHG6rLDMxtwjH0n3XuvhV+/de697917ozPxp+GPyw+ZG6Bs74t/HntrvTNx1UFJkW682ZmM3hNvyVI1Qzbt3VHTptbZ1C4I/ynK1lHTjULuLi/uvdbV/wu/wCEX/y/7NOL3N82+7+v/jDtqTx1FZ1314tL3X264SSLz4vI5TGZHF9VbYkniZvHW0eX3KsbL6qZgffuvdbbvwu/4TjfynPhSMPmNv8Axyxfe3ZGJ+1mXtD5NyUXcGdOQpEXw5PHbTymNo+rNuZCCpBmhqMbgKWqhl0sst0Qr7r3V5tPT09HTwUlJBDS0tLDFT01NTxJBT09PAixQwQQxKscMMMahVVQFVQABb37r3Wb37r3Xvfuvde9+691737r3Xvfuvde9+691BymLxubxuRwuax1Dl8Pl6GrxeWxOUpKevxuUxtfTyUldjsjQ1cctLW0NbSyvHNDIjRyRsVYEEj37r3WsL/MN/4S6fET5MHcfYvxMyEXxG7kr1q8iu2sTj3ynx83PlnVpVgrdiwFMj1utbMiRfcbclTHUUZaQYepkNjYHr3Wix84P5dXy2/l5b/i2H8nOrshtaDKy1C7O7BwzvuDq7sCClAeafZu9qWFMdX1MELpJPj6gUuVo0kQ1NLDrS9ga9a6I/7917q7D+Sx/N47J/lld84rE57L5PcPxI7R3Li6TvHriRpq2LArVtBjD25sOkLkY7e+16MI9VFEFjzuOp/s6geVKKpo9EV631vWfz7+/abYP8nD5R782JnKXJ0/bew9hdebVzeIr45sZndrd3bz2jtnN1lDX00wWsxuW62z2RkiaLyJURuoI8bMwqOPXuvlR+79a6tE/lGfPrrj+Wv8rJvk/v7o7Jd75HE9a7s2fsLCY3d2P2ZPtLc+7qnDUVfvKLKV+29y+WdNmxZTFCJIoy0WVkJew0toivXutnv/AKDI+tv+8DN8f+j+wP8A9qr3rT1vr3/QZH1t/wB4Gb4/9H9gf/tVe/aevde/6DI+tv8AvAzfH/o/sD/9qr37T17rSE7f3VtnfXbPaO99lbZl2Vs3ePYm9t1bS2bPWxZKfaW2dw7lyeXwO2ZsjBS0MGQlwOKrIqVp0hhWUxagiA6RbrXRxP5TG68ns3+Z58AMviX0VVZ8u+g9qStq03xm++yNv7HzSX0t/nMNuKdbW5va4+vvR4de628P+Fin/ZNnw4/8Tjvv/wB4KD3pePW+tAf3brXX1Ev+Ex//AG586D/8Pnvj/wB/FvD3Q8et9X++9de697917r5Tf/CibAYvbf8AOU+aWOw9KlJSVOZ6az80UaoivlN1/HPp/dGcqiI0RS9bm8xUTMbamaQkkkkm44de6pixmSrcNksfl8bN9tkcVXUmSoKjxxTeCtoaiOqpZvFPHLBL4p4lbS6sjWsQRx731rr7NnzF+MmzfmX8X+7vi/v2d6Lbfcuw8ptRsxDTrV1G285eHKbS3dR0jywxVlds/duOocpBC7rHLNSKrHST7b4db6+QN8o/jH3D8Ou9uwfjv3rtip2t2J11m6jF18Lx1Bxedxpdnwu7dr108FP/ABjae6MaY6zHVioomp5V1Kjh0VzrXQcdadn9i9M76232d1NvjdPXHYe0MgmU2xvPZmbr9vbjwlcqPE01BlcbPT1UKz08rxTJqMc0LvHIrI7Kfde624/gr/wra7g2DSYbY/z26mj7vwlL9vSS909RQ4PZ/aa0iFFlrNxdf1TYnrveWSKlrNQVO10CqAySOWc1K+nW+tvz4X/zN/hH8/sQav4y96ba3buWlo1rc31lmxU7P7W29GI0epfJbB3HFQZyrx9FI/jkyNAlbimkBEdU/B96Ip17o+3vXXuvkHfzfO4cl3r/ADO/nH2BkqySvRPkV2HsLCVL/cDybS6ky8nVGzSkVUkc9PH/AHV2XRkRsqNHexUEEe7jh1rquSLxeWPza/DrTy+LT5fFqHk8ev069N7X4v7317reO6//AOFdXSHV+w9k9abK/l9b1w+zevNo7b2NtLERfIDAtHi9s7Sw1FgMDjo2PVIJShxWPiiB/ovuunrfSu/6DI+tv+8DN8f+j+wP/wBqr37T17r3/QZH1t/3gZvj/wBH9gf/ALVXv2nr3Wvr/On/AJr3XX817fHRHYm1fjxmejd19UbV3psvcuSzG/MVviXeW3sxl8HnNo0MMlBs/a8+NTauS/jUhEr1KSnK3RYirmXYFOvdUle99a6+q3/wnT3lU70/k6fD2qr6tqzI4DH9u7NqS0VTGKem2p3x2fh8BSK9SCtQtPtanoV1xM0QN0GkqyLQ8et9aiX/AAq2+Qec7O/mU0nSj5CU7U+M/T2x9uY/CLWLPR0+7uz8bTdpbpz/ANos8v2OTzW38/gKSYMsTyU+Kp20ldLtZeHXut5b+VV8UtpfDP4CfGnpbbWEo8TmE6z2tvfs2sgo1pq3cnbW+8HjtxdgZ3LTMi1dbUjM1ZoaZqgvLT4yipaUER08aLU8evdHS7N622T3H11vjqjsnb9DurYHY+1c7sreW3MlGJaLM7c3JjajFZagmH6k89HVMFkUiSN7OhDKCNde6+Wn/Kr3Tuj4P/zveh9iY/Iz1E2E+W+Y+Iu6hd44c9id+70yvQWQ+/ghGiaODJZeDIxqRojq6SKTjQCLnI690KX/AApv/wC3wff3/hj9D/8Avm9m+/Dh17r6bXWH/MtevP8Awxtpf+6DH+6de6XPv3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r5wv/CxP53fI7bHzV69+IfVvfHZ2w+kqP4u7K3Z2X1xsfeuZ2rt/d+/d5797WiqzvOi29kqQ7koV2RjcMI6TIa4FLM4i5SQ+691pA+/de697917r3v3Xuve/de697917pXbF2BvztDdWH2L1nsnd3Ym99w1UdDgNnbF23md3bqzlbKwSKjw+3sBRZDL5OqkdgFjhhdyTYD37r3WyH8Lv+EnX8075Rri9xdsbU2n8Nuu67xzvle9shJN2NPQPHE5kxnT+1P4puajyCNLpNLuGbbrgo12Hp1e691uE/y8v+E7H8ub+T3vPZny/wC0u9dx9j96dZQ5k4Lt/uPdW3Oo+qNp5bcW1s7svM1+0ut6GvFIK7J7e3NV08UOcze5HhmkjlpfFUxxyBLeX1lt8Jub+WOG3HFnYKP2sQPy6EXK/KHNfO26psXJ223+67zJ8MFpBLcSkeZ0RKzBRxLEaQKkkAdbAG0flb8X9/CD+5PyM6M3XJU6BFTYDtfYuUrfJItO4gloaTOy1kFUoqog0MiLKjSKGUEge0NtzDsF5T6W9tJCfJZoyf2Bq1yMcehdvvsj7y8sFv6xcp8y2KpWrT7beRpQahqDtCFZe1qMpKkKSCQOh3pqmmrIIaqkqIKqlqI1lgqaaWOeCeJxdJIZomaOSNxyCpIPs3VlYBlIKnzHUZywy28rQzqyTKaMrAhgRxBByCPQ9Z/e+m+qZP8AhQz2GOsP5Lv8wHcjTeAZLpzH9ea9Mb6j272FsvqhYbSUtWo+5beojuEDLqurxsBInuvdfKf/AJYXwVn/AJk/zW6i+G9H23gek6/tdd4S0m/NwbdrN2wUa7K2Xn995Kix+3KPLYA5jNVmE23Vfa08lfRRSyqFaZLi/uvdfSL+F/8Awk6/lX/F04jcXa+1N3fMjsWgWOefLd75OKHriLIpMj+bE9PbSXE7aqca0UYQ0e4p9yISzksbqE917rZC2L1/sPq7auH2L1nsnaPXWyNvUsdDgNnbF23hto7VwdFEoSKjw+3tv0WPxGNpY0UBY4YUQAWA9+690rvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+690EfenQ3T/AMmOrt1dL979fbc7N6y3nQtQ5/au56Faujm4JpchQTqY63DZzFzETUWQo5YK2hqFWWCWORVYe6918rn+ct/LFzf8rz5X1XV+NymT3X0j2NiJd/dD70y6RDL1u1GrXocttHc8tLFDQy7w2HlR9rVyQKiVlJLR1vip/u/tobg1611Ul7317rb/AOy+/dwfIz/hJ1tBMzPU5nNfHP5EbE6CzmRLeY0239i7rp6jYEFQFZzSU2F2Bvrb+KiDabiGMgesXr+LrfWoB7t1roxvx/8AiD8o/lZ/e3/ZbOgu1e8f7h/wH++n+jLZuZ3b/df+9H8Z/u5/G/4TTVH8P/jf93a/7byW8v2kum+g2917oxv/AA0J/NF/7wF+VX/om94//W336o6917/hoT+aL/3gL8qv/RN7x/8Arb79Ude69/w0J/NF/wC8BflV/wCib3j/APW336o6917/AIaE/mi/94C/Kr/0Te8f/rb79Ude6sT/AJSf8pD554b+ZP8ADbdfdXxD786v6z6+7q292luTfW++ttxbd2tgz1VBXdj4X+JZjKUUNFSSZPcO16SjpgzBpaqoijS7uoOicdb6vL/4WKf9k2fDj/xOO+//AHgoPel49e60B/dutdfUS/4TH/8AbnzoP/w+e+P/AH8W8PdDx631f77117r3v3Xuvlaf8KQ/+30PzL/8t3/+BS6M93HDr3VHfvfWuvuH+2+t9V1fzDv5XPxO/mXdewbS+QWz5afeW3qOrg657m2e9NiO0OvJqpjM8WKzEtNVUuZ27U1BLVOHycNXjpmYyrFHUrFUR7Bp17rQy+df/CZz+YF8UKjO7p6bwcHzB6dx0dTXQbj6lx8lL2jjsbCWKruLpWqra7ctTkiqk+Pbk+449ADu8ZJRbAjr3WvBk8ZksLka7D5nH12Jy2Mq56DJYvJ0lRQZHH11LI0NTR11FVRxVNJV08yFJI5FV0YEEAj3vrXTzszeu8euN1YHfXX269x7G3ttbJU+Y2zu/aObyW3Nzbfy1KxamyWFzmIqaPJ4yugJOmWGVHF+D7917r6Qf/CfT+ddlv5hO08v8a/kjV4+P5YdTbWgztJu+FaDGUvfGwKOphxlbudcRSx0tLjt/bXqKmmXNU1LGlNVxVMdbTRov3UNNQinW+vnvfL+kylB8s/lFQ5yeSpzVF8ie7KTMVMtS9ZLUZSm7K3NDkJ5KuQtJVSS1aOxkYlnJ1Hk+7jh1roFtp7V3JvvdO2tkbNwmS3Lu/eW4MNtXau3MPSyVuX3BuTcORpsRg8JiqKENLV5LK5OsiggiUFpJZFUcn37r3R+v+GhP5ov/eAvyq/9E3vH/wCtvv1R17r3/DQn80X/ALwF+VX/AKJveP8A9bffqjr3Xv8AhoT+aL/3gL8qv/RN7x/+tvv1R17r3/DQn80X/vAX5Vf+ib3j/wDW336o6917/hoT+aL/AN4C/Kr/ANE3vH/62+/VHXuvpefyZ/j5vD4tfyxPiB0p2DgMptTfO3uuspuXd21s5TVNFm9tbi7P3xuvtTLbfzNDWAVWPy+Grd6vTVNO4VoJ42jsNNhQ8et9aEH/AApr2RldqfzhvkLnsikqUfZmzOhd74FpKcwpLiqDpXZHXEzwSeWT7uIZzr+tUyWSzqyafRqaw4de6+ld8bN74Tsz469C9jbaqYKzb2/emOsN5YOqpp/uaeoxO5tk4TM4+WGfxxeZHpa1bMUQn8qDwKde6Gh3WNWd2VERS7u5CqiqCWZmJAVVAuSeAPfuvdfJ7+P+Wh+UP8+rrjfG1NOdw3bX802k7fgmx0lRSx1uzMh8npez8tkaObGVE9RSxLtKKeoSSGY+JV1CUAeQX8uvdDn/AMKb/wDt8H39/wCGP0P/AO+b2b78OHXuvptdYf8AMtevP/DG2l/7oMf7p17pc+/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3XuvkHf8Ki+x37E/nbfLuKOdp8X1/S9KdcYnU5Ywpt/ozrutzsAUTzxRrHu/L5KypoFiCyrIX9+691r8e/de697917rNT09RV1EFJSQTVVVVTRU9NTU8TzVFRUTOscMEEMatJLNLIwVVUFmYgAX9+691eZ8Lv+E4/wDNj+azYfMbe+OOU6K63y/2sy9o/JqSt6fwIx9XIvhyeO2plMbWdp7jx89MTNDUY3AVdLNEFZZbOhb3XutsH4o/8I5vg50DhIOyf5hXyR3H3nJiIkr89tjb+Uh+PHROMiMUZqaHP7obM1XY+cpqaRH0V9PmtsF0b1UykX9tTzwW0RnuXSOFeLMQqj7SaAdGW0bNu+/7hHtOw2lze7rM1I4beJ5pXPokcas7H5KCerRcV88v5OH8s/a1f138Kun9gVGSp6b+GV2O+NfXmHxNLnqiif8AyaXfPdOZgoJ99IzLf+JCu3HUEKvLAC0c7z7rcqbXqjtpHvLgeUQ7a/ORqKR801/Z1nL7W/3cf3jvcLwr3frO25Y2OSjGTcXpcFDx02UPiTrIP993Itfmw6rQ+Qf8/L5c9nmtxfTuK2j8etuT+SOKfDU1PvvfZp5GIMdRurdGNXCxOYfSJKLDUc6MSyyA6SsUbz7v8x39Y9sWOygP8I8ST/e3Gn81RT8+ukXtZ/dhexXJojvefp7/AJq3ZaErKzWVlqHmttbSeMRXJWW7lRgAGQioNUtZJ8kvlXvOfK1MXc/yG35VSFHnhpN59n7iH3EutaSnipYczWU1N5JAI4IlSJAQqKBYe49Y75zDdGRvqr27PyeVs+QpqIHoBj06zct09ovZLl9bKFuXuVOWUFQpa026DtFNTFjEjNQdzsSxyWYmp6Z+4ehe4/j9msJtvurrrc/Wuf3Ht6DdeGw+6qE47I1eAqcjksTDXmkZ2mpr5DEVEZimEcyGO7IFZCzW57RuezSpBukEkEzprVXFCVJIrTyyCKHPy4dL+Qfc3kD3S2653f283az3ja7S7a2lltn1xrMsccpTVQBuyVGDLqQ6qBiQwCL23vfeezZhU7Q3dufatQJDKKjbefyuDmEpCKZBLjKulfyFY1F73so/p7SwXd1anVbSSRt/RYr/AICOhFu/LnL3MEfg79YWV7FSmm4gimFM4pIrCmTj5noy20/5gHze2SYxt/5Xd9JDCAsNHmOy90bnx0CKrqEhxm58hmMfDH+4TpWIAmxtcAg9t+cea7X+x3G8oPJpXcfscsP5dQ/vn3XPu5cxVO6ckcsGRuLxbfb28hOMmS3SJyccSxNMcCetmzoPJS/zkf5Pvf8A8eO38/Q5ztPeOyuz+jd3bmzFLRRim3+1J/e3pzsmpxOOoaangTb+RrcJWq0MPjkrsRMUsylEyZ9ueY5uY+XFnvX17jDI0chwCT8StQAAVRgMChIP2DgF9+j2M2r2J99J9n5WtfpORt0soL2wjBdliQgwzwiR2dmKXMMj0ZtSxyxVqCGb5fX8vftzcPwe/mXfFfs/eFNVbUyXQnyp2HQ9o4uuc01biMBjN9U+z+3cDWSQNIKeqXa1TlaOQ/uIjk3WRQVYe9Ybdfb49+691737r3Xvfuvde9+691737r3XvfuvdNGb3Bgds0D5XcmbxG38ZG4jkyObyVFiaBHZXdUesr5qenVykbEAteyk/g+25ZoYE8Sd1SP1YgD9px0v27a9z3i5FltFtPdXhFRHDG8rkYFQqBmpUgcPMdBNF8m/jbPWjGw/IPo+bImZ6YUEXbGwpK01EZZXpxSpnzOZkZSCmnUCDx7Lhv2xltAvbQvWlPGjrX7NXQ4f2d93I7f6yTlXmNbTSG1nbb0JpPBtRh00NcGtOhSwG6dsbspDX7W3Hgdy0IOk1uAy+PzNIG1yx2NTjqiphB8kLr+r9SMPqD7Xw3EFwuu3dJE9VIYftBPQL3TZd42Of6XerS5s7n+CeJ4m4A/DIqngQeHAg+Y6fvb3RZ0lt8b22n1rsvdvYm/M9j9rbI2HtrObx3huXLSmDGbf2xtrGVOYzuayEwV2jo8ZjKOWaQgE6ENgTx7917qH112R1/29snbnZPVm9Nsdh7A3djYcvtjeWzc1j9w7czmOnB8dVjctjJ6mjqUDAq4VtUcisjAMpA917pa+/de697917rS3/wCFjuS2YerPg5iKmfGv2EN/9y5LDUxdWy9Psxtu7HpdzToisXgxtZm0xKsWAEssC6SfG9rL17rQ492611uRfy3vjduPvX/hLt/Md2dR4iWbK1/f3ZneGyVSETVeTh6M2L8ZN81oxEfMklZkn6symMREvJKztGoJYA1PHrfWm77t1rrZk/4S2fNDaPxn+eO4OmexszR7e2d8uNmY7rvC5jITw0lBTdv7XzD5nrKir6yd1jhi3LT5PMYelX9U2VyVHGP1H3o8Ot9fSr90691737r3Xvfuvde9+691737r3WnR/wALFP8Asmz4cf8Aicd9/wDvBQe7Lx691oD+7da6+ol/wmP/AO3PnQf/AIfPfH/v4t4e6Hj1vq/33rr3XvfuvdfK0/4Uh/8Ab6H5l/8Alu//AMCl0Z7uOHXuqO/e+tdfcP8AbfW+ve/de697917qnn+aj/Jp+MX8y3rnclfldq4Hrz5QUGBnTrX5B7fxkNDuOPLUMDSYbb/Y32Kwf392LUVCCCSCuE1Vj4ZZJMfLTyM/k2DTr3Xyj9xYDLbU3BndrZ+jkx+d21mcngM1QSlTLQ5bDVs+OyNHIVJUyU1ZTOhsSLr7v1rqyH+TB2znumf5qPwT3Rt6ongqdw/Irr/qbILCQVqcD3flU6ez1PURtJHHLAcTviV+b6GRXUF0X3o8OvdKP+eH0bkOgP5q3zR2pVUL0VBvDuDM90YB9JFLW4bvCODtRZse3CNSUeS3XU0ZVPTDNSyRADx2Hhw691Wj17vfO9Zb+2P2RtedaXc3X279tb327UsCVp87tTNUWexE7BSrFYshQRsbEHj3vr3X2WPid8metvmN8depfkn1NkYq7Zna+z8XuSnpBVQ1ddtrLywiHcmzM48AWOPcWzc/FU4yvQAKKqlcrdCrFvrfRiPfuvde9+691737r3Xvfuvde9+691qpf8KdP5V2/Pl/1TsT5c/HvaWS3p3d8ecFk9rb72Pt+kmyG5t/9H1VbV7iim23joPLVZjcHWe46utrIMbTR/cV1DmK4p5Z4KeCWwPXuqgP5Nv/AApMwnwp6N258TPmL132DvzrPrlq7H9UdndaLhsxvfae3KyvqK+PYu7dq7ozW24M5gcDXVkwoK6nySVVBQ6KJaSaOGEp4jr3Q/8A80L/AIVObD7k6C3r0H8Ddg9p7XznaW3q/aO8u7u0qPA7UyO0tq52kmodw0HW+3Nubk3TWTbly+LqXpVzFXVULYsPJJTQSz+Cpg8F9evdcv8AhLd/Ki7Bouwqb+ZN3ttSu2rtXDbczeF+LeEzlN9pl925XdePrtt7p7ZbG1cH3VFtWg2vWVeNw8zBGycuQmqYrQU8MlT4ny691Uv/AMKb/wDt8H39/wCGP0P/AO+b2b72OHXuvptdYf8AMtevP/DG2l/7oMf7p17pc+/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3XuviK/zb+yP9Lf80P8AmC7+jqPuqHLfL/v+hwtTq1efbe2uydw7W2xLe5A17ewtMdIJC/QEgX9+690pPhl/J2/mR/PlsVX/ABt+KvZG4NjZWRRD23u7Hx9b9PiDxLUT1NN2RvuXA7azv2tM6yPT4uavrCroEhZpEDe691tofEH/AIRW7cwVDR75/mI/LpUo6GniyOb6z+NtNTYXD0CQEVMyZnu7s/ESPLjzGBFVR021qJlXWYa0HTIKu6RqXkIVAKkk0AHzJ6ftbW6vrhLOyjkmu5GCoiKXdmPBVVQWYnyABJ6ud6z3D/IH/lIrS4z4rdP9a7t7bxgNPBubqTby98dx1dbUwHH1CSfIDfWWya45MhUKwqMbQ7ihiiaRvHQqpVPYD3j3L5T2gmITm6uf4IB4mf8AT1EfHy11+XWY/th9wf7yXuXGt/Ls68vbCRU3O8O1kAoFSRbFHvSNOQxthGf9+DNMu7/5in81z5QCoxXw6+FO6epdsZLUmK7C3ntmbK5mWHSBFWUG6+yqTZ/UtC88cgkeGSjyXiuoWZgC7gu5519w9/rHyztcltA3CR1q32h5QkI9aENT19crNh+6h9yL2aKX3v77iWW+7xDQyWNpcCOIHzV7bb2utzcAjSHWW31UJMYrpUutd/J0/mafKzOUm6flh8gtt0BV5JEpt79gbk7KzWBafSJIdvbU25QybDxNKwS7xUeTpIrkWVrsQSv7ac98wyi45ivUHyeRpWX/AEqKPDA+SsB1LFt9/v7nfsltr7L7H8q3coIALWdlb7fDNp4Ge5ncXsrZw0tvI1OLCgBOD1Z/wnW+O+32part3uztPsqqgKSSUO1Mft7rTA1bAnXDWU9RHvrOtTFTb9jI00lxfWB6fYl2/wBldlho25XVxOw8kCxKftH6jU+xgfn1AnOn96/7r7oHh5E5d2XZ4WqA9y8+4TL6FWU2cOr/AE8Ei0NNNc9WP9XfytvgL1E1LUbZ+M/X+YyFL43XJ9hw5LtCrapiHprVj7CyG5KCkqg51qaeGFY3AKBSq2G+38gcn7bQwWELOPOSspr6/qFgD9gFPLrEnnT76H3n+e1eLeOcN0t7V6jw7Ex7cuk/grYpbuy0wRI7lhUMWqanpwuCwm28dT4fbuHxWBxNIgSkxeFx9Ji8dSoAFCU9FQwwU0KBVAsqgWHsWRRRQIIoVVIxwCgAD7AMdY17huW47tdvf7rcTXN9IatJK7SSMfVncliftPVCf/Cgv48Hfvxu2P8AIHD0RmznRO7Ri9xTRKqn/R72RNj8PU1FQwPkqDi96UeHSFSCIo66oe6gtqiH3k2X6zY4t5iFZbSSjf8ANOWikn1o4SnpqY9dOP7rP3W/qx7ubl7W38mnbeZbHxIAf+U6wDyqq+S+JaPdFzUamhiXJpTTj94z9d+uve/de6vj/kBfIYda/KncvSGXrfBt75BbRkp8ZFI6rCvYXXkOS3HgGaSV1jgSt2xPm6ey+ueqemQXOke5c9nt6+h5gfapTSG8jx/zUjqy/tTWPmdI65mf3oftUeb/AGUs/cewj1brytfhpCBn6G+McE+AKkpcLZvU4SMTNgVPWrp/wqV+Gf8Aspn82btzdWAwq4rrj5Y4nFfJfaLU0MKUZ3Ju+WqxHb9M0lOqRHJVHa+DyuWljKrJHBl4C2rWJHye6+fPr6cX8s35Br8qv5e/wz+QUtZ9/luzPjn1ZmN21Hljntv6i2rj8H2JS+aP0y/Y76xWRg1EIzeO7IjXRfde6PH7917r3v3Xuve/de6w1NTT0dPUVlZUQ0tJSwy1NVVVMqQU9NTwI0s9RUTyssUMMMSlndiFVQSTb3pmVVLMQFAqSeAHTkMMtxKsECs87sFVVBLMxNAqgVJJJoAMk4HWq9/MI/nrbhbOZ3qL4S1lJjMPjJqrFZ7v6soqPKV2bqopJKeri6vxWQhqsZR4RApVM1VxTT1ZYvSRU6xxVU+P3OXu1N4r7byqQsakhrggEseB8IGoC/0yCTxULQMe1f3V/wC7U2obbbc9/eKjkmv5lWWHZFd40hUgMp3GRCsjTHibSJkSOgWeSUs8EeuF2B2f2R2xnZtz9ob+3l2JuKcuZM3vXcuY3Pk7SadUaVmZrKyeKEBFAjQqiqoAAAAEJXl/fbjMbi/mlmnP4nZmP7WJ6618rcm8o8j7YuzcmbXt+1bStKQ2lvFbx4rkpEqAnJJYgkkkk1J6Q3tJ0JenjA7hz+1snT5rbGczG3MxSHVS5bA5OtxGTpmDK4anr8fPT1UJDIDdXHIH9PbsM01vIJYHZJRwKkgj8xQ9F+57Vte9WbbfvNtb3e3v8UU0aSxt5dyOGU8TxHV0X8v7+bL82No9vdWdN5vcdb8jdp9gb22tsKl2t2VXz1+7KSo3Tm8fhqeswfZEkNZuanqKRqhTpyTZKhSEP+yhtIkn8ne4nNNtuVvtkrm9t5pUjCSmrjWwWqy5bH9LUtK4HEc9vvR/cc+7tvvIe9e4G3WkfKW+bXt1zetc7eipbMttDJKyTWAK27K2k5txbzF9P6jDsY8v/CqL5qD4/wDwUw/xq2vlmo+w/l/uWTblfHSytHW0XTfX9RiNw9h1XlhcPTrn8zVYXCtHINFZQV9cgv43AynHHr52etF74U/zLvmj/L63BPlvjD3TnNoYLJ11PXbm62zMNLuzq7drwMms5zY2eirMTFXVNOvgbI0IostHCxWGri4IsRXrXWzP07/wsa37j8LT0Pf3wl2nuzcEdNGKjc3UXbmX2FiqmqREWTRsveGz+xKqmiqH1Pq/jshisF0vfUutPW+lF2Z/wsgyc+36ml6c+ClDi90zwTCkznZneNRntv4yp02p3qdq7W6523kc5AXbU6rmcewC6Qx1ak9p691qbfM35sfIj58d15bvn5J70/vZvKupY8Phsdj6NMRtLZG1qWpqqrHbQ2Vt+F5YsNt/HS1krKGeaqqZZHnqp56iSSZ7cOtdAB1519vbtnfez+setts5XeW/9/7jw+0dm7UwdOarLbg3Hnq6HHYnFUEN1Vp6ysqEQFmVEBLOyqCR7r3X17/5b/wyw3wa+DPRPxSmXGZrI7L2RMeyqyKM1mK3J2Hvatr909lVKfeKz12Em3PnqumoxMv/ABbYoYyqqoQUOT1vr5pX85r+W5un+W18yN67Bo8Lkh0F2Nkstv745bslimlx2R2Dka0VE2ypckzzrPufrCsrRiK9JHWpmhjpq5o44q6EGwNevdVMwTz0s8NVSzS01TTSxz09RBI8M8E8LiSKaGWMrJFLFIoZWUgqRcc+99a62x/5f3/CrH5EfH/amE6t+Y3XD/KbaWCpKfG4ftLFbjTaveFBQU4WONd1VWSo8ntrs96emiWOKaoGIyUjFpaquq391K9b6utoP+Fcv8tSowsuRruqvmRjsnAsIbb5626jq6yrlZITMcbWw97DFSU0UkjAPUzUkjrGT4wSoPtJ690Rj5O/8LCMZNt/KYb4c/FPMUu4qyOeHGdgfIncOMSiwoKlIqp+sOvK7J/xirOvWobdFPDC6DUk6sVHtPXugo/4Tz/znu9ey/n92b058yu3812Efmg0OV2PnNyVNPSYfaHdmzMZP/BdrbWwtFFQbc2dtrfWy4JsdFRUUMML5PG4uGGLyVEjP4jGOvdb6nuvXutOj/hYp/2TZ8OP/E477/8AeCg92Xj17rQH926119RL/hMf/wBufOg//D574/8Afxbw90PHrfV/vvXXuve/de6+Vp/wpD/7fQ/Mv/y3f/4FLoz3ccOvdUd+99a6+4RVVVNQ01RW1tRBR0dHBNVVdXVTR09NS01PG0s9RUTyskUEEESFndiFVQSSAPbfW+vnR7x/4VBfLLrr+YJ8g+2uqqzFdtfD7dG/Fwmyeg+wRXUOHTrzZFPFtbb+7Nj5yFajPdabu3vj8c2Xrwi1mOetyD/c0NQ0URjvTHXur5ekP+FYn8tzsHD0Tdu7d726A3MYl/i1DmNk0/Y21oKrwSTOuF3J1/kcnncrRiRBEstTg8dKZHBMQTU4rpPXuoHyZ/4Vdfy+uu+vc5U/G7HdnfIXtWpxVSm0MPU7FzHXWwaLOyJOlHPvncG8mwm4IMPSOiySpisdXz1AKxK0Op5ofaT17r5ze6dy5fee59x7w3BULWZ7deey+5c3VrFHAtVl87kKjKZKoWGJVihWetqnYKoCrewFvd+tdWv/AMh747bi+R/81P4k4jC46eqxXUnY+H+RO8sjGtT9rgNu9IZCi3tQ5HIS0xDwwZDelDiMXEWPjesyMMb3VyDo8OvdbUH/AAqm/lq57vTqfZ/zy6e21Lmt+fHzb1btTvLGYqmlqMrmOjPvKrOYveEdNTxSSVK9U52urpa4galxOUnqJGENBxpT5db6+fF7t1rq1L+WZ/N9+Vv8rzdeUl6fyOK3v1Du/I0+S7C6I36ayo2RuKvhhSiO4sFV0M0OW2RvVMcghXI0L+KpWKBa+mroqeCJNEV631txdT/8K/PhPn8VQr3P8cfkp1puSaOIV0Wxo+uO1NpUc2idqgjO5PeXWm4ZoNSRiMrhWdmkIYKE1NrT17r3aH/Cv/4SYPGV3+iD42fJvsbcNP8AdpR02+E6y6s2xXyR6BSOufxW9u0M7TUtUSxZnwvliAX9tixC+09e610fk9/wpD/mA/Invnp/tTC5jDdKdddH9lbb7K2j0b11VZSDbm5sjgJ/36Ttfc1RIme7Dpc3i56mgqKaQUuKWlqWMVDHPeZt0HXuvpMfGj5A7A+VfQPUfyL6urfvti9w7Hwm9cHrkjkq8b/E6YfxPb2U8JMcWc2xmIqjG18QP7NbSyp9V90690OPv3Xuve/de6qB+cX8tP8Ak5dj1OT7n+Z/RXQu0svlquqr8x2K+7s70Nn925lzqqa3L5brLd2wMnv7cUz1S6jVfxCrmJjUhrIAiv8AdNv2qH6jcp4oIfV2C1+Qqcn5Cp6F/Jnt/wA8e4u6jZOQ9p3Dd91NKx2kEkxUGtGkKKRGmDV5CqAAkkAHoAvgD/LE/kZ72pMr3V8Uvi1tHsKm6739U7AG9e2J+1+xaCs3dt3CbX3RJmcLtPu/cufwsf28O5qR4ayPEUbCoVjEoVVdkex8w7ZzHbSXe0yGS2jmMZbSVBYKrGgYAkUcZIGa/b0Jvdz2W9wfY3fLLln3KtI7Hf77bI79YFmimZIJZriBBK0LPGsmu2kJRXbSumpDEqL94oooIo4II44YYY0ihhiRY4ooo1CRxxxoAiRogAAAAAFh7Oeoq6If3r/K+/l//JrsnMdwd9/Fbq7tHszP0uIoczvLc9Bkp8xkKTA4ymw2HgqJKbJ00Rjx+Lo4oY7ILIgvf36vXuj00FDSYuhosbj4I6Wgx1JT0NFSxAiKmpKSFKemgjBJIjhhjVRcnge/de6l+/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3XuqdPjB/IU/lcfFndmV7MwPxp232/wByZ3cmW3dmO4fkX4O5t5T7kzGZqs7PmsTjNyUY692blIa+scx1OBweLqAttTseffuvdWb72wXbWcaTG7F3xtHrbDLBFTjKvsmo3zuuQMiNJUYla/cG39rbamo2UxRrVUGeimQ6ysRAT2X3cW5TEpazRwR/xaDI/wBoqyotOGVkB444dDTlvcORdrRbnmLbb7dr7UT4Qu1s7ZaVAWTRBPcTqwoxMc1kyEaQXHd0VvO/y6+luzK1cj8jN6d5fJ2ojmjq4MX292pmaTY9DXRz/cCox3WHVsXW3W1Oof0hHxUq6FXVqZVYEE3Je137a97lu781rSaVhGDxqIovCiH+8Hy9Opp2z71/uFyfbm09p9v5b5NiKlWk2vbYmvHQrp0vuO4m/wBwbGai5U1LUoGI6Mn1l8dOgul4oo+pel+ruuWij8f3eztjbbwOSnGlVL1mVx+OhydfM6oA0k0skjWFyfZ5YbLs+1im3WtvB80jVT+ZAqftJ6iLnH3Y9z/cJy/PPMO9bsCa6bq8uJox50SN5DGgFTRUVVHkB0Mvsz6j/r3v3Xuve/de697917r3v3Xugv7s6q2/3l1D2X09ulf9wPZWydx7NyE4jSWagXO4yooYMrSLICgr8RVSx1VO39ieFGHI9oN02+Hdttn2y4/sZ4mQ/LUCKj5g5HzA6GXt3ztuntvz5s/P2yn/AHZ7PuMF2gqQH8GRXaNqZ0SqDHIPNGYefXzXN9bL3D1xvbeHXu7aI43dOxd0Z7Z+5MexYmizu2spVYbLUuplRmEFfRyKDYXAvb3g5d2s1jdS2VwNNxDIyMPRlJUj9o6+vLlrmHaubeXbDmrY5PG2XcrKG6t3/jhuI1libieKOppXHSV9p+jvoQep+ydxdOdn9fdr7TmMG5Oud5bd3phj5ZIY5q7buVpcpFR1TREM9DX/AGxgqEN1lgkdGBViCs26+m2y/h3G2NJ4JVdftUg0PyNKH1BI6CvPHKO1c/8AJu68kb4uraN22+e0lwCQk8bRl1rwdNWtGwVdVYEEA9XYf8K1/jjt35gfyt+h/n11lRnL1nx73Dtbe4ydOhlnfoT5IUO3MDnlmipkaWapxu+U2lO2v0UVOlax03c+85LG8h3Czivrc1gmjV1PyYAj+R6+Qvm7ljdeSuaty5P3xPD3na76e0nX0lt5WienqNSkqfMUIwejCf8ACPn5BDtr+UynVFbXNLl/jF392h1vDRTM71EW096yYrufBV6sXkH2NTm+xMvSwi6srUDroChGdV0Hutqj37r3Xvfuvde9+691Qr/Pr+WWb6Z6A2r0JsjKyYzdHyGqc3TbrrKOQpW0fVe2o6Fc/jUliljmozvHK5alomazJUY+GugIs5IiH3e5il2zZ49otW03F6WDkcREtNQ+WskL81Djz66b/wB2P7Hbd7g+6N77ncxwCbZuVEha2RhVH3K4L+BIQQQ/0sUUkoGCk720gPbnTM94x9fQR1YV8Ff5bnevzvy2Xqtjvi9l9ZbXrUxu6u0d0x1UmGpMrJTJVpt/AYyjArdzbhFLLHNLBE0MFLDIjVE8JlgWUZ8pcj7tzdIzWmmKwjNHleukGldKgZZqUJAoACNRFRXFb7yv3ufbX7s9jBBzGJtx5xvYzJbbdbFRK0QYqZ55G7LeDUGVXYM8jqwiicJIUufP/CbXbv2HjX5d5oZTwqv3h6ToWoPuLDVL/DR2ktR4Sb2j+71D/Vn3J/8ArHwaKfvJ/Epx8AUr9ni1/n+fXPcf3um7fVazyJb/AEWr4P3u+vT6eJ+7tNf6XhU/ojoi3yD/AJC3zA6moazPdXV+0PkJgaSOaaWi2nNLtffsdPTxiSWc7Q3JKtDXllJEcGOylfWSspVYSSoYJ7z7Q8y7chm28xXsI8k7JMf0GwfkFdmPp1kp7Wf3m/sLzxcx7ZzpFf8AKu5yEAPcgXNkWY0A+qtxrT+k89tDEoNTJQEhQ/yKPjDnd1fNLcvYW99tZPD0/wAY9tV9XkMXn8ZXY3IYvsrekWQ2ptrFZTFV9PBLSVlLhhmqwLMqyw1FFEwTV6ke9pdhmuOaJL27jZRYRkkMCCJXqigg8CF1nOQVGPQq/vK/eXbNl+71Z8q8uXkNxLzjeIqSQSJJHJt9oUubiSORGYMrS/SRVQlXjlcFqYYWP+FB38kX5hfPrtnEfKf45dhbe7HqNldZ4rYVD8bN0VkOy8pjqDDV+XzFXW9b7oyNYNnZnLblyualnq4MvJhmURoq1k4SGCPKAGnXz5daIPeXxr+QPxm3TPsr5BdMdldObngnlgXGdhbPze2TXGJUdp8PWZKkhoM5QvFIrx1NHLPTyxsro7KwJt1roEvfuvde9+690fj4g/ywfnT858vj6T46/HnfW5dtVlTTw1fZ+dxsuzupcPDLJEs1VX9i7lXHbcqWo4JPM9JQzVmSkiUmGmlaynRNOvdfQV/k6fyF+nP5Z9PB3B2Pl8R3b8vcriJsdV9gQ0EqbG6soclC0WWwPT+OytLBlUqa+nlalrdw1scORrqXVFDBj6eapp56k1631sBe9de6J384vgp8dv5hPRuY6H+Ru02zm36iV8ttTc+JmTG72633fHR1NHjt6bGzpinONzVAlSwaKaOegroWanrKeop3eJvcOvdfOx+f/wDwnN+fnwzzWbz/AFrsfL/LPoynlqqnF9h9M4Csy29MZio2d4hvzqKhmym8cJWU9IjS1NTjUy+IgjUs9ZGToFwR17qhDJ4zJYXIVeJzGPrsTlMfO9LX43J0lRQZCiqYjpkp6ujqo4qimnjPDI6qwP1HvfWuoPv3Xunrbu2txbvzNDt3aeAzW6NwZOUQY3BbdxVdmszkJyLiGhxmNgqa2rlIH6Y0Y+/de62Hv5b3/Cd7+ZJ392F172vu3E7g+D2wtqbn27vPF9o9k0Vbgu4KDIbfy1PlsZleueqDLQbyg3Pi6+jhqaSozX8ColIWWKeUqI20SOt9fTQpYpYKamhnqZKyeGCGKasljhilqpY41SSpkip44qeOSdwWKxqqAmygCw90691qEf8ACvvbu4NxfHD4fQ7fwWYzs1N3bvqWoiw+MrcnLBE+xIUWSaOignaKNm4BYAE+7Lx691oTf6Meyf8An3u+P/QTz3/1B7t1rr6dX/CaPE5XCfyiOh8dmcZkMRkIt796NLQ5OjqaCsiWXt/d8kTSU1VHFMiyRsGUlRdSCOPdDx631fb7117r3v3Xuvlvf8KMdi73zP8AOU+YuSxGzt1ZXHVP+y9/b1+N29l66in8PxY6Pgl8NVS0csEviniZG0sdLqQeQfdxw691ST/ox7J/597vj/0E89/9Qe99a6+wb/MI6U7w+R/wv+Q/Q3xz35tXrXtrtvrzJbCwW7d6U+Wk2/S4jck1Njd54yqq8FDV5bCz7i2XPkMfBkYKarlx89UtQsLtGB7bHW+vlY/MD+WH86Pgrla2m+R/x53xtTbdNOYqXs3CUQ3n1LlY2cilloux9qtldr009ZFaRaKtnpMlErATU0TgqLg1610Qj3vr3XvfuvdWA/Db+V185vnjuDFYz489B7xzG1shVwQV3bW6cdWbO6c2/TSMhqK/KdiZump8LV/ZUz+ZqLGmvys0Y/yekmcqjaJp17r6Rv8AKB/lE9UfyrencriaLKUnZHyD7LTH1HcncJxzUMVcmOeeXE7H2PQ1LS1eD2HgJKl3tI/3WUrGaqqdIFLS0dSa9b6t8nggqoJqWqhiqaapikgqKeeNJoJ4JkMc0M0MgaOWKWNirKwIYGx496691pJ/zav+EuWS3TubdvyF/lrpgaObO1NZuDdPxOzWQx22cXFlamUzVsvRe5q96Lb2Hx9bNIZRt3MTUdFRHyCirkgNNjobA+vXutMvu/44d/fGnddTsj5A9NdldObppp5YP4T2Js7ObXkrfESPusRUZSjp6LOY2ZRrhq6OSelniIeOR0YMbda6BX37r3XvfuvdH8+Kn8rn58fNLI46D4//ABk7M3Lt7IshHY2ewk2xuq6WBgHeom7I3l/BNo1DRQHyfb01VUVkiW8cLkqDqo6919LL+TR/L67Q/lrfDjHfHntjuek7d3HU71z+/mosBQVNNsbrSTc9LjP4nsfYlflYabP5vBtmKKfJS1dVT0Imrq+dkpIdTtLUmvW+rYfeuvdVD/zqPkl3R8YfibtbePRm9KjYW692d3ba6+yu4KLG4fIZKPbGV2D2Zn6+mxsmZoMjFi62fIbYpCtXAiVUSqwjkQsT7jf3R3zdNh5djudplMNxJdrGWAUnSY5WIGoGhqoyMjyI6zv/ALvL2i9vfeX3xvdg9ydvXc9kseXLi+jgeSVIzcRXu3wI0gieMyIEuJaxOTGxI1qwFOtIvfXYe/ez9w1e7eyd67r3/umut95uLee4ctubN1CrfRHLk8zV1lY0UYNkTXpQcAAe8Vru9vL+Y3N9LJNcHizsWY/mxJ6+jHlrlXljk3ak2LlHbrHa9li+GC0git4V9SI4lRKnzNKk5JJ63Af+E7f/AGRT2h/4tJvX/wB9N0l7yT9lv+VWuP8ApYP/ANWYOuC/965/4kRs3/il2n/dz3fq+v3L3XMfr3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917rSU/nv/Hf/RD8x/8AShiaH7banyI2zT7yjkih8NJHvrba0m2990MPqby1EyJjctUPxqmy7cfk4se7my/u3mb94Rilvex6/l4i0WQf8dc/N+vor/uz/db+vnsD/Uy+l175ypeNakE1Y2dxquLNz6KCbi2QeSWoz6Um+4s66Kde9+691tw/yq67aHz3/ld/IT4H9o1K1UGL2r2F0hl/IiVNZRdad24HcFRtDcEBMyVC5Lb2eq8slCy6DTfwumMcgZRoyj9oN6/eHLbbbIaz2Umn/m3JVkP+9a1+xR188X9517VHkr32h59sY9Ozc12ImJAoPrbMJb3SgUpmI2kzGtWkmckDiaEv+Ed2792fG750fzH/AOXr2eP4Pveh2/BnspgKp3iWg358Ye0cz1Nv2gx61EUTyVlS/aEbunplkp8b5NBWJykr9c3evoU+/de697917r3v3XutNr/hRPkK2X5i9S4mSodsdRfGjbOQpaUhdENbk+0u26avqFIUOXqYMRTK1yQBELAc3xl96XY8zW0ZPYLFSB8zLMD+2g/Z13//ALqG0t09gt8vlUC7k5wuI2bNSke3bYyL6UUyyEYr3GvlSgf3D/XUDrf3/lD4PaOE/l4/HH+58dEIMvt7P5zPVNI6TSVm7q/eG4P7yyV86+uWtpMlC1JpcloIqdIRZY1AzC9t4raLkux+mpRkZmI83LtqqfUHHyAA8uvl7+/duW+7j96zm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e697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3XugS+SXbdT0L0F3B3TSYSDctV1f19ube0G36mukxkGYl2/jZ69MfNkIqatkoo6ow6TIsUhS99J+nsq3zcW2jZ7ndFQSNbws+kmmrSK0rQ0r60PUi+0XIsXud7n7B7eT3LWcO9brb2bTqgkaITyBDIELIHK1rpLLXhUdav8Au7/hRr31X0s8Wxvj31LtmqdFWCq3Pnt37zSAmORZJDTY2fZIlfylWS7hVCkMGvcQLc+9m7upFpZW0berM7/yGjrsvsX90x7ZW0yvzJzVvl5ADlbeG1tC2RQapBeUFKg4qaggrTNQXyk+b3yU+Y2Xoch3p2HV57E4aokqdu7KxFJTbf2Nt2eSN4XqcbtzGpFTT5IwyvGa6saqrzExjM5jsojbf+a985mkD7tMXjU1VAAsa/MKMV/pNVqYrTrPH2X+7l7Q+wNhLa+221JbX1woWe7lZp7ycAghZJ5CWEdQG8GIRw6gGEeqpIGdV9Xb67q7D2j1X1pt+t3Rvje+apMFt/DUKFnnqqp/3KmqmP7NBi8dTq9RWVcxSnpKWKSaVkjRmBTt9hd7pex7fYoZLuVwqqPU+Z9ABlicAAk4HUk8686cte3fKl/zrzhdR2XLe227TTyucBVGFUcXkkakcUSgvLIyxorOwB+jD8W+iML8ZPj51N0VgpYKun662hj8PkcnTwGmizm5J/Jkt2bhWnYs8Az+566rrBGxZoxMFJNr+81tg2iLYdmttpiIKwRBSRjU3F2/2zkt+fXyce9HuZuHvF7p757lbkrRy7tfvKkbHUYbcUjtoC3A+BbpFFUABtFaCtOh89nHUYde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3RMv5if8A2Qn8tf8AxAvY/wD7zlb7DHOv/Kpbj/zxy/8AHT1kF91H/wASW5G/8Wew/wCr6dfPV2xsreW9q5MZszaW5t3ZKWWOCLH7YwOVz9dJPNq8UKUmKpauoeWXSdKhbtY294Z29rdXT+HaxySP6IpY/sAPX1S7zzFy/wAuWxvOYb6zsLRVLF7iaOBABxJaVlUAeZrQdWZfHX+TV84+/K2iqMr1xJ0bs+Yo9XurucVW1auKC4aRKLYwp6nfVXXNCCYVmoKWldrB6iIHUB3svtjzXvDgyQfSWx4vPVD+UdDIT6VUD1YdYe+7H94F92/2wt5IrHdxzJv61C2206blSfIveals1SvxFZpJAKlYnIodrj4K/wAtnof4KYWpqtnpVb57XztAuP3X25uakpoM3W0ReGabC7axUElTS7Q2zLVQLK1LFLPUTuqfc1NR4ofHkLylyPtHKURa2rLuLijzMBqI/hUZCLXNASTjUzUFOIn3lfvd+5n3ldxSHfym28kW0uu22y3ZjCj0IE1xIQrXVwFJUSMqIgLeDDFrk12G+xn1ir1737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+690n91f8AHu5b/qFP/Q6e2Z/7Fvs6M9m/5KkH+n6UHt7os697917r3v3Xuve/de697917r3v3Xuve/de697917r3v3Xuv/9k="
        $base64sos = "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"
        
        $bytes = [System.Convert]::FromBase64String($base64sos)
        Remove-Variable base64sos

        $CompanyLogo = -join($ReportPath,'\','c1B-logo.jpg')
		$p = New-Object IO.MemoryStream($bytes, 0, $bytes.length)
		$p.Write($bytes, 0, $bytes.length)
        Add-Type -AssemblyName System.Drawing
		$picture = [System.Drawing.Image]::FromStream($p, $true)
		$picture.Save($CompanyLogo)

        Remove-Variable bytes
        Remove-Variable p
        Remove-Variable picture
        
        #$CompanyLogo = ".\c1b-logo.jpg"

        $LinkToFile = $false
        $SaveWithDocument = $true
        $Left = 0
        $Top = 0
        $Width = 135
        $Height = 50

        # Add image to the Sheet
        $worksheet.Shapes.AddPicture($CompanyLogo, $LinkToFile, $SaveWithDocument, $Left, $Top, $Width, $Height) | Out-Null

        Remove-Variable LinkToFile
        Remove-Variable SaveWithDocument
        Remove-Variable Left
        Remove-Variable Top
        Remove-Variable Width
        Remove-Variable Height

        $row = 5
        $column = 1
        $worksheet.Cells.Item($row,$column)= "Table of Contents"
        $worksheet.Cells.Item($row,$column).Style = "Heading 2"
        $row++

        For($i=2; $i -le $workbook.Worksheets.Count; $i++)
        {
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item($row,$column) , "" , "'$($workbook.Worksheets.Item($i).Name)'!A1", "", $workbook.Worksheets.Item($i).Name) | Out-Null
            $row++
        }

        $row++
        $worksheet.Cells.Item($row, 1) = "© WinSys Credit One Bank 2018"
        $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item($row,2) , "https://www.creditonebank.com/", "" , "", "www.creditonebank.com/") | Out-Null

        $usedRange = $worksheet.UsedRange
        $usedRange.EntireColumn.AutoFit() | Out-Null

        $excel.Windows.Item(1).Displaygridlines=$false

        $ADStatFileName = -join($ExcelPath,'\',$DomainName,'ADRecon-Report','.xlsx')
        Try
        {
            # Disable prompt if file exists
            $excel.DisplayAlerts = $False
            $workbook.SaveAs($ADStatFileName)
            Write-Output "[+] Excelsheet Saved to: $ADStatFileName"
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject] $excel) | Out-Null
    }
}

Function Get-ADRDomain
{
<#
.SYNOPSIS
    Returns information of the current (or specified) domain.

.DESCRIPTION
    Returns information of the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.PARAMETER ReportPath
    [string]
    Path for ADRecon output folder.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER objDomainRootDSE
    [DirectoryServices.DirectoryEntry]
    RootDSE Directory Entry object.

.PARAMETER DCIP
    [string]
    IP Address of the Domain Controller.

.PARAMETER creds
    [Management.Automation.PSCredential]
    Credentials.

.PARAMETER OutputType
    [array]
    Output Type.

.OUTPUTS
    CSV file is created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [bool] $UseAltCreds,

        [Parameter(Mandatory = $true)]
        [string] $ReportPath,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomainRootDSE,

        [Parameter(Mandatory = $false)]
        [string] $DCIP,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $creds = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $true)]
        [array] $OutputType
    )

    Write-Output "[-] Domain"
    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADDomain = Get-ADDomain
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        If ($ADDomain)
        {
            $ADDomainObj = New-Object PSObject
            $ADDomainObj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Value"
            $ADDomainObj | Add-Member -MemberType NoteProperty -Name "Name" -Value $ADDomain.DNSRoot
            $ADDomainObj | Add-Member -MemberType NoteProperty -Name "NetBIOS" -Value $ADDomain.NetBIOSName
            # Values taken from https://technet.microsoft.com/en-us/library/hh852281(v=wps.630).aspx
            $FLAD = @{
	            0 = "Windows2000";
	            1 = "Windows2003/Interim";
	            2 = "Windows2003";
	            3 = "Windows2008";
	            4 = "Windows2008R2";
	            5 = "Windows2012";
	            6 = "Windows2012R2";
	            7 = "Windows2016"
            }
            $DomainMode = $FLAD[[convert]::ToInt32($ADDomain.DomainMode)] + "Domain"
            Remove-Variable FLAD
            If ($DomainMode)
            {
                $ADDomainObj | Add-Member -MemberType NoteProperty -Name "Functional Level" -Value $DomainMode
                Remove-Variable DomainMode
            }
            Else
            {
                $ADDomainObj | Add-Member -MemberType NoteProperty -Name "Functional Level" -Value $ADDomain.DomainMode
            }
            $ADDomainObj | Add-Member -MemberType NoteProperty -Name "DomainSID "-Value $ADDomain.DomainSID.Value
            For($i=0; $i -lt $ADDomain.ReplicaDirectoryServers.Count; $i++)
            {
                $ADDomainObj | Add-Member -MemberType NoteProperty -Name "Domain Controller -$i" -Value $ADDomain.ReplicaDirectoryServers[$i]
            }
            For($i=0; $i -lt $ADDomain.ReadOnlyReplicaDirectoryServers.Count; $i++)
            {
                $ADDomainObj | Add-Member -MemberType NoteProperty -Name "Read Only Domain Controller -$i" -Value $ADDomain.ReadOnlyReplicaDirectoryServers[$i]
            }

            Try
            {
                $ADForest = Get-ADForest $ADDomain.Forest
            }
            Catch
            {
                Write-Output "[EXCEPTION] $($_.Exception.Message)"
            }

            If (-Not $ADForest)
            {
                Try
                {
                    $ADForest = Get-ADForest -Server $DCIP
                }
                Catch
                {
                    Write-Output "[EXCEPTION] $($_.Exception.Message)"
                }
            }
            If ($ADForest)
            {
                $DomainCreation = Get-ADObject -SearchBase "$($ADForest.PartitionsContainer)" -LDAPFilter "(&(objectClass=crossRef)(systemFlags=3)(Name=$($ADDomain.Name)))" -Properties whenCreated
                If (-Not $DomainCreation)
                {
                    $DomainCreation = Get-ADObject -SearchBase "$($ADForest.PartitionsContainer)" -LDAPFilter "(&(objectClass=crossRef)(systemFlags=3)(Name=$($ADDomain.NetBIOSName)))" -Properties whenCreated
                }
                Remove-Variable ADForest
            }
            # Get RIDAvailablePool
            Try
            {
                $RIDManager = Get-ADObject -Identity "CN=RID Manager$,CN=System,$($ADDomain.DistinguishedName)" -Properties rIDAvailablePool
                $RIDproperty = $RIDManager.rIDAvailablePool
                [int32] $totalSIDS = $($RIDproperty) / ([math]::Pow(2,32))
                [int64] $temp64val = $totalSIDS * ([math]::Pow(2,32))
                $RIDsIssued = [int32]($($RIDproperty) - $temp64val)
                $RIDsRemaining = $totalSIDS - $RIDsIssued
                Remove-Variable RIDManager
                Remove-Variable RIDproperty
                Remove-Variable totalSIDS
                Remove-Variable temp64val
            }
            Catch
            {
                Write-Output "[EXCEPTION] $($_.Exception.Message)"
            }
            If ($DomainCreation)
            {
                $ADDomainObj | Add-Member -MemberType NoteProperty -Name "Creation Date" -Value $DomainCreation.whenCreated
                Remove-Variable DomainCreation
            }
            If ($RIDsIssued)
            {
                $ADDomainObj | Add-Member -MemberType NoteProperty -Name "RIDs Issued" -Value $RIDsIssued
                Remove-Variable RIDsIssued
            }
            If ($RIDsRemaining)
            {
                $ADDomainObj | Add-Member -MemberType NoteProperty -Name "RIDs Remaining" -Value $RIDsRemaining
                Remove-Variable RIDsRemaining
            }
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        If ($UseAltCreds)
        {
            $DomainFQDN = Get-DNtoFQDN($objDomain.distinguishedName)
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain",$($DomainFQDN),$($creds.UserName),$($creds.GetNetworkCredential().password))
            Try
            {
                $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            Catch
            {
                Write-Output "[EXCEPTION] $($_.Exception.Message)"
                Return $null
            }
            Remove-Variable DomainContext
            # Get RIDAvailablePool
            $SearchPath = "CN=RID Manager$,CN=System"
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DCIP)/$SearchPath,$($objDomain.distinguishedName)", $creds.UserName,$creds.GetNetworkCredential().Password
            $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
            $objSearcherPath.PropertiesToLoad.AddRange(("ridavailablepool"))
            $objSearcherResult = $objSearcherPath.FindAll()
            $RIDproperty = $objSearcherResult.Properties.ridavailablepool
            [int32] $totalSIDS = $($RIDproperty) / ([math]::Pow(2,32))
            [int64] $temp64val = $totalSIDS * ([math]::Pow(2,32))
            $RIDsIssued = [int32]($($RIDproperty) - $temp64val)
            $RIDsRemaining = $totalSIDS - $RIDsIssued
            Remove-Variable SearchPath
            $objSearchPath.Dispose()
            $objSearcherPath.Dispose()
            $objSearcherResult.Dispose()
            Remove-Variable RIDproperty
            Remove-Variable totalSIDS
            Remove-Variable temp64val
            $ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest",$($ADDomain.Forest),$($creds.UserName),$($creds.GetNetworkCredential().password))
            Try
            {
                $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestContext)
            }
            Catch
            {
                Write-Output "[EXCEPTION] $($_.Exception.Message)"
                Return $null
            }
            Remove-Variable ForestContext

            $GlobalCatalog = $ADForest.FindGlobalCatalog()
            If ($GlobalCatalog)
            {
                $DN = "GC://$($GlobalCatalog.IPAddress)/$($objDomain.distinguishedname)"
                $ADObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ($($DN),$($creds.UserName),$($creds.GetNetworkCredential().password))
                $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($ADObject.objectSid[0], 0)
                $ADObject.Dispose()
            }
        }
        Else
        {
            $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
            Try
            {
                $GlobalCatalog = $ADForest.FindGlobalCatalog()
                $DN = "GC://$($GlobalCatalog)/$($objDomain.distinguishedname)"
                $ADObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ($DN)
                $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($ADObject.objectSid[0], 0)
                $ADObject.dispose()
            }
            Catch
            {
                Write-Output "[EXCEPTION] $($_.Exception.Message)"
                $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($objDomain.objectSid[0], 0)
            }
            # Get RIDAvailablePool
            $RIDManager = ([ADSI]"LDAP://CN=RID Manager$,CN=System,$($objDomain.distinguishedName)")
            $RIDproperty = $ObjDomain.ConvertLargeIntegerToInt64($RIDManager.Properties.rIDAvailablePool.value)
            [int32] $totalSIDS = $($RIDproperty) / ([math]::Pow(2,32))
            [int64] $temp64val = $totalSIDS * ([math]::Pow(2,32))
            $RIDsIssued = [int32]($($RIDproperty) - $temp64val)
            $RIDsRemaining = $totalSIDS - $RIDsIssued
            Remove-Variable RIDManager
            Remove-Variable RIDproperty
            Remove-Variable totalSIDS
            Remove-Variable temp64val
        }

        If ($ADDomain)
        {
            $ADDomainObj = New-Object PSObject
            $ADDomainObj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Value"
            $ADDomainObj | Add-Member -MemberType NoteProperty -Name "Name" -Value $ADDomain.Name
            $ADDomainObj | Add-Member -MemberType NoteProperty -Name "NetBIOS" -Value $objDomain.name.value
            # Values taken from https://technet.microsoft.com/en-us/library/hh852281(v=wps.630).aspx
            $FLAD = @{
	            0 = "Windows2000";
	            1 = "Windows2003/Interim";
	            2 = "Windows2003";
	            3 = "Windows2008";
	            4 = "Windows2008R2";
	            5 = "Windows2012";
	            6 = "Windows2012R2";
	            7 = "Windows2016"
            }
            $DomainMode = $FLAD[[convert]::ToInt32($objDomainRootDSE.domainFunctionality,10)] + "Domain"
            Remove-Variable FLAD
            $ADDomainObj | Add-Member -MemberType NoteProperty -Name "Functional Level" -Value $DomainMode
            Remove-Variable DomainMode
            $ADDomainObj | Add-Member -MemberType NoteProperty -Name "DomainSID "-Value $ADDomainSID.Value
            For($i=0; $i -lt $ADDomain.DomainControllers.Count; $i++)
            {
                $ADDomainObj | Add-Member -MemberType NoteProperty -Name "Domain Controller -$i" -Value $ADDomain.DomainControllers[$i]
            }
            $ADDomainObj | Add-Member -MemberType NoteProperty -Name "Creation Date" -Value $objDomain.whencreated.value
            If ($RIDsIssued)
            {
                $ADDomainObj | Add-Member -MemberType NoteProperty -Name "RIDs Issued" -Value $RIDsIssued
                Remove-Variable RIDsIssued
            }
            If ($RIDsRemaining)
            {
                $ADDomainObj | Add-Member -MemberType NoteProperty -Name "RIDs Remaining" -Value $RIDsRemaining
                Remove-Variable RIDsRemaining
            }
        }
    }

    If ($ADDomainObj)
    {
        Switch ($OutputType)
        {
            'STDOUT' { $ADDomainObj }
            'CSV'
            {
                Write-Verbose "[+] Domain"
                $ADFileName  = -join($ReportPath,'\','Domain','.csv')
                Try
                {
                    $ADDomainObj | Export-Csv -Path $ADFileName -NoTypeInformation
                }
                Catch
                {
                    Write-Output "[ERROR] Failed to Export CSV File"
                    Write-Output "[EXCEPTION] $($_.Exception.Message)"
                }
                Remove-Variable ADDomainObj
                Remove-Variable ADFileName
            }
        }
    }
}

Function Get-ADRForest
{
<#
.SYNOPSIS
    Returns information of the current (or specified) forest.

.DESCRIPTION
    Returns information of the current (or specified) forest.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.PARAMETER ReportPath
    [string]
    Path for ADRecon output folder.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER objDomainRootDSE
    [DirectoryServices.DirectoryEntry]
    RootDSE Directory Entry object.

.PARAMETER DCIP
    [string]
    IP Address of the Domain Controller.

.PARAMETER creds
    [Management.Automation.PSCredential]
    Credentials.

.PARAMETER OutputType
    [array]
    Output Type.

.OUTPUTS
    CSV file is created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [bool] $UseAltCreds,

        [Parameter(Mandatory = $true)]
        [string] $ReportPath,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomainRootDSE,

        [Parameter(Mandatory = $false)]
        [string] $DCIP,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $creds = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $true)]
        [array] $OutputType
    )

    Write-Output "[-] Forest"
    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADDomain = Get-ADDomain
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }

        Try
        {
            $ADForest = Get-ADForest $ADDomain.Forest
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }
        Remove-Variable ADDomain

        If (-Not $ADForest)
        {
            Try
            {
                $ADForest = Get-ADForest -Server $DCIP
            }
            Catch
            {
                Write-Output "[EXCEPTION] $($_.Exception.Message)"
                Return $null
            }
        }

        If ($ADForest)
        {
            Try
            {
                $ADRecycleBin = Get-ADOptionalFeature -Filter 'name -like "Recycle Bin Feature"'
            }
            Catch
            {
                Write-Output "[EXCEPTION] $($_.Exception.Message)"
            }

            Try
            {
                $ADForestCNC = (Get-ADRootDSE).configurationNamingContext
                $ADForestDSCP = Get-ADObject -Identity "CN=Directory Service,CN=Windows NT,CN=Services,$($ADForestCNC)" -Partition $ADForestCNC -Properties *
                $ADForestTombstoneLifetime = $ADForestDSCP.tombstoneLifetime
                Remove-Variable ADForestCNC
                Remove-Variable ADForestDSCP
            }
            Catch
            {
                Write-Output "[EXCEPTION] $($_.Exception.Message)"
            }

            $ADForestObj = New-Object PSObject
            $ADForestObj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Value"
            $ADForestObj | Add-Member -MemberType NoteProperty -Name "Name" -Value $ADForest.Name
            # Values taken from https://technet.microsoft.com/en-us/library/hh852281(v=wps.630).aspx
            $FLAD = @{
                0 = "Windows2000";
                1 = "Windows2003/Interim";
                2 = "Windows2003";
                3 = "Windows2008";
                4 = "Windows2008R2";
                5 = "Windows2012";
                6 = "Windows2012R2";
                7 = "Windows2016"
            }
            $ForestMode = $FLAD[[convert]::ToInt32($ADForest.ForestMode)] + "Forest"
            Remove-Variable FLAD
            If ($ForestMode)
            {
                $ADForestObj | Add-Member -MemberType NoteProperty -Name "Functional Level" -Value $ForestMode
                Remove-Variable ForestMode
            }
            Else
            {
                $ADForestObj | Add-Member -MemberType NoteProperty -Name "Functional Level" -Value $ADForest.ForestMode
            }
            $ADForestObj | Add-Member -MemberType NoteProperty -Name "Domain Naming Master" -Value $ADForest.DomainNamingMaster
            $ADForestObj | Add-Member -MemberType NoteProperty -Name "Schema Master" -Value $ADForest.SchemaMaster
            $ADForestObj | Add-Member -MemberType NoteProperty -Name "RootDomain" -Value $ADForest.RootDomain
            $ADForestObj | Add-Member -MemberType NoteProperty -Name "Domain Count" -Value $ADForest.Domains.Count
            $ADForestObj | Add-Member -MemberType NoteProperty -Name "Site Count" -Value $ADForest.Sites.Count
            $ADForestObj | Add-Member -MemberType NoteProperty -Name "Global Catalog Count" -Value $ADForest.GlobalCatalogs.Count
            For($i=0; $i -lt $ADForest.Domains.Count; $i++)
            {
                $ADForestObj | Add-Member -MemberType NoteProperty -Name "Domain -$i" -Value $ADForest.Domains[$i]
            }
            For($i=0; $i -lt $ADForest.Sites.Count; $i++)
            {
                $ADForestObj | Add-Member -MemberType NoteProperty -Name "Site -$i" -Value $ADForest.Sites[$i]
            }
            For($i=0; $i -lt $ADForest.GlobalCatalogs.Count; $i++)
            {
                $ADForestObj | Add-Member -MemberType NoteProperty -Name "GlobalCatalog -$i" -Value $ADForest.GlobalCatalogs[$i]
            }
            Remove-Variable ADForest
            If ($ADRecycleBin)
            {
                If ($ADRecycleBin.EnabledScopes.Count -eq 0)
                {
                    $ADForestObj | Add-Member -MemberType NoteProperty -Name "Recycle Bin Enabled" -Value $false
                }
                Else
                {
                    $ADForestObj | Add-Member -MemberType NoteProperty -Name "Recycle Bin Enabled" -Value $true
                }
                Remove-Variable ADRecycleBin
            }
            If ($ADForestTombstoneLifetime)
            {
                $ADForestObj | Add-Member -MemberType NoteProperty -Name "Tombstone Lifetime" -Value $ADForestTombstoneLifetime
                Remove-Variable ADForestTombstoneLifetime
            }
            Else
            {
                $ADForestObj | Add-Member -MemberType NoteProperty -Name "Tombstone Lifetime" -Value "Not Retrieved"
            }
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        If ($UseAltCreds)
        {
            $DomainFQDN = Get-DNtoFQDN($objDomain.distinguishedName)
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain",$($DomainFQDN),$($creds.UserName),$($creds.GetNetworkCredential().password))
            Try
            {
                $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            Catch
            {
                Write-Output "[EXCEPTION] $($_.Exception.Message)"
                Return $null
            }
            Remove-Variable DomainContext

            $ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest",$($ADDomain.Forest),$($creds.UserName),$($creds.GetNetworkCredential().password))
            Remove-Variable ADDomain
            Try
            {
                $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestContext)
            }
            Catch
            {
                Write-Output "[EXCEPTION] $($_.Exception.Message)"
                Return $null
            }
            Remove-Variable ForestContext

            # Check AD Recycle Bin Status
            Try
            {
                $SearchPath = "CN=Recycle Bin Feature,CN=Optional Features,CN=Directory Service,CN=Windows NT,CN=Services,CN=Configuration"
                $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DCIP)/$($SearchPath),$($objDomain.distinguishedName)", $creds.UserName,$creds.GetNetworkCredential().Password
                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                $ADRecycleBin = $objSearcherPath.FindAll()
                Remove-Variable SearchPath
                $objSearchPath.Dispose()
                $objSearcherPath.Dispose()
            }
            Catch
            {
                Write-Output "[EXCEPTION] $($_.Exception.Message)"
            }

            # Get Tombstone Lifetime
            $SearchPath = "CN=Directory Service,CN=Windows NT,CN=Services"
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DCIP)/$SearchPath,$($objDomainRootDSE.configurationNamingContext)", $creds.UserName,$creds.GetNetworkCredential().Password
            $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
            $objSearcherPath.Filter="(name=Directory Service)"
            $objSearcherResult = $objSearcherPath.FindAll()
            $ADForestTombstoneLifetime = $objSearcherResult.Properties.tombstoneLifetime
            Remove-Variable SearchPath
            $objSearchPath.Dispose()
            $objSearcherPath.Dispose()
            $objSearcherResult.Dispose()

        }
        Else
        {
            $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()

            # Check AD Recycle Bin Status
            $ADRecycleBin = ([ADSI]"LDAP://CN=Recycle Bin Feature,CN=Optional Features,CN=Directory Service,CN=Windows NT,CN=Services,CN=Configuration,$($objDomain.distinguishedName)")

            # Get Tombstone Lifetime
            $ADForestTombstoneLifetime = ([ADSI]"LDAP://CN=Directory Service,CN=Windows NT,CN=Services,$($objDomainRootDSE.configurationNamingContext)").tombstoneLifetime.value

        }

        # Values taken from https://technet.microsoft.com/en-us/library/hh852281(v=wps.630).aspx
        $FLAD = @{
	        0 = "Windows2000";
	        1 = "Windows2003/Interim";
	        2 = "Windows2003";
	        3 = "Windows2008";
	        4 = "Windows2008R2";
	        5 = "Windows2012";
	        6 = "Windows2012R2";
	        7 = "Windows2016"
        }

        If ($ADForest)
        {
            $ADForestObj = New-Object PSObject
            $ADForestObj | Add-Member -MemberType NoteProperty -Name "Category" -Value "Value"
            $ADForestObj | Add-Member -MemberType NoteProperty -Name "Name" -Value $ADForest.Name
            $ForestMode = $FLAD[[convert]::ToInt32($objDomainRootDSE.forestFunctionality,10)] + "Forest"
            $ADForestObj | Add-Member -MemberType NoteProperty -Name "Functional Level" -Value $ForestMode
            Remove-Variable ForestMode
            $ADForestObj | Add-Member -MemberType NoteProperty -Name "Domain Naming Master" -Value $ADForest.NamingRoleOwner
            $ADForestObj | Add-Member -MemberType NoteProperty -Name "Schema Master" -Value $ADForest.SchemaRoleOwner
            $ADForestObj | Add-Member -MemberType NoteProperty -Name "RootDomain" -Value $ADForest.RootDomain
            $ADForestObj | Add-Member -MemberType NoteProperty -Name "Domain Count" -Value $ADForest.Domains.Count
            $ADForestObj | Add-Member -MemberType NoteProperty -Name "Site Count" -Value $ADForest.Sites.Count
            $ADForestObj | Add-Member -MemberType NoteProperty -Name "Global Catalog Count" -Value $ADForest.GlobalCatalogs.Count
            For($i=0; $i -lt $ADForest.Domains.Count; $i++)
            {
                $ADForestObj | Add-Member -MemberType NoteProperty -Name "Domain -$i" -Value $ADForest.Domains[$i]
            }
            For($i=0; $i -lt $ADForest.Sites.Count; $i++)
            {
                $ADForestObj | Add-Member -MemberType NoteProperty -Name "Site -$i" -Value $ADForest.Sites[$i]
            }
            For($i=0; $i -lt $ADForest.GlobalCatalogs.Count; $i++)
            {
                $ADForestObj | Add-Member -MemberType NoteProperty -Name "GlobalCatalog -$i" -Value $ADForest.GlobalCatalogs[$i]
            }
            If ($ADRecycleBin)
            {
                If ($ADRecycleBin.Properties.EnabledScopes.Count -eq 0)
                {
                    $ADForestObj | Add-Member -MemberType NoteProperty -Name "Recycle Bin Enabled" -Value $false
                }
                Else
                {
                    $ADForestObj | Add-Member -MemberType NoteProperty -Name "Recycle Bin Enabled" -Value $true
                }
                Remove-Variable ADRecycleBin
            }
            If ($ADForestTombstoneLifetime)
            {
                $ADForestObj | Add-Member -MemberType NoteProperty -Name "Tombstone Lifetime" -Value $ADForestTombstoneLifetime
                Remove-Variable ADForestTombstoneLifetime
            }
            Else
            {
                $ADForestObj | Add-Member -MemberType NoteProperty -Name "Tombstone Lifetime" -Value "Not Retrieved"
            }
            Remove-Variable ADForest
        }
    }

    If ($ADForestObj)
    {
        Switch ($OutputType)
        {
            'STDOUT' { $ADForestObj }
            'CSV'
            {
                Write-Verbose "[+] Forest"
                $ADFileName  = -join($ReportPath,'\','Forest','.csv')
                Try
                {
                    $ADForestObj | Export-Csv -Path $ADFileName -NoTypeInformation
                }
                Catch
                {
                    Write-Output "[ERROR] Failed to Export CSV File"
                    Write-Output "[EXCEPTION] $($_.Exception.Message)"
                }
                Remove-Variable ADForestObj
                Remove-Variable ADFileName
            }
        }
    }
}

Function Get-ADRPassPol
{
<#
.SYNOPSIS
    Returns the Default Password Policy of the current (or specified) domain.

.DESCRIPTION
    Returns the Default Password Policy of the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.PARAMETER ReportPath
    [string]
    Path for ADRecon output folder.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER OutputType
    [array]
    Output Type.

.OUTPUTS
    CSV file is created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [bool] $UseAltCreds,

        [Parameter(Mandatory = $true)]
        [string] $ReportPath,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [array] $OutputType
    )

    Write-Output "[-] Default Password Policy"
    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADpasspolicy = Get-ADDefaultDomainPasswordPolicy
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }

        If ($ADpasspolicy)
        {
            $ADPassPolObj = New-Object PSObject
            $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Policy" -Value "Value"
            $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Enforce password history" -Value $ADpasspolicy.PasswordHistoryCount
            $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Maximum password age (days)" -Value $ADpasspolicy.MaxPasswordAge.days
            $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Minimum password age (days)" -Value $ADpasspolicy.MinPasswordAge.days
            $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Minimum password length" -Value $ADpasspolicy.MinPasswordLength
            $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Password must meet complexity requirements" -Value $ADpasspolicy.ComplexityEnabled
            $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Store password using reversible encryption for all users in the domain" -Value $ADpasspolicy.ReversibleEncryptionEnabled
            $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Account lockout duration (mins)" -Value $ADpasspolicy.LockoutDuration.minutes
            $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Account lockout threshold" -Value $ADpasspolicy.LockoutThreshold
            $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Reset account lockout counter after (mins)" -Value $ADpasspolicy.LockoutObservationWindow.minutes
            Remove-Variable ADpasspolicy
        }

        Write-Output "[-] Fine Grained Password Policy - May need a Privileged Account"
        Try
        {
            $ADFinepasspolicy = Get-ADFineGrainedPasswordPolicy -Filter *
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }

        If ($ADFinepasspolicy)
        {
            $i = 0
            $ADFinepasspolicy | ForEach-Object {
                $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Name -$i" -Value $($_.Name)
                $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Applies To -$i" -Value $($_.AppliesTo)
                $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Enforce password history -$i" -Value $_.PasswordHistoryCount
                $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Maximum password age (days) -$i" -Value $_.MaxPasswordAge.days
                $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Minimum password age (days) -$i" -Value $_.MinPasswordAge.days
                $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Minimum password length -$i" -Value $_.MinPasswordLength
                $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Password must meet complexity requirements -$i" -Value $_.ComplexityEnabled
                $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Store password using reversible encryption -$i" -Value $_.ReversibleEncryptionEnabled
                $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Account lockout duration (mins) -$i" -Value $_.LockoutDuration.minutes
                $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Account lockout threshold -$i" -Value $_.LockoutThreshold
                $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Reset account lockout counter after (mins) -$i" -Value $_.LockoutObservationWindow.minutes
                $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Precedence -$i" -Value $($_.Precedence)
                $i ++
            }
            Remove-Variable ADFinepasspolicy
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        If ($ObjDomain)
        {
            #Value taken from https://msdn.microsoft.com/en-us/library/ms679431(v=vs.85).aspx
            $pwdProperties = @{
                "DOMAIN_PASSWORD_COMPLEX" = 1;
                "DOMAIN_PASSWORD_NO_ANON_CHANGE" = 2;
                "DOMAIN_PASSWORD_NO_CLEAR_CHANGE" = 4;
                "DOMAIN_LOCKOUT_ADMINS" = 8;
                "DOMAIN_PASSWORD_STORE_CLEARTEXT" = 16;
                "DOMAIN_REFUSE_PASSWORD_CHANGE" = 32
            }

            $ADPassPolObj = New-Object PSObject
            $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Policy" -Value "Value"
            $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Enforce password history" -Value $ObjDomain.PwdHistoryLength.value
            $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Maximum password age (days)" -Value $($ObjDomain.ConvertLargeIntegerToInt64($ObjDomain.maxpwdage.value) /-864000000000)
            $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Minimum password age (days)" -Value $($ObjDomain.ConvertLargeIntegerToInt64($ObjDomain.minpwdage.value) /-864000000000)
            $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Minimum password length" -Value $ObjDomain.MinPwdLength.value
            If (($ObjDomain.pwdproperties.value -band $pwdProperties["DOMAIN_PASSWORD_COMPLEX"]) -eq $pwdProperties["DOMAIN_PASSWORD_COMPLEX"])
            {
                $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Password must meet complexity requirements" -Value $true
            }
            Else
            {
                $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Password must meet complexity requirements" -Value $false
            }
            If (($ObjDomain.pwdproperties.value -band $pwdProperties["DOMAIN_PASSWORD_STORE_CLEARTEXT"]) -eq $pwdProperties["DOMAIN_PASSWORD_STORE_CLEARTEXT"])
            {
                $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Store password using reversible encryption for all users in the domain" -Value $true
            }
            Else
            {
                $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Store password using reversible encryption for all users in the domain" -Value $false
            }
            Remove-Variable pwdProperties
            $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Account lockout duration (mins)" -Value $($ObjDomain.ConvertLargeIntegerToInt64($ObjDomain.lockoutduration.value)/-600000000)
            $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Account lockout threshold" -Value $ObjDomain.LockoutThreshold.value
            $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Reset account lockout counter after (mins)" -Value $($ObjDomain.ConvertLargeIntegerToInt64($ObjDomain.lockoutobservationWindow.value)/-600000000)

            Write-Output "[-] Fine Grained Password Policy - May need a Privileged Account"
            $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
            $ObjSearcher.PageSize = $PageSize
            $ObjSearcher.Filter = "(objectClass=msDS-PasswordSettings)"
            #$ObjSearcher.PropertiesToLoad.AddRange(("admincount","canonicalname","description","distinguishedname","lastLogontimestamp","name","objectsid","primarygroupid","pwdLastSet","samaccountName","serviceprincipalname","sidhistory","useraccountcontrol","userworkstations","whenchanged","whencreated"))
            $ObjSearcher.SearchScope = "Subtree"
            Try
            {
                $ADFinepasspolicy = $ObjSearcher.FindAll()
            }
            Catch
            {
                Write-Output "[EXCEPTION] $($_.Exception.Message)"
                Return $null
            }

            If ($ADFinepasspolicy)
            {
                $i = 0
                $ADFinepasspolicy | ForEach-Object {
                    $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Name -$i" -Value $($_.Properties.name)
                    $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Applies To -$i" -Value $($_.Properties.'msds-psoappliesto')
                    $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Enforce password history -$i" -Value $($_.Properties.'msds-passwordhistorylength')
                    $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Maximum password age (days) -$i" -Value $($($_.Properties.'msds-maximumpasswordage') /-864000000000)
                    $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Minimum password age (days) -$i" -Value $($($_.Properties.'msds-minimumpasswordage') /-864000000000)
                    $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Minimum password length -$i" -Value $($_.Properties.'msds-minimumpasswordlength')
                    $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Password must meet complexity requirements -$i" -Value $($_.Properties.'msds-passwordcomplexityenabled')
                    $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Store password using reversible encryption -$i" -Value $($_.Properties.'msds-passwordreversibleencryptionenabled')
                    $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Account lockout duration (mins) -$i" -Value $($($_.Properties.'msds-lockoutduration')/-600000000)
                    $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Account lockout threshold -$i" -Value $($_.Properties.'msds-lockoutthreshold')
                    $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Reset account lockout counter after (mins) -$i" -Value $($($_.Properties.'msds-lockoutobservationwindow')/-600000000)
                    $ADPassPolObj | Add-Member -MemberType NoteProperty -Name "Precedence -$i" -Value $($_.Properties.'msds-passwordsettingsprecedence')
                    $i ++
                }
                Remove-Variable ADFinepasspolicy
            }
        }
    }

    If ($ADPassPolObj)
    {
        Switch ($OutputType)
        {
            'STDOUT' { $ADPassPolObj }
            'CSV'
            {
                Write-Verbose "[+] Default Password Policy"
                $ADFileName = -join($ReportPath,'\','DefaultPasswordPolicy','.csv')
                Try
                {
                    $ADPassPolObj | Export-Csv -Path $ADFileName -NoTypeInformation
                }
                Catch
                {
                    Write-Output "[ERROR] Failed to Export CSV File"
                    Write-Output "[EXCEPTION] $($_.Exception.Message)"
                }
                Remove-Variable ADPassPolObj
                Remove-Variable ADFileName
            }
        }
    }
}

Function Get-ADRDC
{
<#
.SYNOPSIS
    Returns the domain controllers for the current (or specified) forest.

.DESCRIPTION
    Returns the domain controllers for the current (or specified) forest.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.PARAMETER ReportPath
    [string]
    Path for ADRecon output folder.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER OutputType
    [array]
    Output Type.

.OUTPUTS
    CSV file is created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [bool] $UseAltCreds,

        [Parameter(Mandatory = $true)]
        [string] $ReportPath,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [array] $OutputType
    )

    Write-Output "[-] Domain Controllers"
    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $allDCs = Get-ADDomainController -Filter *
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }

        # DC Info
        If ($allDCs)
        {
            # DC Info
            $DCObj = @()
            $allDCs | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Domain" -Value $_.Domain
                $Obj | Add-Member -MemberType NoteProperty -Name "Site" -Value $_.Site
                $Obj | Add-Member -MemberType NoteProperty -Name "Name" -Value $_.Name
                $Obj | Add-Member -MemberType NoteProperty -Name "IPv4Address" -Value $_.IPv4Address
                $OSVersion = $_.OperatingSystem + $_.OperatingSystemHotfix + $_.OperatingSystemServicePack + $_.OperatingSystemVersion
                $Obj | Add-Member -MemberType NoteProperty -Name "Operating System" -Value $OSVersion
                Remove-Variable OSVersion
                $Obj | Add-Member -MemberType NoteProperty -Name "Hostname" -Value $_.HostName
                If ($_.OperationMasterRoles -like 'InfrastructureMaster')
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Infra" -Value $true
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Infra" -Value $false
                }
                If ($_.OperationMasterRoles -like 'DomainNamingMaster')
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Naming" -Value $true
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Naming" -Value $false
                }
                If ($_.OperationMasterRoles -like 'SchemaMaster')
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Schema" -Value $true
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Schema" -Value $false
                }
                If ($_.OperationMasterRoles -like 'RIDMaster')
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "RID" -Value $true
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "RID" -Value $false
                }
                If ($_.OperationMasterRoles -like 'PDCEmulator')
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "PDC" -Value $true
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "PDC" -Value $false
                }
                $DCSMBObj = [PingCastle.Scanners.SmbScanner]::GetPSObject($_.IPv4Address)
                ForEach ($Property in $DCSMBObj.psobject.Properties)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name $Property.Name -Value $Property.value
                }
                $DCObj += $Obj
            }
            Remove-Variable allDCs
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        If ($UseAltCreds)
        {
            $DomainFQDN = Get-DNtoFQDN($objDomain.distinguishedName)
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain",$($DomainFQDN),$($creds.UserName),$($creds.GetNetworkCredential().password))
            Try
            {
                $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            Catch
            {
                Write-Output "[EXCEPTION] $($_.Exception.Message)"
                Return $null
            }
            Remove-Variable DomainContext
        }
        Else
        {
            $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        }

        If ($ADDomain.DomainControllers)
        {
            # DC Info
            $DCObj = @()
            $ADDomain.DomainControllers | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name "Domain" -Value $_.Domain
                $Obj | Add-Member -MemberType NoteProperty -Name "Site" -Value $_.SiteName
                $Obj | Add-Member -MemberType NoteProperty -Name "Name" -Value $_.Name
                $Obj | Add-Member -MemberType NoteProperty -Name "IPAddress" -Value $_.IPAddress
                $Obj | Add-Member -MemberType NoteProperty -Name "Operating System" -Value $_.OSVersion
                $Obj | Add-Member -MemberType NoteProperty -Name "Hostname" -Value $_.Hostname
                If ($null -ne $_.Roles)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Infra" -Value $($_.Roles.Contains("InfrastructureRole"))
                    $Obj | Add-Member -MemberType NoteProperty -Name "Naming" -Value $($_.Roles.Contains("NamingRole"))
                    $Obj | Add-Member -MemberType NoteProperty -Name "Schema" -Value $($_.Roles.Contains("SchemaRole"))
                    $Obj | Add-Member -MemberType NoteProperty -Name "RID" -Value $($_.Roles.Contains("RidRole"))
                    $Obj | Add-Member -MemberType NoteProperty -Name "PDC" -Value $($_.Roles.Contains("PdcRole"))
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name "Infra" -Value $false
                    $Obj | Add-Member -MemberType NoteProperty -Name "Naming" -Value $false
                    $Obj | Add-Member -MemberType NoteProperty -Name "Schema" -Value $false
                    $Obj | Add-Member -MemberType NoteProperty -Name "RID" -Value $false
                    $Obj | Add-Member -MemberType NoteProperty -Name "PDC" -Value $false
                }
                $DCSMBObj = [PingCastle.Scanners.SmbScanner]::GetPSObject($_.IPAddress)
                ForEach ($Property in $DCSMBObj.psobject.Properties)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name $Property.Name -Value $Property.value
                }
                $DCObj += $Obj
            }
            Remove-Variable ADDomain
        }
    }

    If ($DCObj)
    {
        Switch ($OutputType)
        {
            'STDOUT' { $DCObj }
            'CSV'
            {
                Write-Verbose "[+] Domain Controllers"
                $ADFileName  = -join($ReportPath,'\','DCs','.csv')
                Try
                {
                    $DCObj | Export-Csv -Path $ADFileName -NoTypeInformation
                }
                Catch
                {
                    Write-Output "[ERROR] Failed to Export CSV File"
                    Write-Output "[EXCEPTION] $($_.Exception.Message)"
                }
                Remove-Variable DCObj
                Remove-Variable ADFileName
            }
        }
    }
}

Function Get-ADRUser
{
<#
.SYNOPSIS
    Returns all users in the current (or specified) domain.

.DESCRIPTION
    Returns all users in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.PARAMETER ReportPath
    [string]
    Path for ADRecon output folder.

.PARAMETER date
    [DateTime]
    Date when ADRecon was executed.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER DormantTimeSpan
    [int]
    Timespan for Dormant accounts. Default 90 days.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.PARAMETER FlushCount
    [int]
    The number of processed objects which will be flushed to disk. Default -1 (After all objects are processed).

.PARAMETER OutputType
    [array]
    Output Type.

.OUTPUTS
    CSV file is created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [bool] $UseAltCreds,

        [Parameter(Mandatory = $true)]
        [string] $ReportPath,

        [Parameter(Mandatory = $true)]
        [DateTime] $date,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $DormantTimeSpan = 90,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10,

        [Parameter(Mandatory = $false)]
        [int] $FlushCount = -1,

        [Parameter(Mandatory = $true)]
        [array] $OutputType
    )

    Write-Output "[-] Domain Users - May take some time"
    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADUsers = Get-ADUser -Filter * -ResultPageSize $PageSize -Properties AdminCount,AllowReversiblePasswordEncryption,CannotChangePassword,CanonicalName,Description,DistinguishedName,DoesNotRequirePreAuth,Enabled,LastLogonDate,LockedOut,LogonWorkstations,Name,PasswordLastSet,PasswordNeverExpires,PasswordNotRequired,primaryGroupID,pwdlastset,SamAccountName,SID,SIDHistory,TrustedForDelegation,TrustedToAuthForDelegation,whenChanged,whenCreated
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
            Continue
        }

        If ($ADUsers)
        {
            Try
            {
                $ADpasspolicy = Get-ADDefaultDomainPasswordPolicy
                $PassMaxAge = $ADpasspolicy.MaxPasswordAge.days
                Remove-Variable ADpasspolicy
            }
            Catch
            {
                Write-Output "[EXCEPTION] $($_.Exception.Message)"
                $PassMaxAge = 90
            }

            $ADFileName  = -join($ReportPath,'\','Users','.csv')
            [ADRecon.ADWSClass]::UserParser($ADUsers, $date, $PassMaxAge, $ADFileName, $DormantTimeSpan, $Threads, $FlushCount, $OutputType)
            Remove-Variable ADUsers
            Write-Verbose "[+] Domain Users"
        }
        Write-Verbose "[+] Domain Users"
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(samAccountType=805306368)"
        $ObjSearcher.PropertiesToLoad.AddRange(("admincount","canonicalname","description","distinguishedname","lastLogontimestamp","name","objectsid","primarygroupid","pwdLastSet","samaccountName","serviceprincipalname","sidhistory","useraccountcontrol","userworkstations","whenchanged","whencreated"))
        $ObjSearcher.SearchScope = "Subtree"
        Try
        {
            $ADUsers = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADUsers)
        {
            $PassMaxAge = $($ObjDomain.ConvertLargeIntegerToInt64($ObjDomain.maxpwdage.value) /-864000000000)
            If (-Not $PassMaxAge)
            {
                $PassMaxAge = 90
            }
            $icnt = 1
            $cnt = $($ADUsers | Measure-Object | Select-Object -ExpandProperty Count)
            Write-Output "[*] Calculating if the user Cannot Change Password"
            $CannotChangePassword = New-Object 'System.Collections.Generic.Dictionary[String,bool]'
            $StopWatch = [System.Diagnostics.StopWatch]::StartNew()
            $ADUsers | ForEach-Object {
                If ($StopWatch.Elapsed.TotalMilliseconds -ge 1000)
                {
                    Write-Progress -Activity "Calculating if the user Cannot Change Password" -Status "$("{0:N2}" -f (($icnt/$cnt*100),2)) % Complete:" -PercentComplete 100
                    $StopWatch.Reset()
                    $StopWatch.Start()
                }
                # Get ACLs to determine if the user can change their password or not
                $data = $_.GetDirectoryEntry()
                $aclObject = $data.Get_ObjectSecurity()
                ForEach ($access in $aclObject.Access)
                {
                    If (($access.ObjectType -eq "ab721a53-1e2f-11d0-9819-00aa0040529b") -or ($access.ObjectType -eq "AB721A53-1E2F-11D0-9819-00AA0040529B"))
                    {
                        If ($access.AccessControlType -eq "Deny")
                        {
                            If ($access.IdentityReference -eq "Everyone")
                            {
                                $DenyEveryone = $true
                            }
                            Elseif ($access.IdentityReference -eq "NT AUTHORITY\SELF")
                            {
                                $DenySelf = $true
                            }
                        }
                    }
                }
                If ($DenyEveryone -and $DenySelf)
                {
                    $CannotChangePassword.Add($($_.properties.samaccountname),$true)
                    Remove-Variable DenyEveryone
                    Remove-Variable DenySelf
                }
                Else
                {
                    $CannotChangePassword.Add($($_.properties.samaccountname),$false)
                }
                Remove-Variable data
                Remove-Variable aclObject
                $icnt ++
            }
            Write-Progress -Activity "Calculating if the user Cannot Change Password" -Completed -Status "All Done"
            $ADFileName  = -join($ReportPath,'\','Users','.csv')
            [ADRecon.LDAPClass]::UserParser($ADUsers, $date, $PassMaxAge, $ADFileName, $CannotChangePassword, $DormantTimeSpan, $Threads, $FlushCount, $OutputType)
            Remove-Variable ADUsers
            Write-Verbose "[+] Domain Users"
        }
    }
}

Function Get-ADRUserSPN
{
<#
.SYNOPSIS
    Returns all user service principal name (SPN) in the current (or specified) domain.

.DESCRIPTION
    Returns all user service principal name (SPN) in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.PARAMETER ReportPath
    [string]
    Path for ADRecon output folder.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.PARAMETER FlushCount
    [int]
    The number of processed objects which will be flushed to disk. Default -1 (After all objects are processed).

.PARAMETER OutputType
    [array]
    Output Type.

.OUTPUTS
    CSV file is created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [bool] $UseAltCreds,

        [Parameter(Mandatory = $true)]
        [string] $ReportPath,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10,

        [Parameter(Mandatory = $false)]
        [int] $FlushCount = -1,

        [Parameter(Mandatory = $true)]
        [array] $OutputType
    )

    Write-Output "[-] Domain User SPNs"
    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADUsers = Get-ADObject -LDAPFilter "(&(!objectClass=computer)(servicePrincipalName=*))" -Properties Name,sAMAccountName,servicePrincipalName,pwdLastSet,Description -ResultPageSize $PageSize
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
            Continue
        }

        If ($ADUsers)
        {
            $ADFileName  = -join($ReportPath,'\','UserSPNs','.csv')
            $SPNCount = [ADRecon.ADWSClass]::UserSPNParser($ADUsers, $ADFileName, $Threads, $FlushCount, $OutputType)
            # Temporary solution for exception in [ADRecon.ADWSClass]::UserSPNParser
            # System.InvalidCastException: Unable to cast object of type 'Microsoft.ActiveDirectory.Management.ADObject' to type 'System.Management.Automation.PSObject'.
            If ($SPNCount -eq 1)
            {
                $UserSPNObj = @()
                $ADUsers | ForEach-Object {
                    For($i=0; $i -lt $_.servicePrincipalName.count; $i++)
                    {
                        $Obj = New-Object PSObject
                        [array] $SPNObjectArray = $_.servicePrincipalName[$i] -Split("/")
                        $Obj | Add-Member -MemberType NoteProperty -Name "Name" -Value $_.Name
                        $Obj | Add-Member -MemberType NoteProperty -Name "Username" -Value $_.sAMAccountName
                        $Obj | Add-Member -MemberType NoteProperty -Name "Service" -Value $SPNObjectArray[0]
                        $Obj | Add-Member -MemberType NoteProperty -Name "Host" -Value $SPNObjectArray[1]
                        If ($null -ne $_.pwdLastSet)
                        {
                            $pwdlastSet = [datetime]::FromFileTime($_.pwdLastSet)
                        }
                        Else
                        {
                            $pwdlastSet = "-"
                        }
                        $Obj | Add-Member -MemberType NoteProperty -Name "Password Last Set" -Value $pwdlastSet
                        $Obj | Add-Member -MemberType NoteProperty -Name "Description" -Value $_.description
                        $UserSPNObj += $Obj
                    }
                }
                If ($UserSPNObj)
                {
                    Switch ($OutputType)
                    {
                        'STDOUT' { $UserSPNObj }
                        'CSV'
                        {
                            Try
                            {
                                $UserSPNObj | Export-Csv -Path $ADFileName -NoTypeInformation
                            }
                            Catch
                            {
                                Write-Output "[ERROR] Failed to Export CSV File"
                                Write-Output "[EXCEPTION] $($_.Exception.Message)"
                            }
                        }
                    }
                }
            }
            Remove-Variable ADUsers
            Write-Verbose "[+] Domain User SPNs"
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(&(!objectClass=computer)(servicePrincipalName=*))"
        $ObjSearcher.PropertiesToLoad.AddRange(("name","samaccountname","serviceprincipalname","pwdlastset","description"))
        $ObjSearcher.SearchScope = "Subtree"
        Try
        {
            $ADUsers = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADUsers)
        {
            $ADFileName  = -join($ReportPath,'\','UserSPNs','.csv')
            [ADRecon.LDAPClass]::UserSPNParser($ADUsers, $ADFileName, $Threads, $FlushCount, $OutputType)
            Remove-Variable ADUsers
            Write-Verbose "[+] Domain User SPNs"
        }
    }
}

Function Get-ADRGroup
{
<#
.SYNOPSIS
    Returns all groups in the current (or specified) domain.

.DESCRIPTION
    Returns all groups in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.PARAMETER ReportPath
    [string]
    Path for ADRecon output folder.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.PARAMETER FlushCount
    [int]
    The number of processed objects which will be flushed to disk. Default -1 (After all objects are processed).

.PARAMETER OutputType
    [array]
    Output Type.

.OUTPUTS
    CSV file is created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [bool] $UseAltCreds,

        [Parameter(Mandatory = $true)]
        [string] $ReportPath,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10,

        [Parameter(Mandatory = $false)]
        [int] $FlushCount = -1,

        [Parameter(Mandatory = $true)]
        [array] $OutputType
    )

    Write-Output "[-] Domain Groups - May take some time"
    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADGroups = Get-ADGroup -Filter * -ResultPageSize $PageSize -Properties CanonicalName,DistinguishedName,Description,SamAccountName,SID,managedBy,whenChanged,whenCreated
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }

        If ($ADGroups)
        {
            $ADFileName  = -join($ReportPath,'\','Groups','.csv')
            [ADRecon.ADWSClass]::GroupParser($ADGroups, $ADFileName, $Threads, $FlushCount, $OutputType)
            Remove-Variable ADGroups
            Write-Verbose "[+] Domain Groups"
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(objectClass=group)"
        $ObjSearcher.PropertiesToLoad.AddRange(("canonicalname", "distinguishedname", "description", "samaccountname", "managedby", "objectsid", "whencreated", "whenchanged"))
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADGroups = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }
        $ObjSearcher.dispose()

        If ($ADGroups)
        {
            $ADFileName  = -join($ReportPath,'\','Groups','.csv')
            [ADRecon.LDAPClass]::GroupParser($ADGroups, $ADFileName, $Threads, $FlushCount, $OutputType)
            Remove-Variable ADGroups
            Write-Verbose "[+] Domain Groups"
        }
    }
}

Function Get-ADRGroupMember
{
<#
.SYNOPSIS
    Returns all groups and their members in the current (or specified) domain.

.DESCRIPTION
    Returns all groups and their members in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.PARAMETER ReportPath
    [string]
    Path for ADRecon output folder.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.PARAMETER FlushCount
    [int]
    The number of processed objects which will be flushed to disk. Default -1 (After all objects are processed).

.PARAMETER OutputType
    [array]
    Output Type.

.OUTPUTS
    CSV file is created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [bool] $UseAltCreds,

        [Parameter(Mandatory = $true)]
        [string] $ReportPath,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10,

        [Parameter(Mandatory = $false)]
        [int] $FlushCount = -1,

        [Parameter(Mandatory = $true)]
        [array] $OutputType
    )

    Write-Output "[-] Domain Group Memberships - May take some time"
    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADGroups = Get-ADObject -LDAPFilter '(memberof=*)' -Properties DistinguishedName,sAMAccountName,memberof,samaccounttype
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }

        If ($ADGroups)
        {
            $ADFileName  = -join($ReportPath,'\','GroupMembers','.csv')
            [ADRecon.ADWSClass]::GroupMemberParser($ADGroups, $ADFileName, $Threads, $FlushCount, $OutputType)
            Remove-Variable ADGroups
            Write-Verbose "[+] Domain Group Memberships"
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(memberof=*)"
        $ObjSearcher.PropertiesToLoad.AddRange(("samaccountname", "distinguishedname", "dnshostname", "samaccounttype", "memberof"))
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADGroups = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }
        $ObjSearcher.dispose()

        If ($ADGroups)
        {
            $ADFileName  = -join($ReportPath,'\','GroupMembers','.csv')
            [ADRecon.LDAPClass]::GroupMemberParser($ADGroups, $ADFileName, $Threads, $FlushCount, $OutputType)
            Remove-Variable ADGroups
            Write-Verbose "[+] Domain Group Memberships"
        }
    }
}

Function Get-ADROU
{
<#
.SYNOPSIS
    Returns all Organizational Units (OU) in the current (or specified) domain.

.DESCRIPTION
    Returns all Organizational Units (OU) in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.PARAMETER ReportPath
    [string]
    Path for ADRecon output folder.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER OutputType
    [array]
    Output Type.

.OUTPUTS
    CSV file is created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [bool] $UseAltCreds,

        [Parameter(Mandatory = $true)]
        [string] $ReportPath,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $true)]
        [array] $OutputType
    )

    Write-Output "[-] Domain OrganizationalUnits"
    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADOUs = Get-ADOrganizationalUnit -Filter * -Properties Created,DistinguishedName,Description,Name,gPLink,gPOptions
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }

        If ($ADOUs)
        {
            Write-Output "[*] Total OUs: $($ADOUs | Measure-Object | Select-Object -ExpandProperty Count)"
            $OUObj = @()
            $ADOUs | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name Name -Value $_.Name
                $Obj | Add-Member -MemberType NoteProperty -Name Created -Value $_.Created
                $Obj | Add-Member -MemberType NoteProperty -Name DistinguishedName -Value $_.DistinguishedName
                $Obj | Add-Member -MemberType NoteProperty -Name Description -Value $_.Description
                $Obj | Add-Member -MemberType NoteProperty -Name gPLink -Value $_.gPLink
                $Obj | Add-Member -MemberType NoteProperty -Name gPOptions -Value $_.gPOptions
                $Obj | Add-Member -MemberType NoteProperty -Name Depth -Value $(($_.DistinguishedName -split 'OU=').Count -1)
                $OUObj += $Obj
            }
            Remove-Variable ADOUs
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(objectCategory=organizationalunit)"
        $ObjSearcher.PropertiesToLoad.AddRange(("whencreated","distinguishedname","description","name","gplink","gpoptions"))
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADOUs = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }
        $ObjSearcher.dispose()

        If ($ADOUs)
        {
            Write-Output "[*] Total OUs: $($ADOUs | Measure-Object | Select-Object -ExpandProperty Count)"
            $OUObj = @()
            $ADOUs | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name Name -Value ([string] $($_.Properties.name))
                $Obj | Add-Member -MemberType NoteProperty -Name Created -Value ([DateTime] $($_.Properties.whencreated))
                $Obj | Add-Member -MemberType NoteProperty -Name DistinguishedName -Value ([string] $($_.Properties.distinguishedname))
                $Obj | Add-Member -MemberType NoteProperty -Name Description -Value ([string] $($_.Properties.description))
                $Obj | Add-Member -MemberType NoteProperty -Name gPLink -Value ([string] $($_.Properties.gplink))
                $Obj | Add-Member -MemberType NoteProperty -Name gPOptions -Value ([string] $($_.Properties.gpoptions))
                $Obj | Add-Member -MemberType NoteProperty -Name Depth -Value $(($_.Properties.distinguishedname -split 'OU=').Count -1)
                $OUObj += $Obj
            }
            Remove-Variable ADOUs
        }
    }

    If ($OUObj)
    {
        Switch ($OutputType)
        {
            'STDOUT' { $OUObj }
            'CSV'
            {
                Write-Verbose "[+] Domain OrganizationalUnits"
                $ADFileName = -join($ReportPath,'\','OUs','.csv')
                Try
                {
                    $OUObj | Export-Csv -Path $ADFileName -NoTypeInformation
                }
                Catch
                {
                    Write-Output "Failed to Export CSV File"
                    Write-Output "[EXCEPTION] $($_.Exception.Message)"
                }
                Remove-Variable OUObj
                Remove-Variable ADFileName
            }
        }
    }
}

Function Get-ADROUPermission
{
<#
.SYNOPSIS
    Returns all Organizational Units (OU) permissions in the current (or specified) domain.

.DESCRIPTION
    Returns all Organizational Units (OU) permissions in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.PARAMETER ReportPath
    [string]
    Path for ADRecon output folder.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER DCIP
    [string]
    IP Address of the Domain Controller.

.PARAMETER creds
    [Management.Automation.PSCredential]
    Credentials.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.OUTPUTS
    CSV file is created in the folder specified with the information.

.PARAMETER OutputType
    [array]
    Output Type.

.LINK
    https://gallery.technet.microsoft.com/Active-Directory-OU-1d09f989
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [bool] $UseAltCreds,

        [Parameter(Mandatory = $true)]
        [string] $ReportPath,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [string] $DCIP,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $creds = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $true)]
        [array] $OutputType
    )

    Write-Output "[-] Domain OrganizationalUnits Permissions - May take some time"
    # based on https://gallery.technet.microsoft.com/Active-Directory-OU-1d09f989
    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            If (-Not $UseAltCreds)
            {
                Set-Location AD:
            }
            $schemaIDGUID = @{}
            $GUIDs = @{'00000000-0000-0000-0000-000000000000' = 'All'}

            $schemaIDs = Get-ADObject -SearchBase (Get-ADRootDSE).schemaNamingContext -LDAPFilter '(schemaIDGUID=*)' -Properties name, schemaIDGUID

            $schemaIDs | Where-Object {$_} | ForEach-Object {
                # convert the GUID
                $GUIDs[(New-Object Guid (,$_.schemaIDGUID)).Guid] = $_.name
            }
            Remove-Variable schemaIDs

            $schemaIDs = Get-ADObject -SearchBase "CN=Extended-Rights,$((Get-ADRootDSE).configurationNamingContext)" -LDAPFilter '(objectClass=controlAccessRight)' -Properties name, rightsGUID

            $schemaIDs | Where-Object {$_} | ForEach-Object {
                # convert the GUID
                $GUIDs[(New-Object Guid (,$_.rightsGUID)).Guid] = $_.name
            }
            Remove-Variable schemaIDs

            # Get a list of all OUs.  Add in the root containers for good measure (users, computers, etc.).
            $OUs  = @(Get-ADDomain | Select-Object -ExpandProperty DistinguishedName)
            $OUs += Get-ADOrganizationalUnit -Filter * | Select-Object -ExpandProperty DistinguishedName
            $OUs += Get-ADObject -SearchBase (Get-ADDomain).DistinguishedName -SearchScope OneLevel -LDAPFilter '(objectClass=container)' | Select-Object -ExpandProperty DistinguishedName
            ForEach ($OU in $OUs)
            {
                $OUPermissions += Get-Acl -Path "$OU" |
                Select-Object -ExpandProperty Access |
                Select-Object @{name='organizationalUnit';expression={$OU}}, `
                       @{name='objectTypeName';expression={$GUIDs[$_.objectType.ToString()]}}, `
                       @{name='inheritedObjectTypeName';expression={$GUIDs[$_.inheritedObjectType.ToString()]}}, `
                       *
            }
            Remove-Variable OUs
            Remove-Variable GUIDs
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(objectCategory=organizationalunit)"
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADOUs = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }
        $ObjSearcher.dispose()

        $OUPermissions = @()
        If ($ADOUs)
        {
            $GUIDs = @{'00000000-0000-0000-0000-000000000000' = 'All'}

        If ($UseAltCreds)
        {
            $DomainFQDN = Get-DNtoFQDN($objDomain.distinguishedName)
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain",$($DomainFQDN),$($creds.UserName),$($creds.GetNetworkCredential().password))
            Try
            {
                $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            Catch
            {
                Write-Output "[EXCEPTION] $($_.Exception.Message)"
            }

            $ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest",$($ADDomain.Forest),$($creds.UserName),$($creds.GetNetworkCredential().password))
            Try
            {
                $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestContext)
                $SchemaPath = $ADForest.Schema.Name
            }
            Catch
            {
                Write-Output "[EXCEPTION] $($_.Exception.Message)"
            }
        }
        Else
        {
            $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
            $SchemaPath = $ADForest.Schema.Name
            Remove-Variable SchemaPath
        }

            If ($SchemaPath)
            {
                If ($UseAltCreds)
                {
                    $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DCIP)/$($SchemaPath)", $creds.UserName,$creds.GetNetworkCredential().Password
                    $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                }
                Else
                {
                    $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher ([ADSI] "LDAP://$($SchemaPath)")
                }
                $objSearcherPath.PageSize = $PageSize
                $objSearcherPath.filter = "(schemaIDGUID=*)"

                Try
                {
                    $SchemaSearcher = $objSearcherPath.FindAll()
                }
                Catch
                {
                    Write-Output "[EXCEPTION] $($_.Exception.Message)"
                }

                If ($SchemaSearcher)
                {
                    $SchemaSearcher | Where-Object {$_} | ForEach-Object {
                        # convert the GUID
                        $GUIDs[(New-Object Guid (,$_.properties.schemaidguid[0])).Guid] = $_.properties.name[0]
                    }
                    $SchemaSearcher.dispose()
                }
                $objSearcherPath.dispose()

                If ($UseAltCreds)
                {
                    $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DCIP)/$($SchemaPath.replace("Schema","Extended-Rights"))", $creds.UserName,$creds.GetNetworkCredential().Password
                    $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                }
                Else
                {
                    $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher ([ADSI] "LDAP://$($SchemaPath.replace("Schema","Extended-Rights"))")
                }
                $objSearcherPath.PageSize = $PageSize
                $objSearcherPath.filter = "(objectClass=controlAccessRight)"

                Try
                {
                    $RightsSearcher = $objSearcherPath.FindAll()
                }
                Catch
                {
                    Write-Output "[EXCEPTION] $($_.Exception.Message)"
                }

                If ($RightsSearcher)
                {
                    $RightsSearcher | Where-Object {$_} | ForEach-Object {
                        # convert the GUID
                        $GUIDs[$_.properties.rightsguid[0].toString()] = $_.properties.name[0]
                    }
                    $RightsSearcher.dispose()
                }
                $objSearcherPath.dispose()
            }
            If ($UseAltCreds)
            {
                ForEach ($OU in $ADOUs)
                {
                    $OUPermissions += (New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DCIP)/$($OU.Properties.distinguishedname)", $creds.UserName,$creds.GetNetworkCredential().Password).PsBase.ObjectSecurity.access | Select-Object @{name='organizationalUnit';expression={$OU.properties.distinguishedname}}, `
                       @{name='objectTypeName';expression={$GUIDs[$_.objectType.ToString()]}}, `
                       @{name='inheritedObjectTypeName';expression={$GUIDs[$_.inheritedObjectType.ToString()]}}, `
                       *
                }
            }
            Else
            {
                ForEach ($OU in $ADOUs)
                {
                $OUPermissions += (($OU.GetDirectoryEntry()).Get_ObjectSecurity()).Access | Select-Object @{name='organizationalUnit';expression={$OU.properties.distinguishedname}}, `
                       @{name='objectTypeName';expression={$GUIDs[$_.objectType.ToString()]}}, `
                       @{name='inheritedObjectTypeName';expression={$GUIDs[$_.inheritedObjectType.ToString()]}}, `
                       *
                }
            }
            Remove-Variable GUIDs
            Remove-Variable ADOUs
        }
    }

    If ($OUPermissions)
    {
        Switch ($OutputType)
        {
            'STDOUT' { $OUPermissions }
            'CSV'
            {
                Write-Verbose "[+] Domain OrganizationalUnits Permissions"
                $ADFileName = -join($ReportPath,'\','OUPermissions','.csv')
                Try
                {
                    $OUPermissions | Export-Csv -Path $ADFileName -NoTypeInformation
                }
                Catch
                {
                    Write-Output "Failed to Export CSV File"
                    Write-Output "[EXCEPTION] $($_.Exception.Message)"
                }
                Remove-Variable OUPermissions
                Remove-Variable ADFileName
            }
        }
    }
}

Function Get-ADRGPO
{
<#
.SYNOPSIS
    Returns all Group Policy Objects (GPO) in the current (or specified) domain.

.DESCRIPTION
    Returns all Group Policy Objects (GPO) in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.PARAMETER ReportPath
    [string]
    Path for ADRecon output folder.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER creds
    [Management.Automation.PSCredential]
    Credentials.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER OutputType
    [array]
    Output Type.

.OUTPUTS
    CSV file is created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [bool] $UseAltCreds,

        [Parameter(Mandatory = $true)]
        [string] $ReportPath,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $true)]
        [array] $OutputType
    )

    Write-Output "[-] Domain GPOs"
    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADDomainGPOs = Get-ADObject -LDAPFilter '(objectCategory=groupPolicyContainer)' -Properties DisplayName,whenCreated,whenChanged,Name,gPCFileSysPath
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }

        If ($ADDomainGPOs)
        {
            Write-Output "[*] Total GPOs: $($ADDomainGPOs | Measure-Object | Select-Object -ExpandProperty Count)"
            $ADDomainGPOObj = @()
            $ADDomainGPOs | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $_.DisplayName
                $Obj | Add-Member -MemberType NoteProperty -Name Created -Value $_.whenCreated
                $Obj | Add-Member -MemberType NoteProperty -Name Changed -Value $_.whenChanged
                $Obj | Add-Member -MemberType NoteProperty -Name Name -Value $([ADRecon.ADWSClass]::CleanString($_.Name))
                $Obj | Add-Member -MemberType NoteProperty -Name FilePath -Value $_.gPCFileSysPath
                $ADDomainGPOObj += $Obj
            }
            Remove-Variable ADDomainGPOs
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(objectCategory=groupPolicyContainer)"
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADDomainGPOs = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }
        $ObjSearcher.dispose()

        If ($ADDomainGPOs)
        {
            Write-Output "[*] Total GPOs: $($ADDomainGPOs | Measure-Object | Select-Object -ExpandProperty Count)"
            $ADDomainGPOObj = @()
            $ADDomainGPOs | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name DisplayName -Value ([string] $($_.Properties.displayname))
                $Obj | Add-Member -MemberType NoteProperty -Name Created -Value ([DateTime] $($_.Properties.whencreated))
                $Obj | Add-Member -MemberType NoteProperty -Name Changed -Value ([DateTime] $($_.Properties.whenchanged))
                $Obj | Add-Member -MemberType NoteProperty -Name Name -Value $([ADRecon.LDAPClass]::CleanString($_.Properties.name))
                $Obj | Add-Member -MemberType NoteProperty -Name FilePath -Value ([string] $($_.Properties.gpcfilesyspath))
                $ADDomainGPOObj += $Obj
            }
            Remove-Variable ADDomainGPOs
        }
    }

    If ($ADDomainGPOObj)
    {
        Switch ($OutputType)
        {
            'STDOUT' { $ADDomainGPOObj }
            'CSV'
            {
                Write-Verbose "[+] Domain GPOs"
                $ADFileName = -join($ReportPath,'\','GPOs','.csv')
                Try
                {
                    $ADDomainGPOObj | Export-Csv -Path $ADFileName -NoTypeInformation
                }
                Catch
                {
                    Write-Output "Failed to Export CSV File"
                    Write-Output "[EXCEPTION] $($_.Exception.Message)"
                }
                Remove-Variable ADDomainGPOObj
                Remove-Variable ADFileName
            }
        }
    }
}

# Modified Convert-DNSRecord function from https://github.com/PowerShellMafia/PowerSploit/blob/dev/Recon/PowerView.ps1
Function Convert-DNSRecord
{
<#
.SYNOPSIS

Helpers that decodes a binary DNS record blob.

Author: Michael B. Smith, Will Schroeder (@harmj0y)
License: BSD 3-Clause
Required Dependencies: None

.DESCRIPTION

Decodes a binary blob representing an Active Directory DNS entry.
Used by Get-DomainDNSRecord.

Adapted/ported from Michael B. Smith's code at https://raw.githubusercontent.com/mmessano/PowerShell/master/dns-dump.ps1

.PARAMETER DNSRecord

A byte array representing the DNS record.

.OUTPUTS

System.Management.Automation.PSCustomObject

Outputs custom PSObjects with detailed information about the DNS record entry.

.LINK

https://raw.githubusercontent.com/mmessano/PowerShell/master/dns-dump.ps1
#>

    [OutputType('System.Management.Automation.PSCustomObject')]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $True, ValueFromPipelineByPropertyName = $True)]
        [Byte[]]
        $DNSRecord
    )

    BEGIN {
        Function Get-Name
        {
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseOutputTypeCorrectly', '')]
            [CmdletBinding()]
            Param(
                [Byte[]]
                $Raw
            )

            [Int]$Length = $Raw[0]
            [Int]$Segments = $Raw[1]
            [Int]$Index =  2
            [String]$Name  = ''

            while ($Segments-- -gt 0)
            {
                [Int]$SegmentLength = $Raw[$Index++]
                while ($SegmentLength-- -gt 0)
                {
                    $Name += [Char]$Raw[$Index++]
                }
                $Name += "."
            }
            $Name
        }
    }

    PROCESS
    {
        # $RDataLen = [BitConverter]::ToUInt16($DNSRecord, 0)
        $RDataType = [BitConverter]::ToUInt16($DNSRecord, 2)
        $UpdatedAtSerial = [BitConverter]::ToUInt32($DNSRecord, 8)

        $TTLRaw = $DNSRecord[12..15]

        # reverse for big endian
        $Null = [array]::Reverse($TTLRaw)
        $TTL = [BitConverter]::ToUInt32($TTLRaw, 0)

        $Age = [BitConverter]::ToUInt32($DNSRecord, 20)
        If ($Age -ne 0)
        {
            $TimeStamp = ((Get-Date -Year 1601 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0).AddHours($age)).ToString()
        }
        Else
        {
            $TimeStamp = '[static]'
        }

        $DNSRecordObject = New-Object PSObject

        switch ($RDataType)
        {
            1
            {
                $IP = "{0}.{1}.{2}.{3}" -f $DNSRecord[24], $DNSRecord[25], $DNSRecord[26], $DNSRecord[27]
                $Data = $IP
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'A'
            }

            2
            {
                $NSName = Get-Name $DNSRecord[24..$DNSRecord.length]
                $Data = $NSName
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'NS'
            }

            5
            {
                $Alias = Get-Name $DNSRecord[24..$DNSRecord.length]
                $Data = $Alias
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'CNAME'
            }

            6
            {
                $PrimaryNS = Get-Name $DNSRecord[44..$DNSRecord.length]
                $ResponsibleParty = Get-Name $DNSRecord[$(46+$DNSRecord[44])..$DNSRecord.length]
                $SerialRaw = $DNSRecord[24..27]
                # reverse for big endian
                $Null = [array]::Reverse($SerialRaw)
                $Serial = [BitConverter]::ToUInt32($SerialRaw, 0)

                $RefreshRaw = $DNSRecord[28..31]
                $Null = [array]::Reverse($RefreshRaw)
                $Refresh = [BitConverter]::ToUInt32($RefreshRaw, 0)

                $RetryRaw = $DNSRecord[32..35]
                $Null = [array]::Reverse($RetryRaw)
                $Retry = [BitConverter]::ToUInt32($RetryRaw, 0)

                $ExpiresRaw = $DNSRecord[36..39]
                $Null = [array]::Reverse($ExpiresRaw)
                $Expires = [BitConverter]::ToUInt32($ExpiresRaw, 0)

                $MinTTLRaw = $DNSRecord[40..43]
                $Null = [array]::Reverse($MinTTLRaw)
                $MinTTL = [BitConverter]::ToUInt32($MinTTLRaw, 0)

                $Data = "[" + $Serial + "][" + $PrimaryNS + "][" + $ResponsibleParty + "][" + $Refresh + "][" + $Retry + "][" + $Expires + "][" + $MinTTL + "]"
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'SOA'
            }

            12
            {
                $Ptr = Get-Name $DNSRecord[24..$DNSRecord.length]
                $Data = $Ptr
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'PTR'
            }

            13
            {
                [string]$CPUType = ""
                [string]$OSType  = ""
                [int]$SegmentLength = $DNSRecord[24]
                $Index = 25
                while ($SegmentLength-- -gt 0)
                {
                    $CPUType += [char]$DNSRecord[$Index++]
                }
                $Index = 24 + $DNSRecord[24] + 1
                [int]$SegmentLength = $Index++
                while ($SegmentLength-- -gt 0)
                {
                    $OSType += [char]$DNSRecord[$Index++]
                }
                $Data = "[" + $CPUType + "][" + $OSType + "]"
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'HINFO'
            }

            15
            {
                $PriorityRaw = $DNSRecord[24..25]
                # reverse for big endian
                $Null = [array]::Reverse($PriorityRaw)
                $Priority = [BitConverter]::ToUInt16($PriorityRaw, 0)
                $MXHost   = Get-Name $DNSRecord[26..$DNSRecord.length]
                $Data = "[" + $Priority + "][" + $MXHost + "]"
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'MX'
            }

            16
            {
                [string]$TXT  = ''
                [int]$SegmentLength = $DNSRecord[24]
                $Index = 25
                while ($SegmentLength-- -gt 0)
                {
                    $TXT += [char]$DNSRecord[$Index++]
                }
                $Data = $TXT
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'TXT'
            }

            28
            {
        		### yeah, this doesn't do all the fancy formatting that can be done for IPv6
                $AAAA = ""
                for ($i = 24; $i -lt 40; $i+=2)
                {
                    $BlockRaw = $DNSRecord[$i..$($i+1)]
                    # reverse for big endian
                    $Null = [array]::Reverse($BlockRaw)
                    $Block = [BitConverter]::ToUInt16($BlockRaw, 0)
			        $AAAA += ($Block).ToString('x4')
			        If ($i -ne 38)
                    {
                        $AAAA += ':'
                    }
                }
                $Data = $AAAA
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'AAAA'
            }

            33
            {
                $PriorityRaw = $DNSRecord[24..25]
                # reverse for big endian
                $Null = [array]::Reverse($PriorityRaw)
                $Priority = [BitConverter]::ToUInt16($PriorityRaw, 0)

                $WeightRaw = $DNSRecord[26..27]
                $Null = [array]::Reverse($WeightRaw)
                $Weight = [BitConverter]::ToUInt16($WeightRaw, 0)

                $PortRaw = $DNSRecord[28..29]
                $Null = [array]::Reverse($PortRaw)
                $Port = [BitConverter]::ToUInt16($PortRaw, 0)

                $SRVHost = Get-Name $DNSRecord[30..$DNSRecord.length]
                $Data = "[" + $Priority + "][" + $Weight + "][" + $Port + "][" + $SRVHost + "]"
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'SRV'
            }

            default
            {
                $Data = $([System.Convert]::ToBase64String($DNSRecord[24..$DNSRecord.length]))
                $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'UNKNOWN'
            }
        }
        $DNSRecordObject | Add-Member Noteproperty 'UpdatedAtSerial' $UpdatedAtSerial
        $DNSRecordObject | Add-Member Noteproperty 'TTL' $TTL
        $DNSRecordObject | Add-Member Noteproperty 'Age' $Age
        $DNSRecordObject | Add-Member Noteproperty 'TimeStamp' $TimeStamp
        $DNSRecordObject | Add-Member Noteproperty 'Data' $Data
        Return $DNSRecordObject
    }
}

Function Get-ADRDNSZone
{
<#
.SYNOPSIS
    Returns all DNS Zones and Records in the current (or specified) domain.

.DESCRIPTION
    Returns all DNS Zones and Records in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.PARAMETER ReportPath
    [string]
    Path for ADRecon output folder.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER DCIP
    [string]
    IP Address of the Domain Controller.

.PARAMETER creds
    [Management.Automation.PSCredential]
    Credentials.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER OutputType
    [array]
    Output Type.

.OUTPUTS
    CSV files are created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [bool] $UseAltCreds,

        [Parameter(Mandatory = $true)]
        [string] $ReportPath,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [string] $DCIP,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $creds = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $true)]
        [array] $OutputType
    )

    Write-Output "[-] Domain DNS Zones and Records"
    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADDNSZones = Get-ADObject -LDAPFilter '(objectClass=dnsZone)' -Properties Name,whenCreated,whenChanged,usncreated,usnchanged,distinguishedname
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }

        $DNSZoneArray = @()
        If ($ADDNSZones)
        {
            $DNSZoneArray += $ADDNSZones
            Remove-Variable ADDNSZones
        }

        Try
        {
            $ADDNSZones1 = Get-ADObject -LDAPFilter '(objectClass=dnsZone)' -SearchBase "CN=MicrosoftDNS,DC=DomainDnsZones,$((Get-ADDomain).DistinguishedName)" -Properties Name,whenCreated,whenChanged,usncreated,usnchanged,distinguishedname
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }
        If ($ADDNSZones1)
        {
            $DNSZoneArray += $ADDNSZones1
            Remove-Variable ADDNSZones1
        }

        Try
        {
            $ADDNSZones2 = Get-ADObject -LDAPFilter '(objectClass=dnsZone)' -SearchBase "CN=MicrosoftDNS,DC=ForestDnsZones,$((Get-ADDomain).DistinguishedName)" -Properties Name,whenCreated,whenChanged,usncreated,usnchanged,distinguishedname
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }
        If ($ADDNSZones2)
        {
            $DNSZoneArray += $ADDNSZones2
            Remove-Variable ADDNSZones2
        }

        Write-Output "[*] Total DNS Zones: $($DNSZoneArray | Measure-Object | Select-Object -ExpandProperty Count)"

        If ($DNSZoneArray)
        {
            $ADDNSZonesObj = @()
            $ADDNSNodesObj = @()
            $DNSZoneArray | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name Name -Value $([ADRecon.ADWSClass]::CleanString($_.Name))
                $Obj | Add-Member -MemberType NoteProperty -Name USNCreated -Value $_.usncreated
                $Obj | Add-Member -MemberType NoteProperty -Name USNChanged -Value $_.usnchanged
                $Obj | Add-Member -MemberType NoteProperty -Name Created -Value $_.whenCreated
                $Obj | Add-Member -MemberType NoteProperty -Name Changed -Value $_.whenChanged
                Try
                {
                    $DNSNodes = Get-ADObject -SearchBase $($_.DistinguishedName) -LDAPFilter '(objectClass=dnsNode)' -Properties DistinguishedName,dnsrecord,dNSTombstoned,Name,ProtectedFromAccidentalDeletion,showInAdvancedViewOnly,whenChanged,whenCreated
                }
                Catch
                {
                    Write-Output "[EXCEPTION] $($_.Exception.Message)"
                }
                If ($DNSNodes)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name RecordCount -Value $($DNSNodes | Measure-Object | Select-Object -ExpandProperty Count)
                    $DNSNodes | ForEach-Object {
                        $ObjNode = New-Object PSObject
                        $ObjNode | Add-Member -MemberType NoteProperty -Name ZoneName -Value $Obj.Name
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Name -Value $_.Name
                        Try
                        {
                            $DNSRecord = Convert-DNSRecord $_.dnsrecord[0]
                        }
                        Catch
                        {
                            Write-Output "[EXCEPTION] $($_.Exception.Message)"
                        }
                        $ObjNode | Add-Member -MemberType NoteProperty -Name RecordType -Value $DNSRecord.RecordType
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Data -Value $DNSRecord.Data
                        $ObjNode | Add-Member -MemberType NoteProperty -Name TTL -Value $DNSRecord.TTL
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Age -Value $DNSRecord.Age
                        $ObjNode | Add-Member -MemberType NoteProperty -Name TimeStamp -Value $DNSRecord.TimeStamp
                        $ObjNode | Add-Member -MemberType NoteProperty -Name UpdatedAtSerial -Value $DNSRecord.UpdatedAtSerial
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Created -Value $_.whenCreated
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Changed -Value $_.whenChanged
                        # TO DO LDAP part
                        #$ObjNode | Add-Member -MemberType NoteProperty -Name dNSTombstoned -Value $_.dNSTombstoned
                        #$ObjNode | Add-Member -MemberType NoteProperty -Name ProtectedFromAccidentalDeletion -Value $_.ProtectedFromAccidentalDeletion
                        $ObjNode | Add-Member -MemberType NoteProperty -Name showInAdvancedViewOnly -Value $_.showInAdvancedViewOnly
                        $ObjNode | Add-Member -MemberType NoteProperty -Name DistinguishedName -Value $_.DistinguishedName
                        $ADDNSNodesObj += $ObjNode
                        If ($DNSRecord)
                        {
                            Remove-Variable DNSRecord
                        }
                    }
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name RecordCount -Value $null
                }
                $Obj | Add-Member -MemberType NoteProperty -Name DistinguishedName -Value $_.DistinguishedName
                $ADDNSZonesObj += $Obj
            }
            Remove-Variable DNSZoneArray
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.PropertiesToLoad.AddRange(("name","whencreated","whenchanged","usncreated","usnchanged","distinguishedname"))
        $ObjSearcher.Filter = "(objectClass=dnsZone)"
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADDNSZones = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }
        $ObjSearcher.dispose()

        $DNSZoneArray = @()
        If ($ADDNSZones)
        {
            $DNSZoneArray += $ADDNSZones
            Remove-Variable ADDNSZones
        }

        $SearchPath = "CN=MicrosoftDNS,DC=DomainDnsZones"
        If ($UseAltCreds)
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DCIP)/$($SearchPath),$($objDomain.distinguishedName)", $creds.UserName,$creds.GetNetworkCredential().Password
        }
        Else
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($SearchPath),$($objDomain.distinguishedName)"
        }
        $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
        $objSearcherPath.Filter = "(objectClass=dnsZone)"
        $objSearcherPath.PageSize = $PageSize
        $objSearcherPath.PropertiesToLoad.AddRange(("name","whencreated","whenchanged","usncreated","usnchanged","distinguishedname"))
        $objSearcherPath.SearchScope = "Subtree"

        Try
        {
            $ADDNSZones1 = $objSearcherPath.FindAll()
        }
        Catch
        {
            Write-Warning "Try running with a Privileged Account"
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }
        $objSearcherPath.dispose()

        If ($ADDNSZones1)
        {
            $DNSZoneArray += $ADDNSZones1
            Remove-Variable ADDNSZones1
        }

        $SearchPath = "CN=MicrosoftDNS,DC=ForestDnsZones"
        If ($UseAltCreds)
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DCIP)/$($SearchPath),$($objDomain.distinguishedName)", $creds.UserName,$creds.GetNetworkCredential().Password
        }
        Else
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($SearchPath),$($objDomain.distinguishedName)"
        }
        $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
        $objSearcherPath.Filter = "(objectClass=dnsZone)"
        $objSearcherPath.PageSize = $PageSize
        $objSearcherPath.PropertiesToLoad.AddRange(("name","whencreated","whenchanged","usncreated","usnchanged","distinguishedname"))
        $objSearcherPath.SearchScope = "Subtree"

        Try
        {
            $ADDNSZones2 = $objSearcherPath.FindAll()
        }
        Catch
        {
            Write-Warning "Try running with a Privileged Account"
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }
        $objSearcherPath.dispose()

        If ($ADDNSZones2)
        {
            $DNSZoneArray += $ADDNSZones2
            Remove-Variable ADDNSZones2
        }

        Write-Output "[*] Total DNS Zones: $($DNSZoneArray | Measure-Object | Select-Object -ExpandProperty Count)"

        If ($DNSZoneArray)
        {
            $ADDNSZonesObj = @()
            $ADDNSNodesObj = @()
            $DNSZoneArray | ForEach-Object {
                If ($UseAltCreds)
                {
                    $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DCIP)/$($_.Properties.distinguishedname)", $creds.UserName,$creds.GetNetworkCredential().Password
                }
                Else
                {
                    $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($_.Properties.distinguishedname)"
                }
                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                $objSearcherPath.Filter = "(objectClass=dnsNode)"
                $objSearcherPath.PageSize = $PageSize
                $objSearcherPath.PropertiesToLoad.AddRange(("distinguishedname","dnsrecord","name","dc","showinadvancedviewonly","whenchanged","whencreated"))
                Try
                {
                    $DNSNodes = $objSearcherPath.FindAll()
                }
                Catch
                {
                    Write-Output "[EXCEPTION] $($_.Exception.Message)"
                }
                $objSearcherPath.dispose()
                Remove-Variable objSearchPath

                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name Name -Value $([ADRecon.LDAPClass]::CleanString($_.Properties.name))
                $Obj | Add-Member -MemberType NoteProperty -Name USNCreated -Value ([string] $($_.Properties.usncreated))
                $Obj | Add-Member -MemberType NoteProperty -Name USNChanged -Value ([string] $($_.Properties.usnchanged))
                $Obj | Add-Member -MemberType NoteProperty -Name Created -Value ([DateTime] $($_.Properties.whencreated))
                $Obj | Add-Member -MemberType NoteProperty -Name Changed -Value ([DateTime] $($_.Properties.whenchanged))
                If ($DNSNodes)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name RecordCount -Value $($DNSNodes | Measure-Object | Select-Object -ExpandProperty Count)
                    $DNSNodes | ForEach-Object {
                        $ObjNode = New-Object PSObject
                        $ObjNode | Add-Member -MemberType NoteProperty -Name ZoneName -Value $Obj.Name
                        $name = ([string] $($_.Properties.name))
                        If (-Not $name)
                        {
                            $name = ([string] $($_.Properties.dc))
                        }
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Name -Value $name
                        Try
                        {
                            $DNSRecord = Convert-DNSRecord $_.Properties.dnsrecord[0]
                        }
                        Catch
                        {
                            Write-Output "[EXCEPTION] $($_.Exception.Message)"
                        }
                        $ObjNode | Add-Member -MemberType NoteProperty -Name RecordType -Value $DNSRecord.RecordType
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Data -Value $DNSRecord.Data
                        $ObjNode | Add-Member -MemberType NoteProperty -Name TTL -Value $DNSRecord.TTL
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Age -Value $DNSRecord.Age
                        $ObjNode | Add-Member -MemberType NoteProperty -Name TimeStamp -Value $DNSRecord.TimeStamp
                        $ObjNode | Add-Member -MemberType NoteProperty -Name UpdatedAtSerial -Value $DNSRecord.UpdatedAtSerial
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Created -Value ([DateTime] $($_.Properties.whencreated))
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Changed -Value ([DateTime] $($_.Properties.whenchanged))
                        # TO DO
                        #$ObjNode | Add-Member -MemberType NoteProperty -Name dNSTombstoned -Value $null
                        #$ObjNode | Add-Member -MemberType NoteProperty -Name ProtectedFromAccidentalDeletion -Value $null
                        $ObjNode | Add-Member -MemberType NoteProperty -Name showInAdvancedViewOnly -Value ([string] $($_.Properties.showinadvancedviewonly))
                        $ObjNode | Add-Member -MemberType NoteProperty -Name DistinguishedName -Value ([string] $($_.Properties.distinguishedname))
                        $ADDNSNodesObj += $ObjNode
                        If ($DNSRecord)
                        {
                            Remove-Variable DNSRecord
                        }
                    }
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name RecordCount -Value $null
                }
                $Obj | Add-Member -MemberType NoteProperty -Name DistinguishedName -Value ([string] $($_.Properties.distinguishedname))
                $ADDNSZonesObj += $Obj
            }
            Remove-Variable DNSZoneArray
        }
    }

    If ($ADDNSZonesObj)
    {
        Switch ($OutputType)
        {
            'STDOUT' { $ADDNSZonesObj }
            'CSV'
            {
                Write-Verbose "[+] Domain DNS Zones"
                $ADFileName = -join($ReportPath,'\','DNSZones','.csv')
                Try
                {
                    $ADDNSZonesObj | Export-Csv -Path $ADFileName -NoTypeInformation
                }
                Catch
                {
                    Write-Output "Failed to Export CSV File"
                    Write-Output "[EXCEPTION] $($_.Exception.Message)"
                }
                Remove-Variable ADDNSZonesObj
                Remove-Variable ADFileName
            }
        }
    }

    If ($ADDNSNodesObj)
    {
        Switch ($OutputType)
        {
            'STDOUT' { $ADDNSNodesObj }
            'CSV'
            {
                Write-Verbose "[+] Domain DNS Records"
                $ADFileName = -join($ReportPath,'\','DNSNodes','.csv')
                Try
                {
                    $ADDNSNodesObj | Export-Csv -Path $ADFileName -NoTypeInformation
                }
                Catch
                {
                    Write-Output "Failed to Export CSV File"
                    Write-Output "[EXCEPTION] $($_.Exception.Message)"
                }
                Remove-Variable ADDNSNodesObj
                Remove-Variable ADFileName
            }
        }
    }
}

Function Get-ADRPrinter
{
<#
.SYNOPSIS
    Returns all printers in the current (or specified) domain.

.DESCRIPTION
    Returns all printers in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.PARAMETER ReportPath
    [string]
    Path for ADRecon output folder.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER OutputType
    [array]
    Output Type.

.OUTPUTS
    CSV file is created in the folder specified with the information.
#>

    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [bool] $UseAltCreds,

        [Parameter(Mandatory = $true)]
        [string] $ReportPath,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $true)]
        [array] $OutputType
    )

    Write-Output "[-] Domain Printers"
    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADPrinters = Get-ADObject -LDAPFilter '(objectCategory=printQueue)' -Properties serverName,printShareName,driverName,driverVersion,portName,url,whenCreated,whenChanged,Name
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }

        If ($ADPrinters)
        {
            $cnt = $($ADPrinters | Measure-Object | Select-Object -ExpandProperty Count)
            If ($cnt -ge 1)
            {
                Write-Output "[*] Total Printers: $cnt"
                $ADPrintersObj = @()
                $ADPrinters | ForEach-Object {
                    # Create the object for each instance.
                    $Obj = New-Object PSObject
                    $Obj | Add-Member -MemberType NoteProperty -Name Name -Value $_.Name
                    $Obj | Add-Member -MemberType NoteProperty -Name ServerName -Value $_.serverName
                    $Obj | Add-Member -MemberType NoteProperty -Name ShareName -Value ([string]($_.printShareName))
                    $Obj | Add-Member -MemberType NoteProperty -Name DriverName -Value $_.driverName
                    $Obj | Add-Member -MemberType NoteProperty -Name DriverVersion -Value $_.driverVersion
                    $Obj | Add-Member -MemberType NoteProperty -Name PortName -Value ([string]($_.portName))
                    $Obj | Add-Member -MemberType NoteProperty -Name URL -Value ([string]($_.url))
                    $Obj | Add-Member -MemberType NoteProperty -Name Created -Value $_.whenCreated
                    $Obj | Add-Member -MemberType NoteProperty -Name Changed -Value $_.whenChanged
                    $ADPrintersObj += $Obj
                }
            }
            Remove-Variable ADPrinters
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(objectCategory=printQueue)"
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADPrinters = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }
        $ObjSearcher.dispose()

        If ($ADPrinters)
        {
            $cnt = $($ADPrinters | Measure-Object | Select-Object -ExpandProperty Count)
            If ($cnt -ge 1)
            {
                Write-Output "[*] Total Printers: $cnt"
                $ADPrintersObj = @()
                $ADPrinters | ForEach-Object {
                    # Create the object for each instance.
                    $Obj = New-Object PSObject
                    $Obj | Add-Member -MemberType NoteProperty -Name Name -Value ([string] $($_.Properties.name))
                    $Obj | Add-Member -MemberType NoteProperty -Name ServerName -Value ([string] $($_.Properties.servername))
                    $Obj | Add-Member -MemberType NoteProperty -Name ShareName -Value ([string] $($_.Properties.printsharename))
                    $Obj | Add-Member -MemberType NoteProperty -Name DriverName -Value ([string] $($_.Properties.drivername))
                    $Obj | Add-Member -MemberType NoteProperty -Name DriverVersion -Value ([string] $($_.Properties.driverversion))
                    $Obj | Add-Member -MemberType NoteProperty -Name PortName -Value ([string] $($_.Properties.portname))
                    $Obj | Add-Member -MemberType NoteProperty -Name URL -Value ([string] $($_.Properties.url))
                    $Obj | Add-Member -MemberType NoteProperty -Name Created -Value ([DateTime] $($_.Properties.whencreated))
                    $Obj | Add-Member -MemberType NoteProperty -Name Changed -Value ([DateTime] $($_.Properties.whenchanged))
                    $ADPrintersObj += $Obj
                }
            }
            Remove-Variable ADPrinters
        }
    }


    If ($ADPrintersObj)
    {
        Switch ($OutputType)
        {
            'STDOUT' { $ADPrintersObj }
            'CSV'
            {
                Write-Verbose "[+] Domain Printers"
                $ADFileName = -join($ReportPath,'\','Printers','.csv')
                Try
                {
                    $ADPrintersObj | Export-Csv -Path $ADFileName -NoTypeInformation
                }
                Catch
                {
                    Write-Output "Failed to Export CSV File"
                    Write-Output "[EXCEPTION] $($_.Exception.Message)"
                }
                Remove-Variable ADPrintersObj
                Remove-Variable ADFileName
            }
        }
    }
}

Function Get-ADRComputer
{
<#
.SYNOPSIS
    Returns all computers in the current (or specified) domain.

.DESCRIPTION
    Returns all computers in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.PARAMETER ReportPath
    [string]
    Path for ADRecon output folder.

.PARAMETER date
    [DateTime]
    Date when ADRecon was executed.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.PARAMETER FlushCount
    [int]
    The number of processed objects which will be flushed to disk. Default -1 (After all objects are processed).

.PARAMETER OutputType
    [array]
    Output Type.

.OUTPUTS
    CSV file is created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [bool] $UseAltCreds,

        [Parameter(Mandatory = $true)]
        [string] $ReportPath,

        [Parameter(Mandatory = $true)]
        [DateTime] $date,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10,

        [Parameter(Mandatory = $false)]
        [int] $FlushCount = -1,

        [Parameter(Mandatory = $true)]
        [array] $OutputType
    )

    Write-Output "[-] Domain Computers - May take some time"
    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADComputers = Get-ADComputer -Filter * -ResultPageSize $PageSize -Properties Name,DNSHostName,Description,Enabled,IPv4Address,OperatingSystem,LastLogonDate,PasswordLastSet,primaryGroupID,TrustedForDelegation,TrustedToAuthForDelegation,SamAccountName,whenChanged,whenCreated,DistinguishedName
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }

        If ($ADComputers)
        {
            $ADFileName  = -join($ReportPath,'\','Computers','.csv')
            $ComputerCount = [ADRecon.ADWSClass]::ComputerParser($ADComputers, $date, $ADFileName, $Threads, $FlushCount, $OutputType)
            # Temporary solution for exception in [ADRecon.ADWSClass]::ComputerParser
            # System.InvalidCastException: Unable to cast object of type 'Microsoft.ActiveDirectory.Management.ADComputer' to type 'System.Management.Automation.PSObject'.
            If ($ComputerCount -eq 1)
            {
                $ADComputers | ForEach-Object {
                    $Obj = New-Object PSObject
                    $Obj | Add-Member -MemberType NoteProperty -Name Name -Value $($_.Name)
                    $Obj | Add-Member -MemberType NoteProperty -Name DNSHostName -Value $($_.DNSHostName)
                    $Obj | Add-Member -MemberType NoteProperty -Name Enabled -Value $($_.Enabled)
                    $Obj | Add-Member -MemberType NoteProperty -Name IPv4Address -Value $($_.IPv4Address)
                    If ($null -ne $_.OperatingSystem)
                    {
                        $Obj | Add-Member -MemberType NoteProperty -Name OperatingSystem -Value $($_.OperatingSystem)
                    }
                    Else
                    {
                        $Obj | Add-Member -MemberType NoteProperty -Name OperatingSystem -Value "-"
                    }
                    If ($null -eq $_.LastLogonDate)
                    {
                        $Obj | Add-Member -MemberType NoteProperty -Name "Days Since Last Logon" -Value "-"
                    }
                    Else
                    {
                        $DDiff = (Get-DateDiff $_.LastLogonDate $date).Days
                        $Obj | Add-Member -MemberType NoteProperty -Name "Days Since Last Logon" -Value $DDiff
                    }
                    If ($null -eq $_.PasswordLastSet)
                    {
                        $Obj | Add-Member -MemberType NoteProperty -Name "Days Since Last Password Change" -Value "-"
                    }
                    Else
                    {
                        $DDiff = (Get-DateDiff $_.PasswordLastSet $date).Days
                        $Obj | Add-Member -MemberType NoteProperty -Name "Days Since Last Password Change" -Value $DDiff
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name "Trusted for Delegation" -Value $($_.TrustedForDelegation)
                    $Obj | Add-Member -MemberType NoteProperty -Name "Trusted to Auth for Delegation" -Value $($_.TrustedToAuthForDelegation)
                    $Obj | Add-Member -MemberType NoteProperty -Name "Username" -Value $($_.SamAccountName)
                    $Obj | Add-Member -MemberType NoteProperty -Name "Primary Group ID" -Value $($_.primaryGroupID)
                    $Obj | Add-Member -MemberType NoteProperty -Name "Description" -Value $($_.Description)
                    If ($null -eq $_.PasswordLastSet)
                    {
                        $Obj | Add-Member -MemberType NoteProperty -Name "Password LastSet" -Value "-"
                    }
                    Else
                    {
                        $Obj | Add-Member -MemberType NoteProperty -Name "Password LastSet" -Value $($_.PasswordLastSet)
                    }
                    If ($null -eq $_.LastLogonDate)
                    {
                        $Obj | Add-Member -MemberType NoteProperty -Name "Last Logon Date" -Value "-"
                    }
                    Else
                    {
                        $Obj | Add-Member -MemberType NoteProperty -Name "Last Logon Date" -Value $($_.LastLogonDate)
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name "whenCreated" -Value $($_.whenCreated)
                    $Obj | Add-Member -MemberType NoteProperty -Name "whenChanged" -Value $($_.whenChanged)
                    $Obj | Add-Member -MemberType NoteProperty -Name 'Distinguished Name' -Value $($_.DistinguishedName)
                    Switch ($OutputType)
                    {
                        'STDOUT' { $Obj }
                        'CSV'
                        {
                            Try
                            {
                                $Obj | Export-Csv -Path $ADFileName -NoTypeInformation
                            }
                            Catch
                            {
                                Write-Output "[ERROR] Failed to Export CSV File"
                                Write-Output "[EXCEPTION] $($_.Exception.Message)"
                            }
                        }
                    }
                }
            }
            Remove-Variable ADComputers
            Write-Verbose "[+] Domain Computers"
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(samAccountType=805306369)"
        $ObjSearcher.PropertiesToLoad.AddRange(("description","name","pwdlastset","useraccountcontrol","samaccountname","dnshostname","lastlogontimestamp","primarygroupid","whenchanged","whencreated","operatingsystem","distinguishedname"))
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADComputers = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADComputers)
        {
            $ADFileName  = -join($ReportPath,'\','Computers','.csv')
            [ADRecon.LDAPClass]::ComputerParser($ADComputers, $date, $ADFileName, $Threads, $FlushCount, $OutputType)
            Remove-Variable ADComputers
            Write-Verbose "[+] Domain Computers"
        }
    }
}

Function Get-ADRComputerSPN
{
<#
.SYNOPSIS
    Returns all computer service principal name (SPN) in the current (or specified) domain.

.DESCRIPTION
    Returns all computer service principal name (SPN) in the current (or specified) domain.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.PARAMETER ReportPath
    [string]
    Path for ADRecon output folder.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.PARAMETER FlushCount
    [int]
    The number of processed objects which will be flushed to disk. Default -1 (After all objects are processed).

.PARAMETER OutputType
    [array]
    Output Type.

.OUTPUTS
    CSV file is created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [bool] $UseAltCreds,

        [Parameter(Mandatory = $true)]
        [string] $ReportPath,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10,

        [Parameter(Mandatory = $false)]
        [int] $FlushCount = -1,

        [Parameter(Mandatory = $true)]
        [array] $OutputType
    )

    Write-Output "[-] Domain Computer SPNs"
    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADComputers = Get-ADObject -LDAPFilter "(&(objectClass=computer)(servicePrincipalName=*))" -Properties name,dnshostname,servicePrincipalName -ResultPageSize $PageSize
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }

        If ($ADComputers)
        {
            $ADFileName = -join($ReportPath,'\','ComputerSPNs','.csv')
            $ComputerSPNCount = [ADRecon.ADWSClass]::ComputerSPNParser($ADComputers, $ADFileName, $Threads, $FlushCount, $OutputType)
            # Temporary solution for exception in [ADRecon.ADWSClass]::ComputerSPNParser
            # System.InvalidCastException: Unable to cast object of type 'Microsoft.ActiveDirectory.Management.ADComputer' to type 'System.Management.Automation.PSObject'.
            If ($ComputerSPNCount -eq 1)
            {
                $CompSPNObj = @()
                $ADComputers | ForEach-Object {
                    For($i=0; $i -lt $_.servicePrincipalName.count; $i++)
                    {
                        $Obj = New-Object PSObject
                        [array] $SPNObjectArray = $_.servicePrincipalName[$i] -Split("/")
                        $Obj | Add-Member -MemberType NoteProperty -Name "Name" -Value $_.Name
                        $Obj | Add-Member -MemberType NoteProperty -Name "Service" -Value $SPNObjectArray[0]
                        $Obj | Add-Member -MemberType NoteProperty -Name "Host" -Value $SPNObjectArray[1]
                        $CompSPNObj += $Obj
                        Remove-Variable SPNObjectArray
                    }
                }
                If ($CompSPNObj)
                {
                    Switch ($OutputType)
                    {
                        'STDOUT' { $CompSPNObj }
                        'CSV'
                        {
                            $ADFileName = -join($ReportPath,'\','ComputerSPNs','.csv')
                            Try
                            {
                                $CompSPNObj | Export-Csv -Path $ADFileName -NoTypeInformation
                            }
                            Catch
                            {
                                Write-Output "Failed to Export CSV File"
                                Write-Output "[EXCEPTION] $($_.Exception.Message)"
                            }
                        }
                    }
                }
            }
            Remove-Variable ADComputers
            Write-Verbose "[+] Domain Computer SPNs"
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(&(objectClass=computer)(servicePrincipalName=*))"
        $ObjSearcher.PropertiesToLoad.AddRange(("name","samaccountname","serviceprincipalname","pwdlastset","description"))
        $ObjSearcher.SearchScope = "Subtree"
        Try
        {
            $ADComputers = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADComputers)
        {
            $ADFileName = -join($ReportPath,'\','ComputerSPNs','.csv')
            [ADRecon.LDAPClass]::ComputerSPNParser($ADComputers, $ADFileName, $Threads, $FlushCount, $OutputType)
            Remove-Variable ADComputers
            Write-Verbose "[+] Domain Computer SPNs"
        }
    }
}

Function Get-ADRLAPSCheck
{
<#
.SYNOPSIS
    Returns all LAPS (local administrator) stored passwords in the current (or specified) domain.

.DESCRIPTION
    Returns all LAPS (local administrator) stored passwords in the current (or specified) domain. Other details such as the Password Expiration, whether the password is readable by the current user are also returned.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.PARAMETER ReportPath
    [string]
    Path for ADRecon output folder.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER OutputType
    [array]
    Output Type.

.OUTPUTS
    CSV file is created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [bool] $UseAltCreds,

        [Parameter(Mandatory = $true)]
        [string] $ReportPath,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $true)]
        [array] $OutputType
    )

    Write-Output "[-] LAPS - Needs Privileged Account"
    # based on https://github.com/kfosaaen/Get-LAPSPasswords/blob/master/Get-LAPSPasswords.ps1
    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADComputers = Get-ADObject -LDAPFilter "(objectClass=computer)" -Properties cn,dnshostname,'ms-mcs-admpwd','ms-mcs-admpwdexpirationtime' -ResultPageSize $PageSize
        }
        Catch [System.ArgumentException]
        {
            Write-Output "[*] LAPS is not implemented."
            $LAPS = $false
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }

        If ($ADComputers -and $LAPS -ne $false)
        {
            $LAPSObj = @()
            $ADComputers | ForEach-Object {
                [string] $CurrentPassword = $_.'ms-mcs-admpwd'
                If ($_.'ms-mcs-admpwdexpirationtime' -ge 0)
                {
                    $CurrentExpiration = [dateTime]::FromFileTime("$($_.'ms-mcs-admpwdexpirationtime')")
                }
                Else
                {
                    $CurrentExpiration = "NA"
                }
                $PasswordAvailable = $false
                $PasswordStored = $true
                If ($CurrentPassword.length -ge 1)
                {
                    $PasswordAvailable = $true
                }
                If ($CurrentExpiration -eq "NA")
                {
                    $PasswordStored = $false
                    $PasswordAvailable = "NA"
                    $CurrentPassword = $null
                }
                If ($null -ne $_.dnshostname)
                {
                    $CurrentHostname = $_.dnshostname
                }
                Else
                {
                    $CurrentHostname = $_.cn
                }
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name Hostname -Value $CurrentHostname
                $Obj | Add-Member -MemberType NoteProperty -Name Stored -Value $PasswordStored
                $Obj | Add-Member -MemberType NoteProperty -Name Readable -Value $PasswordAvailable
                $Obj | Add-Member -MemberType NoteProperty -Name Password -Value $CurrentPassword
                $Obj | Add-Member -MemberType NoteProperty -Name Expiration -Value $CurrentExpiration
                $LAPSObj += $Obj
                Remove-Variable CurrentHostname
                Remove-Variable PasswordStored
                Remove-Variable PasswordAvailable
                Remove-Variable CurrentPassword
                Remove-Variable CurrentExpiration
            }
            Remove-Variable ADComputers
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(objectClass=computer)"
        $ObjSearcher.PropertiesToLoad.AddRange(("cn","dnshostname","ms-mcs-admpwdexpirationtime","ms-mcs-admpwd"))
        $ObjSearcher.SearchScope = "Subtree"
        Try
        {
            $ADComputers = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        $ObjSearcher.dispose()

        If ($($ADComputers | ForEach-Object {$_.Properties.'ms-mcs-admpwdexpirationtime'} | Measure-Object | Select-Object -ExpandProperty Count) -eq 0)
        {
            Write-Output "[*] LAPS is not implemented."
        }
        Else
        {
            $LAPSObj = @()
            $ADComputers | ForEach-Object {
                [string] $CurrentPassword = $_.properties.'ms-mcs-admpwd'
                If ($_.properties.'ms-mcs-admpwdexpirationtime' -ge 0)
                {
                    $CurrentExpiration = [dateTime]::FromFileTime("$($_.properties.'ms-mcs-admpwdexpirationtime')")
                }
                Else
                {
                    $CurrentExpiration = "NA"
                }
                $PasswordAvailable = $false
                $PasswordStored = $true
                If ($CurrentPassword.length -ge 1)
                {
                    $PasswordAvailable = $true
                }
                If ($CurrentExpiration -eq "NA")
                {
                    $PasswordStored = $false
                    $PasswordAvailable = "NA"
                    $CurrentPassword = $null
                }
                If ($null -ne $_.properties.dnshostname)
                {
                    $CurrentHostname = ([string] $($_.properties.dnshostname))
                }
                Else
                {
                    $CurrentHostname = ([string] $($_.properties.cn))
                }
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name Hostname -Value $CurrentHostname
                $Obj | Add-Member -MemberType NoteProperty -Name Stored -Value $PasswordStored
                $Obj | Add-Member -MemberType NoteProperty -Name Readable -Value $PasswordAvailable
                $Obj | Add-Member -MemberType NoteProperty -Name Password -Value $CurrentPassword
                $Obj | Add-Member -MemberType NoteProperty -Name Expiration -Value $CurrentExpiration
                $LAPSObj += $Obj
                Remove-Variable CurrentHostname
                Remove-Variable PasswordStored
                Remove-Variable PasswordAvailable
                Remove-Variable CurrentPassword
                Remove-Variable CurrentExpiration
            }
            Remove-Variable ADComputers
        }
    }

    If ($LAPSObj)
    {
        Switch ($OutputType)
        {
            'STDOUT' { $LAPSObj }
            'CSV'
            {
                Write-Verbose "[+] LAPS"
                $ADFileName = -join($ReportPath,'\','LAPS','.csv')
                Try
                {
                    $LAPSObj | Export-Csv -Path $ADFileName -NoTypeInformation
                }
                Catch
                {
                    Write-Output "Failed to Export CSV File"
                    Write-Output "[EXCEPTION] $($_.Exception.Message)"
                }
                Remove-Variable LAPSObj
                Remove-Variable ADFileName
            }
        }
    }
}

Function Get-ADRBitLocker
{
<#
.SYNOPSIS
    Returns all BitLocker Recovery Keys stored in the current (or specified) domain.
.DESCRIPTION
    Returns all BitLocker Recovery Keys stored in the current (or specified) domain.
.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.
.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.
.PARAMETER ReportPath
    [string]
    Path for ADRecon output folder.
.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.
.PARAMETER DCIP
    [string]
    IP Address of the Domain Controller.
.PARAMETER creds
    [Management.Automation.PSCredential]
    Credentials.
.PARAMETER OutputType
    [array]
    Output Type.
.OUTPUTS
    CSV file is created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [bool] $UseAltCreds,

        [Parameter(Mandatory = $true)]
        [string] $ReportPath,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [string] $DCIP,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $creds = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $true)]
        [array] $OutputType
    )

    Write-Output "[-] BitLocker Recovery Keys - Needs Privileged Account"
    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            $ADBitLockerRecoveryKeys = Get-ADObject -LDAPFilter '(objectClass=msFVE-RecoveryInformation)' -Properties distinguishedName,msFVE-RecoveryPassword,msFVE-RecoveryGuid,msFVE-VolumeGuid,Name,whenCreated
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }

        If ($ADBitLockerRecoveryKeys)
        {
            $cnt = $($ADBitLockerRecoveryKeys | Measure-Object | Select-Object -ExpandProperty Count)
            If ($cnt -ge 1)
            {
                Write-Output "[*] Total BitLocker Recovery Keys: $cnt"
                $BitLockerObj = @()
                $ADBitLockerRecoveryKeys | ForEach-Object {
                    # Create the object for each instance.
                    $Obj = New-Object PSObject
                    $Obj | Add-Member -MemberType NoteProperty -Name "Distinguished Name" -Value $((($_.distinguishedName -split '}')[1]).substring(1))
                    $Obj | Add-Member -MemberType NoteProperty -Name "Name" -Value $_.Name
                    $Obj | Add-Member -MemberType NoteProperty -Name "Created" -Value $_.whenCreated
                    $Obj | Add-Member -MemberType NoteProperty -Name "Recovery Key ID" -Value $([GUID] $_.'msFVE-RecoveryGuid')
                    $Obj | Add-Member -MemberType NoteProperty -Name "Recovery Key" -Value $_.'msFVE-RecoveryPassword'
                    $Obj | Add-Member -MemberType NoteProperty -Name "Volume GUID" -Value $([GUID] $_.'msFVE-VolumeGuid')
                    #$TempComp = Get-ADComputer -Identity $((($((($_.distinguishedName -split '}')[1]).substring(1)) -Split ("CN=")) -Split (","))[1]) -Properties msTPM-OwnerInformation,msTPM-TpmInformationForComputer
                    $TempComp = Get-ADComputer -Identity $Obj.'Distinguished Name' -Properties msTPM-OwnerInformation,msTPM-TpmInformationForComputer
                    # msTPM-OwnerInformation (Vista/7 or Server 2008/R2)
                    $Obj | Add-Member -MemberType NoteProperty -Name "msTPM-OwnerInformation" -Value $TempComp.'msTPM-OwnerInformation'
                    
                    # msTPM-TpmInformationForComputer (Windows 8/10 or Server 2012/R2)
                    $Obj | Add-Member -MemberType NoteProperty -Name "msTPM-TpmInformationForComputer" -Value $TempComp.'msTPM-TpmInformationForComputer'
                    If ($TempComp.'msTPM-TpmInformationForComputer' -ne $null)
                    {
                        # Grab the TPM Owner Info from the msTPM-InformationObject
                        $TPMObject = Get-ADObject -Identity $TempComp.'msTPM-TpmInformationForComputer' -Properties msTPM-OwnerInformation
                        $TPMRecoveryInfo = $TPMObject.'msTPM-OwnerInformation'
                    }
                    Else
                    {
                        $TPMRecoveryInfo = $null
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name "TPM Owner Password" -Value $TPMRecoveryInfo
                    $BitLockerObj += $Obj
                }
            }
            Remove-Variable ADBitLockerRecoveryKeys
        }
    }

    If ($Protocol -eq 'LDAP')
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = "(objectClass=msFVE-RecoveryInformation)"
        $ObjSearcher.PropertiesToLoad.AddRange(("distinguishedName","msfve-recoverypassword","msfve-recoveryguid","msfve-volumeguid","mstpm-ownerinformation","mstpm-tpminformationforcomputer","name","whencreated"))
        $ObjSearcher.SearchScope = "Subtree"

        Try
        {
            $ADBitLockerRecoveryKeys = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
        }
        $ObjSearcher.dispose()

        If ($ADBitLockerRecoveryKeys)
        {
            $cnt = $($ADBitLockerRecoveryKeys | Measure-Object | Select-Object -ExpandProperty Count)
            If ($cnt -ge 1)
            {
                Write-Output "[*] Total BitLocker Recovery Keys: $cnt"
                $BitLockerObj = @()
                $ADBitLockerRecoveryKeys | ForEach-Object {
                    # Create the object for each instance.
                    $Obj = New-Object PSObject
                    $Obj | Add-Member -MemberType NoteProperty -Name "Distinguished Name" -Value $((($_.Properties.distinguishedname -split '}')[1]).substring(1))
                    $Obj | Add-Member -MemberType NoteProperty -Name "Name" -Value ([string] ($_.Properties.name))
                    $Obj | Add-Member -MemberType NoteProperty -Name "Created" -Value ([DateTime] $($_.Properties.whencreated))
                    $Obj | Add-Member -MemberType NoteProperty -Name "Recovery Key ID" -Value $([GUID] $_.Properties.'msfve-recoveryguid'[0])
                    $Obj | Add-Member -MemberType NoteProperty -Name "Recovery Key" -Value ([string] ($_.Properties.'msfve-recoverypassword'))
                    $Obj | Add-Member -MemberType NoteProperty -Name "Volume GUID" -Value $([GUID] $_.Properties.'msfve-volumeguid'[0])

                    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
                    $ObjSearcher.PageSize = $PageSize
                    $ObjSearcher.Filter = "(&(samAccountType=805306369)(distinguishedName=$($Obj.'Distinguished Name')))"
                    $ObjSearcher.PropertiesToLoad.AddRange(("mstpm-ownerinformation","mstpm-tpminformationforcomputer"))
                    $ObjSearcher.SearchScope = "Subtree"

                    Try
                    {
                        $TempComp = $ObjSearcher.FindAll()
                    }
                    Catch
                    {
                        Write-Output "[EXCEPTION] $($_.Exception.Message)"
                    }
                    $ObjSearcher.dispose()
                    
                    If ($TempComp)
                    {
                        # msTPM-OwnerInformation (Vista/7 or Server 2008/R2)
                        $Obj | Add-Member -MemberType NoteProperty -Name "msTPM-OwnerInformation" -Value $([string] $TempComp.Properties.'mstpm-ownerinformation')

                        # msTPM-TpmInformationForComputer (Windows 8/10 or Server 2012/R2)
                        $Obj | Add-Member -MemberType NoteProperty -Name "msTPM-TpmInformationForComputer" -Value $([string] $TempComp.Properties.'mstpm-tpminformationforcomputer')
                        If ($TempComp.Properties.'mstpm-tpminformationforcomputer' -ne $null)
                        {
                            # Grab the TPM Owner Info from the msTPM-InformationObject
                            If ($UseAltCreds)
                            {
                                $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DCIP)/$($TempComp.Properties.'mstpm-tpminformationforcomputer')", $creds.UserName,$creds.GetNetworkCredential().Password
                                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                                $objSearcherPath.PropertiesToLoad.AddRange(("mstpm-ownerinformation"))
                                $TPMObject = $objSearcherPath.FindAll()
                                $TPMRecoveryInfo = $([string] $TPMObject.Properties.'mstpm-ownerinformation')
                            }
                            Else
                            {
                                $TPMObject = ([ADSI]"LDAP://$($TempComp.Properties.'mstpm-tpminformationforcomputer')")
                                $TPMRecoveryInfo = $([string] $TPMObject.Properties.'mstpm-ownerinformation')
                            }
                        }
                    }
                    Else
                    {
                        $Obj | Add-Member -MemberType NoteProperty -Name "msTPM-OwnerInformation" -Value $null
                        $Obj | Add-Member -MemberType NoteProperty -Name "msTPM-TpmInformationForComputer" -Value $null
                        $TPMRecoveryInfo = $null
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name "TPM Owner Password" -Value $TPMRecoveryInfo
                    $BitLockerObj += $Obj
                }
            }
            Remove-Variable cnt
            Remove-Variable ADBitLockerRecoveryKeys
        }
    }

    If ($BitLockerObj)
    {
        Switch ($OutputType)
        {
            'STDOUT' { $BitLockerObj | Format-List }
            'CSV'
            {
                Write-Verbose "[+] BitLocker Recovery Keys"
                $ADFileName = -join($ReportPath,'\','BitLockerRecoveryKeys','.csv')
                Try
                {
                    $BitLockerObj | Export-Csv -Path $ADFileName -NoTypeInformation
                }
                Catch
                {
                    Write-Output "Failed to Export CSV File"
                    Write-Output "[EXCEPTION] $($_.Exception.Message)"
                }
                Remove-Variable BitLockerObj
                Remove-Variable ADFileName
            }
        }
    }
}

Function Get-ADRGPOReport
{
<#
.SYNOPSIS
    Runs the Get-GPOReport cmdlet if available.

.DESCRIPTION
    Runs the Get-GPOReport cmdlet if available and saves in HTML and XML formats.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.PARAMETER ReportPath
    [string]
    Path for ADRecon output folder.

.OUTPUTS
    CSV file is created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [bool] $UseAltCreds,

        [Parameter(Mandatory = $true)]
        [string] $ReportPath
    )

    Write-Output "[-] Domain GPO Report - May take some time"
    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            Import-Module GroupPolicy -WarningAction Stop -ErrorAction Stop | Out-Null
        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
        Try
        {
            Write-Output "[*] Domain GPO Report XML"
            $ADFileName = -join($ReportPath,'\','GPO-Report','.xml')
            Get-GPOReport -All -ReportType XML -Path $ADFileName

        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
            If ($UseAltCreds)
            {
                Write-Output "[*] Run the tool using RUNAS."
                Write-Output "[*] runas /user:<Domain FQDN>\<Username> /netonly powershell.exe"
            }
            Return $null
        }
        Try
        {
            Write-Output "[*] Domain GPO Report HTML"
            $ADFileName = -join($ReportPath,'\','GPO-Report','.html')
            Get-GPOReport -All -ReportType HTML -Path $ADFileName

        }
        Catch
        {
            Write-Output "[EXCEPTION] $($_.Exception.Message)"
            If ($UseAltCreds)
            {
                Write-Output "[*] Run the tool using RUNAS."
                Write-Output "[*] runas /user:<Domain FQDN>\<Username> /netonly powershell.exe"
            }
            Return $null
        }
    }
    If ($Protocol -eq 'LDAP')
    {
        Write-Output "[*] Currently, the module is only supported with ADWS."
    }
}

Function Remove-EmptyADROutputDir
{
<#
.SYNOPSIS
    Removes ADRecon output folder if empty.

.DESCRIPTION
    Removes ADRecon output folder if empty.

.PARAMETER OutputDir
    [string]
	Path for ADRecon output folder.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $OutputDir
    )
    If (!(Test-Path -Path $OutputDir\*))
    {
        Remove-Item $OutputDir
        Write-Verbose "Removed Empty Directory $OutputDir"
    }
}

Function Invoke-ADRecon
{
<#
.SYNOPSIS
    Wrapper function to run ADRecon modules.

.DESCRIPTION
    Wrapper function to set variables, check dependencies and run ADRecon modules.

.PARAMETER Protocol
    [string]
    Which protocol to use; ADWS (default) or LDAP.

.PARAMETER Collect
    [array]
    What attributes to collect; Forest, Domain, PasswordPolicy, DCs, Users, UserSPNs, Groups, GroupMembers, OUs, OUPermissions, GPOs, GPOReport, DNSZones, Printers, Computers, ComputerSPNs, LAPS, BitLocker

.PARAMETER DCIP
    [string]
    IP Address of the Domain Controller.

.PARAMETER creds
    [Management.Automation.PSCredential]
    Credentials.

.PARAMETER OutputDir
    [string]
	Path for ADRecon output folder to save the CSV files and the ADRecon-Report.xlsx.

.PARAMETER DormantTimeSpan
    [int]
    Timespan for Dormant accounts. Default 90 days.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.PARAMETER FlushCount
    [int]
    The number of processed objects which will be flushed to disk. Default -1 (After all objects are processed).

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.OUTPUTS
    CSV file is created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = $false)]
        [string] $GenExcel,

        [Parameter(Mandatory = $true)]
        [string] $Protocol,

        [Parameter(Mandatory = $true)]
        [array] $Collect,

        [Parameter(Mandatory = $false)]
        [string] $DCIP,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $creds = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $true)]
        [array] $OutputType,

        [Parameter(Mandatory = $false)]
        [string] $ExcelPath,

        [Parameter(Mandatory = $true)]
        [int] $DormantTimeSpan = 90,

        [Parameter(Mandatory = $true)]
        [int] $PageSize = 200,

        [Parameter(Mandatory = $true)]
        [int] $Threads = 10,

        [Parameter(Mandatory = $true)]
        [int] $FlushCount = -1,

        [Parameter(Mandatory = $false)]
        [bool] $UseAltCreds = $false
    )

    [string] $ADReconVersion = "v180429"
    Write-Output "[*] ADRecon $ADReconVersion by Prashant Mahajan (@prashant3535) from Sense of Security."

    If ($GenExcel)
    {
        If (!(Test-Path $GenExcel))
        {
            Write-Output "[ERROR] Invalid Path ... Exiting"
            Return $null
        }
        Get-ADRGenExcel $GenExcel
        Return $null
    }

    Try
    {
        If ($PSVersionTable.PSVersion.Major -ne 2)
        {
            $computer = Get-CimInstance -ClassName Win32_ComputerSystem
            $computerdomainrole = ($computer).DomainRole
        }
        Else
        {
            $computer = Get-WMIObject win32_computersystem
            $computerdomainrole = ($computer).DomainRole
        }
    }
    Catch
    {
        Write-Output "[EXCEPTION] $($_.Exception.Message)"
    }

    switch ($computerdomainrole)
    {
        0
        {
            [string] $computerrole = "Standalone Workstation"
            $Env:ADPS_LoadDefaultDrive = 0
            $UseAltCreds = $true
        }
        1 { [string] $computerrole = "Member Workstation" }
        2
        {
            [string] $computerrole = "Standalone Server"
            $UseAltCreds = $true
            $Env:ADPS_LoadDefaultDrive = 0
        }
        3 { [string] $computerrole = "Member Server" }
        4 { [string] $computerrole = "Backup Domain Controller" }
        5 { [string] $computerrole = "Primary Domain Controller" }
        default { Write-Output "Computer Role could not be identified." }
    }

    If (($DCIP -ne "") -or ($creds -ne [Management.Automation.PSCredential]::Empty))
    {
        $UseAltCreds = $true
    }

    If ($Protocol -eq 'ADWS')
    {
        Try
        {
            Import-Module ActiveDirectory -WarningAction Stop -ErrorAction Stop | Out-Null
        }
        Catch
        {
            Write-Warning "ActiveDirectory Module from RSAT (Remote Server Administration Tools) is not installed. ... Continuing with LDAP"
            $Protocol = 'LDAP'
        }
    }

    Try
    {
        Add-Type -TypeDefinition $PingCastleSMBScannerSource
        $CLR4 = ([System.Reflection.Assembly]::GetExecutingAssembly().ImageRuntimeVersion)[1]
        If ($Protocol -eq 'ADWS')
        {
            If ($CLR4 -eq "4")
            {
                Add-Type -TypeDefinition $ADWSSource -ReferencedAssemblies ([system.reflection.assembly]::LoadWithPartialName("Microsoft.ActiveDirectory.Management")).Location
            }
            Else
            {
                Add-Type -TypeDefinition $ADWSSource -ReferencedAssemblies ([system.reflection.assembly]::LoadWithPartialName("Microsoft.ActiveDirectory.Management")).Location -Language CSharpVersion3
            }
        }

        If ($Protocol -eq 'LDAP')
        {
            If ($CLR4 -eq "4")
            {
                Add-Type -TypeDefinition $LDAPSource -ReferencedAssemblies ([system.reflection.assembly]::LoadWithPartialName("System.DirectoryServices")).Location
            }
            Else
            {
                Add-Type -TypeDefinition $LDAPSource -ReferencedAssemblies ([system.reflection.assembly]::LoadWithPartialName("System.DirectoryServices")).Location -Language CSharpVersion3
            }
            # Allow running using RUNAS from a non-domain joined machine
            # runas /user:<Domain FQDN>\<Username> /netonly powershell.exe
            If (($DCIP -eq "") -and ($creds -eq [Management.Automation.PSCredential]::Empty))
            {
                Try
                {
                    $objDomain = [ADSI]""
                    $UseAltCreds = $false
                    $objDomain.Dispose()
                }
                Catch
                {
                    $UseAltCreds = $true
                }
            }
        }
    }
    Catch
    {
        Write-Output "[ERROR] $($_.Exception.Message)"
        Return $null
    }

    If ($UseAltCreds -and (($DCIP -eq "") -or ($creds -eq [Management.Automation.PSCredential]::Empty)))
    {
        If (($DCIP -ne "") -and ($creds -eq [Management.Automation.PSCredential]::Empty))
        {
            Try
            {
                $creds = Get-Credential
            }
            Catch
            {
                Write-Output "[ERROR] $($_.Exception.Message)"
                Return $null
            }
        }
        Else
        {
            Write-Output "Run Get-Help .\ADRecon.ps1 -Examples for additional information."
            Write-Output "[ERROR] Use the -DomainController and -Credential parameter."`n
            Return $null
        }
    }

    Write-Output "[*] Running on $($computer.domain)\$($env:computername) - $($computerrole)"


    Remove-Variable computer
    Remove-Variable computerdomainrole

    Switch ($Collect)
    {
        'Forest' { $ADRForest = $true }
        'Domain' {$ADRDomain = $true }
        'PasswordPolicy' { $ADRPasswordPolicy = $true }
        'DCs' { $ADRDCs = $true }
        'Users' { $ADRUsers = $true }
        'UserSPNs' { $ADRUserSPNs = $true }
        'Groups' { $ADRGroups = $true }
        'GroupMembers' { $ADRGroupMembers = $true }
        'OUs' { $ADROUs = $true }
        'OUPermissions' { $ADROUPermissions = $true }
        'GPOs' { $ADRGPOs = $true }
        'GPOReport' { $ADRGPOReport = $true }
        'DNSZones' { $ADRDNSZones = $true }
        'Printers' { $ADRPrinters = $true }
        'Computers' { $ADRComputers = $true }
        'ComputerSPNs' { $ADRCopmuterSPNs = $true }
        'BitLocker' { $ADRBitLocker = $true }
        'LAPS' { $ADRLAPS = $true }
        'Default'
        {
            $ADRForest = $true
            $ADRDomain = $true
            $ADRPasswordPolicy = $true
            $ADRDCs = $true
            $ADRUsers = $true
            $ADRUserSPNs = $true
            $ADRGroups = $true
            $ADRGroupMembers = $true
            $ADROUs = $true
            $ADROUPermissions = $true
            $ADRGPOs = $true
            $ADRGPOReport = $true
            $ADRDNSZones = $true
            $ADRPrinters = $true
            $ADRComputers = $true
            $ADRCopmuterSPNs = $true
            $ADRLAPS = $true
            $ADRBitLocker = $true
            If ($OutputType -eq "Default")
            {
                [array] $OutputType = "CSV","Excel"
            }
        }
    }

    Switch ($OutputType)
    {
        'CSV' { $ADRCSV = $true }
        'STDOUT' { $ADRSTDOUT = $true }
        'Excel' { $ADRExcel = $true }
        'Default' { [array] $OutputType = "STDOUT" }
    }

    If ($ADRExcel)
    {
        If (!($ADRCSV))
        {
            $ADRCSV = $true
            If ($ADRSTDOUT)
            {
                [array] $OutputType = "CSV","STDOUT","Excel"
            }
            Else
            {
                [array] $OutputType = "CSV","Excel"
            }
        }
    }

    $returndir = Get-Location
    $date = Get-Date

    If ($ExcelPath)
    {
        If (!(Test-Path $ExcelPath))
        {
            New-Item $ExcelPath -type directory | Out-Null
            If (!(Test-Path $ExcelPath))
            {
                Write-Output "[ERROR] Invalid OutputDir Path ... Exiting"
                Return $null
            }
        }
        $ExcelPath = $((Convert-Path $ExcelPath).TrimEnd("\"))
        Write-Verbose $ExcelPath
    }

    If ($ADRCSV)
    {
        If (-Not $ExcelPath)
        {
            $ExcelPath =  -join($returndir,'\','ADRecon-Report-',$date.day,$date.Month,$date.Year,$date.Hour,$date.Minute,$date.Second)
            New-Item $ExcelPath -type directory | Out-Null
        }
        $ReportPath = [System.IO.DirectoryInfo] -join($ExcelPath,'\','CSV-Files')
        New-Item $ReportPath -type directory | Out-Null

        If (!(Test-Path $ReportPath))
        {
            Write-Output "[ERROR] Could not create output directory"
            return $null
        }
    }
    Else
    {
        If ($ADRGPOReport)
        {
            If (-Not $ExcelPath)
            {
                $ExcelPath =  -join($returndir,'\','ADRecon-Report-',$date.day,$date.Month,$date.Year,$date.Hour,$date.Minute,$date.Second)
                New-Item $ExcelPath -type directory | Out-Null
            }
            $ReportPath = $ExcelPath
        }
        Else
        {
            $ReportPath = $returndir
        }
    }

    If ($UseAltCreds -and ($Protocol -eq 'ADWS'))
    {
        If (!(Test-Path ADR:))
        {
            Try
            {
                New-PSDrive -PSProvider ActiveDirectory -Name ADR -Root "" -Server $DCIP -Credential $creds -ErrorAction Stop | Out-Null
            }
            Catch
            {
                Write-Output "[EXCEPTION] $($_.Exception.Message)"
                If ($ReportPath)
                {
                    Remove-EmptyADROutputDir $ReportPath
                }
                If ($ExcelPath)
                {
                    Remove-EmptyADROutputDir $ExcelPath
                }
                Return $null
            }
        }
        Else
        {
            Remove-PSDrive ADR
            Try
            {
                New-PSDrive -PSProvider ActiveDirectory -Name ADR -Root "" -Server $DCIP -Credential $creds -ErrorAction Stop | Out-Null
            }
            Catch
            {
                Write-Output "[EXCEPTION] $($_.Exception.Message)"
                If ($ReportPath)
                {
                    Remove-EmptyADROutputDir $ReportPath
                }
                If ($ExcelPath)
                {
                    Remove-EmptyADROutputDir $ExcelPath
                }
                Return $null
            }
        }
        Set-Location ADR:
    }

    If ($Protocol -eq 'LDAP')
    {
        If ($UseAltCreds)
        {
            Try
            {
                $objDomain = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DCIP)", $creds.UserName,$creds.GetNetworkCredential().Password
                $objDomainRootDSE = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DCIP)/RootDSE", $creds.UserName,$creds.GetNetworkCredential().Password
            }
            Catch
            {
                Write-Output "[ERROR] $($_.Exception.Message)"
                If ($ReportPath)
                {
                    Remove-EmptyADROutputDir $ReportPath
                }
                If ($ExcelPath)
                {
                    Remove-EmptyADROutputDir $ExcelPath
                }
                Return $null
            }
            If(!($objDomain.name))
            {
                Write-Output "[ERROR] LDAP bind Unsuccessful"
                If ($ReportPath)
                {
                    Remove-EmptyADROutputDir $ReportPath
                }
                If ($ExcelPath)
                {
                    Remove-EmptyADROutputDir $ExcelPath
                }
                Return $null
            }
            Else
            {
                Write-Output "[*] LDAP bind Successful"
            }
        }
        Else
        {
            $objDomain = [ADSI]""
            $objDomainRootDSE = ([ADSI] "LDAP://RootDSE")
            If(!($objDomain.name))
            {
                Write-Output "[ERROR] LDAP bind Unsuccessful"
                If ($ReportPath)
                {
                    Remove-EmptyADROutputDir $ReportPath
                }
                If ($ExcelPath)
                {
                    Remove-EmptyADROutputDir $ExcelPath
                }
                Return $null
            }
        }
    }

    Write-Output "[*] Commencing - $date"
    If ($ADRDomain) { Get-ADRDomain $Protocol $UseAltCreds $ReportPath $objDomain $objDomainRootDSE $DCIP $creds $OutputType }
    If ($ADRForest) { Get-ADRForest $Protocol $UseAltCreds $ReportPath $objDomain $objDomainRootDSE $DCIP $creds $OutputType }
    If ($ADRPasswordPolicy) { Get-ADRPassPol $Protocol $UseAltCreds $ReportPath $objDomain $OutputType }
    If ($ADRDCs) { Get-ADRDC $Protocol $UseAltCreds $ReportPath $objDomain $OutputType }
    If ($ADRUsers) { Get-ADRUser $Protocol $UseAltCreds $ReportPath $date $objDomain $DormantTimeSpan $PageSize $Threads $FlushCount $OutputType }
    If ($ADRUserSPNs) { Get-ADRUserSPN $Protocol $UseAltCreds $ReportPath $objDomain $PageSize $Threads $FlushCount $OutputType }
    If ($ADRGroups) { Get-ADRGroup $Protocol $UseAltCreds $ReportPath $objDomain $PageSize $Threads $FlushCount $OutputType }
    If ($ADRGroupMembers) { Get-ADRGroupMember $Protocol $UseAltCreds $ReportPath $objDomain $PageSize $Threads $FlushCount $OutputType }
    If ($ADROUs) { Get-ADROU $Protocol $UseAltCreds $ReportPath $objDomain $PageSize $OutputType }
    If ($ADROUPermissions) { Get-ADROUPermission $Protocol $UseAltCreds $ReportPath $objDomain $DCIP $creds $PageSize $OutputType }
    If ($ADRGPOs) { Get-ADRGPO $Protocol $UseAltCreds $ReportPath $objDomain $PageSize $OutputType }
    If ($ADRDNSZones) { Get-ADRDNSZone $Protocol $UseAltCreds $ReportPath $objDomain $DCIP $creds $PageSize $OutputType }
    If ($ADRPrinters) { Get-ADRPrinter $Protocol $UseAltCreds $ReportPath $objDomain $PageSize $OutputType }
    If ($ADRComputers) { Get-ADRComputer $Protocol $UseAltCreds $ReportPath $date $objDomain $PageSize $Threads $FlushCount $OutputType }
    If ($ADRCopmuterSPNs) { Get-ADRComputerSPN $Protocol $UseAltCreds $ReportPath $objDomain $PageSize $Threads $FlushCount $OutputType }
    If ($ADRLAPS) { Get-ADRLAPSCheck $Protocol $UseAltCreds $ReportPath $objDomain $PageSize $OutputType }
    If ($ADRBitLocker) { Get-ADRBitLocker $Protocol $UseAltCreds $ReportPath $objDomain $DCIP $creds $OutputType }
    If ($ADRGPOReport) { Get-ADRGPOReport $Protocol $UseAltCreds $ReportPath }
    Switch ($OutputType)
    {
        'CSV'
        {
            $AboutADRecon = New-Object PSObject
            $AboutADRecon | Add-Member -MemberType NoteProperty -Name "Category" -Value "Value"
            $AboutADRecon | Add-Member -MemberType NoteProperty -Name "Date" -Value $($date)
            $AboutADRecon | Add-Member -MemberType NoteProperty -Name "ADRecon" -Value "https://github.com/sense-of-security/ADRecon"
            If ($Protocol -eq 'ADWS')
            {
                $AboutADRecon | Add-Member -MemberType NoteProperty -Name "RSAT Version" -Value $($ADReconVersion)
            }
            Else
            {
                $AboutADRecon | Add-Member -MemberType NoteProperty -Name "LDAP Version" -Value $($ADReconVersion)
            }
            If ($UseAltCreds)
            {
                $AboutADRecon | Add-Member -MemberType NoteProperty -Name "Ran as user" -Value $($creds.UserName)
            }
            Else
            {
                $AboutADRecon | Add-Member -MemberType NoteProperty -Name "Ran as user" -Value $([Environment]::UserName)
            }
            $AboutADRecon | Add-Member -MemberType NoteProperty -Name "Ran from" -Value $([Environment]::MachineName)
            $AboutADRecon | Add-Member -MemberType NoteProperty -Name "Computer Role" -Value $($computerrole)
            $TotalTime = "{0:N2}" -f ((Get-DateDiff (Get-Date) $date).TotalMinutes)
            $AboutADRecon | Add-Member -MemberType NoteProperty -Name "Execution Time (mins)" -Value $($TotalTime)

            Write-Verbose "[+] AboutADRecon"
            If ($AboutADRecon)
            {
                $ADFileName = -join($ReportPath,'\','AboutADRecon','.csv')
                Try
                {
                    $AboutADRecon | Export-Csv -Path $ADFileName -NoTypeInformation
                }
                Catch
                {
                    Write-Output "[ERROR] Failed to Export CSV File"
                    Write-Output "[EXCEPTION] $($_.Exception.Message)"
                }
                Remove-Variable AboutADRecon
                Remove-Variable ADFileName
            }
            Write-Output "[*] Total Execution Time (mins): $($TotalTime)"
            Remove-Variable TotalTime
            Write-Output "[*] Completed."
            Write-Output "[*] Output Directory: $ExcelPath"
        }
        'STDOUT'
        {
            $TotalTime = "{0:N2}" -f ((Get-DateDiff (Get-Date) $date).TotalMinutes)
            Write-Output "[*] Total Execution Time (mins): $($TotalTime)"
            Remove-Variable TotalTime
        }
        'EXCEL'
        {
            Get-ADRGenExcel($ExcelPath)
        }
    }
    Set-Location $returndir
    Remove-Variable returndir

    If (($Protocol -eq 'ADWS') -and $UseAltCreds)
    {
        Remove-PSDrive ADR
    }

    If ($Protocol -eq 'LDAP')
    {
        $objDomain.Dispose()
        $objDomainRootDSE.Dispose()
    }

    If ($ReportPath)
    {
        Remove-EmptyADROutputDir $ReportPath
    }
    If ($ExcelPath)
    {
        Remove-EmptyADROutputDir $ExcelPath
    }

    Remove-Variable ADReconVersion
    Remove-Variable computerrole
}

If ($Log)
{
    Start-Transcript -Path "$(Get-Location)\ADRecon-Console-Log.txt"
}

Invoke-ADRecon $GenExcel $Protocol $Collect $DomainController $Credential $OutputType $OutputDir $DormantTimeSpan $PageSize $Threads $FlushCount

If ($Log)
{
    Stop-Transcript
}
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 

function readResponse {
$result = ""
while($stream.DataAvailable) 
   {  
      $read = $stream.Read($buffer, 0, 1024)    
      $result = $result + ($encoding.GetString($buffer, 0, $read))  
      ""
   } 
return $result
}

function showresult ([String]$result) {
	$Resform = new-object System.Windows.Forms.form 
	$Resform.Text = "Query Result"
	$Resform.topmost = $true
	$Resform.Add_Shown({$form.Activate()})
	$Resbox =  new-object System.Windows.Forms.RichTextBox
	$Resbox.Location = new-object System.Drawing.Size(10,10)
	$Resbox.Size = new-object System.Drawing.Size(260,180)
	$Resbox.Text = $result
	$Resbox.Dock = "Fill"
	$Resform.Controls.Add($Resbox)
	$Resform.ShowDialog()
}


function ptrLookup() {
	
	$ipIpaddressSplit = $logTable.DefaultView[$dgDataGrid.CurrentCell.RowNumber][2].Split(".")
	$revipaddress = $ipIpaddressSplit.GetValue(3) + "." + $ipIpaddressSplit.GetValue(2) + "." + $ipIpaddressSplit.GetValue(1) + "." + $ipIpaddressSplit.GetValue(0) + ".in-addr.arpa"
	$qrQueryresults = [PAB.DnsUtils.DNS]::GetRecords($revipaddress,"PTR")
	showresult($qrQueryresults)
}

function mxLookup() {
	
	if ($logTable.DefaultView[$dgDataGrid.CurrentCell.RowNumber][5] -eq "MAIL" -bor  $logTable.DefaultView[$dgDataGrid.CurrentCell.RowNumber][5] -eq "RCPT"){
		$cmdSplit = $logTable.DefaultView[$dgDataGrid.CurrentCell.RowNumber][6].Split("@")
		$qrQueryresults = [PAB.DnsUtils.DNS]::GetRecords($cmdSplit[$cmdSplit.length-1].Replace(">",""),"MX")
		showresult($qrQueryresults)
	}
	else {
		$msgbox = new-object -comobject wscript.shell
		[void]$msgbox.popup("This is not a RCPT or MAIL command",0,"Cant Do MX lookup",1)
	}
}


function spfLookup() {
	
	if ($logTable.DefaultView[$dgDataGrid.CurrentCell.RowNumber][5] -eq "MAIL" -bor  $logTable.DefaultView[$dgDataGrid.CurrentCell.RowNumber][5] -eq "RCPT"){
		$cmdSplit = $logTable.DefaultView[$dgDataGrid.CurrentCell.RowNumber][6].Split("@")
		$qrQueryresults = [PAB.DnsUtils.DNS]::GetRecords($cmdSplit[$cmdSplit.length-1].Replace(">",""),"SPF")
		showresult($qrQueryresults)
	}
	else {
		$msgbox = new-object -comobject wscript.shell
		[void]$msgbox.popup("This is not a RCPT or MAIL command",0,"Cant Do MX lookup",1)
	}
}

function aLookup() {
	
	if ($logTable.DefaultView[$dgDataGrid.CurrentCell.RowNumber][5] -eq "MAIL" -bor  $logTable.DefaultView[$dgDataGrid.CurrentCell.RowNumber][5] -eq "RCPT"){
		$cmdSplit = $logTable.DefaultView[$dgDataGrid.CurrentCell.RowNumber][6].Split("@")
		$qrQueryresults = [PAB.DnsUtils.DNS]::GetRecords($cmdSplit[$cmdSplit.length-1].Replace(">",""),"A")
		showresult($qrQueryresults)
	}
	else {
		$msgbox = new-object -comobject wscript.shell
		[void]$msgbox.popup("This is not a RCPT or MAIL command",0,"Cant Do MX lookup",1)
	}
}

function RBLLookup() {
	
	$RBLService = "dnsbl.sorbs.net"
	$ipIpaddressSplit = $logTable.DefaultView[$dgDataGrid.CurrentCell.RowNumber][2].Split(".")
	$revipaddress = $ipIpaddressSplit.GetValue(3) + "." + $ipIpaddressSplit.GetValue(2) + "." + $ipIpaddressSplit.GetValue(1) + "." + $ipIpaddressSplit.GetValue(0) + "." + $RBLService
	$qrQueryresults = [PAB.DnsUtils.DNS]::GetRecords($revipaddress,"RBL")
	if ($qrQueryresults -ne "No Record Found"){
	$qrQueryresults = "IP Listed in RBL: " + $qrQueryresults}
	else {$qrQueryresults = "IP Not Listed in RBL: " + $qrQueryresults}
	showresult($qrQueryresults)
}


function HELOchk{

$port = 25 
$socket = new-object System.Net.Sockets.TcpClient($logTable.DefaultView[$dgDataGrid.CurrentCell.RowNumber][2], $port) 
if($socket -eq $null) { showresult("Could not establish connection on Port 25") } 
else{
$stream = $socket.GetStream() 
$writer = new-object System.IO.StreamWriter($stream) 
$buffer = new-object System.Byte[] 1024 
$encoding = new-object System.Text.AsciiEncoding 
readResponse($stream)
$command = "QUIT" 
$writer.WriteLine($command) 
$writer.Flush()
start-sleep -m 500 
$qrQueryresults = readResponse($stream)
$writer.Close() 
$stream.Close()
showresult($qrQueryresults)
}
}

function whoischk{

$whoisserver = "whois.arin.net"
$port = 43 
$socket = new-object System.Net.Sockets.TcpClient($whoisserver, $port) 
if($socket -eq $null) { showresult("Could not establish connection on Port 43") } 
else{
$stream = $socket.GetStream() 
$writer = new-object System.IO.StreamWriter($stream) 
$buffer = new-object System.Byte[] 1024 
$encoding = new-object System.Text.AsciiEncoding 
readResponse($stream)
$command = $logTable.DefaultView[$dgDataGrid.CurrentCell.RowNumber][2]
$writer.WriteLine($command) 
$writer.Flush()
start-sleep -m 500 
$qrQueryresults = readResponse($stream)
showresult($qrQueryresults)
}


}
function Compile-Csharp ([string] $code, [Array]$References) {

# Get an instance of the CSharp code provider
$cp = New-Object Microsoft.CSharp.CSharpCodeProvider

$refs = New-Object Collections.ArrayList
$refs.AddRange( @("${framework}System.dll",
# "${PsHome}\System.Management.Automation.dll",
# "${PsHome}\Microsoft.PowerShell.ConsoleHost.dll",
"${framework}System.Windows.Forms.dll",
"${framework}System.Data.dll",
"${framework}System.Drawing.dll",
"${framework}System.XML.dll"))
if ($References.Count -ge 1) {
$refs.AddRange($References)
}

# Build up a compiler params object...
$cpar = New-Object System.CodeDom.Compiler.CompilerParameters
$cpar.GenerateInMemory = $true
$cpar.GenerateExecutable = $false
$cpar.IncludeDebugInformation = $false
$cpar.CompilerOptions = "/target:library"
$cpar.ReferencedAssemblies.AddRange($refs)
$cr = $cp.CompileAssemblyFromSource($cpar, $code)

if ( $cr.Errors.Count) {
$codeLines = $code.Split("`n");
foreach ($ce in $cr.Errors) {
write-host "Error: $($codeLines[$($ce.Line - 1)])"
$ce | out-default
}
Throw "INVALID DATA: Errors encountered while compiling code"
}
}

$code = @'
namespace PAB.DnsUtils
{
    using System;
    using System.Collections;
    using System.ComponentModel;
    using System.Runtime.InteropServices;
    public class Dns 
        { 
        public Dns() 
        { 
        } 

        [DllImport("Dnsapi", EntryPoint="DnsQuery_W", CharSet=CharSet.Unicode, SetLastError=true, ExactSpelling=true)] 
        private static extern Int32 DnsQuery([MarshalAs(UnmanagedType.VBByRefStr)]ref string sName, QueryTypes wType, QueryOptions options, UInt32 aipServers, ref IntPtr ppQueryResults, UInt32 pReserved); 
        [DllImport("Dnsapi", CharSet=CharSet.Auto, SetLastError=true)] 
        private static extern void DnsRecordListFree(IntPtr pRecordList, int FreeType); 

        public enum ErrorReturnCode 
           { 
            DNS_ERROR_RCODE_NO_ERROR = 0, 
            DNS_ERROR_RCODE_FORMAT_ERROR = 9001, 
            DNS_ERROR_RCODE_SERVER_FAILURE = 9002, 
            DNS_ERROR_RCODE_NAME_ERROR = 9003, 
            DNS_ERROR_RCODE_NOT_IMPLEMENTED = 9004, 
            DNS_ERROR_RCODE_REFUSED = 9005, 
            DNS_ERROR_RCODE_YXDOMAIN = 9006, 
            DNS_ERROR_RCODE_YXRRSET = 9007, 
            DNS_ERROR_RCODE_NXRRSET = 9008, 
            DNS_ERROR_RCODE_NOTAUTH = 9009, 
            DNS_ERROR_RCODE_NOTZONE = 9010, 
            DNS_ERROR_RCODE_BADSIG = 9016, 
            DNS_ERROR_RCODE_BADKEY = 9017, 
            DNS_ERROR_RCODE_BADTIME = 9018 
            } 

            private enum QueryOptions 
            { 
            DNS_QUERY_ACCEPT_TRUNCATED_RESPONSE = 1, 
            DNS_QUERY_BYPASS_CACHE = 8, 
            DNS_QUERY_DONT_RESET_TTL_VALUES = 0x100000, 
            DNS_QUERY_NO_HOSTS_FILE = 0x40, 
            DNS_QUERY_NO_LOCAL_NAME = 0x20, 
            DNS_QUERY_NO_NETBT = 0x80, 
            DNS_QUERY_NO_RECURSION = 4, 
            DNS_QUERY_NO_WIRE_QUERY = 0x10, 
            DNS_QUERY_RESERVED = -16777216, 
            DNS_QUERY_RETURN_MESSAGE = 0x200, 
            DNS_QUERY_STANDARD = 0, 
            DNS_QUERY_TREAT_AS_FQDN = 0x1000, 
            DNS_QUERY_USE_TCP_ONLY = 2, 
            DNS_QUERY_WIRE_ONLY = 0x100 
            } 

            public enum QueryTypes 
            { 
            DNS_TYPE_A = 1, 
            DNS_TYPE_CNAME = 5, 
            DNS_TYPE_MX = 15, 
            DNS_TYPE_TEXT = 16, 
            DNS_TYPE_SRV = 33, 
            DNS_TYPE_PTR = 12

            } 

            [StructLayout(LayoutKind.Explicit)] 
            private struct DnsRecord 
            { 
            [FieldOffset(0)] 
            public IntPtr pNext; 
            [FieldOffset(4)] 
            public string pName; 
            [FieldOffset(8)] 
            public short wType; 
            [FieldOffset(10)] 
            public short wDataLength; 
            [FieldOffset(12)] 
            public uint flags; 
            [FieldOffset(16)] 
            public uint dwTtl; 
            [FieldOffset(20)] 
            public uint dwReserved; 

            // below is a partial list of the unionized members for this struct 

            // for DNS_TYPE_A records 
            [FieldOffset(24)] 
            public uint a_IpAddress; 

            // for DNS_TYPE_ PTR, CNAME, NS, MB, MD, MF, MG, MR records 
            [FieldOffset(24)] 
            public IntPtr ptr_pNameHost; 

            // for DNS_TXT_ DATA, HINFO, ISDN, TXT, X25 records 
            [FieldOffset(24)] 
            public uint data_dwStringCount; 
            [FieldOffset(28)] 
            public IntPtr data_pStringArray; 

            // for DNS_TYPE_MX records 
            [FieldOffset(24)] 
            public IntPtr mx_pNameExchange; 
            [FieldOffset(28)] 
            public short mx_wPreference; 
            [FieldOffset(30)] 
            public short mx_Pad; 

            // for DNS_TYPE_SRV records 
            [FieldOffset(24)] 
            public IntPtr srv_pNameTarget; 
            [FieldOffset(28)] 
            public short srv_wPriority; 
            [FieldOffset(30)] 
            public short srv_wWeight; 
            [FieldOffset(32)] 
            public short srv_wPort; 
            [FieldOffset(34)] 
            public short srv_Pad; 

            } 

            public static string[] GetRecords(string domain, string dnsqtype) 
            { 
            IntPtr ptr1 = IntPtr.Zero ; 
            IntPtr ptr2 = IntPtr.Zero ;
            DnsRecord rec;
            Dns.QueryTypes qtype = QueryTypes.DNS_TYPE_PTR;
            switch(dnsqtype){
                case "MX":
                    qtype = QueryTypes.DNS_TYPE_MX;
                    break;
                case "PTR":
                    qtype = QueryTypes.DNS_TYPE_PTR;
                    break;
                case "SPF":
                    qtype = QueryTypes.DNS_TYPE_TEXT;
                    break;
                case "A":
                    qtype = QueryTypes.DNS_TYPE_A;
                    break;
                case "RBL":
                    qtype = QueryTypes.DNS_TYPE_A;
                    break;
                case "MULTIRBL":
                    qtype = QueryTypes.DNS_TYPE_A;
                    break;
		case "SMTPTEST":
                    qtype = QueryTypes.DNS_TYPE_MX;
                    break;
		
            }
           
            if(Environment.OSVersion.Platform != PlatformID.Win32NT) 
            { 
            throw new NotSupportedException(); 
            } 

            ArrayList list1 = new ArrayList(); 
            int num1 = DnsQuery(ref domain, qtype, QueryOptions.DNS_QUERY_USE_TCP_ONLY|QueryOptions.DNS_QUERY_BYPASS_CACHE, 0, ref ptr1, 0); 
            if (num1 != 0) 
            {
                if (num1 == 9003 | num1 == 9501)
                {
                    String[] emErrormessage = new string[1];
                    emErrormessage.SetValue("No Record Found",0);
                    return emErrormessage;
                }
                else
                {
                    String[] emErrormessage = new string[1];
                    emErrormessage.SetValue("Error During Query Error Number " + num1 , 0);
                    return emErrormessage;  
                } 
            } 
            for (ptr2 = ptr1; !ptr2.Equals(IntPtr.Zero); ptr2 = rec.pNext) 
            { 
            rec = (DnsRecord) Marshal.PtrToStructure(ptr2, typeof(DnsRecord)); 
            if (rec.wType == (short)qtype) 
            { 
            string text1 = String.Empty; 
            switch(qtype) 
            { 
            case Dns.QueryTypes.DNS_TYPE_A: 
                System.Net.IPAddress ip = new System.Net.IPAddress(rec.a_IpAddress); 
                text1 = ip.ToString(); 
                break; 
                case Dns.QueryTypes.DNS_TYPE_CNAME: 
                text1 = Marshal.PtrToStringAuto(rec.ptr_pNameHost); 
                break; 
                case Dns.QueryTypes.DNS_TYPE_MX: 
                text1 = Marshal.PtrToStringAuto(rec.mx_pNameExchange);
		if (dnsqtype == "MX") {
			string[] mxalookup = PAB.DnsUtils.Dns.GetRecords(Marshal.PtrToStringAuto(rec.mx_pNameExchange), "A");
			text1 = text1 + " : " + rec.mx_wPreference.ToString()  + " : " ;
			foreach (string st in mxalookup)
			{
			text1 = text1 + st.ToString() + " ";
			}}               
                break; 
           case Dns.QueryTypes.DNS_TYPE_SRV: 
                text1 = Marshal.PtrToStringAuto(rec.srv_pNameTarget); 
                break; 
           case Dns.QueryTypes.DNS_TYPE_PTR:
                text1 = Marshal.PtrToStringAuto(rec.ptr_pNameHost);
                break;
           case Dns.QueryTypes.DNS_TYPE_TEXT:
                    if (Marshal.PtrToStringAuto(rec.data_pStringArray).ToLower().IndexOf("v=spf") == 0)
                    {
                        text1 = Marshal.PtrToStringAuto(rec.data_pStringArray);
                    }
                break; 
            default: 
            continue; 
            } 
            list1.Add(text1); 
            } 
            } 

            DnsRecordListFree(ptr2, 0); 
            return (string[]) list1.ToArray(typeof(string)); 
            } 
            } 
} 

'@


function openLog{
$logTable.clear()
$exFileName = new-object System.Windows.Forms.openFileDialog
$exFileName.ShowDialog()
$fnFileNamelableBox.Text = $exFileName.FileName
$tcountline = -1
if ($rbVeiwAllOnlyRadioButton.Checked -eq $true){$tcountline = $lnLogfileLineNum.value}
get-content $exFileName.FileName -totalCount $tcountline | %{ 
	$linarr = $_.split(" ")
	$lfDate = ""
	$lfTime = ""
	$lfSourceIP = ""
	$lfHostName = ""
	$lfDestIP = ""
	$lfSMTPVerb = ""
	$lfCommandText = ""
	if ($linarr[0].substring(0, 1) -ne "#"){
		 if ($linarr.Length -gt 0){$lfDate = $linarr[0]}
		 if ($linarr.Length -gt 1){$lfTime = $linarr[1]}
		 if ($linarr.Length -gt 2){$lfSourceIP= $linarr[2]}
		 if ($linarr.Length -gt 3){$lfHostName = $linarr[3]}
		 if ($linarr.Length -gt 6){$lfDestIP = $linarr[6]}
		 if ($linarr.Length -gt 8){$lfSMTPVerb = $linarr[8]}
		 if ($linarr.Length -gt 10){$lfCommandText = $linarr[10]}	
		 $logTable.Rows.Add($lfDate,$lfTime,$lfSourceIP,$lfHostName,$lfDestIP,$lfSMTPVerb,$lfCommandText)
	}
}

$dgDataGrid.DataSource = $logTable
}

Compile-Csharp $code
$form = new-object System.Windows.Forms.form 
$form.Text = "SMTP Log Test Tool"
$Dataset = New-Object System.Data.DataSet
$logTable = New-Object System.Data.DataTable
$logTable.TableName = "SMTPLogs"
$logTable.Columns.Add("Date");
$logTable.Columns.Add("Time");
$logTable.Columns.Add("SourceIPAddress");
$logTable.Columns.Add("HostName");
$logTable.Columns.Add("DestIPAddress");
$logTable.Columns.Add("SMTPVerb");
$logTable.Columns.Add("CommandText");

# Content
$cmClickMenu = new-object System.Windows.Forms.ContextMenuStrip
$cmClickMenu.Items.add("test122")

# Add Open Log file Button

$olButton = new-object System.Windows.Forms.Button
$olButton.Location = new-object System.Drawing.Size(20,19)
$olButton.Size = new-object System.Drawing.Size(75,23)
$olButton.Text = "Select file"
$olButton.Add_Click({openLog})
$form.Controls.Add($olButton)

# Add Reverse DNS Lookup Button

$rdnsbutton = new-object System.Windows.Forms.Button
$rdnsbutton.Location = new-object System.Drawing.Size(500,19)
$rdnsbutton.Size = new-object System.Drawing.Size(85,23)
$rdnsbutton.Text = "Reverse DNS"
$rdnsbutton.Add_Click({ptrLookup})
$form.Controls.Add($rdnsbutton)

# Add MX Lookup Button

$mxbutton = new-object System.Windows.Forms.Button
$mxbutton.Location = new-object System.Drawing.Size(500,44)
$mxbutton.Size = new-object System.Drawing.Size(85,23)
$mxbutton.Text = "MX Lookup"
$mxbutton.Add_Click({mxLookup})
$form.Controls.Add($mxbutton)


# Add SPF Lookup Button

$spfbutton = new-object System.Windows.Forms.Button
$spfbutton.Location = new-object System.Drawing.Size(590,19)
$spfbutton.Size = new-object System.Drawing.Size(85,23)
$spfbutton.Text = "SPF Lookup"
$spfbutton.Add_Click({spfLookup})
$form.Controls.Add($spfbutton)

# Add A Lookup Button

$Abutton = new-object System.Windows.Forms.Button
$Abutton.Location = new-object System.Drawing.Size(590,44)
$Abutton.Size = new-object System.Drawing.Size(85,23)
$Abutton.Text = "A Rec Lookup"
$Abutton.Add_Click({ALookup})
$form.Controls.Add($Abutton)

# Add RBL Single Lookup Button

$RBLButton = new-object System.Windows.Forms.Button
$RBLButton.Location = new-object System.Drawing.Size(680,19)
$RBLButton.Size = new-object System.Drawing.Size(85,23)
$RBLButton.Text = "RBL Lookup"
$RBLButton.Add_Click({RBLLookup})
$form.Controls.Add($RBLButton)

# Add HELO chk Button

$HelochkButton = new-object System.Windows.Forms.Button
$HelochkButton.Location = new-object System.Drawing.Size(680,44)
$HelochkButton.Size = new-object System.Drawing.Size(85,23)
$HelochkButton.Text = "HELO Banner Check"
$HelochkButton.Add_Click({HELOChk})
$form.Controls.Add($HelochkButton)

# Add WhoIS chk Button

$Whois = new-object System.Windows.Forms.Button
$Whois.Location = new-object System.Drawing.Size(770,19)
$Whois.Size = new-object System.Drawing.Size(85,23)
$Whois.Text = "Whois Check"
$Whois.Add_Click({whoischk})
$form.Controls.Add($Whois)


# Add FileName Lable
$fnFileNamelableBox = new-object System.Windows.Forms.Label
$fnFileNamelableBox.Location = new-object System.Drawing.Size(110,25)
$fnFileNamelableBox.forecolor = "MenuHighlight"
$fnFileNamelableBox.size = new-object System.Drawing.Size(200,20) 
$form.Controls.Add($fnFileNamelableBox) 

# Add Veiw RadioButtons
$rbVeiwAllRadioButton = new-object System.Windows.Forms.RadioButton
$rbVeiwAllRadioButton.Location = new-object System.Drawing.Size(310,19)
$rbVeiwAllRadioButton.size = new-object System.Drawing.Size(69,17) 
$rbVeiwAllRadioButton.Checked = $true
$rbVeiwAllRadioButton.Text = "View All"
$rbVeiwAllRadioButton.Add_Click({if ($rbVeiwAllRadioButton.Checked -eq $true){$lnLogfileLineNum.Enabled = $false}})
$form.Controls.Add($rbVeiwAllRadioButton) 

$rbVeiwAllOnlyRadioButton = new-object System.Windows.Forms.RadioButton
$rbVeiwAllOnlyRadioButton.Location = new-object System.Drawing.Size(310,42)
$rbVeiwAllOnlyRadioButton.size = new-object System.Drawing.Size(89,17) 
$rbVeiwAllOnlyRadioButton.Text = "View Only #"
$rbVeiwAllOnlyRadioButton.Add_Click({if ($rbVeiwAllOnlyRadioButton.Checked -eq $true){$lnLogfileLineNum.Enabled = $true}})
$form.Controls.Add($rbVeiwAllOnlyRadioButton) 

# Add Numeric log line number 
$lnLogfileLineNum =  new-object System.Windows.Forms.numericUpDown
$lnLogfileLineNum.Location = new-object System.Drawing.Size(401,39)
$lnLogfileLineNum.Size = new-object System.Drawing.Size(69,20)
$lnLogfileLineNum.Enabled = $false
$lnLogfileLineNum.Maximum = 10000000000
$form.Controls.Add($lnLogfileLineNum)


# File setting Group Box

$OfGbox =  new-object System.Windows.Forms.GroupBox
$OfGbox.Location = new-object System.Drawing.Size(12,0)
$OfGbox.Size = new-object System.Drawing.Size(464,75)
$OfGbox.Text = "Log File Settings"
$form.Controls.Add($OfGbox)

# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGrid
$dgDataGrid.AllowSorting = $True
#$dgDataGrid.Add_Click({MXlookup($logTable.DefaultView[$dgDataGrid.CurrentCell.RowNumber])})
$dgDataGrid.Location = new-object System.Drawing.Size(12,81) 
$dgDataGrid.size = new-object System.Drawing.Size(1024,750) 
$form.Controls.Add($dgDataGrid)

#
$form.topmost = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()



'==========================================================================
'
' NAME: RemoveNetPrinters.vbs
' COMMENT: Removes and add all net printers
'
'==========================================================================
ON ERROR RESUME NEXT
Set wshNet = CreateObject("WScript.Network")
Set wshPrn = wshNet.EnumPrinterConnections
For x = 0 To wshPrn.Count - 1 Step 2
If Left(wshPrn.Item(x+1),2) = "\\" Then wshNet.RemovePrinterConnection wshPrn.Item(x+1),True,True
Next

wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\NVCI-NursePod-2335dn"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\NVCI-Doctor-2330dn"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\NVCI-CheckIn-2335dn"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\RR-Scheduling-XeroxWorkCenter7435"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\RR-Authorizations-2335dn"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\RR-NurseDesk-2330dn"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\RR-Triage-HP1320n"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\RR-Billing-1320c"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\RR-CheckOut-2330dn"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\RR-NurseStation-2330dn"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\RR-CheckIn-2350dn"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\RR-Scheduling-2335dn-Secondary"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\RR-Peds-HP1320n"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\RR-SurgeryScheduling-2335dn"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\RR-Billing-HICFA"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\RR-Peds-2335dn"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\RR-Scheduling-2335dn--Primary"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\RR-Billing-1815dn"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\RR-MedicalRecords-23335nd"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\RR-Billing-2330dn"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\RR-ChartCheck-1815dn"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\GV-NurseDesk-1135n"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\GV-BackOffice-HP1320n"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\GV-CheckIn-HP1320n"
wshnet.AddWindowsPrinterConnection "\\USONPSVRFPF\GV-CheckIn-2335dn"

'==========================================================================
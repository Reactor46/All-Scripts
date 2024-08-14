Option Explicit
Dim objNetwork, strDefaultPrinter, strPrinter1,strPrinter2
Dim strPrinter3,strPrinter4,strPrinter5,strPrinter6,strPrinter7

strPrinter1 = "\\servermain\Billing_206"
strPrinter2 = "\\server\printershare"
strPrinter3 = "\\server\printershare"
strPrinter4 = "\\server\printershare"
strPrinter5 = "\\server\printershare"
strPrinter6 = "\\server\printershare"

strDefaultPrinter =strPrinter1

Set objNetwork = CreateObject("WScript.Network") 
objNetwork.AddWindowsPrinterConnection strPrinter1
objNetwork.AddWindowsPrinterConnection strPrinter2
objNetwork.AddWindowsPrinterConnection strPrinter3
objNetwork.AddWindowsPrinterConnection strPrinter4
objNetwork.AddWindowsPrinterConnection strPrinter5
objNetwork.AddWindowsPrinterConnection strPrinter6

' Set Default Printer
objNetwork.SetDefaultPrinter strDefaultPrinter
WScript.Echo "Check the Printers folder for : " &  strDefaultPrinter

WScript.Quit
'==========================================================================
' NAME: RebootPending.vbs
' AUTHOR: Matthias CECILLON
' DATE  : 15/07/2014
' COMMENT: 
' This is script is designed for use in a 2-state monitor in Operations Manager 2012. Its function is to
' verify the value registry key in the Windows Registry. 
'==========================================================================

Dim SearchKey, KeyFound
Dim oAPI, oBag
Set oAPI = CreateObject("MOM.ScriptAPI")
Set oBag = oAPI.CreatePropertyBag()

ssig="Unable to open registry key"

set wshShell= Wscript.CreateObject("WScript.Shell")
SearchKey = "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations"
on error resume next
present = WshShell.RegRead(SearchKey)
if err.number<>0 then
    if right(SearchKey,1)="\" then    'SearchKey is a registry key
        if instr(1,err.description,ssig,1)<>0 then
            KeyFound=true
        else
            KeyFound=false
        end if
    else    'SearchKey is a registry valuename
        KeyFound=false
    end if
    err.clear
else
    KeyFound=true
end if
on error goto 0
if KeyFound=vbFalse then
    wscript.echo SearchKey & " does not exist."
    Call oBag.AddValue("State","OK") 
    Call oAPI.Return(oBag)
else
    wscript.echo SearchKey & " exists."
    Call oBag.AddValue("State","KO")
    Call oAPI.Return(oBag)
end if
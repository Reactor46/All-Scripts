homepage = "http://www.google.com.au/search?hl=en&q=coffee+tasting&btnG=Google+Search&meta=cr%3DcountryAU"
dwVersion = "02"
dwType = "00000001"
dwFlags = "00000001"
dwUnused = "00000000000000000000000000000000000000000000000000000000"
bData = AsciiToHex(homepage)
cbDataSize = cstr(ubound(bData)+1)
propval = dwVersion & dwType & dwFlags & dwUnused & "000000" & Hex(cbDataSize) & "000000" & Join(bData,"")

set rec = createobject("ADODB.Record")
rec.open "file://./backofficestorage/domain.com/MBX/mailbox/webfolder/", ,3,33562624 
rec.fields("http://schemas.microsoft.com/mapi/proptag/0x36DF0102").value = propval
rec.fields.update
rec.close


Function AsciiToHex(sData)
 Dim i, aTmp()
 ReDim aTmp((Len(sData)*2) + 1)
 arnum = 0 
 For i = 1 To Len(sData)
  aTmp(arnum) = Hex(Asc(Mid(sData, i)))
  arnum = arnum + 1
  aTmp(arnum) = "00" 
  arnum = arnum + 1
 Next
  aTmp(arnum) = "00" 
  arnum = arnum + 1
  aTmp(arnum) = "00"
 ASCIItoHex = aTmp
End Function
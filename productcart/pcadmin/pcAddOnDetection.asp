<%

'//Check for MailUp: no longer necessary: built into v4.5

'//Check for Google Analytics: no longer necessary: built into v4.5

'//Check for eBAY
Set fs=Server.CreateObject("Scripting.FileSystemObject")
If (fs.FileExists(Server.MapPath("ebay_Listings.asp")))=0 Then
   isEbayApplied="0"
Else
   isEbayApplied="1"
End If
set fs=nothing

'//Check for Mobile Commerce (since statusM.inc was not part of the general build before v4.5)
isMobileComApplied="0"
if PPD="1" then
	pcStrFolder=Server.Mappath ("/"&scPcFolder&"/m")
else
	pcStrFolder=server.MapPath("../m")
end if	
Set fs=Server.CreateObject("Scripting.FileSystemObject")
If (fs.FileExists(pcStrFolder & "\checkout.asp"))=0 Then
   isMobileComApplied="0"
Else
   isMobileComApplied="1"
End If
set fs=nothing
%>
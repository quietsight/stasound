<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/stringfunctions.asp"-->
<% 
'/////////////////////////////////////////////////////
'// Write all changes to Settings.asp file
'/////////////////////////////////////////////////////
Dim objFS
Dim objFile

Set objFS = Server.CreateObject ("Scripting.FileSystemObject")
if PPD="1" then
	pcStrFileName=Server.Mappath ("/"&scPcFolder&"/includes/GCConstants.asp")
else
	pcStrFileName=Server.Mappath ("../includes/GCConstants.asp")
end if

If pcGCIncludeShipping="" then
	pcGCIncludeShipping = "0"
End If

Set objFile = objFS.OpenTextFile (pcStrFileName, 2, True, 0)
objFile.WriteLine CHR(60)&CHR(37)&"'// Gift Certificate Constants //" & vbCrLf
objFile.WriteLine "private const GC_INCSHIPPING = """&pcGCIncludeShipping&"""" & vbCrLf
objFile.WriteLine "'// Gift Certificate Constants // " &CHR(37)&CHR(62)& vbCrLf
objFile.Close

set objFS=nothing
set objFile=nothing

%>
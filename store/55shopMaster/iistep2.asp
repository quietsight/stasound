<%@ LANGUAGE="VBSCRIPT" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/pcSanitizeUpload.asp"-->
<%
Server.ScriptTimeout = 5400
Response.Expires=0
response.Buffer=true
Response.Clear
byteCount=Request.TotalBytes
RequestBin=Request.BinaryRead(byteCount)
'//DETAILED MSG
dim pcTmpErr, pcTmpErrSize
pcTmpErr = Cstr("")
pcTmpErrSize = Cint(0)

pcTmpErr = err.description

If pcTmpErr&""<>"" Then
	If instr(pcTmpErr, "007") Then
		pcTmpErrSize = 1
	End If
End If
'//END DETAILED MSG

Dim UploadRequest
Set UploadRequest=CreateObject("Scripting.Dictionary")
BuildUploadRequest  RequestBin
Dim InValidCSV 
InValidCSV=0

filepathname=UploadRequest.Item("file1").Item("FileName") 
contentType=Right(filepathname,4)
if (right(ucase(filepathname),4)=".XLS") AND IsUploadAllowed(filepathname) then
	if PPD="1" then
		filename="/"&scPcFolder&"/pc/catalog" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	else
		filename="../pc/catalog" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	if not filename="" then 
	value=UploadRequest.Item("file1").Item("Value")
	Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
	Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
	For i=1 to LenB(value)
	MyFile.Write chr(AscB(MidB(value,i,1)))
	Next
	MyFile.Close
	set myfile=nothing

	Sub BuildUploadRequest(RequestBin)
		PosBeg=1
		PosEnd=InstrB(PosBeg,RequestBin,getByteString(chr(13)))
		boundary=MidB(RequestBin,PosBeg,PosEnd-PosBeg)
		boundaryPos=InstrB(1,RequestBin,boundary)
		Do until (boundaryPos=InstrB(RequestBin,boundary & getByteString("--")))
		Dim UploadControl
		Set UploadControl=CreateObject("Scripting.Dictionary")
		Pos=InstrB(BoundaryPos,RequestBin,getByteString("Content-Disposition"))
		Pos=InstrB(Pos,RequestBin,getByteString("name="))
		PosBeg=Pos+6
		PosEnd=InstrB(PosBeg,RequestBin,getByteString(chr(34)))
		Name=getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
		PosFile=InstrB(BoundaryPos,RequestBin,getByteString("filename="))
		PosBound=InstrB(PosEnd,RequestBin,boundary)
		If  PosFile<>0 AND (PosFile<PosBound) Then
				PosBeg=PosFile + 10
				PosEnd=InstrB(PosBeg,RequestBin,getByteString(chr(34)))
				FileName=getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
				UploadControl.Add "FileName", FileName
				Pos=InstrB(PosEnd,RequestBin,getByteString("Content-Type:"))
				PosBeg=Pos+14
				PosEnd=InstrB(PosBeg,RequestBin,getByteString(chr(13)))
				ContentType=getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
				UploadControl.Add "ContentType",ContentType
				PosBeg=PosEnd+4
				PosEnd=InstrB(PosBeg,RequestBin,boundary)-2
				Value=MidB(RequestBin,PosBeg,PosEnd-PosBeg)
			Else
				Pos=InstrB(Pos,RequestBin,getByteString(chr(13)))
				PosBeg=Pos+4
				PosEnd=InstrB(PosBeg,RequestBin,boundary)-2
				Value=getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			End If
			UploadControl.Add "Value" , Value	
			UploadRequest.Add name, UploadControl	
			BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
			Loop
		
		End Sub

		Function getByteString(StringStr)
		 For i=1 to Len(StringStr)
			char=Mid(StringStr,i,1)
			getByteString=getByteString & chrB(AscB(char))
		 Next
		End Function

		Function getString(StringBin)
		 getString=""
		 For intCount=1 to LenB(StringBin)
			getString=getString & chr(AscB(MidB(StringBin,intCount,1))) 
		 Next
		End Function

		lgCSV=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		session("importfile")=lgCSV
		
		Response.redirect "iistep1.asp?a=next&s=1&msg=" & Server.URLEncode("The product data file " & ucase(lgCSV) & " was uploaded successfully.")
	end if
else
	if UploadRequest.Item("file1").Item("FileName")<>"" then
		InValidCSV=InValidCSV+1
	end if
end if

If InValidCSV>0 then
	'//DETAILED MSG
	if pcTmpErrSize = 1 then
		pcTmpErrSize = 0
		
		response.write "An Error occurred while attempting to upload your images: "&err.description&"<br><br>This is probably due to a default configuration option in IIS, which only allows the entity size in a POST request to be 200,000 bytes (~200KB).<br><br>You have to change this setting if you want to upload multiple images files or large image files through the ProductCart application (otherwise you can just FTP the file to the same location: e.g. the &quot;pc/catalog&quot; folder).<br><br>To change this setting:<br><br>- Open IIS Manager<br>- Navigate the tree to your application<br>- Double click the &quot;ASP&quot; icon in the main panel<br>- Expand the &quot;Limits&quot; category<br>- Modify the &quot;Maximum Requesting Entity Body Limit&quot; to a larger value."
	else
		response.redirect "iistep1.asp?s=0&msg="&Server.URLEncode("Invalid file type. Only XLS files can be uploaded to the server.")
	end if
	'//END DETAILED MSG
Else
	response.redirect "iistep1.asp?s=0&msg="&Server.URLEncode("You did not select a file to upload.")
end if
%>
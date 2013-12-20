<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"-->
<!--#include file="../../includes/ppdstatus.inc"-->
<!--#include file="../../includes/productcartFolder.asp"-->
<!--#include file="../../includes/stringfunctions.asp"-->
<!--#include file="../../includes/pcSanitizeUpload.asp"-->

<%
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
Set UploadRequest=server.CreateObject("Scripting.Dictionary")
BuildUploadRequest  RequestBin
Dim invalidTextFile, fileCnt
invalidTextFile=0
fileCnt=0

'=======
'file 1 
'=======
contentType=UploadRequest.Item("one").Item("ContentType")
filepathname=UploadRequest.Item("one").Item("FileName")

if instr(ucase(contentType),"TEXT") AND IsUploadAllowed(filepathname) then
	fileCnt=1
	if filepathname="" Then
		one=""
	Else
    
	If LCase(right(filepathname,3)) <> "txt" Then
       invalidTextFile = 1
    else
        if PPD="1" then
            filename="/"&scPcFolder&"/pc/library" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
        else
            filename="../../pc/library" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
        end If
        if not filename="" then 
            value=UploadRequest.Item("one").Item("Value")
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
                    Set UploadControl=server.CreateObject("Scripting.Dictionary")
                    
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
            File1=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
        end if
        response.write "library/"&one&"<br>"
        end If
    End if
else
	if UploadRequest.Item("one").Item("FileName")<>"" then
		fileCnt=1
		invalidTextFile=1
	end if
end if


If invalidTextFile>0 then
	'//DETAILED MSG
	if pcTmpErrSize = 1 then
		pcTmpErrSize = 0
		
		response.write "An Error occurred while attempting to upload your images: "&err.description&"<br><br>This is probably due to a default configuration option in IIS, which only allows the entity size in a POST request to be 200,000 bytes (~200KB).<br><br>You have to change this setting if you want to upload multiple images files or large image files through the ProductCart application (otherwise you can just FTP the file to the same location: e.g. the &quot;pc/catalog&quot; folder).<br><br>To change this setting:<br><br>- Open IIS Manager<br>- Navigate the tree to your application<br>- Double click the &quot;ASP&quot; icon in the main panel<br>- Expand the &quot;Limits&quot; category<br>- Modify the &quot;Maximum Requesting Entity Body Limit&quot; to a larger value."
	else
		response.redirect "FileUploada.asp?msg="&Server.URLEncode(invalidTextFile&" is not a valid TEXT file (with extension .TXT) and could not be uploaded to the server.")
	end if
	'//END DETAILED MSG
Else
	if fileCnt>0 Then
		response.redirect "FileUploada.asp?s=1&f=" & File1 & "&msg="&Server.URLEncode("Your file was successfully uploaded.")
	else
		response.redirect "FileUploada.asp?msg="&Server.URLEncode("You need to supply a file to upload.")
	end if
end if
%>







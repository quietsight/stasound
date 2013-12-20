<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin="1*6*"%><!--#include file="adminv.asp"-->
<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/pcSanitizeUpload.asp"-->

<%
Response.Expires=0
response.Buffer=true
Response.Clear
byteCount=Request.TotalBytes
RequestBin=Request.BinaryRead(byteCount)
Dim UploadRequest
Set UploadRequest=CreateObject("Scripting.Dictionary")
BuildUploadRequest  RequestBin

'=======
'file 1 
'=======
contentType=UploadRequest.Item("one").Item("ContentType")
filepathname=UploadRequest.Item("one").Item("FileName")

if (instr(ucase(contentType),"APPLICATION") OR instr(ucase(contentType),"TEXT")) AND IsUploadAllowed(filepathname) then
	if instr(ucase(filepathname),".CSV") then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
			one=""
		else
			if PPD="1" then
				filename="/"&scPcFolder&"/pc/tax" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
			else
				filename="../pc/tax" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
			end if
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
				File1=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
			end if
		end if
	else
		if UploadRequest.Item("one").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if
else
	if UploadRequest.Item("one").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

If InValidImage>0 then
	response.write "<br><div align=center><font face=arial size=2>Your file does not appear to be in the correct format. <br>Invalid  formats are not allowed to be uploaded to the server.<br><br><a href=""javascript:history.go(-1)""><font face="&Link&">Back</font></a><br><br>If you are certain that your file is of the proper format and you are receiving this error, you will need to manually upload your ""Tax Rate File"" file to your server using your ftp client.<br><br>The file needs to be uploaded to the folder:<br><br> <font color=""FF0000"">""/productcart/pc/tax/""</font></div><p align=""center""><input type=""button"" value=""Close Window"" onClick=""javascript:window.close();""></p></font>"
Else
	if ImageCnt>0 then
		response.redirect "taxupl_popup_confirm.html"
	else
		response.write "<br><div align=center><font face=arial size=2>You need to supply a file to upload.<br><br><a href=""javascript:history.go(-1)""><font face="&Link&">Back</font></a></font></div>"
	end if
end if
%>
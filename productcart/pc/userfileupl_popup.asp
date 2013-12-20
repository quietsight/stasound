<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>

<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/pcSanitizeUpload.asp"-->

<html>
<head>
<title>Upload Data File(s)</title>
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
</head>
<body>
<div id="pcMain">
	<table class="pcMainTable" cellpadding="4">
		<tr>
			<td>

<% dim mySQL, conntemp, rstemp
on error resume next
call openDB()
response.write"<b>Please wait while your files are being uploaded ...</b><br><br>"
Response.Expires=0
'response.Buffer=true
'Response.Clear
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
Dim InValidFile, FileCnt
InValidFile=0
FileCnt=0
TempName=month(now()) & day(now()) & year(now()) & hour(now()) & minute(now()) & second(now())


'=======
'file 1 
'=======
contentType=UploadRequest.Item("one").Item("ContentType")
filepathname=UploadRequest.Item("one").Item("FileName")
checkfile=0
if filepathname<>"" AND IsUploadAllowed(filepathname) then
extfile=Right(ucase(filepathname),4)
if (extfile=".TXT") or (extfile=".HTM") or (extfile=".GIF") or (extfile=".JPG") or (extfile=".PDF") or (extfile=".DOC") or (extfile=".ZIP") then
checkfile=1
else
extfile=Right(ucase(filepathname),5)
if (extfile=".HTML") then
checkfile=1
end if
end if
end if

if checkfile=1 then
	FileCnt=FileCnt+1
	filepathname=UploadRequest.Item("one").Item("FileName")
	if filepathname="" Then
	one=""
	else
	filename="Library" & "/" & TempName & "_" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
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
		File1=TempName & "_" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	if File1<>"" then
	MySQL="insert into pcUploadFiles (pcUpld_IDFeedback,pcUpld_FileName) values (" & session("UIDFeedback") & ",'" & File1 & "')"
	set rstemp=connTemp.execute(mySQL)
	end if
	end if
else
	if UploadRequest.Item("one").Item("FileName")<>"" then
		FileCnt=FileCnt+1
		InValidFile=InValidFile+1
	end if
end if

'=======
'file 2 
'=======
contentType=UploadRequest.Item("two").Item("ContentType")

filepathname=UploadRequest.Item("two").Item("FileName")
checkfile=0
if filepathname<>"" AND IsUploadAllowed(filepathname) then
extfile=Right(ucase(filepathname),4)
if (extfile=".TXT") or (extfile=".HTM") or (extfile=".GIF") or (extfile=".JPG") or (extfile=".PDF") or (extfile=".DOC") or (extfile=".ZIP") then
checkfile=1
else
extfile=Right(ucase(filepathname),5)
if (extfile=".HTML") then
checkfile=1
end if
end if
end if

if checkfile=1 then
	FileCnt=FileCnt+1
	filepathname=UploadRequest.Item("two").Item("FileName")
	if filepathname="" Then
	add=""
	else
	filename="Library" & "/" & TempName & "_" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	if not filename="" then 
	value=UploadRequest.Item("two").Item("Value")
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
	File2=TempName & "_" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	
	if File2<>"" then
	MySQL="insert into pcUploadFiles (pcUpld_IDFeedback,pcUpld_FileName) values (" & session("UIDFeedback") & ",'" & File2 & "')"
	set rstemp=connTemp.execute(mySQL)
	end if
	
	end if
else
	if UploadRequest.Item("two").Item("FileName")<>"" then
		FileCnt=FileCnt+1
		InValidFile=InValidFile+1
	end if
end if

'=======
'file 3 
'=======
contentType=UploadRequest.Item("three").Item("ContentType")

filepathname=UploadRequest.Item("three").Item("FileName")
checkfile=0
if filepathname<>"" AND IsUploadAllowed(filepathname) then
extfile=Right(ucase(filepathname),4)
if (extfile=".TXT") or (extfile=".HTM") or (extfile=".GIF") or (extfile=".JPG") or (extfile=".PDF") or (extfile=".DOC") or (extfile=".ZIP") then
checkfile=1
else
extfile=Right(ucase(filepathname),5)
if (extfile=".HTML") then
checkfile=1
end if
end if
end if



if checkfile=1 then
	FileCnt=FileCnt+1
	filepathname=UploadRequest.Item("three").Item("FileName")
	if filepathname="" Then
	file3=""
	else
	filename="Library" & "/" & TempName & "_" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	if not filename="" then 
	value=UploadRequest.Item("three").Item("Value")
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
	File3=TempName & "_" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	
	if File3<>"" then
	MySQL="insert into pcUploadFiles (pcUpld_IDFeedback,pcUpld_FileName) values (" & session("UIDFeedback") & ",'" & File3 & "')"
	set rstemp=connTemp.execute(mySQL)
	end if
	
	end if
else
	if UploadRequest.Item("three").Item("FileName")<>"" then
		FileCnt=FileCnt+1
		InValidFile=InValidFile+1
	end if
end if

'=======
'file 4 
'=======
contentType=UploadRequest.Item("four").Item("ContentType")

filepathname=UploadRequest.Item("four").Item("FileName")
checkfile=0
if filepathname<>"" AND IsUploadAllowed(filepathname) then
extfile=Right(ucase(filepathname),4)
if (extfile=".TXT") or (extfile=".HTM") or (extfile=".GIF") or (extfile=".JPG") or (extfile=".PDF") or (extfile=".DOC") or (extfile=".ZIP") then
checkfile=1
else
extfile=Right(ucase(filepathname),5)
if (extfile=".HTML") then
checkfile=1
end if
end if
end if



if checkfile=1 then
	FileCnt=FileCnt+1
	filepathname=UploadRequest.Item("four").Item("FileName")
	if filepathname="" Then
	file4=""
	else
	filename="Library" & "/" & TempName & "_" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	if not filename="" then 
	value=UploadRequest.Item("four").Item("Value")
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
	File4=TempName & "_" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	
	if File4<>"" then
	MySQL="insert into pcUploadFiles (pcUpld_IDFeedback,pcUpld_FileName) values (" & session("UIDFeedback") & ",'" & File4 & "')"
	set rstemp=connTemp.execute(mySQL)
	end if
	
	end if
else
	if UploadRequest.Item("four").Item("FileName")<>"" then
		FileCnt=FileCnt+1
		InValidFile=InValidFile+1
	end if
end if

'=======
'file 5 
'=======
contentType=UploadRequest.Item("five").Item("ContentType")

filepathname=UploadRequest.Item("five").Item("FileName")
checkfile=0
if filepathname<>"" AND IsUploadAllowed(filepathname) then
extfile=Right(ucase(filepathname),4)
if (extfile=".TXT") or (extfile=".HTM") or (extfile=".GIF") or (extfile=".JPG") or (extfile=".PDF") or (extfile=".DOC") or (extfile=".ZIP") then
checkfile=1
else
extfile=Right(ucase(filepathname),5)
if (extfile=".HTML") then
checkfile=1
end if
end if
end if



if checkfile=1 then
	FileCnt=FileCnt+1
	filepathname=UploadRequest.Item("five").Item("FileName")
	if filepathname="" Then
	file5=""
	else
	filename="Library" & "/" & TempName & "_" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	if not filename="" then 
	value=UploadRequest.Item("five").Item("Value")
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
	File5=TempName & "_" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	
	if File5<>"" then
	MySQL="insert into pcUploadFiles (pcUpld_IDFeedback,pcUpld_FileName) values (" & session("UIDFeedback") & ",'" & File5 & "')"
	set rstemp=connTemp.execute(mySQL)
	end if	
	
	end if
else
	if UploadRequest.Item("five").Item("FileName")<>"" then
		FileCnt=FileCnt+1
		InValidFile=InValidFile+1
	end if
end if

'=======
'file 6 
'=======
contentType=UploadRequest.Item("six").Item("ContentType")

filepathname=UploadRequest.Item("six").Item("FileName")
checkfile=0
if filepathname<>"" AND IsUploadAllowed(filepathname) then
extfile=Right(ucase(filepathname),4)
if (extfile=".TXT") or (extfile=".HTM") or (extfile=".GIF") or (extfile=".JPG") or (extfile=".PDF") or (extfile=".DOC") or (extfile=".ZIP") then
checkfile=1
else
extfile=Right(ucase(filepathname),5)
if (extfile=".HTML") then
checkfile=1
end if
end if
end if



if checkfile=1 then
	FileCnt=FileCnt+1
	filepathname=UploadRequest.Item("six").Item("FileName")
	if filepathname="" Then
	file6=""
	else
	filename="images" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	if not filename="" then 
	value=UploadRequest.Item("six").Item("Value")
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
	File6=TempName & "_" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	
	if File6<>"" then
	MySQL="insert into pcUploadFiles (pcUpld_IDFeedback,pcUpld_FileName) values (" & session("UIDFeedback") & ",'" & File6 & "')"
	set rstemp=connTemp.execute(mySQL)
	end if
	
	end if
else
	if UploadRequest.Item("six").Item("FileName")<>"" then
		FileCnt=FileCnt+1
		InValidFile=InValidFile+1
	end if
end if

If InValidFile>0 then
	'//DETAILED MSG
	if pcTmpErrSize = 1 then
		pcTmpErrSize = 0
		
		response.write "<div class=pcErrorMessage>An Error occurred while attempting to upload your images: <strong>"&err.description&"</strong></div><div>This is probably due to a default configuration option in IIS, which only allows the entity size in a POST request to be 200,000 bytes (~200KB).<br><br>You have to change this setting if you want to upload multiple images files or large image files through the ProductCart application (otherwise you can just FTP the file to the same location: e.g. the &quot;pc/catalog&quot; folder).<br><br>To change this setting:<br><br>- Open IIS Manager<br>- Navigate the tree to your application<br>- Double click the &quot;ASP&quot; icon in the main panel<br>- Expand the &quot;Limits&quot; category<br>- Modify the &quot;Maximum Requesting Entity Body Limit&quot; to a larger value<br><br><a href=""javascript:history.go(-1)"">Back</a></font></div>"
	else
		response.write "<br><div align=center>"&InValidFile&" of your "&FileCnt&" files were not a valid file type. <br>Invalid file types are not allowed to be uploaded to the server.<br><br><a href=""javascript:history.go(-1)"">Back</a></div>"
	end if
	'//END DETAILED MSG
Else
	if FileCnt>0 then
		session("uploaded")="1"%>
		<script>
		location="userfileupl_popup_confirm.asp";
		</script>
		<%
	else
		response.write "<br><div align=center><font face=arial size=2>You need to supply at least one file to upload.<br><br><a href=""javascript:history.go(-1)"">Back</a></font></div>"
	end if
end if
%>
</td>
</tr>
</table>
</div>
</body>
</html>
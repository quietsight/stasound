<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"-->
<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/pcSanitizeUpload.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
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
Dim InValidImage, ImageCnt
InValidImage=0
ImageCnt=0

'=======
'file 1 
'=======
contentType=UploadRequest.Item("erroricon").Item("ContentType")
filepathname=UploadRequest.Item("erroricon").Item("FileName")
	
if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
	ImageCnt=ImageCnt+1
	if filepathname="" Then
	file1=""
	else
	if PPD="1" then
		filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	else
		filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	if not filename="" then 
		value=UploadRequest.Item("erroricon").Item("Value")
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
		file1=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	end if
else
	if UploadRequest.Item("erroricon").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'=======
'file 2 
'=======
contentType=UploadRequest.Item("requiredicon").Item("ContentType")
filepathname=UploadRequest.Item("requiredicon").Item("FileName")
	
if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
	ImageCnt=ImageCnt+1
	if filepathname="" Then
	file2=""
	else
	if PPD="1" then
		filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	else
		filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	if not filename="" then 
	value=UploadRequest.Item("requiredicon").Item("Value")
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
	file2=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	response.write "images/"& file2 &"<br>"
	end if
else
	if UploadRequest.Item("requiredicon").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'=======
'file 3 
'=======
contentType=UploadRequest.Item("errorfieldicon").Item("ContentType")
filepathname=UploadRequest.Item("errorfieldicon").Item("FileName")
	
if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
	ImageCnt=ImageCnt+1
	if filepathname="" Then
	file3=""
	else
	if PPD="1" then
		filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	else
		filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	if not filename="" then 
	value=UploadRequest.Item("errorfieldicon").Item("Value")
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
	file3=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	response.write "images/"& file3 &"<br>"
	end if
else
	if UploadRequest.Item("errorfieldicon").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'=======
'file 4 
'=======
contentType=UploadRequest.Item("previousicon").Item("ContentType")
filepathname=UploadRequest.Item("previousicon").Item("FileName")
	
if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
	ImageCnt=ImageCnt+1
	if filepathname="" Then
	file4=""
	else
	if PPD="1" then
		filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	else
		filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	if not filename="" then 
	value=UploadRequest.Item("previousicon").Item("Value")
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
	file4=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	response.write "images/"& file4 &"<br>"
	end if
else
	if UploadRequest.Item("previousicon").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'=======
'file 5 
'=======
contentType=UploadRequest.Item("nexticon").Item("ContentType")
filepathname=UploadRequest.Item("nexticon").Item("FileName")
	
if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
	ImageCnt=ImageCnt+1
	if filepathname="" Then
	file5=""
	else
	if PPD="1" then
		filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	else
		filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	if not filename="" then 
	value=UploadRequest.Item("nexticon").Item("Value")
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
	file5=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	response.write "images/"& file5 &"<br>"
	end if
else
	if UploadRequest.Item("nexticon").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'=======
'file 6
'=======
contentType=UploadRequest.Item("zoom").Item("ContentType")
filepathname=UploadRequest.Item("zoom").Item("FileName")
	
if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
	ImageCnt=ImageCnt+1
	if filepathname="" Then
	file6=""
	else
	if PPD="1" then
		filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	else
		filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	if not filename="" then 
	value=UploadRequest.Item("zoom").Item("Value")
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
	file6=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	response.write "images/"& file6 &"<br>"
	end if
else
	if UploadRequest.Item("zoom").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'=======
'file 7
'=======
contentType=UploadRequest.Item("discount").Item("ContentType")
filepathname=UploadRequest.Item("discount").Item("FileName")
	
if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
	ImageCnt=ImageCnt+1
	if filepathname="" Then
	file7=""
	else
	if PPD="1" then
		filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	else
		filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	if not filename="" then 
	value=UploadRequest.Item("discount").Item("Value")
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
	file7=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	response.write "images/"& file7 &"<br>"
	end if
else
	if UploadRequest.Item("discount").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'=======
'file 8
'=======
contentType=UploadRequest.Item("arrowUp").Item("ContentType")
filepathname=UploadRequest.Item("arrowUp").Item("FileName")
	
if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
	ImageCnt=ImageCnt+1
	if filepathname="" Then
	file8=""
	else
	if PPD="1" then
		filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	else
		filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	if not filename="" then 
	value=UploadRequest.Item("arrowUp").Item("Value")
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
	file8=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	response.write "images/"& file8 &"<br>"
	end if
else
	if UploadRequest.Item("arrowUp").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'=======
'file 9
'=======
contentType=UploadRequest.Item("arrowDown").Item("ContentType")
filepathname=UploadRequest.Item("arrowDown").Item("FileName")
	
if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
	ImageCnt=ImageCnt+1
	if filepathname="" Then
	file9=""
	else
	if PPD="1" then
		filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	else
		filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	if not filename="" then 
	value=UploadRequest.Item("arrowDown").Item("Value")
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
	file9=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	response.write "images/"& file9 &"<br>"
	end if
else
	if UploadRequest.Item("arrowDown").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if


'response.end
If InValidImage>0 then

	'//DETAILED MSG
	if pcTmpErrSize = 1 then
		pcTmpErrSize = 0
		
		response.write "An Error occurred while attempting to upload your images: "&err.description&"<br><br>This is probably due to a default configuration option in IIS, which only allows the entity size in a POST request to be 200,000 bytes (~200KB).<br><br>You have to change this setting if you want to upload multiple images files or large image files through the ProductCart application (otherwise you can just FTP the file to the same location: e.g. the &quot;pc/catalog&quot; folder).<br><br>To change this setting:<br><br>- Open IIS Manager<br>- Navigate the tree to your application<br>- Double click the &quot;ASP&quot; icon in the main panel<br>- Expand the &quot;Limits&quot; category<br>- Modify the &quot;Maximum Requesting Entity Body Limit&quot; to a larger value<br><br><a href=""javascript:history.go(-1)""><font face="&Link&">Back</font></a></font></div>"
	else
		If Cint(InValidImage)=Cint(ImageCnt) then
			response.redirect "AdminIcons.asp?msg="&Server.URLEncode(InValidImage&" of your "&ImageCnt&" images were not a valid image format. Invalid image formats are not allowed to be uploaded to the server.")
		else
			response.redirect "dbicons.asp?file1="&file1&"&file2="&file2&"&file3="&file3&"&file4="&file4&"&file5="&file5&"&file6="&file6&"&file7="&file7&"&file8="&file8&"&file9="&file9&"&msg="&Server.URLEncode(InValidImage&" of your "&ImageCnt&" images were not a valid image format. Invalid image formats are not allowed to be uploaded to the server.")
		end if
	end if
	'//END DETAILED MSG


Else
	if ImageCnt>0 then
		response.redirect "dbicons.asp?file1="&file1&"&file2="&file2&"&file3="&file3&"&file4="&file4&"&file5="&file5&"&file6="&file6&"&file7="&file7&"&file8="&file8&"&file9="&file9&"&s=1&msg="&Server.URLEncode("Your images were successfully uploaded.")
	else
		response.redirect "AdminIcons.asp?msg="&Server.URLEncode("You need to supply at least one file to upload.")
	end if
end if

%>
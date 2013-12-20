<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
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

'=======
'file 1
'=======
file1=""
contentType=UploadRequest.Item("add2").Item("ContentType")
filepathname=UploadRequest.Item("add2").Item("FileName")

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
			value=UploadRequest.Item("add2").Item("Value")
			Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
			Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
			For i=1 to LenB(value)
				MyFile.Write chr(AscB(MidB(value,i,1)))
			Next
			MyFile.Close
			set myfile=nothing

			file1=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file1 &"<br>"
	end if
else
	if UploadRequest.Item("add2").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'=======
'file 2
'=======
file2=""
contentType=UploadRequest.Item("addtocart").Item("ContentType")
filepathname=UploadRequest.Item("addtocart").Item("FileName")

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
	value=UploadRequest.Item("addtocart").Item("Value")
	Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
	Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
	For i=1 to LenB(value)
	MyFile.Write chr(AscB(MidB(value,i,1)))
	Next
	MyFile.Close
	set myfile=nothing


	file2=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	response.write "images/"& file2 &"<br>"
	end if
else
	if UploadRequest.Item("addtocart").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'=======
'file 3
'=======
file3=""
contentType=UploadRequest.Item("addtowl").Item("ContentType")
filepathname=UploadRequest.Item("addtowl").Item("FileName")

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
	value=UploadRequest.Item("addtowl").Item("Value")
	Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
	Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
	For i=1 to LenB(value)
	MyFile.Write chr(AscB(MidB(value,i,1)))
	Next
	MyFile.Close
	set myfile=nothing


	file3=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	response.write "images/"& file3 &"<br>"
	end if
else
	if UploadRequest.Item("addtowl").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'=======
'file 4
'=======
file4=""
contentType=UploadRequest.Item("checkout").Item("ContentType")
filepathname=UploadRequest.Item("checkout").Item("FileName")

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
	value=UploadRequest.Item("checkout").Item("Value")
	Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
	Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
	For i=1 to LenB(value)
	MyFile.Write chr(AscB(MidB(value,i,1)))
	Next
	MyFile.Close
	set myfile=nothing

	file4=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	response.write "images/"& file4 &"<br>"
	end if
else
	if UploadRequest.Item("checkout").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'=======
'file 5
'=======
file5=""
contentType=UploadRequest.Item("cancel").Item("ContentType")
filepathname=UploadRequest.Item("cancel").Item("FileName")

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
	value=UploadRequest.Item("cancel").Item("Value")
	Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
	Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
	For i=1 to LenB(value)
	MyFile.Write chr(AscB(MidB(value,i,1)))
	Next
	MyFile.Close
	set myfile=nothing

	file5=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	response.write "images/"& file5 &"<br>"
	end if
else
	if UploadRequest.Item("cancel").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'=======
'file 6
'=======
file6=""
contentType=UploadRequest.Item("continueshop").Item("ContentType")
filepathname=UploadRequest.Item("continueshop").Item("FileName")

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
	value=UploadRequest.Item("continueshop").Item("Value")
	Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
	Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
	For i=1 to LenB(value)
	MyFile.Write chr(AscB(MidB(value,i,1)))
	Next
	MyFile.Close
	set myfile=nothing

	file6=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	response.write "images/"& file6 &"<br>"
	end if
else
	if UploadRequest.Item("continueshop").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'=======
'file 7
'=======
file7=""
contentType=UploadRequest.Item("morebtn").Item("ContentType")
filepathname=UploadRequest.Item("morebtn").Item("FileName")

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
	value=UploadRequest.Item("morebtn").Item("Value")
	Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
	Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
	For i=1 to LenB(value)
	MyFile.Write chr(AscB(MidB(value,i,1)))
	Next
	MyFile.Close
	set myfile=nothing

	file7=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	response.write "images/"& file7 &"<br>"
	end if
else
	if UploadRequest.Item("morebtn").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'=======
'file 8
'=======
file8=""
contentType=UploadRequest.Item("login").Item("ContentType")
filepathname=UploadRequest.Item("login").Item("FileName")

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
	value=UploadRequest.Item("login").Item("Value")
	Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
	Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
	For i=1 to LenB(value)
	MyFile.Write chr(AscB(MidB(value,i,1)))
	Next
	MyFile.Close
	set myfile=nothing

	file8=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	response.write "images/"& file8 &"<br>"
	end if
else
	if UploadRequest.Item("login").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'=======
'file 10
'=======
file10=""
contentType=UploadRequest.Item("recalculate").Item("ContentType")
filepathname=UploadRequest.Item("recalculate").Item("FileName")

if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
	ImageCnt=ImageCnt+1
	if filepathname="" Then
	file10=""
	else
	if PPD="1" then
		filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	else
		filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	if not filename="" then
	value=UploadRequest.Item("recalculate").Item("Value")
	Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
	Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
	For i=1 to LenB(value)
	MyFile.Write chr(AscB(MidB(value,i,1)))
	Next
	MyFile.Close
	set myfile=nothing

	file10=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	response.write "images/"& file10 &"<br>"
	end if
else
	if UploadRequest.Item("recalculate").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'=======
'file 11
'=======
file11=""
contentType=UploadRequest.Item("register").Item("ContentType")
filepathname=UploadRequest.Item("register").Item("FileName")

if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
	ImageCnt=ImageCnt+1
	if filepathname="" Then
	file11=""
	else
	if PPD="1" then
		filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	else
		filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	if not filename="" then
	value=UploadRequest.Item("register").Item("Value")
	Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
	Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
	For i=1 to LenB(value)
	MyFile.Write chr(AscB(MidB(value,i,1)))
	Next
	MyFile.Close
	set myfile=nothing

	file11=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	response.write "images/"& file11 &"<br>"
	end if
else
	if UploadRequest.Item("register").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'=======
'file 12
'=======
file12=""
contentType=UploadRequest.Item("remove").Item("ContentType")
filepathname=UploadRequest.Item("remove").Item("FileName")

if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
	ImageCnt=ImageCnt+1
	if filepathname="" Then
	file12=""
	else
	if PPD="1" then
		filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	else
		filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	if not filename="" then
	value=UploadRequest.Item("remove").Item("Value")
	Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
	Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
	For i=1 to LenB(value)
	MyFile.Write chr(AscB(MidB(value,i,1)))
	Next
	MyFile.Close
	set myfile=nothing

	file12=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	response.write "images/"& file12 &"<br>"
	end if
else
	if UploadRequest.Item("remove").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'=======
'file 14
'=======
file14=""
contentType=UploadRequest.Item("back").Item("ContentType")
filepathname=UploadRequest.Item("back").Item("FileName")

if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
	ImageCnt=ImageCnt+1
	if filepathname="" Then
	file14=""
	else
	if PPD="1" then
		filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	else
		filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	if not filename="" then
	value=UploadRequest.Item("back").Item("Value")
	Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
	Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
	For i=1 to LenB(value)
	MyFile.Write chr(AscB(MidB(value,i,1)))
	Next
	MyFile.Close
	set myfile=nothing

	file14=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	response.write "images/"& file14 &"<br>"
	end if
else
	if UploadRequest.Item("back").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'=======
'file 16
'=======
file16=""
contentType=UploadRequest.Item("viewcartbtn").Item("ContentType")
filepathname=UploadRequest.Item("viewcartbtn").Item("FileName")

if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
	ImageCnt=ImageCnt+1
	if filepathname="" Then
	file16=""
	else
	if PPD="1" then
		filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	else
		filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	if not filename="" then
	value=UploadRequest.Item("viewcartbtn").Item("Value")
	Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
	Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
	For i=1 to LenB(value)
	MyFile.Write chr(AscB(MidB(value,i,1)))
	Next
	MyFile.Close
	set myfile=nothing

	file16=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	response.write "images/"& file16 &"<br>"
	end if
else
	if UploadRequest.Item("viewcartbtn").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'=======
'file 17
'=======
file17=""
contentType=UploadRequest.Item("checkoutbtn").Item("ContentType")
filepathname=UploadRequest.Item("checkoutbtn").Item("FileName")

if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
	ImageCnt=ImageCnt+1
	if filepathname="" Then
	file17=""
	else
	if PPD="1" then
		filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	else
		filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	if not filename="" then
	value=UploadRequest.Item("checkoutbtn").Item("Value")
	Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
	Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
	For i=1 to LenB(value)
	MyFile.Write chr(AscB(MidB(value,i,1)))
	Next
	MyFile.Close
	set myfile=nothing

	file17=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
	end if
	response.write "images/"& file17 &"<br>"
	end if
else
	if UploadRequest.Item("checkoutbtn").Item("FileName")<>"" then
		ImageCnt=ImageCnt+1
		InValidImage=InValidImage+1
	end if
end if

'===========
'file 25
'=======
file25=""
contentType=UploadRequest.Item("pcv_placeOrder").Item("ContentType")
filepathname=UploadRequest.Item("pcv_placeOrder").Item("FileName")

if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file25=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
		value=UploadRequest.Item("pcv_placeOrder").Item("Value")
		Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
		Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
		For i=1 to LenB(value)
		MyFile.Write chr(AscB(MidB(value,i,1)))
		Next
		MyFile.Close
		set myfile=nothing

		file25=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file25 &"<br>"
		end if
	else
		if UploadRequest.Item("pcv_placeOrder").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if

	'=======
	'file 26
	'=======
	file26=""
	contentType=UploadRequest.Item("pcv_checkoutWR").Item("ContentType")
	filepathname=UploadRequest.Item("pcv_checkoutWR").Item("FileName")

	if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file26=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
		value=UploadRequest.Item("pcv_checkoutWR").Item("Value")
		Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
		Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
		For i=1 to LenB(value)
		MyFile.Write chr(AscB(MidB(value,i,1)))
		Next
		MyFile.Close
		set myfile=nothing

		file26=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file26 &"<br>"
		end if
	else
		if UploadRequest.Item("pcv_checkoutWR").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if

	'=======
	'file 27
	'=======
	file27=""
	contentType=UploadRequest.Item("pcv_processShip").Item("ContentType")
	filepathname=UploadRequest.Item("pcv_processShip").Item("FileName")

	if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file27=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
		value=UploadRequest.Item("pcv_processShip").Item("Value")
		Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
		Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
		For i=1 to LenB(value)
		MyFile.Write chr(AscB(MidB(value,i,1)))
		Next
		MyFile.Close
		set myfile=nothing

		file27=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file27 &"<br>"
		end if
	else
		if UploadRequest.Item("pcv_processShip").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if

	'=======
	'file 28
	'=======
	file28=""
	contentType=UploadRequest.Item("pcv_finalShip").Item("ContentType")
	filepathname=UploadRequest.Item("pcv_finalShip").Item("FileName")

	if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file28=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
		value=UploadRequest.Item("pcv_finalShip").Item("Value")
		Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
		Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
		For i=1 to LenB(value)
		MyFile.Write chr(AscB(MidB(value,i,1)))
		Next
		MyFile.Close
		set myfile=nothing

		file28=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file28 &"<br>"
		end if
	else
		if UploadRequest.Item("pcv_finalShip").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if

	'=======
	'file 29
	'=======
	file29=""
	contentType=UploadRequest.Item("pcv_backtoOrder").Item("ContentType")
	filepathname=UploadRequest.Item("pcv_backtoOrder").Item("FileName")

	if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file29=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
		value=UploadRequest.Item("pcv_backtoOrder").Item("Value")
		Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
		Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
		For i=1 to LenB(value)
		MyFile.Write chr(AscB(MidB(value,i,1)))
		Next
		MyFile.Close
		set myfile=nothing

		file29=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file29 &"<br>"
		end if
	else
		if UploadRequest.Item("pcv_backtoOrder").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if

	'=======
	'file 30
	'=======
	file30=""
	contentType=UploadRequest.Item("pcv_previous").Item("ContentType")
	filepathname=UploadRequest.Item("pcv_previous").Item("FileName")

	if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file30=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
		value=UploadRequest.Item("pcv_previous").Item("Value")
		Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
		Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
		For i=1 to LenB(value)
		MyFile.Write chr(AscB(MidB(value,i,1)))
		Next
		MyFile.Close
		set myfile=nothing

		file30=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file30 &"<br>"
		end if
	else
		if UploadRequest.Item("pcv_previous").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if


	'=======
	'file 31
	'=======
	file31=""
	contentType=UploadRequest.Item("pcv_next").Item("ContentType")
	filepathname=UploadRequest.Item("pcv_next").Item("FileName")

	if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file31=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
		value=UploadRequest.Item("pcv_next").Item("Value")
		Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
		Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
		For i=1 to LenB(value)
		MyFile.Write chr(AscB(MidB(value,i,1)))
		Next
		MyFile.Close
		set myfile=nothing


		file31=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file31 &"<br>"
		end if
	else
		if UploadRequest.Item("pcv_next").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if

'GGG Add-on start
	'=======
	'file 32
	'=======
	file32=""
	contentType=UploadRequest.Item("crereg").Item("ContentType")
	filepathname=UploadRequest.Item("crereg").Item("FileName")

	if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file32=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
		value=UploadRequest.Item("crereg").Item("Value")
		Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
		Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
		For i=1 to LenB(value)
		MyFile.Write chr(AscB(MidB(value,i,1)))
		Next
		MyFile.Close
		set myfile=nothing

		file32=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file32 &"<br>"
		end if
	else
		if UploadRequest.Item("crereg").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if

	'=======
	'file 33
	'=======
	file33=""
	contentType=UploadRequest.Item("delreg").Item("ContentType")
	filepathname=UploadRequest.Item("delreg").Item("FileName")

	if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file33=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
		value=UploadRequest.Item("delreg").Item("Value")
		Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
		Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
		For i=1 to LenB(value)
		MyFile.Write chr(AscB(MidB(value,i,1)))
		Next
		MyFile.Close
		set myfile=nothing

		file33=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file33 &"<br>"
		end if
	else
		if UploadRequest.Item("delreg").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if

	'=======
	'file 34
	'=======
	file34=""
	contentType=UploadRequest.Item("addreg").Item("ContentType")
	filepathname=UploadRequest.Item("addreg").Item("FileName")

	if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file34=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
		value=UploadRequest.Item("addreg").Item("Value")
		Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
		Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
		For i=1 to LenB(value)
		MyFile.Write chr(AscB(MidB(value,i,1)))
		Next
		MyFile.Close
		set myfile=nothing

		file34=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file34 &"<br>"
		end if
	else
		if UploadRequest.Item("addreg").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if

	'=======
	'file 35
	'=======
	file35=""
	contentType=UploadRequest.Item("updreg").Item("ContentType")
	filepathname=UploadRequest.Item("updreg").Item("FileName")

	if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file35=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
		value=UploadRequest.Item("updreg").Item("Value")
		Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
		Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
		For i=1 to LenB(value)
		MyFile.Write chr(AscB(MidB(value,i,1)))
		Next
		MyFile.Close
		set myfile=nothing

		file35=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file35 &"<br>"
		end if
	else
		if UploadRequest.Item("updreg").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if

	'=======
	'file 36
	'=======
	file36=""
	contentType=UploadRequest.Item("sendmsgs").Item("ContentType")
	filepathname=UploadRequest.Item("sendmsgs").Item("FileName")

	if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file36=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
		value=UploadRequest.Item("sendmsgs").Item("Value")
		Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
		Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
		For i=1 to LenB(value)
		MyFile.Write chr(AscB(MidB(value,i,1)))
		Next
		MyFile.Close
		set myfile=nothing

		file36=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file36 &"<br>"
		end if
	else
		if UploadRequest.Item("sendmsgs").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if

	'=======
	'file 37
	'=======
	file37=""
	contentType=UploadRequest.Item("retreg").Item("ContentType")
	filepathname=UploadRequest.Item("retreg").Item("FileName")

	if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file37=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
		value=UploadRequest.Item("retreg").Item("Value")
		Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
		Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
		For i=1 to LenB(value)
		MyFile.Write chr(AscB(MidB(value,i,1)))
		Next
		MyFile.Close
		set myfile=nothing

		file37=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file37 &"<br>"
		end if
	else
		if UploadRequest.Item("retreg").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if

'GGG Add-on end

	'=======
	'file 38
	'=======
	file38=""
	contentType=UploadRequest.Item("yellowupd").Item("ContentType")
	filepathname=UploadRequest.Item("yellowupd").Item("FileName")

	if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file38=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
			value=UploadRequest.Item("yellowupd").Item("Value")
			Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
			Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
			For i=1 to LenB(value)
			MyFile.Write chr(AscB(MidB(value,i,1)))
			Next
			MyFile.Close
			set myfile=nothing

			file38=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file38 &"<br>"
		end if
	else
		if UploadRequest.Item("yellowupd").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if

	'=======
	'file 39
	'=======
	file39=""
	contentType=UploadRequest.Item("savecart").Item("ContentType")
	filepathname=UploadRequest.Item("savecart").Item("FileName")

	if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file39=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
			value=UploadRequest.Item("savecart").Item("Value")
			Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
			Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
			For i=1 to LenB(value)
			MyFile.Write chr(AscB(MidB(value,i,1)))
			Next
			MyFile.Close
			set myfile=nothing

			file39=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file39 &"<br>"
		end if
	else
		if UploadRequest.Item("savecart").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if


If scBTO=1 then
	'BTO ADDON-S

	'=======
	'file 22
	'=======
	file22=""
	contentType=UploadRequest.Item("revorder").Item("ContentType")
filepathname=UploadRequest.Item("revorder").Item("FileName")

if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file22=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
		value=UploadRequest.Item("revorder").Item("Value")
		Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
		Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
		For i=1 to LenB(value)
		MyFile.Write chr(AscB(MidB(value,i,1)))
		Next
		MyFile.Close
		set myfile=nothing

		file22=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file22 &"<br>"
		end if
	else
		if UploadRequest.Item("revorder").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if

	'=======
	'file 23
	'=======
	file23=""
	contentType=UploadRequest.Item("submitquote").Item("ContentType")
	filepathname=UploadRequest.Item("submitquote").Item("FileName")

	if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file23=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
		value=UploadRequest.Item("submitquote").Item("Value")
		Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
		Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
		For i=1 to LenB(value)
		MyFile.Write chr(AscB(MidB(value,i,1)))
		Next
		MyFile.Close
		set myfile=nothing

		file23=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file23 &"<br>"
		end if
	else
		if UploadRequest.Item("submitquote").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if

	'=======
	'file 24
	'=======
	file24=""
	contentType=UploadRequest.Item("pcv_requestQuote").Item("ContentType")
	filepathname=UploadRequest.Item("pcv_requestQuote").Item("FileName")

	if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file24=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
		value=UploadRequest.Item("pcv_requestQuote").Item("Value")
		Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
		Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
		For i=1 to LenB(value)
		MyFile.Write chr(AscB(MidB(value,i,1)))
		Next
		MyFile.Close
		set myfile=nothing

		file24=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file24 &"<br>"
		end if
	else
		if UploadRequest.Item("pcv_requestQuote").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if

	'=======
	'file 18
	'=======
	file18=""
	contentType=UploadRequest.Item("customize").Item("ContentType")
	filepathname=UploadRequest.Item("customize").Item("FileName")

	if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file18=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
		value=UploadRequest.Item("customize").Item("Value")
		Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
		Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
		For i=1 to LenB(value)
		MyFile.Write chr(AscB(MidB(value,i,1)))
		Next
		MyFile.Close
		set myfile=nothing

		file18=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file18 &"<br>"
		end if
	else
		if UploadRequest.Item("customize").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if

	'=======
	'file 19
	'=======
	file19=""
	contentType=UploadRequest.Item("reconfigure").Item("ContentType")
	filepathname=UploadRequest.Item("reconfigure").Item("FileName")

	if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file19=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
		value=UploadRequest.Item("reconfigure").Item("Value")
		Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
		Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
		For i=1 to LenB(value)
		MyFile.Write chr(AscB(MidB(value,i,1)))
		Next
		MyFile.Close
		set myfile=nothing

		file19=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file19 &"<br>"
		end if
	else
		if UploadRequest.Item("reconfigure").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if

	'=======
	'file 20
	'=======
	file20=""
	contentType=UploadRequest.Item("resetdefault").Item("ContentType")
	filepathname=UploadRequest.Item("resetdefault").Item("FileName")

	if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file20=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
		value=UploadRequest.Item("resetdefault").Item("Value")
		Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
		Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
		For i=1 to LenB(value)
		MyFile.Write chr(AscB(MidB(value,i,1)))
		Next
		MyFile.Close
		set myfile=nothing

		file20=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file20 &"<br>"
		end if
	else
		if UploadRequest.Item("resetdefault").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if

	'=======
	'file 21
	'=======
	file21=""
	contentType=UploadRequest.Item("savequote").Item("ContentType")
	filepathname=UploadRequest.Item("savequote").Item("FileName")

	if instr(ucase(contentType),"IMAGE") AND IsUploadAllowed(filepathname) then
		ImageCnt=ImageCnt+1
		if filepathname="" Then
		file21=""
		else
		if PPD="1" then
			filename="/"&scPcFolder&"/pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		else
			filename="../pc/images/pc" & "/" & Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		if not filename="" then
		value=UploadRequest.Item("savequote").Item("Value")
		Set ScriptObject=Server.CreateObject("Scripting.FileSystemObject")
		Set MyFile=ScriptObject.CreateTextFile(Server.mappath(filename))
		For i=1 to LenB(value)
		MyFile.Write chr(AscB(MidB(value,i,1)))
		Next
		MyFile.Close
		set myfile=nothing

		file21=Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
		end if
		response.write "images/"& file21 &"<br>"
		end if
	else
		if UploadRequest.Item("savequote").Item("FileName")<>"" then
			ImageCnt=ImageCnt+1
			InValidImage=InValidImage+1
		end if
	end if
	'=======

	If InValidImage>0 then
		'//DETAILED MSG
		if pcTmpErrSize = 1 then
			pcTmpErrSize = 0

			response.write "An Error occurred while attempting to upload your images: "&err.description&"<br><br>This is probably due to a default configuration option in IIS, which only allows the entity size in a POST request to be 200,000 bytes (~200KB).<br><br>You have to change this setting if you want to upload multiple images files or large image files through the ProductCart application (otherwise you can just FTP the file to the same location: e.g. the &quot;pc/catalog&quot; folder).<br><br>To change this setting:<br><br>- Open IIS Manager<br>- Navigate the tree to your application<br>- Double click the &quot;ASP&quot; icon in the main panel<br>- Expand the &quot;Limits&quot; category<br>- Modify the &quot;Maximum Requesting Entity Body Limit&quot; to a larger value."
		else
			response.redirect "dbbuttons.asp?file1="&file1&"&file2="&file2&"&file3="&file3&"&file4="&file4&"&file5="&file5&"&file6="&file6&"&file7="&file7&"&file8="&file8&"&file9="&file9&"&file10="&file10&"&file11="&file11&"&file12="&file12&"&file13="&file13&"&file14="&file14&"&file15="&file15&"&file16="&file16&"&file17="&file17&"&file18="&file18&"&file19="&file19&"&file20="&file20&"&file21="&file21&"&file22="&file22&"&file23="&file23&"&file24="&file24&"&file25="&file25&"&file26="&file26&"&file27="&file27&"&file28="&file28&"&file29="&file29&"&file30="&file30&"&file31="&file31&"&file32="&file32&"&file33="&file33&"&file34="&file34&"&file35="&file35&"&file36="&file36&"&file37="&file37&"&file38="&file38&"&file39="&file39&"&msg="&Server.URLEncode(InValidImage&" of your "&ImageCnt&" images were not a valid image format. Invalid image formats are not allowed to be uploaded to the server.")
		end if
		'//END DETAILED MSG
	Else
		if ImageCnt>0 then
			response.redirect "dbbuttons.asp?file1="&file1&"&file2="&file2&"&file3="&file3&"&file4="&file4&"&file5="&file5&"&file6="&file6&"&file7="&file7&"&file8="&file8&"&file9="&file9&"&file10="&file10&"&file11="&file11&"&file12="&file12&"&file13="&file13&"&file14="&file14&"&file15="&file15&"&file16="&file16&"&file17="&file17&"&file18="&file18&"&file19="&file19&"&file20="&file20&"&file21="&file21&"&file22="&file22&"&file23="&file23&"&file24="&file24&"&file25="&file25&"&file26="&file26&"&file27="&file27&"&file28="&file28&"&file29="&file29&"&file30="&file30&"&file31="&file31&"&file32="&file32&"&file33="&file33&"&file34="&file34&"&file35="&file35&"&file36="&file36&"&file37="&file37&"&file38="&file38&"&file39="&file39&"&s=1&msg="&Server.URLEncode("Your images were successfully uploaded.")
		else
			response.redirect "AdminButtons.asp?msg="&Server.URLEncode("You need to supply at least one file to upload.")
		end if
	end if
	'BTO ADDON-E
Else
	If InValidImage>0 then
		If Cint(InValidImage)=Cint(ImageCnt) then
			response.redirect "AdminButtons.asp?msg="&Server.URLEncode(InValidImage&" of your "&ImageCnt&" images were not a valid image format. Invalid image formats are not allowed to be uploaded to the server.")
		else
			response.redirect "dbbuttons.asp?file1="&file1&"&file2="&file2&"&file3="&file3&"&file4="&file4&"&file5="&file5&"&file6="&file6&"&file7="&file7&"&file8="&file8&"&file9="&file9&"&file10="&file10&"&file11="&file11&"&file12="&file12&"&file13="&file13&"&file14="&file14&"&file15="&file15&"&file16="&file16&"&file17="&file17&"&file18="&file18&"&file19="&file19&"&file20="&file20&"&file21="&file21&"&file22="&file22&"&file23="&file23&"&file24="&file24&"&file25="&file25&"&file26="&file26&"&file27="&file27&"&file28="&file28&"&file29="&file29&"&file30="&file30&"&file31="&file31&"&file32="&file32&"&file33="&file33&"&file34="&file34&"&file35="&file35&"&file36="&file36&"&file37="&file37&"&file38="&file38&"&file39="&file39&"&msg="&Server.URLEncode(InValidImage&" of your "&ImageCnt&" images were not a valid image format. Invalid image formats are not allowed to be uploaded to the server.")
		end if
	Else
		if ImageCnt>0 then
			response.redirect "dbbuttons.asp?file1="&file1&"&file2="&file2&"&file3="&file3&"&file4="&file4&"&file5="&file5&"&file6="&file6&"&file7="&file7&"&file8="&file8&"&file9="&file9&"&file10="&file10&"&file11="&file11&"&file12="&file12&"&file13="&file13&"&file14="&file14&"&file15="&file15&"&file16="&file16&"&file17="&file17&"&file18="&file18&"&file19="&file19&"&file20="&file20&"&file21="&file21&"&file22="&file22&"&file23="&file23&"&file24="&file24&"&file25="&file25&"&file26="&file26&"&file27="&file27&"&file28="&file28&"&file29="&file29&"&file30="&file30&"&file31="&file31&"&file32="&file32&"&file33="&file33&"&file34="&file34&"&file35="&file35&"&file36="&file36&"&file37="&file37&"&file38="&file38&"&file39="&file39&"&s=1&msg="&Server.URLEncode("Your images were successfully uploaded.")
		else
			response.redirect "AdminButtons.asp?msg="&Server.URLEncode("You need to supply at least one file to upload.")
		end if
	end if
End If %>
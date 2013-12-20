<% pageTitle = "Import 'Order Shipped' Information - Import Results" %>
<% section = "orders" %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/sendmail.asp"-->

<%
on error resume next
Server.ScriptTimeout = 5400
dim CSVRecord(100),f, mySQL, conntemp, rstemp, rstemp1,TopRecord(100), IDcustom(2), Customcontent(2)
%>
<!--#include file="ship-common.asp"-->
<%
	Append=session("append")
	FileCSV = "../pc/catalog/" & session("importfile")
	if PPD="1" then
		FileCSV="/"&scPcFolder&"/pc/catalog/"&session("importfile")
	end if
	findit = Server.MapPath(FileCSV)
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	Err.number=0
	Set f = fso.OpenTextFile(findit, 1)
	
		if Err.number<>0 then
			session("importfilename")=""
			response.redirect "msgb.asp?message=" & server.URLEncode("Either the import file can no longer be located, or it was not properly formatted. Try uploading the file again.")
		end if

	call openDb()

	TotalCSVlines=0
	ImportedRecords=0
	fields=session("totalfields")
	TopLine=f.Readline
	if (customfieldsid(0)>-1) or (customfieldsid(1)>-1) or (customfieldsid(2)>-1) then
	B=split(TopLine,",")
	For i=lbound(B) to ubound(B)
		if left(B(i),1)=chr(34) then
    	B(i)=mid(B(i),2,len(B(i)))
    	end if
    	if right(B(i),1)=chr(34) then
    	B(i)=mid(B(i),1,len(B(i))-1)
    	end if
    Next

	end if
	
	if session("append")="1" then
	UpdateType="UPDATE"
	else
	UpdateType="IMPORT"
	end if
	PreRecords=""
	CATRecords=""
	
	SKUError=0
	
	Do While not f.AtEndofStream
	
	TempCSV=f.Readline
	TempCSV=replace(TempCSV,chr(34) & chr(34),"&quot;")
	TempCSV=replace(TempCSV,"'","''")
	A=split(TempCSV,",")
	RecordError=false
	TotalCSVlines=TotalCSVlines+1

		'Maybe Data Text Fields have commas or "End Text Line - VBCrLf" characters, fix these problems
	 	i=0
	 	j=0
	 	Do while j<fields
	 	if i>ubound(a) then
			ErrorsReport=ErrorsReport & "Record " & TotalCSVlines & ": Does not have enough data fields." & vbcrlf
			RecordError=true
			exit do
		end if
	 	if left(A(i),1)= chr(34) then
	 		if len(a(i))>1 then
	 		CSVRecord(j)=mid(A(i),2,len(A(i)))
	 		else
	 		CSVRecord(j)=""
	 		end if
	 		if (right(A(i),1)=chr(34)) and (len(a(i))>1) then
	 			CSVRecord(j) = mid(CSVRecord(j),1,len(CSVRecord(j))-1)
	 		else
	 			Do
	 				i=i+1
	 				if i<=ubound(a) then 
	 					CSVRecord(j)=CSVRecord(j) & "," & A(i)
	 				else
	 				Do
	 					Templine=f.ReadLine
	 					Templine=replace(Templine,chr(34) & chr(34),"&quot;")
	 					Templine=replace(Templine,"'","''")
	 					check=instr(Templine,chr(34) & ",")
	 					if check=0 then
	 						CSVRecord(j)=CSVRecord(j) & vbcrlf & Templine
	 					end if
	 				Loop Until (check>0)
	 				A=split(Templine,",")
	 				i=0
	 				CSVRecord(j)=CSVRecord(j) & vbcrlf & A(i)
	 				end if
	 			Loop Until right(A(i),1)=chr(34)
	 			CSVRecord(j)=mid(CSVRecord(j),1,len(CSVRecord(j))-1)
	 		end if
	 		i=i+1
	 	else
		 	CSVRecord(j)=A(i)	
		 	i=i+1
		end if
		j=j+1
		Loop 	
	 	
if RecordError=False then%>
	<!--#include file="ship-common1.asp"-->
<%end if%>
<%if RecordError=false then%>
	<!--#include file="ship-checkdata1.asp"-->
<%end if%>
<%if RecordError=false then

'***************
'Get product information before update
			mySQL="select * from Orders where IDOrder=" & ship_order & ";"
			set rstemp=conntemp.execute(mySQL)
			
			IF not rstemp.eof THEN

			PreRecord1=""
			PreRecord1=PreRecord1 & rstemp("idOrder") & "****"
			PreRecord1=PreRecord1 & "Ord" & "****"
			
			iCols = rstemp.Fields.Count
		    for dd=1 to iCols-1
		    FType="" & Rstemp.Fields.Item(dd).Type
		    if (Ftype="202") or (Ftype="203") or (Ftype="135") then
		    	PTemp=Rstemp.Fields.Item(dd).Value
		    	if PTemp<>"" then
		    		PTemp=replace(PTemp,"'","''")
		    		PTemp=replace(PTemp,vbcrlf,"DuLTVDu")
		    	end if
		    	if dd=1 then
		    		if (scDB="Access") and (Ftype="135") then
		    		PreRecord1=PreRecord1 & "#" & PTemp & "#"
		    		else
		    		PreRecord1=PreRecord1 & "'" & PTemp & "'"
		    		end if
		    	else
		    		if (scDB="Access") and (Ftype="135") then
		    		PreRecord1=PreRecord1 & "@@@@@#" & PTemp & "#"
		    		else
		    		PreRecord1=PreRecord1 & "@@@@@'" & PTemp & "'"
		    		end if
		    	end if
		    else
		    PTemp="" & Rstemp.Fields.Item(dd).Value
		    if PTemp<>"" then
		    else
		    PTemp="0"
		    end if
		    if dd=1 then
		    PreRecord1=PreRecord1 & PTemp
		    else
		    PreRecord1=PreRecord1 & "@@@@@" & PTemp
		    end if
		    end if
			next
			PreRecords=PreRecords & PreRecord1 & vbcrlf
			END IF

'***************

		if instr(ship_shipmethod,"'")>0 then
			SP_WPrice=replace(ship_shipmethod,"'","''")
		end if
		
		if instr(ship_tracking,"'")>0 then
			SP_WPrice=replace(ship_tracking,"'","''")
		end if
					
		query1=""
		
		if shipid<>-1 then
		if ship_ship="1" then
			query1=query1 & " orderstatus=4 "
		end if
		end if
		
		if ship_shipdate<>"" then
			if query1<>"" then
				query1=query1 & "," 
			end if
			if scDB="Access" then
		    	query1=query1 & " shipDate=#" & ship_shipdate & "# "
		    else
		    	query1=query1 & " shipDate='" & ship_shipdate & "' "
			end if
		end if
		
		if shipmethodid<>-1 then
			if query1<>"" then
				query1=query1 & "," 
			end if
	    	query1=query1 & " shipVia='" & ship_shipmethod & "' "
		end if
		
		if trackingid<>-1 then
			if query1<>"" then
				query1=query1 & "," 
			end if
	    	query1=query1 & " TrackingNum='" & ship_tracking & "' "
		end if
		
		if query1<>"" then
		query="UPDATE orders SET " & query1 &" WHERE idOrder="& ship_order
		set rs=connTemp.execute(query)
		query1=""
		end if
%>
	<!--#include file="ship-sendmail.asp"-->
<%		
	ImportedRecords=ImportedRecords+1
end if
	
	Loop

	f.Close
	Set f = nothing
	'Set f = fso.GetFile(findit)
	'f.Delete
	'Set fso = nothing
	'Set f = nothing
	call closeDB()
	
	if ImportedRecords>0 then
	
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	Set afi=fs.CreateTextFile(server.MapPath(".") & "\importlogs\ship-prologs.txt",True)
		
	afi.Writeline(UpdateType)
	afi.Writeline(PreRecords)
	afi.Close

	end if
	
	session("importfile")=""
	session("totalfields")=0
	
%>

<!--#include file="AdminHeader.asp"-->

<table class="pcCPcontent">
    <tr>  
        <td colspan="2"><h2>Steps:</h2></td>
    </tr>
    <tr>
        <td  width="5%" align="right"><img border="0" src="images/step1.gif"></td>
        <td width="95%"><font color="#A8A8A8">Upload data file</font></td>
    </tr>
    <tr>
        <td align="right"><img border="0" src="images/step2.gif"></td>
        <td><font color="#A8A8A8">Map fields</font></td>
    </tr>
    <tr>
        <td align="right"><img border="0" src="images/step3.gif"></td>
        <td><font color="#A8A8A8">Confirm mapping</font></td>
    </tr>
    <tr>
        <td align="right"><img border="0" src="images/step4a.gif"></td>
        <td><font color="#000000"><strong><%if session("append")="1" then%>Update<%else%>Import<%end if%> results</strong></font></td>
    </tr>
</table>

<div class="pcCPmessageSuccess">Total <b><%=ImportedRecords%></b> records were <%if session("append")="1" then%>updated<%else%>imported<%end if%> successfully!
<%if TotalCSVlines-ImportedRecords>0 then%><br>
Total <b><%=TotalCSVlines-ImportedRecords%></b> records could not be <%if session("append")="1" then%>updated<%else%>imported<%end if%> successfully!
<%end if%>
</div>

<% if ErrorsReport<>"" then%> 
	<div class="pcCPmessage">
    <strong>Error(s) Report</strong>
    <p><textarea rows="7" name="S1" cols="62" style="font-family: Arial; font-size: 10pt"><%=ErrorsReport%></textarea></p>
    </div>
<%end if%>

<%
session("append")=0
%>
<!--#include file="AdminFooter.asp"-->
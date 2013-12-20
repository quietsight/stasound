<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Export Affiliate Table" %>
<% section="" %>
<%PmAdmin=10%><!--#include file="adminv.asp"-->  
<!--#include file="../includes/utilities.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"--> 
<% 	
Response.Buffer=true
Response.Expires=0
	
'on error resume next 
dim query, conntemp, rstemp

call openDb()
' Choose the records to display
idaffiliate=request.form("idaffiliate")
affiliateName=request.form("affiliateName")
affiliateEmail=request.form("affiliateEmail")
commission=request.form("commission")
	
query="SELECT idaffiliate, affiliateName,affiliateEmail,commission FROM affiliates WHERE idaffiliate>1"
set rstemp=Server.CreateObject("ADODB.Recordset")     
rstemp.Open query, conntemp, adOpenForwardOnly, adLockReadOnly, adCmdText
%>

<html>
<head>
	<title>Affiliate Information</title>
</head>
<body>
<p align="center">
<%dim strReturnAs
strReturnAs=request.Form("ReturnAS")
select case strReturnAS
	case "CSV"
		CreateCSVFile()
	case "HTML"
		GenHTML()
	case "XLS"
		CreateXlsFile()
end select		
   
Set rstemp=Nothing
Set conntemp=Nothing		
%>

<%

Function GenFileName()
	dim fname
	fname="File"
	systime=now()
	fname= fname & cstr(year(systime)) & cstr(month(systime)) & cstr(day(systime))
	fname= fname  & cstr(hour(systime)) & cstr(minute(systime)) & cstr(second(systime))
	GenFileName=fname
End Function

Function GenHTML()%>
	<h2><font face="Verdana">Affiliates Report</font></h2>
	<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=2>
		<tr>
			<% If idaffiliate="1" then	%>
				<TD valign="middle" style="border-right: 1 solid #FFFFFF" bgcolor="#CCCCCC" height="32" align="center" nowrap><font face="Verdana" size="2"><strong>Affiliate ID</strong></font></TD>
			<% End If
			If affiliateName="1" then	%>
				<TD valign="middle" style="border-right: 1 solid #FFFFFF" bgcolor="#CCCCCC" height="32" align="center"><font face="Verdana" size="2"><strong>Name</strong></font></TD>
			<% End If
			If affiliateEmail="1" then	%>
				<TD valign="middle" style="border-right: 1 solid #FFFFFF" bgcolor="#CCCCCC" height="32" align="center"><font face="Verdana" size="2"><strong>Email</strong></font></TD>
			<% End If
			If commission="1" then	%>
				<TD valign="middle" style="border-right: 1 solid #FFFFFF" bgcolor="#CCCCCC" height="32" align="center"><font face="Verdana" size="2"><strong>Commission</strong></font></TD><%
			End If%>
		</TR>
		<%if(rstemp.BOF=True and rstemp.EOF=True) then%>
			<tr>
				<TD valign="top" nowrap><font face="Verdana" size="2" color="#FF0000"><br><b>No records found.</b><br></font></TD>
			</tr>
		<%	else
			rstemp.MoveFirst
			Do While Not rstemp.EOF
				pidaffiliate=rstemp("idaffiliate")
				paffiliateName=rstemp("affiliateName")
				paffiliateEmail=rstemp("affiliateEmail")
				pcommission=rstemp("commission")%>
				<TR>
					<%If idaffiliate="1" then	%>
						<TD nowrap height="26" style="border-bottom: 1 solid #000000" width="5%"><p align="right"><font face="Verdana" size="2"><%=pidaffiliate%>&nbsp;</font></TD>
					<%End If
					If affiliateName="1" then	%>
						<TD nowrap height="26" style="border-bottom: 1 solid #000000" width="40%"><font face="Verdana" size="2"><%=paffiliateName%>&nbsp;</font></TD>
					<%End If
					If affiliateEmail="1" then	%>
						<TD nowrap height="26" style="border-bottom: 1 solid #000000" width="40%"><font face="Verdana" size="2"><%=paffiliateEmail%>&nbsp;</font></TD>
					<%End If
					if commission="1" then	%>
						<TD nowrap height="26" style="border-bottom: 1 solid #000000" width="5%" nowrap><p align="right"><font face="Verdana" size="2"><%=scCurSign & money(pcommission)%>&nbsp;</font></TD>
					<%End If%>
				</TR>
				<%
				rstemp.movenext
			loop
		End if%>
	</TABLE>
	<br>
	<font face="Verdana" size="1"><b>Reported Date: <%=now()%></b></font>
	<br>	
<% End Function

Function CreateCSVFile()
	strFile=GenFileName()   
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	Set a=fs.CreateTextFile(server.MapPath(".") & "\" & strFile & ".csv",True)
	If Not rstemp.EOF Then
		strtext=""
		If idaffiliate="1" then
			strtext=strtext & chr(34) & "Affiliate ID" & chr(34) & ","
		End If
		If affiliateName="1" then
			strtext=strtext & chr(34) & "Name" & chr(34) & ","
		End If
		If affiliateEmail="1" then
			strtext=strtext & chr(34) & "Email" & chr(34) & ","
		End If
		If commission="1" then
			strtext=strtext & chr(34) & "Commission" & chr(34) & ","
		End If
		a.WriteLine(strtext)
		Do Until rstemp.EOF
			pidaffiliate=rstemp("idaffiliate")
			paffiliateName=rstemp("affiliateName")
			paffiliateEmail=rstemp("affiliateEmail")
			pcommission=rstemp("commission")
			set StringBuilderObj = new StringBuilder
			If idaffiliate="1" then 
				StringBuilderObj.append chr(34) & pidaffiliate & chr(34) & ","
			End If
			If affiliateName="1" then
				StringBuilderObj.append chr(34) & paffiliateName & chr(34) & ","
			End If
			If affiliateEmail="1" then
				StringBuilderObj.append chr(34) & paffiliateEmail & chr(34) & ","
			End If
			If commission="1" then
				StringBuilderObj.append chr(34) & pcommission & chr(34) & ","
			End If
			a.Writeline(StringBuilderObj.toString())
			set StringBuilderObj = nothing
			rstemp.MoveNext			
		Loop
	End If
	a.Close
	Set fs=Nothing
	response.redirect "getFile.asp?file="& strFile &"&Type=csv"	
End Function

Function CreateXlsFile()
	Dim xlWorkSheet
	Dim xlApplication 
				
	Set xlApplication=CreateObject("Excel.Application")
	xlApplication.Visible=False
	xlApplication.Workbooks.Add
	Set xlWorksheet=xlApplication.Worksheets(1)
	t=0
	If idaffiliate="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Affiliate ID"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If affiliateName="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Name"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If affiliateEmail="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Email"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If commission="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Commission"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	
	iRow=2
	If Not rstemp.EOF Then
		Do Until rstemp.EOF
			pidaffiliate=rstemp("idaffiliate")
			paffiliateName=rstemp("affiliateName")
			paffiliateEmail=rstemp("affiliateEmail")
			pcommission=rstemp("commission")
			t=0
			If idaffiliate="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=pidaffiliate
			End If
			If affiliateName="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=paffiliateName
			End If
			If affiliateEmail="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=paffiliateEmail
			End If
			If commission="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=pcommission
			End If
						
			iRow=iRow + 1
			rstemp.MoveNext
		Loop
	End If
	strFile=GenFileName()
	xlWorksheet.SaveAs Server.MapPath(".") & "\" & strFile & ".xls"
	xlApplication.Quit
	Set xlWorksheet=Nothing
	Set xlApplication=Nothing
	response.redirect "getFile.asp?file="& strFile &"&Type=xls"
End Function

set rstemp=nothing
call closeDb()
%>

</body>
</html>
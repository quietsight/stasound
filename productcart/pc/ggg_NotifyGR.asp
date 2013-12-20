<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"--> 
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="header.asp"-->
<% 
dim query, conntemp, pid

'Check to see if store has been turned off by admin, then display message
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
End If

'open database
call openDB()

pIdCustomer=session("idCustomer")

query="SELECT name,lastName,email FROM customers WHERE idCustomer=" &pIdCustomer
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
CustName=rs("name") & " " & rs("lastname")
CustEmail=rs("email")

set rs=nothing

gIDEvent=getUserInput(request("IDEvent"),0)

if gIDEvent<>"" then
	query="select pcEv_Name,pcEv_Date,pcEv_Code from pcEvents where pcEv_IDCustomer=" & pIDCustomer & " and pcEv_IDEvent=" & gIDEvent
	set rstemp=connTemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	if rstemp.eof then
		set rstemp=nothing
		call closedb()
		response.redirect "ggg_manageGRs.asp"
	else
		geName=rstemp("pcEv_Name")
		geDate=rstemp("pcEv_Date")
		if gedate<>"" then
			if scDateFrmt="DD/MM/YY" then
				gedate=(day(gedate)&"/"&month(gedate)&"/"&year(gedate))
			else
				gedate=(month(gedate)&"/"&day(gedate)&"/"&year(gedate))
			end if
		end if
		gCode=rstemp("pcEv_Code")
	
		SPath1=Request.ServerVariables("PATH_INFO")
		mycount1=0
		do while mycount1<2
			if mid(SPath1,len(SPath1),1)="/" then
				mycount1=mycount1+1
			end if
			if mycount1<2 then
				SPath1=mid(SPath1,1,len(SPath1)-1)
			end if
		loop
		SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & SPath1
	
		if Right(SPathInfo,1)="/" then
			gStr=SPathInfo & "pc/ggg_viewGR.asp?grcode=" & gCode
		else
			gStr=SPathInfo & "/pc/ggg_viewGR.asp?grcode=" & gCode
		end if
	
		gmsg=gename & vbcrlf & dictLanguage.Item(Session("language")&"_NotifyGR_10") & gedate & vbcrlf & dictLanguage.Item(Session("language")&"_NotifyGR_11") & gStr
	
	end if
	set rstemp=nothing
end if

%>
<script language="JavaScript">
<!--		
function Form1_Validator(theForm)
{
	if (theForm.yourname.value == "")
  	{
			alert("<%=dictLanguage.Item(Session("language")&"_alert_4")%>");
		    theForm.yourname.focus();
		    return (false);
	}

	if (theForm.youremail.value == "")
  	{
			alert("<%=dictLanguage.Item(Session("language")&"_alert_4")%>");
		    theForm.youremail.focus();
		    return (false);
	}
	
	if (theForm.friendsemail.value == "")
  	{
			alert("<%=dictLanguage.Item(Session("language")&"_alert_4")%>");
		    theForm.friendsemail.focus();
		    return (false);
	}
	
	if (theForm.title.value == "")
  	{
			alert("<%=dictLanguage.Item(Session("language")&"_alert_4")%>");
		    theForm.title.focus();
		    return (false);
	}
	
	if (theForm.message.value == "")
  	{
			alert("<%=dictLanguage.Item(Session("language")&"_alert_4")%>");
		    theForm.message.focus();
		    return (false);
	}
	
return (true);
}
//-->
</script>

<div id="pcMain">
<form name="Form1" action="ggg_NotifyGRb.asp?action=send" method="POST" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcMainTable">
<tr>
	<td colspan="2"> 
		<p><h1><%response.write dictLanguage.Item(Session("language")&"_NotifyGR_1")%>"<%=geName%>"<%response.write dictLanguage.Item(Session("language")&"_NotifyGR_1a")%></h1>
		<br>
		<br>
		<%response.write dictLanguage.Item(Session("language")&"_NotifyGR_2")%>
		</p>
	</td>
</tr>
<tr> 
	<td colspan="2">
		<input type="hidden" name="pid" value="<%=pid%>">
		<input type="hidden" name="pname" value="<%=productname%>">
	</td>
</tr>
	<tr> 
		<td width="21%"><%response.write dictLanguage.Item(Session("language")&"_NotifyGR_3")%></td>
		<td width="79%">
			<input type="text" size="40" name="yourname" value="<%=CustName%>">
		</td>
	</tr>
<tr> 
	<td width="21%"><%response.write dictLanguage.Item(Session("language")&"_NotifyGR_4")%></td>
	<td width="79%">
		<input type="text" size="40" name="youremail" value="<%=CustEmail%>">
	</td>
</tr>
<tr> 
	<td width="21%" valign="top" nowrap><%response.write dictLanguage.Item(Session("language")&"_NotifyGR_5")%></td>
	<td width="79%">
		<textarea name="friendsemail" rows="8" cols="40"></textarea>
		<br><i><%response.write dictLanguage.Item(Session("language")&"_NotifyGR_6")%></i>
		<br><br>
	</td>
</tr>
<tr> 
	<td width="21%"><%response.write dictLanguage.Item(Session("language")&"_NotifyGR_7")%></td>
	<td width="79%">
		<input type="text" size="40" name="title">
	</td>
</tr>
<tr> 
	<td width="21%"><%response.write dictLanguage.Item(Session("language")&"_NotifyGR_8")%></td>
	<td width="79%">&nbsp;</td>
</tr>
<tr> 
	<td width="21%">&nbsp;</td>
	<td width="79%">
		<textarea rows="15" cols="40" name="message"><%=gmsg%></textarea>
	</td>
</tr>
<tr>
	<td colspan="2">&nbsp;</td>
</tr>
<tr> 
	<td colspan="2"> 
		<p><br>
			<input type="image" id="submit" src="<%=rslayout("SendMsgs")%>" border="0" name="Submit" value="<%response.write dictLanguage.Item(Session("language")&"_tellafriend_9")%> ">
		</p>
	</td>
</tr>
</table>
</form>
</div>
<% call closeDb() %><!--#include file="footer.asp"-->
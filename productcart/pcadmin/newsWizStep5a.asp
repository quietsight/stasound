<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
Response.Buffer = False
Server.ScriptTimeout = 5400
%>
<%PmAdmin=7%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/sendmail.asp"--> 
<!--#include file="../includes/SQLFormat.txt"-->
<%
on error resume next

'Start SDBA
if request("pagetype")="1" then
	pcv_PageType="1"
	pcv_Title="Drop-Shippers"
else
	if request("pagetype")="0" then
		pcv_PageType="0"
		pcv_Title="Suppliers"
	else
		pcv_PageType=""
		pcv_Title="Customers"
	end if
end if
'End SDBA

dim rstemp, conntemp, mysql

call opendb()
Response.Buffer = False
%>
<% pageTitle="Newsletter Wizard: Message is Being Sent" %>
<% section="mngAcc" %>
<!--#include file="AdminHeader.asp"-->

<form name="d1">
<table class="pcCPcontent">
        <tr>
            <td colspan="2">
                <table width="100%">
                <tr>
                    <td width="5%" align="center"><img border="0" src="images/step1.gif"></td>
                    <td width="95%"><font color="#A8A8A8">Select <%=pcv_Title%></font></td>
                </tr>
                <tr>
                    <td align="center"><img border="0" src="images/step2.gif"></td>
                    <td><font color="#A8A8A8">Verify <%=pcv_Title%></font></td>
                </tr>
                <tr>
                    <td align="center"><img border="0" src="images/step3.gif"></td>
                    <td><font color="#A8A8A8">Enter message</font></td>
                </tr>
                <tr>
                    <td align="center"><img border="0" src="images/step4.gif"></td>
                    <td><font color="#A8A8A8">Test message</font></td>
                </tr>
                <tr>
                    <td align="center"><img border="0" src="images/step5a.gif"></td>
                    <td><b>Send message</b></td>
                </tr>
                </table>
            <p>&nbsp;</p>
            </td>
        </tr>
<tr>
	<td>Please wait while the Newsletter Wizard is sending your message to the <%=pcv_Title%> list that you have created....<br><br>
	<input type="text" size="5" name="s1" value="0" readonly style="border:medium none; font-size: 10pt; font-weight: bold; text-align:center"> e-mails of <%=session("AddrCount")%> have been sent.</td>
</tr>
<tr>
	<td align="center">&nbsp;</td>
</tr>
<tr>
	<td align="center">&nbsp;
</td>
</tr>
<tr>
	<td align="center">&nbsp;</td>
</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->
<%AddrList=session("AddrList")
MsgBody=replace(session("News_MsgBody"),"&quot;",chr(34))
Sentmails=0
Totalmails=session("AddrCount")
For dem=lbound(AddrList) to ubound(AddrList)
if trim(AddrList(dem))<>"" then
	Tn1=""
	For dd=1 to 10
	Randomize
	Tn1=Tn1 & Cstr(Fix(10*Rnd))
	Next
if session("News_MsgType")="1" then
MsgBody1="<html><body>" & MsgBody & "<br><br>" & Tn1 & "</body></html>"
else
MsgBody1=MsgBody & vbcrlf & vbcrlf & Tn1
end if
call sendMail(session("News_FromName"), session("News_FromEmail"), AddrList(dem), session("News_Title") & " - " & Tn1, MsgBody1)
Sentmails=Sentmails+1
Response.Write(vbCrLf & "<script>document.d1.s1.value = " & Sentmails & ";</script>")
end if
next

Tn1=""
For dd=1 to 16
Randomize
Tn1=Tn1 & Cstr(Fix(10*Rnd))
Next

CustFile=Cstr(day(now())) & Cstr(month(now())) & Cstr(year(now())) & Tn1 & ".txt"
MyList=""
For dem=lbound(AddrList) to ubound(AddrList)
if trim(AddrList(dem))<>"" then
myList=myList & AddrList(dem) & vbcrlf
end if
next

	findit = Server.MapPath("newslists/" & CustFile)
	Set fso = server.CreateObject("Scripting.FileSystemObject")
 	Set f = fso.CreateTextFile(findit,2)
	f.WriteLine(MyList)
	f.close
	Set fso = nothing
	Set f = nothing 

MsgFromName=replace(session("News_FromName"),"'","''")
MsgFromEmail=replace(session("News_FromEmail"),"'","''")	
MsgBody=replace(session("News_MsgBody"),"'","''")
MsgTitle=replace(session("News_Title"),"'","''")

MsgFromName=replace(MsgFromName,chr(34),"&quot;")
MsgFromEmail=replace(MsgFromEmail,chr(34),"&quot;")	
MsgBody=replace(MsgBody,chr(34),"&quot;")
MsgTitle=replace(MsgTitle,chr(34),"&quot;")

dim pTodaysDate
if SQL_Format="1" then
	pTodaysDate=(day(Date())&"/"&month(Date())&"/"&year(Date()))
else
	pTodaysDate=(month(Date())&"/"&day(Date())&"/"&year(Date()))
end if
if scDB="SQL" then
	mySQL="insert into News (FromDate,FromEmail,FromName,Title,MsgBody,MsgType,CustFile,CustTotal) values ('" & pTodaysDate & "','" & MsgFromEmail & "','" & MsgFromName & "','" & MsgTitle & "','" & MsgBody & "'," & session("News_MsgType") & ",'" & CustFile & "'," & session("AddrCount") & ")"
else
	mySQL="insert into News (FromDate,FromEmail,FromName,Title,MsgBody,MsgType,CustFile,CustTotal) values (#" & pTodaysDate & "#,'" & MsgFromEmail & "','" & MsgFromName & "','" & MsgTitle & "','" & MsgBody & "'," & session("News_MsgType") & ",'" & CustFile & "'," & session("AddrCount") & ")"
end if 

set rstemp=connTemp.execute(mySQL)

if err.number <> 0 then
  response.write Err.Description 
end If

session("News_FromEmail")=""
session("News_FromName")=""
session("News_Title")=""
session("News_MsgBody")=""
session("News_MsgType")=""
dim BList(1)
session("AddrList")=BList
session("AddrCount")=0
set rstemp=nothing
call closedb()
CustFile=""
%>

<script>
location="newsWizStep6.asp?pagetype=<%=pcv_PageType%>&sent=<%=Sentmails%>&total=<%=Totalmails%>";
</script>
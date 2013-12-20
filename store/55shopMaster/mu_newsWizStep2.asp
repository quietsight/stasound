<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=0%><!--#include file="adminv.asp"-->  
<!--#include file="../includes/utilities.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
dim rstemp, conntemp, query

Function GenFileName()
	dim fname
	fname="EmailList-"
	systime=now()
	fname= fname & cstr(year(systime)) & cstr(month(systime)) & cstr(day(systime))
	fname= fname  & cstr(hour(systime)) & cstr(minute(systime)) & cstr(second(systime))
	GenFileName=fname
End Function

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

AddrList=""

if (request("action")="update") and (request("submit1")<>"") then
	AddrList=request("AddrList")
else
	if Not IsNull(session("AddrList")) then
		AddrList=join(session("AddrList"),vbcrlf)
	end if
end if

if AddrList<>"" then

AList=split(AddrList,vbcrlf)
checkStrOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-@_."

For k=lbound(AList) to ubound(AList)
MyTest=0
AList(k)=trim(AList(k))
'have spaces in e-mail address
if instr(AList(k)," ")>0 then
	MyTest=1
end if
'do not have @ in e-mail address or do not have any character infront of @
if instr(AList(k),"@")<=1 then
	MyTest=1
end if
'have more than 1 @ in e-mail address
if instr(instr(AList(k),"@")+1,AList(k),"@")>0 then
	MyTest=1
end if		
'do not have domain behind @
if instr(instr(AList(k),"@")+1,AList(k),".")<=0 then
	MyTest=1
end if
'do not have domain behind @
if instr(instr(AList(k),"@")+1,AList(k),".")<=instr(AList(k),"@")+1 then
	MyTest=1
end if

TmpStr=AList(k)
For chay=1 to len(TmpStr)
	if instr(checkStrOK,mid(TmpStr,chay,1))<=0 then
	 MyTest=1
	end if
Next

if MyTest=1 then
AList(k)=""
end if
Next

For k=lbound(AList) to ubound(AList)
For k1=k+1 to ubound(AList)
if AList(k)>AList(k1) then
TempStr=AList(k1)
AList(k1)=AList(k)
AList(k)=TempStr
end if
next
next

mCount=0
For k=lbound(AList) to ubound(AList)
if trim(AList(k))<>"" then
mCount=mCount+1
end if
next

session("AddrList")=AList
session("AddrCount")=mCount

else

dim BList(1)
session("AddrList")=BList
session("AddrCount")=0

end if

if (request("action")="update") and (request("submit2")<>"") then
	AddrList=session("AddrList")
	set StringBuilderObj = new StringBuilder
	StringBuilderObj.append """First Name"",""Last Name"",""E-mail""" & vbcrlf
	call opendb()
	For k=lbound(AddrList) to ubound(AddrList)
		if trim(AddrList(k))<>"" then
			query="SELECT [name],lastName FROM customers WHERE email like '" & AddrList(k) & "';"
			set rs=connTemp.execute(query)
			FirstName=""
			LastName=""
			if not rs.eof then
				FirstName=rs("name")
				LastName=rs("lastName")
			end if
			set rs=nothing
			StringBuilderObj.append """" & FirstName & """,""" & LastName & """,""" & AddrList(k) & """" & vbcrlf
		end if
	Next
	strFile=GenFileName()   
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	Set a=fs.CreateTextFile(server.MapPath(".") & "\" & strFile & ".csv",True)
	a.Write(StringBuilderObj.toString)
	a.Close
	Set fs=Nothing
	call closeDb()
	set StringBuilderObj = nothing
	response.redirect "getFile.asp?file="& strFile &"&Type=csv&frompage=1"
end if

msg=""
if (request("action")="update") and (request("submit3")<>"") then
	AddrList=session("AddrList")
	tmpGrpName=request("tempGrpName")
	if tmpGrpName<>"" then
		tmpGrpName=replace(tmpGrpName,"'","''")
	else
		tmpGrpName="Temporary Group " & Date()
	end if
	MyList=""
	For k=lbound(AddrList) to ubound(AddrList)
		if trim(AddrList(k))<>"" then
			myList=myList & AddrList(k) & vbcrlf
		end if
	Next
	if MyList<>"" then
		call opendb()
		query="INSERT INTO pcMailUpSavedGroups (pcMailUpSavedGroups_Name,pcMailUpSavedGroups_Data) VALUES ('" & tmpGrpName & "','" & MyList & "');"
		set rs=connTemp.execute(query)
		set rs=nothing
		call closedb()
		msg="1"
	end if
end if

%>
<% pageTitle="Newsletter Wizard - STEP 2: Verify Group" %>
<% section="mngAcc" %>
<!--#include file="AdminHeader.asp"-->
<form name="Form1" method="post" action="mu_newsWizStep2.asp?action=update" class="pcForms">
<table class="pcCPcontent">
<%if msg<>"" then%>
<tr>
	<td colspan="2">
			<%if msg="1" then%>
				<div class="pcCPmessageSuccess">Email addresses list has been saved successfully.</div>
			<%end if%>
	</td>
</tr>
<%end if%>
<tr>
	<td colspan="2">
		<img src="images/pc2008_MailUp_Wizard.gif" alt="Newsletter Wizard - MailUp Integration" style="margin-bottom: 10px;" />
	</td>
</tr>
<tr>
	<td colspan="2">
		The <%=pcv_Title%> list that you have created contains <strong><%=session("AddrCount")%></strong> <%=pcv_Title%>. 
		<%if int(session("AddrCount"))>0 then%>
		The addresses are listed below:
		<%end if%>
	</td>
</tr>
<tr>
	<td colspan="2">
		<%AddrList=session("AddrList")
		set StringBuilderObj = new StringBuilder
		For k=lbound(AddrList) to ubound(AddrList)
			if trim(AddrList(k))<>"" then
				StringBuilderObj.append AddrList(k) & vbcrlf
			end if
		Next
		if session("AddrCount")="0" then
			'set StringBuilderObj = nothing
		end if
		%>
		<textarea name="AddrList" cols="60" rows="10"><%=StringBuilderObj.toString()%></textarea>
        <%
		set StringBuilderObj = nothing
		%>
	</td>
</tr>
<tr>
	<td colspan="2"><p>You can <strong>edit/add addresses</strong> using the field above. When adding addresses, please make sure to place one email address per line. When you are done, click on &quot;Update&quot;.</p></td>
</tr>
<tr>
	<td align="center" colspan="2">&nbsp;</td>
</tr>
<tr>
	<td align="center" colspan="2">
		<%'Start SDBA%>
		<input type=hidden name="pagetype" value="<%=pcv_PageType%>">
		<%'End SDBA%>
		<input type="submit" name="submit1" value="Update" class="submit2">
		&nbsp;
		<%if session("AddrCount")<>"0" then%>
		<input type="submit" name="submit2" value="Export list" class="submit2">
		&nbsp;
		<%end if%>
		<input type="button" value="Send to MailUp" onClick="location='mu_sendgroup.asp'" class="submit2">
		&nbsp;
				<input type="button" value="Continue in ProductCart" onClick="location='mu_newsWizStep3.asp?pagetype=<%=pcv_PageType%>'" class="submit2">
		&nbsp;
		<input type="button" name="back" value="Back" onClick="location='<%if pcv_pageType<>"" then%>mu_sds_newsWizStep1.asp?pagetype=<%=pcv_pageType%><%else%>mu_newsWizStep1.asp<%end if%>'">
	</td>
</tr>
<tr>
	<td colspan="2"><hr size="1" noshade="noshade" /></td>
</tr>
<tr>
	<td colspan="2">
		<div style="margin-bottom: 10px;"><strong>Save</strong> the customers you have filtered.</div>
		<div>Save as: 
		  <input type="text" name="tempGrpName" size="40" value="Filtered customers - <%=Date()%>">&nbsp;
	<input type="submit" name="submit3" value="Save" class="submit2" onclick="javascript: if (document.Form1.tempGrpName.value=='') {alert('Enter a nickname for this saved filter'); return(false);}">
	  </div>
	</td>
</tr>
</table>
</form>
<%call closedb()%><!--#include file="AdminFooter.asp"-->
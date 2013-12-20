<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->  
<!--#include file="../includes/openDb.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<%
'*****************************************************************************************************
'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
%><!--#include file="pcVerifySession.asp"--><%
pcs_VerifySession
'*****************************************************************************************************
'// END: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************

'Check to see if store has been turned off by admin, then display message
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
End If

pcCartArray=Session("pcCartSession")
ppcCartIndex=Session("pcCartIndex")

if countCartRows(pcCartArray, ppcCartIndex)=0 then
	response.redirect "msg.asp?message=9" 
end if

If session("customerType")=1 Then
	if calculateCartTotal(pcCartArray, ppcCartIndex)<scWholesaleMinPurchase then  
		response.redirect "msgb.asp?message="&Server.URLEncode(dictLanguage.Item(Session("language")&"_checkout_2")& scCurSign &money(scWholesaleMinPurchase)&" for wholesale purchases.<BR><BR><a href=""javascript:history.go(-1)"">Back</a>")
	end if
Else
	if calculateCartTotal(pcCartArray, ppcCartIndex)<scMinPurchase then  
		response.redirect "msgb.asp?message="&server.URLEncode(dictLanguage.Item(Session("language")&"_checkout_2")& scCurSign &money(scMinPurchase)&"<BR><BR><a href=""javascript:history.go(-1)"">Back</a>")
	end if
End If

dim conntemp

call openDb()

pIdCustomer=session("idCustomer")

if request("action")="add" then
	session("Cust_GcReName")=getUserInput(request("GcReName"),0)
	session("Cust_GcReEmail")=getUserInput(request("GcReEmail"),0)
	session("Cust_GcReMsg")=getUserInput(request("GcReMsg"),0)
	if request("pa")<>"" then
		'response.redirect "OrderVerify.asp?idDbSession=" & Server.URLEncode(getUserInput(request("idDbSession"),0)) & "&randomKey=" & Server.URLEncode(getUserInput(request("randomKey"),0))
	else
		response.redirect "checkout.asp"
	end if
else
	if session("Cust_GcReEmail")="" then
		if Session("Cust_IDEvent")<>"" then
			query="select Customers.name,Customers.lastname,Customers.email from pcEvents,Customers where pcEvents.pcEv_IDEvent=" & Session("Cust_IDEvent") & " and Customers.idcustomer=pcEvents.pcEv_IdCustomer"
			set rsGc=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rsGc=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			if not rsGc.eof then
				session("Cust_GcReName")=rsGc("name") & " " & rsGc("lastname")
				session("Cust_GcReEmail")=rsGc("email")
			end if
			set rsGc=nothing
		end if
	end if
end if

HaveGcsTest=0

For f=1 to pcCartIndex
	if pcCartArray(f,10)="0" then
		query="select pcprod_Gc from Products where idproduct=" & pcCartArray(f,0) & " AND pcprod_Gc=1"
		set rsGc=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsGc=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		if not rsGc.eof then
			HaveGcsTest=1
			set rsGc=nothing
			exit for
		end if
		set rsGc=nothing
	end if
next

if HaveGcsTest=0 then
	if request("pa")<>"" then
		'response.redirect "OrderVerify.asp?idDbSession=" & Server.URLEncode(getUserInput(request("idDbSession"),0)) & "&randomKey=" & Server.URLEncode(getUserInput(request("randomKey"),0))
	else
		response.redirect "checkout.asp"
	end if
end if

GcReName=session("Cust_GcReName")
if GcReName<>"" then
	GcReName=replace(GcReName,"''","'")
end if
GcReEmail=session("Cust_GcReEmail")
if GcReEmail<>"" then
	GcReEmail=replace(GcReEmail,"''","'")
end if
GcReMsg=session("Cust_GcReMsg")
if GcReMsg<>"" then
	GcReMsg=replace(GcReMsg,"''","'")
end if

%>
<!--#include file="header.asp"--> 
<script language="JavaScript">
<!--

function echeck(str) {

	var at="@"
	var dot="."
	var lat=str.indexOf(at)
	var lstr=str.length
	var ldot=str.indexOf(dot)
		if (str.indexOf(at)==-1){
		   return(false);
		}

		if (str.indexOf(at)==-1 || str.indexOf(at)==0 || str.indexOf(at)==lstr){
		   return(false);
		}

		if (str.indexOf(dot)==-1 || str.indexOf(dot)==0 || str.indexOf(dot)==lstr){
		    return(false);
		}

		 if (str.indexOf(at,(lat+1))!=-1){
		    return(false);
		 }

		 if (str.substring(lat-1,lat)==dot || str.substring(lat+1,lat+2)==dot){
		    return(false);
		 }

		 if (str.indexOf(dot,(lat+2))==-1){
		    return(false);
		 }
		
		 if (str.indexOf(" ")!=-1){
		    return(false);
		 }

	 return(true);
	}


function Form1_Validator(theForm)
{
	if (theForm.GcReEmail.value !="")
	{
		if (echeck(theForm.GcReEmail.value)==false)
		{
		alert("<%response.write dictLanguage.Item(Session("language")&"_NotifyRe_6")%>");
		theForm.GcReEmail.focus();
		return(false);
		}
	}
return (true);
}
//-->
</script>
<div id="pcMain">
<form method="post" name="Form1" action="ggg_NotifyRecipient.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcMainTable">
<tr> 
	<td colspan="2">
		<h1><%response.write dictLanguage.Item(Session("language")&"_NotifyRe_1")%></h1>
		<p><%response.write dictLanguage.Item(Session("language")&"_NotifyRe_2")%></p>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcSpacer"></td>
</tr>
<tr> 
	<td width="25%" nowrap> 
		<p><%response.write dictLanguage.Item(Session("language")&"_NotifyRe_3")%></p>
	</td>
	<td width="75%"><input type="text" size="40" name="GcReName" value="<%=GcReName%>"></td>
</tr>
<tr> 
	<td nowrap> 
		<p><%response.write dictLanguage.Item(Session("language")&"_NotifyRe_4")%></p>
	</td>
	<td><input type="text" size="40" name="GcReEmail" value="<%=GcReEmail%>"></td>
</tr>
<tr> 
	<td width="20%" valign="top"> 
		<p><%response.write dictLanguage.Item(Session("language")&"_NotifyRe_5")%></p>
	</td>
	<td width="80%"><textarea cols="35" rows=13 name="GcReMsg"><%=GcReMsg%></textarea></td>
</tr>
<tr> 
	<td colspan="2"> 
		<p>
			<br><br>
			<input type="image" id="submit" name="submit" value="<%response.write dictLanguage.Item(Session("language")&"_GiftWrap_7")%>" src="<%=RSlayout("submit")%>" border="0">
			<input type="hidden" name="pa" value="<%=request("pa")%>">
			<input type="hidden" name="idDbSession" value="<%=getUserInput(request("idDbSession"),0)%>">
			<input type="hidden" name="randomKey" value="<%=getUserInput(request("randomKey"),0)%>">
			<br>
		</p>
	</td>
</tr>
</table>
</form>
</div>
<%call closedb()%><!--#include file="footer.asp"-->
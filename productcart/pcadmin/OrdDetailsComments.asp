<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=9%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!-- #Include file="../pc/checkdate.asp" -->
<% Dim pageTitle, Section
pageTitle="Order Details - Administrator Comments"
Section="orders" %>
<!-- #Include File="Adminheader.asp" -->
<%

Dim rs, connTemp, query, intIdOrder

'// Get order ID
intIdOrder=getUserInput(request("IDOrder"),0)
if not validNum(intIdOrder) or intIdOrder="0" then
   response.redirect "msg.asp?message=45"
end if
'//

'Update or Retrieve Admin Comments
if (request("action")="add") then
	adminComments=replace(request("adminComments"),"'","''")
	call opendb()	
	query="UPDATE orders SET adminComments='"&adminComments&"' WHERE idOrder="& intIdOrder
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)	
	call closedb()
	response.redirect "Orddetails.asp?id="&intIdOrder&"&s=1&msg="&Server.URLEncode("Administrator comments updated successfully.")
else
	call opendb()	
	query="SELECT adminComments FROM orders WHERE idOrder="& intIdOrder
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)	
	padminComments=rs("adminComments")
	call closedb()	
end if
%>
<script language="JavaScript">
<!--
function Form1_Validator(theForm)
{
			if (theForm.adminComments.value == "")
 	{
		    alert("Please enter Administrator Comments for this order.");
		    theForm.Details.focus();
		    return (false);
	}
return (true);
}
function newWindow(file,window) {
		msgWindow=open(file,window,'resizable=no,width=400,height=500');
		if (msgWindow.opener == null) msgWindow.opener = self;
}
//-->
</script>

<form name="hForm" method="post" action="ordDetailsComments.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
<input type="hidden" value="<%=intIdOrder%>" name="IDOrder">
    <table class="pcCPcontent">
        <tr>
            <td colspan="2"><h2>You are editing Administrator Comments for <strong>Order #<%=clng(scpre)+intIdOrder%></strong></h2></td>
        </tr>
        <tr>
            <td align="right" valign="top">
            <input type="button" value="Use HTML Editor" onClick="newWindow('pop_HtmlEditor.asp?fi=adminComments','window2')">
            </td>
            <td><textarea name="adminComments" cols="70" rows="15" id="adminComments"><%=padminComments%></textarea></td>
        </tr>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
            <td colspan="2" align="center">
                <input type="submit" name="Submit" value="Add or Edit Comments" class="submit2">
                <input type="button" name="Back" value="Back" onClick="document.location.href='ordDetails.asp?id=<%=intIdOrder%>'">
            </td>
        </tr>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
    </table>
</form>
<!-- #Include File="Adminfooter.asp" -->
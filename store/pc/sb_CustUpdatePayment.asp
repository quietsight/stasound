<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/bto_language.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/ShipFromSettings.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="opc_contentType.asp" -->
<% dim conntemp,query,rs 

HaveSecurity=0
if session("idCustomer")=0 OR session("idCustomer")="" then
	HaveSecurity=1
end if

call openDb()

'Check to see if ARB has been turned off by admin, then display message
If scSBStatus="0" then
	response.redirect "msg.asp?message=212"
End If 

Call SetContentType()

IF HaveSecurity=0 THEN

	qry_ID = getUserInput(Request("ID"),0)
	if not validNum(qry_ID) then
	   qry_ID=0
	end if
				
	query="SELECT orders.idOrder, orders.orderDate, orders.total, orders.ord_OrderName, ProductsOrdered.idProductOrdered,ProductsOrdered.UnitPrice,ProductsOrdered.quantity, ProductsOrdered.pcSubscription_ID, ProductsOrdered.pcPO_SubAmount, ProductsOrdered.pcPO_SubActive, ProductsOrdered.pcPO_IsTrial, ProductsOrdered.pcPO_SubTrialAmount, ProductsOrdered.pcPO_SubStartDate, ProductsOrdered.pcPO_SubType FROM orders, productsordered WHERE orders.idCustomer=" & Session("idcustomer") &" AND orders.OrderStatus>1 AND orders.idOrder = ProductsOrdered.idOrder AND ProductsOrdered.pcSubscription_ID > 0 AND orders.idOrder = " & qry_ID
	'set rstemp=Server.CreateObject("ADODB.Recordset")
	'set rstemp=conntemp.execute(query)
	
	if request("action")="add" then
		
	
		UpdateSuccess="1"
	end if
	
END IF
%>
<html>
<head>
<TITLE><%=dictLanguage.Item(Session("language")&"_opc_gwa_title")%></TITLE>
<%
Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN  

Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")
%>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
<script language="JavaScript">
<!--
	
	function Form1_Validator(theForm)
	{
		return (true);
	}

//-->
</script>
</head>
<body style="margin: 0;">
<div id="pcMain">
<form method="post" name="Form1" action="sb_CustUpdatePayment.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
<input name="ID" type="hidden" value="<%=qry_ID%>">
<table class="pcMainTable">
	<%IF HaveSecurity=1 THEN%>
        <tr>
            <td>
                <div class="pcErrorMessage">You were logged out due to inactivity.</div>
            </td>
        </tr>
    <%ELSE%>
    
        <%IF UpdateSuccess="1" THEN%>
            <tr>
                <td>
                    <div class="pcSuccessMessage">Action Complete!</div>
                    <script>
                        setTimeout(function(){parent.closeManageSubscription()},1000);
                    </script>
                </td>
            </tr>        
		<%ELSE%>            

                <tr>
                	<td colspan="2"><h2>Update Payment Method</h2></td>
                </tr>
				<tr>
					<td valign="top">
					
                        <table class="pcShowContent">
                            <tr>
                                <td>
                                	put form here
                                </td>
                            </tr>
                        </table>
                    
					</td>
				</tr>
				<tr>
					<td colspan="2"><hr></td>
				</tr>

            <tr>
                <td colspan="2" class="pcSpacer"></td>
            </tr>	
            <tr>
                <td colspan="2"> 
                    <input type="image" id="DSubmit" name="DSubmit" value="DSubmit" src="images/sample/pc_button_update.gif" border="0" class="clearBorder">
                </td>
            </tr>
        <%END IF%>
    <%END IF%>
</table>
</form>
</div>
</body>
</html>
<% call closedb() %>
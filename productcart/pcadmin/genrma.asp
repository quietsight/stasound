<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Process a Return - RMA (Return Merchandise Authorization)" %>
<% section="mngRma"%>
<%PmAdmin=7%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/currencyformatinc.asp" -->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="AdminHeader.asp"-->

<%
dim query, connTemp, rs

Dim pIdOrder
pIdOrder=Request.QueryString("IdOrder")
if not validNum(pIdOrder) then response.Redirect("resultsAdvancedAll.asp?B1=View+All&dd=1")

call openDb()
query="SELECT ProductsOrdered.idProduct, ProductsOrdered.idOrder, products.description, products.sku, products.idProduct, orders.idOrder FROM ProductsOrdered, products, orders WHERE orders.idorder=ProductsOrdered.idOrder AND ProductsOrdered.idProduct=products.idProduct AND orders.idOrder=" &pIdOrder
set rs=Server.CreateObject("ADODB.Recordset")
set rs=conntemp.execute(query)
%>

<script language="JavaScript">
<!--
function Form1_Validator(theForm)
{
		// require that at least one checkbox be checked
		if (typeof theForm.idProduct.length != 'undefined') {
			var checkSelected = false;
			for (i = 0;  i < theForm.idProduct.length;  i++)
			{
			if (theForm.idProduct[i].checked)
			checkSelected = true;
			}
			if (!checkSelected)
			{
			alert("Please select at least one product.");
			return (false);
			}
		} else {
			if (!theForm.idProduct.checked)
			{
			alert("Please select at least one product");
			return (false);
			}
	}
	if (theForm.rmaReturnReason.value == "")
  	{
			 alert("Please enter a Return Reason.");
		    theForm.rmaReturnReason.focus();
		    return (false);
	}

return (true);
}
//-->
</script>
<%
if request.form("Submit")<>"" then
	rmaReturnReason=request.form("rmaReturnReason")
	Session("rmaReturnReason")=rmaReturnReason
end if
%>
<form method="POST" action="rmaAdminRequest.asp" name="orderform" onsubmit="return Form1_Validator(this)" class="pcForms">
<input name="rmaApproved" type="hidden" value="1">
<table class="pcCPcontent">
	<tr>
		<td colspan="2">Use this feature to generate an <strong>Return Merchandise Authorization</strong> number (RMA) for your customer for use when returning a product.
		</td>
	</tr>
	<tr>
		<td  colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2">Order ID: <%=pIdOrder%></td>
	</tr>
	<tr>
		<td colspan="2">The order contained the following <strong>products</strong>. Check the ones that will be returned:</td>
	</tr>
			<% 
			While Not rs.EOF
				pIdProduct=rs("idProduct") 
				pSku=rs("sku")
				pDescription=rs("description")
				%>
				<tr>
					<td colspan="2">
					<input name="idProduct" type="checkbox" id="idProduct" value="<% =pIdProduct %>" class="clearBorder">
					&nbsp;<%=psku%> - <% =pDescription %>
					</td>
				</tr>
				<%
				rs.MoveNext
			Wend
			set rs = nothing
			call closeDb()
			%>
	<tr>
		<td colspan="2"><hr></td>
	</tr>
	<%
	function makePassword(byVal maxLen)
			Dim strNewPass
			Dim whatsNext, upper, lower, intCounter
			Randomize
	
			For intCounter = 1 To maxLen
				whatsNext = Int((1 - 0 + 1) * Rnd + 0)
				If whatsNext = 0 Then
					'character
					upper = 90
					lower = 65
				Else
					upper = 57
					lower = 48
				End If
				strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
			Next
			makePassword = strNewPass
		end function

		pRmaNumber = makePassword(16)
	%>
	<tr>
		<td nowrap><strong>RMA Number</strong>:</td>
		<td width="100%">
			<input name="pRmaNumber" value="<%=pRmaNumber%>" type="text" size="20"> <em>A random RMA number was created for you.</em>
		</td>
	</tr>
	<tr>
		<td colspan="2">
			<input name="sendEmail" value="1" type="checkbox" class="clearBorder">
			&nbsp;Send RMA Number to customer via e-mail
		</td>
	</tr>
	<tr>
		<td colspan="2"><strong>Comments</strong> (e.g. Reasons why the products were returned): <div class="pcSmallText">These comments are included in the e-mail message sent to the customer.</div></td>
	</tr>
	<tr>	    
		<td colspan="2">
			<textarea rows="6" cols="60" name="rmaReturnReason" value="<%session("rmaReturnReason")%>"><%session("rmaReturnReason")%></textarea>
		</td>
	</tr>
	<tr>
		<td colspan="2"><hr></td>
	</tr>
	<tr>
		<td  colspan="2" align="center">
		<input type="hidden" name="idCustomer" value="<%session("idCustomer")%>">
		<input type="hidden" name="idOrder" value="<%=pIdOrder%>">
		<input type="submit" name="Submit" value="Submit" class="submit2">&nbsp;
		<input name="back" type="button" value="Back" onClick="JavaScript: history.back();"></a>
		</td>
	</tr>
</table>
</form>
<%
Session("rmaReturnReason")=""
%>
<!--#include file="AdminFooter.asp"-->
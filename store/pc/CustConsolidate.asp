<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%'Allow Guest Account
AllowGuestAccess=1
%>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/shipFromsettings.asp"--> 
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/bto_language.asp"-->
<!--#include file="../includes/rewards_language.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<%
err.number=0
dim query, conntemp, rs, rstemp, pIdOrder

call opendb()
%>
<!--#include file="header.asp"-->
<div id="pcMain">
	<div id="GlobalAjaxErrorDialog" title="Communication Error" style="display:none">
		<div class="pcErrorMessage">
			<%=dictLanguage.Item(Session("language")&"_ajax_globalerror")%>
		</div>
	</div>
	<table class="pcMainTable">   
	<tr>
		<td>
			<h1>
				<%=dictLanguage.Item(Session("language")&"_opc_cons_title")%>
			</h1>
		</td>
	</tr>
	<tr>
		<td>
        	<%
			query = "SELECT email FROM customers WHERE idCustomer = " & Session("idCustomer")
			set rs = Server.CreateObject("ADODB.Recordset")
			set rs = conntemp.execute(query)
			If Not rs.EOF Then
				pEmail = rs("email")
			End If 
			set rs = nothing
			%>
			<!--#include file="opc_inc_CustConsolidate.asp"-->
		</td>
	</tr>
    <tr>
        <td class="pcSpacer"></td>
    </tr>
</table>
</div>
<% call closedb() %>
<!--#include file="footer.asp"-->
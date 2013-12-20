<%@ language="vbscript" %>
<% 'option explicit %>
<% response.Buffer=true %>
<% Response.CacheControl = "no-cache" %>
<% Response.AddHeader "Pragma", "no-cache" %> 
<% Response.Expires = -1 %>

<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!-- #include file="Centinel_Config.asp"-->
<% dim conntemp, rs, query
'==========================================================================================
'= CardinalCommerce (http://www.cardinalcommerce.com)
'= Page used to create the inline frame window
'==========================================================================================
dim headerText, imageSRC

headerText = Messaging
%>
<%
	'=====================================================================================
	' Check the transaction Id value to verify that this transaction has not already
	' been processed. This attempts to block the user from using the back button.
	'=====================================================================================

	if Session("Centinel_TransactionId") = "" then
		If scSSL="" OR scSSL="0" Then
			tempCAURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
			tempCAURL=replace(tempCAURL,"https:/","https://")
			tempCAURL=replace(tempCAURL,"http:/","http://") 
		Else
			tempCAURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
			tempCAURL=replace(tempCAURL,"https:/","https://")
			tempCAURL=replace(tempCAURL,"http:/","http://")
		End If

			Session("Centinel_Message") = "Order Already Processed, User Hit the Back Button"
		redirectPage = tempCAURL&"?psslurl="&session("redirectPage")&"&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"&ordertotal="&session("x_amount")
	end if

%>
<!--#include file="header.asp"-->
<div id="pcMain">
<table class="pcMainTable">
	<tr>
		<td><img src="images/checkout_bar_step5.gif" alt=""></td>
	</tr>
	<tr>
		<td class="pcSpacer"></td>
	</tr>
	<tr valign="top"> 
		<td align="center">
		<%=headerText%>	
		</td>
	</tr>
	<tr>
		<td>
		<IFRAME SRC="Centinel_Launch.asp" WIDTH="500" HEIGHT="500" FRAMEBORDER="0">
					Frames are currently disabled or not supported by your browser.  Please click <A HREF="Centinel_Launch.asp">here</A> to continue processing your transaction.
		</IFRAME>
		</td>
	</tr>
</table>
</div>
<!--#include file="footer.asp"-->

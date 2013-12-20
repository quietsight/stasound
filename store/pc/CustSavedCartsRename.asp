<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/bto_language.asp"-->  
<!--#include file="../includes/rewards_language.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include FILE="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<% dim conntemp, query, rs, pcv_SavedCartName, tmpID

	tmpID=getUserInput(request("id"),10)
	if tmpID="" or IsNull(tmpID) then
		tmpID=0
	end if
	if not IsNumeric(tmpID) then
		tmpID=0
	end if
	if tmpID=0 then response.Redirect "CustSavedCarts.asp"
	
	if request("submit")<>"" then
		pcv_SavedCartName=getUserInput(request("SavedCartName"),250)
		pcv_SavedCartName=pcf_ReplaceQuotes(pcv_SavedCartName)
		call opendb()
		set rs=Server.CreateObject("ADODB.Recordset")
		query="UPDATE pcSavedCarts SET SavedCartName='" & pcv_SavedCartName & "' WHERE SavedCartID=" & tmpID & ";"
		set rs=conntemp.execute(query)
		set rs=nothing
		call closedb()
		response.Redirect("CustSavedCarts.asp")
		response.End()
	end if
	
	call opendb()
	query="SELECT SavedCartName FROM pcSavedCarts WHERE SavedCartID=" & tmpID & ";"
	set rs=connTemp.execute(query)
	If rs.eof then 
		response.Redirect "CustSavedCarts.asp"
	else
		pcv_SavedCartName=rs("SavedCartName")
	End if
	set rs=nothing
	call closeDb()
	
%>
<!--#include file="header.asp"-->
<div id="pcMain">
    <form action="CustSavedCartsRename.asp" method="post" class="pcForms">
    <input type="hidden" value="<%=tmpID%>" name="id">
	<table class="pcMainTable">
		<tr>
			<td>
				<h1>
					<%response.write dictLanguage.Item(Session("language")&"_CustPref_16")%>
				</h1>
			</td>
		</tr>
        <tr>
            <td class="pcSpacer"></td>
        </tr>
        <tr>
            <th><%response.write dictLanguage.Item(Session("language")&"_CustSavedCarts_8")%></th>
            <th></th>
        </tr>
        <tr>
            <td class="pcSpacer"></td>
        </tr>
        <tr>
            <td><input type="text" value="<%=pcv_SavedCartName%>" name="SavedCartName" size="50"></td>
        </tr>
        <tr>
            <td class="pcSpacer"><hr></td>
        </tr>
       	<tr>
        	<td>
            	<input type="submit" name="submit" class="submit2" value="<%response.write dictLanguage.Item(Session("language")&"_CustSavedCarts_9")%>">
                <input type="button" value="<%response.write dictLanguage.Item(Session("language")&"_msg_back")%>" onClick="document.location.href='CustSavedCarts.asp';">
            </td>
        </tr>
	</table>
    </form>
</div>
<!--#include file="footer.asp"-->
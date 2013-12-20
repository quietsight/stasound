<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Edit ZIP Code (Postal Code)" %>
<% section="shipOpt" %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/rc4.asp"-->
<%PmAdmin="1*4*"%><!--#include file="adminv.asp"--> 
<%
Dim rs, connTemp, query

zipcode=request("zipcode")
oldzipcode=request("oldzipcode")

if request("action")="update" then

	call openDB()
	query="update ZipCodeValidation set zipcode='" & zipcode & "' where zipcode='" & oldzipcode & "'"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	set rs=nothing
	call closeDb()
	response.redirect "DeliveryZipCodes_main.asp?s=1&msg=Zip Code updated successfully!"

end if

%>
	
<!--#include file="AdminHeader.asp"-->

<%
	call openDb()
	query="select * from ZipCodeValidation where zipcode='" & zipcode & "'"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	zipcode=rs("zipcode")
	set rs=nothing
	call closeDb()
%>
<form name="updateform" method="post" action="DeliveryZipCodes_edit.asp?action=update" class="pcForms">
    <input type="hidden" name="oldzipcode" value="<%=zipcode%>">

    <table class="pcCPcontent">
        <tr>
            <td colspan="2" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
        <tr> 
			<td width="20%" nowrap>Zip Code:</td>
            <td width="80%"><input type="text" name="zipcode" size="20" value="<%=zipcode%>"></td>
        </tr>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr> 
            <td></td>
            <td>
            	<input type="submit" name="submit" value="Update" class="submit2">
                &nbsp;<input type="button" name="back" value="Back" onClick="javascript:history.back()"> 
            </td>
        </tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->
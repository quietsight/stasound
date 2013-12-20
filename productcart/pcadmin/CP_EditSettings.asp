<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Canada Post Shipping Configuration - Edit Settings" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/CPconstants.asp"-->
<!--#include file="AdminHeader.asp"-->

<% 
if request.form("submit")<>"" then

    Session("ship_CP_Height")=request.form("CP_Height")
	if not validNum(Session("ship_CP_Height")) then Session("ship_CP_Height")=12
	
    Session("ship_CP_Width")=request.form("CP_Width")
	if not validNum(Session("ship_CP_Width")) then Session("ship_CP_Width")=12

    Session("ship_CP_Length")=request.form("CP_Length")
	if not validNum(Session("ship_CP_Length")) then Session("ship_CP_Length")=12
    
    response.redirect "../includes/PageCreateCPConstants.asp?refer=viewShippingOptions.asp#CP"
    response.end
else %>
    <form name="form1" method="post" action="CP_EditSettings.asp" class="pcForms">
        <table class="pcCPcontent">
            <tr>
                <td colspan="2" class="pcCPspacer">
                    <% ' START show message, if any %>
                        <!--#include file="pcv4_showMessage.asp"-->
                    <% 	' END show message %>
                </td>
            </tr>
            <tr> 
            <td colspan="2">
            <h2>Default Package Size</h2>
            Choose a default package size. This size will be the common size of the majority of packages that you ship.
            </td>
            </tr>
            <tr> 
                <td width="15%" align="right">Height:</td>
                <td width="85%"><input name="CP_Height" type="text" id="CP_Height" value="<%=CP_Height%>" size="4" maxlength="4"></td>
            </tr>
            <tr> 
                <td align="right">Width:</td>
                <td><input name="CP_Width" type="text" id="CP_Width" value="<%=CP_Width%>" size="4" maxlength="4"></td>
            </tr>
            <tr> 
                <td align="right">Length:</td>
                <td><input name="CP_Length" type="text" id="CP_Length" value="<%=CP_Length%>" size="4" maxlength="4">&nbsp;<span class="pcSmallText">This is the measurement of the longest side</span></td>
            </tr>
            <tr>
                <td colspan="2" class="pcCPspacer"></td>
            </tr>
            <tr> 
                <td></td>
                <td><input type="submit" name="Submit" value="Submit" class="submit2"></td>
            </tr>
        </table>
    </form>
<% end if %>
<!--#include file="AdminFooter.asp"-->
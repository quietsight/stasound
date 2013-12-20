<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Canada Post Shipping Configuration" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->

		<% if request.form("submit")<>"" then
			CP_Height=request.form("CP_Height")
			Session("ship_CP_Height")=CP_Height
			CP_Width=request.form("CP_Width")
			Session("ship_CP_Width")=CP_Width
			CP_Length=request.form("CP_Length")
			Session("ship_CP_Length")=CP_Length
		
			response.redirect "4_Step5.asp"
			response.end
		else %>
		<% ' START show message, if any %>
            <!--#include file="pcv4_showMessage.asp"-->
        <% 	' END show message %>
		<form name="form1" method="post" action="4_Step3.asp" class="pcForms">
        <table class="pcCPcontent">
          <tr> 
            <td colspan="2"><h2>Default Package Size</h2>
              If you are using your own packaging to ship products, please specify a default package size below. This size should be the most common size for the majority of packages that you ship.<br> </td>
          </tr>
          <tr> 
            <td align="right" width="20%">Height:</td>
            <td width="80%"><input name="CP_Height" type="text" id="CP_Height" value="10" size="4" maxlength="4"></td>
          </tr>
          <tr> 
            <td align="right">Width:</td>
            <td><input name="CP_Width" type="text" id="CP_Width" value="10" size="4" maxlength="4"></td>
          </tr>
          <tr> 
            <td align="right">Length:</td>
            <td><input name="CP_Length" type="text" id="CP_Length" value="10" size="4" maxlength="4"> 
              <span class="pcSmallText">This is the measurement of the longest side</span></td>
          </tr>
          <tr> 
            <td colspan="2">&nbsp;</td>
          <tr> 
            <td>&nbsp;</td>
            <td> <input type="submit" name="Submit" value="Submit"class="submit2"></td>
          </tr>
        </table>
		</form>
		<% end if %>
<!--#include file="AdminFooter.asp"-->
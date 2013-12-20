<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="FedEX Shipping Configuration" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
if request.form("submit")<>"" then
	FEDEX_SERVICE=request.form("FEDEX_SERVICE")
	Session("ship_FEDEX_SERVICE")=FEDEX_SERVICE
	if FEDEX_SERVICE="" then
		response.redirect "3_Step2.asp?msg="&Server.URLEncode("Select at least one service.")
		response.end
	end if
	freeshipStr=""
	handlingStr=""
	If request.form("free111")="YES" then
		freeamt=request.form("amt111")
		freeshipStr=freeshipStr&"111|"&replacecomma(freeamt)&","
	End if
	If request.form("handling111")<>"0" AND request.form("handling111")<>"" then
		If isNumeric(request.form("handling111"))=true then
			handlingStr=handlingStr&"111|"&replacecomma(request.form("handling111"))&"|"&request.form("shfee111")&","
		End If
	End if
	If request.form("free222")="YES" then
		freeamt=request.form("amt222")
		freeshipStr=freeshipStr&"222|"&replacecomma(freeamt)&","
	End if
	If request.form("handling222")<>"0" AND request.form("handling222")<>"" then
		If isNumeric(request.form("handling222"))=true then
			handlingStr=handlingStr&"222|"&replacecomma(request.form("handling222"))&"|"&request.form("shfee222")&","
		End If
	End if
	If request.form("free333")="YES" then
		freeamt=request.form("amt333")
		freeshipStr=freeshipStr&"333|"&replacecomma(freeamt)&","
	End if
	If request.form("handling333")<>"0" AND request.form("handling333")<>"" then
		If isNumeric(request.form("handling333"))=true then
			handlingStr=handlingStr&"333|"&replacecomma(request.form("handling333"))&"|"&request.form("shfee333")&","
		End If
	End if
	If request.form("free444")="YES" then
		freeamt=request.form("amt444")
		freeshipStr=freeshipStr&"444|"&replacecomma(freeamt)&","
	End if
	If request.form("handling444")<>"0" AND request.form("handling444")<>"" then
		If isNumeric(request.form("handling444"))=true then
			handlingStr=handlingStr&"444|"&replacecomma(request.form("handling444"))&"|"&request.form("shfee444")&","
		End If
	End if
	If request.form("free555")="YES" then
		freeamt=request.form("amt555")
		freeshipStr=freeshipStr&"555|"&replacecomma(freeamt)&","
	End if
	If request.form("handling555")<>"0" AND request.form("handling555")<>"" then
		If isNumeric(request.form("handling555"))=true then
			handlingStr=handlingStr&"555|"&replacecomma(request.form("handling555"))&"|"&request.form("shfee555")&","
		End If
	End if
	If request.form("free666")="YES" then
		freeamt=request.form("amt666")
		freeshipStr=freeshipStr&"666|"&replacecomma(freeamt)&","
	End if
	If request.form("handling666")<>"0" AND request.form("handling666")<>"" then
		If isNumeric(request.form("handling666"))=true then
			handlingStr=handlingStr&"666|"&replacecomma(request.form("handling666"))&"|"&request.form("shfee666")&","
		End If
	End if
	If request.form("free777")="YES" then
		freeamt=request.form("amt777")
		freeshipStr=freeshipStr&"777|"&replacecomma(freeamt)&","
	End if
	If request.form("handling777")<>"0" AND request.form("handling777")<>"" then
		If isNumeric(request.form("handling777"))=true then
			handlingStr=handlingStr&"777|"&replacecomma(request.form("handling777"))&"|"&request.form("shfee777")&","
		End If
	End if
	If request.form("free888")="YES" then
		freeamt=request.form("amt888")
		freeshipStr=freeshipStr&"888|"&replacecomma(freeamt)&","
	End if
	If request.form("handling888")<>"0" AND request.form("handling888")<>"" then
		If isNumeric(request.form("handling888"))=true then
			handlingStr=handlingStr&"888|"&replacecomma(request.form("handling888"))&"|"&request.form("shfee888")&","
		End If
	End if
	If request.form("free999")="YES" then
		freeamt=request.form("amt999")
		freeshipStr=freeshipStr&"999|"&replacecomma(freeamt)&","
	End if
	If request.form("handling999")<>"0" AND request.form("handling999")<>"" then
		If isNumeric(request.form("handling999"))=true then
			handlingStr=handlingStr&"999|"&replacecomma(request.form("handling999"))&"|"&request.form("shfee999")&","
		End If
	End if
	If request.form("freei111")="YES" then
		freeamt=request.form("amti111")
		freeshipStr=freeshipStr&"i111|"&replacecomma(freeamt)&","
	End if
	If request.form("handlingi111")<>"0" AND request.form("handlingi111")<>"" then
		If isNumeric(request.form("handlingi111"))=true then
			handlingStr=handlingStr&"i111|"&replacecomma(request.form("handlingi111"))&"|"&request.form("shfeei111")&","
		End If
	End if
	If request.form("freei222")="YES" then
		freeamt=request.form("amti222")
		freeshipStr=freeshipStr&"i222|"&replacecomma(freeamt)&","
	End if
	If request.form("handlingi222")<>"0" AND request.form("handlingi222")<>"" then
		If isNumeric(request.form("handlingi222"))=true then
			handlingStr=handlingStr&"i222|"&replacecomma(request.form("handlingi222"))&"|"&request.form("shfeei222")&","
		End If
	End if
	Session("ship_FedEX_freeshipStr")=freeshipStr
	Session("ship_FedEX_handlingStr")=handlingStr
	response.redirect "3_Step3.asp"
	response.end
else %>
	<form name="form1" method="post" action="3_Step2.asp">
		<table width="94%" border="0" cellpadding="4" cellspacing="0" align="center">
			<% if request.querystring("msg")<>"" then %>
				<tr class="normal"> 
				<td colspan="2"> 
				<table width="100%" border="0" cellspacing="0" cellpadding="4">
				<tr> 
				<td width="4%"><img src="images/pcadmin_note.gif" alt="Alert" width="20" height="20"></td>
				<td width="96%" class="message"><font color="#FF9900"><b><%=request.querystring("msg")%></b></font></td>
				</tr>
				</table></td>
				</tr>
			<% end if %>
			<tr class="normal"> 
			<td colspan="2" bgcolor="#e1e1e1"><b>Choose 
			one or more shipping services to offer to your 
			customers.</b></td>
			</tr>
				<tr class="normal"> 
				<td colspan="2"> 
				<table width="400" border="0" cellspacing="0" cellpadding="1">
				<tr> 
				<td width="57"><img src="images/fedex_express.gif" width="112" height="17"></td>
				<td width="339"><font color="#9933CC">&nbsp;</font></td>
				</tr>
				</table></td>
				</tr>
				<tr bgcolor="#DDEEFF" class="normal"> 
				<td width="5%"> 
				<input type="checkbox" name="FEDEX_SERVICE" value="111"></td>
				<td width="95%"><font color="#000000"><b>FedEx SameDay</b></font></td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td>  
				<input name="free111" type="checkbox" id="free111" value="YES">
				Offer free shipping for orders over <%=scCurSign%> 
				<input name="amt111" type="text" id="amt111" size="10" maxlength="10">
				</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td> 
				<hr align="left" width="325" size="1" noshade></td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td>Add Handling Fee <%=scCurSign%> 
				<input name="handling111" type="text" id="handling111" size="10" maxlength="10">
				</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td>  
				<input type="radio" name="shfee111" value="-1" checked>
				Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
				<input type="radio" name="shfee111" value="0">
				Integrate into shipping rate.</td>
				</tr>
				<tr bgcolor="#DDEEFF" class="normal"> 
				<td> 
				<input type="checkbox" name="FEDEX_SERVICE" value="222">
				</td>
				<td><font color="#000000"><b>FedEx First Overnight </b></font></td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td>  
				<input name="free222" type="checkbox" id="free222" value="YES">
				Offer free shipping for orders over <%=scCurSign%> 
				<input name="amt222" type="text" id="amt222" size="10" maxlength="10">
				</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td> 
				<hr align="left" width="325" size="1" noshade></td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				
				<td>&nbsp;</td>
				<td>Add Handling Fee <%=scCurSign%> 
				<input name="handling222" type="text" id="handling222" size="10" maxlength="10">
				</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td> 
				<input type="radio" name="shfee222" value="-1" checked>
				Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
				<input type="radio" name="shfee222" value="0">
				Integrate into shipping rate.</td>
				</tr>
				<tr bgcolor="#DDEEFF" class="normal"> 
				<td> 
				<input type="checkbox" name="FEDEX_SERVICE" value="333">
				</td>
				<td><font color="#000000"><b>FedEx Priority Overnight </b></font></td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td> 
				<input name="free333" type="checkbox" id="free333" value="YES">
				Offer free shipping for orders over <%=scCurSign%> 
				<input name="amt333" type="text" id="amt333" size="10" maxlength="10">
				</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td> 
				<hr align="left" width="325" size="1" noshade></td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td>Add Handling Fee <%=scCurSign%> 
				<input name="handling333" type="text" id="handling333" size="10" maxlength="10">
				</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td>  
				<input type="radio" name="shfee333" value="-1" checked>
				Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
				<input type="radio" name="shfee333" value="0">
				Integrate into shipping rate.</td>
				</tr>
				<tr bgcolor="#DDEEFF" class="normal"> 
				<td> 
				<input type="checkbox" name="FEDEX_SERVICE" value="444">
				</td>
				<td><font color="#000000"><b>FedEx Standard Overnight </b></font></td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td>  
				<input name="free444" type="checkbox" id="free444" value="YES">
				Offer free shipping for orders over <%=scCurSign%> 
				<input name="amt444" type="text" id="amt444" size="10" maxlength="10">
				</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td> 
				<hr align="left" width="325" size="1" noshade></td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td>Add Handling Fee <%=scCurSign%> 
				<input name="handling444" type="text" id="handling444" size="10" maxlength="10">
				</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td> 
				<input type="radio" name="shfee444" value="-1" checked>
				Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
				<input type="radio" name="shfee444" value="0">
				Integrate into shipping rate.</td>
				</tr>
				<tr bgcolor="#DDEEFF" class="normal"> 
				<td> 
				<input type="checkbox" name="FEDEX_SERVICE" value="555">
				</td>
				<td><font color="#000000"><b>FedEx 2Day</b></font></td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td> 
				<input name="free555" type="checkbox" id="free555" value="YES">
				Offer free shipping for orders over <%=scCurSign%> 
				<input name="amt555" type="text" id="amt555" size="10" maxlength="10">
				</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td> 
				<hr align="left" width="325" size="1" noshade></td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td>Add Handling Fee <%=scCurSign%> 
				<input name="handling555" type="text" id="handling555" size="10" maxlength="10">
				</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td>  
				<input type="radio" name="shfee555" value="-1" checked>
				Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
				<input type="radio" name="shfee555" value="0">
				Integrate into shipping rate.</td>
				</tr>
				<tr bgcolor="#DDEEFF" class="normal"> 
				<td> 
				<input type="checkbox" name="FEDEX_SERVICE" value="666">
				</td>
				<td><font color="#000000"><b>FedEx Express Saver </b></font></td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td>  
				<input name="free666" type="checkbox" id="free666" value="YES">
				Offer free shipping for orders over <%=scCurSign%> 
				<input name="amt666" type="text" id="amt666" size="10" maxlength="10">
				</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td> 
				<hr align="left" width="325" size="1" noshade></td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td>Add Handling Fee <%=scCurSign%> 
				<input name="handling666" type="text" id="handling666" size="10" maxlength="10">
				</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td> 
				<input type="radio" name="shfee666" value="-1" checked>
				Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
				<input type="radio" name="shfee666" value="0">
				Integrate into shipping rate.</td>
				</tr>
				<tr class="normal"> 
				<td colspan="2"> 
				<table width="400" border="0" cellspacing="0" cellpadding="1">
				<tr> 
				<td width="57"><img src="images/fedex_ground.gif" width="110" height="17"></td>
				<td width="339"><font color="#9933CC">&nbsp;</font></td>
				</tr>
				</table></td>
				</tr>
				<tr bgcolor="#DDEEFF" class="normal"> 
				<td> 
				<input type="checkbox" name="FEDEX_SERVICE" value="777">
				</td>
				<td><font color="#000000"><b>FedEx Ground</b></font></td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td> 
				<input name="free777" type="checkbox" id="free777" value="YES">
				Offer free shipping for orders over <%=scCurSign%> 
				<input name="amt777" type="text" id="amt777" size="10" maxlength="10">
				</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td> 
				<hr align="left" width="325" size="1" noshade></td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td>Add Handling Fee <%=scCurSign%> 
				<input name="handling777" type="text" id="handling777" size="10" maxlength="10">
				</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td>  
				<input type="radio" name="shfee777" value="-1" checked>
				Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
				<input type="radio" name="shfee777" value="0">
				Integrate into shipping rate.</td>
				</tr>
				<tr bgcolor="#DDEEFF" class="normal"> 
				<td> 
				<input type="checkbox" name="FEDEX_SERVICE" value="888">
				</td>
				<td><font color="#000000"><b>FedEx Home Delivery</b></font></td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td> 
				<input name="free888" type="checkbox" id="free888" value="YES">
				Offer free shipping for orders over <%=scCurSign%> 
				<input name="amt888" type="text" id="amt888" size="10" maxlength="10">
				</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td>Add Handling Fee <%=scCurSign%> 
				<input name="handling888" type="text" id="handling888" size="10" maxlength="10">
				</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				
				<td>&nbsp;</td>
				<td>  
				<input type="radio" name="shfee888" value="-1" checked>
				Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
				<input type="radio" name="shfee888" value="0">
				Integrate into shipping rate.</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td colspan="2" bgcolor="#FFFFFF"> 
				<table width="400" border="0" cellspacing="0" cellpadding="1">
				<tr> 
				<td width="57"><img src="images/fedex.gif" width="50" height="18"></td>
				<td width="339"><font color="#9933CC"><b><font color="#9966CC">International</font></b></font></td>
				</tr>
				</table></td>
				</tr>
				<tr bgcolor="#DDEEFF" class="normal"> 
				<td> 
				<input type="checkbox" name="FEDEX_SERVICE" value="i111">
				</td>
				<td><font color="#000000"><b>FedEx International Priority</b></font></td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td> 
				<input name="freei111" type="checkbox" id="freei111" value="YES">
				Offer free shipping for orders over <%=scCurSign%> 
				<input name="amti111" type="text" id="amti111" size="10" maxlength="10">
				</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td> 
				<hr align="left" width="325" size="1" noshade></td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td>Add Handling Fee <%=scCurSign%> 
				<input name="handlingi111" type="text" id="handlingi111" size="10" maxlength="10">
				</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td>  
				<input type="radio" name="shfeei111" value="-1" checked>
				Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
				<input type="radio" name="shfeei111" value="0">
				Integrate into shipping rate.</td>
				</tr>
				<tr bgcolor="#DDEEFF" class="normal"> 
				<td> 
				<input type="checkbox" name="FEDEX_SERVICE" value="i222">
				</td>
				<td><font color="#000000"><b>FedEx International Economy </b></font></td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td>  
				<input name="freei222" type="checkbox" id="freei222" value="YES">
				Offer free shipping for orders over <%=scCurSign%> 
				<input name="amti222" type="text" id="amti222" size="10" maxlength="10">
				</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td> 
				<hr align="left" width="325" size="1" noshade></td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td>Add Handling Fee <%=scCurSign%> 
				<input name="handlingi222" type="text" id="handlingi222" size="10" maxlength="10">
				</td>
				</tr>
				<tr bgcolor="#F1F1F1" class="normal"> 
				<td>&nbsp;</td>
				<td> 
				<input type="radio" name="shfeei222" value="-1" checked>
				Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
				<input type="radio" name="shfeei222" value="0">
				Integrate into shipping rate.</td>
			<tr class="normal"> 
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			</tr>
			<tr class="normal"> 
			<td colspan="2">
			<input type="submit" name="Submit" value="Submit" class="ibtnGrey"></td>
			</tr>
		</table>
	</form>
<% end if %>
<!--#include file="AdminFooter.asp"-->
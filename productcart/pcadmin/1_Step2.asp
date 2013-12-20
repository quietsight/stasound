<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="UPS OnLine&reg; Tools Shipping Configuration" %>
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
			UPS_Service=request.form("UPS_Service")
			Session("ship_UPS_ServiceStr")=UPS_Service
			if UPS_Service="" then
				response.redirect "1_Step2.asp?msg="&Server.URLEncode("Select at least one service.")
				response.end
			end if
			freeshipStr=""
			handlingStr=""
			If request.form("free01")="YES" then
				freeamt=request.form("amt01")
				freeshipStr=freeshipStr&"01|"&replacecomma(freeamt)&","
			End if
			If request.form("handling01")<>"0" AND request.form("handling01")<>"" then
				If isNumeric(request.form("handling01"))=true then
					handlingStr=handlingStr&"01|"&replacecomma(request.form("handling01"))&"|"&request.form("shfee01")&","
				End If
			End if
			If request.form("free02")="YES" then
				freeamt=request.form("amt02")
				freeshipStr=freeshipStr&"02|"&replacecomma(freeamt)&","
			End if
			If request.form("handling02")<>"0" AND request.form("handling02")<>"" then
				If isNumeric(request.form("handling02"))=true then
					handlingStr=handlingStr&"02|"&replacecomma(request.form("handling02"))&"|"&request.form("shfee02")&","
				End If
			End if
			If request.form("free03")="YES" then
				freeamt=request.form("amt03")
				freeshipStr=freeshipStr&"03|"&replacecomma(freeamt)&","
			End if
			If request.form("handling03")<>"0" AND request.form("handling03")<>"" then
				If isNumeric(request.form("handling03"))=true then
					handlingStr=handlingStr&"03|"&replacecomma(request.form("handling03"))&"|"&request.form("shfee03")&","
				End If
			End if
			If request.form("free07")="YES" then
				freeamt=request.form("amt07")
				freeshipStr=freeshipStr&"07|"&replacecomma(freeamt)&","
			End if
			If request.form("handling07")<>"0" AND request.form("handling07")<>"" then
				If isNumeric(request.form("handling07"))=true then
					handlingStr=handlingStr&"07|"&replacecomma(request.form("handling07"))&"|"&request.form("shfee07")&","
				End If
			End if
			If request.form("free08")="YES" then
				freeamt=request.form("amt08")
				freeshipStr=freeshipStr&"08|"&replacecomma(freeamt)&","
			End if
			If request.form("handling08")<>"0" AND request.form("handling08")<>"" then
				If isNumeric(request.form("handling08"))=true then
					handlingStr=handlingStr&"08|"&replacecomma(request.form("handling08"))&"|"&request.form("shfee08")&","
				End If
			End if
			If request.form("free11")="YES" then
				freeamt=request.form("amt11")
				freeshipStr=freeshipStr&"11|"&replacecomma(freeamt)&","
			End if
			If request.form("handling11")<>"0" AND request.form("handling11")<>"" then
				If isNumeric(request.form("handling11"))=true then
					handlingStr=handlingStr&"11|"&replacecomma(request.form("handling11"))&"|"&request.form("shfee11")&","
				End If
			End if
			If request.form("free12")="YES" then
				freeamt=request.form("amt12")
				freeshipStr=freeshipStr&"12|"&replacecomma(freeamt)&","
			End if
			If request.form("handling12")<>"0" AND request.form("handling12")<>"" then
				If isNumeric(request.form("handling12"))=true then
					handlingStr=handlingStr&"12|"&replacecomma(request.form("handling12"))&"|"&request.form("shfee12")&","
				End If
			End if
			If request.form("free13")="YES" then
				freeamt=request.form("amt13")
				freeshipStr=freeshipStr&"13|"&replacecomma(freeamt)&","
			End if
			If request.form("handling13")<>"0" AND request.form("handling13")<>"" then
				If isNumeric(request.form("handling13"))=true then
					handlingStr=handlingStr&"13|"&replacecomma(request.form("handling13"))&"|"&request.form("shfee13")&","
				End If
			End if
			If request.form("free14")="YES" then
				freeamt=request.form("amt14")
				freeshipStr=freeshipStr&"14|"&replacecomma(freeamt)&","
			End if
			If request.form("handling14")<>"0" AND request.form("handling14")<>"" then
				If isNumeric(request.form("handling14"))=true then
					handlingStr=handlingStr&"14|"&replacecomma(request.form("handling14"))&"|"&request.form("shfee14")&","
				End If
			End if
			If request.form("free54")="YES" then
				freeamt=request.form("amt54")
				freeshipStr=freeshipStr&"54|"&replacecomma(freeamt)&","
			End if
			If request.form("handling54")<>"0" AND request.form("handling54")<>"" then
				If isNumeric(request.form("handling54"))=true then
					handlingStr=handlingStr&"54|"&replacecomma(request.form("handling54"))&"|"&request.form("shfee54")&","
				End If
			End if
			If request.form("free59")="YES" then
				freeamt=request.form("amt59")
				freeshipStr=freeshipStr&"59|"&replacecomma(freeamt)&","
			End if
			If request.form("handling59")<>"0" AND request.form("handling59")<>"" then
				If isNumeric(request.form("handling59"))=true then
					handlingStr=handlingStr&"59|"&replacecomma(request.form("handling59"))&"|"&request.form("shfee59")&","
				End If
			End if
			If request.form("free65")="YES" then
				freeamt=request.form("amt65")
				freeshipStr=freeshipStr&"65|"&replacecomma(freeamt)&","
			End if
			If request.form("handling65")<>"0" AND request.form("handling65")<>"" then
				If isNumeric(request.form("handling65"))=true then
					handlingStr=handlingStr&"65|"&replacecomma(request.form("handling65"))&"|"&request.form("shfee65")&","
				End If
			End if
			Session("ship_UPS_freeshipStr")=freeshipStr
			Session("ship_UPS_handlingStr")=handlingStr
			response.redirect "1_Step3.asp"
			response.end
		else %>
<form name="form1" method="post" action="1_Step2.asp" class="pcForms">
<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>
<table class="pcCPcontent">
<tr> 
<td colspan="2">Choose one or more shipping services to offer to your customers.</td></tr>
<tr> 
<td width="5%" bgcolor="#DDEEFF">  
<input type="checkbox" name="UPS_Service" value="01"></td>
<td width="95%" bgcolor="#DDEEFF"><font color="#000000"> <b>UPS Next Day Air&reg;</b></font></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>  
<input name="free01" type="checkbox" id="free01" value="YES"> Offer free shipping for orders over <%=scCurSign%> 
<input name="amt01" type="text" id="amt01" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>
<hr align="left" width="325" size="1" noshade></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>Add Handling Fee <%=scCurSign%> 
<input name="handling01" type="text" id="handling01" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>  
<input type="radio" name="shfee01" value="-1" checked> Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
<input type="radio" name="shfee01" value="0"> Integrate into shipping rate.</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
<td bgcolor="#DDEEFF"> 
<input type="checkbox" name="UPS_Service" value="13"></td>
<td bgcolor="#DDEEFF"><font color="#000000"><b>UPS Next Day Air Saver&reg;</b></font></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td> 
<input name="free13" type="checkbox" id="free13" value="YES"> Offer free shipping for orders over <%=scCurSign%> 
<input name="amt13" type="text" id="amt13" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td><hr align="left" width="325" size="1" noshade></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>Add Handling Fee <%=scCurSign%> 
<input name="handling13" type="text" id="handling13" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>  
<input type="radio" name="shfee13" value="-1" checked> Display as a &quot;Shipping &amp; Handling&quot; charge.
<input type="radio" name="shfee13" value="0"> Integrate into shipping rate.</td>
</tr>

<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>                          
<td bgcolor="#DDEEFF"> 
<input type="checkbox" name="UPS_Service" value="14"></td>
<td bgcolor="#DDEEFF"><b>UPS Next Day Air&reg; Early A.M.&reg;</b></td>
</tr>
<tr> 
<td bgcolor="F1F1F1">&nbsp;</td>
<td bgcolor="F1F1F1">  
<input name="free14" type="checkbox" id="free14" value="YES"> Offer free shipping for orders over <%=scCurSign%> 
<input name="amt14" type="text" id="amt14" size="10" maxlength="10"></td>
</tr>
<tr> 
<td bgcolor="F1F1F1">&nbsp;</td>
<td bgcolor="F1F1F1">
<hr align="left" width="325" size="1" noshade></td>
</tr>
<tr> 
<td bgcolor="F1F1F1">&nbsp;</td>
<td bgcolor="F1F1F1">Add Handling Fee  <input name="handling14" type="text" id="handling14" size="10" maxlength="10"></td>
</tr>
<tr> 
<td bgcolor="F1F1F1">&nbsp;</td>
<td bgcolor="F1F1F1">  
<input type="radio" name="shfee14" value="-1" checked> Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
<input type="radio" name="shfee14" value="0"> Integrate into shipping rate.</td>
</tr>

<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>                          
<td bgcolor="#DDEEFF"> 
<input type="checkbox" name="UPS_Service" value="02"></td>
<td bgcolor="#DDEEFF"><b>UPS 2nd Day Air&reg;</b></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td> 
<input name="free02" type="checkbox" id="free02" value="YES"> Offer free shipping for orders over <%=scCurSign%> 
<input name="amt02" type="text" id="amt02" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>
<hr align="left" width="325" size="1" noshade></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>Add Handling Fee <%=scCurSign%> <input name="handling02" type="text" id="handling02" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>  
<input type="radio" name="shfee02" value="-1" checked> Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
<input type="radio" name="shfee02" value="0"> Integrate into shipping rate.</td>
</tr>

<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>                               
<td bgcolor="#DDEEFF"> 
<input type="checkbox" name="UPS_Service" value="59"></td>
<td bgcolor="#DDEEFF"><b>UPS 2nd Day Air A.M.&reg;</b></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>  
<input name="free59" type="checkbox" id="freE59" value="YES"> Offer free shipping for orders over <%=scCurSign%> 
<input name="amt59" type="text" id="amt59" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td><hr align="left" width="325" size="1" noshade></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>Add Handling Fee <%=scCurSign%> <input name="handling59" type="text" id="handling59" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>  
<input type="radio" name="shfee59" value="-1" checked> Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
<input type="radio" name="shfee59" value="0"> Integrate into shipping rate.</td>
</tr>

<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>													
<tr>                          
<td bgcolor="#DDEEFF"> 
<input type="checkbox" name="UPS_Service" value="12"></td>
<td bgcolor="#DDEEFF"><b>UPS 3 Day Select<sup>SM</sup></b></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>  
<input name="free12" type="checkbox" id="free12" value="YES"> Offer free shipping for orders over <%=scCurSign%> 
<input name="amt12" type="text" id="amt12" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td><hr align="left" width="325" size="1" noshade></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>Add Handling Fee <%=scCurSign%> <input name="handling12" type="text" id="handling12" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>  
<input type="radio" name="shfee12" value="-1" checked> Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
<input type="radio" name="shfee12" value="0"> Integrate into shipping rate.</td>
</tr>

<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
<td bgcolor="#DDEEFF"> 
<input type="checkbox" name="UPS_Service" value="65"></td>
<td bgcolor="#DDEEFF"><b>UPS Express Saver<sup>SM</sup></b></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>  
<input name="free65" type="checkbox" id="free65" value="YES"> Offer free shipping for orders over <%=scCurSign%> 
<input name="amt65" type="text" id="amt65" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td><hr align="left" width="325" size="1" noshade></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>Add Handling Fee <%=scCurSign%> <input name="handling65" type="text" id="handling65" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>  
<input type="radio" name="shfee65" value="-1" checked> Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
<input type="radio" name="shfee65" value="0"> Integrate into shipping rate.</td>
</tr>

<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
<td bgcolor="#DDEEFF"> 
<input type="checkbox" name="UPS_Service" value="11"></td>
<td bgcolor="#DDEEFF"><b>UPS Standard To Canada</b></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td> 
<input name="free11" type="checkbox" id="free11" value="YES"> Offer free shipping for orders over <%=scCurSign%> 
<input name="amt11" type="text" id="amt11" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td><hr align="left" width="325" size="1" noshade></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>Add Handling Fee <%=scCurSign%> <input name="handling11" type="text" id="handling11" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>  
<input type="radio" name="shfee11" value="-1" checked> Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
<input type="radio" name="shfee11" value="0"> Integrate into shipping rate.</td>
</tr>

<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
<td bgcolor="#DDEEFF"> 
<input type="checkbox" name="UPS_Service" value="03"></td>
<td bgcolor="#DDEEFF"><b>UPS Ground</b></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>  
<input name="free03" type="checkbox" id="free03" value="YES"> Offer free shipping for orders over <%=scCurSign%> 
<input name="amt03" type="text" id="amt03" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td><hr align="left" width="325" size="1" noshade></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>Add Handling Fee <%=scCurSign%> <input name="handling03" type="text" id="handling03" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>  
<input type="radio" name="shfee03" value="-1" checked> Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
<input type="radio" name="shfee03" value="0">Integrate into shipping rate.</td>
</tr>

<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>                          
<td bgcolor="#CCCC99"> 
<input type="checkbox" name="UPS_Service" value="07"></td>
<td bgcolor="#CCCC99"><b>UPS Worldwide Express<sup>SM</sup></b></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td> 
<input name="free07" type="checkbox" id="free07" value="YES"> Offer free shipping for orders over <%=scCurSign%> 
<input name="amt07" type="text" id="amt07" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td><hr align="left" width="325" size="1" noshade></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>Add Handling Fee <%=scCurSign%> <input name="handling07" type="text" id="handling07" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>  
<input type="radio" name="shfee07" value="-1" checked> Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
<input type="radio" name="shfee07" value="0"> Integrate into shipping rate.</td>
</tr>

<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
<td bgcolor="#CCCC99"> 
<input type="checkbox" name="UPS_Service" value="08"></td>
<td bgcolor="#CCCC99"><b>UPS Worldwide Expedited<sup>SM</sup></b></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>  
<input name="free08" type="checkbox" id="free08" value="YES"> Offer free shipping for orders over <%=scCurSign%> 
<input name="amt08" type="text" id="amt08" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td><hr align="left" width="325" size="1" noshade></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>Add Handling Fee <%=scCurSign%> <input name="handling08" type="text" id="handling08" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>  
<input type="radio" name="shfee08" value="-1" checked> Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
<input type="radio" name="shfee08" value="0"> Integrate into shipping rate.</td>
</tr>

<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>                         
<td bgcolor="#CCCC99"> 
<input type="checkbox" name="UPS_Service" value="54"></td>
<td bgcolor="#CCCC99"><b>UPS Worldwide Express Plus<sup>SM</sup></b></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td> 
<input name="free54" type="checkbox" id="free54" value="YES"> Offer free shipping for orders over <%=scCurSign%> 
<input name="amt54" type="text" id="amt54" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td><hr align="left" width="325" size="1" noshade></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>Add Handling Fee <%=scCurSign%> <input name="handling54" type="text" id="handling54" size="10" maxlength="10"></td>
</tr>
<tr bgcolor="#F1F1F1"> 
<td>&nbsp;</td>
<td>  
<input type="radio" name="shfee54" value="-1" checked> Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
<input type="radio" name="shfee54" value="0"> Integrate into shipping rate.</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"><hr></td>
</tr>
<tr> 
<td colspan="2">
<input type="submit" name="Submit" value="Submit" class="submit2"></td>
</tr>
<tr>
  <td colspan="2">&nbsp;</td>
</tr>
<tr>
  <td colspan="2"><div align="center">
    <table>
        <tr>
          <td width="58" valign="top" bgcolor="#FFFFFF"><div align="right"><img src="../UPSLicense/LOGO_S2.jpg" width="45" height="50" /></div></td>
          <td width="457" valign="top" bgcolor="#FFFFFF"><div align="center">UPS, THE UPS SHIELD TRADEMARK, THE UPS READY MARK, <br />THE UPS ONLINE TOOLS MARK AND THE COLOR BROWN ARE TRADEMARKS OF<br />UNITED PARCEL SERVICE OF AMERICA, INC. ALL RIGHTS RESERVED.</div></td>
        </tr>
      </table>
  </div></td>
</tr>
</table>
		</form>
		<% end if %>
<!--#include file="AdminFooter.asp"-->
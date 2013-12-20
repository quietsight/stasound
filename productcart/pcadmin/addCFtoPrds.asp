<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add a Custom Input Field to Selected Products" %>
<% nav=request("nav")
if nav="bto" then
	Section="services"
else
	Section="products"
end if %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
Dim query, conntemp, rs
If request("action")="new" then
	xfield=request.Form("xfield")
	if xfield="" then
		response.redirect "addCFtoPrds.asp?nav="&nav&"&x="&x&"&idproduct="&idproduct&"&message="&Server.UrlEncode("Name of your field can not be left blank.")        
	else
		if instr(xfield, ":") then
			response.redirect "addCFtoPrds.asp?nav="&nav&"&x="&x&"&idproduct="&idproduct&"&message="&Server.UrlEncode("Name of your field can not contain a colon (:).")        
		end if
		xfield=replace(xfield,"""","&quot;")
		xfield=replace(xfield,"'","''")
	end if
	fieldtype=request.Form("fieldtype")
	if fieldtype="1" then
		textareainput="0"
		rows="0"
		length=request.Form("length1")
		if length="" then
			length="15"
		else
			if Not isNumeric(length) then
				response.redirect "addCFtoPrds.asp?nav="&nav&"&x="&x&"&idproduct="&idproduct&"&message="&Server.UrlEncode("Length of your field must be a numeric value.")        
			end if
		end if
		maxchars=request.Form("maxchars")
		if maxchars="" then
			maxchars="150"
		else
			if Not isNumeric(maxchars) then
				response.redirect "addCFtoPrds.asp?nav="&nav&"&x="&x&"&idproduct="&idproduct&"&message="&Server.UrlEncode("The maximum number of characters for your field must be a numeric value.")        
			end if
		end if
	else
		textareainput="-1"
		length=request.Form("length2")
		if length="" then
			length="15"
		else
			if Not isNumeric(length) then
				response.redirect "addCFtoPrds.asp?nav="&nav&"&x="&x&"&idproduct="&idproduct&"&message="&Server.UrlEncode("Length of your field must be a numeric value.")        
			end if
		end if
		rows=request.Form("rows")
		if rows="" then
			rows="4"
		else
			if Not isNumeric(rows) then
				response.redirect "addCFtoPrds.asp?nav="&nav&"&x="&x&"&idproduct="&idproduct&"&message="&Server.UrlEncode("The number of rows for your text area must be a numeric value.")        
			end if
		end if
		maxchars=request.Form("maxchars2")
		if maxchars="" then
			maxchars="150"
		else
			if Not isNumeric(maxchars) then
				response.redirect "addCFtoPrds.asp?nav="&nav&"&x="&x&"&idproduct="&idproduct&"&message="&Server.UrlEncode("The maximum number of characters for your field must be a numeric value.")        
			end if
		end if
	end if
	xreq=request.Form("xreq")
	' randomNumber function, generates a number between 1 and limit
	function randomNumber(limit)
	 randomize
	 randomNumber=int(rnd*limit)+2
	end function

	randnum=randomNumber(9999999)
	call openDb()
	query="INSERT INTO xfields (xfield, textarea, widthoffield, maxlength, rowlength, randnum) VALUES ('"&xfield&"',"&textareainput&","&length&","&maxchars&","&rows&","&randnum&");"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=conntemp.execute(query)
	
	query="SELECT idxfield FROM xfields WHERE randnum="&randnum
	set rs=conntemp.execute(query)
	idxfield=rs("idxfield")
	set rs=nothing
	call closeDb()
	
	session("admin_idxfield")=idxfield
	session("admin_xreq")=xreq
	session("admin_customtype")=2
	session("admin_useExist")=0
	
	response.redirect "addCFtoPrds1.asp"
	response.end
Else
	if request("action")="exist" then
	idxfield=request("idx")
	xreq=request("xreq")
	
	if xreq="" then
	xreq=0
	end if
	
	session("admin_idxfield")=idxfield
	session("admin_xreq")=xreq
	session("admin_customtype")=2
	session("admin_useExist")=1
	response.redirect "addCFtoPrds1.asp"
	response.end
	end if
End if

' START show message, if any 
%>
<!--#include file="pcv4_showMessage.asp"-->
<%
' END show message 

call opendb()
query="SELECT idxfield,xfield FROM xfields ORDER BY xfield asc;"
set rs=Server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query)

if not rs.eof then
	pcArray=rs.getRows()
	intCount=ubound(pcArray,2)
%>
	<form name="form1" method="post" action="addCFtoPrds.asp?action=exist" class="pcForms">
	<input name="nav" type="hidden" value="<%=nav%>">
		<table class="pcCPcontent">
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
				<th colspan="2">Using an existing custom input field:</th>
			</tr>
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td width="25%" align="right">Text to display:</td>
				<td width="75%">
					<select name="idx">
						<%for i=0 to intCount%>
						<option value="<%=pcArray(0,i)%>"><%=pcArray(1,i)%></option>
						<%next%>
					</select>
				</td>
			</tr>
			<tr>
				<td align="right">Required:</td>
				<td> 
				<input type="radio" name="xreq" value="-1" class="clearBorder">Yes 
				<input type="radio" name="xreq" value="0" checked class="clearBorder">No
				</td>
			</tr>
			<tr>
				<td width="25%">&nbsp;</td>
				<td> 
				<input type="submit" name="submit" value="Continue" class="submit2">
				&nbsp;<input type="button" name="back" value="Back" onClick="javascript:history.back()">
				</td>
			</tr>
		</table>
	</form>          
<%
end if

set rs = nothing
call closedb()
%>

<form name="form2" method="post" action="addCFtoPrds.asp?action=new" class="pcForms">
<input name="nav" type="hidden" value="<%=nav%>">          
	<table class="pcCPcontent">
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<th colspan="2">Adding a new custom input field:</th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td width="25%" align="right">Text to display:</td>
			<td width="75%" valign="top"> 
			<input name="xfield" type="text" id="xfield" size="30" maxlength="150"> <span class="pcCPnotes">colons (:) are not permitted in field names.</span>
			</td>
		</tr>
		<tr> 
			<td colspan="2" align="right">&quot;Text to display&quot; is the name of the input field that customers will see on the product details page (e.g. &quot;Embroidery - Front Pocket&quot;)</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td align="right">Field Type:</td>
			<td><input name="fieldtype" type="radio" value="1" checked class="clearBorder">Text Field</td>
		</tr>
		<tr> 
			<td></td>
			<td>&nbsp;Length of field:
			<input name="length1" type="text" id="length1" size="5" maxlength="10">
			</td>
		</tr>
		<tr> 
			<td height="10"></td>
			<td>&nbsp;Maximum chars: <input name="maxchars" type="text" id="maxchars" size="5" maxlength="10">
			</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td height="10">&nbsp;</td>
			<td><input type="radio" name="fieldtype" value="2" class="clearBorder">Text Area</td>
		</tr>
		<tr> 
			<td></td>
			<td>&nbsp;Length of field: <input name="length2" type="text" id="length2" size="5" maxlength="10"></td>
		</tr>
		<tr> 
			<td></td>
			<td>&nbsp;Height (number of rows): <input name="rows" type="text" id="rows" size="5" maxlength="10"></td>
		</tr>
		<tr> 
			<td height="10"></td>
			<td>&nbsp;Maximum chars: <input name="maxchars2" type="text" id="maxchars" size="5" maxlength="10">
			</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td align="right">Required:</td>
			<td> 
			<input type="radio" name="xreq" value="-1" class="clearBorder">Yes 
			<input type="radio" name="xreq" value="0" checked class="clearBorder">No
			</td>
		</tr>
		<tr> 
		<td>&nbsp;</td>
		<td>If this option is set to 'Yes', customers will be required to enter information in the input field before proceeding.</td>
		</tr>
		<tr> 
			<td>&nbsp;</td>
			<td>
			<input type="submit" name="submit1" value="Continue" class="submit2">&nbsp;
			<input type="button" name="back" value="Back" onClick="javascript:history.back()">
			</td>
		</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Edit Custom Input Field" %>
<% Section="layout" %>
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
If request.form("submit")<>"" then
	x=request.Form("x")
	idproduct=request.Form("idproduct")
	idxfield=request.Form("idxfield")
	xfield=request.Form("xfield")
	if xfield="" then
		response.redirect "modCustomFields.asp?x="&x&"&idproduct="&idproduct&"&message="&Server.UrlEncode("Name of your field can not be left blank.")        
	else
		if instr(xfield, ":") then
			response.redirect "modCustomFields.asp?x="&x&"&idproduct="&idproduct&"&message="&Server.UrlEncode("Name of your field can not contain a colon (:).")        
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
				response.redirect "modCustomFields.asp?x="&x&"&idproduct="&idproduct&"&message="&Server.UrlEncode("Length of your field must be a numeric value.")        
			end if
		end if
		maxchars=request.Form("maxchars")
		if maxchars="" then
			maxchars="150"
		else
			if Not isNumeric(maxchars) then
				response.redirect "modCustomFields.asp?x="&x&"&idproduct="&idproduct&"&message="&Server.UrlEncode("The maximum number of characters for your field must be a numeric value.")        
			end if
		end if
	else
		textareainput="-1"
		length=request.Form("length2")
		if length="" then
			length="15"
		else
			if Not isNumeric(length) then
				response.redirect "modCustomFields.asp?x="&x&"&idproduct="&idproduct&"&message="&Server.UrlEncode("Length of your field must be a numeric value.")        
			end if
		end if
		rows=request.Form("rows")
		if rows="" then
			rows="4"
		else
			if Not isNumeric(rows) then
				response.redirect "modCustomFields.asp?x="&x&"&idproduct="&idproduct&"&message="&Server.UrlEncode("The number of rows for your text area must be a numeric value.")        
			end if
		end if
		maxchars=request.Form("maxchars2")
		if maxchars="" then
			maxchars="150"
		else
			if Not isNumeric(maxchars) then
				response.redirect "modCustomFields.asp?x="&x&"&idproduct="&idproduct&"&message="&Server.UrlEncode("The maximum number of characters for your field must be a numeric value.")        
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
	query="UPDATE xfields SET xfield='"&xfield&"', textarea="&textareainput&", widthoffield="&length&", maxlength="&maxchars&", rowlength="&rows&" WHERE idxfield="&idxfield&";"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=conntemp.execute(query)
	
	if x="1" then
		query="UPDATE products SET xfield1="&idxfield&", x1req="&xreq&" WHERE idproduct="&idproduct
	end if
	if x="2" then
		query="UPDATE products SET xfield2="&idxfield&", x2req="&xreq&" WHERE idproduct="&idproduct
	end if
	if x="3" then
    query="UPDATE products SET xfield3="&idxfield&", x3req="&xreq&" WHERE idproduct="&idproduct
	end if
	set rs=conntemp.execute(query)
	set rs=nothing
	call closeDb()
	
	response.redirect "AdminCustom.asp?idproduct="&idproduct&"&message="&Server.URLEncode("Succesfully updated custom input fields for this product.")
	response.end
Else
	call openDb()
	x=request.querystring("x")
	idproduct=request.querystring("idproduct")
	query="SELECT * FROM products WHERE idproduct="&idproduct
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	productName=rs("description")
	xfield1=rs("xfield1")
	x1req=rs("x1req")
	xfield2=rs("xfield2")
	x2req=rs("x2req")
	xfield3=rs("xfield3")
	x3req=rs("x3req")

	'get xfield info
	if x="1" then
		xreq=x1req
		query="SELECT * FROM xfields WHERE idxfield="&xfield1
	end if
	if x="2" then
		xreq=x2req
		query="SELECT * FROM xfields WHERE idxfield="&xfield2
	end if
	if x="3" then
		xreq=x3req
		query="SELECT * FROM xfields WHERE idxfield="&xfield3
	end if
	set rs=conntemp.execute(query)
	idxfield=rs("idxfield")
	xfield=rs("xfield")
	textarea=rs("textarea")
	widthoffield=rs("widthoffield")
	maxlength=rs("maxlength")
	rowlength=rs("rowlength")

	set rs=nothing
	call closeDb()
	%>
	
	<form name="form1" method="post" action="modCustomFields.asp" class="pcForms">
	<input name="x" type="hidden" value="<%=x%>">
	<input name="idxfield" type="hidden" value="<%=idxfield%>">
	<input name="idproduct" type="hidden" value="<%=idproduct%>">
	<table class="pcCPcontent">
		<tr> 
			<td colspan="3">
            
            	<div style="float: right; margin-top: 5px;" class="cpLinksList"><a href="FindProductType.asp?id=<%=idproduct%>" target="_blank">Edit</a> | <a href="../pc/viewPrd.asp?idproduct=<%=idproduct%>&adminPreview=1" target="_blank">Preview</a></div>
                <h2>Product: <strong><%=productName%></strong></h2>
                
				<% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>

				</td>
			</tr>
		<tr> 
			<td width="20%" align="right" nowrap="nowrap">Text to display:</td>
			<td width="80%" colspan="2"><input name="xfield" type="text" id="xfield" value="<%=xfield%>" size="30" maxlength="150">&nbsp;<span class="pcCPnotes">Colons (:) are not permitted in field names.</span></td>
		</tr>
		<tr> 
			<td>&nbsp;</td>
			<td colspan="2">&quot;Text to display&quot; is the name of the input field that the customer will see on the product details page.</font></td>
		</tr>
        <tr>
        	<td colspan="3" class="pcCPspacer"></td>
        </tr>
		<tr> 
			<td align="right">Type of field:</td>
			<td colspan="2"><input type="radio" name="fieldtype" value="1" checked class="clearBorder"> <strong>Text Field</strong></td>
		</tr>
		<tr> 
			<td width="20%">&nbsp;</td>
			<td width="15%" align="right">Length of field: </td>
            <td width="65%"><input name="length1" type="text" id="length1" value="<%=widthoffield%>" size="5" maxlength="10"></td>
		</tr>
		<tr> 
			<td>&nbsp;</td>
			<td align="right">Maximum chars:</td>
            <td><input name="maxchars" type="text" id="maxchars" value="<%=maxlength%>" size="5" maxlength="10"></td>
			</td>
		</tr>
        <tr>
        	<td colspan="3" class="pcCPspacer"></td>
        </tr>
		<tr> 
			<td>&nbsp;</td>
			<td colspan="2"><input type="radio" name="fieldtype" value="2" <% if textarea="-1" then%>checked<%end if%> class="clearBorder"> <strong>Text Area</strong></td>
		</tr>
		<tr> 
			<td>&nbsp;</td>
			<td align="right">Length of field:</td>
            <td><input name="length2" type="text" id="length2" value="<%=widthoffield%>" size="5" maxlength="10"></td>
		</tr>
		<tr> 
			<td>&nbsp;</td>
			<td align="right">Number of rows:</td>
            <td><input name="rows" type="text" id="rows" value="<%=rowlength%>" size="5" maxlength="10"></td>
		</tr>
		<tr> 
			<td>&nbsp;</td>
			<td align="right">Maximum chars:</td>
            <td><input name="maxchars2" type="text" id="maxchars2" value="<%=maxlength%>" size="5" maxlength="10"></td>
		</tr>
        <tr>
        	<td colspan="3" class="pcCPspacer"></td>
        </tr>
		<tr> 
			<td align="right">Required:</td>
			<td colspan="2"> 
				<input type="radio" name="xreq" value="-1" checked class="clearBorder">	Yes 
				<input type="radio" name="xreq" value="0" <% if xreq="0" then%>checked<%end if%> class="clearBorder">	No
			</td>
		</tr>
		<tr> 
			<td>&nbsp;</td>
			<td colspan="2">If you require that this information be entered into the field before allowing a customer to purchase this product, then set the radio button to &quot;Yes&quot;.
			</td>
		</tr>
		<tr> 
			<td colspan="3" class="pcCPspacer"><hr></td>
		</tr>
		<tr> 
        	<td></td>
			<td colspan="2">
			<input type="submit" name="submit" value="Submit" class="submit2">
			&nbsp;<input type="button" name="back" value="Back" onClick="JavaScript: history.go(-1);">
			</td>
		</tr>
	</table>
</form>
<% end if %>
<!--#include file="AdminFooter.asp"-->
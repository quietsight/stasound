<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Removing Custom Product Fields - Confirmation" %>
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
<script>
function newWindow(file,window)
	{
		msgWindow=open(file,window,'resizable=yes,scrollbars=yes,width=400,height=500');
		if (msgWindow.opener == null) msgWindow.opener = self;
	}
</script>
<%
Dim query, conntemp, rs
removaltype=request.QueryString("type")
x=request.QueryString("x")
idproduct=request.QueryString("idproduct")
if idproduct="" then
	idproduct=request.Form("idproduct")
end if

If request.Form("submitRemoveCustom")<>"" then
	If request.Form("removecustom")="1" then
		call openDb()
		set rs=Server.CreateObject("ADODB.Recordset")
		idxfield=request.Form("idxfield")
		'remove all associations in products first
		query="UPDATE products SET xfield1=0, x1req=0 WHERE xfield1="&idxfield&";"
		set rs=conntemp.execute(query)
		query="UPDATE products SET xfield2=0, x2req=0 WHERE xfield2="&idxfield&";"
		set rs=conntemp.execute(query)
		query="UPDATE products SET xfield3=0, x3req=0 WHERE xfield3="&idxfield&";"
		set rs=conntemp.execute(query)
		'remove from idxfields database
		query="DELETE FROM xfields WHERE idxfield="&idxfield&";"
		set rs=conntemp.execute(query)
		set rs=nothing 
		call closeDb()
		'redirect
		response.redirect "manageCFields.asp?s=1&msg=" & Server.URLEncode("Custom input field successfully removed from the database")
		response.end()
	Else
		'redirect back
		response.redirect "AdminCustom.asp?nav="&nav&"&idproduct="&idproduct&"&message="&Server.URLEncode("Succesfully updated custom input fields for this product.")
		response.end()
	End If
End If

If request.Form("submitRemoveSearch")<>"" then
	If request.Form("removeSearch")="1" then
		call openDb()
		set rs=Server.CreateObject("ADODB.Recordset")
		idcustom=request.Form("idcustom")
		'remove all associations in products first
		query="UPDATE products SET custom1=0, content1='' WHERE custom1="&idcustom&";"
		set rs=conntemp.execute(query)
		query="UPDATE products SET custom2=0, content2='' WHERE custom2="&idcustom&";"
		set rs=conntemp.execute(query)
		query="UPDATE products SET custom3=0, content3='' WHERE custom3="&idcustom&";"
		set rs=conntemp.execute(query)
		'remove from idxfields database
		query="DELETE FROM customfields WHERE idcustom="&idcustom&";"
		set rs=conntemp.execute(query)
		set rs=nothing
		call closeDb()
		'redirect back
		response.redirect "manageCFields.asp?message=Custom search field successfully removed from the database"
		response.end()
	Else
		'redirect back
		response.redirect "AdminCustom.asp?nav="&nav&"&idproduct="&idproduct&"&message="&Server.URLEncode("Succesfully updated custom search fields for this product.")
		response.end() 
	End If
End If

If removaltype="2" then
	idxfield=request.QueryString("idxfield")
	call openDb()
	set rs=Server.CreateObject("ADODB.Recordset") 
	if x="1" then
		query="UPDATE products SET xfield1=0, x1req=0 WHERE idproduct="&idproduct
	end if
	if x="2" then
		query="UPDATE products SET xfield2=0, x2req=0 WHERE idproduct="&idproduct
	end if
	if x="3" then
    query="UPDATE products SET xfield3=0, x3req=0 WHERE idproduct="&idproduct
	end if
	set rs=conntemp.execute(query)
	query="SELECT * FROM xfields WHERE idxfield="&idxfield&";"
	set rs=conntemp.execute(query)
	if not rs.eof then
		fieldname=rs("xfield")
	else
		fieldname="N/A (Field no longer exists in the database)"
	end if
	set rs=nothing
	call closeDb()
	%>
		
			<table class="pcCPcontent">
				<tr>
					<td class="pcCPspacer"></td>
				</tr>	
				<tr>
					<td>
						<div class="pcCPmessageSuccess">You have successfully removed the custom input field &quot;<%=fieldname%>&quot; from this product.</div>
						<div style="margin: 20px;">
						Select one of the following options:
						<ul class="pcListIcon">
							<li>Return to the <a href="AdminCustom.asp?idproduct=<%=idproduct%>">Custom Fields</a> page.</li>
							<li style="padding-top: 10px;"><a href="javascript: newWindow('showCFProducts.asp?idcustom=C<%=idxfield%>','products');">Show</a> other products using this custom field</li>
							<li style="padding-top: 10px;"><a href="JavaScript:;" onClick="JavaScript:document.getElementById('removeALL').style.display=''">Remove</a> this field from all products</li>
						</ul>
						</div>

						<div id="removeALL" style="display:none; margin: 0 30px 0 40px;">
						<form name="removeALL" method="post" action="removecustomfields.asp" class="pcForms">
							<div>You are about to <strong>permanently delete</strong> this custom input field from the database. You should not use this feature unless you are absolutely sure that you no longer need this field for this or any other product. All other associations to this field will be removed from all products in your store.</div>
							<div style="padding-top: 10px;">
							<input name="nav" type="hidden" value="<%=nav%>">
							<input name="idxfield" type="hidden" value="<%=idxfield%>">
							<input name="idproduct" type="hidden" value="<%=idproduct%>">
							<input type="hidden" name="removeCustom" value="1">
							<input type="submit" name="submitRemoveCustom" value="Remove &quot;<%=fieldname%>&quot; From All Products" class="submit2">
							</div>
						</form>
						</div>
					</td>
				</tr>
		</table>
<% 
end if
%>
<!--#include file="AdminFooter.asp"-->
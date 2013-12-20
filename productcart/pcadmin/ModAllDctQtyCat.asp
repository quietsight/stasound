<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% pageTitle="Category - Modify Quantity Discounts" %>
<% Section="specials" %>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<% Dim rs, connTemp, query

call openDb()
'~~~~~~~~~~~~~~~ DELETE ~~~~~~~~~~~~~~~~~~~~~
dMode=Request.QueryString("Delete")
if dMode<>"" then
	Session("adminidcategory")=Request.QueryString("idcategory")
	idcategory=Session("adminidcategory")
	iUnitCnt=request("iUnitCnt")
	query="DELETE FROM pcCatDiscounts WHERE pcCD_idcategory="&idcategory 
	Set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	Set rs=Nothing
	call closedb()
	
	Session("adminidcategory")=""
	Session("admindiscountdesc")=""
	
	for i=1 to iUnitCnt
		Session("admindiscountPerUnit"&i)=""
		Session("adminquantityfrom"&i)=""
		Session("adminquantityUntil"&i)=""
		Session("adminidDiscountPerQuantity"&i)=""
	next
	
	response.redirect "viewCatDisc.asp"
end if


'~~~~~~~~~~~~~ SAVE ~~~~~~~~~~~~~~~~~~~~~~~~~~`
sMode=Request.Form("Submit")
If sMode="Save" Then
	iUnitCnt=request("iUnitCnt")
	'save all inputs in temporary session state
	Session("adminidcategory")=Request("idcategory")
	idcategory=Session("adminidcategory")
	Session("admindiscountdesc")=Request("discountdesc")
	Session("adminpercentage")=Request("percentage")
	Session("adminbaseproductonly")=Request("baseproductonly")
	
	for s=1 to iUnitCnt
		
		Session("admindiscountPerUnit" & s)=replacecomma(Request("discountPerUnit" & s))
		Session("admindiscountPerWUnit" & s)=replacecomma(Request("discountPerWUnit" & s))
		if Session("admindiscountPerUnit" & s)="" then
			Session("admindiscountPerUnit" & s)="0"
		end if
		if Session("admindiscountPerWUnit" & s)="" then
			Session("admindiscountPerWUnit" & s)="0"
		end if
		Session("adminquantityfrom" & s)=Request("quantityfrom" & s)
		Session("adminquantityUntil" & s)=Request("quantityuntil" & s)
		Session("adminidDiscountPerQuantity" & s)=Request("idDiscountPerQuantity" & s)
	next
	
	for c=1 to iUnitCnt
		'check to make sure there are no overlaps
		quantityfrom=Request("quantityfrom" & c)
		if quantityfrom="" then
			quantityfrom=0
		end if
		quantityUntil=Request("quantityuntil" & c)
		if quantityUntil="" then
			quantityUntil=99999
		end if

		if quantityfrom <> "" AND quantityUntil <> "" AND replacecomma(Request("discountPerUnit" & c))="" then
			msg="Error: You must specify a discount price for each tier."
			response.redirect "ModAllDctQtyCat.asp?idcategory="&idcategory&"&msg="&msg
		end if
		
		'make sure the from < until
		if int(quantityfrom)>int(quantityUntil) then
			msg="Error: Your quantity ""To"" must be greater then your quantity ""From""."
			response.redirect "ModAllDctQtyCat.asp?idcategory="&idcategory&"&msg="&msg
		end if

		if c<>1 then
			d=c-1
			if NOT quantityfrom="" then
				If Clng(quantityfrom) <> "" AND Clng(quantityfrom)=> Clng(Request("quantityfrom" & d)) AND Clng(quantityfrom) > Clng(Request("quantityUntil" & d)) AND Clng(quantityUntil) <> "" AND Clng(quantityUntil) >= Clng(quantityfrom) then
				else
					msg="Conflict: Your entries are conflicting with each other. It appears that you have created two or more entries that contain at least one value that is the same. You cannot have more then one discount assigned to any one quantity per category."
					response.redirect "ModAllDctQtyCat.asp?idcategory="&idcategory&"&msg="&msg
				end if
			end if
		end if
	next
	
	for i=1 to iUnitCnt
		
		discountPerUnit=replacecomma(Request("discountPerUnit" & i))
		if discountPerUnit="" then
			discountPerUnit="0"
		end if

		discountPerWUnit=replacecomma(Request("discountPerWUnit" & i))				
		if discountPerWUnit="" Then
			discountPerWUnit="0"
		end if

		discountdesc=Request("discountdesc")
		percentage=Request("percentage")
		baseproductonly=Request("baseproductonly")
		quantityfrom=Request("quantityfrom" & i)
		quantityUntil=Request("quantityuntil" & i)
		if quantityfrom="" then
			quantityfrom=0
		end if
		if quantityUntil="" then
			quantityUntil=99999
		end if

		idDiscountPerQuantity=Request("idDiscountPerQuantity" & i)
		idcategory=Request("idcategory")
		
		If (money(discountPerUnit) > 0 OR money(discountPerWUnit)>0) AND quantityfrom <> "" AND quantityUntil <> "" AND idDiscountPerQuantity <> "" Then
			query="UPDATE pcCatDiscounts SET pcCD_discountPerUnit="&discountPerUnit&",pcCD_discountPerWUnit="&discountPerWUnit&",pcCD_quantityfrom="&quantityfrom&", pcCD_quantityuntil= "&quantityuntil&", pcCD_percentage= "&percentage&", pcCD_baseproductonly= "&baseproductonly&" WHERE pcCD_idDiscount="&idDiscountPerQuantity 
			Set rs=Server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)
			Set rs=Nothing
		End If
		
		If quantityfrom="" AND quantityUntil="" AND idDiscountPerQuantity <> "" Then
			query="DELETE FROM pcCatDiscounts WHERE pcCD_idDiscount="&idDiscountPerQuantity 
			Set rs=Server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)
			Set rs=Nothing
		End If
		
		If discountPerUnit <> "" AND quantityfrom <> "" AND quantityUntil <> "" AND idDiscountPerQuantity="" Then
			query="INSERT INTO pcCatDiscounts (pcCD_idcategory,pcCD_discountPerUnit,pcCD_discountPerWUnit,pcCD_quantityuntil,pcCD_quantityfrom,pcCD_num,pcCD_percentage,pcCD_baseproductonly) VALUES ("&idcategory&","&discountPerUnit&","&discountPerWUnit&","&quantityuntil&","&quantityfrom&","&i&","&percentage&","&baseproductonly&");"
			Set rs=Server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)
			Set rs=Nothing
		End If
		
	next
	
	call closedb()
	
	Session("adminidcategory")=""
	Session("admindiscountdesc")=""
	Session("adminpercentage")=""
	Session("adminbaseproductonly")=""
	
	for i=1 to iUnitCnt
		Session("admindiscountPerUnit1")=""
		Session("admindiscountPerWUnit1")=""
		Session("adminquantityfrom1")=""
		Session("adminquantityUntil1")=""
		Session("adminidDiscountPerQuantity1")=""
	next
	
	response.redirect "ModDctQtyCat.asp?idcategory="&idcategory
End If

'~~~~~~~~~~~~~~ SHOW ADMIN ~~~~~~~~~~~~~~~~~~
msg=request.QueryString("msg")
idcategory=request.QueryString("idcategory")
%>


<form method="POST" action="ModAllDctQtyCat.asp" class="pcForms">
	<table class="pcCPcontent">
		<tr>
			<td colspan="5">
                    <% ' START show message, if any %>
                        <!--#include file="pcv4_showMessage.asp"-->
                    <% 	' END show message %>
				<% 
				query="SELECT categoryDesc,idcategory FROM categories WHERE idcategory="&idcategory
				set rsCatObj=server.CreateObject("ADODB.RecordSet") 
				set rsCatObj=conntemp.execute(query)%>
        		<h2>Category Name: <strong><%=rsCatObj("categoryDesc")%></strong> - ID#: <%=rsCatObj("idcategory")%></h2>
				
				<input type="hidden" name="discountdesc" size="40" value="PD">
				<input type="hidden" name="idcategory" size="40" value="<%=idcategory%>">
				<% set rsCatObj=nothing %>
			</td>
		</tr>
		
		<% 
		query="SELECT pcCD_num,pcCD_percentage,pcCD_baseproductonly,pcCD_discountPerUnit,pcCD_discountPerWUnit,pcCD_quantityfrom,pcCD_quantityUntil,pcCD_discountPerUnit,pcCD_IDDiscount FROM pcCatDiscounts WHERE pcCD_idcategory="&idcategory&" ORDER BY pcCD_num"
		Set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if rs.eof then
			msg="This category does not contains quantity discounts, please choose ADD instead."
			set rs=nothing
			call closedb()
			response.redirect "ModAllDctQtyCat.asp?idcategory="&idcategory&"&msg="&msg
		else
			r=rs("pcCD_num")
			Session("adminpercentage")=rs("pcCD_percentage")
			Session("adminbaseproductonly")=rs("pcCD_baseproductonly")
			percentage=Session("adminpercentage")
			baseproductonly=Session("adminbaseproductonly") %>
			<tr> 
				<td colspan="5">Discount based on:<b> 
					<% if percentage="" then %>
						<input type="radio" name="percentage" value="0">
						<%=scCurSign%> 
						<input type="radio" name="percentage" value="-1">
						% 
					<% else %>
						<% if percentage="0" then %>
							<input type="radio" name="percentage" value="0" checked>
							<%=scCurSign%> 
							<input type="radio" name="percentage" value="-1">
							% 
						<%else %>
							<input type="radio" name="percentage" value="0">
							<%=scCurSign%> 
							<input type="radio" name="percentage" value="-1" checked>
							% 
						<% end if %>
					<% end if %></b>
				</td>
			</tr>
			
			<tr> 
				<td colspan="5"> 
					<% if baseproductonly="-1" then %>
						<input type="radio" name="baseproductonly" value="-1" checked>
					<% else 
						if baseproductonly="" then %>
							<input type="radio" name="baseproductonly" value="-1" checked>
						<% else %>
							<input type="radio" name="baseproductonly" value="-1">
						<% end if %>
					<% end if %>
					Apply discount to base price only (product options not included)
				</td>
			</tr>

			<tr> 
				<td colspan="5"> 
					<% if baseproductonly="0" then %>
						<input type="radio" name="baseproductonly" value="0" checked>
					<% else %>
						<input type="radio" name="baseproductonly" value="0">
					<% end if %>
					Apply discount to base price + options prices (if any)
				</td>
			</tr>

			<tr> 
				<td colspan="5" class="pcCPspacer"></td>
			</tr>
			<tr> 
				<th>&nbsp;</th>
				<th>From</th>
				<th>To</th>
				<th><%=scCurSign%> or % (retail)</th>
				<th><%=scCurSign%> or % (wholesale)</th>
			</tr>
			<tr> 
                <td colspan="5" class="pcCPspacer"></td>
			</tr>
			<% do until rs.eof
				Session("admindiscountPerUnit" & r)=rs("pcCD_discountPerUnit")
				Session("admindiscountPerWUnit" & r)=rs("pcCD_discountPerWUnit")
				Session("adminquantityfrom" & r)=rs("pcCD_quantityfrom")
				Session("adminquantityUntil" & r)=rs("pcCD_quantityUntil")
				Session("admindiscountPerUnit" & r)=rs("pcCD_discountPerUnit")
				Session("adminidDiscountPerQuantity" & r)=rs("pcCD_IDDiscount")
				%>
	
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
					<td>Quantity:</td>
					<td> 
					<input name="quantityFrom<%=r%>" size="6" value="<%=Session("adminquantityfrom" & r)%>">
					<input type="hidden" name="idDiscountPerQuantity<%=r%>" size="40" value="<%=Session("adminidDiscountPerQuantity" & r)%>"></td>
					<td><input name="quantityUntil<%=r%>" size="6" value="<%=Session("adminquantityUntil" & r)%>"></td>
					<td><input name="discountPerUnit<%=r%>" size="10" value="<%=money(Session("admindiscountPerUnit" & r))%>"></td>
					<td><input name="discountPerWUnit<%=r%>" size="10" value="<%=money(Session("admindiscountPerWUnit" & r))%>"></td>
				</tr>
	
				<% r=r + 1
				rs.movenext
			loop
			Set rs=Nothing
		call closedb()
		end if
		%>
		<tr> 
			<td colspan="5">&nbsp;</td>
		</tr>
		<tr> 
			<td colspan="5">
            <input type="hidden" name="iUnitCnt" value="<%=r-1%>">
			<input type="submit" name="Submit" value="Save" class="submit2">&nbsp; 
			<input type="button" name="Delete" value="Delete discount" onClick="javascript:if (confirm('You are about to permanantly delete this discount from the database. Are you sure you want to complete this action?')) location='modAllDctQtyCat.asp?Delete=Yes&idcategory=<%=idcategory%>'">&nbsp;
			<input type="button" name="Button" value="Back" onClick="javascript:history.back()">&nbsp;
			</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->
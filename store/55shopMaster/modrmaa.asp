<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="View/Update RMA Information" %>
<% section="mngRma"%>
<%PmAdmin=7%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/currencyformatinc.asp" -->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/SQLFormat.txt" -->
<%
dim conntemp, query, rs, rstemp, rstemp2
pidRMA=trim(request("idRMA"))
if not validNum(pidRMA) then response.Redirect("resultsAdvancedAll.asp?B1=View+All&dd=1")

if request.form("Modify")<>"" then
	'update reason, comments, rmaIdProducts
	if pidRMA="" then
		pidRMA=request.Form("idRMA")
	end if
	idOrder=request.Form("idOrder")
	rmaReturnReason=replace(request.form("rmaReturnReason"),"'","''")
	rmaReturnStatus=replace(request.form("rmaReturnStatus"),"'","''")
	rmaApproved=request.form("rmaApproved")
	genNumber=request.form("genNumber")
	iCnt=request.form("iCnt")
	pSendEmail=request.form("sendEmail")
	pRMADate=now()
	if SQL_Format = "1" then
	    pRMADate=(Day(pRMADate)&"/"&Month(pRMADate)&"/"&Year(pRMADate))
	else
	    pRMADate=(Month(pRMADate)&"/"&Day(pRMADate)&"/"&Year(pRMADate))
	end if
	call opendb()
	for i=1 to iCnt
		'update productOrders with return flag
		rmaSubmitted=trim(request.form("rmaSubmitted"&i))
		if Not isNumeric(rmaSubmitted) then
			rmaSubmitted=0
		end if
		query="UPDATE productsOrdered SET rmaSubmitted="&rmaSubmitted&" WHERE idProductOrdered="&request.form("idProductOrdered"&i)&";"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
	next
	set rs=nothing
	
	if genNumber="YES" then
		function makePassword(byVal maxLen)
			Dim strNewPass
			Dim whatsNext, upper, lower, intCounter
			Randomize
	
			For intCounter = 1 To maxLen
				whatsNext = Int((1 - 0 + 1) * Rnd + 0)
				If whatsNext = 0 Then
					'character
					upper = 90
					lower = 65
				Else
					upper = 57
					lower = 48
				End If
				strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
			Next
			makePassword = strNewPass
		end function
		pRmaNumber = request("pRmaNumber1")
		if pRmaNumber<>"" then
			pRmaNumber=replace(pRmaNumber,"'","''")
		else
			pRmaNumber = makePassword(16)
		end if
		rmaApproved="1"
	end if
		
	query="UPDATE PCReturns set "
	if genNumber="YES" then
		query=query&"rmaNumber='"&pRmaNumber&"',"
	end if
	query=query&"rmaReturnReason='"&rmaReturnReason&"', rmaReturnStatus='"&rmaReturnStatus&"', rmaApproved="&rmaApproved&" WHERE idRMA="&pidRMA&";"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	'update amount credited to account
	rmaCredit=trim(request.Form("rmaCredit"))
	if NOT isNumeric(rmaCredit) then
		rmaCredit=0
	end if
	
	if instr(rmaCredit,",")>0 then
		rmaCredit=replacecomma(rmaCredit)
	end if
	
	query="UPDATE orders SET rmaCredit="&rmaCredit&" WHERE idOrder="&idOrder&";"
	set rs=conntemp.execute(query)
	set rs=nothing
	
	'send out email
	if pSendEmail="1" then
		
		rmaReturnStatus=replace(rmaReturnStatus,"''","'")

		'// Get Order details
		query="SELECT orders.idOrder, orders.idCustomer, customers.name, customers.lastName, customers.email FROM customers INNER JOIN orders ON customers.idcustomer = orders.idCustomer WHERE (((orders.idOrder)="&idOrder&"));"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		pFirstName=rs("name")
		pLastName=rs("lastName")
		pRcpt=rs("email")
		
		'// Get RMA number
		query="SELECT rmaNumber FROM PCReturns WHERE idRMA="&pidRMA&";"	
		set rs=conntemp.execute(query)
		if rs.eof then
			prmaNumber=""
		else
			prmaNumber=rs("rmaNumber")
		end if
		set rs=nothing
		
		msgTitle= dictLanguage.Item(Session("language")&"_sendMail_rma_1") & scCompanyName
		MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_sendMail_rma_7") & VBCRLF
		select case rmaApproved
		case 1
			'If the RMA has been approved:
			MsgBody=MsgBody & "" & VBCRLF
			if rmaApproved="1" then
				MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_sendMail_rma_8") & VBCRLF
			else
				MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_sendMail_rma_9") & VBCRLF
			end if
			MsgBody=MsgBody & "" & VBCRLF
			if prmaNumber<>"" then
				MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_sendMail_rma_4") & pRmaNumber & VBCRLF
			end if
			MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_sendMail_rma_5") & ShowDateFrmt(pRMADate) & VBCRLF
			MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_sendMail_rma_6") & rmaReturnStatus & VBCRLF
			MsgBody=MsgBody & "" & VBCRLF
		case 2
			'If the RMA request has been denied:
			MsgBody=MsgBody & "" & VBCRLF
			MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_sendMail_rma_10") & VBCRLF
			MsgBody=MsgBody & "" & VBCRLF
			MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_sendMail_rma_5") & ShowDateFrmt(pRMADate) & VBCRLF
			MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_sendMail_rma_6") & rmaReturnStatus & VBCRLF
			MsgBody=MsgBody & "" & VBCRLF
		case 0			
			'If new updates are posted:
			MsgBody=MsgBody & "" & VBCRLF
			MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_sendMail_rma_11") & VBCRLF
			MsgBody=MsgBody & "" & VBCRLF
			if prmaNumber<>"" then
				MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_sendMail_rma_4") & pRmaNumber & VBCRLF
			end if
			MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_sendMail_rma_5") & ShowDateFrmt(pRMADate) & VBCRLF
			MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_sendMail_rma_6") & rmaReturnStatus & VBCRLF
			MsgBody=MsgBody & "" & VBCRLF
		end select
		call sendmail (scCompanyName,scFrmEmail,pRcpt,MsgTitle,MsgBody)
	end if
	call closedb()
end if

'// Retrieve RMA information
call openDb()
query="SELECT PCReturns.idRMA, PCReturns.rmaNumber, PCReturns.rmaReturnReason, PCReturns.rmaDateTime, PCReturns.rmaReturnStatus, PCReturns.idOrder, PCReturns.rmaIdProducts, PCReturns.rmaApproved, customers.name, customers.lastName, customers.phone, customers.email FROM PCReturns INNER JOIN (customers INNER JOIN orders ON customers.idcustomer = orders.idCustomer) ON PCReturns.idOrder = orders.idOrder WHERE (((PCReturns.idRMA)="&pidRMA&"));"	
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(query)

if err.number <> 0 then
	set rstemp=nothing
	response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&Err.Description) 
	call closeDb()
end If

pidRMA=rstemp("idRMA")
prmaNumber=rstemp("rmaNumber")
prmaReturnReason=rstemp("rmaReturnReason")
	prmaReturnReason=replace(prmaReturnReason,"''","'")
prmaDateTime=rstemp("rmaDateTime")
prmaReturnStatus=rstemp("rmaReturnStatus")
pIdOrder=rstemp("idOrder")
prmaIdProducts=rstemp("rmaIdProducts")
prmaApproved=rstemp("rmaApproved")
pname=rstemp("name")
plastName=rstemp("lastName")
pphone=rstemp("phone")
pemail=rstemp("email")
set rstemp=nothing
call closedb()
%>
<!--#include file="AdminHeader.asp"-->

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<form method="post" name="modRma" action="modRmaa.asp" class="pcForms">
	<table class="pcCPcontent">
	<tr>
		<td><div align="right">Order ID:</div></td>
		<td><a href="Orddetails.asp?id=<%=pIdOrder%>"><%=(int(pIdOrder)+scpre)%></a></td>
	</tr>
	<tr>
		<td><div align="right">Customer Name:</div></td>
		<td><%=pName&"&nbsp;"&plastName%> </td>
	</tr>
	<tr> 
	<td><div align="right">Date Submitted:</div></td>
	<td><%=showDateFrmt(prmaDateTime)%></td>
	</tr>
	<%
	if prmaNumber<>"" then
		prmaexists=1 %>
		<tr> 
		<td><div align="right">RMA Number:</div></td>
		<td><%=prmaNumber%></td>
		</tr>
        <tr>
            <td><div align="right"><strong>Status</strong>:</div></td>
            <td><input type="radio" name="rmaApproved" value="0" checked onclick="document.modRma.genNumber.disabled=false;" class="clearBorder">
                Pending
                    <% if prmaApproved="1" then %>
                    <input type="radio" name="rmaApproved" value="1" checked onclick="document.modRma.genNumber.disabled=false;" class="clearBorder">
                    <% else %>
                    <input type="radio" name="rmaApproved" value="1" onclick="document.modRma.genNumber.disabled=false;" class="clearBorder">
                    <% end if %>
                Approved
                <% if prmaApproved="2" then %>
                <input type="radio" name="rmaApproved" value="2" checked onclick="javascript:document.modRma.genNumber.checked=false;document.modRma.genNumber.disabled=true;document.modRma.pRmaNumber1.value='';" class="clearBorder">
                <% else %>
                <input type="radio" name="rmaApproved" value="2" onclick="javascript:document.modRma.genNumber.checked=false;document.modRma.genNumber.disabled=true;document.modRma.pRmaNumber1.value='';" class="clearBorder">
                <% end if %>
                Denied</td>
        </tr>
	<%
	else
		prmaexists=0
	end if %>
	<tr> 
	<td align="right" valign="top">Return Reason:</td>
	<td align="left">  
	<textarea cols="50" rows="5" name="rmaReturnReason"><%=prmaReturnReason%></textarea>
	<% if msg="" then %>
	<img src="images/pc_required.gif" width="9" height="9"> 
	<% else
	if prmaReturnReason="" then %>
	<img src="images/prev.gif" width="10" height="10"> 
	<% end if %>
	<% end if %>
	</td>
	</tr> 
	<tr valign="top">
		<td><div align="right">Products:</div></td>
		<td>
			<table width="100%" border="0" cellspacing="0" cellpadding="4">
			<tr bgcolor="#E1E1E1">
				<td>SKU</td>
				<td>Product</td>
			</tr>
			<% idArray=split(prmaIdProducts,",")
			call openDb()
			for i = 0 to ubound(idArray)
				query="SELECT idProduct,SKU,Description FROM products WHERE idProduct="&idArray(i)
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=conntemp.execute(query)
				pSKU=rs("SKU")
				pDescription=rs("Description") %>
				<tr>
					<td width="17%"><%=pSKU%></td>
					<td width="83%"><%=pDescription%></td>
				</tr>
			<%
				next
				set rs=nothing	
			%>
		</table>
	</td>
	</tr>
	<tr>
		<td colspan="2">
		<hr>
		<input type="hidden" name="idOrder" value="<%=pIdOrder%>">
		<input type="hidden" name="idRMA" value="<%=pIdRMA%>">
	</td>
	</tr>
	<% if prmaexists=0 then %>
	<tr>
		<td><div align="right"><strong>Status</strong>:</div></td>
		<td><input type="radio" name="rmaApproved" value="0" checked onclick="document.modRma.genNumber.disabled=false;" class="clearBorder">
			Pending
				<% if prmaApproved="1" then %>
				<input type="radio" name="rmaApproved" value="1" checked onclick="document.modRma.genNumber.disabled=false;" class="clearBorder">
				<% else %>
				<input type="radio" name="rmaApproved" value="1" onclick="document.modRma.genNumber.disabled=false;" class="clearBorder">
				<% end if %>
			Approved
			<% if prmaApproved="2" then %>
			<input type="radio" name="rmaApproved" value="2" checked onclick="javascript:document.modRma.genNumber.checked=false;document.modRma.genNumber.disabled=true;document.modRma.pRmaNumber1.value='';" class="clearBorder">
			<% else %>
			<input type="radio" name="rmaApproved" value="2" onclick="javascript:document.modRma.genNumber.checked=false;document.modRma.genNumber.disabled=true;document.modRma.pRmaNumber1.value='';" class="clearBorder">
			<% end if %>
			Denied</td>
	</tr>
	<%
		function makePassword1(byVal maxLen)
				Dim strNewPass
				Dim whatsNext, upper, lower, intCounter
				Randomize
		
				For intCounter = 1 To maxLen
					whatsNext = Int((1 - 0 + 1) * Rnd + 0)
					If whatsNext = 0 Then
						'character
						upper = 90
						lower = 65
					Else
						upper = 57
						lower = 48
					End If
					strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
				Next
				makePassword1 = strNewPass
			end function
	
			pRmaNumber1 = makePassword1(16)
		%>
	<tr>
		<td colspan="2"><hr></td>
	</tr>
	<tr>
		<td align="right" nowrap="nowrap">Create an <strong>RAM Number</strong>?</td>
		<td valign="top">Yes <input onclick="javascript: if ((this.checked==true) && (document.modRma.pRmaNumber1.value=='')) {document.modRma.pRmaNumber1.value=StrpRmaNumber1;} else {if (this.checked==false) document.modRma.pRmaNumber1.value='';}" name="genNumber" type="checkbox" id="genNumber" value="YES" <% if prmaApproved<>"2" then %>checked<%end if%> <% if prmaApproved="2" then %>disabled<%end if%> class="clearBorder">
		<script>
			StrpRmaNumber1="<%=pRmaNumber1%>";
		</script>
	</td>
	</tr>
	<tr>
	  <td align="right">RAM Number</td>
		<td><input name="pRmaNumber1" value="<%if trim(prmaNumber)="" AND prmaApproved<>"2" then%><%=pRmaNumber1%><%end if%>" type="text" size="20"> <em>(a random RMA was created for you)</em></td>
	</tr>
	<tr>
		<td colspan="2"><hr></td>
	</tr>
	<tr>
		<td valign="top"><div align="right"><strong>Comments</strong>:</div></td>
		<td align="left" valign="top">
			<textarea cols="50" rows="5" name="rmaReturnStatus" value="<%=prmaReturnStatus%>" size="25"><%=prmaReturnStatus%></textarea>
		</td>
	</tr>
	<% else %>
	<tr>
		<td colspan="2" align="center"><div align="left"><strong>Update Order Information</strong></div></td>
	</tr>
	<tr>
		<td colspan="2" align="center"><div align="left">If a product has been returned and you need to credit the corresponding order, flag the product as being returned and credit the account for the appropriate amount. This information will be shown to the customer when they log into their account and view information about that order, along with your &quot;Comments/Status&quot;</div></td>
	</tr>
	<% 
		query="SELECT orders.idOrder, orders.rmaCredit, products.idProduct, products.description, products.sku, ProductsOrdered.rmaSubmitted, ProductsOrdered.unitPrice, ProductsOrdered.quantity, ProductsOrdered.idProductOrdered FROM (ProductsOrdered INNER JOIN products ON ProductsOrdered.idProduct = products.idProduct) INNER JOIN orders ON ProductsOrdered.idOrder = orders.idOrder WHERE (((orders.idOrder)="&pIdOrder&"));"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
	%>
	<tr>
		<td colspan="2" align="center"><table width="100%" border="0" cellspacing="0" cellpadding="4">
			<tr bgcolor="#E1E1E1">
				<td width="5%">QTY</td>
				<td width="15%">SKU</td>
				<td width="50%">Product</td>
				<td width="15%"><div align="center">Unit Price</div></td>
				<td width="15%"><div align="center">Units Returned</div></td>
			</tr>
			<% iCnt=0
			do until rs.eof
				iCnt=iCnt+1
				prmaCredit=rs("rmaCredit")
				pIdProduct=rs("idProduct")
				pdescription=rs("description")
				pSKU=rs("sku")
				prmaSubmitted=rs("rmaSubmitted")
				if isNull(prmaSubmitted) OR prmaSubmitted="" then
					prmaSubmitted=0
				end if
				punitPrice=rs("unitPrice")
				pquantity=rs("quantity")
				pidProductOrdered=rs("idProductOrdered") %>
				<tr>
					<td><%=pquantity%></td>
					<td><%=pSKU%></td>
					<td><%=pDescription%></td>
					<td><div align="center"><%=scCurSign&money(pUnitPrice)%></div></td>
					<td><div align="center">
						<input name="rmaSubmitted<%=iCnt%>" type="text" value="<%=prmaSubmitted%>" size="4" maxlength="4">
						<input name="idProductOrdered<%=iCnt%>" type="hidden" value="<%=pidProductOrdered%>">
						<input name="idProduct<%=iCnt%>" type="hidden" value="<%=pIdProduct%>">
					</div></td>
				</tr>
				<%
			rs.movenext
			loop
			set rs=nothing
			call closedb()
			%>
			<input name="iCnt" type="hidden" value="<%=iCnt%>">
		</table>
		</td>
	</tr>
	<% if IsNull(prmaCredit) OR prmaCredit="" then
	prmaCredit="0"
	end if %>
	<tr>
		<td><div align="right">Credit Account:</div></td>
		<td><%=scCurSign%>&nbsp;<input type="text" name="rmaCredit" value="<%=money(prmaCredit)%>" size="10" maxlength="20"></td>
	</tr>
	<tr> 
	<td  width="18%" valign="top"><div align="right">Comments:</div></td>
	<td  width="82%" align="left" valign="top">  
	<textarea cols="40" rows="5" name="rmaReturnStatus" value="<%=prmaReturnStatus%>" size="25"><%=prmaReturnStatus%></textarea>
    </td>
	</tr>
	<% end if %>
	<tr>
		<td align="center"><div align="right">
				<input type="checkbox" name="sendEmail" value="1" class="clearBorder">
			</div>
		</td>
		<td>Send e-mail to customer.</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr> 
	<td colspan="2" align="center">
		<input type="submit" name="Modify" class="submit2" value="Create/Update">
		&nbsp;
		<% if pIdOrder<>"" then %>
		<input type="button" name="OrderDetails" value="Back to Order Details" class="ibtnGrey" onclick="location.href='Orddetails.asp?id=<%=pIdOrder%>'">
		<% end if %>
		</td>
	</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Modify Custom Payment Option" %>
<% section="paymntOpt" %>
<%PmAdmin=5%><!--#include file="adminv.asp"-->  
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<%
dim query, conntemp, rstemp

pidCustomCardType=request.QueryString("idc")
if NOT isNumeric(pidCustomCardType) then
	pidCustomCardType=0
end if
pidPayment=Request.QueryString("id")
pgwCode= Request.QueryString("gwCode")
call openDb()

if request.form("SubmitOrder")="Update Order" then
	iCnt=request.form("TotalCnt")
	pidCustomCardType=request.form("idc")
	pidPayment=Request.form("id")
	pgwCode= Request.form("gwCode")
	for i=1 to iCnt
		intOrder=request.form("intOrder"&i)
		idcustomcardrule=request.form("id"&i)
		query="UPDATE customCardRules SET intOrder="&intOrder&" WHERE idCustomCardRules="&idcustomcardrule
		set rstemp=Server.CreateObject("ADODB.Recordset")     
		set rstemp=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rstemp=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If	
	next
	Cbtob=request.Form("Cbtob")
	if Cbtob = "" then
		Cbtob = "0"
	end if
	if Cbtob="1" then
		Cbtob="-1"
	end if
	if Cbtob="2" then
		Cbtob2 = request("Cbtob2")
		if Cbtob2 = "" then
			Cbtob = "0"
		end if
	end if

	for i=1 to iCnt
		intOrder=request.form("intOrder"&i)
		idcustomcardrule=request.form("id"&i)
		query="UPDATE customCardRules SET intOrder="&intOrder&" WHERE idCustomCardRules="&idcustomcardrule
		set rstemp=Server.CreateObject("ADODB.Recordset")     
		set rstemp=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rstemp=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If	
	next
	customCardDesc=replace(request.Form("customCardDesc"),"'","''")
	pPriceToAddType=request.Form("priceToAddType")
	If pPriceToAddType="price" Then
		pPriceToAdd=replacecomma(Request("priceToAdd"))
		pPercentageToAdd="0"
		if pPriceToAdd="" then
			pPriceToAdd="0"
		end if
	Else
		pPercentageToAdd=request.Form("percentageToAdd")
		pPriceToAdd="0"
		if pPercentageToAdd="" then
			pPercentageToAdd="0"
		end if
	End If
	query="UPDATE customCardTypes SET customCardDesc='"&customCardDesc&"' WHERE idcustomCardType="&pidCustomCardType&";"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
			strErrorDescription=err.description
			set rstemp=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	'Start SDBA
		pcv_processOrder=request.Form("pcv_processOrder")
		if pcv_processOrder="" then
			pcv_processOrder="0"
		end if
		pcv_setPayStatus=request.Form("pcv_setPayStatus")
		if pcv_setPayStatus="" then
			pcv_setPayStatus="0"
		end if
	'End SDBA
	
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",paymentDesc='"&customCardDesc&"', Cbtob="&Cbtob&", priceToAdd="&pPriceToAdd&", percentageToAdd="&pPercentageToAdd&" WHERE idpayment="&pidPayment&";"

	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
			strErrorDescription=err.description
			set rstemp=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	query = "DELETE FROM CustCategoryPayTypes WHERE idPayment = "&pidPayment&";"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	
	if Cbtob="2" then
		CbtobArray=split(Cbtob2,",")
		for t=lbound(CbtobArray) to ubound(CbtobArray)
			query = "INSERT INTO CustCategoryPayTypes (idCustomerCategory,idPayment) VALUES ("&CbtobArray(t)&" ,"&pidPayment&");"
			set rstemp=Server.CreateObject("ADODB.Recordset")     
			set rstemp=conntemp.execute(query)
		next
	end if
		
	call closeDb()
	response.redirect "modcustomCardPaymentOpt.asp?id="&pidPayment&"&idc="&pidCustomCardType&"&gwCode="&pgwcode
end if

iMode=Request.queryString("mode")
If iMode="Del" Then
	query="DELETE FROM paytypes WHERE idPayment="&pidPayment
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
			strErrorDescription=err.description
			set rstemp=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="DELETE FROM customCardRules WHERE idCustomCardType="&pidCustomCardType
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
			strErrorDescription=err.description
			set rstemp=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="DELETE FROM customCardTypes WHERE idCustomCardType="&pidCustomCardType
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
			strErrorDescription=err.description
			set rstemp=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	set rstemp=nothing
	call closeDb()
	response.redirect "paymentOptions.asp"
End If

sMode=Request.Form("Submit")
If sMode <> "" Then
	iCnt=request.form("TotalCnt")
	pidCustomCardType=request.form("idc")
	pidPayment=Request.form("id")
	pgwCode= Request.form("gwCode")
	Cbtob=request.Form("Cbtob")
	if Cbtob = "" then
		Cbtob = "0"
	end if
	if Cbtob="1" then
		Cbtob="-1"
	end if
	if Cbtob="2" then
		Cbtob2 = request("Cbtob2")
		if Cbtob2 = "" then
			Cbtob = "0"
		end if
	end if

	for i=1 to iCnt
		intOrder=request.form("intOrder"&i)
		idcustomcardrule=request.form("id"&i)
		query="UPDATE customCardRules SET intOrder="&intOrder&" WHERE idCustomCardRules="&idcustomcardrule
		set rstemp=Server.CreateObject("ADODB.Recordset")     
		set rstemp=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rstemp=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If	
	next
	customCardDesc=replace(request.Form("customCardDesc"),"'","''")
	pPriceToAddType=request.Form("priceToAddType")
	If pPriceToAddType="price" Then
		pPriceToAdd=replacecomma(Request("priceToAdd"))
		pPercentageToAdd="0"
		if pPriceToAdd="" then
			pPriceToAdd="0"
		end if
	Else
		pPercentageToAdd=request.Form("percentageToAdd")
		pPriceToAdd="0"
		if pPercentageToAdd="" then
			pPercentageToAdd="0"
		end if
	End If
	
	query="UPDATE customCardTypes SET customCardDesc='"&customCardDesc&"' WHERE idcustomCardType="&pidCustomCardType&";"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	'Start SDBA
		pcv_processOrder=request.Form("pcv_processOrder")
		if pcv_processOrder="" then
			pcv_processOrder="0"
		end if
		pcv_setPayStatus=request.Form("pcv_setPayStatus")
		if pcv_setPayStatus="" then
			pcv_setPayStatus="0"
		end if
	'End SDBA
	
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",paymentDesc='"&customCardDesc&"', Cbtob="&Cbtob&", priceToAdd="&pPriceToAdd&", percentageToAdd="&pPercentageToAdd&" WHERE idpayment="&pidPayment&";"
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)
	
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
		
	query = "DELETE FROM CustCategoryPayTypes WHERE idPayment = "&pidPayment&";"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	
	if Cbtob="2" then
		CbtobArray=split(Cbtob2,",")
		for t=lbound(CbtobArray) to ubound(CbtobArray)
			query = "INSERT INTO CustCategoryPayTypes (idCustomerCategory,idPayment) VALUES ("&CbtobArray(t)&" ,"&pidPayment&");"
			set rstemp=Server.CreateObject("ADODB.Recordset")     
			set rstemp=conntemp.execute(query)
		next
	end if
	
	set rstemp=nothing
	call closeDb()
	response.redirect "PaymentOptions.asp"
End If

if pidCustomCardType="" then
	pidCustomCardType=request.querystring("idcustomCardType")
end if
if pidCustomCardType<>0 then
	query= "SELECT customCardTypes.idcustomCardType, customCardTypes.customCardDesc, payTypes.gwCode, payTypes.idPayment, payTypes.priceToAdd, payTypes.percentageToAdd, payTypes.Cbtob,payTypes.pcPayTypes_processOrder,payTypes.pcPayTypes_setPayStatus FROM customCardTypes, payTypes WHERE (((customCardTypes.idcustomCardType)="&pidCustomCardType&") AND ((payTypes.gwCode)=10"&pidCustomCardType&"));"
	
	set rstemp=Server.CreateObject("ADODB.Recordset")
	
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	pcustomCardDesc=rstemp("customCardDesc")
	pCbtob=rstemp("Cbtob")
	percentageToAdd=rstemp("percentageToAdd")
	priceToAdd=rstemp("priceToAdd")
		tempidPayment=rstemp("idPayment")
	if pidPayment="" then
		pidPayment=tempidPayment
	end if
	
	'Start SDBA
	pcv_processOrder=rstemp("pcPayTypes_processOrder")
	pcv_setPayStatus=rstemp("pcPayTypes_setPayStatus")
	'End SDBA
	
	set rstemp=nothing
end if

'Get Customer Categories
query="SELECT idCustomerCategory, pcCC_Name FROM pcCustomerCategories Order by pcCC_Name;"
set rs=Server.CreateObject("ADODB.Recordset")     
rs.Open query, conntemp
if err.number <> 0 then
	strErrDescription = err.description
	set rs=nothing
	call closeDb()
	response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCCPaymentOpt 345: "&strErrDescription) 
end If

pcv_NoCPC=1
pcv_CPCString = ""

do until rs.eof
	pcv_NoCPC=0
	pcv_IdCustomerCategory = rs("idCustomerCategory")
	pcv_CCName = rs("pcCC_Name")
	'check if this is selected for the CC already.
	query = "SELECT * FROM CustCategoryPayTypes WHERE idPayment="&pidPayment&" AND idCustomerCategory="&pcv_IdCustomerCategory&";"
	set rsCPCObj=Server.CreateObject("ADODB.Recordset")     
	set rsCPCObj=connTemp.execute(query)
	if rsCPCObj.eof then
		pcv_Checked=""
	else
		pcv_Checked=" checked"
	end if
	set rsCPCObj=nothing
	pcv_CPCString = pcv_CPCString&"<tr><td width=""6%""><input type=""checkbox"" name=""Cbtob2"" value='"&pcv_IdCustomerCategory&"' class='clearBorder'"&pcv_Checked&"></td><td width=""94%"">Apply to "&pcv_CCName&"</td></tr>"
	rs.movenext
loop
set rs=nothing
%>
<!--#include file="AdminHeader.asp"-->

	<% if pidCustomCardType=0 then %>
		<div class="pcCPmessage">Invalid input: use the browser's back button and please try again.</div>
	<% else %>
	
<form method="post" name="form1" action="modCustomCardPaymentOpt.asp" class="pcForms">
	<table class="pcCPcontent">
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2">Enter a description for this payment option and specify whether it applies to all customers or only wholesale customers.&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=303')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
		</tr>
		<tr>
			<td colspan="2">
				<p>Option Name: <input name="customCardDesc" type="text" value="<%=replace(pcustomCardDesc,"""","&quot;")%>" size="30" maxlength="70"></p>
			</td>
		</tr>
                <tr bgcolor="#FFFFFF">
                  <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="2%" align="right"><input type="radio" name="Cbtob" id="Cbtob" value="0" <% if pCbtob="0" then%>checked<% end if %> /></td>
                      <td width="98%">Apply to all customers</td>
                    </tr>
                    <tr>
                      <td width="2%" align="right"><input type="radio" name="Cbtob" id="Cbtob" value="1" <% if pCbtob="-1" then%>checked<% end if %>  /></td>
                      <td width="98%">Apply to Wholesale Customers only</td>
                    </tr>
                    <% if pcv_CPCString&""<>"" then %>
                    <tr>
                      <td width="2%" align="right"><input type="radio" name="Cbtob" id="Cbtob" value="2" <% if pCbtob="2" then%>checked<% end if %> /></td>
                      <td width="98%">Apply to the following customer pricing categories</td>
                    </tr>
                    <tr>
                      <td align="right">&nbsp;</td>
                      <td><table width="95%" border="0" cellspacing="0" cellpadding="2">
                        <%=pcv_CPCString%>
                      </table></td>
                    </tr>
                    <% end if %>
                    <tr>
                      <td colspan="2" height="10"></td>
                    </tr>
                  </table></td>
                </tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2" bgcolor="#e5e5e5">Input Fields - <a href="OptionFieldsAdd.asp?idc=<%=pidCustomCardType%>&id=<%=pidPayment%>">Add New</a></td>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2">
			<table width="60%" cellpadding="4" cellspacing="0" style="border: 1px solid #e1e1e1">
			<%
			query="Select idCustomCardRules, ruleName, intRuleRequired, intOrder from customCardRules WHERE idcustomCardType="&pidcustomCardType&" ORDER BY intorder"
			set rstemp=Server.CreateObject("ADODB.Recordset")
			set rstemp=connTemp.execute(query)
			if rstemp.eof then
			%>
			<tr>								
				<td colspan="4"><div class="pcCPmessage">No Input Fields Found</div></td>
			</tr>									
			<%else %>									
			<tr>									
				<th>Field name</th>
				<th align="center">Required</th>
				<th align="center">Order</th>
				<th align="center">Action</th>
			</tr>
            <tr> 
                <td colspan="4" class="pcCPspacer"></td>
            </tr>
			<% iCnt=0
			do while not rstemp.eof
				pidCustomCardRule=rstemp("idCustomCardRules")
				pruleName=rstemp("ruleName")
				pintrulerequired=rstemp("intRuleRequired")
				pintOrder=rstemp("intOrder")
				iCnt=iCnt+1
			%>
			<tr>
				<td width="55%"><p><%=pruleName%></p></td>
				<td width="15%" align="center">
					<%if pintrulerequired="-1" then%><img src="images/pc_required.gif" width="9" height="9"><% else %>&nbsp;<% end if %>
				</td>
				<td width="10%" align="center">
					<input name="intOrder<%=iCnt%>" type="text" size="2" maxlength="4" value="<%=pintOrder%>">
					<input name="id<%=iCnt%>" type="hidden" value="<%=pidCustomCardRule%>">
				</td>
				<td width="25%" nowrap="nowrap" align="center"><a href="OptionFieldsEdit.asp?idc=<%=pidCustomCardType%>&id=<%=pidPayment%>&idccr=<%=pidCustomCardRule%>">Edit</a> - <a href="javascript:if (confirm('You are about to remove this option from your database. Are you sure you want to complete this action?')) location='OptionFieldsEdit.asp?m=del&idc=<%=pidCustomCardType%>&id=<%=pidPayment%>&idccr=<%=pidCustomCardRule%>'">Delete</a></td>
			</tr>
				<%
					rstemp.movenext
					loop
					end if
					set rstemp=nothing
					call closeDb()
				%>
			<tr>
				<td colspan="4"><hr></td>
			</tr>
			<tr>
				<td colspan="2"></td>
				<td colspan="2">
					<input name="TotalCnt" type="hidden" id="TotalCnt" value="<%=iCnt%>">
					<input name="idc" type="hidden" value="<%=pidCustomCardType%>">
					<input name="id" type="hidden" value="<%=pidPayment%>">
					<input name="gwCode" type="hidden" value="<%=pgwCode%>">
					<input name="SubmitOrder" type="submit" id="SubmitOrder" value="Update Order" class="submit2">
				</td>
			</tr>
			</table>
			</td>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2" bgcolor="#e5e5e5">Additional Fees</td>
		</tr>
		<tr> 
			<td colspan="2">You have the option to charge a processing fee for this payment option.&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=304')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
		</tr>
		<tr>
			<td>Processing fee:</td>
			<td>
			<% if percentageToAdd="0" then %>
				<input type="radio" name="priceToAddType" value="price" checked class="clearBorder">
			<% else %>
				<input type="radio" name="priceToAddType" value="price" class="clearBorder">
			<% end if %>
			 Flat Fee&nbsp;&nbsp;<%=scCurSign%>
			<input name="priceToAdd" size="6" value="<%=money(priceToAdd)%>">
			</td>
		</tr>
		<tr> 
			<td>&nbsp; </td>
			<td width="80%">
			<% if percentageToAdd<>"0" then %>
				<input type="radio" name="priceToAddType" value="percentage" checked class="clearBorder">
			<% else %>
				<input type="radio" name="priceToAddType" value="percentage" class="clearBorder">
			<% end if %>
			 Percentage of Order Total&nbsp;&nbsp;% 
				<input name="percentageToAdd" size="6" value="<%=percentageToAdd%>">
			</td>
		</tr>

		<%'Start SDBA%>
		<tr>
			<td colspan="4" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="4" bgcolor="#e5e5e5">Order Processing: Order Status and Payment Status</td>
		</tr>
		<tr>
			<td colspan="2">Process orders when they are placed: <input type="checkbox" name="pcv_processOrder" value="1" <%if pcv_processOrder="1" then%>checked<%end if%> class="clearBorder">&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=301')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
		</tr>
		<tr>
			<td colspan="2">When orders are placed, set the payment status to:
				<select name="pcv_setPayStatus">
					<option value="0" selected>Pending</option>
					<option value="1" <%if pcv_setPayStatus="1" then%>selected<%end if%>>Authorized</option>
					<option value="2" <%if pcv_setPayStatus="2" then%>selected<%end if%>>Paid</option>
				</select>
				&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=302')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>
			</td>
		</tr>
	<%'End SDBA%>

	<input type="hidden" name="idPayment" value="<%=idPayment%>">
	<input type="hidden" name="gwCode" value="<%=gwCode%>">

		<tr>
			<td colspan="2"><hr></td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td colspan="2" align="center"> 
				<input type="submit" name="Submit" value="Update" class="submit2">
				&nbsp;<input type="button" name="back" value="Back" onClick="javascript:history.back()">
				&nbsp;<input type="button" value="Payment Options Summary" onClick="location.href='PaymentOptions.asp'">
			</td>
		</tr>
	</table>
	</form>
<% end if %>
<!--#include file="AdminFooter.asp"-->
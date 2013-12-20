<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add New Custom Payment Option" %>
<% Section="paymntOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->  
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
err.clear
dim query, conntemp, rstemp
call openDb()
		
'Get Customer Categories
query="SELECT idCustomerCategory, pcCC_Name FROM pcCustomerCategories Order by pcCC_Name;"
set rstemp=Server.CreateObject("ADODB.Recordset")     
rstemp.Open query, conntemp
if err.number <> 0 then
	strErrDescription = err.description
	set rstemp=nothing
	call closeDb()
	response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCCPaymentOpt 345: "&strErrDescription) 
end If

Dim pcv_NoCPC, pcv_CPCString
pcv_NoCPC=1
pcv_CPCString = ""

do until rstemp.eof
	pcv_NoCPC=0
	pcv_IdCustomerCategory = rstemp("idCustomerCategory")
	pcv_CCName = rstemp("pcCC_Name")
	set rsCPCObj=nothing
	pcv_CPCString = pcv_CPCString&"<tr><td width=""6%""><input type=""checkbox"" name=""Cbtob2"" value='"&pcv_IdCustomerCategory&"' class='clearBorder'"&pcv_Checked&"></td><td width=""94%"">Apply to "&pcv_CCName&"</td></tr>"
	rstemp.movenext
loop

If request.form("Continue")<>"" then
	CustomCardDesc=request.form("customCardDesc")
	session("adminCustomCardDesc")=CustomCardDesc
	Cbtob=request.Form("Cbtob")
	session("adminCbtob")=Cbtob
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
	priceToAddType=request.form("priceToAddType")
	session("adminpriceToAddType")=priceToAddType
	if priceToAddType="price" then
		priceToAdd=replacecomma(Request("priceToAdd"))
		percentageToAdd="0"
	else
		priceToAdd="0"
		percentageToAdd=request.form("percentageToAdd")
	end if
	If priceToAdd="" Then
		priceToAdd="0"
	End If
	If percentageToAdd="" Then
		percentageToAdd="0"
	End If
	session("adminpercentageToAdd")=percentageToAdd
	session("adminpriceToAdd")=priceToAdd
	ruleName=request.form("ruleName")
	session("adminruleName")=ruleName
	ruleRequired=request.form("ruleRequired")
	session("adminruleRequired")=ruleRequired
	if ruleRequired="YES" then
		Creq="-1"
	else
		Creq="0"
	end if
	lengthOfField=request.form("lengthOfField")
	if lengthOfField="" then
		lengthOfField="20"
	end if
	session("adminlengthOfField")=lengthOfField
	maxInput=request.form("maxInput")
	if maxInput="" then
		maxInput=lengthOfField
	end if
	session("adminmaxInput")=maxInput
	if CustomCardDesc="" then
		response.redirect "AddCustomCardPaymentOpt.asp?msg=1"
	end if
	if ruleName="" then
		response.redirect "AddCustomCardPaymentOpt.asp?msg=2"
	end if
	if NOT isNumeric(lengthOfField) then
		response.redirect "AddCustomCardPaymentOpt.asp?msg=3"
	end if
	if NOT isNumeric(maxInput) then
		response.redirect "AddCustomCardPaymentOpt.asp?msg=4"
	end if
	
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

	'check to see if this CustomCardDesc exists in DB
	query="SELECT customCardDesc FROM customCardTypes WHERE customCardDesc='"&replace(CustomCardDesc,"'","''")&"';"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCustomCardPaymentOpt: "&Err.Description) 
	end If
	if NOT rstemp.eof then
		set rstemp=nothing
		call closedb()
		response.redirect "AddCustomCardPaymentOpt.asp?msg=5"
	else
		'insert new type
		query="INSERT INTO customCardTypes (customCardDesc) VALUES ('"&replace(customCardDesc,"'","''")&"');"
		set rstemp=conntemp.execute(query)
		if err.number <> 0 then
			set rstemp=nothing
			call closedb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCustomCardPaymentOpt: "&Err.Description) 
		end If
		
		query="SELECT idCustomCardType, customCardDesc FROM customCardTypes WHERE customCardDesc='"&replace(CustomCardDesc,"'","''")&"';"
		set rstemp=conntemp.execute(query)
		idCustomCardType=rstemp("idCustomCardType")
		
		'insert new paytype
		query="INSERT INTO payTypes (paymentDesc, sslURL, active, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode, terms, Cbtob, Creq, Cprompt, Type,pcPayTypes_processOrder,pcPayTypes_setPayStatus) VALUES ('"&replace(customCardDesc,"'","''")&"','paymnta_customcard.asp',-1,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",10"&idCustomCardType&",'',"& Cbtob &","& Creq &",'','CU'," & pcv_processOrder & "," & pcv_setPayStatus & ")"
		set rstemp=conntemp.execute(query)
		if err.number <> 0 then
			set rstemp=nothing
			call closedb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCustomCardPaymentOpt: "&Err.Description) 
		end If
		
		query="SELECT idPayment FROM payTypes WHERE gwCode=10"&idCustomCardType&";"
		set rstemp=conntemp.execute(query)
		pcv_idPayment=rstemp("idPayment")
		
		query="INSERT INTO customCardRules (idCustomCardType,ruleName,intRuleRequired,intLengthOfField,intMaxInput,intOrder) VALUES ("&idCustomCardType&",'"&replace(ruleName,"'","''")&"',"&Creq&","&lengthOfField&","&maxInput&",0);"
		set rstemp=conntemp.execute(query)
		if err.number <> 0 then
			set rstemp=nothing
			call closedb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCustomCardPaymentOpt: "&Err.Description) 
		end If
		
		query = "DELETE FROM CustCategoryPayTypes WHERE idPayment = "&pcv_idPayment&";"
		set rstemp=Server.CreateObject("ADODB.Recordset")     
		set rstemp=conntemp.execute(query)
	
		if Cbtob="2" then
			CbtobArray=split(Cbtob2,",")
			for t=lbound(CbtobArray) to ubound(CbtobArray)
				query = "INSERT INTO CustCategoryPayTypes (idCustomerCategory,idPayment) VALUES ("&CbtobArray(t)&" ,"&pcv_idPayment&");"
				set rstemp=Server.CreateObject("ADODB.Recordset")     
				set rstemp=conntemp.execute(query)
			next
		end if

	end if
	set rstemp=nothing
	call closedb()	
	response.redirect "modCustomCardPaymentOpt.asp?idcustomCardType="&idCustomCardType
End if

if request.QueryString("msg")<>"" then
	msg=request.QueryString("msg")
	if isNumeric(msg) then
	 select case msg
	 case "1"
	 	errmsg="'Option Name' is a required field."
	 case "2"
	 	errmsg="'Text Label' is a required field."
	 case "3"
	 	errmsg="'Length of Field' must be a numeric value."
	 case "4"
	 	errmsg="'Maximum Input length' must be a numeric value."
	 case "5"
	 	errmsg="There is already a custom card payment containing this option name."
	 end select
	end if
end if 
%>
<!--#include file="AdminHeader.asp"-->
<form name="form" method="post" action="AddCustomCardPaymentOpt.asp" class="pcForms">
	<table class="pcCPcontent">
		<% if errmsg<>"" then %>
			<tr>
				<td colspan="2" align="center">
					<div class="pcCPmessage"><%=errmsg%></div>
				</td>
			</tr>
			<tr> 
				<td colspan="2">&nbsp;</td>
			</tr>
		<% end if %>
		<tr>
			<th colspan="2">General Information</th>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2">Enter a description for this payment option and specify whether it applies to all customers or only wholesale customers.&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=303')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
		</tr>
		<tr> 
			<td width="25%" align="right">Option Name:</td>
			<td width="75%">
				<input name="customCardDesc" type="text" id="customCardDesc" value="<%=session("admincustomCardDesc")%>">
				<span class="pcSmallText">(e.g.: My Store Card) Maximum 70 characters</span>
			</td>
		</tr>
		<tr>
		  <td>&nbsp;</td>
		  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
		    <tr>
		      <td width="2%" align="right"><input type="radio" name="Cbtob" value="0" class="clearBorder" checked></td>
		      <td width="98%">Apply to all customers</td>
		      </tr>
		    <tr>
		      <td width="2%" align="right"><input type="radio" name="Cbtob" value="1" class="clearBorder"></td>
		      <td width="98%">Apply to Wholesale Customers only</td>
		      </tr>
				<% if pcv_CPCString&""<>"" then %>
                <tr>
                <td width="2%" align="right"><input type="radio" name="Cbtob" value="2" class="clearBorder"></td>
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
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<th colspan="2">Input Fields</th>
	</tr>
    <tr> 
        <td colspan="2" class="pcCPspacer"></td>
    </tr>
	<tr>
		<td colspan="2">You have the option to add an unlimited amount of input fields for this payment option. This is the information that customers will be asked to fill out when they checkout using this payment option. After you click &quot;Continue&quot;, you will be able to add <u>more input fields</u>.</td>
	</tr>
	<tr>
		<td align="right">Text Label:</td>
		<td>
			<input name="ruleName" type="text" id="ruleName" size="20" maxlength="50" value="<%=session("adminruleName")%>">
      <span class="pcSmallText">(e.g.: Name on Card) Maximum of 50 characters</span>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>
			<% if session("adminruleRequired")="YES" then %>
				<input name="ruleRequired" type="checkbox" id="ruleRequired" value="YES" checked>
			<% else %>
				<input name="ruleRequired" type="checkbox" id="ruleRequired" value="YES">
			<% end if %>
			This is a required field
		</td>
	</tr>
	<tr>
		<td align="right">Length of Field:</td>
		<td>
			<input name="lengthOfField" type="text" id="lengthOfField" value="<%=session("adminlengthOfField")%>" size="4" maxlength="4">
	    <span class="pcSmallText">The size of the input field itself</span>
		</td>
	</tr>
	<tr>
		<td align="right">Maximum Input length:</td>
		<td>
			<input name="maxInput" type="text" id="maxInput" size="4" maxlength="3" value="<%=session("adminmaxInput")%>">
			<span class="pcSmallText">Not to excede 250 characters</span>
		</td>
	</tr>
    
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="2">Additional Fees</th>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td colspan="2">You have the option to charge a processing fee for this payment option.&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=304')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
		</tr>
		<tr>
			<td align="right">
				<% if priceToAdd <> "0" then %>
					Processing fee:
				<% else %>
				<% end if %>
			</td>
			<td>
			<% if session("adminpriceToAddType")="price" then %>
				<input name="priceToAddType" type="radio" value="price" checked>
			<% else %>
				<input type="radio" name="priceToAddType" value="price">
			<% end if %>
						Flat rate&nbsp;&nbsp; <%=scCurSign%>
				<input name="priceToAdd" size="6" value="<%=money(session("adminpriceToAdd"))%>">
			</td>
		</tr>
		<tr>
		<td>&nbsp;</td>
		<td>
			<% if session("adminpriceToAddType")="percentage" then %>
				<input type="radio" name="priceToAddType" value="percentage" checked>
			<% else %>
				<input type="radio" name="priceToAddType" value="percentage">
			<% end if %>
						Percentage of total order&nbsp;&nbsp;%
			<input name="percentageToAdd" size="6" value="<%=session("adminpercentageToAdd")%>">
		</td>
	</tr>
	
	<%'Start SDBA%>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">Order Processing: Order Status and Payment Status</th>
	</tr>
    <tr> 
        <td colspan="2" class="pcCPspacer"></td>
    </tr>
	<tr>
		<td colspan="2">Process orders when they are placed: <input type="checkbox" name="pcv_processOrder" value="1">&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=301')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
	</tr>
	<tr>
		<td colspan="2">When orders are placed, set the payment status to:
			<select name="pcv_setPayStatus">
				<option value="0" selected>Pending</option>
				<option value="1">Authorized</option>
				<option value="2">Paid</option>
			</select>
			&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=302')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>
		</td>
	</tr>
	<%'End SDBA%>
	
	<tr> 
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2" align="center">
		<input name="Continue" type="submit" id="Continue" value="Continue" class="submit2">
		&nbsp;<input type="button" name="Button" value="Back" onClick="javascript:history.back()">
		</td>
	</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->
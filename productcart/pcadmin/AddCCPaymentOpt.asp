<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add Custom Payment Option" %>
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
dim query, conntemp, rstemp

call openDb()
sMode=Request.Form("SubmitA")

If sMode <> "" Then
	If sMode="Add" Then
		priceToAddCCType=request.form("priceToAddCCType")
		if priceToAddCCType="price" then
			priceToAddCC=replacecomma(Request("priceToAddCC"))
			percentageToAddCC="0"
		else
			priceToAddCC="0"
			percentageToAddCC=request.form("percentageToAddCC")
		end if
		if priceToAddCC="" then
			priceToAddCC="0"
		end if
		if percentageToAddCC="" then
			percentageToAddCC="0"
		end if
		cvv="0"
		paymentNickName=replace(request.Form("paymentNickName"),"'","''")
		if paymentNickName="" then
			paymentNickName="Credit Card"
		end if
		gwCC=request.form("gwCC")
		
		Cbtob=request.Form("CbtobCC")
		if Cbtob = "" then
			Cbtob = "0"
		end if
		if Cbtob="1" then
			Cbtob="-1"
		end if
		if Cbtob="2" then
			Cbtob2 = request("Cbtob2CC")
			if Cbtob2 = "" then
				Cbtob = "0"
			end if
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
		
		If gwCC="1" then
			'request which Credit Cards the store wishes to accept
			M=request.form("M")
			V=request.form("V")
			D=request.form("D")
			A=request.form("A")
			DC=request.form("DC")
			dit="0"
			If M="1" Then
				dit="1"
				query="UPDATE CCTypes SET active=-1 WHERE CCcode='M'"
				set rstemp=Server.CreateObject("ADODB.Recordset")     
				rstemp.Open query, conntemp
				if err.number <> 0 then
					set rstemp=nothing
					call closeDb()
				  	response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCCPaymentOpt: "&Err.Description) 
				end If
			Else
				query="UPDATE CCTypes SET active=0 WHERE CCcode='M'"
				set rstemp=Server.CreateObject("ADODB.Recordset")     
				rstemp.Open query, conntemp
				if err.number <> 0 then
					set rstemp=nothing
					call closeDb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCCPaymentOpt: "&Err.Description) 
				end If
			End If
			If V="1" Then
				dit="1"
				query="UPDATE CCTypes SET active=-1 WHERE CCcode='V'"
				set rstemp=Server.CreateObject("ADODB.Recordset")     
				rstemp.Open query, conntemp
				if err.number <> 0 then
					set rstemp=nothing
					call closeDb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCCPaymentOpt: "&Err.Description) 
				end If
			Else
				query="UPDATE CCTypes SET active=0 WHERE CCcode='V'"
				set rstemp=Server.CreateObject("ADODB.Recordset")     
				rstemp.Open query, conntemp
				if err.number <> 0 then
					set rstemp=nothing
					call closeDb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCCPaymentOpt: "&Err.Description) 
				end If
			End If
			If D="1" Then
				dit="1"
				query="UPDATE CCTypes SET active=-1 WHERE CCcode='D'"
				set rstemp=Server.CreateObject("ADODB.Recordset")
				rstemp.Open query, conntemp
				if err.number <> 0 then
					set rstemp=nothing
					call closeDb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCCPaymentOpt: "&Err.Description) 
				end If
			Else
				query="UPDATE CCTypes SET active=0 WHERE CCcode='D'"
				set rstemp=Server.CreateObject("ADODB.Recordset")     
				rstemp.Open query, conntemp
				if err.number <> 0 then
					set rstemp=nothing
					call closeDb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCCPaymentOpt: "&Err.Description) 
				end If
			End If
			If A="1" Then
				dit="1"
				query="UPDATE CCTypes SET active=-1 WHERE CCcode='A'"
				set rstemp=Server.CreateObject("ADODB.Recordset")     
				rstemp.Open query, conntemp
				if err.number <> 0 then
					set rstemp=nothing
					call closeDb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCCPaymentOpt: "&Err.Description) 
				end If
			Else
				query="UPDATE CCTypes SET active=0 WHERE CCcode='A'"
				set rstemp=Server.CreateObject("ADODB.Recordset")     
				rstemp.Open query, conntemp
				if err.number <> 0 then
					set rstemp=nothing
					call closeDb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCCPaymentOpt: "&Err.Description) 
				end If
			End If
			If DC="1" Then
				dit="1"
				query="UPDATE CCTypes SET active=-1 WHERE CCcode='DC'"
				set rstemp=Server.CreateObject("ADODB.Recordset")     
				rstemp.Open query, conntemp
				if err.number <> 0 then
					set rstemp=nothing
					call closeDb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCCPaymentOpt: "&Err.Description) 
				end If
			Else
				query="UPDATE CCTypes SET active=0 WHERE CCcode='DC'"
				set rstemp=Server.CreateObject("ADODB.Recordset")     
				rstemp.Open query, conntemp
				if err.number <> 0 then
					set rstemp=nothing
					call closeDb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCCPaymentOpt: "&Err.Description) 
				end If
			End If
			If dit="0" then
				query="UPDATE payTypes SET active=0,pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & " WHERE gwcode=6"
				set rstemp=Server.CreateObject("ADODB.Recordset")     
				rstemp.Open query, conntemp
				if err.number <> 0 then
					set rstemp=nothing
					call closeDb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCCPaymentOpt: "&Err.Description) 
				end If
			Else
				query="UPDATE payTypes SET active=-1, priceToAdd="&priceToAddCC&", percentageToAdd="&percentageToAddCC&", cvv="&cvv&", paymentNickName='"&paymentNickName&"',pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & " WHERE gwcode=6"
				set rstemp=Server.CreateObject("ADODB.Recordset")     
				rstemp.Open query, conntemp
				if err.number <> 0 then
					set rstemp=nothing
					call closeDb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCCPaymentOpt: "&Err.Description) 
				end If
			End If
			
			query = "SELECT idPayment FROM payTypes WHERE gwcode=6" 
			set rstemp=Server.CreateObject("ADODB.Recordset")     
			set rstemp=conntemp.execute(query)
			pcv_idPayment = rstemp("idPayment")

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
		End If
		If gwCC="" Then
			'set credit cards as inactive
			query="UPDATE CCTypes SET active=0 WHERE active=-1"
			set rstemp=Server.CreateObject("ADODB.Recordset")     
			rstemp.Open query, conntemp
			if err.number <> 0 then
				set rstemp=nothing
				call closeDb()				
				response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCCPaymentOpt: "&Err.Description) 
			end If
			query="UPDATE payTypes SET active=0,pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & " WHERE gwCode=6"
			set rstemp=Server.CreateObject("ADODB.Recordset")     
			rstemp.Open query, conntemp
			if err.number <> 0 then
				set rstemp=nothing
				call closeDb()
				response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCCPaymentOpt: "&Err.Description) 
			end If
		End If
	End If
	
	If sMode="Delete" Then
		'code
	End If
	set rstemp=nothing
	call closeDb()
	response.redirect "PaymentOptions.asp"
End If

sMode2=Request.Form("SubmitB")

If sMode2 <> "" Then
	If sMode2="Add" Then
		priceToAddType=request.form("priceToAddType")
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
		paymentDesc=replace(request.form("CDesc"),"'","''")
		terms=replace(request.form("CTerms"),"'","''")
		terms=replace(terms,vbcrlf,"<br>")
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

		CReq=request.form("Creq")
		If  CReq="1" then
			CReq="-1"
			Cprompt=replace(request.form("Cprompt"),"'","''")
		Else
			Creq="0"
		End If
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
		
		'insert new paytype into db
		query="INSERT INTO payTypes (paymentDesc, sslURL, active, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode, terms, Cbtob, Creq, Cprompt, Type,pcPayTypes_processOrder,pcPayTypes_setPayStatus) VALUES ('"& paymentDesc &"','paymnta_c.asp',-1,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",7,'"& terms &"',"& Cbtob &","& Creq &",'"& Cprompt &"','C'," & pcv_processOrder & "," & pcv_setPayStatus & ")"
		set rstemp=Server.CreateObject("ADODB.Recordset")     
		rstemp.Open query, conntemp
		if err.number <> 0 then
			set rstemp=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCCPaymentOpt: "&Err.Description) 
		end If
		
		query = "SELECT idPayment FROM payTypes WHERE paymentDesc='"& paymentDesc &"';" 
		set rstemp=Server.CreateObject("ADODB.Recordset")     
		set rstemp=conntemp.execute(query)
		pcv_idPayment = rstemp("idPayment")

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
	End If
	set rstemp=nothing
	call closeDb()
	response.redirect "PaymentOptions.asp"
End If
		
		
query="SELECT idPayment, gwCode, Cbtob FROM paytypes WHERE active=-1;"
set rstemp=Server.CreateObject("ADODB.Recordset")     
rstemp.Open query, conntemp

if err.number <> 0 then
	strErrDescription = err.description
	set rstemp=nothing
	call closeDb()
	response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCCPaymentOpt 279: "&strErrDescription) 
end If

if NOT rstemp.eof then
	iCnt=1
	oCnt=0
	do until rstemp.eof
		pcv_GwCode = rstemp("gwCode")
		Cbtob = rstemp("Cbtob")
		select case pcv_GwCode
			case "6"
				gwCC="1"
				gwCCPaymentType = rstemp("idPayment")
				CbtobCC = Cbtob
			case "7"
				gwOL="1"
				oCnt=oCnt + 1
		end select
		rstemp.moveNext
	loop
	set rstemp = nothing

	if gwCC="1" Then
		' get CCtypes
		query="SELECT * FROM CCTypes"
		set rstemp=Server.CreateObject("ADODB.Recordset")     
		rstemp.Open query, conntemp
		if err.number <> 0 then
			strErrDescription = err.description
			set rstemp=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCCPaymentOpt 307: "&strErrDescription) 
		end If
		do until rstemp.eof
			varTemp=rstemp("CCcode")
			active=rstemp("active")
			if varTemp="0" then
			else
				if varTemp="M" AND active="-1" then
					M="1"
				end if
				if varTemp="V" AND active="-1" then
					V="1"
				end if
				if varTemp="A" AND active="-1" then
					A="1"
				end if
				if varTemp="D" AND active="-1" then
					D="1"
				end if
				if varTemp="DC" AND active="-1" then
					DC="1"
				end if
			end if
			rstemp.movenext
		loop
	end if
end if

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

Dim pcv_NoCPC, pcv_CPCString, pcv_CPCStringCC
pcv_NoCPC=1
pcv_CPCString = ""
pcv_CPCStringCC = ""

do until rstemp.eof
	pcv_NoCPC=0
	pcv_IdCustomerCategory = rstemp("idCustomerCategory")
	pcv_CCName = rstemp("pcCC_Name")
	'check if this is selected for the CC already.
	if gwCCPaymentType<>"" then
		query = "SELECT * FROM CustCategoryPayTypes WHERE idPayment="&gwCCPaymentType&" AND idCustomerCategory="&pcv_IdCustomerCategory&";"
		set rsCPCObj=Server.CreateObject("ADODB.Recordset")     
		set rsCPCObj=connTemp.execute(query)
		if rsCPCObj.eof then
			pcv_Checked=""
		else
			pcv_Checked=" checked"
		end if
		set rsCPCObj=nothing
	end if
	pcv_CPCStringCC = pcv_CPCStringCC&"<tr><td width=""6%""><input type=""checkbox"" name=""Cbtob2CC"" value='"&pcv_IdCustomerCategory&"' class='clearBorder'"&pcv_Checked&"></td><td width=""94%"">Apply to "&pcv_CCName&"</td></tr>"
	pcv_CPCString = pcv_CPCString&"<tr><td width=""6%""><input type=""checkbox"" name=""Cbtob2"" value='"&pcv_IdCustomerCategory&"' class='clearBorder'"&pcv_Checked&"></td><td width=""94%"">Apply to "&pcv_CCName&"</td></tr>"
	rstemp.movenext
loop
 
call closeDb()
%>
<!--#include file="AdminHeader.asp"-->
<form name="form1" method="post" action="AddCCPaymentOpt.asp" class="pcForms">
	<table class="pcCPcontent">
		<tr>
        	<td class="pcCPspacer"></td>
        </tr>
		<tr> 
			<th>Offline credit card processing:</td>
		</tr>
		<tr>
        	<td class="pcCPspacer"></td>
        </tr>
		<tr> 
			<td valign="top"> 
				<table class="pcCPcontent">
					<tr> 
						<td colspan="3"> 
							<% if gwCC="1" then %>
								Offline credit card payments are currently <strong>enabled</strong>. The following credit card types are shown as accepted on your store:
                                <input name="gwCC" type="hidden" value="1" checked>
							<% else %>
                            	Offline credit card payments are currently disabled. 
								<input name="gwCC" type="hidden" value="1" checked>
							<% end if %>
						</td>
					</tr>
					<tr> 
						<td width="4%" height="21">&nbsp;</td>
						<td width="5%" height="21"> 
							<% if M="1" then %>
								<input type="checkbox" name="M" value="1" checked class="clearBorder">
							<% else %>
								<input type="checkbox" name="M" value="1" class="clearBorder">
							<% end if %>
						</td>
						<td width="91%" height="21">MasterCard</td>
					</tr>
					<tr> 
						<td width="4%" height="21">&nbsp;</td>
						<td width="5%" height="21">  
						<% if V="1" then %>
							<input type="checkbox" name="V" value="1" checked class="clearBorder">
						<% else %>
							<input type="checkbox" name="V" value="1" class="clearBorder">
						<% end if %>
						</td>
						<td width="91%" height="21">Visa</td>
					</tr>
					<tr> 
						<td width="4%" height="21">&nbsp;</td>
						<td width="5%" height="21"> 
						<% if D="1" then %>
							<input type="checkbox" name="D" value="1" checked class="clearBorder">
						<% else %>
							<input type="checkbox" name="D" value="1" class="clearBorder">
						<% end if %>
						</td>
						<td width="91%" height="21">Discover</td>
					</tr>
					<tr> 
						<td width="4%" height="21">&nbsp;</td>
						<td width="5%" height="21"> 
						<% if A="1" then %>
							<input type="checkbox" name="A" value="1" checked class="clearBorder">
						<% else %>
							<input type="checkbox" name="A" value="1" class="clearBorder">
						<% end if %>
						</td>
						<td width="91%" height="21">American Express</td>
					</tr>
					<tr> 
						<td width="4%" height="21">&nbsp;</td>
						<td width="5%" height="21"> 
							<% if DC="1" then %>
								<input type="checkbox" name="DC" value="1" checked class="clearBorder">
							<% else %>
								<input type="checkbox" name="DC" value="1" class="clearBorder">
							<% end if %>
						</td>
						<td width="91%" height="21">Diner's Club</td>
					</tr>
				</table>
				<table class="pcCPcontent">
                    <tr>
                        <td class="pcCPspacer" colspan="2"></td>
                    </tr>
					<tr> 
						<td colspan="2">You have the option to charge a processing fee for this payment option.</td>
						<td>&nbsp;</td>
					</tr>
					<tr>
						<td>Processing fee:</td>
						<td>
						<input type="radio" name="priceToAddCCType" value="price"class="clearBorder">
						Flat rate&nbsp;&nbsp;
						<%=scCurSign%>
						<input name="priceToAddCC" size="6" value="<%=money(0)%>">
						</td>
						<td>&nbsp;</td>
					</tr>
					<tr> 
						<td width="16%">&nbsp;</td>
						<td width="84%">
						<input type="radio" name="priceToAddCCType" value="percentage" class="clearBorder">
						Percentage of total order&nbsp;&nbsp;%  
						<input name="percentageToAddCC" size="6" value="0">
						</td>
						<td>&nbsp;</td>
					</tr>
					<%'Start SDBA%>
					<tr> 
						<td colspan="2">
						Process orders when they are placed: <input type="checkbox" name="pcv_processOrder" value="1" class="clearBorder"></td>
						<td>&nbsp;</td>
					</tr>
					<tr> 
						<td colspan="2">When orders are placed, set payment status is:
							<select name="pcv_setPayStatus">
								<option value="0" selected>Pending</option>
								<option value="1">Authorized</option>
								<option value="2">Paid</option>
							</select></td>
						<td>&nbsp;</td>
					</tr>
					<%'End SDBA%>
					<tr> 
						<td colspan="2">&nbsp; </td>
						<td>&nbsp;</td>
					</tr>
					<tr>
						<td colspan="2">You can change the display name that is shown for this payment type. </td>
						<td>&nbsp;</td>
					</tr>
					<tr>
						<td>Payment Name:&nbsp;</td>
						<td><input type="hidden" name="cvv" value="0"><input name="paymentNickName" value="Credit Card" size="35" maxlength="255"></td>
						<td>&nbsp;</td>
					</tr>
						<tr> 
							<td align="right"><input type="radio" name="CbtobCC" value="0" class="clearBorder"></td>
							<td>Apply to all customers</td>
                        </tr>
						<tr>
						  <td align="right"><input type="radio" name="CbtobCC" value="1" class="clearBorder" <% if CbtobCC=-1 then%>checked<% end if%>></td>
						  <td>Apply to wholesale customers only</td>
					  </tr>
                      <% if pcv_CPCStringCC&""<>"" then %>
                        <tr>
                            <td align="right"><input type="radio" name="CbtobCC" value="2" class="clearBorder" <% if CbtobCC=2 then%>checked<% end if%>></td>
                            <td>Apply to the following customer pricing categories</td>
                        </tr>
                        <tr> 
                            <td align="right">&nbsp;</td>
                            <td><table width="95%" border="0" cellspacing="0" cellpadding="2">
                            <%=pcv_CPCStringCC%>
                            </table></td>
                        </tr>
                     <% end if %>
                    <tr>
                    	<td colspan="3"><hr></td>
                    </tr>
                    <tr>
                    	<td colspan="3">
                        <input type="submit" value="Add" name="SubmitA" class="submit2">&nbsp;
                        <input type="button" name="Button" value="Back" onClick="javascript:history.back()">
                        </td>
                    </tr>
				</table>
			</td>
		</tr>
	</table>
</form>
<script language="JavaScript">
<!--

function Form1_Validator(theForm)
{
    if (theForm.CDesc.value == "")
        {
            alert("Description is a required field.");
                theForm.CDesc.focus();
                return (false);
    }

return (true);
}
//-->
</script>
<form name="Offline" method="post" action="AddCCPaymentOpt.asp" onSubmit="return Form1_Validator(this)" class="pcForms">
    <table class="pcCPcontent">
        <tr>
            <td class="pcCPspacer"></td>
        </tr>
        <tr> 
            <th>Other custom payment options:</th>
        </tr>
        <tr>
            <td class="pcCPspacer"></td>
        </tr>
        <tr> 
            <td valign="top"> 
                <table class="pcCPcontent">
                    <tr> 
                        <td colspan="2"> 
                            <% if gwOL="1" then %>
                                You currently have <b><font color=#ff0000><%=oCnt%></font></b> active custom payment option(s).
                            <% else %>
                                You have not set up any other custom payment option. You can start using the form below.
                            <% end if %>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" class="pcCPspacer"></td>
                    </tr>
                    <tr> 
                        <td width="24%"> 
                            <div align="right">Description:</div></td>
                        <td width="76%">  
                            <input type="text" name="CDesc">
                            <img src="images/pc_required.gif" width="9" height="9">&nbsp;<span class="pcSmallText">(e.g.: C.O.D., Money Order, Net 30)</span></td>
                    </tr>
                    <tr> 
                        <td width="24%" valign="top"> 
                            <div align="right">Terms:<br>
                            <span class="pcSmallText">(Displayed to customers during check out)</div></td>
                        <td width="76%">  
                            <textarea name="CTerms" cols="30" rows="5"></textarea>
                        </td>
                    </tr>
                    <tr> 
                        <td align="right"><input type="radio" name="Cbtob" value="0" class="clearBorder"></td>
                        <td>Apply to all customers</td>
                    </tr>
                    <tr>
                      <td align="right"><input type="radio" name="Cbtob" value="1" class="clearBorder"></td>
                      <td>Apply to wholesale customers only</td>
                  </tr>
					<% if pcv_CPCString&""<>"" then %>
                        <tr>
                            <td align="right"><input type="radio" name="Cbtob" value="2" class="clearBorder"></td>
                            <td>Apply to the following customer pricing categories</td>
                        </tr>
                        <tr> 
                            <td align="right">&nbsp;</td>
                            <td><table width="95%" border="0" cellspacing="0" cellpadding="2">
                            <%=pcv_CPCString%>
                            </table></td>
                        </tr>
                     <% end if %>
                    <tr> 
                        <td colspan="2"><b>Optional</b>:<br>
                        Checking the box below will prompt the customer to input more information, such as an account number or purchase order number, before completing their order.</td>
                    </tr>
                    <tr> 
                        <td align="right"><input type="checkbox" name="Creq" value="1" class="clearBorder"></td>
                        <td>Require additional information for this payment option</td>
                    </tr>
                    <tr> 
                        <td align="right">Description:</td>
                        <td><input type="text" name="Cprompt">&nbsp;<span class="pcSmallText">(e.g.: Purchase Order #, Account Number)</span> </td>
                    </tr>
                    <tr> 
                        <td colspan="2">&nbsp;</td>
                    </tr>
                    <tr> 
                        <td colspan="2">You have the option to charge a processing fee for this payment option.</td>
                    </tr>
                    <tr>
                        <% if priceToAdd <> "0" then %>
                            <td>
                                <div align="right">Processing fee:</div></td>
                        <% end if %>
                        <td>
                            <input type="radio" name="priceToAddType" value="price" class="clearBorder">
                            Flat rate&nbsp;&nbsp; <%=scCurSign%>
                            <input name="priceToAdd" size="6" value="<%=money(0)%>">
                        </td>
                    </tr>
                    <tr>
                        <td>&nbsp;</td>
                        <td>
                        <input type="radio" name="priceToAddType" value="percentage" class="clearBorder">
                        Percentage of total order&nbsp;&nbsp;%
                        <input name="percentageToAdd" size="6" value="0">
                        </td>
                    </tr>
                    <%'Start SDBA%>
                    <tr>
                        <td colspan="2">Process orders when they are placed: <input type="checkbox" name="pcv_processOrder" value="1" class="clearBorder"></td>
                    </tr>
                    <tr> 
                        <td colspan="2">When orders are placed, set payment status is:
                            <select name="pcv_setPayStatus">
                                <option value="0" selected>Pending</option>
                                <option value="1">Authorized</option>
                                <option value="2">Paid</option>
                            </select>
                        </td>
                    </tr>
                    <%'End SDBA%>
                    <tr> 
                        <td colspan="2"><hr></td>
                    </tr>
                    <tr>
                        <td>
                            <input type="submit" value="Add" name="SubmitB" class="submit2">&nbsp;
                            <input type="button" name="Button" value="Back" onClick="javascript:history.back()">
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</form>
<!--#include file="AdminFooter.asp"-->
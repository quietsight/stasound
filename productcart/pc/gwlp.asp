<% 'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="header.asp"-->
<% 
'//////////////////////////////////////////////////////////////////////////
'// LinkPoint will generate it's own Order id for each transaction.
'// This will prevent errors with duplicate order ids that customer 
'// receive if their initial order attempt was declined. If you wish 
'// to use the ProductCart Order ID instead, you should set the following
'// variable below to equal "1"
'//
'// pcv_UsePcOrderID=1 
'//
'//////////////////////////////////////////////////////////////////////////
dim pcv_UsePcOrderID
pcv_UsePcOrderID=0
'//////////////////////////////////////////////////////////////////////////


'// Check if LP is posting back
If request("approval_code")<>"" OR request("failReason")<>""  then
	pcIntOrderFailed=0
	if request("failReason")<>"" then
		pcIntOrderFailed=1
		pcStrFailedReason=request("failReason")
	end if
	
	tempTransId=request("oid")
	if instr(tempTransId,",") then
		arryTransId=split(tempTransId,",")
		pTransID=arryTransId(0)
	else
		pTransID=tempTransId
	end if

	tempIdOrder=request("userid")
	if instr(tempIdOrder,",") then
		arryIdOrder=split(tempIdOrder,",")
		pIdOrder=arryIdOrder(0)
	else
		pIdOrder=tempIdOrder
	end if
	if session("GWOrderId")="" then
		session("GWOrderId")=pIdOrder
	end if
	
	if pcIntOrderFailed=0 then
		session("GWAuthCode")=request("approval_code")
		session("GWTransId")=pTransID
		session("GWSessionID")=Session.SessionID 
		
		Response.redirect "gwReturn.asp?s=true&gw=LinkPoint"
	end if
end if

if session("pcStrFailedReason")=1 then
	pcIntOrderFailed=1
	pcStrFailedReason=session("pcIntOrderFailed")
	session("pcIntOrderFailed")=""
	session("pcStrFailedReason")=""
end if

'//Set redirect page to the current file name
session("redirectPage")="gwlp.asp"

'//Declare and Retrieve Customer's IP Address	
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
	
'//Declare URL path to gwSubmit.asp	
Dim tempURL
If scSSL="" OR scSSL="0" Then
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://") 
Else
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
End If
		
'//Get Order ID
if session("GWOrderId")="" then
	session("GWOrderId")=request("idOrder")
end if

'//Retrieve customer data from the database using the current session id		
pcGatewayDataIdOrder=session("GWOrderID")
%>
<!--#include file="pcGateWayData.asp"-->
<% '//Set customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer
if pcIntOrderFailed=1 then
		response.redirect "msgb.asp?message="&server.URLEncode("<b>Error</b>: "&pcStrFailedReason&"<br><br><a href="""&tempURL&"?psslurl="&session("redirectPage")&"&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"&ordertotal="&pcBillingTotal&"""><img src="""&rslayout("back")&""" border=0></a>")
end if

dim connTemp, rs
call openDb()

'//Retrieve any gateway specific data from database or hard-code the variables
query="SELECT storeName, transType, lp_testmode, lp_cards, CVM, lp_yourpay FROM LinkPoint where id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
lp_storeName=rs("storeName") 
lp_transType=rs("transType")
lp_testmode=rs("lp_testmode")
lp_cards=rs("lp_cards")
lp_CVM=rs("CVM")
lp_yourpay=rs("lp_yourpay")
if lp_CVM<>1 then
	lp_CVM=0
end if

set rs=nothing
call closedb()

	
	' Define FORM action
	pcStrOnSubmit=""
	if lp_CVM="1" then
		pcStrOnSubmit="return checkproqty(document.PaymentForm.cvm);"
	end if

%>
<script language="JavaScript">
<!--
function checkproqty(cvm)
{
if (cvm.value == "")
{
	alert("<%=dictLanguage.Item(Session("language")&"_GateWay_11")%> is required.");
	cvm.focus();
	return (false);
	}
}
//-->
</script>

<% 
if lp_testmode="YES" then 
	pcStrPostingURL="https://www.staging.linkpointcentral.com/lpc/servlet/lppay"
else
	if lp_yourpay="YES" then
		pcStrPostingURL="https://secure.linkpt.net/lpcentral/servlet/lppay"
	else 
		pcStrPostingURL="https://www.linkpointcentral.com/lpc/servlet/lppay"
	end if
end if

if scSSL="1" then
	pcStrReferURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwlp.asp"),"//","/")
	pcStrReferURL=replace(pcStrReferURL,"http:/","http://")
	pcStrReferURL=replace(pcStrReferURL,"https:/","https://")
else
	pcStrReferURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwlp.asp"),"//","/")
	pcStrReferURL=replace(pcStrReferURL,"http:/","http://")
	pcStrReferURL=replace(pcStrReferURL,"https:/","https://")
end if

%>
<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td>
				<img src="images/checkout_bar_step5.gif" alt="">
			</td>
		</tr>
		<tr>
			<td>
				<form method="POST" action="<%=pcStrPostingURL%>" name="PaymentForm" onsubmit="<%=pcStrOnSubmit%>" class="pcForms">
					<input type='hidden' name='referurl' value='<%=pcStrReferURL%>'>
					<input type='hidden' name='responseURL' value='<%=pcStrReferURL%>'>
					<input type='hidden' name='responseSuccessURL' value='<%=pcStrReferURL%>'>
					<input type='hidden' name='responseFailURL' value='<%=pcStrReferURL%>'>
					<input type='hidden' name='mode' value='payonly'>
					<input type='hidden' name='storename' value='<%=lp_storeName%>'>
					<input type='hidden' name='transid' value='<%=session("GWOrderId")%>'>
					<input type="hidden" name="txntype" value="<%=lp_transType%>">
					<% if lp_yourpay="YES" then
					else %>
						<input type='hidden' name='2000' value='Submit'>
					<% end if %>
                    <% if pcBillingCountryCode="US" then
						pcv_bstate=pcBillingState
						pcv_bstate2=""
					else
						pcv_bstate=""
						pcv_bstate2=pcBillingState
					end if %>
                    <% if pcv_UsePcOrderID=1 then
                  		response.write "<input type='hidden' name='oid' value='"&session("GWOrderId")&"'>"
					end if %>
					<input type='hidden' name='userid' value='<%=session("GWOrderId")%>'>
					<input type="hidden" name="bname" value="<%=pcBillingFirstName&" "&pcBillingLastName%>">
					<input type="hidden" name="bcompany" value="<%=pcBillingCompany%>">
					<input type="hidden" name="baddr1" value="<%=pcBillingAddress%>">
					<input type="hidden" name="baddr2" value="<%=pcBillingAddress2%>">
					<input type="hidden" name="bcity" value="<%= pcBillingCity%>">
					<input type="hidden" name="bstate" value="<%=pcv_bstate%>">
					<input type="hidden" name="bstate2" value="<%=pcv_bstate2%>">
					<input type="hidden" name="bzip" value="<%=pcBillingPostalCode%>">
					<input type="hidden" name="bcountry" value="<%=pcBillingCountryCode%>">
					<input type="hidden" name="phone" value="<%=pcBillingPhone%>">
					<input type='hidden' name='chargetotal' value='<%=replace(money(pcBillingTotal),",","")%>'>
                    
                    <% if pcShippingCountryCode="US" then
						pcv_sstate=pcShippingState
						pcv_sstate2=""
					else
						pcv_sstate=""
						pcv_sstate2=pcShippingState
					end if %>
					<input type='hidden' name='sname' value="<%=pcShippingFirstName&" "&pcShippingLastName%>">
					<input type='hidden' name='saddr1' value="<%=pcShippingAddress%>">
					<input type='hidden' name='saddr2' value="<%=pcShippingAddress2%>">
					<input type='hidden' name='scity' value="<%=pcShippingCity%>">
					<input type='hidden' name='sstate' value="<%=pcv_sstate%>">
					<input type='hidden' name='sstate2' value="<%=pcv_sstate2%>">
					<input type='hidden' name='szip' value="<%=pcShippingPostalCode%>">
					<input type='hidden' name='scountry' value="<%=pcShippingCountryCode%>">
					<input type='hidden' name='email' value="<%=pcCustomerEmail%>">
					
					<input type="hidden" name="txnorg" value="eci" />
					<input type="hidden" name="authenticateTransaction" value="False" />

					<table class="pcShowContent">
			
					<% if Msg<>"" then %>
						<tr valign="top"> 
							<td colspan="2">
								<div class="pcErrorMessage"><%=Msg%></div>
							</td>
						</tr>
					<% end if %>
					<% if len(pcCustIpAddress)>0 AND CustomerIPAlert="1" then %>
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
						<tr>
							<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_GateWay_6")&pcCustIpAddress%></p></td>
						</tr>
					<% end if %>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr class="pcSectionTitle">
						<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_GateWay_1")%></p></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr>
						<td colspan="2"><p><%=pcBillingFirstName&" "&pcBillingLastName%></p></td>
					</tr>
					<tr>
						<td colspan="2"><p><%=pcBillingAddress%></p></td>
					</tr>
					<% if pcBillingAddress2<>"" then %>
					<tr>
						<td colspan="2"><p><%=pcBillingAddress2%></p></td>
					</tr>
					<% end if %>
					<tr>
						<td colspan="2"><p><%=pcBillingCity&", "&pcBillingState%><% if pcBillingPostalCode<>"" then%>&nbsp;<%=pcBillingPostalCode%><% end if %></p></td>
					</tr>
					<tr>
						<td colspan="2"><p><a href="pcModifyBillingInfo.asp"><%=dictLanguage.Item(Session("language")&"_GateWay_2")%></a></p></td>
					</tr>
					<% if lp_testmode="YES" then %>
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
						<tr>
							<td colspan="2"><div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_3")%></div></td>
						</tr>
					<% end if %>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					
					<tr class="pcSectionTitle">
						<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_GateWay_5")%></p></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr>
						<td><%=dictLanguage.Item(Session("language")&"_GateWay_12")%></td> 
						<td>
							<select name="cctype">
							<% cardTypeArray=split(lp_cards,", ")
							i=ubound(cardTypeArray)
							cardCnt=0
							do until cardCnt=i+1
								'response.write cardTypeArray(cardCnt)
								if cardTypeArray(cardCnt)="V" then %>
									<option value="V" selected>Visa</option>
								<% end if 
								if cardTypeArray(cardCnt)="M" then %>
									<option value="M">MasterCard</option>
								<% end if 
								if cardTypeArray(cardCnt)="A" then %>
									<option value="A">American Express</option>
								<% end if 
								if cardTypeArray(cardCnt)="D" then %>
									<option value="D">Discover</option>
								<% end if 
								cardCnt=cardCnt+1
							loop
							%>
						</select>
					</td>
				</tr>
				<tr> 
					<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></p></td>
					<td> 
						<input type="text" name="cardnumber" value="">
					</td>
				</tr>
				<tr> 
					<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></p></td>
					<td><%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
						<select name="expmonth">
							<option value="01">1</option>
							<option value="02">2</option>
							<option value="03">3</option>
							<option value="04">4</option>
							<option value="05">5</option>
							<option value="06">6</option>
							<option value="07">7</option>
							<option value="08">8</option>
							<option value="09">9</option>
							<option value="10">10</option>
							<option value="11">11</option>
							<option value="12">12</option>
						</select>
						<% dtCurYear=Year(date()) %>
						&nbsp;<%=dictLanguage.Item(Session("language")&"_GateWay_10")%> 
						<select name="expyear">
							<option value="<%=right(dtCurYear,4)%>" selected><%=dtCurYear%></option>
							<option value="<%=right(dtCurYear+1,4)%>"><%=dtCurYear+1%></option>
							<option value="<%=right(dtCurYear+2,4)%>"><%=dtCurYear+2%></option>
							<option value="<%=right(dtCurYear+3,4)%>"><%=dtCurYear+3%></option>
							<option value="<%=right(dtCurYear+4,4)%>"><%=dtCurYear+4%></option>
							<option value="<%=right(dtCurYear+5,4)%>"><%=dtCurYear+5%></option>
							<option value="<%=right(dtCurYear+6,4)%>"><%=dtCurYear+6%></option>
							<option value="<%=right(dtCurYear+7,4)%>"><%=dtCurYear+7%></option>
							<option value="<%=right(dtCurYear+8,4)%>"><%=dtCurYear+8%></option>
							<option value="<%=right(dtCurYear+9,4)%>"><%=dtCurYear+9%></option>
							<option value="<%=right(dtCurYear+10,4)%>"><%=dtCurYear+10%></option>
						</select>
						</td>
					</tr>
					<% if lp_CVM="1" then %>
						<tr> 
							<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></p></td>
							<td> 
								<input type="hidden" name="cvmnotpres" value="0">
								<input name="cvm" type="text" size="4" maxlength="4">
							</td>
						</tr>
						<tr> 
							<td>&nbsp;</td>
							<td><img src="images/CVC.gif" alt="cvc code" width="212" height="155"></td>
						</tr>
					<% end If %>
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></p></td>
						<td><%=money(pcBillingTotal)%></td>
					</tr>
					
					<tr> 
						<td colspan="2" align="center">
							<!--#include file="inc_gatewayButtons.asp"-->
						</td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
</table>
</div>
<!--#include file="footer.asp"-->
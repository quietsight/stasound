<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
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
'//Set redirect page to the current file name
session("redirectPage")="gwPSI.asp"

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

dim connTemp, rs
call openDb()

'//Retrieve any gateway specific data from database or hard-code the variables
query="SELECT Config_File_Name, Config_File_Name_Full, Host, Port, psi_testmode FROM PSIGate WHERE id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
Config_File_Name=rs("Config_File_Name")
Config_File_Name_Full=rs("Config_File_Name_Full")
Host=rs("Host")
Port=rs("Port")
psi_testmode=rs("psi_testmode")

set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	Set PSIObj=CreateObject("MyServer.PsiGate")
	
	'This is supplied by PSiGate, must match configurate filename on 
	'PSiGate payment gateway

	PSIObj.Configfile=trim(Config_File_Name)
    
	'This is the location of the certificate file you download from
	'PSiGate

	PSIObj.Keyfile=Server.MapPath(trim(Config_File_Name_Full))
	
	PSIObj.Host=trim(Host)
	PSIObj.Port=trim(Port)
		
	PSIObj.Oid=session("GWOrderId")
	PSIObj.Userid="COM Sample FORM"
	PSIObj.Bname=pcBillingFirstName&" "&pcBillingLastName
	PSIObj.Bcompany=pcBillingCompany
	PSIObj.Baddr1=pcBillingAddress
	PSIObj.Baddr2=pcBillingAddress2
	PSIObj.Bcity=pcBillingCity
	PSIObj.Bstate=pcBillingState
	PSIObj.Bzip=pcBillingPostalCode
	PSIObj.Bcountry=pcBillingCountry
	PSIObj.Sname=pcShippingFirstName&" "&pcShippingLastName
	PSIObj.Saddr1=pcShippingAddress
	PSIObj.Saddr2=pcShippingAddress2
	PSIObj.Scity=pcShippingCity
	PSIObj.Sstate=pcShippingState
	PSIObj.Szip=pcShippingPostalCode
	PSIObj.Scountry=pcShippingCountryCode
	PSIObj.Phone=pcBillingPhone
	PSIObj.Fax=""
	PSIObj.Comments=""
	PSIObj.Cardnumber=Request.Form("Cardnumber")
	PSIObj.Chargetype="0"
	PSIObj.Expmonth=Request.Form("expMonth")
	PSIObj.Expyear=Request.Form("expYear")
	PSIObj.Email=pcCustomerEmail

	'Used during the testing process, 0=Live
	if psi_testmode="YES" then
	else
	PSIObj.Result=0
   end if
	'Used with AVS processing (only in the US)
	'PSIObj.Addrnum="111"
    
	'----------------------------Add items
	ItemID1=Request.Form("ItemID1")			
	Description1=Request.Form("Description1")
	Price1=Request.Form("Price1")
	Quantity1=Request.Form("Quantity1")
	SoftFile1=Request.Form("SoftFile1")
	EsdType1=Request.Form("EsdType1")
	Serial1=Request.Form("Serial1")

  intErr=0  
   
	ret_code=PSIObj.AddItem(ItemID1, Description1, Price1, Quantity1, SoftFile1, EsdType1, Serial1)
	If Not ret_code=1 Then
		Msg="ERROR   " & PSIObj.ErrMsg
		intErr=1
		Set PSIObj=Nothing
	End If


	'-------------------------------Process
	if intErr=0 then
		ret_code=PSIObj.ProcessOrder()
		If Not ret_code=1 Then
			Msg="ERROR   " & PSIObj.ErrMsg
			intErr=1
			Set PSIObj=Nothing
		End If
	end if
	
	'-------------------------------Get results
	if intErr=0 then
		pcv_Response_Approved=PSIObj.Appr
		pcv_Response_Code=PSIObj.code
		pcv_Response_TransTime=PSIObj.transtime
		pcv_Response_Refno=PSIObj.refno
		pcv_Response_Error=PSIObj.Err
		pcv_Response_Orderno=PSIObj.OrdNo
		pcv_Response_Subtotal=CStr(PSIObj.Subtotal)
		pcv_Response_Shiptotal=CStr(PSIObj.Shiptotal)
		pcv_Response_Taxtotal=CStr(PSIObj.Taxtotal)
		pcv_Response_Total=CStr(PSIObj.Total)
		
		Set PSIObj=Nothing
		
		If pcv_Response_Approved="APPROVED" then
			session("GWAuthCode")=pcv_Response_Code
			session("GWTransId")=pcv_Response_Refno
			response.redirect "gwReturn.asp?s=true&gw=PSIGate"
		Else
			Msg=pcv_Response_Error
		End if
	end if

	'*************************************************************************************
	' END
	'*************************************************************************************
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
				<form method="POST" action="<%=session("redirectPage")%>" name="PaymentForm" class="pcForms">
					<input type="hidden" name="PaymentSubmitted" value="Go">
					<input type="hidden" name="ItemID1" value="Online Order">
					<% If scCompanyName="" then %>
						<input type="hidden" name="Description1" value="Shopping Cart"> 
					<%else %>
						<input type="hidden" name="Description1" value="<%=scCompanyName%>"> 
					<% end if %>
					<input type="hidden" name="Price1" value="<%=pcBillingTotal%>"> 
					<input type="hidden" name="Quantity1" value="1">
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
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<% if psi_testmode="YES" then %>
						<tr>
							<td colspan="2"><div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_3")%></div></td>
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
					<% end if %>

					<tr class="pcSectionTitle">
						<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_GateWay_5")%></p></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></p></td>
						<td> 
							<input type="text" name="CardNumber" value="">
						</td>
					</tr>
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></p></td>
						<td><%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
							<select name="expMonth">
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
							<select name="expYear">
								<option value="<%=right(dtCurYear,2)%>" selected><%=dtCurYear%></option>
								<option value="<%=right(dtCurYear+1,2)%>"><%=dtCurYear+1%></option>
								<option value="<%=right(dtCurYear+2,2)%>"><%=dtCurYear+2%></option>
								<option value="<%=right(dtCurYear+3,2)%>"><%=dtCurYear+3%></option>
								<option value="<%=right(dtCurYear+4,2)%>"><%=dtCurYear+4%></option>
								<option value="<%=right(dtCurYear+5,2)%>"><%=dtCurYear+5%></option>
								<option value="<%=right(dtCurYear+6,2)%>"><%=dtCurYear+6%></option>
								<option value="<%=right(dtCurYear+7,2)%>"><%=dtCurYear+7%></option>
								<option value="<%=right(dtCurYear+8,2)%>"><%=dtCurYear+8%></option>
								<option value="<%=right(dtCurYear+9,2)%>"><%=dtCurYear+9%></option>
								<option value="<%=right(dtCurYear+10,2)%>"><%=dtCurYear+10%></option>
							</select>
						</td>
					</tr>
					<% If pcv_CVV="1" Then %>
						<tr> 
							<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></p></td>
							<td> 
								<input name="CVV" type="text" id="CVV" value="" size="4" maxlength="4">
							</td>
						</tr>
						<tr> 
							<td>&nbsp;</td>
							<td><img src="images/CVC.gif" alt="cvc code" width="212" height="155"></td>
						</tr>
					<% End If %>
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
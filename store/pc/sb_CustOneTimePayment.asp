<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/bto_language.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/ShipFromSettings.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="opc_contentType.asp" -->
<!--#include file="inc_sb.asp"-->
<% dim conntemp,query,rs 

HaveSecurity=0
if session("idCustomer")=0 OR session("idCustomer")="" then
	HaveSecurity=1
end if

call openDb()

'Check to see if ARB has been turned off by admin, then display message
If scSBStatus="0" then
	response.redirect "msg.asp?message=212"
End If 

Call SetContentType()

IF HaveSecurity=0 THEN

	qry_GUID = getUserInput(Request("GUID"),0)
	qry_ID = getUserInput(Request("ID"),0)
	if not validNum(qry_ID) then
	   qry_ID=0
	end if

	if request("action")="add" then

		query="SELECT Setting_APIUser,Setting_APIPassword,Setting_APIKey,Setting_RegSuccess FROM SB_Settings;"
		set rsAPI=connTemp.execute(query)
		if not rsAPI.eof then
			Setting_APIUser=rsAPI("Setting_APIUser")
			Setting_APIPassword=enDeCrypt(rsAPI("Setting_APIPassword"), scCrypPass)
			Setting_APIKey=enDeCrypt(rsAPI("Setting_APIKey"), scCrypPass)
		end if
		set rsAPI=nothing
		
		pcBillingFirstName=getUserInput(request.Form("billfname"),0)
		pcBillingLastName=getUserInput(request.Form("billlname"),0)
		pcBillingCompany=getUserInput(request.Form("billcompany"),0)
		pcBillingAddress=getUserInput(request.Form("billaddr"),0)
		pcBillingAddress2=getUserInput(request.Form("billaddr2"),0)
		pcBillingCity=getUserInput(request.Form("billcity"),0)
		pcBillingPostalCode=getUserInput(request.Form("billzip"),0)
		pcBillingStateCode=getUserInput(request.Form("billstate"),0)
		pcBillingProvince=getUserInput(request.Form("billprovince"),0)
		pcBillingCountryCode=getUserInput(request.Form("billcountry"),0)

	
		CardNumber=getUserInput(request.Form("CardNumber"),0)
		expYear=getUserInput(request.Form("expYear"),0)
		expMonth=getUserInput(request.Form("expMonth"),0)
		CVV=getUserInput(request.Form("CVV"),0)
		CC_TYPE=getUserInput(request.Form("creditCardType"),0)
		RegularAmt=getUserInput(Session("pcSF_OutstandingBalance"),0)
		GUID=getUserInput(request.Form("GUID"),0)

		Set objSB = NEW pcARBClass
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// BILLING ADDRESS
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
		objSB.BillingFirstName = pcBillingFirstName
		objSB.BillingLastName = pcBillingLastName
		objSB.BillingCompany = pcBillingCompany
		objSB.BillingAddress = pcBillingAddress
		objSB.BillingAddress2 = pcBillingAddress2
		objSB.BillingCity = pcBillingCity
		objSB.BillingPostalCode = pcBillingPostalCode
		objSB.BillingStateCode = pcBillingStateCode
		objSB.BillingProvince = pcBillingProvince
		objSB.BillingCountryCode = pcBillingCountryCode
		objSB.BillingPhone = pcBillingPhone
		objSB.CustomerEmail = pcCustomerEmail

		objSB.CartRegularAmt = RegularAmt
		objSB.GUID = GUID
		objSB.PayInfoType="CC"
		objSB.PayInfoExpMonth= expMonth
		If len(expYear)=2 Then
			objSB.PayInfoExpYear = "20" & expYear	
		Else
			objSB.PayInfoExpYear = expYear	
		End If
		objSB.PayInfoCardNumber = left(CardNumber,16)
		objSB.PayInfoAccountNumber = right(PayInfoCardNumber,4)
		objSB.PayInfoCardType = CC_TYPE
		objSB.PayInfoCVVNumber = CVV
		
		Dim result
		result = objSB.OneTimePaymentRequest(Setting_APIUser, Setting_APIPassword, Setting_APIKey)

		If len(SB_ErrMsg)>0 Then	
			UpdateSuccess="0"
		Else
			UpdateSuccess="1"
		End If 
		
		Set objSB = Nothing

	end if
	
END IF
%>
<html>
<head>
<TITLE><%=dictLanguage.Item(Session("language")&"_opc_gwa_title")%></TITLE>
<%
Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN  

Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")
%>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
<script language="JavaScript">
<!--
	
	function Form1_Validator(theForm)
	{
		return (true);
	}

//-->
</script>
</head>
<body style="margin: 0;">
<div id="pcMain">
<form method="post" name="BillingForm" id="BillingForm" action="sb_CustOneTimePayment.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
<input name="ID" type="hidden" value="<%=qry_ID%>">
<input name="GUID" type="hidden" value="<%=qry_GUID%>">
<table class="pcMainTable">
	<%IF HaveSecurity=1 THEN%>
        <tr>
            <td colspan="2">
                <div class="pcErrorMessage"><%response.write dictLanguage.Item(Session("language")&"_SB_20")%></div>
            </td>
        </tr>
    <%ELSE%>
    
        <%IF UpdateSuccess="1" THEN%>
            <tr>
                <td colspan="2">
                    <div class="pcSuccessMessage"><%response.write dictLanguage.Item(Session("language")&"_SB_22")%></div>
                    <script>
                        setTimeout(function(){parent.closeManageSubscription()},4000);
                    </script>
                </td>
            </tr>        
		<%ELSE%>            


				<% If SB_ErrMsg <> "" Then %>
                <tr>
                    <td colspan="2">
                        <div class="pcErrorMessage"><%=SB_ErrMsg%></div>
                    </td>
                </tr>
				<% End If %>
                <tr>
                	<td colspan="2"><h2><%response.write dictLanguage.Item(Session("language")&"_SB_15")%></h2></td>
                </tr>
				<tr>
					<td colspan="2">
					
                        <table class="pcShowContent">
                        		<!--#include file="../includes/pcServerSideValidation.asp" -->
								<%
                                '// Get Bill Information from database if Registered Customer
                                if Session("idCustomer")>0 then
                                    query="SELECT idcustomer, customers.pcCust_Guest, [name], lastName, customerCompany, phone, email, address, zip, stateCode, state, city, countryCode, address2, suspend, idCustomerCategory, customerType FROM customers WHERE ((customers.idcustomer)="&session("idCustomer")&");"
                                    set rs=connTemp.execute(query)
                                    if not rs.eof then
                                        pcStrBillingFirstName=rs("name")
                                        pcStrBillingLastName=rs("lastName")
                                        pcStrBillingCompany=rs("customerCompany")
                                        pcStrBillingPhone=rs("phone")
                                        pcStrCustomerEmail=rs("email")
                                        pcStrBillingAddress=rs("address")
                                        pcStrBillingPostalCode=rs("zip")
                                        pcStrBillingStateCode=rs("stateCode")
                                        pcStrBillingProvince=rs("state")
                                        pcStrBillingCity=rs("city")
                                        pcStrBillingCountryCode=rs("countryCode")
                                        pcStrBillingAddress2=rs("address2")
                                    end if
                                set rs=nothing
                                end if 
                              %>
                                
                             <tr>
                                <td width="16%" nowrap><%=dictLanguage.Item(Session("language")&"_opc_10")%></td>
                                <td width="34%"><input type="text" name="billfname" id="billfname" value="<%=pcStrBillingFirstName%>" /></td>
                                <td width="16%" nowrap><%=dictLanguage.Item(Session("language")&"_opc_11")%></td>
                                <td width="34%"><input type="text" name="billlname" id="billlname" value="<%=pcStrBillingLastName%>" /></td>
                              </tr>
                              <tr>
                                <td nowrap><%=dictLanguage.Item(Session("language")&"_opc_12")%></td>
                                <td colspan="3"><input type="text" name="billcompany" id="billcompany" value="<%=pcStrBillingCompany%>" /></td>
                              </tr>
                              <tr>
                                <td nowrap><%=dictLanguage.Item(Session("language")&"_opc_13")%></td>
                                <td><input type="text" name="billaddr" id="billaddr" value="<%=pcStrBillingAddress%>" /></td>
                                <td nowrap><%=dictLanguage.Item(Session("language")&"_opc_14")%></td>
                                <td><input type="text" name="billaddr2" id="billaddr2" value="<%=pcStrBillingAddress2%>" /></td>
                              </tr>
                              <tr>
                                <td nowrap><%=dictLanguage.Item(Session("language")&"_opc_15")%></td>
                                <td><input type="text" name="billcity" id="billcity" value="<%=pcStrBillingCity%>" /></td>
                                <td nowrap><%=dictLanguage.Item(Session("language")&"_opc_16")%></td>
                                <td><input type="text" name="billzip" id="billzip" value="<%=pcStrBillingPostalCode%>" /></td>
                              </tr>
                                <%
                                pcv_strTargetForm = "BillingForm" '// Name of Form
                                pcv_strCountryBox = "billcountry" '// Name of Country Dropdown
                                pcv_strTargetBox = "billstate" '// Name of State Dropdown
                                pcv_strProvinceBox =  "billprovince" '// Name of Province Field
                            
                                '// Set local Country to Session
                                if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
                                    Session(pcv_strSessionPrefix&pcv_strCountryBox) = pcStrBillingCountryCode
                                end if
                            
                                '// Set local State to Session
                                if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
                                    Session(pcv_strSessionPrefix&pcv_strTargetBox) = pcStrBillingStateCode
                                end if
                            
                                '// Set local Province to Session
                                if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
                                    Session(pcv_strSessionPrefix&pcv_strProvinceBox) =  pcStrBillingProvince
                                end if
                                %>
                                <!--#include file="../includes/javascripts/opc_pcStateAndProvince.asp"-->
                                <%
                                pcs_CountryDropdown
                                %>
    
                                <%
                                '// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
                                pcs_StateProvince
                                %>
                              <tr>
                                <td nowrap><%=dictLanguage.Item(Session("language")&"_opc_17")%></td>
                                <td><input type="text" name="billphone" id="billphone" value="<%=pcStrBillingPhone%>" /></td>
                                <td nowrap><%=dictLanguage.Item(Session("language")&"_opc_18")%></td>
                                <td><input type="text" name="billfax" id="billfax" value="<%=pcStrBillingFax%>"/></td>
                              </tr>
                        </table>
                    
					</td>
				</tr>
				<tr>
					<td colspan="2"><hr></td>
				</tr>    
				<tr>
					<td colspan="2">
                    
						<table class="pcShowContent">
                              <tr>
                                  <td nowrap="nowrap" width="20%"><p><%=dictLanguage.Item(Session("language")&"_GateWay_12")%></p></td>
                                  <td width="80%">
  
                                      <select name="creditCardType">		
                                          <option value="Visa" <%if CC_TYPE="Visa" then %>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_GateWay_16")%></option>
                                          <option value="MasterCard" <%if CC_TYPE="MasterCard" then %>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_GateWay_17")%></option>
                                          <option value="Discover" <%if CC_TYPE="Discover" then %>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_GateWay_18")%></option>
                                          <option value="Amex" <%if CC_TYPE="Amex" then %>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_GateWay_19")%></option>
                                          <% If PaymentAction="Authorization" AND pcPay_PayPal_Currency="GBP" Then %>
                                          <option value="Maestro" <%if CC_TYPE="Maestro" then %>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_GateWay_20")%></option>
                                          <option value="Solo" <%if CC_TYPE="Solo" then %>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_GateWay_21")%></option>
                                          <% End If %>
                                      </select>							
                                      
                              	  </td>
                              </tr>
                              <tr> 
                                  <td nowrap="nowrap" width="20%"><p><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></p></td>
                                  <td> 
                                      <input type="text" name="CardNumber" value="">
                                  </td>
                              </tr>
                              <tr> 
                                  <td colspan="4">
  
                                      <%
                                      '// Visa/ MasterCard/ Discover/ AMEX
                                      %>
                                      <table width="100%">
                                          <tr> 
                                              <td nowrap="nowrap" width="20%"><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></td>
                                              <td width="80%"><%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
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
                                                      <option value="<%=(dtCurYear)%>" selected><%=dtCurYear%></option>
                                                      <option value="<%=(dtCurYear+1)%>"><%=dtCurYear+1%></option>
                                                      <option value="<%=(dtCurYear+2)%>"><%=dtCurYear+2%></option>
                                                      <option value="<%=(dtCurYear+3)%>"><%=dtCurYear+3%></option>
                                                      <option value="<%=(dtCurYear+4)%>"><%=dtCurYear+4%></option>
                                                      <option value="<%=(dtCurYear+5)%>"><%=dtCurYear+5%></option>
                                                      <option value="<%=(dtCurYear+6)%>"><%=dtCurYear+6%></option>
                                                      <option value="<%=(dtCurYear+7)%>"><%=dtCurYear+7%></option>
                                                      <option value="<%=(dtCurYear+8)%>"><%=dtCurYear+8%></option>
                                                      <option value="<%=(dtCurYear+9)%>"><%=dtCurYear+9%></option>
                                                      <option value="<%=(dtCurYear+10)%>"><%=dtCurYear+10%></option>
                                                  </select>
                                              </td>
                                          </tr>					
          
                                      </table>
                                  
                                  </td>
                              </tr>
                              <tr> 
                                  <td nowrap="nowrap" width="20%"><p><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></p></td>
                                  <td> 
                                      <input name="CVV" type="text" id="CVV" value="" size="4" maxlength="4">
                                  </td>
                              </tr>
                        </table>
                    
					</td>
				</tr>
				<tr>
					<td colspan="2"><hr></td>
				</tr>
                <tr>
                    <td colspan="2" class="pcSpacer"></td>
                </tr>
				<%
                  query="SELECT Setting_APIUser,Setting_APIPassword,Setting_APIKey,Setting_RegSuccess FROM SB_Settings;"
                  set rsAPI=connTemp.execute(query)
                  if not rsAPI.eof then
                      Setting_APIUser=rsAPI("Setting_APIUser")
                      Setting_APIPassword=enDeCrypt(rsAPI("Setting_APIPassword"), scCrypPass)
                      Setting_APIKey=enDeCrypt(rsAPI("Setting_APIKey"), scCrypPass)
                  end if
                  set rsAPI=nothing                  
                  
                  Set objSB = NEW pcARBClass                  
                  objSB.GUID = qry_GUID
                  If scSBLanguageCode<>"" Then
                      objSB.CartLanguageCode = scSBLanguageCode
                  Else
                      objSB.CartLanguageCode = "en-EN"
                  End If
                
                  result = objSB.GetSubscriptionDetailsRequest(Setting_APIUser, Setting_APIPassword, Setting_APIKey)

                  If SB_ErrMsg="" Then 
                     
						'// Total Balance
					  	pcv_strBalance = objSB.pcf_GetNode(result, "BalanceTotal", "//GetSubscriptionDetailsResponse/Subscription")
						
						'// Reason
						pcv_strReason = ""
						Set myXmlDoc = Server.CreateObject("Msxml2.DOMDocument"&scXML)				 
						myXmlDoc.loadXml(result)
						Set Nodes = myXmlDoc.selectnodes("//GetSubscriptionDetailsResponse/Subscription/BalanceDetails/Balance")	
						For Each Node In Nodes	
							pcv_strReason = pcv_strReason & "<li>" & objSB.pcf_CheckNode(Node,"Reason","")		 & "</li>"		
						Next
						Set Node = Nothing
						Set Nodes = Nothing
						Set myXmlDoc = Nothing
						                  
                  End If

                '// Query the remaining balance
                Dim OutstandingBalance
                OutstandingBalance = pcv_strBalance
				Session("pcSF_OutstandingBalance")=OutstandingBalance
                %>  
                <tr>
                    <td nowrap="nowrap" width="20%"><p><%response.write dictLanguage.Item(Session("language")&"_SB_19")%></p></td>
                    <td width="80%">	
                          <%=money(OutstandingBalance)%>			
                    </td>
                </tr>
                <tr>
                    <td nowrap="nowrap" width="20%"><p><%response.write dictLanguage.Item(Session("language")&"_SB_18")%></p></td>
                    <td width="80%">
						<%=pcv_strReason%>
                    </td>
                </tr>
                <tr>
                    <td colspan="2" class="pcSpacer"></td>
                </tr>	
                <tr>
                    <td colspan="2"> 
                        <input type="image" id="DSubmit" name="DSubmit" value="DSubmit" src="images/sample/pc_button_update.gif" border="0" class="clearBorder">
                    </td>
                </tr>
        <%END IF%>
    <%END IF%>
</table>
</form>
</div>
</body>
</html>
<% call closedb() %>
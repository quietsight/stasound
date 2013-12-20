<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>

<% response.Buffer=true
Response.Expires = -1
Response.Expiresabsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#INCLUDE FILE="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/bto_language.asp"--> 
<!--#include file="../includes/rewards_language.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include FILE="../includes/SearchConstants.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/securitysettings.asp" -->
<!--#include file="../includes/MailUpFunctions.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/rc4.asp" -->
<%Dim iAddDefaultPrice,	iAddDefaultWPrice%>
<%
'*******************************
' Page Name
'*******************************
Dim pcStrPageName
pcStrPageName = "OnePageCheckout.asp"

Dim TurnOffDiscountCodesWhenHasSale, HavePrdsOnSale

TurnOffDiscountCodesWhenHasSale=scDisableDiscountCodes
'=1: True - Default
'=0: False

HavePrdsOnSale=0

'*******************************
' Pay Panel Open or Closed
'*******************************
Dim pcv_strPayPanel
pcv_strPayPanel=getUserInput(request("PayPanel"),2)

'*******************************
' START: Check store on/off, start PC session, check affiliate ID
'*******************************
%>
<!--#include file="pcStartSession.asp"-->
<%
'*******************************
' END: Check store on/off, start PC session, check affiliate ID
'*******************************

'*****************************************************************************************************
'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
%><!--#include file="pcVerifySession.asp"--><%
pcs_VerifySession
'*****************************************************************************************************
'// END: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
pcCartArray=Session("pcCartSession")
ppcCartIndex=Session("pcCartIndex")

if countCartRows(pcCartArray, ppcCartIndex)=0 then
	response.redirect "msg.asp?message=9" 
end if

%><!--#include file="inc_checkPrdQtyCart.asp"--><%
Call CheckALLCartStock()
%>
<!--#include file="DBsv.asp"-->
<%
'SB S
session("pcIsSubscription") = ""
pcIsSubscription = findSubscription(pcCartArray, ppcCartIndex)			
session("pcIsSubscription") = pcIsSubscription
'SB E

If session("customerType")=1 Then
	if calculateCartTotal(pcCartArray, ppcCartIndex)<scWholesaleMinPurchase then  
		response.redirect "msgb.asp?message="&Server.URLEncode(dictLanguage.Item(Session("language")&"_checkout_2")& scCurSign &money(scWholesaleMinPurchase)& dictLanguage.Item(Session("language")&"_techErr_3") & "<BR><BR><a href=viewcart.asp>"& dictLanguage.Item(Session("language")&"_msg_back") &"</a>")
	end if
Else
	if calculateCartTotal(pcCartArray, ppcCartIndex)<scMinPurchase then  
		response.redirect "msgb.asp?message="&server.URLEncode(dictLanguage.Item(Session("language")&"_checkout_2")& scCurSign &money(scMinPurchase)&"<BR><BR><a href=viewcart.asp>"& dictLanguage.Item(Session("language")&"_msg_back") &"</a>")
	end if
End If

if session("ExpressCheckoutPayment") <> "YES" then
%>
<!--#include file="opc_checkpayment.asp"-->
<%
else
	session("pcSFIdPayment")=999999
end if

'MailUp-S

Dim MaxRequestTime,StopHTTPRequests

'maximum seconds for each HTTP request time
MaxRequestTime=5

StopHTTPRequests=0

'MailUp-E

'MAILUP-S

	tmp_setup=0
	pcMailUpSett_APIUser=""
	pcMailUpSett_APIPassword=""
	pcMailUpSett_URL=""

	call opendb()
	query="SELECT pcMailUpSett_APIUser,pcMailUpSett_APIPassword,pcMailUpSett_URL,pcMailUpSett_AutoReg,pcMailUpSett_RegSuccess FROM pcMailUpSettings;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcMailUpSett_APIUser=rs("pcMailUpSett_APIUser")
		session("SF_MU_APIUser")=pcMailUpSett_APIUser
		pcMailUpSett_APIPassword=enDeCrypt(rs("pcMailUpSett_APIPassword"), scCrypPass)
		session("SF_MU_APIPassword")=pcMailUpSett_APIPassword
		pcMailUpSett_URL=rs("pcMailUpSett_URL")
		session("SF_MU_URL")=pcMailUpSett_URL
		tmp_Auto=rs("pcMailUpSett_AutoReg")
		if IsNull(tmp_Auto) or tmp_Auto="" then
			tmp_Auto=0
		end if
		session("SF_MU_Auto")=tmp_Auto
		tmp_setup=rs("pcMailUpSett_RegSuccess")
		if IsNull(tmp_setup) or tmp_setup="" then
			tmp_setup=0
		end if
		session("SF_MU_Setup")=tmp_setup
	end if
	set rs=nothing
	call closedb()

'MAILUP-E

call openDb()

    '// We'll check the cart stock levels on entry (plus each time the ajax panels slide)
	Dim strCCSLCheck
	strCCSLcheck = checkCartStockLevels(pcCartArray, pcCartIndex, aryBadItems)
	If Len(Trim(strCCSLCheck))>0 Then
		response.redirect "viewcart.asp"
	End If

'// Vat Settings
pcv_ShowVatId = false
pcv_isVatIdRequired = false
pcv_ShowSSN = false
pcv_isSSNRequired = false
if pshowVatID="1" then pcv_ShowVatId = true
if pVatIdReq="1" then pcv_isVatIdRequired = true
if pshowSSN="1" then pcv_ShowSSN = true
if pSSNReq="1" then pcv_isSSNRequired = true

%>
<!--#include file="header.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<%'MailUp-S%>
<script>
function newWindow(file,window) {
	msgWindow=open(file,window,'resizable=no,width=530,height=150');
	if (msgWindow.opener == null) msgWindow.opener = self;
}
</script>
<%'MailUp-E%>
<!--#include file="../includes/pcServerSideValidation.asp" -->
<link href="../includes/spry/SpryAccordionOPC.css" rel="stylesheet" type="text/css" />
<script src="../includes/spry/SpryEffects.js" type="text/javascript" > </script>
<script src="../includes/spry/SpryAccordionOPC.js" type="text/javascript"></script>
<!--#include file="onepagecheckoutJS.asp" -->
<!--#include file="pcValidateQty.asp"-->
<!--#include file="inc_PrdCatTip.asp"-->
<div id="pcMain">
<a name="top"></a>

<%'// Dialogs - START %>
<div id="AskEnterPassDialog" title="<%=dictLanguage.Item(Session("language")&"_opc_35")%>" style="display:none">
	<div id="AskEnterPassMsg" class="ui-main"><%=dictLanguage.Item(Session("language")&"_opc_37")%></div>
</div>
<%'MailUp-S%>
<script>var tmpNListChecked=0;</script>
<div id="PleaseWaitDialogMU" title="<%=dictLanguage.Item(Session("language")&"_MailUp_SynNote4")%>" style="display:none">
	<div id="PleaseWaitMsgMU">
		<img src="images/ajax-loader1.gif" width="20" height="20" align="absmiddle"> <font size="2"><%=dictLanguage.Item(Session("language")&"_MailUp_SynNote3")%></font>
    </div>
</div>
<%'MailUp-E%>
<div id="PleaseWaitDialog" title="" style="display:none">
	<div id="PleaseWaitMsg" class="ui-main"></div>
</div>
<div id="GlobalAjaxErrorDialog" title="<%=dictLanguage.Item(Session("language")&"_opc_js_68")%>" style="display:none">
	<div class="pcErrorMessage">
		<%=dictLanguage.Item(Session("language")&"_ajax_globalerror")%>
	</div>
</div>
<div id="GlobalErrorDialog" title="<%=dictLanguage.Item(Session("language")&"_opc_js_69")%>" style="display:none">
	<div id="GlobalErrorMsg" class="pcErrorMessage">
	</div>
</div>
<div id="ValidationDialog" title="<%=dictLanguage.Item(Session("language")&"_opc_js_69")%>" style="display:none">
	<div id="ValidationErrorMsg" class="pcErrorMessage">
	</div>
</div>
<div id="GWDialog" title="Gift Wrapping" style="display:none; overflow:hidden">
	<div id="GWframeloader" style="overflow:hidden"><img src="images/ajax-loader1.gif" width="20" height="20" align="absmiddle"><%=dictLanguage.Item(Session("language")&"_opc_loadcontent")%></div>
	<iframe name="GWframe" id="GWframe" src="" width="600" height="450" frameborder="0" scrolling="auto"></iframe>
</div>
<%'// Dialogs - END %>

<div id="opcMainContainer">
	<table class="pcMainTable">
        <tr>
            <td>
				<% '// ORDER TOTAL - START %>
                    <div id="pcOPCtotal">
                        <div id="pcOPCtotalAmount"><%=scCurSign & money(opcOrderTotal)%></div>
                        <div id="pcOPCtotalLinks"><a href="#orderPreview"><%=dictLanguage.Item(Session("language")&"_opc_1")%></a></div>
                    </div>
                <% '// ORDER TOTAL - END %>
                <h1>Checkout</h1>
            </td>
        </tr>
		<%if request("msg")<>"" and IsNumeric(request("msg")) then%>
		<tr>
			<td>
				<div class="pcErrorMessage">
					<%Select case Clng(request("msg"))
						Case 1:
							response.write dictLanguage.Item(Session("language")&"_opc_msg1")
					End Select%>
				</div>
			</td>
		</tr>
		<%end if%>
        <% if session("ExpressCheckoutPayment") = "YES" then %>
		<tr>
			<td>
				<div class="pcInfoMessage">
					<%=dictLanguage.Item(Session("language")&"_opc_55")%>
				</div>
			</td>
		</tr>
        <% end if %>
		<tr>
			<td>
                <%
				'/////////////////////////////////////////////////////////////////////////
				'// LOGIN:  START
				'/////////////////////////////////////////////////////////////////////////
				%>              
                <table class="pcShowContent">
                    <tr>
                      <td colspan="2">
                        <%
                        if (Session("idCustomer")=0 or Session("idCustomer")="") OR (session("idCustomer")>"0" AND session("CustomerGuest")>"0") then
                            'session("pcSFIdDbSession")=""
                            'session("pcSFRandomKey")=""
                            %>
                            <div id="LoginOptions">                                      
                              <table id="opcLoginTable">
                                <tr>
                                  <td class="leftCell">
                                        <h3><%=dictLanguage.Item(Session("language")&"_opc_3")%></h3>
                                        <form name="loginForm" id="loginForm">
                                            <div id="pcShowLoginFields">
                                                <table>
                                                    <tr>
                                                        <td><%=dictLanguage.Item(Session("language")&"_opc_5")%></td>
                                                        <td><input type="text" id="email" name="email" size="25"></td>
                                                    </tr>
                                                    <tr>
                                                        <td><%=dictLanguage.Item(Session("language")&"_opc_6")%></td>
                                                        <td><input type="password" id="password" name="password" size="25" autocomplete="off"></td>
                                                    </tr>
													<% 'If Advanced Security is turned on
                                                    if scSecurity=1 then
                                                        Session("store_userlogin")="1"
                                                        session("store_adminre")="1"	
                                                        if (scUserLogin=1 OR scUserReg=1) and (scUseImgs=1) then %>
                                                            <tr>
                                                                <td colspan="2" class="pcSpacer"></td>
                                                            </tr>
                                                            <tr>
                                                                <td></td>
                                                                <td>
                                                                  	<!--#include file="../CAPTCHA/CAPTCHA_form_inc.asp" -->
                                                                </td>
                                                            </tr>
                                                        <% else 
                                                            response.write "<div id=""show_security""></div>"
                                                        end if %>
                                                    <% else
                                                        response.write "<div id=""show_security""></div>"
                                                    end if %>
                                                    <tr>
                                                        <td></td>
                                                        <td>
                                                            <div id="LoginLoader" style="display:none"></div>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td>&nbsp;</td>
                                                        <td>
                                                        	<input type="image" name="LoginSubmit" id="LoginSubmit" src="<%=RSlayout("login")%>" align="absmiddle" border="0">
															<div style="margin: 8px 2px 0 2px;" class="pcSmallText"><%response.write dictLanguage.Item(Session("language")&"_Custva_3")%></div><div style="margin: 4px 2px 6px 2px;" class="pcSmallText">
			<a href="checkout.asp?cmode=2&fmode=<%=pcPageMode%>&orderReview=no"><%response.write dictLanguage.Item(Session("language")&"_Custva_8")%></a></div>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </div>
                                        </form>
                                  </td>
                                  <td class="rightCell">
                                    <div id="pcShowLoginFields2">                                            
                                        <h3><%if scGuestCheckoutOpt=0 OR scGuestCheckoutOpt="" then%><%=dictLanguage.Item(Session("language")&"_opc_4")%><%else%><%if scGuestCheckoutOpt="1" then%><%=dictLanguage.Item(Session("language")&"_opc_4d")%><%else%><%if scGuestCheckoutOpt="2" then%><%=dictLanguage.Item(Session("language")&"_opc_4f")%><%end if%><%end if%><%end if%></h3>
                                        <div style="padding: 0 5px 5px 5px;"><%if scGuestCheckoutOpt=0 OR scGuestCheckoutOpt="" then%><%=dictLanguage.Item(Session("language")&"_opc_4b")%><%else%><%if scGuestCheckoutOpt="1" then%><%=dictLanguage.Item(Session("language")&"_opc_4e")%><%else%><%if scGuestCheckoutOpt="2" then%><%=dictLanguage.Item(Session("language")&"_opc_4g")%><%end if%><%end if%><%end if%></div>
                                        <div style="padding: 10px 5px 5px 5px;"><input type="image" name="GuestSubmit" id="GuestSubmit" src="<%=RSlayout("pcLO_Update")%>" align="absmiddle" border="0" onclick="$('#LoginOptions').hide(); $('#acc1').show(); acc1.openPanel('opcLogin'); $('#BillingArea').show(); <%if scGuestCheckoutOpt="2" then%>document.BillingForm.billemail.focus();<%else%>document.BillingForm.billfname.focus();<%end if%> "></div>
										<!--#include file="pcPay_PayPal.asp"-->
                                        <!--#include file="pcPay_GoogleCheckout.asp"-->                                  
                                    </div>
                                  </td>
                                </tr>
                              </table>
                        	</div>
                        <%
                        else
                            if session("CustomerGuest")="0" then
								pcIntCustomerId = session("idCustomer")
                                if not validNum(pcIntCustomerId) then
                                    session("idCustomer") = Cdbl(0)
                                    response.Redirect("default.asp")
                                end if
							
                                ' START - Retrieve customer name
                                if session("pcStrCustName") = "" then
                                
									query = "SELECT name, lastName FROM customers WHERE idCustomer = " & pcIntCustomerId
									set rs = Server.CreateObject("ADODB.Recordset")
									set rs = conntemp.execute(query)
									if not rs.eof then
										pcStrCustName = rs("name") & " " & rs("lastName")
										session("pcStrCustName") = pcStrCustName
									end if
									set rs = nothing
								end if
								' END - Retrieve customer name
								
								query = "SELECT suspend FROM customers WHERE idCustomer = " & pcIntCustomerId
                                set rs = Server.CreateObject("ADODB.Recordset")
                                set rs = conntemp.execute(query)
								pcIntSuspend=0
								if not rs.eof then
									pcIntSuspend=rs("suspend")
								end if
                                set rs = nothing
								if pcIntSuspend="1" then
									call closedb()
									response.clear
									response.redirect "msg.asp?message=310"
								end if
								%>
								<%=dictLanguage.Item(Session("language")&"_opc_7")%>&nbsp;<a href="custPref.asp"><%=session("pcStrCustName")%></a><br><br>
								<%
                            end if
                        end if
                        %>
                      </td>
                    </tr>
                </table>
                <%
				'/////////////////////////////////////////////////////////////////////////
				'// LOGIN:  END
				'/////////////////////////////////////////////////////////////////////////
				%> 
				<%
				'// Check Gift Registry
				if session("Cust_IDEvent")<>"" then
					query="select pcEv_IDCustomer, pcEv_Delivery, pcEv_MyAddr, pcEv_HideAddress from PcEvents where pcEv_IDEvent=" & session("Cust_IDEvent")
					set rstemp=connTemp.execute(query)					
					if not rstemp.eof then
						gIDCustomer=rstemp("pcEv_IDCustomer")
						gDelivery=rstemp("pcEv_Delivery")
						if gDelivery<>"" then
						else
							gDelivery=0
						end if
						gMyAddr=rstemp("pcEv_MyAddr")
						if gMyAddr<>"" then
						else
							gMyAddr=0
						end if
						if gDelivery="1" then
							GRTest=1
						end if
						gHideAddress=rstemp("pcEv_HideAddress")
						session("gHideAddress")=gHideAddress						
					end if
					set rstemp=nothing
				end if

				'// Set global shipping mode to local
				pcv_AlwAltShipAddress = scAlwAltShipAddress

				pcv_NOShippingAtAll = "1" 
				query="SELECT active FROM ShipmentTypes WHERE active<>0"
				set rs=connTemp.execute(query)
				if rs.eof then '// There are NO active dynamic shipping services
					pcv_NoDynamicShipping="1"
				end if
				
				query="SELECT idFlatShipType FROM FlatShipTypes"
				set rs=connTemp.execute(query)
				if rs.eof then '// There are NO active custom shipping services
					pcv_NoCustomShipping="1"
				end if
				if pcv_NoDynamicShipping="1" and pcv_NoCustomShipping="1" then '// There are NO active shipping options
					pcv_AlwAltShipAddress = "1"
					pcv_NOShippingAtAll = "2"
				end if

				pcCartArray=Session("pcCartSession")
				pcCartIndex=Session("pcCartIndex")
				ppcCartIndex=Session("pcCartIndex")

				'// If No shipping at all is still set to "1" - check if products in cart qualify for no shipping and if so - is the store owner hiding the address?
				pShipTotal=Cdbl(calculateShipCartTotal(pcCartArray, ppcCartIndex))
				pShipWeight=Cdbl(calculateShipWeight(pcCartArray, ppcCartIndex))
				pShipQuantity=Int(calculateCartShipQuantity(pcCartArray, ppcCartIndex))

				if session("Cust_IDEvent")="" AND pShipTotal=0 AND pShipWeight=0 AND pShipQuantity=0 AND scHideShipAddress="1" AND pcv_NOShippingAtAll="1" then
					pcv_NOShippingAtAll = "2"
				end if
				if session("Cust_IDEvent")<>"" then
					pcv_AlwAltShipAddress = "2"
				end if
				%>
			  <script>
					<%if Session("CustomerGuest")="0" AND Session("IdCustomer")>"0" then%>
						var tmpNewCust=0;
					<%else%>
						var tmpNewCust=1
					<%end if%>
					
					var NeedPreLoadShipContent=0;
					var HaveShipTypeArea=0;
					var HaveDeliveryArea=0;
					
					<%if pcv_NOShippingAtAll = "1" then%>
						var NeedLoadShipChargeContent=1;
					<%else%>
						var NeedLoadShipChargeContent=0;
					<%end if%>
					
					<%if session("Cust_IDEvent")<>"" then%>
						var HaveGRAddress=1;
						<%if gDelivery=1 then%>
							var GRAddrOnly=1;
						<%else%>
							var GRAddrOnly=0;
						<%end if%>
					<%else%>
						var HaveGRAddress=0;
						var GRAddrOnly=0;
					<%end if%>
				</script>
                <div id="acc1" class="Accordion">
                <%
				'/////////////////////////////////////////////////////////////////////////
				'// ADDRESS:  START
				'/////////////////////////////////////////////////////////////////////////
				%>
                <a name="opcLoginAnchor"></a>  
                <div class="AccordionPanel" id="opcLogin">
                    <div class="AccordionPanelTab">                   
                        <div class="StatusIndicators">
                        	<a id="btnEditCO" href="javascript:;" onclick="javascript: acc1.openPanel('opcLogin'); $('#BillingAddress').hide();" alt="<%=dictLanguage.Item(Session("language")&"_opc_45")%>" title="<%=dictLanguage.Item(Session("language")&"_opc_45")%>" style="display:none;">Edit</a>
                        	<img id="btnOKCO" src="images/pc_checkmark_sm.gif" alt="<%=dictLanguage.Item(Session("language")&"_opc_46")%>" title="<%=dictLanguage.Item(Session("language")&"_opc_46")%>" <%If (session("ExpressCheckoutPayment")="YES") OR (NOT (session("idCustomer")>0 and session("CustomerGuest")="0")) then%>style="display:none"<%end if%>>
                            <img id="btnErrorCO" src="images/pc_icon_error.gif" width="18" height="18" border="0" alt="<%=dictLanguage.Item(Session("language")&"_opc_47")%>" title="<%=dictLanguage.Item(Session("language")&"_opc_47")%>" style="display:none;">                        
                        </div>                        
                        <span class="pcCheckoutTitle"><%=dictLanguage.Item(Session("language")&"_opc_2")%></span>                    
                    	<div id="BillingAddress"></div>
                        <div id="ShippingAddress"></div>
                    </div>                    
                    <div class="AccordionPanelContent">

						<%		
                        pcIntSuspend=0
                        pcIntRecvNews=0
                        pcAgreeTerms=0
						
                        '//* Get Bill Information from database if Registered Customer
                        if Session("idCustomer")>0 then
                        	query="SELECT idcustomer, customers.pcCust_Guest, customers.pcCust_VATID, customers.pcCust_SSN, [name], lastName, customerCompany, phone, email, address, zip, stateCode, state, city, countryCode, address2, suspend, idCustomerCategory, customerType, RecvNews, fax, pcCust_Locked,pcCust_AgreeTerms FROM customers WHERE ((customers.idcustomer)="&session("idCustomer")&");"
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
								pcIntSuspend=rs("suspend")
								pcIntRecvNews=rs("RecvNews")
								pcStrBillingFax=rs("fax")
								pcStrBillingVATID=Trim(rs("pcCust_VATID"))
								pcStrBillingSSN=Trim(rs("pcCust_SSN"))
								pcAgreeTerms=rs("pcCust_AgreeTerms")
								if IsNull(pcAgreeTerms) OR pcAgreeTerms="" then
									pcAgreeTerms=0
								end if
							end if
                        set rs=nothing
                        end if %>
                        
						
						<%
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						'// Billing Address - Start
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
						%>
						<div id="BillingArea" style="display:none">
                            <div class="pcCheckoutSubTitle"><%=dictLanguage.Item(Session("language")&"_opc_8")%></div>
                            <form name="BillingForm" id="BillingForm">
                            <table class="pcShowContent">
							  <%if scGuestCheckoutOpt=2 then%>
							  <%if Not (Session("idCustomer")>0 AND session("CustomerGuest")="0") then%>
							  <tr>
							 	<td id="billArea1" nowrap><%=dictLanguage.Item(Session("language")&"_opc_5")%></td>
                                <td id="billArea2" colspan="3"><input type="text" name="billemail" id="billemail" value="<%=pcStrCustomerEmail%>" size="40" /></td>
							  </tr>
							  <tr id="billArea3">
							    <td nowrap><%=dictLanguage.Item(Session("language")&"_opc_6")%></td>
                                <td><input type="password" name="billpass" id="billpass" value="" autocomplete="off" /></td>
								<td nowrap><%=dictLanguage.Item(Session("language")&"_opc_38")%></td>
                                <td><input type="password" name="billrepass" id="billrepass" value="" autocomplete="off" /></td>
							  </tr>
							  <tr>
							  	<td colspan="4"><hr /></td>
							  </tr>
							  <%end if%>
							  <%end if%>
                              <tr>
                                <td width="16%" nowrap><%=dictLanguage.Item(Session("language")&"_opc_10")%></td>
                                <td width="34%"><input type="text" name="billfname" id="billfname" value="<%=pcStrBillingFirstName%>" /></td>
                                <td width="16%" nowrap><%=dictLanguage.Item(Session("language")&"_opc_11")%></td>
                                <td width="34%">
                                	<input type="text" name="billlname" id="billlname" value="<%=pcStrBillingLastName%>" />
                                    <input type="hidden" name="billemail2" id="billemail2" value="<%=pcStrCustomerEmail%>" />
                              	</td>
                              </tr>
                              <tr>
							  	<%tmpShowE=0
								if scGuestCheckoutOpt=0 OR scGuestCheckoutOpt="" OR scGuestCheckoutOpt=1 then%>
                                <%if Not (Session("idCustomer")>0 AND session("CustomerGuest")="0") then
								tmpShowE=1%>
                                <td id="billArea1" nowrap><%=dictLanguage.Item(Session("language")&"_opc_5")%></td>
                                <td id="billArea2"><input type="text" name="billemail" id="billemail" value="<%=pcStrCustomerEmail%>" /></td>
                                <%end if%>
								<%end if%>
                                <td nowrap><%=dictLanguage.Item(Session("language")&"_opc_12")%></td>
                                <td <%if tmpShowE=0 then%>colspan="3"<%end if%>><input type="text" name="billcompany" id="billcompany" value="<%=pcStrBillingCompany%>" /></td>
                              </tr>
                              
                              <%
							  '//////////////////////////////////////////////////////////
							  '// START:  VAT
							  '//////////////////////////////////////////////////////////
							  If pcv_ShowVatId OR pcv_ShowSSN Then
                              %> 
                              <tr>                       
								  <% if pcv_ShowVatId then %>
                                        <td nowrap><%=dictLanguage.Item(Session("language")&"_Custmoda_26")%></td>
                                        <td>
                                            <input type="text" name="billVATID" value="<%=pcStrBillingVATID%>">
                                        </td>
                                  <% end if %>	
                                      
                                  <% if pcv_ShowSSN then %>
                                          <td nowrap><%=dictLanguage.Item(Session("language")&"_Custmoda_24")%></td>
                                          <td>
                                              <input type="text" name="billSSN" value="<%=pcStrBillingSSN%>">
                                          </td>
                                  <% end if %>
                              </tr>
                              <%
							  End If
							  '//////////////////////////////////////////////////////////
							  '// END:  VAT
							  '//////////////////////////////////////////////////////////
                              %>
                              
                              <tr>
                                <td nowrap><%=dictLanguage.Item(Session("language")&"_opc_13")%></td>
                                <td><input type="text" name="billaddr" id="billaddr" value="<%=pcStrBillingAddress%>" /></td>
                                <td nowrap><%=dictLanguage.Item(Session("language")&"_opc_14")%></td>
                                <td><input type="text" name="billaddr2" id="billaddr2" value="<%=pcStrBillingAddress2%>" /></td>
                              </tr>
                              <tr>
                                <td nowrap><%=dictLanguage.Item(Session("language")&"_opc_15")%></td>
                                <td><input type="text" name="billcity" id="billcity" value="<%=pcStrBillingCity%>" /></td>
                                <td nowrap><span id="billzipname"><%=dictLanguage.Item(Session("language")&"_opc_16")%></span></td>
                                <td><input type="text" name="billzip" id="billzip" value="<%=pcStrBillingPostalCode%>" />
								<script>
									function switchZipName1(tmpValue)
									{
										if (tmpValue=="CA")
										{
											$("#billzipname").html('<%=FixLang(dictLanguage.Item(Session("language")&"_opc_16a"))%>');
										}
										else
										{
											$("#billzipname").html('<%=FixLang(dictLanguage.Item(Session("language")&"_opc_16"))%>');
										}
									}
								</script>
									</td>
                              </tr>
                                <%
                                pcv_strTargetForm = "BillingForm" '// Name of Form
                                pcv_strCountryBox = "billcountry" '// Name of Country Dropdown
                                pcv_strTargetBox = "billstate" '// Name of State Dropdown
                                pcv_strProvinceBox =  "billprovince" '// Name of Province Field
								tmp_CountryBoxFunc="switchZipName1(this.value);"
                            
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
                            
							<% if ((pcv_NOShippingAtAll="1") AND (pcv_AlwAltShipAddress="1")) then %>
                                <tr>
                                    <td colspan="4">
										<%
                                        '// ProductCart v4.5 - Commercial vs. Residential
                                        Dim pcComResShipAddress
                                        if scComResShipAddress = "0" then
                                        %>
										<script>
                                    		HaveShipTypeArea=1;
                                   		</script>
                                    	<table class="pcShowContent" id="shipAddrTypeArea">
                                    		<tr>
                                    			<td colspan="4">
														<%=dictLanguage.Item(Session("language")&"_opc_23")%>&nbsp;<input type="radio" name="pcAddressType" value="1" checked>&nbsp;<%=dictLanguage.Item(Session("language")&"_opc_24")%>&nbsp;<input type="radio" name="pcAddressType" value="0">&nbsp;<%=dictLanguage.Item(Session("language")&"_opc_25")%>
                                                </td>
                                            </tr>
                                        </table>
										<%
                                        else
                                            Select Case scComResShipAddress
                                            Case "1"
                                                pcComResShipAddress="1"
                                            Case "2"
                                                pcComResShipAddress="0"
                                            Case "3"
                                                if session("customerType")="1" then
                                                    pcComResShipAddress="0"
                                                else
                                                    pcComResShipAddress="1"
                                                end if
                                            End Select
                                        %>
                                        <input type="hidden" name="pcAddressType" value="<%=pcComResShipAddress%>">
										<%
                                        end if
                                        %>
									</td>
								</tr>
							<% end if %>

                            <%'Start Special Customer Fields
                            tmpCustCFList=""
                            pcSFCustFieldsExist=""
        
                            query="SELECT pcCField_ID, pcCField_Name, pcCField_FieldType, pcCField_Value, pcCField_Length, pcCField_Maximum, pcCField_Required, pcCField_PricingCategories, pcCField_ShowOnReg, pcCField_ShowOnCheckout,'',pcCField_Description,0 FROM pcCustomerFields ORDER BY pcCField_Order ASC, pcCField_Name ASC;"
                            set rs=server.CreateObject("ADODB.RecordSet")
                            set rs=connTemp.execute(query)
                            if not rs.eof then
                                pcSFCustFieldsExist="YES"
                                tmpCustCFList=rs.GetRows()
                            end if
                            set rs=nothing
    
                            if pcSFCustFieldsExist="YES" AND Session("idCustomer")<>0 then
                            pcArr=tmpCustCFList
                            For k=0 to ubound(pcArr,2)
                                pcArr(10,k)=""
                                query="SELECT pcCFV_Value FROM pcCustomerFieldsValues WHERE idcustomer=" & Session("idCustomer") & " AND pcCField_ID=" & pcArr(0,k) & ";"
                                set rs=server.CreateObject("ADODB.RecordSet")
                                set rs=connTemp.execute(query)
                                if not rs.eof then
                                    pcArr(10,k)=rs("pcCFV_Value")
                                end if
                                set rs=nothing
                            Next
                            tmpCustCFList=pcArr
                            end if
    
                            if pcSFCustFieldsExist="YES" then
                                pcArr=tmpCustCFList
                                For k=0 to ubound(pcArr,2)						
                                    pcv_ShowField=0
                                    if pcArr(9,k)="1" then
                                        pcv_ShowField=1
                                    end if
                                    if (pcv_ShowField=1) AND (pcArr(7,k)="1") then
                                    if session("idCustomer")>"0" then
                                        query="SELECT pcCustFieldsPricingCats.idcustomerCategory FROM pcCustFieldsPricingCats INNER JOIN Customers ON (pcCustFieldsPricingCats.pcCField_ID=" & pcArr(0,k) & " AND pcCustFieldsPricingCats.idCustomerCategory=customers.idCustomerCategory) WHERE customers.idcustomer=" & session("idCustomer")
                                        set rs=Server.CreateObject("ADODB.Recordset")
                                        set rs=conntemp.execute(query)												
                                        if NOT rs.eof then
                                            pcv_ShowField=1
                                        else
                                            pcv_ShowField=0
                                        end if
                                        set rs=nothing
                                    else
                                        pcv_ShowField=0
                                    end if
                                    end if	
                                    pcArr(12,k)=pcv_ShowField
                                Next
                                tmpCustCFList=pcArr
                            end if
    
                            if pcSFCustFieldsExist="YES" then
                                pcArr=tmpCustCFList
    
                                For k=0 to ubound(pcArr,2)
                            
                                pcv_ShowField=pcArr(12,k)
                            
                                if pcv_ShowField=1 then 
                                %>
                                <tr>
                                    <td colspan="4">
                                        <table class="pcShowContent">
                                            <tr>
                                                <td width="30%"><%=pcArr(1,k)%>:</td>
                                                <td width="70%">
                                                    <%if pcArr(2,k)="1" then%>
                                                        <input type="checkbox" id="custfield<%=pcArr(0,k)%>" name="custfield<%=pcArr(0,k)%>" <%if pcArr(6,k)="1" then%>class="required"<%end if%> <%if pcArr(10,k)<>"" then%>value="<%=pcArr(10,k)%>"<%else%><%if pcArr(3,k)<>"" then%>value="<%=pcArr(3,k)%>"<%else%>value="1"<%end if%><%end if%> <%if pcArr(10,k)<>"" then%>checked<%end if%> class="clearBorder">
                                                    <%else%>
                                                        <input type="text" id="custfield<%=pcArr(0,k)%>" name="custfield<%=pcArr(0,k)%>" <%if pcArr(6,k)="1" then%>class="required"<%end if%> value="<%if session("idcustomer")=0 then%><%if pcArr(3,k)<>"" then%><%=pcArr(3,k)%><%end if%><%else%><%if pcArr(10,k)<>"" then%><%=pcArr(10,k)%><%else%><%=pcArr(3,k)%><%end if%><%end if%>" size="<%=pcArr(4,k)%>" <%if pcArr(5,k)>"0" then%>maxlength="<%=pcArr(5,k)%>"<%end if%>>
                                                    <%end if%>
                                              </td>
                                          </tr>
                                            <%if trim(pcArr(11,k))<>"" then%>
                                            <tr>
                                                <td></td>
                                                <td class="pcSmallText"><%=pcArr(11,k)%></td>
                                            </tr>
                                            <%end if%>
                                        </table>
                                  </td>
                              </tr>
                                <%
                                end if
                            Next
    
                            end if%>
                            
                            <% ' If the referrer drop down field is enabled, show it for a new customer
							if (Session("idCustomer")>"0") then
								query="SELECT IDRefer FROM Customers WHERE idCustomer=" & Session("idCustomer") & ";"
								set rs=Server.CreateObject("ADODB.Recordset")
                                set rs=connTemp.execute(query)
								if not rs.eof then
									Session("pcSFIDrefer")=rs("IDRefer")
								end if
								set rs=nothing
							end if
                            if ((Session("idCustomer")=0) OR ((Session("idCustomer")>"0") AND (Session("CustomerGuest")<>"0"))) AND (RefNewCheckout="1") then %>
                            <tr align="top"> 
                                <td colspan="4">
                                    <%=ReferLabel%>&nbsp;
                                    <select name="IDRefer" id="IDRefer" <%if ViewRefer="1" then%>class="required"<%end if%>>
                                        <option value="" <%if Session("pcSFIDrefer")="" then%>selected<%end if%>></option>
                                        <%
                                        query="select idrefer, [name] from Referrer where removed=0 order by SortOrder;"
                                        set rs=Server.CreateObject("ADODB.Recordset")
                                        set rs=connTemp.execute(query)
                                        do while not rs.eof
                                            intIdrefer=rs("idrefer")
                                            strName=rs("name") %>
                                            <option value="<%=intIdrefer%>" <%if Session("pcSFIDrefer")=trim(intIdrefer) then%>selected<%end if%>><%=strName%></option>
                                            <% rs.movenext
                                        loop
                                        set rs = nothing 
                                        %>
                                    </select>
                                </td>
                            </tr>
                        <% end if
                        'End If the referrer drop down field is enabled, show it for a new customer %>
                        
                        <% 'MAILUP-S: MailUp Lists, show it for new customer and when existing customers edit their account
						IF (session("SF_MU_Setup")="1" AND Session("idCustomer")<>0) OR (session("SF_MU_Setup")="1" AND Session("idCustomer")=0 AND ((NewsCheckout="1") OR (NewsReg="1"))) THEN
							call opendb()
							query="SELECT pcMailUpLists_ID,pcMailUpLists_ListID,pcMailUpLists_ListGuid,pcMailUpLists_ListName,pcMailUpLists_ListDesc,0 FROM pcMailUpLists WHERE pcMailUpLists_Active>0 and pcMailUpLists_Removed=0;"
							set rs=connTemp.execute(query)
							if not rs.eof then
								pcv_TurnMUOn=1%>
								<tr> 
								<td colspan="4">
								<div id="MailUpArea">
								<table>
								<%tmpArr=rs.getRows()
								set rs=nothing
								intCount=ubound(tmpArr,2)
								tmpNListChecked=0
								pcv_MUSynError=0
								if Session("idCustomer")<>0 then
								'Synchronizing
								For j=0 to intCount
									tmpResult=CheckUserStatus(Session("idCustomer"),pcStrCustomerEmail,tmpArr(1,j),tmpArr(2,j),session("SF_MU_URL"),session("SF_MU_Auto"))
									if tmpResult="-1" then
										pcv_MUSynError=1
										exit for
									else
										if tmpResult="2" then
											query="SELECT pcMailUpSubs_ID FROM pcMailUpSubs WHERE idCustomer=" & Session("idCustomer") & " AND pcMailUpLists_ID=" & tmpArr(0,j) & ";"
											set rs=connTemp.execute(query)
											dtTodaysDate=Date()
											if SQL_Format="1" then
												dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
											else
												dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
											end if
											if not rs.eof then
												if scDB="SQL" then
													query="UPDATE pcMailUpSubs SET idCustomer=" & Session("idCustomer") & ",pcMailUpLists_ID=" & tmpArr(0,j) & ",pcMailUpSubs_LastSave='" & dtTodaysDate & "',pcMailUpSubs_SyncNeeded=0,pcMailUpSubs_Optout=0 WHERE idCustomer=" & Session("idCustomer") & " AND pcMailUpLists_ID=" & tmpArr(0,j) & ";"
												else
													query="UPDATE pcMailUpSubs SET idCustomer=" & Session("idCustomer") & ",pcMailUpLists_ID=" & tmpArr(0,j) & ",pcMailUpSubs_LastSave=#" & dtTodaysDate & "#,pcMailUpSubs_SyncNeeded=0,pcMailUpSubs_Optout=0 WHERE idCustomer=" & Session("idCustomer") & " AND pcMailUpLists_ID=" & tmpArr(0,j) & ";"
												end if
											else
												if scDB="SQL" then
													query="INSERT INTO pcMailUpSubs (idCustomer,pcMailUpLists_ID,pcMailUpSubs_LastSave,pcMailUpSubs_SyncNeeded,pcMailUpSubs_Optout) VALUES (" & Session("idCustomer") & "," & tmpArr(0,j) & ",'" & dtTodaysDate & "',0,0);"
												else
													query="INSERT INTO pcMailUpSubs (idCustomer,pcMailUpLists_ID,pcMailUpSubs_LastSave,pcMailUpSubs_SyncNeeded,pcMailUpSubs_Optout) VALUES (" & Session("idCustomer") & "," & tmpArr(0,j) & ",#" & dtTodaysDate & "#,0,0);"
												end if
											end if
											set rs=nothing
											set rs=connTemp.execute(query)
											set rs=nothing
										end if
										if tmpResult="4" then
											tmpArr(5,j)=4
										end if
										if tmpResult="1" or tmpResult="3" then
											query="DELETE FROM pcMailUpSubs WHERE idCustomer=" & Session("idCustomer") & " AND pcMailUpLists_ID=" & tmpArr(0,j) & ";"
											set rs=connTemp.execute(query)
											set rs=nothing
										end if
									end if
								Next
								For j=0 to intCount
									query="SELECT idcustomer FROM pcMailUpSubs WHERE idcustomer=" & Session("idCustomer") & " AND pcMailUpLists_ID=" & tmpArr(0,j) & " AND pcMailUpSubs_Optout=0;"
									set rs=connTemp.execute(query)
									tmpOptedIn=0
									if not rs.eof then
										tmpOptedIn=1
										tmpNListChecked=1
									end if
									set rs=nothing
									if tmpArr(5,j)<>4 then
										tmpArr(5,j)=tmpOptedIn
									end if
								Next
								end if%>
								<%if pcv_MUSynError=1 then%>
								<tr> 
									<td colspan="4">
										<div class="pcErrorMessage">
											<%response.write dictLanguage.Item(Session("language")&"_MailUp_SynNote1")%>
										</div>
									</td>
								</tr>
								<%end if%>
								<tr> 
									<td colspan="4" class="pcSpacer"><script>tmpNListChecked=<%=tmpNListChecked%>;</script><input type="hidden" name="newslistcount" value="<%=intCount%>"></td>
								</tr>
								<tr> 
									<td colspan="4"><%response.write dictLanguage.Item(Session("language")&"_MailUp_RegisterLabel")%></td>
								</tr>
								<%For j=0 to intCount%>
								<tr> 
									<td align="right" valign="top"><input type="checkbox" onclick="javascript: tmpNListChecked=1;" value="<%=tmpArr(0,j)%>" name="newslist<%=j%>" <%if tmpArr(5,j)="1" or tmpArr(5,j)="4" or (Session("idCustomer")=0 AND Session("pcSFpcNewsList" & j)&""=tmpArr(0,j)&"") then%>checked<%end if%> class="clearBorder" /></td>
									<td valign="top" colspan="3">
										<b><%=tmpArr(3,j)%></b><%if tmpArr(5,j)="4" then%>&nbsp;(<span class="pcTextMessage"><%response.write dictLanguage.Item(Session("language")&"_MailUp_RegisterLabel1")%></span>)<%end if%><%if tmpArr(4,j)<>"" then%><br><%=tmpArr(4,j)%><%end if%>
										<%if tmpArr(5,j)="4" then%><div class="pcErrorMessage"><%response.write dictLanguage.Item(Session("language")&"_MailUp_RegisterLabel2a")%><a href="javascript:newWindow('mu_subscribe.asp?listid=<%=tmpArr(1,j)%>','window1');"><%response.write dictLanguage.Item(Session("language")&"_MailUp_RegisterLabel2b")%></a><%response.write dictLanguage.Item(Session("language")&"_MailUp_RegisterLabel2c")%></div><%end if%>
									</td>
								</tr>
								<%Next%>
								</table>
								</div>
								</td>
								</tr>
							<%end if
							set rs=nothing	
							'End If MailUp Lists
						ELSE%>
							<% 'If newsletter is enabled, show it for new customer and when existing customers edit their account
							if (session("SF_MU_Setup")<>"1") AND (AllowNews="1") AND (NewsCheckout="1") then%>
								<tr align="top"> 
									<td colspan="4">
										<input type="checkbox" value="1" name="CRecvNews" <%if pcIntRecvNews="1" then%>checked<%end if%> class="clearBorder" />&nbsp;<%=NewsLabel%>
									</td>
								</tr>
							<% end if
							'End If newsletter is enabled, show it for new customer
						END IF
						'MAILUP-E %>
  
						<%
                        '=========================================
                        ' TERMS Area : NOT A PANEL : START
                        ' Hide after it has been agreed to
                        '=========================================
      
                        pcv_AgreedToTerms=0
      
                        if scTermsShown=1 then
                            pcv_AgreedToTerms=1
                        else
                            If (pcAgreeTerms="0" OR pcAgreeTerms="") then
                                pcv_AgreedToTerms=1
                            End If
                        End if
      
                        if scTerms=1 AND pcv_AgreedToTerms=1 then
							%>
                             <tr>
                                <td colspan="4">
								  <script>
                                        var pcCustomerTermsAgreed=0;
                                    </script>
                                    <%
                                    Session("pcCustomerTermsAgreed")="0"
                        
                                    query="SELECT pcStoreSettings_TermsLabel, pcStoreSettings_TermsCopy FROM pcStoreSettings;"
                                    set rs=server.CreateObject("ADODB.RecordSet")
                                    set rs=conntemp.execute(query)
                    
                                    pcStrTermsLabel=rs("pcStoreSettings_TermsLabel")
									
                                    pcStrTermsCopy=rs("pcStoreSettings_TermsCopy")
									if trim(pcStrTermsCopy)<>"" and not isNull(pcStrTermsCopy) then
										pcStrTermsCopy=replace(pcStrTermsCopy, CHR(10),"<br>")
										pcStrTermsCopy=replace(pcStrTermsCopy, "&lt;","<")
										pcStrTermsCopy=replace(pcStrTermsCopy, "&gt;",">")
									end if
									
                                    set rs=nothing
                                    %>                                    
                                    <div id="AgreeArea" name="AgreeArea">
                                    	<input type="checkbox" value="1" id="AgreeTerms" name="AgreeTerms" class="clearBorder" />&nbsp;<%=pcStrTermsLabel%>. <a href="javascript:;" id="ViewTerms"><%=dictLanguage.Item(Session("language")&"_opc_50")%></a>.
                                        <div id="TermsDialog" title="<%=pcStrTermsLabel%>" style="display:none">
                                        	<div id="TermsMsg"><%=pcStrTermsCopy%></div>
                                        </div>
                                    </div>
                                </td>
                             </tr>
                      <% else %>                        
                            <script>
                                var pcCustomerTermsAgreed=1;
                            </script>                          
                          <%
                          Session("pcCustomerTermsAgreed")="1"					
                      end if                  
                      %>    
                             
					  <%
                      'SB S
                      if pcIsSubscription Then
                          
						  '// Find if we have required agreements
						  Dim pcv_strRegAgree
						  pcv_strRegAgree=0
                          for f=1 to pcCartIndex
                              pcSubscriptionId = pcCartArray(f,38)
                              if pcSubscriptionId<>"0" then
							  	  '// Check Package Level
                                  query= "SELECT SB_Agree, SB_AgreeText FROM SB_Packages WHERE SB_PackageID ="&pcSubscriptionId&" AND  SB_Agree=1" 
                                  set rss=server.CreateObject("ADODB.RecordSet")
                                  set rss=connTemp.execute(query)	
                                  If not rss.eof Then
                                      session("pcCartSession")=pcCartArray		
                                      session("pcIsRegAgree") = true
									  pSubAgreeText=rss("SB_AgreeText")
									  pcv_strRegAgree=1								 
								  Else
										'// Check Global
										if scSBRegAgree="1" then
                                      		session("pcCartSession")=pcCartArray		
                                      		session("pcIsRegAgree") = true
											pSubAgreeText=scSBAgreeText
											pcv_strRegAgree=1
										end if
                                  End if 
                              end if
                          Next
                          set  rss = nothing
                      End if
					  
					  If pcv_strRegAgree=1 Then
						%>
                             <tr>
                                <td colspan="4">
								  <script>
                                        var pcCustomerRegAgreed=0;
                                    </script>
                                    <%
                                    Session("pcCustomerRegAgreed")="0"

                                    pcStrTermsLabel_SB=scSBLang1

									if trim(pSubAgreeText)<>"" and not isNull(pSubAgreeText) then
										pSubAgreeText=replace(pSubAgreeText, CHR(10),"<br>")
										pSubAgreeText=replace(pSubAgreeText, "&lt;","<")
										pSubAgreeText=replace(pSubAgreeText, "&gt;",">")
									end if
                                    %>                                    
                                    <div id="sb_AgreeArea" name="sb_AgreeArea">
                                    	<div style="padding:2px">
                                        	<input type="checkbox" value="1" id="sb_AgreeTerms" name="sb_AgreeTerms" class="clearBorder" />
                                            &nbsp;<%=pcStrTermsLabel_SB%>. 
                                            <a href="javascript:;" id="sb_ViewTerms"><%=scSBLang4%></a>.
                                         </div>
                                        <div id="sb_TermsDialog" title="<%=pcStrTermsLabel_SB%>" style="display:none">
                                        	<div id="TermsMsg"><%=pSubAgreeText%></div>
                                        </div>
                                    </div>
                                </td>
                             </tr>
                      <% else %>                        
                            <script>
                                var pcCustomerRegAgreed=1;
                            </script> 
                           	<%
							Session("pcCustomerRegAgreed")="1"				
                      end if  				 			
                      'SB E            
                      %>   
                             
                             <tr>
                                <td colspan="4">
                                    <div id="BillingLoader" style="display:none"></div>
                                    <% 'SB S %>
                                    <div id="BillingLoaderSB" style="display:none"></div>
                                    <% 'SB E %>
                                </td>
                             </tr>
                              <tr>
                                <td colspan="4" style="padding-top: 10px;"><input type="image" name="BillingSubmit" id="BillingSubmit" src="<%=RSlayout("pcLO_Update")%>" align="absmiddle" border="0"><% if (session("idCustomer")="0" OR session("idCustomer")="") then %>&nbsp;&nbsp;<input type="image" name="BillingCancel" id="BillingCancel" src="<%=RSlayout("back")%>" align="absmiddle" border="0"><% end if %></td>
                              </tr>
                            </table>
                            </form>
						</div>
                        <%
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						'// Billing Address - End
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
						%>

                        <%
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						'// Shipping Address - Start
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
						%>

						<div id="ShipLoadContentMsg" style="display: none;"></div>
                        <div id="ShippingArea" style="display: none;">
                            <% if (pcv_NOShippingAtAll = "2" AND pcv_AlwAltShipAddress="0") OR (pcv_AlwAltShipAddress="1") then %>
                            	
								<script>
                                	var HaveShipArea=0;
                            	</script>
                            
							<% Else %>
                            
                                <div class="pcCheckoutSubTitle"><%=dictLanguage.Item(Session("language")&"_opc_19")%></div>
                       	  		<script>
                                	var HaveShipArea=1;
                            	</script>
                               
                            	<form name="ShippingForm" id="ShippingForm">
                            		<div style="padding: 8px 0 4px 3px;">
                                        
                                         <div id="radios">
                                         	<%if (session("Cust_IDEvent")="") OR (session("Cust_IDEvent")<>"" AND gDelivery=0) then%>
                                            	<input id="rad_0" type="radio" name="ShipArrOpts" value="-1" onclick="javascript:FillShipForm(this.value,0);" /> 
												<label for="rad_0"><%=dictLanguage.Item(Session("language")&"_opc_20")%></label>
 												<br />
                                            <%end if%>
                                          	<%if session("Cust_IDEvent")<>"" then%>  
                                            	<input id="rad_1" type="radio" name="ShipArrOpts" value="-2" onclick="javascript:FillShipForm(this.value,0);" /> 
												<label for="rad_1"><%=dictLanguage.Item(Session("language")&"_opc_21")%></label> 
                                            	<br />
                                            <%end if%>                                             
                                         </div>

										 <script>											
											var ShipContents="";
											<%if pcv_AlwAltShipAddress="0" OR pcv_AlwAltShipAddress="2" then%>
												<%if (session("Cust_IDEvent")="") OR (session("Cust_IDEvent")<>"" AND gDelivery=0) then%>
													var NeedLoadShipContent=1;
													var CanCreateNewShip=1;
												<%else%>
													var NeedLoadShipContent=0;
													var CanCreateNewShip=0;
												<%end if%>
											<%else%>
												var NeedLoadShipContent=0;
												var CanCreateNewShip=0;
											<%end if%>
                                          </script>
                                    </div>
                            
                                    <table class="pcShowContent" id="shippingAddressArea" style="display: none;">
									  <tr>
									  <td colspan="4">
									  <a href="javascript:copyfromBillAddr();"><%=dictLanguage.Item(Session("language")&"_opc_msg2")%></a>
									  <script>
									  function copyfromBillAddr()
									  {
										//$("#shipnickname").val("");
										$("#shipfname").val($("#billfname").val());
										$("#shiplname").val($("#billlname").val());
										if ($("#billemail").length) {
											$("#shipemail").val($("#billemail").val());
										} else {
											$("#shipemail").val($("#billemail2").val());
										}								
										$("#shipphone").val($("#billphone").val());
										$("#shipfax").val($("#billfax").val());
										$("#shipcompany").val($("#billcompany").val());
										$("#shipaddr").val($("#billaddr").val());
										$("#shipaddr2").val($("#billaddr2").val());
										$("#shipcity").val($("#billcity").val());
										$("#shipprovince").val($("#billprovince").val());
										$("#shipstate").val($("#billstate").val());
										$("#shipzip").val($("#billzip").val());
										$("#shipcountry").val($("#billcountry").val());
										SwitchStates('ShippingForm',document.ShippingForm.shipcountry.options.selectedIndex, 'shipcountry', 'shipstate', 'shipprovince', $("#billstate").val(), '');
									  }
									  </script>
									  </td>
                                      <tr id="shipnicknameArea">
                                        <td width="20%" nowrap><%=dictLanguage.Item(Session("language")&"_opc_22")%></td>
                                        <td colspan="3"><input type="text" name="shipnickname" id="shipnickname" /></td>
                                      </tr>
                                      <tr id="shipnameArea">
                                        <td width="20%" nowrap><%=dictLanguage.Item(Session("language")&"_opc_10")%></td>
                                        <td width="30%"><input type="text" name="shipfname" id="shipfname" /></td>
                                        <td width="20%" nowrap><%=dictLanguage.Item(Session("language")&"_opc_11")%></td>
                                        <td width="30%"><input type="text" name="shiplname" id="shiplname" /></td>
                                      </tr>
                                      <tr>
                                        <td nowrap><%=dictLanguage.Item(Session("language")&"_opc_5")%></td>
                                        <td><input type="text" name="shipemail" id="shipemail" /></td>
                                        <td width="20%" nowrap><%=dictLanguage.Item(Session("language")&"_opc_12")%></td>
                                        <td width="30%"><input type="text" name="shipcompany" id="shipcompany" /></td>
                                      </tr>
                                      <tr>
                                        <td nowrap><%=dictLanguage.Item(Session("language")&"_opc_13")%></td>
                                        <td><input type="text" name="shipaddr" id="shipaddr" /></td>
                                        <td nowrap><%=dictLanguage.Item(Session("language")&"_opc_14")%></td>
                                        <td><input type="text" name="shipaddr2" id="shipaddr2" /></td>
                                      </tr>
                                      <tr>
                                        <td nowrap><%=dictLanguage.Item(Session("language")&"_opc_15")%></td>
                                        <td><input type="text" name="shipcity" id="shipcity" /></td>
                                        <td nowrap><span id="shipzipname"><%=dictLanguage.Item(Session("language")&"_opc_16")%></span></td>
                                        <td><input type="text" name="shipzip" id="shipzip" />
										<script>
										function switchZipName2(tmpValue)
										{
											if (tmpValue=="CA")
											{
												$("#shipzipname").html('<%=FixLang(dictLanguage.Item(Session("language")&"_opc_16a"))%>');
											}
											else
											{
												$("#shipzipname").html('<%=FixLang(dictLanguage.Item(Session("language")&"_opc_16"))%>');
											}
										}
										</script>
										</td>
                                      </tr>
                                        <%
                                        pcv_strTargetForm = "ShippingForm" '// Name of Form
                                        pcv_strCountryBox = "shipcountry" '// Name of Country Dropdown
                                        pcv_strTargetBox = "shipstate" '// Name of State Dropdown
                                        pcv_strProvinceBox =  "shipprovince" '// Name of Province Field
										tmp_CountryBoxFunc="switchZipName2(this.value);"
                                    
                                        '// Set local Country to Session
                                        if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
                                            Session(pcv_strSessionPrefix&pcv_strCountryBox) = ""
                                        end if
                                    
                                        '// Set local State to Session
                                        if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
                                            Session(pcv_strSessionPrefix&pcv_strTargetBox) = ""
                                        end if
                                    
                                        '// Set local Province to Session
                                        if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
                                            Session(pcv_strSessionPrefix&pcv_strProvinceBox) =  ""
                                        end if
            
                                        pcs_CountryDropdown
                                        %>
            
                                        <%
                                        '// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
                                        pcs_StateProvince
                                        %>
                                      <tr>
                                        <td nowrap><%=dictLanguage.Item(Session("language")&"_opc_17")%></td>
                                        <td><input type="text" name="shipphone" id="shipphone" /></td>
                                        <td nowrap><%=dictLanguage.Item(Session("language")&"_opc_18")%></td>
                                        <td><input type="text" name="shipfax" id="shipfax" /></td>
                                      </tr>
                                    </table>
                            		<% if (pcv_NOShippingAtAll="1") OR (pcv_NOShippingAtAll="2" AND pcv_AlwAltShipAddress="2") then %>
										<%
                                        '// ProductCart v4.5 - Commercial vs. Residential
                                        if scComResShipAddress = "0" then
                                        %>
										<script>
                                        	HaveShipTypeArea=1;
                                        </script>
                                        <table class="pcShowContent" id="shipAddrTypeArea" style="display:none;">
                                        	<tr>
                                            	<td colspan="4">
													<%=dictLanguage.Item(Session("language")&"_opc_23")%>&nbsp;<input type="radio" name="pcAddressType" value="1" checked>&nbsp;<%=dictLanguage.Item(Session("language")&"_opc_24")%>&nbsp;<input type="radio" name="pcAddressType" value="0">&nbsp;<%=dictLanguage.Item(Session("language")&"_opc_25")%>
                                                </td>
                                          	</tr>
                                    	</table>
										<%
                                        else
                                            Select Case scComResShipAddress
                                            Case "1"
                                                pcComResShipAddress="1"
                                            Case "2"
                                                pcComResShipAddress="0"
                                            Case "3"
                                                if session("customerType")="1" then
                                                    pcComResShipAddress="0"
                                                else
                                                    pcComResShipAddress="1"
                                                end if
                                            End Select
                                        %>
                                        <input type="hidden" name="pcAddressType" value="<%=pcComResShipAddress%>">
										<%
                                        end if
                                        %>
                            		<% end if %>
                            
									<%if DFShow="1" OR TFShow="1" then%>
                       			  		<script>
                            				HaveDeliveryArea=1;
                            			</script>
                            			<table class="pcShowContent" id="shipDeliveryArea" style="display:none;">
                            			<% if DFShow="1" then 'show delivery date field %>
                                			<script language="javascript">
                                				function CalPop(sInputName) {
                                    			window.open('../Calendar/Calendar.asp?N=' + escape(sInputName) + '&DT=' + escape(window.eval(sInputName).value), 'CalPop','toolbar=0,width=378,height=225' );
                                				}
                                			</script>			
                                			<tr> 
                                    			<td width="20%" nowrap><%=DFLabel%>:</td>
                                    			<td colspan="3">
                                                    <a href="javascript:CalPop('document.ShippingForm.DF1');"><img SRC="../Calendar/icon_Cal.gif" border="0"></a>  
                                                    <input type="text" readonly size="10" name="DF1" value="<%if session("DF1")<>"" then%><%=showdateFrmt(session("DF1"))%><%else%><%end if%>" <%if DFReq="1" then%>class="required"<%end if%>>
                                                </td>
                                            </tr>
                            				<% ' If the store is using blackout dates, show a message here and a link a list of dates
                            				Dim blackoutdates
                            				query="SELECT * FROM Blackout ORDER BY Blackout_Date ASC;"
                            				set rs=connTemp.execute(query)
											If rs.eof Then
												blackoutdates="0"
											else
												blackoutdates="1"
											end if
											set rs = nothing
                           					if blackoutdates="1" then 
											 	%>
										  <script language="JavaScript">
                                                <!--
                                                function optwin2(fileName)
                                                    {
                                                    myFloater = window.open('','myWindow','scrollbars=yes,status=no,width=400,height=300')
                                                    myFloater.location.href = fileName;
                                                    }	
                                                //-->
                                                </script>  
        
                                              <tr> 
                                                  <td width="20%">&nbsp;</td>
                                                  <td colspan="3">
                                                  <a href="javascript:optwin2('blackoutDates.asp')"><%response.write dictLanguage.Item(Session("language")&"_catering_20")%></a><%response.write dictLanguage.Item(Session("language")&"_catering_21")%></td>
                                              </tr>
                                          <% end if
									end if
									
									if TFShow="1" then %>			
                                        <tr> 
                                            <td width="20%" nowrap><%=TFLabel%>:</td>
                                            <td colspan="3"> 
                                            <select name="TF1" <%if TFReq="1" then%>class="required"<%end if%>>
                                                <% if pcSFTF1="" then %>
                                                    <option value=""><%response.write dictLanguage.Item(Session("language")&"_viewCatOrder_6")%></option>
                                                <% end if %>
                                                <% if scDateFrmt="DD/MM/YY" then %>
                                                    <option value="7:00">7:00</option>
                                                    <option value="7:30">7:30</option>
                                                    <option value="8:00">8:00</option>
                                                    <option value="8:30">8:30</option>
                                                    <option value="9:00">9:00</option>
                                                    <option value="9:30">9:30</option>
                                                    <option value="10:00">10:00</option>
                                                    <option value="10:30">10:30</option>
                                                    <option value="11:00">11:00</option>
                                                    <option value="11:30">11:30</option>
                                                    <option value="12:00">12:00</option>
                                                    <option value="12:30">12:30</option>
                                                    <option value="13:00">13:00</option>
                                                    <option value="13:30">13:30</option>
                                                    <option value="14:00">14:00</option>
                                                    <option value="14:30">14:30</option>
                                                    <option value="15:00">15:00</option>
                                                    <option value="15:30">15:30</option>
                                                    <option value="16:00">16:00</option>
                                                    <option value="16:30">16:30</option>
                                                    <option value="17:00">17:00</option>
                                                    <option value="17:30">17:30</option>
                                                    <option value="18:00">18:00</option>								
                                                    <option value="18:30">18:30</option>								
                                                    <option value="19:00">19:00</option>								
                                                    <option value="19:30">19:30</option>								
                                                    <option value="20:00">20:00</option>								
                                                    <option value="20:30">20:30</option>
                                                    <option value="21:00">21:00</option>													
                                                <% Else  %>
                                                    <option value="7:00 AM">7:00 AM</option>
                                                    <option value="7:30 AM">7:30 AM</option>
                                                    <option value="8:00 AM">8:00 AM</option>
                                                    <option value="8:30 AM">8:30 AM</option>
                                                    <option value="9:00 AM">9:00 AM</option>
                                                    <option value="9:30 AM">9:30 AM</option>
                                                    <option value="10:00 AM">10:00 AM</option>
                                                    <option value="10:30 AM">10:30 AM</option>
                                                    <option value="11:00 AM">11:00 AM</option>
                                                    <option value="11:30 AM">11:30 AM</option>
                                                    <option value="12:00 PM">12:00 PM</option>
                                                    <option value="12:30 PM">12:30 PM</option>
                                                    <option value="1:00 PM">1:00 PM</option>
                                                    <option value="1:30 PM">1:30 PM</option>
                                                    <option value="2:00 PM">2:00 PM</option>
                                                    <option value="2:30 PM">2:30 PM</option>
                                                    <option value="3:00 PM">3:00 PM</option>
                                                    <option value="3:30 PM">3:30 PM</option>
                                                    <option value="4:00 PM">4:00 PM</option>
                                                    <option value="4:30 PM">4:30 PM</option>
                                                    <option value="5:00 PM">5:00 PM</option>
                                                    <option value="5:30 PM">5:30 PM</option>
                                                    <option value="6:00 PM">6:00 PM</option>
                                                    <option value="7:00 PM">7:00 PM</option>
                                                    <option value="7:30 PM">7:30 PM</option>
                                                    <option value="8:00 PM">8:00 PM</option>
                                                    <option value="8:30 PM">8:30 PM</option>
                                                    <option value="9:00 PM">9:00 PM</option>
                                                <% End If %>
                                            </select>
                                            </td>
                                        </tr>
                                    <% end if  '// if TFShow="1" then %>
									<% if (DTCheck="1") then %>
                                        <tr> 
                                            <td width="20%" nowrap>&nbsp;</td>
                                            <td colspan="3"><i><%response.write dictLanguage.Item(Session("language")&"_catering_6")%></i></td>
                                        </tr>
                                    <% end if%>
                            	</table>
                            <%end if%>
                            
                            <table class="pcShowContent">
                            	<tr>
                                	<td colspan="4">
                                    	<div id="ShippingLoader" style="display:none"></div>
                                	</td>
                             	</tr>
                            </table>
                            <div style="padding-top: 10px;"><input type="image" name="ShippingSubmit" id="ShippingSubmit" src="<%=RSlayout("pcLO_Update")%>" border="0"></div>
                            </form>
                            <% End If '// If NOT pcv_AlwAltShipAddress = "-1" Then %>
                        </div>
                        <%
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						'// Shipping Address - End
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
						%>


                    </div>
                </div>
                <%
				'/////////////////////////////////////////////////////////////////////////
				'// ADDRESS:  END
				'/////////////////////////////////////////////////////////////////////////
				%>



                <%
				'/////////////////////////////////////////////////////////////////////////
				'// SHIPPING INFORMATION:  START
				'/////////////////////////////////////////////////////////////////////////
				if Not (pcv_NOShippingAtAll = "2") then
				%>  
                <a name="opcShippingAnchor"></a>            
				<div class="AccordionPanel" id="opcShipping">
                    <div class="AccordionPanelTab">
                        <div class="StatusIndicators">
                        	<a id="btnEditShip" href="javascript:;" onclick="javascript: acc1.openPanel('opcShipping'); " alt="<%=dictLanguage.Item(Session("language")&"_opc_45")%>" title="<%=dictLanguage.Item(Session("language")&"_opc_45")%>" style="display:none;">Edit</a>
                        	<img id="btnOKShip" src="images/pc_checkmark_sm.gif" alt="<%=dictLanguage.Item(Session("language")&"_opc_46")%>" title="<%=dictLanguage.Item(Session("language")&"_opc_46")%>" <%If (session("ExpressCheckoutPayment")="YES") OR (NOT (session("idCustomer")>0 and session("CustomerGuest")="0")) then%>style="display:none"<%end if%>>
                            <img id="btnErrorShip" src="images/pc_icon_error.gif" width="18" height="18" border="0" alt="<%=dictLanguage.Item(Session("language")&"_opc_47")%>" title="<%=dictLanguage.Item(Session("language")&"_opc_47")%>" style="display:none;">                        
                        </div>
 						<span class="pcCheckoutTitle"><%=dictLanguage.Item(Session("language")&"_opc_27")%></span>
                    	<div id="ShippingMethod"></div>
                    </div>
                    <div class="AccordionPanelContent">		
						<div id="ShipChargeLoadContentMsg" style="display: none;"></div>
						<div id="ShippingChargeArea" style="display: none;"></div>
                    </div>
                </div>
                <%
				end if
				'/////////////////////////////////////////////////////////////////////////
				'// SHIPPING INFORMATION:  END
				'/////////////////////////////////////////////////////////////////////////
				%>   



                
                <%
				'/////////////////////////////////////////////////////////////////////////
				'// PAYMENT INFORMATION:  START
				'/////////////////////////////////////////////////////////////////////////
				%>  
                <a name="opcPaymentAnchor"></a>
                <div class="AccordionPanel" id="opcPayment">
                  <div class="AccordionPanelTab">
                        <div class="StatusIndicators">
                        	<a id="btnEditPay" href="javascript:;" onclick="javascript: acc1.openPanel('opcShipping') " alt="<%=dictLanguage.Item(Session("language")&"_opc_45")%>" title="<%=dictLanguage.Item(Session("language")&"_opc_45")%>" style="display:none;">Edit</a>
                        	<img id="btnOKPay" src="images/pc_checkmark_sm.gif" alt="<%=dictLanguage.Item(Session("language")&"_opc_46")%>" title="<%=dictLanguage.Item(Session("language")&"_opc_46")%>" <%If (session("ExpressCheckoutPayment")="YES") OR (NOT (session("idCustomer")>0 and session("CustomerGuest")="0")) then%>style="display:none"<%end if%>>
                            <img id="btnErrorPay" src="images/pc_icon_error.gif" width="18" height="18" border="0" alt="<%=dictLanguage.Item(Session("language")&"_opc_47")%>" title="<%=dictLanguage.Item(Session("language")&"_opc_47")%>" style="display:none;">                        
                        </div>                    	
                    	<span class="pcCheckoutTitle"><%=dictLanguage.Item(Session("language")&"_opc_28")%></span>
                        <%' IF RewardsActive=1 THEN response.write "&amp; " & RewardsLabel %>
                  </div>
                    <div class="AccordionPanelContent">
						<div id="TaxLoadContentMsg" style="display: none;"></div>
						<div id="TaxContentArea" style="display: none;"></div>
                        <div id="PaymentContentArea">
						<%
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						'// Password - Start
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
						%>
						<% if scGuestCheckoutOpt=0 OR scGuestCheckoutOpt="" OR scGuestCheckoutOpt=1 then
						if NOT (session("idCustomer")>"0" AND (session("CustomerGuest")="0" OR session("CustomerGuest")="2")) then %>
                            <div id="PwdArea" style="display: none;">
                                <form id="PwdForm" name="PwdForm">
                                <table class="pcShowContent">
                                <tr>
                                    <td colspan="4"><%=dictLanguage.Item(Session("language")&"_opc_common_3")%></td>
                                </tr>
                                <tr>
                                    <td width="20%" nowrap><%=dictLanguage.Item(Session("language")&"_opc_6")%></td>
                                    <td width="30%"><input type="password" name="newPass1" id="newPass1" size="20" autocomplete="off"></td>
                                    <td width="20%" nowrap><%=dictLanguage.Item(Session("language")&"_opc_38")%></td>
                                    <td width="30%"><input type="password" name="newPass2" id="newPass2" size="20" autocomplete="off"></td>
                                </tr>
                                <tr>
                                    <td colspan="4" style="padding-top: 10px;"><div id="PwdLoader"></div></td>
                                </tr>
                                <tr>
                                    <td colspan="4" style="padding-top: 10px;">
                                    	<input type="image" name="PwdSubmit" id="PwdSubmit" src="<%=RSlayout("pcLO_Update")%>" align="absmiddle" border="0">
										<%if scGuestCheckoutOpt=0 OR scGuestCheckoutOpt="" then%>
                                        &nbsp;&nbsp;
                                        <a href="javascript:;" onclick="$('#PwdLoader').hide(); $('#PwdWarning').hide(); $('#PwdArea').hide();"><%=dictLanguage.Item(Session("language")&"_opc_51")%></a>
										<%end if%>
                                    </td>
                                </tr>
                                </table>
                                </form>
                            </div>
						<% end if
						end if %>
                        <%
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						'// Password - End
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
						%>


                        <%
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						'// Other - Start
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
						%>
						<div id="OtherArea">
							<form id="OtherForm" name="OtherForm">
							<table class="pcShowContent">
								<%
                                '// Nickname - S
                                %>
                                <% if scOrderName="1" then 'Allow customers to nickname their order %>
                                <tr>
                                    <td colspan="4">
                                    	<%
										If pcv_strPayPanel = "1" Then
										
											savOrderNickName = session("pord_OrderName")
											
											query="SELECT pcCustSession_OrderName FROM pcCustomerSessions WHERE pcCustomerSessions.idDbSession="&session("pcSFIdDbSession")&" AND pcCustomerSessions.randomKey="&session("pcSFRandomKey")&" AND pcCustomerSessions.idCustomer="&session("idCustomer")&";"
											set rs=server.CreateObject("ADODB.RecordSet")
											set rs=conntemp.execute(query)
											if NOT rs.eof then
												savOrderNickName=rs("pcCustSession_OrderName")
											end if									
											set rs=nothing
										
										End If
										%>
                                    	<div>
                                        	<a href="javascript:;" onclick="togglediv('nickname');"><img src="images/edit.gif" alt="" border="0"></a>
											<a href="javascript:;" onclick="togglediv('nickname');"><%response.write dictLanguage.Item(Session("language")&"_catering_13b")%></a>
                                        </div>
                                    	<table class="pcShowContent" id="nickname" style="display:none; border: 1px #eee solid; margin-bottom: 10px;">
                                            <tr> 
                                                <td colspan="4" class="pcSpacer"></td>
                                            </tr>
                                            <tr>
                                                <td colspan="4"><%response.write dictLanguage.Item(Session("language")&"_catering_1")%></td>
                                            </tr>					
                                            <tr> 
                                                <td><%response.write dictLanguage.Item(Session("language")&"_catering_12")%></td>
                                                <td colspan="3" align="left"> 
                                                    <input type="text" name="OrderNickName" value="<%=savOrderNickName%>" size="20">
                                                </td>
                                            </tr>
                                       	</table>
                                        
                                    </td>
                                </tr>  
                                <% end if 'End allow customers to nickname their order %>
                                <%
                                '// Nickname - E
                                %>
                                
                                <%
                                '// Comments - S
                                %>
                                <tr>
                                    <td colspan="4">
                                    	<%
										If pcv_strPayPanel = "1" Then
										
											savOrderComments = ""
											
											query="SELECT pcCustSession_Comment FROM pcCustomerSessions WHERE pcCustomerSessions.idDbSession="&session("pcSFIdDbSession")&" AND pcCustomerSessions.randomKey="&session("pcSFRandomKey")&" AND pcCustomerSessions.idCustomer="&session("idCustomer")&";"
											set rs=server.CreateObject("ADODB.RecordSet")
											set rs=conntemp.execute(query)
											if NOT rs.eof then
												savOrderComments=rs("pcCustSession_Comment")
											end if									
											set rs=nothing
										
										End If
										%>
                                   		<div>
                                        	<a href="javascript:;" onclick="togglediv('comments');"><img src="images/edit.gif" alt="" border="0"></a>
											<a href="javascript:;" onclick="togglediv('comments');"><%response.write dictLanguage.Item(Session("language")&"_order_Ub")%></a>
                                        </div>
                                    	<table class="pcShowContent" id="comments" style="display:none; border: 1px #eee solid; margin-bottom: 10px;">
                                            <tr> 
                                                <td class="pcSpacer"></td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <textarea name="OrderComments" cols="50" rows="3"><%=savOrderComments%></textarea>
                                                </td>
                                            </tr>                                
                                       	</table>
                                        
                                    </td>
                                </tr>  
                                <%
                                '// Comments - E
                                %>
                            
								<%
                                '// Gift Certs - S
                                %>
								<% HaveGcsTest=0    
                                For f=1 to pcCartIndex
                                    if pcCartArray(f,10)="0" then
                                        query="SELECT pcprod_Gc FROM Products WHERE idproduct=" & pcCartArray(f,0) & " AND pcprod_Gc=1;"
                                        set rsGc=conntemp.execute(query)
                                        if not rsGc.eof then
                                            HaveGcsTest=1
                                            set rsGc=nothing
                                            exit for
                                        end if
                                        set rsGc=nothing
                                    end if
                                Next
    
                                IF HaveGcsTest=1 THEN%>
                                <tr> 
                                    <td colspan="4" class="pcSpacer"></td>
                                </tr>
                                <tr> 
                                    <td colspan="4" class="pcSectionTitle">
                                    	<img src="images/pc4_notify.png" alt="" style="margin-right: 4px;">
                                        <%response.write dictLanguage.Item(Session("language")&"_NotifyRe_1")%>
                                    </td>
                                </tr>
                                <tr> 
                                    <td colspan="4" class="pcSpacer"></td>
                                </tr>
                                <tr>
                                    <td colspan="4">
                                        <%response.write dictLanguage.Item(Session("language")&"_NotifyRe_2")%>
                                    </td>
                                </tr>
                                <tr valign="top">
                                    <td>
                                        <%response.write dictLanguage.Item(Session("language")&"_NotifyRe_3")%>
                                    </td>
                                    <td><input type="text" size="20" name="GcReName" value=""></td>
                                    <td>
                                        <%response.write dictLanguage.Item(Session("language")&"_NotifyRe_4")%>
                                    </td>
                                    <td><input type="text" size="20" name="GcReEmail" value="" class="email"></td>
                                </tr>
                                <tr valign="top">
                                    <td>
                                        <%response.write dictLanguage.Item(Session("language")&"_NotifyRe_5")%>
                                    </td>
                                    <td colspan="3"><textarea cols="50" rows="3" name="GcReMsg"></textarea></td>
                                </tr>
                                <%END IF%>
                                <%
                                '// Gift Certs - E
                                %>
    
                                
                                <tr>
                                    <td colspan="4" style="padding-top: 10px;"><div id="OtherLoader"></div></td>
                                </tr>
							</table>
                            <input type="image" name="OtherSubmit" id="OtherSubmit" src="<%=RSlayout("pcLO_Update")%>" align="absmiddle" border="0" style="display:none">
							</form>
						</div>
                        <%
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						'// Other - End
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
						%>


                        <%
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						'// Gift Wrapping - Start
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
						%>
						<div id="GiftArea">
							<%
							
							'// Gift Wrapping - S
							
							query="SELECT pcGWSet_Show,pcGWSet_Overview,pcGWSet_HTML FROM pcGWSettings;"
							set rs=connTemp.execute(query)
							pcvGW=0
							pcvOverview=0
							if not rs.eof then
								pcvGW=rs("pcGWSet_Show")
								if IsNull(pcvGW) OR pcvGW="" then
									pcvGW="0"
								end if			
								session("Cust_GW")=pcvGW					
								pcvOverview=rs("pcGWSet_Overview")
								if pcvOverview="0" then
									pcvOverview=""
								end if
								session("Cust_GWText")=pcvOverview
								pcvGWDetails=rs("pcGWSet_HTML")
							end if
							set rs=nothing
							
							tmpStrList=""
							PrdCanGW=0
							tmpCheckedList=""
							
							IF pcvGW="1" THEN
								For f=1 to pcCartIndex
									if pcCartArray(f,10)=0 then
										PrdCanGWchecks=0
										query="SELECT pcGC_EOnly FROM pcGC WHERE pcGC_idproduct=" & pcCartArray(f,0)
										set rs1=connTemp.execute(query)
										if rs1.eof then
											query="SELECT pcPE_IDProduct FROM pcProductsExc WHERE pcPE_IDProduct=" & pcCartArray(f,0)
											set rs1=connTemp.execute(query)
											if rs1.eof then
												PrdCanGWchecks=1
											else 
												PrdCanGWchecks=0
											end if
										end if
										if PrdCanGWchecks=1 then
											PrdCanGW=PrdCanGW+1
											tmpStrList=tmpStrList & "<tr><td width='70%' nowrap>" & pcCartArray(f,1) & "</td><td>"										
											if (pcCartArray(f,34)<>"") and (pcCartArray(f,34)<>"0") then
												tmpStrList=tmpStrList & "<div id=""GWMarker" & pcCartArray(f,0) & """>"
												tmpStrList=tmpStrList & "<a href=""javascript:;"" onclick=""GWAdd('" & pcCartArray(f,0) & "','" & f & "');"">" & dictLanguage.Item(Session("language")&"_opc_giftWrap_1") & "</a>"
												tmpStrList=tmpStrList & "</div>"
											else
												tmpStrList=tmpStrList & "<div id=""GWMarker" & pcCartArray(f,0) & """>"
												tmpStrList=tmpStrList & "<a href=""javascript:;"" onclick=""GWAdd('" & pcCartArray(f,0) & "','" & f & "');"">" & dictLanguage.Item(Session("language")&"_opc_giftWrap_2") & "</a>"
												tmpStrList=tmpStrList & "</div>"
											end if											
											tmpStrList=tmpStrList & "</td></tr>"
										end if
									end if
								Next
								IF tmpStrList<>"" THEN 
									%>
                                    <form id="GWForm" name="GWForm">
                                        <table class="pcShowContent">
                                        <tr class="pcSectionTitle">
                                            <td colspan="2"><%=dictLanguage.Item(Session("language")&"_opc_32")%></td>
                                        </tr>
										<% if (session("Cust_GW")="1") and (pcvGWDetails<>"") and (session("Cust_GWText")="1") then %>
                                        <tr>
                                            <td colspan="2">
                                                <div style="margin-top:2px"><%=pcvGWDetails%></div>
                                            </td>
                                        </tr>
                                        <% end if %>
                                        <tr> 
                                            <td colspan="2" class="pcSpacer"></td>
                                        </tr>
                                        <tr>
                                            <td><%=dictLanguage.Item(Session("language")&"_opc_33")%></td>
                                            <td><%=dictLanguage.Item(Session("language")&"_opc_34")%></td>
                                        </tr>
                                        <%=tmpStrList%>
                                        <tr>
                                            <td colspan="2">
                                                <div id="GWLoader" style="display:none"></div>
                                            </td>
                                        </tr>
                                        </table>
                                        <script>
                                            var PrdCanGW=<%=PrdCanGW%>;
                                            var tmpCheckedList="<%=tmpCheckedList%>";
                                        </script>
                                        <input type="image" name="GWSubmit" id="GWSubmit" src="<%=RSlayout("pcLO_Update")%>" align="absmiddle" border="0" style="display:none">
                                    </form>
                                <%ELSE%>
                          <script>
									var PrdCanGW="";
									var tmpCheckedList="";
								</script>
                                <input type="image" name="GWSubmit" id="GWSubmit" src="<%=RSlayout("pcLO_Update")%>" align="absmiddle" border="0" style="display:none">
								<%END IF%>
                            <%ELSE%>
						  <script>
                                var PrdCanGW="";
                                var tmpCheckedList="";
                            </script>
                            <input type="image" name="GWSubmit" id="GWSubmit" src="<%=RSlayout("pcLO_Update")%>" align="absmiddle" border="0" style="display:none">
							<%END IF%>
                        <%
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						'// Gift Wrapping - End
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
						%>

                        <%
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						'// Payment - Start
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
						%>
						<div id="PayArea">
							<div id="PayNoNeed" style="display:none">
								<div class="pcSuccessMessage"><%=dictLanguage.Item(Session("language")&"_opc_40")%></div>
							</div>
							<div id="PayAreaSub">
							<script>
								var NeedToUpdatePay=0;
								var PaymentSelected="0";
								var CustomPayment=0;
							</script>
							<form id="PayForm" name="PayForm">
							<div class="pcSectionTitle"><%=dictLanguage.Item(Session("language")&"_orderverify_12")%></div>
							
							<% if session("ExpressCheckoutPayment") = "YES" then %>
                            
                            	<%
								query="SELECT payTypes.idPayment, payTypes.paymentDesc, payTypes.priceToAdd, payTypes.percentageToAdd, payTypes.Type, payTypes.gwCode, payTypes.paymentNickName, payTypes.sslUrl FROM payTypes WHERE (payTypes.active = - 1) AND (payTypes.pcPayTypes_ppab <> 1) AND (payTypes.gwCode = 999999 OR payTypes.gwCode = 46 OR payTypes.gwCode = 80 OR payTypes.gwCode = 99)"
                                set rs=server.CreateObject("Adodb.recordset")
                                set rs=conntemp.execute(query)
                                if not rs.eof then 
									tempidPayment=rs("idPayment")
									temppaymentDesc=rs("paymentDesc")
									temppriceToAdd=rs("priceToAdd")
									temppercentageToAdd=rs("percentageToAdd")
									tempType=rs("Type")
									tempgwCode=rs("gwCode")
									tempPaymentNickName=rs("paymentNickName")
									if isNull(temppriceToAdd) OR temppriceToAdd="" then
										temppriceToAdd=0
									end if
									if isNull(temppercentageToAdd) OR temppercentageToAdd="" then
										temppercentageToAdd=0
									end if
									if ccur(temppriceToAdd)<>0 then
										HowMuch=temppriceToAdd 
										HowMuch=roundTo(HowMuch,.01)           
									else
										HowMuch=""
									end if
									if ccur(temppercentageToAdd)<>0 then
										HowMuch1=temppercentageToAdd & "% of Order Total"         
									else
										HowMuch1=""
									end if
									tmpStrPay=tmpStrPay & "if (tmpid==" & tempidPayment & ") {CustomPayment=1; }" & vbcrlf
									%>
                                    <table id="PayList" class="pcShowContent">
                                        <tr>
                                            <td width="5%" nowrap>
                                                <input type="radio" id="chkPayment" name="chkPayment" class="chkPay" value="<%=tempidPayment%>" onclick="CheckPayType('<%=tempidPayment%>',1);" checked>
                                                <script type="text/javascript">
													$(document).ready(function() { $('#chkPayment').click(); });			
												</script>
                                            </td>
                                            <td><%=dictLanguage.Item(Session("language")&"_opc_56")%></td>
                                        </tr>
                                    </table>									
                                 <% end if %>
                            
							<% else %>
                                  <table id="PayList" class="pcShowContent">
                                  <% ' get available paytypes
                                  'If customer session
                                  If session("customerCategory")<>0 AND session("customerCategory")<>"" then
									  'SB S
									  strAndSub = ""
									  if pcIsSubscription = True Then
										 strAndSub = " AND pcPayTypes_Subscription = 1 ORDER BY payTypes.paymentPriority"
									  else
										 strAndSub = " ORDER BY payTypes.paymentPriority"
									  End if 
									  'SB E	
                                      query="SELECT payTypes.idPayment, payTypes.paymentDesc, payTypes.priceToAdd, payTypes.percentageToAdd, payTypes.Type, payTypes.gwCode, payTypes.paymentNickName, payTypes.sslUrl, CustCategoryPayTypes.idCustomerCategory FROM payTypes INNER JOIN CustCategoryPayTypes ON payTypes.idPayment = CustCategoryPayTypes.idPayment WHERE (payTypes.active = - 1) AND (payTypes.pcPayTypes_ppab <> 1) AND (payTypes.gwCode <> 50) AND (payTypes.gwCode <> 999999) AND (CustCategoryPayTypes.idCustomerCategory = "&session("customerCategory")&")" & strAndSub
                                      set rs=server.CreateObject("Adodb.recordset")
                                      set rs=conntemp.execute(query)
                                      if not rs.eof then
                                          tmpStrPay="" %>
                                      
                                          <% while not rs.eof
                                              tempidPayment=rs("idPayment")
                                              temppaymentDesc=rs("paymentDesc")
                                              temppriceToAdd=rs("priceToAdd")
                                              temppercentageToAdd=rs("percentageToAdd")
                                              tempType=rs("Type")
                                              tempgwCode=rs("gwCode")
                                              tempPaymentNickName=rs("paymentNickName")
                                              if isNull(temppriceToAdd) OR temppriceToAdd="" then
                                                  temppriceToAdd=0
                                              end if
                                              if isNull(temppercentageToAdd) OR temppercentageToAdd="" then
                                                  temppercentageToAdd=0
                                              end if
                                              if ccur(temppriceToAdd)<>0 then
                                                  HowMuch=temppriceToAdd 
                                                  HowMuch=roundTo(HowMuch,.01)           
                                              else
                                                  HowMuch=""
                                              end if
                                              if ccur(temppercentageToAdd)<>0 then
                                                  HowMuch1=temppercentageToAdd & "% of Order Total"
                                                  'HowMuch=temppriceToAdd + (temppercentageToAdd*intCalPaymnt/100)
                                                  'HowMuch=roundTo(HowMuch,.01)           
                                              else
                                                  HowMuch1=""
                                              end if
                                              CustomPayType=0
                                              payURL=rs("sslURL")
                                              if payURL<>"" then
                                                  if Instr(UCase(payURL),UCASE("paymnta_"))=1 then
                                                      CustomPayType=1
                                                      payURL="opc_" & payURL
                                                  end if
                                              end if
                                              if CustomPayType=1 then
                                                  tmpStrPay=tmpStrPay & "if (tmpid==" & tempidPayment & ") {CustomPayment=1; getPayDetails(" & tempidPayment & ",'" & payURL & "'); tmpOK=1;}" & vbcrlf
                                              end if
                                              %>
                                              <tr>
                                              <td width="5%" nowrap>
                                              <input type="radio" name="chkPayment" class="chkPay" class="required" value="<%=tempidPayment%>" onclick="CheckPayType('<%=tempidPayment%>',0);">
                                              </td>
                                              <td valign="middle">
												<%
                                                if tempgwCode="999999" OR tempgwCode="3" or tempgwCode="80" then
													if tempgwCode="3" or tempgwCode="80" then %>
                                                        Credit/Debit Card or PayPal
													<% else %>
                                                        <table cellpadding="2" cellspacing="2">
                                                        <tr>
                                                        <td style="vertical-align:middle"><img src="images/PayPal_mark_50x34.gif" width="50" height="34"></td>
                                                        <td style="vertical-align:middle" nowrap><span style="font-size:smaller;"><a href="https://www.paypal.com/us/cgibin/webscr?cmd=xpt/Marketing/popup/OLCWhatIsPayPal-outside" target="_blank">What is PayPal?</a></span></td>
                                                        </tr>
                                                        </table>
													<% end if
                                                else										  
												  if tempgwCode<>"7" then
													  if tempType="CU" then
														  response.write temppaymentDesc
													  else
														  response.write tempPaymentNickName
													  end if
												  else
													  response.write temppaymentDesc
												  end if 
												end if 
												
											  ' chedck to see if it's an interac online method
                                              if tempgwCode ="66" then
                                                  strInteractrade = "<i>&reg; Trade-mark of Interac Inc. Used under licence.<i/>"
                                              end if 
                                              if HowMuch<>"" then								
                                                  response.write " - "&scCurSign&money(HowMuch)
                                              else
                                                  if HowMuch1<>"" then
                                                      response.write " - "& HowMuch1
                                                  end if
                                              end if %>
                                              </td>
                                              </tr>
                                              <% rs.movenext
                                          wend
                                          set rs=nothing %>
                                      <%end if
                                  End if
								  
								  'SB S
								  strAndSub = ""
								  if pcIsSubscription = True Then
									 strAndSub = " AND pcPayTypes_Subscription = 1 ORDER by pcPayTypes_Subscription, paymentPriority"
								  else
									 strAndSub = " ORDER by paymentPriority"
								  End if 
								  'SB E
								  
                                  if session("customerType")=1 then
                                      query="SELECT idPayment,paymentDesc,priceToAdd,percentageToAdd, type, gwcode, paymentNickName,sslURL FROM paytypes WHERE active=-1 AND Cbtob<>2 AND (payTypes.pcPayTypes_ppab <> 1) AND (gwcode<>50 AND gwcode<>999999)" & strAndSub
                                  else
                                      query="SELECT idPayment,paymentDesc,priceToAdd,percentageToAdd, type, gwcode, paymentNickName,sslURL FROM paytypes WHERE active=-1 AND Cbtob=0 AND (payTypes.pcPayTypes_ppab <> 1) AND (gwcode<>50 AND gwcode<>999999)" & strAndSub
                                  end if
                                  set rs=server.CreateObject("Adodb.recordset")
                                  set rs=conntemp.execute(query)
      
                                  if err.number<>0 then
                                      call LogErrorToDatabase()
                                      set rs=nothing
                                      call closedb()
                                      response.redirect "techErr.asp?err="&pcStrCustRefID
                                  end if
      
                                  if not rs.eof then %>
                                      <% while not rs.eof
                                          tempidPayment=rs("idPayment")
                                          temppaymentDesc=rs("paymentDesc")
                                          temppriceToAdd=rs("priceToAdd")
                                          temppercentageToAdd=rs("percentageToAdd")
                                          tempType=rs("Type")
                                          tempgwCode=rs("gwCode")
                                          tempPaymentNickName=rs("paymentNickName")
                                          if isNull(temppriceToAdd) OR temppriceToAdd="" then
                                              temppriceToAdd=0
                                          end if
                                          if isNull(temppercentageToAdd) OR temppercentageToAdd="" then
                                              temppercentageToAdd=0
                                          end if
                                          if ccur(temppriceToAdd)<>0 then
                                              HowMuch=temppriceToAdd 
                                              HowMuch=roundTo(HowMuch,.01)           
                                          else
                                              HowMuch=""
                                          end if
                                          if ccur(temppercentageToAdd)<>0 then
                                              HowMuch1=temppercentageToAdd & "% of Order Total"
                                              'HowMuch=temppriceToAdd + (temppercentageToAdd*intCalPaymnt/100)
                                              'HowMuch=roundTo(HowMuch,.01)           
                                          else
                                              HowMuch1=""
                                          end if
                                          CustomPayType=0
                                          payURL=rs("sslURL")
                                          if payURL<>"" then
                                              if Instr(UCase(payURL),UCASE("paymnta_"))=1 then
                                                  CustomPayType=1
                                                  payURL="opc_" & payURL
                                              end if
                                          end if
                                          if CustomPayType=1 then
                                              tmpStrPay=tmpStrPay & "if (tmpid==" & tempidPayment & ") {CustomPayment=1; getPayDetails(" & tempidPayment & ",'" & payURL & "'); tmpOK=1;}" & vbcrlf
                                          end if
                                          %>
                                          <tr>
                                          <td width="5%" nowrap>
                                          <input type="radio" name="chkPayment" class="chkPay" class="required" value="<%=tempidPayment%>" onclick="CheckPayType('<%=tempidPayment%>',0);">
                                          </td>
                                          <td valign="middle">
                                          	<% if tempgwCode="999999" OR tempgwCode="3" or tempgwCode="80" then
													if tempgwCode="3" or tempgwCode="80" then %>
                                                        Credit/Debit Card or PayPal
													<% else %>
                                                        <table cellpadding="2" cellspacing="2">
                                                            <tr>
                                                                <td style="vertical-align:middle"><img src="images/PayPal_mark_50x34.gif" width="50" height="34"></td>
                                                                <td style="vertical-align:middle" nowrap><span style="font-size:smaller;"><a href="https://www.paypal.com/us/cgibin/webscr?cmd=xpt/Marketing/popup/OLCWhatIsPayPal-outside" target="_blank">What is PayPal?</a></span></td></tr></table>
                                                	<% end if 
											else		
											  if tempgwCode<>"7" then
												  if tempType="CU" then
													  response.write temppaymentDesc
												  else
													  response.write tempPaymentNickName
												  end if
											  else
												  response.write temppaymentDesc
											  end if 
											end if
											
										  ' chedck to see if it's an interac online method
                                          if tempgwCode ="66" then
                                              strInteractrade = "<i>&reg; Trade-mark of Interac Inc. Used under licence.<i/>"
                                          end if 
                                          if HowMuch<>"" then								
                                              response.write " - "&scCurSign&money(HowMuch)
                                          else
                                              if HowMuch1<>"" then
                                                  response.write " - "& HowMuch1
                                              end if
                                          end if %>
                                          </td>
                                          </tr>
                                          <% rs.movenext
                                      wend
                                      set rs=nothing %>
                                  <%end if%>
                                  </table>
                            
							<% end if %>
							<div id="PayFormArea" style="display:none"></div>
							</form>
							<script>
								function CheckPayType(tmpid,ctype)
								{
									PaymentSelected=tmpid;
									CustomPayment=0;
									var tmpOK=0;
									 $('.chkPay').attr('disabled','disabled');
									 $('#PayFormArea').html(); $('#PayFormArea').hide();
									 <%=tmpStrPay%>
									 if (tmpOK==0) {NeedToUpdatePay=0;}
									 if (ctype==0) 
									 {GetOrderInfo(tmpid,'#PayLoader1',ctype,'')}
									 else
									 { $('.chkPay').attr('disabled',false)}
								}
								
								function PreSelectPayType(tmpid)
								{
									if (tmpid!="")
									{
										var totalradio=document.getElementsByName("chkPayment").length;
										for (var i=0;i<totalradio;i++)
										{
											if (document.getElementsByName("chkPayment")[i].value+""==tmpid+"")
											{
												document.getElementsByName("chkPayment")[i].checked=true;
												CheckPayType(tmpid,1);
												$('#PayFormArea').hide();
												break;
											}
										}
									}
								}
								
							</script>
							</div>
							<div id="PayLoader" style="display:none"></div>
							<div id="PayLoader1" style="display:none"></div>
						</div>
                        <%
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						'// Payment - End
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
						%>
                        
                        <%
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						'// Discounts - Start
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
						%>
							<% DiscPanelHaveInfo=0
							
							IF (TurnOffDiscountCodesWhenHasSale="1") AND (scDB="SQL") AND (not scHideDiscField="1") THEN
								Dim tmpPrdList
								tmpPrdList=""
								for f=1 to ppcCartIndex
									if pcCartArray(f,10)=0 then
										if tmpPrdList<>"" then
											tmpPrdList=tmpPrdList & ","
										end if
										tmpPrdList=tmpPrdList & pcCartArray(f,0)
									end if
								next
								if tmpPrdList="" then
									tmpPrdList="0"
								end if
								tmpPrdList="(" & tmpPrdList & ")"
								query="SELECT idProduct FROM Products WHERE idProduct IN " & tmpPrdList & " AND pcSC_ID>0;"
								set rsQ=connTemp.execute(query)
								if not rsQ.eof then
									HavePrdsOnSale=1
								end if
								set rsQ=nothing
							END IF
							
							If (not scHideDiscField="1") AND (HavePrdsOnSale=0) then
								displayDiscountCode=""
								
								If pcv_strPayPanel = "1" Then
								
									pdiscountDetails = ""
									
									query="SELECT pcCustSession_discountcode FROM pcCustomerSessions WHERE pcCustomerSessions.idDbSession="&session("pcSFIdDbSession")&" AND pcCustomerSessions.randomKey="&session("pcSFRandomKey")&" AND pcCustomerSessions.idCustomer="&session("idCustomer")&";"
									set rs=server.CreateObject("ADODB.RecordSet")
									set rs=conntemp.execute(query)
									if NOT rs.eof then
										pdiscountDetails=rs("pcCustSession_discountcode")
									end if									
									set rs=nothing
									
									displayDiscountCode=pdiscountDetails
								
								Else
								
									query="SELECT discountcode FROM discounts WHERE pcDisc_Auto=1 AND active=-1 ORDER BY percentagetodiscount DESC,pricetodiscount DESC;"
									set rs=server.CreateObject("ADODB.RecordSet")
									set rs=conntemp.execute(query)
									if NOT rs.eof then
										pcStrAutoDiscCode=""
										do until rs.eof
											pcIntADCnt=pcIntADCnt+1
											if pcStrAutoDiscCode<>"" then
												pcStrAutoDiscCode=pcStrAutoDiscCode & ","
											end if
											pcStrAutoDiscCode=pcStrAutoDiscCode&rs("discountcode")
											rs.movenext
										loop
									end if
									displayDiscountCode=pcStrAutoDiscCode
									set rs=nothing
									
								End If
								if DiscPanelHaveInfo=0 then%>
								<div id="DiscArea">
								<form id="DiscForm" name="DiscForm">
								<table class="pcShowContent">
								<%DiscPanelHaveInfo=1
								end if%>
								<tr>
									<td width="35%" valign="top"><%=dictLanguage.Item(Session("language")&"_order_T")%></td>
									<td width="65%">
										<input type="text" id="DiscountCode" name="DiscountCode" value="<%=displayDiscountCode%>" size="30">&nbsp;
										<input type="image" name="DiscRecal" id="DiscRecal" src="<%=RSlayout("recalculate")%>"> 
										<input type="image" name="DiscSubmit" id="DiscSubmit" src="<%=RSlayout("recalculate")%>" style="display:none">                               
									</td>
								</tr>
								<tr>
									<td colspan="2" style="text-align: center;"><i><%=dictLanguage.Item(Session("language")&"_orderverify_41")%></i></td>
								</tr>
							<%End If%>
							
							<%
							IF RewardsActive=1 THEN
							if session("idCustomer")>"0" AND (session("CustomerGuest")="0" OR session("CustomerGuest")="2") then
								'Add visual separator
							if DiscPanelHaveInfo=1 then%>
								<tr><td colspan="2"><hr></td></tr>
                            <%end if
								'query="SELECT iRewardPointsAccrued, iRewardPointsUsed FROM Customers WHERE idcustomer=" & session("idCustomer") & " AND pcCust_Guest=0;"
								query="SELECT iRewardPointsAccrued, iRewardPointsUsed FROM Customers WHERE idcustomer=" & session("idCustomer") & ";"
								set rs=connTemp.execute(query)
								pcIntRewardPointsAccrued = 0
								pcIntRewardPointsUsed = 0
								if not rs.eof then
									pcIntRewardPointsAccrued = rs("iRewardPointsAccrued")
									pcIntRewardPointsUsed = rs("iRewardPointsUsed")
								end if
								set rs=nothing
								
								If RewardsActive = 1 Then
									opcSFIntBalance = 0
									If IsNull(pcIntRewardPointsAccrued) or pcIntRewardPointsAccrued="" Then 
										pcIntRewardPointsAccrued = 0
									End if
									If IsNull(pcIntRewardPointsUsed) or pcIntRewardPointsUsed="" Then 
										pcIntRewardPointsUsed = 0
									End if
									pcIntBalance = pcIntRewardPointsAccrued - pcIntRewardPointsUsed
									pcIntDollarValue = pcIntBalance * (RewardsPercent / 100)
									opcSFIntBalance = pcIntBalance
								End If
								
								'if customer has reward points - show total here 
								if opcSFIntBalance > 0 AND ((pcIntDollarValue > 0 AND session("customerType")<>"1") OR (session("customerType")="1" AND RewardsIncludeWholesale=1)) then
								if DiscPanelHaveInfo=0 then%>
								<div id="DiscArea">
								<form id="DiscForm" name="DiscForm">
								<table class="pcShowContent">
								<%DiscPanelHaveInfo=1
								end if%>
								<tr> 
									<td colspan="2">
										<i><%response.write ship_dictLanguage.Item(Session("language")&"_login_e")%> <%=opcSFIntBalance%>&nbsp;<%=RewardsLabel%> <%response.write ship_dictLanguage.Item(Session("language")&"_login_f")%> <%Response.Write scCurSign & money(pcIntDollarValue)%> <%response.write ship_dictLanguage.Item(Session("language")&"_login_g")%></i>
									</td>
								</tr>
								<% end if %>
                                <% If pcIntDollarValue>0 AND ((session("customerType")<>"1") OR (session("customerType")="1" AND RewardsIncludeWholesale=1)) Then
                                if DiscPanelHaveInfo=0 then%>
								<div id="DiscArea">
								<form id="DiscForm" name="DiscForm">
								<table class="pcShowContent">
								<%DiscPanelHaveInfo=1
								end if%>
                                <tr>
                                    <td width="35%">				
                                        <%response.write dictRewardsLanguage.Item(Session("language")&"_order_AA")%>
                                    </td>
                                    <td width="65%">
                                        <input type="text" id="UseRewards" name="UseRewards" size="30" maxlength="10" value="0">&nbsp;
                                        <input type="image" name="RewardsRecal" id="RewardsRecal" src="<%=RSlayout("recalculate")%>">  
                                        <input type="image" name="RewardsSubmit" id="RewardsSubmit" src="<%=RSlayout("recalculate")%>" style="display:none">  
                                    </td>
                                </tr>
                                <%end if 'if customer has reward points %>
							<%end if 'Customer Logged in
							END IF%>
							<%if DiscPanelHaveInfo=1 then
							session("NoNeedStep5")="0"%>
							<tr>
								<td colspan="2"><div id="DiscLoader"></div>
								<div id="DiscLoader1"></div></td>
							</tr>
							<%else
							session("NoNeedStep5")="1"
							end if%>
							<%if DiscPanelHaveInfo=1 then%>
							</table>
							</form>
						</div>
						<%end if%>
                        <%
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						'// Discounts - End
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
						%>
 
                        <%
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						'// Payment Button - Start
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
						%>
                        <table class="pcMainTable" style="margin-top: 20px;">
                            <tr>
                                <td align="center">
                                    <div id="ButtonArea" style="display:none;">
                                        <div id="PlaceOrderButton" style="display:none;">
                                            <a href="javascript:;" onclick="ValidateGroup1();"><img src="<%=RSlayout("pcLO_placeOrder")%>" border="0"></a>
                                        </div>
                                        <div id="ContinueButton" style="display:none;">
                                            <a href="javascript:;" onclick="ValidateGroup2();"><img src="<%=RSlayout("submit")%>" border="0"></a>
                                        </div>
                                        <div id="PlaceOrderTips" style="display:none;">
                                        </div>
                                    </div>
                                </td>
                            </tr>
                        </table>
                        <%
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						'// Payment Button - End
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
						%>
                       </div>
                    </div>                    
                </div>
                <%
				'/////////////////////////////////////////////////////////////////////////
				'// PAYMENT INFORMATION:  END
				'/////////////////////////////////////////////////////////////////////////
				%>   
				
				
				<%
				' END SPRY ACCORDION WITH FINAL DIV TAG
				%>

            </div>           
		</td>
    </table>
	</div>
    
	<%
    '/////////////////////////////////////////////////////////////////////////
    '// ORDER PREVIEW:  START
    '/////////////////////////////////////////////////////////////////////////
    %>   
    <a name="orderPreview"></a>
    <div id="opcOrderPreviewDIV">
        <div class="pcCheckoutSubTitle"><%=dictLanguage.Item(Session("language")&"_opc_41")%></div>
	<table class="pcMainTable" id="opcOrderPreview">
        <tr>
        	<td>
				<script>
					var OPCReady="NO";
					var tmpchkFree="";
				</script>
				<div id="OPRWarning">
					<!--#include file="opc_inc_viewitems.asp"-->
				</div>
				<div id="OPRArea" style="display:none">
				</div>
			</td>
        </tr>
		<tr>
			<td class="pcSpacer">&nbsp;</td>
		</tr>
    </table>
        </div>
	<%
    '/////////////////////////////////////////////////////////////////////////
    '// ORDER PREVIEW:  END
    '/////////////////////////////////////////////////////////////////////////
    %>  


  
    
 
<script type="text/javascript">
	var acc1 = new Spry.Widget.Accordion("acc1", { useFixedPanelHeights: false, enableAnimation: false });
	var currentPanel = 0;

	<% if session("idCustomer")>"0" then
		session("OPCstep")=2
	else
		session("OPCstep")=0
	end if %>
	
	//* Find Current Panel
	<% if len(Session("CurrentPanel"))=0 AND pcv_strPayPanel="" then %> 

		  <% if session("idCustomer")>"0" then %>
				acc1.openPanel('opcLogin');
				GoToAnchor('opcLoginAnchor');
				$('#LoginOptions').hide();
				$('#ShippingArea').hide(); 
				$('#BillingArea').show();
			$.ajax({
				type: "GET",
				url: "opc_cartcheck.asp",
				data: "{}",
				timeout: 45000,
				success: function(data, textStatus){
					if (data!="OK")
					{
						location="viewcart.asp";
					}
				}
	 		});
		  <% else %>
				$('#LoginOptions').show();
				$('#acc1').hide(); 
			$.ajax({
				type: "GET",
				url: "opc_cartcheck.asp",
				data: "{}",
				timeout: 45000,
				success: function(data, textStatus){
					if (data!="OK")
					{
						location="viewcart.asp";
					}
				}
	 		});
		  <% end if %>
		
	<% else %>
			
			<% If pcv_strPayPanel = "1" Then %>		
					$(document).ready(function() {
						$('#LoginOptions').hide();						
						pcf_LoadPaymentPanel();
			$.ajax({
				type: "GET",
				url: "opc_cartcheck.asp",
				data: "{}",
				timeout: 45000,
				success: function(data, textStatus){
					if (data!="OK")
					{
						location="viewcart.asp";
					}
				}
	 		});
					});
			<% Else %>
					acc1.openPanel('opcLogin');
					$('#LoginOptions').hide();
					$('#ShippingArea').hide(); 
					$('#BillingArea').show();
			$.ajax({
				type: "GET",
				url: "opc_cartcheck.asp",
				data: "{}",
				timeout: 45000,
				success: function(data, textStatus){
					if (data!="OK")
					{
						location="viewcart.asp";
					}
				}
	 		});
			<% End If %>
			
	<% end if %>

	function openme(pNumber) {
              acc1.openPanel(pNumber);                           
    }
	function toggle(pNumber) {
              var ele = acc1.getCurrentPanel();
              var panelNumber = acc1.getPanelIndex(ele);
              if (panelNumber == pNumber) {
                     acc1.closePanel(pNumber);
              } else {
                     acc1.openPanel(pNumber);    
              }                          
    }
    function togglediv(id) {
       var div = document.getElementById(id);
       if(div.style.display == 'block')
          div.style.display = 'none';
       else
          div.style.display = 'block';
    }
	
	<% If pcv_strPayPanel = "1" Then %>
		togglediv('comments');
		<% if scOrderName="1" then 'Allow customers to nickname their order %>
		togglediv('nickname');
		<% end if %>
	<% End If %>
	
	function win(fileName)
		{
			myFloater=window.open('','myWindow','scrollbars=yes,status=no,width=300,height=250')
			myFloater.location.href=fileName;
		}

</script>
<input type="image" name="LoadPaymentPanel" id="LoadPaymentPanel" src="<%=RSlayout("pcLO_Update")%>" border="0" style="display:none">
<%
Session("CurrentPanel") = ""
session("SF_DiscountTotal") = ""
session("SF_RewardPointTotal") = ""
%>
</div>
<!--#include file="footer.asp" -->
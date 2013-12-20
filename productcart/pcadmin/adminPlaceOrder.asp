<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin="7*9*"%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->    
<!--#include file="../includes/encrypt.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/stringfunctions.asp" -->
<!--#include file="../includes/languagesCP.asp" -->
<%
	'// Set Admin Order session
	session("pcAdminOrder")=Cint(1)

	'// Get customer ID and build redirect link
	'// Retrieve email and password from the database
	Dim pidcustomer, rs, query, conntemp
	call openDb()
	pidcustomer = getUserInput(request.QueryString("idcustomer"),15)
	if not validNum(pidcustomer) then response.redirect "viewCusta.asp"
	query="SELECT email,password FROM customers WHERE idCustomer=" & pidcustomer
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)
	if rs.eof then
		set rs=nothing
		call closedb()
		response.redirect "viewCusta.asp"
	else
		pemail=rs("email")
		ppassword=enDeCrypt(rs("password"), scCrypPass)
		ppassword=encrypt(ppassword, 9286803311968)
	end if
	set rs=nothing
	call closeDb()

	'// Clear any previous "place order" sessions
	'// Session variables are ordered alphabetically
		dim pcCartArray2(100,45)
		session("admin-" & session("GWPaymentId") & "-expMonth")=""
		session("admin-" & session("GWPaymentId") & "-expYear")=""
		session("admin-" & session("GWPaymentId") & "-pCardNumber")=""
		session("admin-" & session("GWPaymentId") & "-pCardType")=""
		session("ATBCustomer")=Cint(0)
		session("ATBPercentage")=Cint(0)
		session("ATBPercentOff")=Cint(0)
		Session("ContinueRef")=""
		Session("CurrentPanel")=""
		session("Cust_BuyGift")=""
		session("Cust_IDEvent")=""
		session("customerCategory")=Cint(0)
		session("customerCategoryType")=""
		session("CustomerGuest")=""
		session("customerType")=Cint(0)
		session("DCODE")=""
		session("DF1")=""
		session("Entered-" & session("GWPaymentId"))=""
		session("EPN_idOrder")=""
		session("ExpressCheckoutPayment")=""
		session("Gateway")=""
		session("gHideAddress")=""
		session("GWAuthCode")=""
		session("GWOrderDone")=""
		session("GWOrderId")=""
		session("GWPaymentId")=""
		session("GWSessionID")=""
		session("GWTransId")=""
		session("GWTransType")=""
		session("idAffiliate")=Cint(1)
		session("idcustomer")=""
		session("idGWSubmit")=""
		session("idGWSubmit2")=""
		session("idGWSubmit3")=""
		session("idOrder")=""
		session("idOrderSaved")=""
		session("IDRefer")=""
		session("iOrderTotal")=""
		session("NeedToUpdatePay")=""
		session("OPCstep")=""
		session("pc_pidOrder")=""
		Session("pcCartIndex")=Cint(0)
		Session("pcCartSession")=pcCartArray2
		Session("pcPromoIndex")=""
		Session("pcPromoSession")=""
		session("pcSFCartRewards")=Cint(0)
		Session("pcSFDF1")=""
		session("pcSFIdDbSession")=""
		session("pcSFIdDbSession")=""
		session("pcSFLoginEmail")=""
		session("pcSFLoginPassword")=""
		session("pcSFPassWordExists")=""
		Session("pcSFpcBillingAddress")=""
		Session("pcSFpcBillingAddress2")=""
		Session("pcSFpcBillingCity")=""
		Session("pcSFpcBillingCompany")=""
		Session("pcSFpcBillingCountryCode")=""
		Session("pcSFpcBillingFirstName")=""
		Session("pcSFpcBillingLastName")=""
		Session("pcSFpcBillingPhone")=""
		Session("pcSFpcBillingPostalCode")=""
		Session("pcSFpcBillingProvince")=""
		Session("pcSFpcBillingStateCode")=""
		Session("pcSFpcCustomerEmail2")=""
		Session("pcSFpcShippingAddress")=""
		Session("pcSFpcShippingAddress2")=""
		Session("pcSFpcShippingCity")=""
		Session("pcSFpcShippingCompany")=""
		Session("pcSFpcShippingCountryCode")=""
		Session("pcSFpcShippingEmail")=""
		Session("pcSFpcShippingFax")=""
		Session("pcSFpcShippingFirstName")=""
		Session("pcSFpcShippingIdRefer")=""
		Session("pcSFpcShippingLastName")=""
		Session("pcSFpcShippingNickName")=""
		Session("pcSFpcShippingPhone")=""
		Session("pcSFpcShippingPostalCode")=""
		Session("pcSFpcShippingProvince")=""
		Session("pcSFpcShippingReferenceId")=""
		Session("pcSFpcShippingResidential")=""
		Session("pcSFpcShippingStateCode")=""
		session("pcSFRandomKey")=""
		Session("pcSFTF1")=""
		session("pcSFUseRewards")=Cint(0)
		session("pcStrCustName")=""
		session("redirectPage")=""
		session("RefRewardPointsTest")=""
		session("SaveOrder")=""
		session("SF_DiscountTotal")=""
		session("SF_RewardPointTotal")=""
		Session("SFStrRedirectUrl")=""
		session("shippingAddress")=""
		session("shippingAddress2")=""
		session("shippingCity")=""
		session("shippingCompany")=""
		session("shippingCountryCode")=""
		session("shippingFullName")=""
		session("shippingPhone")=""
		session("shippingState")=""
		session("shippingStateCode")=""
		session("shippingZip")=""
		session("specialdiscount")=""
		session("TF1")=""
		
		session("ppassword") = ppassword
		
	'// Redirect to storefront
	response.Redirect "../pc/checkout.asp?cmode=3&LoginEmail=" & pemail
	response.End()
%>
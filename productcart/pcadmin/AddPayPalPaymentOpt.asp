<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add PayPal Payment Option" %>
<% section="paymntOpt" %>
<%PmAdmin=5%><!--#include file="adminv.asp"-->  
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/rc4.asp"--> 
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="AdminHeader.asp"-->
<% dim query, connTemp, rs

sMode=Request.Form("Submit")

dim ppAB
ppAB=0 '// Default value for PayPal Accelerated Boarding

'//////////////////////////////////////////////////////////////////////////////////
'// START: CHOOSE MODE
'//////////////////////////////////////////////////////////////////////////////////
If sMode <> "" Then	

	'******************************************************************************
	'// START: ADD SELECTED OPTIONS
	'******************************************************************************
	If sMode="Add Selected Options" Then
		dim varCheck
		varCheck=0
		
		'Start SDBA
		pcv_processOrder=request.Form("pcv_processOrder")
		if pcv_processOrder="" then
			pcv_processOrder="0"
		end if
		pcv_setPayStatus=request.Form("pcv_setPayStatus")
		if pcv_setPayStatus="" then
			pcv_setPayStatus="3"
		end if
		'End SDBA

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Gateways
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		
		'---Start PayPal ---
		gwpp=request.form("gwpp")	
		pcv_intComponentPass = 1
		If gwpp="1" AND pcv_intComponentPass = 1 then
			varCheck=1
			'// request gateway variables and insert them into the paypal table check payment type
			PayPalSelection=request.Form("PayPalSelection")
			if PayPalSelection = "" then 
				response.redirect "AddPayPalPaymentOpt.asp?gwchoice=PayPal&ppv=0&msg="&Server.URLEncode("You did not select a PayPal Payment Option.")
			end if
		
			if PayPalSelection="PayPal" then
				PayPalPaymentURL="gwpp.asp"
				PayPalName="PayPal"
				PayPal_Email=request.Form("PayPal_Email")
				pcPay_PayPal_Sandbox=request.Form("PayPal_Sandbox")
				if PayPal_Email = "" then 
					response.redirect "AddPayPalPaymentOpt.asp?gwchoice=PayPal&ppv=0&msg="&Server.URLEncode("You must enter your PayPal email address to activate PayPal Standard.")
				end if
				PayPal_Currency=request.Form("PayPal_Currency")
				ppGwcode=3
			else
				
				if PayPalSelection="PayPalWP" then
					pcPay_PayPal_TransType=request.Form("pcPay_PayPal_TransType")
					pcPay_PayPal_Username=request.Form("pcPay_PayPal_Username")
					pcPay_PayPal_Subject=request.Form("pcPay_PayPal_Subject")
					pcPay_PayPal_Password=request.Form("pcPay_PayPal_Password")
					pcPay_PayPal_Sandbox=request.Form("pcPay_PayPal_Sandbox")
					pcPay_PayPal_Vendor=""
					pcPay_PayPal_Partner=""
					pcPay_PayPal_Signature=request.Form("pcPay_PayPal_Signature")
					pcPay_PayPal_Currency=request.Form("pcPay_PayPal_Currency")
					pcPay_PayPal_CVC=request.Form("pcPay_PayPal_CVC")
					pcPay_PayPal_CardTypes=request.Form("CardTypes")
					PayPalPaymentURL="gwPayPal.asp"
					PayPalName="PayPal Website Payments Pro"
					ppGwcode=46
				
				elseif PayPalSelection="PayPalWPUK" then
					pcPay_PayPal_TransType=request.Form("pcPay_PayPalUK_TransType")
					pcPay_PayPal_Username=request.Form("pcPay_PayPalUK_Username")
					pcPay_PayPal_Subject=request.Form("pcPay_PayPalUK_Subject")
					pcPay_PayPal_Password=request.Form("pcPay_PayPalUK_Password")
					pcPay_PayPal_Sandbox=request.Form("pcPay_PayPalUK_Sandbox")
					pcPay_PayPal_Vendor=request.Form("pcPay_PayPalUK_Vendor")
					pcPay_PayPal_Partner=request.Form("pcPay_PayPalUK_Partner")
					pcPay_PayPal_Signature=""
					pcPay_PayPal_Currency=request.Form("pcPay_PayPalUK_Currency")
					pcPay_PayPal_CVC=request.Form("pcPay_PayPalUK_CVC")
					pcPay_PayPal_CardTypes=request.Form("CardTypesUK")
					PayPalPaymentURL="gwPayPalUK.asp"
					PayPalName="PayPal Website Payments Pro"
					ppGwcode=53
					If pcPay_PayPal_Vendor="" OR pcPay_PayPal_Partner="" then
						Session("pcAdminUserName")=pcPay_PayPal_Username
						Session("pcAdminSubject")=pcPay_PayPal_Subject
						Session("pcAdminPassword")=pcPay_PayPal_Password
						Session("pcAdminPartner")=pcPay_PayPal_Partner										
						Session("pcAdminVendor")=pcPay_PayPal_Vendor
						response.redirect "AddPayPalPaymentOpt.asp?gwchoice=PayPal&ppv=4&msg="&Server.URLEncode("An error occurred while trying to add Website Payments Pro as your payment gateway. <b>""Vendor""</b> and <b>""Partner""</b> are required fields.")
					End If
				
				elseif PayPalSelection="PayPalExpUK" then 
					pcPay_PayPal_TransType=request.Form("pcPay_PayPalUKe_TransType")
					pcPay_PayPal_Username=request.Form("pcPay_PayPalUKe_Username")
					pcPay_PayPal_Subject=request.Form("pcPay_PayPalUKe_Subject")
					pcPay_PayPal_Password=request.Form("pcPay_PayPalUKe_Password")
					pcPay_PayPal_Sandbox=request.Form("pcPay_PayPalUKe_Sandbox")
					pcPay_PayPal_Vendor=request.Form("pcPay_PayPalUKe_Vendor")
					pcPay_PayPal_Partner=request.Form("pcPay_PayPalUKe_Partner")
					pcPay_PayPal_Signature=""
					pcPay_PayPal_Currency=request.Form("pcPay_PayPalUKe_Currency")
					pcPay_PayPal_CVC=request.Form("pcPay_PayPalUKe_CVC")
					pcPay_PayPal_CardTypes=request.Form("CardTypesUKe")
					PayPalPaymentURL=""
					PayPalName="PayPal Express Checkout"
					ppGwCode=999999
					If pcPay_PayPal_Vendor="" OR pcPay_PayPal_Partner="" then
						Session("pcAdminUserName")=pcPay_PayPal_Username
						Session("pcAdminSubject")=pcPay_PayPal_Subject
						Session("pcAdminPassword")=pcPay_PayPal_Password
						Session("pcAdminPartner")=pcPay_PayPal_Partner										
						Session("pcAdminVendor")=pcPay_PayPal_Vendor
						response.redirect "AddPayPalPaymentOpt.asp?gwchoice=PayPal&ppv=3&msg="&Server.URLEncode("An error occurred while trying to add Website Payments Pro as your payment gateway. <b>""Vendor""</b> and <b>""Partner""</b> are required fields.")
					End If
					
				else
					pcPay_PayPal_TransType=request.Form("pcPay_PayPale_TransType")
					pcPay_PayPal_Username=request.Form("pcPay_PayPale_Username")
					pcPay_PayPal_Subject=request.Form("pcPay_PayPale_Subject")
					pcPay_PayPal_Password=request.Form("pcPay_PayPale_Password")
					pcPay_PayPal_Sandbox=request.Form("pcPay_PayPale_Sandbox")
					pcPay_PayPal_Vendor=""
					pcPay_PayPal_Partner=""
					pcPay_PayPal_Signature=request.Form("pcPay_PayPale_Signature")
					pcPay_PayPal_Currency=request.Form("pcPay_PayPale_Currency")
					pcPay_PayPal_CVC=request.Form("pcPay_PayPale_CVC")
					pcPay_PayPal_CardTypes=request.Form("CardTypese")
					PayPalPaymentURL=""
					PayPalName="PayPal Express Checkout"
					ppGwCode=999999
				end if
				
			end if
			if pcPay_PayPal_Sandbox="YES" then
				pcPay_PayPal_Sandbox=1
			else
				pcPay_PayPal_Sandbox=0
			end if
			if pcPay_PayPal_CVC="" then
				pcPay_PayPal_CVC=0
			end if

			priceToAddType=request.Form("priceToAddType")
			if priceToAddType="price" then
				priceToAdd=replacecomma(Request("priceToAdd"))
				percentageToAdd="0"
				If priceToAdd="" Then
					priceToAdd="0"
				end if
			else
				priceToAdd="0"
				percentageToAdd=request.Form("percentageToAdd")
				If percentageToAdd="" Then
					percentageToAdd="0"
				end if
			end if
			paymentNickName=replace(request.Form("paymentNickName"),"'","''")
			if paymentNickName="" then
				paymentNickName="PayPal"
			End If
			
			If (pcPay_PayPal_Signature="" OR pcPay_PayPal_Username="" OR pcPay_PayPal_Password="") AND (ppGwCode=46 OR ppGwCode=999999) Then
				ppAB=1
			Else 
				ppAB=0
			End If
			
			err.clear
			err.number=0
			
			call openDb() 
			
			'// Update the PayPal Table
			if PayPalSelection="PayPal" then
				query="UPDATE paypal SET Pay_To='"&PayPal_Email&"',URL='https://www.paypal.com/cgi-bin/webscr',PP_Currency='"&PayPal_Currency&"', PP_Sandbox="&pcPay_PayPal_Sandbox&" WHERE ID=1"
				set rs=Server.CreateObject("ADODB.Recordset")     
				set rs=connTemp.execute(query)

				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if

			else
				query="UPDATE pcPay_PayPal SET pcPay_PayPal_TransType="&pcPay_PayPal_TransType&", pcPay_PayPal_Username='"&pcPay_PayPal_Username&"', pcPay_PayPal_Subject='"&pcPay_PayPal_Subject&"', pcPay_PayPal_Password='"&pcPay_PayPal_Password&"', pcPay_PayPal_Sandbox="&pcPay_PayPal_Sandbox&", pcPay_PayPal_Signature='"&pcPay_PayPal_Signature&"', pcPay_PayPal_Currency='"&pcPay_PayPal_Currency&"', pcPay_PayPal_CVC="&pcPay_PayPal_CVC&", pcPay_PayPal_Vendor='"&pcPay_PayPal_Vendor&"', pcPay_PayPal_Partner='"&pcPay_PayPal_Partner&"', pcPay_PayPal_CardTypes='"&pcPay_PayPal_CardTypes&"' WHERE (((pcPay_PayPal_ID)=1));"
				set rs=Server.CreateObject("ADODB.Recordset")     
				set rs=connTemp.execute(query)

				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
			end if
			

			'// Update the PayTypes Table
			query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName,pcPayTypes_ppab) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'"&PayPalName&"','"&PayPalPaymentURL&"',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&","&ppGwcode&",'"&paymentNickName&"',"&ppAB&")"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			if ppGwcode<>3 then
				session("pcSetupPayPalExpress") = ""
			end if
			set rs=nothing
			
			call closedb()
		end if
		'---End PayPal ---
		
		'---Start PayPal PayFlow Pro ---
		gwvpfp=request.form("gwvpfp")
		If gwvpfp="1" then
			varCheck=1
			'request gateway variables and insert them into the verisign_pfp table
			v_Url="na"
			v_Type=request.Form("pfp_Type")
			v_User=request.Form("pfp_User")
			v_Partner=request.Form("pfp_Partner")
			v_Password=request.Form("pfp_Password")
			v_Tender=request.Form("pfp_Tender")
			pfp_testmode=request.Form("pfp_testmode")
			if pfp_testmode="" then
				pfp_testmode=0
			end if
			pfp_transtype=request.Form("pfp_transtype")
			pfp_CSC=request.Form("pfp_CSC")
			priceToAddType=request.Form("priceToAddType")
			if priceToAddType="price" then
				priceToAdd=replacecomma(Request("priceToAdd"))
				percentageToAdd="0"
				If priceToAdd="" Then
					priceToAdd="0"
				end if
			else
				priceToAdd="0"
				percentageToAdd=request.Form("percentageToAdd")
				If percentageToAdd="" Then
					percentageToAdd="0"
				end if
			end if
			paymentNickName=replace(request.Form("paymentNickName"),"'","''")
			if paymentNickName="" then
				paymentNickName="Credit Card"
			end if
			v_Vendor=request.Form("pfp_Vendor")
			If v_Vendor="" then
				session("adminv_Type")=v_Type
				session("adminv_User")=v_User
				session("adminv_Partner")=v_Partner
				session("adminv_Password")=v_Password
				session("adminv_Tender")=v_Tender
				session("adminpfp_testmode")=v_testmode
				response.redirect "AddPayPalPaymentOpt.asp?msg="&Server.URLEncode("An error occurred while trying to add Payflow Pro as your payment gateway. <b>""Vendor""</b> is a required field.")
			End If 
			
			'check to see if centinel is activated
			pcPay_Cent_Active=request.Form("pcPay_Cent_Active_pfp")
			if pcPay_Cent_Active="YES" then
				pcPay_Cent_Active=1
				pcPay_Cent_TransactionURL=request.Form("pcPay_Cent_TransactionURL_pfp")
				pcPay_Cent_MerchantID=request.Form("pcPay_Cent_MerchantID_pfp")
				pcPay_Cent_ProcessorId=request.Form("pcPay_Cent_ProcessorId_pfp")
				if pcPay_Cent_TransactionURL="" or pcPay_Cent_MerchantID="" OR	pcPay_Cent_ProcessorId="" then
					response.redirect "AddPayPalPaymentOpt.asp?msg="&Server.URLEncode("An error occurred while trying to add Cardinal Centinel for Authorize.Net. <b>""Tranaction URL, Merchant ID and Process ID""</b> are all required fields.")
				else
					err.clear
					err.number=0
					
					call openDb() 
					query="UPDATE pcPay_Centinel SET pcPay_Cent_TransactionURL='"&pcPay_Cent_TransactionURL&"',	pcPay_Cent_MerchantID='"&pcPay_Cent_MerchantID&"',	pcPay_Cent_ProcessorId='"&pcPay_Cent_ProcessorId&"', pcPay_Cent_Active=1 WHERE pcPay_Cent_ID=1;"
					set rs=Server.CreateObject("ADODB.Recordset")     
					set rs=connTemp.execute(query)
					
					if err.number<>0 then
						call LogErrorToDatabase()
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					
					set rs=nothing
					
					call closedb()
				end if
			end if
			
			err.clear
			err.number=0
			
			call openDb() 

			query="UPDATE verisign_pfp SET v_Url='na',v_Type='"&v_Type&"',v_User='"&v_User&"',v_Partner='"&v_Partner&"' ,v_Password='"&v_Password &"' ,v_Vendor='"&v_Vendor&"',v_Tender='na',pfl_testmode='"&pfp_testmode&"',pfl_transtype='"&pfp_transtype&"',pfl_CSC='"&pfp_CSC&"' WHERE id=1"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=connTemp.execute(query)

			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if

			query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName, pcPayTypes_ppab) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'PayPal-PayFlow-Pro','gwpfp.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",2,'"&paymentNickName&"',"&ppAB&")"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			set rs=nothing
			
			call closedb()
		end if
		'---End PayPal PayFlow Pro---
		'
		'---Start PayPal PayFlow Link---
		gwvpfl=request.form("gwvpfl")
		If gwvpfl="1" then
			varCheck=1
			'request gateway variables and insert them into the verisign_pfp table
			v_User=request.Form("v_User")
			v_Partner=request.Form("v_Partner")
			pfl_testmode=request.Form("pfl_testmode")
			if pfl_testmode="" then
				pfl_testmode=0
			end if
			pfl_transtype=request.Form("pfl_transtype")
			pfl_CSC=request.Form("pfl_CSC")
			priceToAddType=request.Form("priceToAddType")
			if priceToAddType="price" then
				priceToAdd=replacecomma(Request("priceToAdd"))
				percentageToAdd=request.Form("percentageToAdd")
				If priceToAdd="" Then
					priceToAdd="0"
				end if
			else
				priceToAdd="0"
				percentageToAdd=request.Form("percentageToAdd")
				If percentageToAdd="" Then
					percentageToAdd="0"
				end if
			end if
			paymentNickName=replace(request.Form("paymentNickName"),"'","''")
			if paymentNickName="" then
				paymentNickName="Credit Card"
			End If 
			
			err.clear
			err.number=0
			
			call openDb() 

			query="UPDATE verisign_pfp SET v_Url='na',v_Type='na',v_User='"&v_User&"',v_Partner='"&v_Partner&"' ,v_Password='na' ,v_Vendor='na',v_Tender='na',pfl_testmode='"&pfl_testmode&"',pfl_transtype='"&pfl_transtype&"',pfl_CSC='"&pfl_CSC&"' WHERE id=1"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=connTemp.execute(query)

			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if

			query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName,pcPayTypes_ppab) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'PayPal-PayFlow-Link','gwpfl.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",9,'"&paymentNickName&"',"&ppAB&")"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=connTemp.execute(query)
			
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			set rs=nothing
			
			call closedb()
		end if
		'---End PayPal PayFlow Link---

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Gateways
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Centinel
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'check to see if centinel is activated
		pcPay_Cent_Active=request.Form("pcPay_Cent_Active")
		if pcPay_Cent_Active="YES" then
			
			pcPay_Cent_Active=1
			pcPay_Cent_TransactionURL=request.Form("pcPay_Cent_TransactionURL")
			pcPay_Cent_MerchantID=request.Form("pcPay_Cent_MerchantID")
			pcPay_Cent_ProcessorId=request.Form("pcPay_Cent_ProcessorId")
			pcPay_Cent_Password=request.Form("pcPay_Cent_Password")
			
			if pcPay_Cent_TransactionURL="" OR pcPay_Cent_MerchantID="" OR pcPay_Cent_ProcessorId="" OR pcPay_Cent_Password="" then
				
				'// response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Cardinal Centinel for Authorize.Net. <b>""Tranaction URL, Merchant ID and Process ID""</b> are all required fields.")
				
			else

				call openDb()
				query="UPDATE pcPay_Centinel SET pcPay_Cent_TransactionURL='"&pcPay_Cent_TransactionURL&"',	pcPay_Cent_MerchantID='"&pcPay_Cent_MerchantID&"',	pcPay_Cent_ProcessorId='"&pcPay_Cent_ProcessorId&"', pcPay_Cent_Password='"&pcPay_Cent_Password&"', pcPay_Cent_Active=1 WHERE pcPay_Cent_ID=1;"
				set rs=Server.CreateObject("ADODB.Recordset")     
				set rs=connTemp.execute(query)
				set rs=nothing
				call closedb()

			end if
			
		end if
		
		call openDb()
		query="SELECT pcPay_Centinel.pcPay_Cent_TransactionURL, pcPay_Centinel.pcPay_Cent_ProcessorId, pcPay_Centinel.pcPay_Cent_MerchantID, pcPay_Centinel.pcPay_Cent_Active, pcPay_Centinel.pcPay_Cent_Password FROM pcPay_Centinel WHERE (((pcPay_Centinel.pcPay_Cent_ID)=1));"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		pcPay_Cent_TransactionURL = rs("pcPay_Cent_TransactionURL")
		pcPay_Cent_ProcessorId = rs("pcPay_Cent_ProcessorId")
		pcPay_Cent_MerchantID = rs("pcPay_Cent_MerchantID")
		pcPay_Cent_Password=rs("pcPay_Cent_Password")
		pcPay_Cent_Active=rs("pcPay_Cent_Active")
		set rs=nothing
		call closedb()
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Centinel
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	

	end if
	'******************************************************************************
	'// END: ADD SELECTED OPTIONS
	'******************************************************************************
	
	If sMode="Add Selected Options" and varCheck=1 then
		response.redirect "PaymentOptions.asp"
	else
		response.redirect "AddPayPalPaymentOpt.asp?msg="&Server.URLEncode("You did not specify a payment option to add. Make sure that you check the box next to the payment option that you wish to add.")
	end if
end if
'//////////////////////////////////////////////////////////////////////////////////
'// END: CHOOSE MODE
'//////////////////////////////////////////////////////////////////////////////////




'//////////////////////////////////////////////////////////////////////////////////
'// START: PAYMENT TYPES
'//////////////////////////////////////////////////////////////////////////////////

err.clear
err.number=0
call openDb()  

query="SELECT idPayment, gwCode, active FROM paytypes;"
set rs=Server.CreateObject("ADODB.Recordset")     
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if NOT rs.eof then 
	iCnt=1
	'******************************************************************************
	'// START: ACTIVE GATEWAYS
	'******************************************************************************
	do until rs.eof
		varIDTemp=rs("idPayment")
		varTemp=rs("gwCode")
		varActive=rs("active")
		actGW=0
		if varActive<>"0" then
			select case varTemp
			case 2
				gwvpfp="1"
				actGW=1
				pcv_strPFPID=varIDTemp
				pcv_strPFPGW=varTemp
				exit do
			case 3
				gwpp="1"
				actGW=1
				pcv_strPPID=varIDTemp
				pcv_strPPGW=varTemp
				exit do
			case 9
				gwvpfl="1"
				actGW=1
				pcv_strPFLID=varIDTemp
				pcv_strPFLGW=varTemp
				exit do
			case 46
				gwpp="1"
				actGW=1
				pcv_strPPID=varIDTemp
				pcv_strPPGW=varTemp
				exit do
			case 53
				gwpp="1"
				actGW=1
				pcv_strPPID=varIDTemp
				pcv_strPPGW=varTemp
				exit do
			case 999999
				gwpp="1"
				actGW=1
				pcv_strPPID=varIDTemp
				pcv_strPPGW=varTemp
				exit do
			end select
		end if
		rs.moveNext
	loop
	'******************************************************************************
	'// END: ACTIVE GATEWAYS
	'******************************************************************************
	set rs=nothing
end if
call closedb()

'//////////////////////////////////////////////////////////////////////////////////
'// END: PAYMENT TYPES
'//////////////////////////////////////////////////////////////////////////////////
%>

<script language="JavaScript"><!--
<!-- Hide me
function wincomcheckobj(fileName)
	{
	myFloater=window.open('','myWindow','scrollbars=no,status=no,width=300,height=250')
	myFloater.location.href=fileName;
	}
function wingwaycheckobj(fileName)
	{
	myFloater=window.open('','myWindow','scrollbars=yes,status=no,width=475,height=425')
	myFloater.location.href=fileName;
	}
// show me-->
//-->
</script>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<%
pcPay_PayPal_Subject=scEmail
%>

<form name="form1" method="post" action="AddPayPalPaymentOpt.asp" class="pcForms">
	
<% if actGW=1 then %>	
	<div style="float: right; margin: 10px 25px 0 0;"><a href="AdminPaymentOptions.asp">Return to Payment Options</a></div>
	<table class="pcCPcontent">		
		<tr>
			 <th colspan="2">The following PayPal payment options have been configured: </th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>	
		<tr> 
			<td valign="top"> 
				<!--Enabled start here -->
				<table width="100%">					
					<% if gwpp="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<% 
						pcv_strPayPalOption=""
						if varTemp=3 then
							pcv_strPayPalOption="PayPal Standard"
						elseif varTemp=46 then
							pcv_strPayPalOption="PayPal Website Payment Pro (US)"
						elseif varTemp=53 then
							pcv_strPayPalOption="PayPal Website Payment Pro (UK)"
						else
							pcv_strPayPalOption="PayPal Express Checkout"
						end if
						%>
						<td height="21">
                            <% if varTemp=3 then %>
                            	<div class="pcCPnotes">
                                	<strong>NOTE:</strong> PayPal Standard must be removed to add Express Checkout or WPP.
                                </div>
                            <% end if %> 
							<%=pcv_strPayPalOption%> is enabled - <a href="modPayPalPaymentOpt.asp?mode=Edit&id=<%=pcv_strPPID%>&gwCode=<%=pcv_strPPGW%>">Modify</a>
                            &nbsp;|&nbsp;
                            <%
							call opendb()
							query="SELECT active,idPayment,gwCode,paymentDesc FROM paytypes WHERE gwCode=3 ORDER BY paymentDesc"
							set rs=Server.CreateObject("ADODB.Recordset")     
							set rs=conntemp.execute(query)	
							id=rs("idPayment")
							gwCode=rs("gwCode")					
							call closedb()
							%>
                            <a href="javascript:if (confirm('You are about to remove this payment option. Are you sure you want to complete this action?')) location='modPayPalPaymentOpt.asp?mode=Del&id=<%=id%>&gwCode=<%=gwCode%>'">Remove</a>                           
                        </td>
					</tr>
					<% end if %>
					
					<% if gwvpfp="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">PayPal Payflow Pro is enabled - <a href="modPayPalPaymentOpt.asp?mode=Edit&id=<%=pcv_strPFPID%>&gwCode=<%=pcv_strPFPGW%>">Modify</a></td>
					</tr>
					<% end if %>
					
					<% if gwvpfl="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">PayPal Payflow Link is enabled - <a href="modPayPalPaymentOpt.asp?mode=Edit&id=<%=pcv_strPFLID%>&gwCode=<%=pcv_strPFLGW%>">Modify</a></td>
					</tr>
					<% end if %>
			
					<% if actGW=1 then %>
					<tr> 
						<td colspan="2">&nbsp;</td>
					</tr>
					<% end if %>								
				</table>
				<!--Enabled end here -->			
			</td>
		</tr>	
	</table>
<% end if %>

<% 
if request("gwChoice")="" then '// If gwChoice is empty, hide the rest of the page 
%>
	
	<% if gwvpfl<>"1" OR gwpp<>"1" then %>
	<table class="pcCPcontent">
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<th colspan="2">Select &amp; Configure the PayPal payment option that you would like to use:</th>
		</tr>
		<tr> 
			<td colspan="2">&nbsp;</td>
		</tr>
		<tr> 
			<td valign="top"> 	
				<!--Choices start here -->
				<table width="100%">								
					<% if gwpp<>"1" then %>
					<tr> 
						<td width="6%" height="21"></td>
						<td height="21"><a href="AddPayPalPaymentOpt.asp?gwchoice=PayPal&ppv=0"><strong>PayPal Website Payments Standard</strong></a></td>
					</tr>
					<% end if %>
					
					<% if gwpp<>"1" then %>
					<tr> 
						<td width="6%" height="21"></td>
						<td height="21"><a href="AddPayPalPaymentOpt.asp?gwchoice=PayPal&ppv=1"><strong>Express Checkout</strong></a></td>
					</tr>
					<% end if %>
					
					<% if gwpp<>"1" then %>
					<tr> 
						<td colspan="2" width="6%" height="21"><hr /></td>
					</tr>
					<tr> 
						<td width="6%" height="21"></td>
						<td height="21"><a href="AddPayPalPaymentOpt.asp?gwchoice=PayPal&ppv=2"><strong>Website Payments Pro (US)</strong></a></td>
					</tr>
					<% end if %>
					
					<% if gwpp<>"1" then %>
					<tr> 
						<td colspan="2" width="6%" height="21"><hr /></td>
					</tr>
					<tr> 
						<td width="6%" height="21"></td>
						<td height="21"><a href="AddPayPalPaymentOpt.asp?gwchoice=PayPal&ppv=3"><strong>Express Checkout (UK)</strong></a></td>
					</tr>
					<% end if %>
					
					<% if gwpp<>"1" then %>
					<tr> 
						<td width="6%" height="21"></td>
						<td height="21"><a href="AddPayPalPaymentOpt.asp?gwchoice=PayPal&ppv=4"><strong>Website Payments Pro (UK)</strong></a></td>
					</tr>
					<% end if %>
					
					<% if gwpp<>"1" AND gwvpfp<>"1" then %>
					<tr> 
						<td colspan="2" width="6%" height="21"><hr /></td>
					</tr>
					<% end if %>
					
					<% if gwvpfp<>"1" then %>
					<tr> 
						<td width="6%" height="21"></td>
						<td height="21"><a href="AddPayPalPaymentOpt.asp?gwchoice=VeriSignPP#GWA"><strong>PayPal Payflow Pro</strong></a></td>
					</tr>
					<% end if %>
					
					<% if gwvpfl<>"1" then %>
					<tr> 
						<td width="6%" height="21"></td>
						<td height="21"><a href="AddPayPalPaymentOpt.asp?gwchoice=VeriSignLk#GWA"><strong>PayPal Payflow Link</strong></a></td>
					</tr>
					<% end if %>

					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>	
				</table>
				<!-- Choices end here -->
			</td>
		</tr>	
	</table>
	<% end if %>
<% 
else '// If gwChoice is empty, hide the rest of the page 
%>
		<% if NOT actGW=1 then %><div style="float: right; margin: 10px 25px 0 0;"><a href="AdminPaymentOptions.asp">Return to Payment Options</a></div><% end if %>
       	<table class="pcCPcontent">
		<tr> 
			<td colspan="2"> 			
				
				
				<!-- Main Config Table start here -->
				<table width="100%">		

					<!-- START PAYPAL -->
					<% if gwpp<>"1" AND request("gwchoice")="PayPal" then %>
					<% pcv_strIsGateway = 1 '// A Gateway is Active %>
					<tr> 
						<th colspan="2"><input name="gwpp" type="checkbox" class="clearBorder" value="1" <% if request("gwchoice")="PayPal" then%>Checked<%end if%>> 
							<a name="GWA"></a>Enable PayPal Settings - <a href="https://www.paypal.com/cgi-bin/webscr?cmd=_profile-comparison" target="_blank">More Information</a>
                        </th>
					</tr>
                    <tr>
                        <td colspan="2" class="pcCPspacer"></td>
                    </tr>
					<tr> 
						<td colspan="2">
							
                            
                            <% If (Request("ppv") = "0" AND Request("ppv") <> "") OR Request("gwCode")="3" Then %>
                            <div class="pcCPmessage">
                            		<strong>PayPal Website Payments Standard</strong> is a fast, affordable way to start accepting credit cards and PayPal payments online. Your buyers pay on a secure PayPal page, but do not need a PayPal account to pay you (<a href="http://www.paypal-marketing.com/html/partner/portal/standard.html" target="_blank">View a Demo</a>).
                                    <br /><br />                            
                            		To use <strong>PayPal Payments Standard</strong> you must modify settings inside your PayPal account. <a href="http://wiki.earlyimpact.com/productcart/paypal_standard" target="_blank">See the documentation</a> for instructions.
                            </div>         
						  	<% End If %>
                            
                            <% If (Request("ppv") = "2") OR (Request("gwCode")="46") Then %>
                            <div class="pcCPmessage">
                            		<strong>PayPal Website Payments Pro</strong> is an all-in-one solution that allows you to accept credit cards and PayPal. Buyers can enter their credit card numbers directly on your Web store (<a href="http://www.paypal-marketing.com/html/partner/portal/paymentspro.html" target="_blank">View a Demo</a>).
                                    <br /><br />
                          			Before activating <strong>PayPal Website Payments Pro</strong> please <a href="http://wiki.earlyimpact.com/productcart/paypal_pro" target="_blank">review the documentation</a>.
                          	</div> 
							<% End If %>   
                            
                               <% If (Request("ppv") = "1") Then %>
                            <div class="pcCPmessage">
                          			Before activating <strong>PayPal Express Checkout</strong> please <a href="http://wiki.earlyimpact.com/productcart/paypal_pro" target="_blank">review the documentation</a>.
                          	</div> 
							<% End If %> 
                            
                                                 
                          
                    <SCRIPT LANGUAGE="JavaScript" TYPE="text/JavaScript">
					<!--
					function whatPayPalSelected(){
						var selectValDom = document.forms['form1'];
						if (selectValDom.PayPalSelection[0].checked == true) {
						document.getElementById('PayPal_table').style.display='';
						//document.getElementById('chpaytype_table').style.display='none';
						}else{
						document.getElementById('PayPal_table').style.display='none';
						}
						
						if (selectValDom.PayPalSelection[1].checked == true) {
						document.getElementById('PayPalExp_table').style.display='';
						//document.getElementById('chpaytype_table').style.display='none';
						}else{
						document.getElementById('PayPalExp_table').style.display='none';
						}
						
						if (selectValDom.PayPalSelection[2].checked == true) {
						document.getElementById('PayPalWP_table').style.display='';
						//document.getElementById('chpaytype_table').style.display='';
						}else{
						document.getElementById('PayPalWP_table').style.display='none';
						}
						
						if (selectValDom.PayPalSelection[3].checked == true) {
						document.getElementById('PayPalExpUK_table').style.display='';
						//document.getElementById('chpaytype_table').style.display='';
						}else{
						document.getElementById('PayPalExpUK_table').style.display='none';
						}
						
						if (selectValDom.PayPalSelection[4].checked == true) {
						document.getElementById('PayPalWPUK_table').style.display='';
						//document.getElementById('chpaytype_table').style.display='';
						}else{
						document.getElementById('PayPalWPUK_table').style.display='none';
						}					
					}					
					function expandWhatSelected(tID){						
						var selectValDom = document.forms['form1'];
						selectValDom.PayPalSelection[tID].checked = true;
						whatPayPalSelected()
					}
					 //-->
					</SCRIPT>
                          					
						</td>
					</tr>

                    <% 
					'// PAYPAL PAYMENT STANDARD - START
					if (ucase(request("gwchoice"))="PAYPAL" and request("ppv")="0") or gwcode="3" then
					%>
					<tr> 
						<td valign="top">
                        	<div align="right"> 
							<input type="radio" class="clearBorder" name="PayPalSelection" value="PayPal" onClick="whatPayPalSelected();" checked>						
                            </div>
                        </td>
						<td><strong>PayPal Website Payments Standard</strong> - <a href="https://www.paypal.com/cgi-bin/webscr?cmd=_profile-comparison" target="_blank" class="pcSmallText">Compare Solutions</a><br>
								Accept credit/debit cards and PayPal. PayPal payments via PayPal Web site: Customers shop on your web site and pay on the PayPal web site. No merchant account required. </td>
					</tr>
                    <%
					else
					%>
					<tr><td style="height: 1px; padding: 0px;"><div style="display: none;"><input type="radio" name="PayPalSelection" value="PayPal"></div></td></tr>
                    <%
					end if
					'// PAYPAL PAYMENT STANDARD - END

					'// PAYPAL EXPRESS - START
					if (ucase(request("gwchoice"))="PAYPAL" and request("ppv")="1") or gwcode="999999" then
					%>

			  		<tr> 
						<td valign="top">
                        <div align="right"> 
							<input type="radio" class="clearBorder" name="PayPalSelection" value="PayPalExp" onClick="whatPayPalSelected();" checked>
						</div>
                        </td>
						<td><strong>Express Checkout</strong> (Alternative Checkout Process) - <a href="https://www.paypal.com/cgi-bin/webscr?cmd=_profile-comparison" target="_blank" class="pcSmallText">Compare Solutions</a>
                        <br>
                        Accept credit/debit cards and PayPal using a Business or Premier PayPal account.  Customers benefit from a rapid checkout process without the need to enter their shipping address, resulting in fewer abandoned carts. Alternative payment solution: Customers shop on your web site and pay on PayPal without creating an account in your webstore. No merchant account required.
                       	</td>
					</tr>
                    
                    <%
					else
					%>
                    <tr><td style="height: 1px; padding: 0px;"><div style="display: none;"><input type="radio" name="PayPalSelection" value="PayPalExp"></div></td></tr>
                    <%
					end if
					'// PAYPAL EXPRESS - END

					'// PAYPAL WPP - START
					if (ucase(request("gwchoice"))="PAYPAL" and request("ppv")="2") or gwcode="46" then
					%>

					<tr> 
						<td valign="top">
                        	<div align="right"> 
							<input type="radio" class="clearBorder" name="PayPalSelection" value="PayPalWP" onClick="whatPayPalSelected();" checked>
							</div>
                        </td>
						<td><strong>Website Payment Pro</strong> - <a href="https://www.paypal.com/cgi-bin/webscr?cmd=_wp-pro-overview-outside" target="_blank">Solution Overview</a><br>
                        Includes both Express Checkout and Direct Payment - with Direct Payment the customer shops and pays on your website. PayPal merchant account required.
                        </td>						
					</tr>
                    <%
					else
					%>
                    <tr><td style="height: 1px; padding: 0px;"><div style="display: none;"><input type="radio" name="PayPalSelection" value="PayPalWP"></div></td></tr>
                    <%
					end if
					'// PAYPAL WPP - END

					'// PAYPAL EXPRESS UK - START
					if (ucase(request("gwchoice"))="PAYPAL" and request("ppv")="3") or request("ppeSetup")="1" then
					%>
			  		<tr> 
						<td valign="top">
                        	<div align="right"> 
							<input type="radio" class="clearBorder" name="PayPalSelection" value="PayPalExpUK" onclick="whatPayPalSelected();">
							</div>
          				</td>
						<td><strong>Express Checkout (UK)</strong> (Alternative checkout process) - <a href="https://www.paypal.com/uk/cgi-bin/webscr?cmd=_profile-comparison" target="_blank" class="pcSmallText">Compare Solutions</a>
                        	<br>
							For UK merchants <u>only with a PayPal partner name of &quot;<strong>PayPalUK</strong>&quot;, merchant ID and username</u>. Accept credit/debit cards and PayPal using a Business or Premier PayPal account.  Customers benefit from a rapid checkout process without the need to enter their shipping address, resulting in fewer abandoned carts. Alternative payment solution: Customers shop on your web site and pay on PayPal without creating an account in your webstore. No merchant account required.
                            </td>
					</tr>
                    <%
					else
					%>
                    <tr><td style="height: 1px; padding: 0px;"><div style="display: none;"><input type="radio" name="PayPalSelection" value="PayPalExpUK"></div></td></tr>
                    <%
					end if
					'// PAYPAL WPP - END

					'// PAYPAL WPP UK - START
					if (ucase(request("gwchoice"))="PAYPAL" and request("ppv")="4") or gwcode="53" then
					%>
					<tr> 
						<td valign="top">
                        	<div align="right"> 
							<input type="radio" class="clearBorder" name="PayPalSelection" value="PayPalWPUK" onclick="whatPayPalSelected();">
							</div>
                      </td>
					  <td><strong>Website Payments Pro (UK)</strong> - <a href="https://www.paypal.com/uk/cgi-bin/webscr?cmd=_profile-comparison" target="_blank" class="pcSmallText">Compare Solutions</a>
                      		<br>
							For UK merchants <u>only with a PayPal partner name of &quot;<strong>PayPalUK</strong>&quot;, merchant ID and username</u>. Accept credit/debit cards on your web site, and payments from PayPal account holders using Express Checkout. PayPal merchant account required.
                       </td>
					</tr>
                    <%
					else
					%>
                    <tr><td style="height: 1px; padding: 0px;"><div style="display: none;"><input type="radio" name="PayPalSelection" value="PayPalWPUK"></div></td></tr>
                    <%
					end if
					'// PAYPAL WPP UK - END
					%>

					<tr> 
						<td colspan="2" valign="top"> 							
							
							<table ID="PayPal_table" width="100%" border="0" cellspacing="0" cellpadding="4" style="display:none">
								<tr> 
									<th colspan="2">Configure Settings - Website Payments Standard</th>
								</tr>
                                <tr>
                                    <td colspan="2" class="pcCPspacer"></td>
                                </tr>
								<tr> 
									<td width="127" valign="top"><div align="right">Email Address: </div></td>
									<td><input type="text" value="" name="PayPal_Email" size="30" maxlength="50"> 
									<input type="hidden" value="https://www.paypal.com/cgi-bin/webscr" name="URL2">
                                    </td>
								</tr>
								<tr> 
									<td valign="top"><div align="right">Currency:</div></td>
									<td> <select name="PayPal_Currency">
											<option value="USD">U.S. Dollars ($)</option>
                                            <option value="AUD">Australian Dollars ($)</option>
											<option value="CAD">Canadian Dollars (C $)</option>
                                            <option value="CZK">Czech Koruna</option>
                                            <option value="DKK">Danish Krone</option>                                            
											<option value="EUR">Euros (€)</option>
                                            <option value="HKD">Hong Kong Dollar</option>
                                            <option value="HUF">Hungarian Forint</option>
                                            <option value="ILS">Israeli New Shekel</option>
                                            <option value="JPY">Yen (¥)</option>
                                            <option value="MXN">Mexican Peso</option> 
                                            <option value="NOK">Norwegian Krone</option>
                                            <option value="NZD">New Zealand Dollar</option>
                                            <option value="PHP">Philippine Peso</option> 
                                            <option value="PLN">Polish Zloty</option>
											<option value="GBP">Pounds Sterling (£)</option>											
											<option value="SGD">Singapore Dollar</option>
                                            <option value="SEK">Swedish Krona</option>                                            
                                            <option value="CHF">Swiss Franc</option>     
                                            <option value="TWD">Taiwan New Dollar</option>    
                                            <option value="THB">Thai Baht</option>                               
										</select>
                                        </td>
								</tr>
								<tr> 
									<td valign="top"> <div align="right"> 
									<input name="PayPal_Sandbox" type="checkbox" class="clearBorder" value="YES" checked>
									</div></td>
									<td><b>Enable Test Mode </b>(Credit cards will not be charged)<div><font color="#FF0000">Visit <a href="https://developer.paypal.com" target="_blank">PayPal's Developer Site</a> to obtain "SandBox" (Test) credentials. <strong>Do not use your real account information.</strong></font></div></td>
								</tr>
							</table>										
							
							
							<table ID="PayPalWP_table" width="100%" border="0" cellspacing="0" cellpadding="4" style="display:none">
								<tr> 
									<th colspan="2">Configure Settings - Website Payments Pro (US)</th>
								</tr>	
                                <tr>
                                    <td colspan="2" class="pcCPspacer"></td>
                                </tr>
								<tr> 
									<td colspan="2">
                                    
                                    	<table>
                                            <tr> 
                                                <td width="127" valign="top" nowrap><div align="right">Email Address:</div></td>
                                                <td>
                                                	<input type="text" value="<%=pcPay_PayPal_Subject%>" name="pcPay_PayPal_Subject" size="30" maxlength="50">
                                                </td>
                                            </tr>
                                            <tr> 
                                                <td width="127" valign="top" nowrap></td>
                                                <td>
                                                	<div align="left">
                                                    	This is the email address to receive PayPal payment.                                        
                                                    </div>
                                                </td>
                                            </tr>
                                        </table>
                                    
                                    </td>
								</tr>
								<tr> 
									<td width="127" valign="top" nowrap><div align="right">API Credentials: <br /> (can be added later)</div></td>
                                    <td>
                                    
                                    	<table style="border:1px #CCC dashed">
                                            <tr> 
                                                <td width="127" valign="top" nowrap><div align="right">API Account Name:</div></td>
                                                <td><input type="text" value="<%=pcPay_PayPal_Username%>" name="pcPay_PayPal_Username" size="30" maxlength="50"></td>
                                                <td>&nbsp;</td>
                                                <td rowspan="3" valign="top">
                                                    <div align="center">
                                                        <div class="pcCPnotes">
                                                            API Credentials are required for direct payments and post-checkout operations (for example, Authorization & Capture and Refund) 
                                                        </div>
                                                    </div>                                             
                                                </td>
                                            </tr>
                                            <tr> 
                                                <td valign="top"><div align="right">Password:</div></td>
                                                <td><input type="text" value="<%=pcPay_PayPal_Password%>" name="pcPay_PayPal_Password" size="30" maxlength="50"></td>
                                                <td>&nbsp;</td>
                                            </tr>
                                            <tr> 
                                                <td valign="top"><div align="right">API Signature:</div></td>
                                                <td>
                                                <input type="text" value="<%=pcPay_PayPal_Signature%>" name="pcPay_PayPal_Signature" size="30" maxlength="250"></td>
                                                <td>&nbsp;</td>
                                            </tr>
                                        </table>
                                    
                                    </td>
								</tr>
								<tr> 
									<td width="127" valign="top" nowrap><div align="right">Card Types:</div></td>
									<td>
										<% if V="1" then %>
                                            <input name="cardTypes" type="checkbox" class="clearBorder" id="cardTypes" value="V" checked> 
                                        <% else %>
                                            <input name="cardTypes" type="checkbox" class="clearBorder" id="cardTypes" value="V"> 
                                        <% end if %> Visa 
                                        <% if M="1" then %>
                                            <input name="cardTypes" type="checkbox" class="clearBorder" id="cardTypes" value="M" checked> 
                                        <% else %>
                                            <input name="cardTypes" type="checkbox" class="clearBorder" id="cardTypes" value="M"> 
                                        <% end if %> MasterCard 
                                        <% if A="1" then %>
                                            <input name="cardTypes" type="checkbox" class="clearBorder" id="cardTypes" value="A" checked> 
                                        <% else %>
                                            <input name="cardTypes" type="checkbox" class="clearBorder" id="cardTypes" value="A"> 
                                        <% end if %>  American Express 
                                        <% if D="1" then %>
                                            <input name="cardTypes" type="checkbox" class="clearBorder" id="cardTypes" value="D" checked> 
                                        <% else %>
                                            <input name="cardTypes" type="checkbox" class="clearBorder" id="cardTypes" value="D"> 
                                        <% end if %> Discover
                                	</td>
								</tr>
								<tr> 
									<td valign="top"> <div align="right">Currency:</div></td>
									<td> <select name="pcPay_PayPal_Currency">
											<option value="USD">U.S. Dollars ($)</option>
											<option value="AUD">Australian Dollars ($)</option>
											<option value="CAD">Canadian Dollars (C $)</option>
											<option value="EUR">Euros (&euro;)</option>
											<option value="GBP">Pounds Sterling (&pound;)</option>
											<option value="JPY">Yen (&yen;)</option>														
										</select>
                                    </td>
								</tr>
								<tr> 
									<td valign="top"> <div align="right">Transaction Type:</div></td>
									<td> 
                                    	<select name="pcPay_PayPal_TransType">
                                        	<option value="1">Sale (Authorize and Capture)</option>	
											<option value="2">Authorize Only</option>																								
										</select>
                                        &nbsp;
                                        <span class="pcCPnotes"><strong>NOTE:</strong> &quot;Sale&quot; will be used when API Credentials are unavailable.</span>
                                    </td>
								</tr>
								<tr> 
									<td></td>
									<td>
									<div style="padding-bottom:4px">
										<span class="pcCPsectionTitle">Transaction Type tells PayPal how you want to obtain payment.</span>												</div>
									<li><strong>Sale</strong> indicates that authorization and capture occur at the same time. The sale is final. For more information, please consult the ProductCart User Guide.</li>
									<li><strong>Authorize Only</strong> indicates that the transaction is only authorized, but not settled. The payment will be captured at a later time.</li>											  </td>
								</tr>
								<tr> 
									<td valign="top"> <div align="right"> 
									<input name="pcPay_PayPal_Sandbox" type="checkbox" class="clearBorder" value="YES">
									</div></td>
									<td><b>Enable Test Mode </b>(Credit cards will not be charged)<div><font color="#FF0000">Visit <a href="https://developer.paypal.com" target="_blank">PayPal's Developer Site</a> to obtain "SandBox" (Test) credentials. <strong>Do not use your real account information.</strong></font></div></td>
								</tr>
								<tr> 
									<td valign="top"> <div align="right"> 
									<input name="pcPay_PayPal_CVC" type="checkbox" class="clearBorder" value="1" checked>
									</div></td>
									<td><b>Enable Credit Card Security Code</b> 
									<div>Every Credit Card has a 3 or 4 digit CVV Code (also known as CVV2), which is a security code designed for your safety and security. Check the box above and customers are required to enter their Credit Card's CVV Code. Note: You must also enable CVV in your PayPal Account.</div>												</td>
								</tr>
							</table>
							
							
							<table ID="PayPalExp_table" width="100%" border="0" cellspacing="0" cellpadding="4" style="display:none">
								<tr> 
									<th colspan="2">Configure Settings - Express Checkout</th>
								</tr>	
                                <tr>
                                    <td colspan="2" class="pcCPspacer"></td>
                                </tr>
								<tr> 
									<td colspan="2">
                                    
                                    	<table>
                                            <tr> 
                                                <td width="127" valign="top" nowrap><div align="right">Email Address:</div></td>
                                                <td>
                                                	<input type="text" value="<%=pcPay_PayPal_Subject%>" name="pcPay_PayPale_Subject" size="30" maxlength="50">
                                                </td>
                                            </tr>
                                            <tr> 
                                                <td width="127" valign="top" nowrap></td>
                                                <td>
                                                	<div align="left">
                                                    	This is the email address to receive PayPal payment.                                        
                                                    </div>
                                                </td>
                                            </tr>
                                        </table>
                                    
                                    </td>
								</tr>
								<tr> 
									<td width="127" valign="top" nowrap><div align="right">API Credentials: <br /> (can be added later)</div></td>
                                    <td>
                                    
                                    	<table style="border:1px #CCC dashed">
                                            <tr> 
                                                <td width="127" valign="top" nowrap><div align="right">API Account Name:</div></td>
                                                <td><input type="text" value="<%=pcPay_PayPal_Username%>" name="pcPay_PayPale_Username" size="30" maxlength="50"></td>
                                                <td>&nbsp;</td>
                                                <td rowspan="3" valign="top">
                                                    <div align="center">
                                                        <div class="pcCPnotes">
                                                            API Credentials are required for direct payments and post-checkout operations (for example, Authorization & Capture and Refund) 
                                                        </div>
                                                    </div>                                             
                                                </td>
                                            </tr>
                                            <tr> 
                                                <td valign="top"><div align="right">Password:</div></td>
                                                <td><input type="text" value="<%=pcPay_PayPal_Password%>" name="pcPay_PayPale_Password" size="30" maxlength="50"></td>
                                                <td>&nbsp;</td>
                                            </tr>
                                            <tr> 
                                                <td valign="top"><div align="right">API Signature:</div></td>
                                                <td>
                                                <input type="text" value="<%=pcPay_PayPal_Signature%>" name="pcPay_PayPale_Signature" size="30" maxlength="250"></td>
                                                <td>&nbsp;</td>
                                            </tr>
                                        </table>
                                    
                                    </td>
								</tr>
								<tr> 
									<td valign="top"> <div align="right">Currency:</div></td>
									<td> <select name="pcPay_PayPale_Currency">
											<option value="USD">U.S. Dollars ($)</option>
											<option value="AUD">Australian Dollars ($)</option>
											<option value="CAD">Canadian Dollars (C $)</option>
											<option value="EUR">Euros (&euro;)</option>
											<option value="GBP">Pounds Sterling (&pound;)</option>
											<option value="JPY">Yen (&yen;)</option>														
										</select>
                                    </td>
								</tr>
								<tr> 
									<td valign="top"> <div align="right">Transaction Type:</div></td>
									<td> 
                                    	<select name="pcPay_PayPale_TransType">
											<option value="1">Sale (Authorize and Capture)</option>
                                            <option value="2">Authorize Only</option>																									
										</select>
                                        &nbsp;
                                        <span class="pcCPnotes"><strong>NOTE:</strong> &quot;Sale&quot; will be used when API Credentials are unavailable.</span>
                                    </td>
								</tr>
								<tr> 
									<td></td>
									<td>
									<div style="padding-bottom:4px">
										<span class="pcCPsectionTitle">Transaction Type tells PayPal how you want to obtain payment.</span>												</div>
									<li><strong>Sale</strong> indicates that authorization and capture occur at the same time. The sale is final. For more information, please consult the ProductCart User Guide.</li>
									<li><strong>Authorize Only</strong> indicates that the transaction is only authorized, but not settled. The payment will be captured at a later time.</li>
                                 </td>
								</tr>
								<tr> 
									<td valign="top"> <div align="right"> 
									<input name="pcPay_PayPale_Sandbox" type="checkbox" class="clearBorder" value="YES">
									</div></td>
									<td>
									<b>Enable Test Mode </b>(Credit cards will not be charged)
									<div>
									<font color="#FF0000">
									Visit <a href="https://developer.paypal.com" target="_blank">PayPal's Developer Site</a> to obtain "SandBox" (Test) credentials. <strong>
									Do not use your real account information.</strong></font></div></td>
								</tr>
							</table>								
							
							
							<table ID="PayPalWPUK_table" width="100%" border="0" cellspacing="0" cellpadding="4" style="display:none">
								<tr> 
									<th colspan="2">Configure Settings - Website Payments Pro (UK)</th>
								</tr>	
                                <tr>
                                    <td colspan="2" class="pcCPspacer"></td>
                                </tr>
								<tr> 
									<td width="127" valign="top"><div align="right">User Name:</div></td>
									<td><input type="text" value="<%=Session("pcAdminUserName")%>" name="pcPay_PayPalUK_Username" size="30" maxlength="50"></td>
								</tr>
								
								<tr> 
									<td valign="top"><div align="right">Password:</div></td>
									<td><input type="text" value="<%=Session("pcAdminPassword")%>" name="pcPay_PayPalUK_Password" size="30" maxlength="50"></td>
								</tr>						  
								<tr> 
									<td width="127" valign="top"><div align="right">Partner:</div></td>
									<td>
									<input type="text" value="<%=Session("pcAdminPartner")%>" name="pcPay_PayPalUK_Partner" size="30" maxlength="250">									</td>
								</tr>
								<tr> 
									<td width="127" valign="top"><div align="right">Vendor:</div></td>
									<td>
									<input type="text" value="<%=Session("pcAdminVendor")%>" name="pcPay_PayPalUK_Vendor" size="30" maxlength="250">									</td>
								</tr>
								<tr> 
									<td width="127" valign="top" nowrap><div align="right">Card Types:</div></td>
									<td>
										<% if V="1" then %>
                                            <input name="cardTypesUK" type="checkbox" class="clearBorder" id="cardTypes" value="V" checked> 
                                        <% else %>
                                            <input name="cardTypesUK" type="checkbox" class="clearBorder" id="cardTypes" value="V"> 
                                        <% end if %> Visa 
                                        <% if M="1" then %>
                                            <input name="cardTypesUK" type="checkbox" class="clearBorder" id="cardTypes" value="M" checked> 
                                        <% else %>
                                            <input name="cardTypesUK" type="checkbox" class="clearBorder" id="cardTypes" value="M"> 
                                        <% end if %> MasterCard 
                                        <% if A="1" then %>
                                            <input name="cardTypesUK" type="checkbox" class="clearBorder" id="cardTypes" value="A" checked> 
                                        <% else %>
                                            <input name="cardTypesUK" type="checkbox" class="clearBorder" id="cardTypes" value="A"> 
                                        <% end if %>  American Express 
                                        <% if D="1" then %>
                                            <input name="cardTypesUK" type="checkbox" class="clearBorder" id="cardTypes" value="D" checked> 
                                        <% else %>
                                            <input name="cardTypesUK" type="checkbox" class="clearBorder" id="cardTypes" value="D"> 
                                        <% end if %> Discover
                                	</td>
								</tr>
								<tr> 
									<td valign="top"> <div align="right">Currency:</div></td>
									<td> <select name="pcPay_PayPalUK_Currency">
											<option value="GBP">Pounds Sterling (&pound;)</option>
											<option value="USD">U.S. Dollars ($)</option>
											<option value="AUD">Australian Dollars ($)</option>
											<option value="CAD">Canadian Dollars (C $)</option>
											<option value="EUR">Euros (&euro;)</option>														
											<option value="JPY">Yen (&yen;)</option>														
										</select>												</td>
								</tr>
								<tr> 
									<td valign="top"> <div align="right">Transaction Type:</div></td>
									<td> <select name="pcPay_PayPalUK_TransType">
											<option value="2">Authorize Only</option>
											<option value="1">Sale (Authorize and Capture)</option>														
										</select></td>
								</tr>
								<tr> 
									<td></td>
									<td>
									<div style="padding-bottom:4px">
										<span class="pcCPsectionTitle">Transaction Type tells PayPal how you want to obtain payment.</span>												</div>
									<li><strong>Sale</strong> indicates that authorization and capture occur at the same time. The sale is final. For more information, please consult the ProductCart User Guide.</li>
									<li><strong>Authorize Only</strong> indicates that the transaction is only authorized, but not settled. The payment will be captured at a later time.</li>											  </td>
								</tr>
								<tr> 
									<td valign="top"> <div align="right"> 
									<input name="pcPay_PayPalUK_Sandbox" type="checkbox" class="clearBorder" value="YES">
									</div></td>
									<td><b>Enable Test Mode </b>(Credit cards will not be charged)<div><font color="#FF0000">Visit <a href="https://developer.paypal.com" target="_blank">PayPal's Developer Site</a> to obtain "SandBox" (Test) credentials. <strong>Do not use your real account information.</strong></font></div></td>
								</tr>
								<tr> 
									<td valign="top"> <div align="right"> 
									<input name="pcPay_PayPalUK_CVC" type="checkbox" class="clearBorder" value="1" checked>
									</div></td>
									<td><b>Enable Credit Card Security Code</b> 
									<div>Every Credit Card has a 3 or 4 digit CVV Code (also known as CVV2), which is a security code designed for your safety and security. Check the box above and customers are required to enter their Credit Card's CVV Code. Note: You must also enable CVV in your PayPal Account.</div>												</td>
								</tr>

								<% if gwpp<>"1" AND request("gwchoice")="PayPal" then %>
                                    <tr>
                                        <td colspan="2" class="pcCPspacer"></td>
                                    </tr>
                                    <tr> 
                                        <th colspan="2">Minimize fraud by enabling Centinel by CardinalCommerce</th>
                                    </tr>
                                    <tr>
                                        <td colspan="2" class="pcCPspacer"></td>
                                    </tr>
                                    <tr> 
                                        <td colspan="2">Note: Additional charges apply. <a href="http://billing.cardinalcommerce.com/centinel/registration/frame_services.asp?RefId=PRDCTCART" target="_blank">Contact CardinalCommerce for more information &gt;&gt;</a></td>
                                    </tr> 
                                    
                                    
                                    <% if pcPay_Cent_Active=1 then %>
                                        <tr>
                                            <td colspan="2">Centinel has already been activated for this or another payment gateway. To edit its settings or remove it, simply activate this payment gateway and then click on the &quot;Modify&quot; button on the payment options summary page.</td>
                                        </tr>
                                    <% else %>
                                        <tr> 
                                            <td><div align="right"> 
                                                <input name="pcPay_Cent_Active" type="checkbox" class="clearBorder" value="YES" <%if pcPay_Cent_Active=1 then%>checked<%end if%>></div></td>
                                            <td><strong>Enable Centinel for PayPal Website Payments Pro UK</strong></td>
                                        </tr>
                                        <% if trim(pcPay_Cent_TransactionURL)="" then
                                            pcPay_Cent_TransactionURL="https://centineltest.cardinalcommerce.com/maps/txns.asp"
                                        end if %>
                                        <tr> 
                                            <td><div align="right">Transaction Url:</div></td>
                                            <td><input name="pcPay_Cent_TransactionURL" size="60" maxlength="255" value="<%=pcPay_Cent_TransactionURL%>"></td>
                                        </tr>
                                        <% if pcPay_Cent_MerchantID<>"" then
                                            pcPay_Cent_MerchantID=replace(pcPay_Cent_MerchantID,"""","&quot;")
                                        end if %>
                                        <tr> 
                                            <td><div align="right">Merchant ID: </div></td>
                                            <td><input name="pcPay_Cent_MerchantID" size="35" maxlength="255" value="<%=pcPay_Cent_MerchantID%>"></td>
                                        </tr>
                                        <% if pcPay_Cent_ProcessorID<>"" then
                                            pcPay_Cent_ProcessorID=replace(pcPay_Cent_ProcessorID,"""","&quot;")
                                        end if %>
                                        <tr> 
                                            <td><div align="right">Processor ID: </div></td>
                                            <td><input name="pcPay_Cent_ProcessorId" size="35" maxlength="255" value="<%=pcPay_Cent_ProcessorID%>"></td>
                                        </tr>
                                        <% if pcPay_Cent_Password<>"" then
                                            pcPay_Cent_Password=replace(pcPay_Cent_Password,"""","&quot;")
                                        end if %>
                                        <tr> 
                                            <td><div align="right">Password: </div></td>
                                            <td><input name="pcPay_Cent_Password" size="35" maxlength="255" value="<%=pcPay_Cent_Password%>"></td>
                                        </tr>
                                    <% end if %>
								<% end if '// if gwpp<>"1" AND request("gwchoice")="PayPal" then %>
							</table>
							
							
							<table ID="PayPalExpUK_table" width="100%" border="0" cellspacing="0" cellpadding="4" style="display:none">
								<tr> 
									<th colspan="2">Configure Settings - Express Checkout (UK)</th>
								</tr>
                                <tr>
                                    <td colspan="2" class="pcCPspacer"></td>
                                </tr>	
								<tr> 
									<td width="127" valign="top"><div align="right">User Name:</div></td>
									<td><input type="text" value="<%=Session("pcAdminUserName")%>" name="pcPay_PayPalUKe_Username" size="30" maxlength="50"></td>
								</tr>
								
								<tr> 
									<td valign="top"><div align="right">Password:</div></td>
									<td><input type="text" value="<%=Session("pcAdminPassword")%>" name="pcPay_PayPalUKe_Password" size="30" maxlength="50"></td>
								</tr>						  
								<tr> 
									<td width="127" valign="top"><div align="right">Partner:</div></td>
									<td>
									<input type="text" value="<%=Session("pcAdminPartner")%>" name="pcPay_PayPalUKe_Partner" size="30" maxlength="250">									</td>
								</tr>
								<tr> 
									<td width="127" valign="top"><div align="right">Vendor:</div></td>
									<td>
									<input type="text" value="<%=Session("pcAdminVendor")%>" name="pcPay_PayPalUKe_Vendor" size="30" maxlength="250">									</td>
								</tr>
								<tr> 
									<td width="127" valign="top" nowrap><div align="right">Card Types:</div></td>
									<td>
										<% if V="1" then %>
                                            <input name="cardTypesUKe" type="checkbox" class="clearBorder" id="cardTypes" value="V" checked> 
                                        <% else %>
                                            <input name="cardTypesUKe" type="checkbox" class="clearBorder" id="cardTypes" value="V"> 
                                        <% end if %> Visa 
                                        <% if M="1" then %>
                                            <input name="cardTypesUKe" type="checkbox" class="clearBorder" id="cardTypes" value="M" checked> 
                                        <% else %>
                                            <input name="cardTypesUKe" type="checkbox" class="clearBorder" id="cardTypes" value="M"> 
                                        <% end if %> MasterCard 
                                        <% if A="1" then %>
                                            <input name="cardTypesUKe" type="checkbox" class="clearBorder" id="cardTypes" value="A" checked> 
                                        <% else %>
                                            <input name="cardTypesUKe" type="checkbox" class="clearBorder" id="cardTypes" value="A"> 
                                        <% end if %>  American Express 
                                        <% if D="1" then %>
                                            <input name="cardTypesUKe" type="checkbox" class="clearBorder" id="cardTypes" value="D" checked> 
                                        <% else %>
                                            <input name="cardTypesUKe" type="checkbox" class="clearBorder" id="cardTypes" value="D"> 
                                        <% end if %> Discover
                                	</td>
								</tr>
								<tr> 
									<td valign="top"> <div align="right">Currency:</div></td>
									<td> <select name="pcPay_PayPalUKe_Currency">
											<option value="GBP">Pounds Sterling (&pound;)</option>
											<option value="USD">U.S. Dollars ($)</option>
											<option value="AUD">Australian Dollars ($)</option>
											<option value="CAD">Canadian Dollars (C $)</option>
											<option value="EUR">Euros (&euro;)</option>														
											<option value="JPY">Yen (&yen;)</option>														
										</select>												</td>
								</tr>
								<tr> 
									<td valign="top"> <div align="right">Transaction Type:</div></td>
									<td> <select name="pcPay_PayPalUKe_TransType">
											<option value="2">Authorize Only</option>
											<option value="1">Sale (Authorize and Capture)</option>														
										</select></td>
								</tr>
								<tr> 
									<td></td>
									<td>
									<div style="padding-bottom:4px">
										<span class="pcCPsectionTitle">Transaction Type tells PayPal how you want to obtain payment.</span>												</div>
									<li><strong>Sale</strong> indicates that authorization and capture occur at the same time. The sale is final. For more information, please consult the ProductCart User Guide.</li>
									<li><strong>Authorize Only</strong> indicates that the transaction is only authorized, but not settled. The payment will be captured at a later time.</li>												</td>
								</tr>
								<tr> 
									<td valign="top"> <div align="right"> 
									<input name="pcPay_PayPalUKe_Sandbox" type="checkbox" class="clearBorder" value="YES">
									</div></td>
									<td>
									<b>Enable Test Mode </b>(Credit cards will not be charged)
									<div>
									<font color="#FF0000">
									Visit <a href="https://developer.paypal.com" target="_blank">PayPal's Developer Site</a> to obtain "SandBox" (Test) credentials. <strong>
									Do not use your real account information.</strong></font></div></td>
								</tr>
								<tr> 
									<td valign="top"> <div align="right"> 
									<input name="pcPay_PayPalUKe_CVC" type="checkbox" class="clearBorder" value="1" checked>
									</div></td>
									<td><b>Enable Credit Card Security Code</b> 
									<div>Every Credit Card has a 3 or 4 digit CVV Code (also known as CVV2), which is a security code designed for your safety and security. Check the box above and customers are required to enter their Credit Card's CVV Code. Note: You must also enable CVV in your PayPal Account.</div>												</td>
								</tr>
							</table>							
												
						</td>
					</tr>
						
						<%
						pcv_strPayPalVersion = Request("ppv")
						if isNumeric(pcv_strPayPalVersion)=False then pcv_strPayPalVersion = 0
						if pcv_strPayPalVersion <> "" then
						%>
						<SCRIPT LANGUAGE="JavaScript" TYPE="text/JavaScript">
						<!--
							expandWhatSelected(<%=pcv_strPayPalVersion%>)
						 //-->
						</SCRIPT>
						<% end if %>
						
					<% end if %>
					<!-- END PAYPAL -->
					
					<!-- START PAYFLOW LINK -->
						
					<% 'check if Centinel has previously been activated.
					dim intCentActive
					intCentActive=0
					
					err.clear
					err.number=0
					call openDb()  
	
					query="SELECT pcPay_Cent_Active FROM pcPay_Centinel WHERE pcPay_Cent_ID=1;"
					set rs=Server.CreateObject("ADODB.Recordset")     
					set rs=connTemp.execute(query)
					pcPay_Cent_Active=rs("pcPay_Cent_Active")
					if pcPay_Cent_Active=1 then
						intCentActive=1
					end if
					
					set rs=nothing
					
					call closedb()
					%>
					<% if gwvpfl<>"1" AND request("gwchoice")="VeriSignLk" then %>
									<% pcv_strIsGateway = 1 '// A Gateway is Active %>
					<tr> 
						<th colspan="2"> <input name="gwvpfl" type="checkbox" class="clearBorder" value="1" <% if request("gwchoice")="VeriSignLk" then%>Checked<%end if%>> 
							<a name="GWA"></a>Enable PayPal Payflow Link - <a href="https://www.paypal.com/cgi-bin/webscr?cmd=_profile-comparison" target="_blank">More Information</a></th>
					</tr>
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>	
					<tr> 						
						<td colspan="2"> 
						
							<table width="100%" border="0" cellspacing="0" cellpadding="4">
								<tr> 
									<th colspan="2">Configure Settings - Payflow Link</th>
								</tr>
                                <tr>
                                    <td colspan="2" class="pcCPspacer"></td>
                                </tr>
								<tr> 
									<td>&nbsp;</td>
									<td> <input type="checkbox" class="clearBorder" name="pfl_testmode" value="YES"> 
										<b>Enable Test Mode </b>(Credit cards will not be charged)</td>
								</tr>
								<tr> 
									<td width="102"> <div align="right">Login:</div></td>
									<td width="475"> <input type="text" value="" name="v_User" size="24">								</td>
								</tr>
								<tr> 
									<td width="102"> <div align="right">Partner:</div></td>
									<td width="475"> <input type="text" value="VeriSign" name="v_Partner" size="24">								</td>
								</tr>
								<tr> 
									<td width="102"> <div align="right">Transaction Type:</div></td>
									<td width="475"> <select name="pfl_transtype">
											<option value="S">Sale</option>
											<option value="A" selected>Authorize Only</option>
										</select> </td>
								</tr>
								<tr> 
									<td> <div align="right">Require CSC:</div></td>
									<td> <input type="radio" class="clearBorder" name="pfl_CSC" value="YES">
										Yes 
										<input name="pfl_CSC" type="radio" class="clearBorder" value="NO" checked>
										No</td>
								</tr>
							</table>							
												
						</td>
					</tr>
					<% end if %>
					<!-- END PAYFLOW LINK -->




					<!-- START PAYFLOW PRO -->
					<% if gwvpfp<>"1" AND request("gwchoice")="VeriSignPP" then %>
					<% pcv_strIsGateway = 1 '// A Gateway is Active %>
									<tr> 
						<th colspan="2"> <input name="gwvpfp" type="checkbox" class="clearBorder" value="1" <% if request("gwchoice")="VeriSignPP" then%>Checked<%end if%>> 
							<a name="GWA"></a>Enable PayPal  Payflow Pro - <a href="https://www.paypal.com/cgi-bin/webscr?cmd=_profile-comparison" target="_blank">More information</a></th>
					</tr>
                    <tr>
                        <td colspan="2" class="pcCPspacer"></td>
                    </tr>
					<tr> 						
						<td colspan="2"> 
						
							<table width="100%" border="0" cellspacing="0" cellpadding="4">
								<tr> 
									<th colspan="2">Configure Settings - Payflow Pro</th>
								</tr>
                                <tr>
                                    <td colspan="2" class="pcCPspacer"></td>
                                </tr>
								<tr> 
									<td width="133"> <div align="right"></div></td>
									<td width="444"> <input type="checkbox" class="clearBorder" name="pfp_testmode" value="YES"> 
										<b>Enable Test Mode </b>(Credit cards will not be charged) 
										<input type="hidden" value="A" name="pfp_Type" size="24"> 
										<input type="hidden" value="TENDER=C&ZIP=12345&COMMENT1=ASP/COM Test Transaction" name="pfp_Tender" size="24">								</td>
								</tr>
								<tr> 
									<td width="133"> <div align="right">User:</div></td>
									<td width="444"> <input type="text" name="pfp_User" size="24" value="<%=session("adminv_User")%>">								</td>
								</tr>
								<tr> 
									<td width="133"> <div align="right">Partner:</div></td>
									<td width="444"> <% if  session("adminv_Partner")="" then
	session("adminv_Partner")="VeriSign"
	end if %> <input type="text" name="pfp_Partner" size="24" value="<%=session("adminv_Partner")%>">								</td>
								</tr>
								<tr> 
									<td width="133"> <div align="right">Password:</div></td>
									<td width="444"> <input type="password" name="pfp_Password" size="24" value="<%=session("adminv_Password")%>">								</td>
								</tr>
								<tr> 
									<td width="133"> <div align="right">Vendor:</div></td>
									<td width="444"> <input type="text" value="" name="pfp_Vendor" size="24">								</td>
								</tr>
								<tr> 
									<td width="133"> <div align="right">Transaction Type:</div></td>
									<td width="444"> <select name="pfp_transtype">
											<option value="S" selected>Sale</option>
											<option value="A"<% if session("adminv_Type")="A" then
											response.write " selected"
											end if %>
											>Authorize Only</option>
										</select> </td>
								</tr>
								<% session("adminv_Type")=""
								session("adminv_User")=""
								session("adminv_Partner")=""
								session("adminv_Password")=""
								session("adminv_Tender")=""
								session("adminpfp_testmode")="" %>
								<tr> 
									<td> <div align="right">Require CSC:</div></td>
									<td> <input type="radio" class="clearBorder" name="pfp_CSC" value="YES">
										Yes 
										<input name="pfp_CSC" type="radio" class="clearBorder" value="NO" checked>
										No
									</td>
								</tr>
							</table>							
												
						</td>
					</tr>
					<% end if %>
					<!-- END PAYFLOW PRO -->
				</table>
				<!-- Main Config Table end here -->
				
				
			</td>
		</tr>
		
	<% if pcv_strIsGateway = 1 then '// A Gateway is Active %>
		<tr>
			<th colspan="2">You have the option to charge a processing fee for this payment option.</th>
		</tr>
         <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
		<tr> 
			<td colspan="2"> 
					
				<% '// Processing Fees %>
				<table width="100%" border="0" cellspacing="0" cellpadding="4">
					 <tr>
						<% if priceToAdd <> "0" then %>
						<td width="133"> <div align="right">Processing fee: </div></td>
						<% end if %>
						<td> <input type="radio" class="clearBorder" name="priceToAddType" value="price">
							Flat fee&nbsp;&nbsp; &nbsp;<%=scCurSign%> <input name="priceToAdd" size="6" value="<%=money(0)%>">								</td>
					</tr>
					<tr> 
						<td>&nbsp;</td>
						<td> <input type="radio" class="clearBorder" name="priceToAddType" value="percentage">
							Percentage of Order Total&nbsp;&nbsp; &nbsp; % 
							<input name="percentageToAdd" size="6" value="0"> </td>
					</tr>
                    <tr>
                        <td colspan="2" class="pcCPspacer"></td>
                    </tr>
                    
                    <% if Request("ppv")<>1 then %>
					<tr>
						<th colspan="2">You can change the display name that is shown for this payment type. </th>
					</tr>
                    <tr>
                        <td colspan="2" class="pcCPspacer"></td>
                    </tr>
					<tr> 
						<td><div align="right">Payment Name:&nbsp;</div></td>
						<td><input name="paymentNickName" value="Credit Card" size="35" maxlength="255"></td>
					</tr>
                    <% end if %>
					
					<% if gwvpfp<>"1" AND request("gwchoice")="VeriSignPP" then %>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Minimize fraud by enabling Centinel by CardinalCommerce</th>
						</tr>
                        <tr>
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
						<tr> 
							<td colspan="2">Note: Additional charges apply. <a href="http://billing.cardinalcommerce.com/centinel/registration/frame_services.asp?RefId=PRDCTCART" target="_blank">Contact CardinalCommerce for more information &gt;&gt;</a></td>
						</tr>
						<% if intCentActive=0 then %>
							<tr> 
								<td><div align="right"> 
										<input name="pcPay_Cent_Active_pfp" type="checkbox" class="clearBorder" value="YES">
									</div></td>
								<td><strong>Enable Centinel for PayPal Payflow Pro</strong></td>
							</tr>
							<tr> 
								<td><div align="right">Transaction Url:</div></td>
								<td><input name="pcPay_Cent_TransactionURL_pfp" value="https://centineltest.cardinalcommerce.com/maps/txns.asp" size="60" maxlength="255"></td>
							</tr>
							<tr> 
								<td><div align="right">Merchant ID: </div></td>
								<td><input name="pcPay_Cent_MerchantID_pfp" size="35" maxlength="255"></td>
							</tr>
							<tr> 
								<td><div align="right">Processor ID: </div></td>
								<td><input name="pcPay_Cent_ProcessorId_pfp" size="35" maxlength="255"></td>
							</tr>
						<% else %>
							<tr>
								<td colspan="2">Centinel has already been activated for this or another payment gateway. To edit its settings or remove it, simply activate this payment gateway and then click on the &quot;Modify&quot; button on the payment options summary page.</td>
							</tr>
						<% end if %>
					
					<% end if '//  if gwvpfp<>"1" AND request("gwchoice")="VeriSignPP" then  %>
				</table>
				
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
			<td>&nbsp;</td>
			<td>Process orders when they are placed: <input type="checkbox" class="clearBorder" name="pcv_processOrder" value="1">&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=301')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
		</tr>
		<tr> 
			<td>&nbsp;</td>
			<td>When orders are placed, set the payment status to:
				<select name="pcv_setPayStatus">
					<option value="3" selected="selected">Default</option>
					<option value="0">Pending</option>
					<option value="1">Authorized</option>
					<option value="2">Paid</option>
				</select>
				&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=302')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>					</td>
		</tr>
		<%'End SDBA%>			
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
        <tr>
            <td colspan="2"><hr></td>
        </tr>
		<tr> 
			<td height="27" colspan="2" align="center">
				<input type="submit" value="Add Selected Options" name="Submit" class="submit2"> 
				&nbsp;
				<input type="button" value="Back" onclick="javascript:history.back()">					
			</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>	
	<% end if %>	

</table>
<% end if 'gwChoice is empty %>
</form>
<%
Session("pcAdminUserName")=""
Session("pcAdminPassword")=""
Session("pcAdminPartner")=""										
Session("pcAdminVendor")=""
%>
<!--#include file="AdminFooter.asp"-->
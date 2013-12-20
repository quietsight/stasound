<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Modify PayPal Payment Option" %>
<% section="paymntOpt" %>
<%PmAdmin=5%><!--#include file="adminv.asp"-->  
<!--#include file="../includes/settings.asp"-->
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


idPayment=Request.QueryString("id")
gwCode= Request.QueryString("gwCode")
iMode=Request.queryString("mode")

If sslUrl="" then
	sslUrl="Undefined"
End If

dim ppAB
ppAB=0 '// Default value for PayPal Accelerated Boarding

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

'Delete
If iMode="Del" Then
	call openDb()
	If request.QueryString("TYPE")="CC" then
		CCcode=request.queryString("CCCode")
		query="UPDATE CCTypes SET active=0 WHERE CCcode='" & CCcode & "'"
		set rs=Server.CreateObject("ADODB.Recordset")  
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		query="SELECT * FROM CCTypes WHERE active=-1"
		set rs=Server.CreateObject("ADODB.Recordset")  
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		if rs.eof then
			query= "UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",active=0, paymentNickName='' WHERE gwCode=6"
		end if
	Else		
		query= "DELETE FROM payTypes WHERE gwCode="& gwCode			
	End If
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	set rs=nothing
	call Closedb()
	response.redirect "paymentOptions.asp"
End If

'Edit
If iMode="Edit" Then
	call opendb()
	query= "SELECT paymentDesc, priceToAdd, cvv, percentageToAdd, sslUrl, terms, CReq, Cprompt, Cbtob, paymentNickName, pcPayTypes_processOrder, pcPayTypes_setPayStatus FROM payTypes WHERE gwCode= "& gwCode &" AND idPayment= "& idPayment
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	
	if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	if NOT rs.eof then
		paymentDesc=rs("paymentDesc")
		priceToAdd=rs("priceToAdd")
		cvv=0
		percentageToAdd=rs("percentageToAdd")
		sslUrl=rs("sslUrl")
		terms=rs("terms")
		CReq=rs("CReq")
		Cprompt=rs("Cprompt")
		Cbtob=rs("Cbtob")
		paymentNickName=rs("paymentNickName")
		'Start SDBA
		pcv_processOrder=rs("pcPayTypes_processOrder")
		pcv_setPayStatus=rs("pcPayTypes_setPayStatus")
		if pcv_setPayStatus="" then
			pcv_setPayStatus="3"
		end if
		'End SDBA
	End If
	
	set rs=nothing

	
	select case gwCode


		case "2" 'VeriSign
			query= "SELECT v_Url,v_User,v_Partner,v_Password,v_Vendor,pfl_testmode,pfl_transtype,pfl_CSC FROM verisign_pfp where id=1"
			set rs=Server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)
			if err.number <> 0 then
				strErrorDescription=err.description
				set rs=nothing
				call closeDb()
				response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
			end If
			pfp_Url=rs("v_Url")
			pfp_User=rs("v_User")
			pfp_Partner=rs("v_Partner")
			pfp_Password=rs("v_Password")
			pfp_Vendor=rs("v_Vendor")
			pfp_testmode=rs("pfl_testmode")
			pfp_transtype=rs("pfl_transtype")
			pfp_CSC=rs("pfl_CSC") 
			set rs=nothing


		case "9" 'pfl
			query= "SELECT v_User,v_Partner,pfl_testmode,pfl_transtype,pfl_CSC FROM verisign_pfp where id=1"
			set rs=Server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)
			if err.number <> 0 then
				strErrorDescription=err.description
				set rs=nothing
				call closeDb()
				response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
			end If
			v_User=rs("v_User")
			v_Partner=rs("v_Partner")
			pfl_testmode=rs("pfl_testmode")
			pfl_transtype=rs("pfl_transtype")
			pfl_CSC=rs("pfl_CSC") 
			set rs=nothing			


		case "3" 'Paypal
			query= "SELECT Pay_To, PP_Currency, PP_Sandbox FROM paypal WHERE ID=1"
			set rs=Server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)
			if err.number <> 0 then
				strErrorDescription=err.description
				set rs=nothing
				call closeDb()
				response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
			end If
			PayPal_Email=rs("Pay_To")
			PayPal_Currency=rs("PP_Currency")
			PayPal_Sandbox=rs("PP_Sandbox")
			set rs=nothing



		case "46", "53", "999999" 'PayPalWP, PayPalExp
			query="SELECT pcPay_PayPal.pcPay_PayPal_TransType, pcPay_PayPal.pcPay_PayPal_Username, pcPay_PayPal.pcPay_PayPal_Password,  pcPay_PayPal.pcPay_PayPal_Sandbox, pcPay_PayPal.pcPay_PayPal_Signature, pcPay_PayPal.pcPay_PayPal_Currency, pcPay_PayPal.pcPay_PayPal_CVC, pcPay_PayPal.pcPay_PayPal_Partner, pcPay_PayPal.pcPay_PayPal_Vendor, pcPay_PayPal.pcPay_PayPal_Subject, pcPay_PayPal.pcPay_PayPal_CardTypes FROM pcPay_PayPal WHERE (((pcPay_PayPal.pcPay_PayPal_ID)=1));"
			set rs=Server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)
			if err.number <> 0 then
				strErrorDescription=err.description
				set rs=nothing
				call closeDb()
				response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
			end If
			pcPay_PayPal_TransType=rs("pcPay_PayPal_TransType")
			pcPay_PayPal_Username=rs("pcPay_PayPal_Username")
			pcPay_PayPal_Password=rs("pcPay_PayPal_Password")
			pcPay_PayPal_Sandbox=rs("pcPay_PayPal_Sandbox")	
			pcPay_PayPal_Signature=rs("pcPay_PayPal_Signature")	
			pcPay_PayPal_Currency=rs("pcPay_PayPal_Currency")
			pcPay_PayPal_CVC=rs("pcPay_PayPal_CVC")
			pcPay_PayPal_Partner=rs("pcPay_PayPal_Partner")
			pcPay_PayPal_Vendor=rs("pcPay_PayPal_Vendor")
			pcPay_PayPal_Subject=rs("pcPay_PayPal_Subject")
			pcPay_PayPal_CardTypes=rs("pcPay_PayPal_CardTypes")
			If len(pcPay_PayPal_Subject)=0 Then
				pcPay_PayPal_Subject=""
			End If
			if pcPay_PayPal_Partner<>"" AND pcPay_PayPal_Vendor<>"" then
				pcPay_PayPal_Version = "UK"			
			else
				pcPay_PayPal_Version = "US"						
			end if
			if IsNull(pcPay_PayPal_CardTypes) then pcPay_PayPal_CardTypes=""
			set rs=nothing
			
			
	end select
	call closedb()
End If
'end of Edit part

sMode=Request.Form("Submit")
If sMode <> "" Then
	PaymentDesc=request.Form("PaymentDesc")
	idPayment=request.Form("idPayment")
	priceToAddType=request.Form("priceToAddType")
	gwCode=request.form("gwCode")
	If priceToAddType="price" Then
		priceToAdd=replacecomma(Request("priceToAdd"))
		percentageToAdd="0"
		if priceToAdd="" then
			priceToAdd="0"
		end if
	Else
		percentageToAdd=request.Form("percentageToAdd")
		priceToAdd="0"
		if percentageToAdd="" then
			percentageToAdd="0"
		end if
	End If
	
	sslUrl=request.Form("sslUrl")
	paymentNickName=replace(request.Form("paymentNickName"),"'","''")
	
	call opendb()	

	select case gwCode

		case "3", "46", "53", "999999"
			'request gateway variables and insert them into the paypal table
			'check payment type
			PayPalSelection=request.Form("PayPalSelection")
			if PayPalSelection="PayPal" then
				PayPalPaymentURL="gwpp.asp"
				PayPalName="PayPal"
				PayPal_Email=request.Form("PayPal_Email")
				if PayPal_Email = "" then 
					response.redirect "modPayPalPaymentOpt.asp?Mode=Edit&id="&idPayment&"&gwCode="&gwCode&"&msg="&Server.URLEncode("You must enter your PayPal email address to activate PayPal Standard.")
				end if
				PayPal_Currency=request.Form("PayPal_Currency")	
				pcPay_PayPal_Sandbox=request.Form("PayPal_Sandbox")	
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
			
			If (pcPay_PayPal_Signature="" OR pcPay_PayPal_Username="" OR pcPay_PayPal_Password="") AND (ppGwCode=46 OR ppGwCode=999999) Then
				ppAB=1
			Else 
				ppAB=0
			End If
			
			err.clear
			err.number=0

			sslUrl = PayPalPaymentURL

			if PayPalSelection="PayPal" then
				query="UPDATE paypal SET Pay_To='"&PayPal_Email&"', PP_Currency='"&PayPal_Currency&"', PP_Sandbox="&pcPay_PayPal_Sandbox&" WHERE ID=1;"
				set rs=Server.CreateObject("ADODB.Recordset")   
				rs.Open query, conntemp
				if err.number <> 0 then
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddPaymentOpt: "&Err.Description) 
				end if
			else
				query="UPDATE pcPay_PayPal SET pcPay_PayPal_TransType="&pcPay_PayPal_TransType&", pcPay_PayPal_Username='"&pcPay_PayPal_Username&"', pcPay_PayPal_Subject='"&pcPay_PayPal_Subject&"', pcPay_PayPal_Password='"&pcPay_PayPal_Password&"', pcPay_PayPal_Sandbox="&pcPay_PayPal_Sandbox&", pcPay_PayPal_Signature='"&pcPay_PayPal_Signature&"', pcPay_PayPal_Currency='"&pcPay_PayPal_Currency&"', pcPay_PayPal_CVC="&pcPay_PayPal_CVC&", pcPay_PayPal_Vendor='"&pcPay_PayPal_Vendor&"', pcPay_PayPal_Partner='"&pcPay_PayPal_Partner&"', pcPay_PayPal_CardTypes='"&pcPay_PayPal_CardTypes&"' WHERE (((pcPay_PayPal_ID)=1));"
				set rs=Server.CreateObject("ADODB.Recordset")     
				rs.Open query, conntemp
				if err.number <> 0 then
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddPaymentOpt: "&Err.Description) 
				end if
			end if			

			
			query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",gwCode="&ppGwcode&",paymentDesc='"&PayPalName&"',sslUrl='"&sslUrl&"',priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"',pcPayTypes_ppab="&ppAB&" WHERE idPayment="&idPayment
			set rs=Server.CreateObject("ADODB.Recordset")  
			rs.Open query, conntemp
			if err.number <> 0 then
				set rs=nothing
				call closedb()
				session("ppGwCode")=""
				response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddPaymentOpt: "&Err.Description) 
			end if



		case "2"
			'request gateway variables and insert them into the verisign_pfp table
			query= "SELECT v_User,v_Password FROM verisign_pfp where id=1"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=conntemp.execute(query)
			if err.number <> 0 then
				set rs=nothing
				call closedb()
			  response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddPaymentOpt: "&Err.Description) 
			end If
			v_User2=rs("v_User")
			v_Password2=rs("v_Password")
			set rs=nothing
			v_Url=request.Form("pfp_Url")
			v_User=request.Form("pfp_User")
			if v_User="" then
				v_User=v_User2
			end if
			v_Partner=request.Form("pfp_Partner")
			v_Password=request.Form("pfp_Password")
			if v_Password="" then
				v_Password=v_Password2
			end if
			v_Vendor=request.Form("pfp_Vendor")
			pfl_testmode=request.Form("pfp_testmode")
			pfl_CSC=request.Form("pfp_CSC")
			if pfl_testmode="" then
				pfl_testmode=0
			end if
			pfl_transtype=request.Form("pfp_transtype") 


			query="UPDATE verisign_pfp SET v_Url='"&v_Url&"',v_User='"&v_User&"',v_Partner='"&v_Partner&"',v_Password='"&v_Password &"',v_Vendor='"&v_Vendor&"',pfl_testmode='"&pfl_testmode&"',pfl_transtype='"&pfl_transtype&"',pfl_CSC='"&pfl_CSC&"' where id=1"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=conntemp.execute(query)
			set rs=nothing
			if err.number <> 0 then
				call closedb()
			  response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddPaymentOpt: "&Err.Description) 
			end If
			query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"',pcPayTypes_ppab="&ppAB&" WHERE gwCode=2"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=conntemp.execute(query)
				set rs=nothing
			if err.number <> 0 then
				call closedb()
			  response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddPaymentOpt: "&Err.Description) 
			end If



		case "9"
			'request gateway variables for pfLink and insert them into the verisign_pfp table
			query= "SELECT v_User FROM verisign_pfp where id=1"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=conntemp.execute(query)
			if err.number <> 0 then
				set rs=nothing
				call closedb()
			  response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddPaymentOpt: "&Err.Description) 
			end If
			v_User2=rs("v_User")
			set rs=nothing
			v_User=request.Form("v_User")
			if v_User="" then
				v_User=v_User2
			end if
			v_Partner=request.Form("v_Partner")
			pfl_testmode=request.Form("pfl_testmode")
			if pfl_testmode="" then
				pfl_testmode=0
			end if
			pfl_transtype=request.Form("pfl_transtype")
			pfl_CSC=request.Form("pfl_CSC") 
			
			query="UPDATE verisign_pfp SET v_User='"&v_User&"',v_Partner='"&v_Partner &"',pfl_testmode='"&pfl_testmode&"',pfl_transtype='"&pfl_transtype&"',pfl_CSC='"&pfl_CSC&"' where id=1"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=conntemp.execute(query)
			set rs=nothing
			if err.number <> 0 then
				call closedb()
			  response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddPaymentOpt: "&Err.Description) 
			end If
			query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"',pcPayTypes_ppab="&ppAB&" WHERE gwCode=9"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=conntemp.execute(query)
			set rs=nothing
			if err.number <> 0 then
				call closedb()
			  response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddPaymentOpt: "&Err.Description) 
			end If

	end select
	call closedb()
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// Start: Centinel
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcPay_Cent_Active=request.Form("pcPay_Cent_Active")
	pcPay_Cent_TransactionURL=request.Form("pcPay_Cent_TransactionURL")
	pcPay_Cent_MerchantID=request.Form("pcPay_Cent_MerchantID")
	pcPay_Cent_ProcessorId=request.Form("pcPay_Cent_ProcessorId")
	pcPay_Cent_Password=request.Form("pcPay_Cent_Password")
	
	if pcPay_Cent_Active="YES" then
		pcPay_Cent_Active=1
	else
		pcPay_Cent_Active=0
	end if
	
	if (pcPay_Cent_Active=1) AND (pcPay_Cent_TransactionURL="" OR pcPay_Cent_MerchantID="" OR pcPay_Cent_ProcessorId="") then
		
		response.redirect "modPayPalPaymentOpt.asp?Mode=Edit&id="&idPayment&"&gwCode="&gwCode&"&msg="&Server.URLEncode("An error occured while trying to add Cardinal Centinel for PayPal. <b>""Tranaction URL, Merchant ID and Process ID""</b> are all required fields.")
		
	else

		call openDb()
		query="UPDATE pcPay_Centinel SET pcPay_Cent_TransactionURL='"&pcPay_Cent_TransactionURL&"',	pcPay_Cent_MerchantID='"&pcPay_Cent_MerchantID&"', pcPay_Cent_ProcessorId='"&pcPay_Cent_ProcessorId&"', pcPay_Cent_Password='"&pcPay_Cent_Password&"', pcPay_Cent_Active="&pcPay_Cent_Active&" WHERE pcPay_Cent_ID=1;"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=connTemp.execute(query)
		set rs=nothing
		call closedb()

	end if
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// End: Centinel
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	response.redirect "PaymentOptions.asp"
End If
'end of Submit section


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// Start: Centinel
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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


'// Card Types
M="0"
V="0"
D="0"
A="0"
if NOT len(pcPay_PayPal_CardTypes)>0 then
	pcPay_PayPal_CardTypes="V, M, D"
end if
cardTypeArray=split(pcPay_PayPal_CardTypes,", ")
for i=0 to ubound(cardTypeArray)
	select case cardTypeArray(i)
		case "M"
			M="1" 
		case "V"
			V="1"
		case "D"
			D="1"
		case "A"
			A="1"
	end select
next
%>

<script language="JavaScript">
<!--
function Form1_Validator(theForm)
{
<%if gwCode="7" then%>
	if (theForm.CDesc.value == "")
		{
			alert("Description is a required field.");
				theForm.CDesc.focus();
				return (false);
	}
<%end if%>
return (true);
}
function winemailobj(fileName)
	{
	myFloater=window.open('','myWindow','scrollbars=yes,status=no,width=300,height=350')
	myFloater.location.href=fileName;
	}
function wingwaycheckobj(fileName)
	{
	myFloater=window.open('','myWindow','scrollbars=yes,status=no,width=475,height=425')
	myFloater.location.href=fileName;
	}
//-->
</script>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<form method="post" name="form1" action="modPayPalPaymentOpt.asp" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<td colspan="3"> 				
			
			<table border="0" width="100%" cellpadding="4" cellspacing="0">
				
				<% 				
				select case gwCode
				
					case "3", "46", "53", "999999" 
					%>
						<tr> 
							<th colspan="2">Modify PayPal Settings - <a href="https://www.paypal.com/cgi-bin/webscr?cmd=_profile-comparison" target="_blank">More Information</a></th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<td colspan="2">
                            
                          
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
                             //-->
                            </SCRIPT>	
							</td>
						</tr>
                        <% 
						'// PAYPAL PAYMENT STANDARD - START
						if gwcode="3" then
						%>
						<tr> 
							<td valign="top"><div align="right"> 
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
                        if gwcode="999999" and pcPay_PayPal_Version="US" then
                        %>
                        <tr> 
							<td valign="top">
                            	<div align="right"> 
								<input type="radio" class="clearBorder" name="PayPalSelection" value="PayPalExp" onClick="whatPayPalSelected();" checked>
								</div>
                            </td>
							<td><strong>Express Checkout</strong> - (Alternative Payment Option) <a href="https://www.paypal.com/cgi-bin/webscr?cmd=_profile-comparison" target="_blank" class="pcSmallText">Compare Solutions</a><br>
						    Accept credit/debit cards and PayPal using a Business or Premier PayPal account. Customers benefit from a rapid checkout process without the need to enter their shipping address, resulting in fewer abandoned carts. Alternative payment solution: Customers shop on your web site and pay on PayPal without creating an account in your webstore. No merchant account required.</td>
						</tr>
						<%
                        else
                        %>
                        <tr><td style="height: 1px; padding: 0px;"><div style="display: none;"><input type="radio" name="PayPalSelection" value="PayPalExp"></div></td></tr>
                        <%
                        end if
                        '// PAYPAL EXPRESS - END
    
                        '// PAYPAL WPP - START
                        if gwcode="46" then
                        %>
                        <tr> 
							<td valign="top">
                            	<div align="right"> 
								<input type="radio" class="clearBorder" name="PayPalSelection" value="PayPalWP" onClick="whatPayPalSelected();" checked>
								</div>
                            </td>
							<td><strong>Website Payment Pro</strong> - <a href="https://www.paypal.com/cgi-bin/webscr?cmd=_wp-pro-overview-outside" target="_blank" class="pcSmallText">Solution Overview</a><br>
						    Includes both Express Checkout and Direct Payment - with Direct Payment the customer shops and pays on your website. PayPal merchant account required. </td>						
						</tr>
                        
						<%
                        else
                        %>
                        <tr><td style="height: 1px; padding: 0px;"><div style="display: none;"><input type="radio" name="PayPalSelection" value="PayPalWP"></div></td></tr>
                        <%
                        end if
                        '// PAYPAL WPP - END
    
                        '// PAYPAL EXPRESS UK - START
                        if gwcode="999999" AND pcPay_PayPal_Version="UK" then
                        %>
                        <tr> 
							<td valign="top">
                            <div align="right"> 
							<input type="radio" class="clearBorder" name="PayPalSelection" value="PayPalExpUK" onclick="whatPayPalSelected();" checked>
							</div></td>
							<td><strong>Express Checkout (UK)</strong> (Alternative Payment Option) - <a href="https://www.paypal.com/uk/cgi-bin/webscr?cmd=_profile-comparison" target="_blank" class="pcSmallText">Compare Solutions</a><br>
						   For UK merchants <u>only with a PayPal partner name of &quot;<strong>PayPalUK</strong>&quot;, merchant ID and username</u>. Accept credit/debit cards and PayPal using a Business or Premier PayPal account. Customers benefit from a rapid checkout process without the need to enter their shipping address, resulting in fewer abandoned carts. Alternative payment solution: Customers shop on your web site and pay on PayPal without creating an account in your webstore. No merchant account required.</td>
						</tr>
						<%
                        else
                        %>
                        <tr><td style="height: 1px; padding: 0px;"><div style="display: none;"><input type="radio" name="PayPalSelection" value="PayPalExpUK"></div></td></tr>
                        <%
                        end if
                        '// PAYPAL WPP - END
    
                        '// PAYPAL WPP UK - START
                        if gwcode="53" then
                        %>
						<tr> 
							<td valign="top">
                            <div align="right"> 
							<input type="radio" class="clearBorder" name="PayPalSelection" value="PayPalWPUK" onclick="whatPayPalSelected();" checked>
							</div>
                            </td>
							<td><strong>Website Payments Pro (UK)</strong> <a href="https://www.paypal.com/uk/cgi-bin/webscr?cmd=_profile-comparison" target="_blank">Compare Solutions</a><br>
						    For UK merchants <u>only with a PayPal partner name of &quot;<strong>PayPalUK</strong>&quot;, merchant ID and username</u>. Accept credit/debit cards on your web site, and payments from PayPal account holders using Express Checkout. PayPal merchant account required.</td>
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

							<!--PayPal Standard start here -->
							<table ID="PayPal_table" width="100%" border="0" cellspacing="1" cellpadding="1" <% if gwCode<>"3" then%>style="display:none"<%end if %>>
								<tr> 
									<th colspan="2">Configure Settings - Website Payments Standard</th>
								</tr>	
								<tr> 
									<td width="127" nowrap><div align="right">Email Address:</div></td>
									<td>
										<input type="text" value="<%=PayPal_Email%>" name="PayPal_Email" size="24"></td>
								</tr>
								<tr> 
									<td><div align="right">Currency:</div></td>
									<td> 
										<select name="PayPal_Currency">
											<option value="USD" <% if PayPal_Currency="USD" then%>selected<% end if %>>U.S. Dollars ($)</option>
                                            <option value="AUD" <% if PayPal_Currency="AUD" then%>selected<% end if %>>Australian Dollars ($)</option>
											<option value="CAD" <% if PayPal_Currency="CAD" then%>selected<% end if %>>Canadian Dollars (C $)</option>
                                            <option value="CZK" <% if PayPal_Currency="CZK" then%>selected<% end if %>>Czech Koruna</option>
                                            <option value="DKK" <% if PayPal_Currency="DKK" then%>selected<% end if %>>Danish Krone</option>
                                            <option value="EUR" <% if PayPal_Currency="EUR" then%>selected<% end if %>>Euros (€)</option>
                                            <option value="HKD" <% if PayPal_Currency="HKD" then%>selected<% end if %>>Hong Kong Dollar</option>
                                            <option value="HUF" <% if PayPal_Currency="HUF" then%>selected<% end if %>>Hungarian Forint</option>
                                            <option value="ILS" <% if PayPal_Currency="ILS" then%>selected<% end if %>>Israeli New Shekel</option>
                                            <option value="JPY" <% if PayPal_Currency="JPY" then%>selected<% end if %>>Yen (¥)</option>
                                            <option value="MXN" <% if PayPal_Currency="MXN" then%>selected<% end if %>>Mexican Peso</option> 
                                            <option value="NOK" <% if PayPal_Currency="NOK" then%>selected<% end if %>>Norwegian Krone</option>
                                            <option value="NZD" <% if PayPal_Currency="NZD" then%>selected<% end if %>>New Zealand Dollar</option>
                                            <option value="PHP" <% if PayPal_Currency="PHP" then%>selected<% end if %>>Philippine Peso</option> 
                                            <option value="PLN" <% if PayPal_Currency="PLN" then%>selected<% end if %>>Polish Zloty</option>
											<option value="GBP" <% if PayPal_Currency="GBP" then%>selected<% end if %>>Pounds Sterling (£)</option>											
											<option value="SGD" <% if PayPal_Currency="SGD" then%>selected<% end if %>>Singapore Dollar</option>
                                            <option value="SEK" <% if PayPal_Currency="SEK" then%>selected<% end if %>>Swedish Krona</option>                                   
                                            <option value="CHF" <% if PayPal_Currency="CHF" then%>selected<% end if %>>Swiss Franc</option>     
                                            <option value="TWD" <% if PayPal_Currency="TWD" then%>selected<% end if %>>Taiwan New Dollar</option>    
                                            <option value="THB" <% if PayPal_Currency="THB" then%>selected<% end if %>>Thai Baht</option>   
                                      	</select>						
                                	</td>
								</tr>
								<tr> 
									<td> <div align="right"> 
											<input name="PayPal_Sandbox" type="checkbox" class="clearBorder" value="YES" <% if PayPal_Sandbox=1 then%>checked<% end if %>>
										</div></td>
									<td><b>Enable Test Mode </b>(Credit cards will not be charged)
									<div>
									<font color="#FF0000">
									Visit <a href="https://developer.paypal.com" target="_blank">PayPal's Developer Site</a> to obtain "SandBox" (Test) credentials. <strong>
									Do not use your real account information.</strong>									</font>									</div>									</td>
								</tr>
							</table>
							<!--PayPal Standard end here -->




							<!--PayPal WPP start here -->
							<table ID="PayPalWP_table" width="100%" border="0" cellspacing="4" cellpadding="0" <% if gwCode<>"46" then%>style="display:none"<%end if %>>
								<tr> 
									<th colspan="2">Configure Settings - Website Payments Pro (US)</th>
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
									<td width="127" valign="top" nowrap><div align="right">Currency:</div></td>
									<td> <select name="pcPay_PayPal_Currency">
											<option value="AUD" <% if pcPay_PayPal_Currency="AUD" then%>selected<% end if %>>Australian Dollars ($)</option>
											<option value="CAD" <% if pcPay_PayPal_Currency="CAD" then%>selected<% end if %>>Canadian Dollars (C $)</option>
											<option value="EUR" <% if pcPay_PayPal_Currency="EUR" then%>selected<% end if %>>Euros (€)</option>
											<option value="GBP" <% if pcPay_PayPal_Currency="GBP" then%>selected<% end if %>>Pounds Sterling (£)</option>
											<option value="JPY" <% if pcPay_PayPal_Currency="JPY" then%>selected<% end if %>>Yen (¥)</option>
											<option value="USD" <% if pcPay_PayPal_Currency="USD" then%>selected<% end if %>>U.S. Dollars ($)</option>
										</select>									</td>
								</tr>						
								<tr> 
									<td valign="top" nowrap><div align="right">Transaction Type:</div></td>
									<td> 
                                    	<select name="pcPay_PayPal_TransType">
											<option value="1" <% if pcPay_PayPal_TransType=1 then%>selected<%end if %>>Sale (Authorize and Capture)</option>
											<option value="2" <% if pcPay_PayPal_TransType=2 then%>selected<%end if %>>Authorize Only</option>
										</select>
                                        &nbsp;
                                        <span class="pcCPnotes"><strong>NOTE:</strong> &quot;Sale&quot; will be used when API Credentials are unavailable.</span>									
                                    </td>
								</tr>
								<tr> 
									<td></td>
									<td>
										<div style="padding-bottom:4px">
											<span class="pcCPsectionTitle">Transaction Type tells PayPal how you want to obtain payment.</span>										</div>
										<li><strong>Sale</strong> indicates that authorization and capture occur at the same time. The sale is final. For more information, please consult the ProductCart User Guide.</li>
										<li><strong>Authorize Only</strong> indicates that the transaction is only authorized, but not settled. The payment will be captured at a later time.</li>									</td>
								</tr>
								<tr> 
									<td> <div align="right"> 
											<input name="pcPay_PayPal_Sandbox" type="checkbox" class="clearBorder" value="YES" <% if pcPay_PayPal_Sandbox=1 then%>checked<% end if %>>
										</div></td>
									<td><b>Enable Test Mode </b>(Credit cards will not be charged)
									<div>
									<font color="#FF0000">
									Visit <a href="https://developer.paypal.com" target="_blank">PayPal's Developer Site</a> to obtain "SandBox" (Test) credentials. <strong>
									Do not use your real account information.</strong>									</font>									</div>									</td>
								</tr>
								<tr> 
									<td> <div align="right">  
									<input name="pcPay_PayPal_CVC" type="checkbox" class="clearBorder" value="1" <% if pcPay_PayPal_CVC=1 then%>checked<% end if %>>
										</div></td>
									<td><b>Enable Credit Card Security Code</b> 
									<div>Every Credit Card has a 3 or 4 digit CVV Code (also known as CVV2), which is a security code designed for your safety and security. Check the box above and customers are required to enter their Credit Card's CVV Code. Note: You must also enable CVV in your PayPal Account.</div>									</td>
								</tr>
							</table>
							<!--PayPal WPP end here -->




							<!--PayPal Express start here -->
							<table ID="PayPalExp_table" width="100%" border="0" cellspacing="4" cellpadding="0" <% if gwCode<>"999999" OR pcPay_PayPal_Version<>"US" then%>style="display:none"<%end if %>>
								<tr> 
									<th colspan="2">Configure Settings - Express Checkout</th>
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
									<td valign="top" nowrap><div align="right">Currency:</div></td>
									<td> <select name="pcPay_PayPale_Currency">
											<option value="AUD" <% if pcPay_PayPal_Currency="AUD" then%>selected<% end if %>>Australian Dollars ($)</option>
											<option value="CAD" <% if pcPay_PayPal_Currency="CAD" then%>selected<% end if %>>Canadian Dollars (C $)</option>
											<option value="EUR" <% if pcPay_PayPal_Currency="EUR" then%>selected<% end if %>>Euros (€)</option>
											<option value="GBP" <% if pcPay_PayPal_Currency="GBP" then%>selected<% end if %>>Pounds Sterling (£)</option>
											<option value="JPY" <% if pcPay_PayPal_Currency="JPY" then%>selected<% end if %>>Yen (¥)</option>
											<option value="USD" <% if pcPay_PayPal_Currency="USD" then%>selected<% end if %>>U.S. Dollars ($)</option>
										</select>									</td>
								</tr>						
								<tr> 
									<td valign="top" nowrap><div align="right">Transaction Type:</div></td>
									<td> 
                                    	<select name="pcPay_PayPale_TransType">
											<option value="1" <% if pcPay_PayPal_TransType=1 then%>selected<%end if %>>Sale (Authorize and Capture)</option>
											<option value="2" <% if pcPay_PayPal_TransType=2 then%>selected<%end if %>>Authorize Only</option>
										</select>
                                        &nbsp;
                                        <span class="pcCPnotes"><strong>NOTE:</strong> &quot;Sale&quot; will be used when API Credentials are unavailable.</span>											
                                	</td>
								</tr>
								<tr> 
									<td></td>
									<td>
										<div style="padding-bottom:4px">
											<span class="pcCPsectionTitle">Transaction Type tells PayPal how you want to obtain payment.</span>										</div>
										<li><strong>Sale</strong> indicates that authorization and capture occur at the same time. The sale is final. For more information, please consult the ProductCart User Guide.</li>
										<li><strong>Authorize Only</strong> indicates that the transaction is only authorized, but not settled. The payment will be captured at a later time.</li>									</td>
								</tr>
								<tr> 
									<td> <div align="right"> 
											<input name="pcPay_PayPale_Sandbox" type="checkbox" class="clearBorder" value="YES" <% if pcPay_PayPal_Sandbox=1 then%>checked<% end if %>>
										</div></td>
									<td><b>Enable Test Mode </b>(Credit cards will not be charged)
									<div>
									<font color="#FF0000">
									Visit <a href="https://developer.paypal.com" target="_blank">PayPal's Developer Site</a> to obtain "SandBox" (Test) credentials. <strong>
									Do not use your real account information.</strong>									</font>									</div>									</td>
								</tr>
							</table>
							<!--PayPal Express end here -->




							<!--PayPal WPP UK start here -->								
							<table ID="PayPalWPUK_table"width="100%" border="0" cellspacing="4" cellpadding="0" <% if gwCode<>"53" then%>style="display:none"<%end if %>>
								<tr> 
									<th colspan="2">Configure Settings - Website Payments Pro (UK)</th>
								</tr>	
								<tr> 
									<td width="127" valign="top"><div align="right">User Name:</div></td>
									<td><input type="text" value="<%=pcPay_PayPal_Username%>" name="pcPay_PayPalUK_Username" size="30" maxlength="50"></td>
								</tr>
								
								<tr> 
									<td valign="top"><div align="right">Password:</div></td>
									<td><input type="text" value="<%=pcPay_PayPal_Password%>" name="pcPay_PayPalUK_Password" size="30" maxlength="50"></td>
								</tr>						  
								<tr> 
									<td width="127" valign="top"><div align="right">Partner:</div></td>
									<td>
									<input type="text" value="<%=pcPay_PayPal_Partner%>" name="pcPay_PayPalUK_Partner" size="30" maxlength="250">									</td>
								</tr>
								<tr> 
									<td width="127" valign="top"><div align="right">Vendor:</div></td>
									<td>
									<input type="text" value="<%=pcPay_PayPal_Vendor%>" name="pcPay_PayPalUK_Vendor" size="30" maxlength="250">									</td>
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
											<option value="AUD" <% if pcPay_PayPal_Currency="AUD" then%>selected<% end if %>>Australian Dollars ($)</option>
											<option value="CAD" <% if pcPay_PayPal_Currency="CAD" then%>selected<% end if %>>Canadian Dollars (C $)</option>
											<option value="EUR" <% if pcPay_PayPal_Currency="EUR" then%>selected<% end if %>>Euros (€)</option>
											<option value="GBP" <% if pcPay_PayPal_Currency="GBP" then%>selected<% end if %>>Pounds Sterling (£)</option>
											<option value="JPY" <% if pcPay_PayPal_Currency="JPY" then%>selected<% end if %>>Yen (¥)</option>
											<option value="USD" <% if pcPay_PayPal_Currency="USD" then%>selected<% end if %>>U.S. Dollars ($)</option>
										</select>												
									</td>
								</tr>
								<tr> 
									<td valign="top"> <div align="right">Transaction Type:</div></td>
									<td> <select name="pcPay_PayPalUK_TransType">
											<option value="1" <% if pcPay_PayPal_TransType=1 then%>selected<%end if %>>Sale (Authorize and Capture)</option>
											<option value="2" <% if pcPay_PayPal_TransType=2 then%>selected<%end if %>>Authorize Only</option>
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
									<input name="pcPay_PayPalUK_Sandbox" type="checkbox" class="clearBorder" value="YES" <% if pcPay_PayPal_Sandbox=1 then%>checked<% end if %>>
									</div></td>
									<td><b>Enable Test Mode </b>(Credit cards will not be charged)<div><font color="#FF0000">Visit <a href="https://developer.paypal.com" target="_blank">PayPal's Developer Site</a> to obtain "SandBox" (Test) credentials. <strong>Do not use your real account information.</strong></font></div></td>
								</tr>
								<tr> 
									<td valign="top"> <div align="right"> 
									<input name="pcPay_PayPalUK_CVC" type="checkbox" class="clearBorder" value="1" <% if pcPay_PayPal_CVC=1 then%>checked<% end if %>>
									</div></td>
									<td><b>Enable Credit Card Security Code</b> 
									<div>Every Credit Card has a 3 or 4 digit CVV Code (also known as CVV2), which is a security code designed for your safety and security. Check the box above and customers are required to enter their Credit Card's CVV Code. Note: You must also enable CVV in your PayPal Account.</div>												</td>
								</tr>
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
							</table>
							<!--PayPal WPP UK end here -->




							<!--PayPal Express UK end here -->
							<table ID="PayPalExpUK_table" width="100%" border="0" cellspacing="4" cellpadding="0" <% if gwCode<>"999999" OR pcPay_PayPal_Version<>"UK" then%>style="display:none"<%end if %>>
								<tr> 
									<th colspan="2">Configure Settings - Express Checkout (UK)</th>
								</tr>	
								<tr> 
									<td width="127" valign="top"><div align="right">User Name:</div></td>
									<td><input type="text" value="<%=pcPay_PayPal_Username%>" name="pcPay_PayPalUKe_Username" size="30" maxlength="50"></td>
								</tr>
								
								<tr> 
									<td valign="top"><div align="right">Password:</div></td>
									<td><input type="text" value="<%=pcPay_PayPal_Password%>" name="pcPay_PayPalUKe_Password" size="30" maxlength="50"></td>
								</tr>						  
								<tr> 
									<td width="127" valign="top"><div align="right">Partner:</div></td>
									<td>
									<input type="text" value="<%=pcPay_PayPal_Partner%>" name="pcPay_PayPalUKe_Partner" size="30" maxlength="250">									</td>
								</tr>
								<tr> 
									<td width="127" valign="top"><div align="right">Vendor:</div></td>
									<td>
									<input type="text" value="<%=pcPay_PayPal_Vendor%>" name="pcPay_PayPalUKe_Vendor" size="30" maxlength="250">									</td>
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
											<option value="AUD" <% if pcPay_PayPal_Currency="AUD" then%>selected<% end if %>>Australian Dollars ($)</option>
											<option value="CAD" <% if pcPay_PayPal_Currency="CAD" then%>selected<% end if %>>Canadian Dollars (C $)</option>
											<option value="EUR" <% if pcPay_PayPal_Currency="EUR" then%>selected<% end if %>>Euros (€)</option>
											<option value="GBP" <% if pcPay_PayPal_Currency="GBP" then%>selected<% end if %>>Pounds Sterling (£)</option>
											<option value="JPY" <% if pcPay_PayPal_Currency="JPY" then%>selected<% end if %>>Yen (¥)</option>
											<option value="USD" <% if pcPay_PayPal_Currency="USD" then%>selected<% end if %>>U.S. Dollars ($)</option>
										</select>
									</td>
								</tr>
								<tr> 
									<td valign="top"> <div align="right">Transaction Type:</div></td>
									<td> <select name="pcPay_PayPalUKe_TransType">
											<option value="1" <% if pcPay_PayPal_TransType=1 then%>selected<%end if %>>Sale (Authorize and Capture)</option>
											<option value="2" <% if pcPay_PayPal_TransType=2 then%>selected<%end if %>>Authorize Only</option>
										</select>
									</td>
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
									<input name="pcPay_PayPalUKe_Sandbox" type="checkbox" class="clearBorder" value="YES" <% if pcPay_PayPal_Sandbox=1 then%>checked<% end if %>>
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
									<input name="pcPay_PayPalUKe_CVC" type="checkbox" class="clearBorder" value="1" <% if pcPay_PayPal_CVC=1 then%>checked<% end if %>>
									</div></td>
									<td><b>Enable Credit Card Security Code</b> 
									<div>Every Credit Card has a 3 or 4 digit CVV Code (also known as CVV2), which is a security code designed for your safety and security. Check the box above and customers are required to enter their Credit Card's CVV Code. Note: You must also enable CVV in your PayPal Account.</div>												</td>
								</tr>
							</table>
							<!--PayPal Express UK end here -->




				<% case "2" %>
					<tr> 
						<th colspan="2">Modify PayPal PayFlow Pro settings - <a href="https://www.paypal.com/cgi-bin/webscr?cmd=_profile-comparison" target="_blank">More Information</a></th>
					</tr>
					<tr> 
						<td>&nbsp;</td>
						<td> <% if pfp_testmode="YES" then %> <input type="checkbox" class="clearBorder" name="pfp_testmode" value="YES" checked> 
							<% else %> <input type="checkbox" class="clearBorder" name="pfp_testmode" value="YES"> 
							<% end if %> <b>Enable Test Mode </b>(Credit 
							cards will not be charged)</td>
					</tr>
					<% dim pfp_UserCnt,pfp_UserEnd,pfp_UserStart
					pfp_UserCnt=(len(pfp_User)-2)
					pfp_UserEnd=right(pfp_User,2)
					pfp_UserStart=""
					for c=1 to pfp_UserCnt
					pfp_UserStart=pfp_UserStart&"*"
					next %>
					<tr> 
						<td>&nbsp;</td>
						<td>Current User:&nbsp;<%=pfp_UserStart&pfp_UserEnd%></td>
					</tr>
					<tr> 
						<td>&nbsp;</td>
						<td> For security reasons, your &quot;User&quot; is only 
							partially shown on this page. The password is not shown. 
							If you need to edit your account information, please re-enter 
							your &quot;User&quot; and password below.</td>
					</tr>
					<tr> 
						<td width="20%">&nbsp;</td>
						<td>Change User: 
							<input type="text" value="" name="pfp_User" size="24"> 
							<input type="hidden" value="<%=pfp_Url%>" name="pfp_Url" size="24"></td>
					</tr>
					<tr> 
						<td width="20%">&nbsp;</td>
						<td>Partner: 
							<input type="text" value="<%=pfp_Partner%>" name="pfp_Partner" size="24"></td>
					</tr>
					<tr> 
						<td width="20%">&nbsp;</td>
						<td>Change Password: 
							<input type="password" value="" name="pfp_Password" size="24"></td>
					</tr>
					<tr> 
						<td width="20%">&nbsp;</td>
						<td>Vendor: 
							<input type="text" value="<%=pfp_Vendor%>" name="pfp_Vendor" size="24"></td>
					</tr>
						<tr> 
							<td width="20%">&nbsp;</td>
							<td>Transaction Type: 
								<select name="pfp_transtype">
								<option value="S" selected>Sale</option>
								<option value="A"<% if pfp_transtype="A" then
								response.write " selected"
								end if %>
								>Authorize Only</option>
								</select></td>
						</tr>
						<tr> 
							<td>&nbsp;</td>
							<td>Require CSC: 
								<% if pfp_CSC="YES" then %> <input type="radio" class="clearBorder" name="pfp_CSC" value="YES" checked>
								Yes 
								<input name="pfp_CSC" type="radio" class="clearBorder" value="NO">
								No 
								<% else %> <input type="radio" class="clearBorder" name="pfp_CSC" value="YES">
								Yes 
								<input name="pfp_CSC" type="radio" class="clearBorder" value="NO" checked>
								No 
								<% end if %> </td>
						</tr>

					<% case "9" %>
						<tr> 
							<th colspan="2">Modify PayPal PayFlow Link settings - <a href="https://www.paypal.com/cgi-bin/webscr?cmd=_profile-comparison" target="_blank">More Information</a></th>
						</tr>
						<tr> 
							<td>&nbsp;</td>
							<td> <% if pfl_testmode="YES" then %> <input type="checkbox" class="clearBorder" name="pfl_testmode" value="YES" checked> 
								<% else %> <input type="checkbox" class="clearBorder" name="pfl_testmode" value="YES"> 
								<% end if %> <b>Enable Test Mode </b>(Credit 
								cards will not be charged)</td>
						</tr>
						<% dim v_UserCnt,v_UserEnd,v_UserStart
						v_UserCnt=(len(v_User)-2)
						v_UserEnd=right(v_User,2)
						v_UserStart=""
						for c=1 to v_UserCnt
						v_UserStart=v_UserStart&"*"
						next %>
						<tr> 
							<td>&nbsp;</td>
							<td>Current Login:&nbsp;<%=v_UserStart&v_UserEnd%></td>
						</tr>
						<tr> 
							<td>&nbsp;</td>
							<td> For security reasons, your &quot;Login&quot; is only 
								partially shown on this page. If you need to edit your 
								account information, please re-enter your &quot;Login&quot; 
								below.</td>
						</tr>
						<tr> 
							<td width="20%">&nbsp;</td>
							<td>Change Login: 
								<input type="text" value="" name="v_User" size="24"></td>
						</tr>
						<tr> 
							<td width="20%">&nbsp;</td>
							<td>Partner: 
								<input type="text" value="<%=v_Partner%>" name="v_Partner" size="24"></td>
						</tr>
						<tr> 
							<td width="20%">&nbsp;</td>
							<td>Transaction Type: 
								<select name="pfl_transtype">
									<option value="S" selected>Sale</option>
									<option value="A"<% if pfl_transtype="A" then
									response.write " selected"
									end if %>
									>Authorize Only</option>
								</select></td>
						</tr>
						<tr> 
							<td>&nbsp;</td>
							<td>Require CSC: 
								<% if pfl_CSC="YES" then %> <input type="radio" class="clearBorder" name="pfl_CSC" value="YES" checked>
								Yes 
								<input name="pfl_CSC" type="radio" class="clearBorder" value="NO">
								No 
								<% else %> <input type="radio" class="clearBorder" name="pfl_CSC" value="YES">
								Yes 
								<input name="pfl_CSC" type="radio" class="clearBorder" value="NO" checked>
								No 
								<% end if %> </td>
						</tr>	
		
				<% 
				end select
				' End payment gateway specific settings
				' Show fields shared by all payment gateways %>
			</table>		
			
		</td>
	</tr>
	<tr> 
		<td height="29" colspan="4"><hr size="1"></td>
	</tr>
	<tr>
		<th colspan="4">You have the option to charge a processing fee for this payment option.</th>
	</tr>
	<tr> 
		<td colspan="4"> 
		
			<% '// Processing Fees %>
			<table width="100%" border="0" cellspacing="0" cellpadding="4">
              <tr>
                <td>Processing fee:</td>
                <td><% if percentageToAdd="0" then %>
                    <input type="radio" class="clearBorder" name="priceToAddType" value="price" checked>
                    <% else %>
                    <input type="radio" class="clearBorder" name="priceToAddType" value="price">
                    <% end if %>
                  Flat Fee&nbsp;&nbsp;<%=scCurSign%>
                  <input name="priceToAdd" size="6" value="<%=money(priceToAdd)%>">                </td>
              </tr>
              <tr>
                <td width="20%">&nbsp;</td>
                <td width="80%"><% if percentageToAdd<>"0" then %>
                    <input type="radio" class="clearBorder" name="priceToAddType" value="percentage" checked>
                    <% else %>
                    <input type="radio" class="clearBorder" name="priceToAddType" value="percentage">
                    <% end if %>
                  Percentage of Order Total&nbsp;&nbsp;%
                  <input name="percentageToAdd" size="6" value="<%=percentageToAdd%>">                </td>
              </tr>
             
             <% if gwCode<>"999999" then %>
              <tr>
                <th colspan="2">You can change the display name that is shown for this payment type. </th>
              </tr>
              <tr>
                <td>Payment Name:</td>
                <% if isNull(paymentNickName) OR paymentNickName="" then
else
paymentNickName=replace(paymentNickName,"""","&quot;")
end if %>
                <td><input name="paymentNickName" value="<%=paymentNickName%>" size="35" maxlength="255">                </td>
              </tr>
              <% end if %>
              
              <% if gwCode="1" OR gwCode="18" OR  gwCode="24" or gwCode="27" or gwCode="35" or gwCode="37" then %>
              <tr>
                <td>eCheck Name:</td>
                <td><input name="eCheckNickName" value="<%=replace(eCheckNickName,"""","&quot;")%>" size="35" maxlength="255">                </td>
              </tr>
              <% end if %>
              <% if gwCode="1" OR gwCode="2" OR gwCode="34" OR gwCode="35" then %>
              <tr>
                <td height="29" colspan="4"><hr size="1">                </td>
              </tr>
              <tr>
                <td colspan=2><strong>Centinel Settings</strong></td>
              </tr>
              <tr>
                <td><div align="right">
                    <input name="pcPay_Cent_Active" type="checkbox" class="clearBorder" id="pcPay_Cent_Active" value="YES" <%if pcPay_Cent_Active=1 then%>checked<%end if%>>
                </div></td>
                <td><b>Enable Centinel</b></td>
              </tr>
              <tr>
                <td><div align="right">Transaction URL:</div></td>
                <% if pcPay_Cent_TransactionURL<>"" then
				pcPay_Cent_TransactionURL=replace(pcPay_Cent_TransactionURL,"""","&quot;")
			end if %>
                <td><input name="pcPay_Cent_TransactionURL" value="<%=pcPay_Cent_TransactionURL%>" size="50" maxlength="255">                </td>
              </tr>
              <tr>
                <td><div align="right">Merchant ID:</div></td>
                <% if pcPay_Cent_MerchantID<>"" then
				pcPay_Cent_MerchantID=replace(pcPay_Cent_MerchantID,"""","&quot;")
			end if %>
                <td><input name="pcPay_Cent_MerchantID" value="<%=pcPay_Cent_MerchantID%>" size="35" maxlength="255">                </td>
              </tr>
              <tr>
                <td><div align="right">Processor ID: </div></td>
                <% if pcPay_Cent_ProcessorID<>"" then
				pcPay_Cent_ProcessorID=replace(pcPay_Cent_ProcessorID,"""","&quot;")
			end if %>
                <td><input name="pcPay_Cent_ProcessorID" value="<%=pcPay_Cent_ProcessorID%>" size="35" maxlength="255">                </td>
              </tr>
              <% end if %>
            </table>
			
		</td>
	</tr>
	<%'Start SDBA%>
	<tr>
		<td colspan="4" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="4">Order Processing: Order Status and Payment Status</th>
	</tr>
	<tr>
		<td colspan="4">Process orders when they are placed: <input type="checkbox" class="clearBorder" name="pcv_processOrder" value="1" <%if pcv_processOrder="1" then%>checked<%end if%>>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=301')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
	</tr>
	<tr>
		<td colspan="4">When orders are placed, set the payment status to:
			<select name="pcv_setPayStatus">
				<option value="3" selected="selected">Default</option>
				<option value="0" <%if pcv_setPayStatus="0" then%>selected<%end if%>>Pending</option>
				<option value="1" <%if pcv_setPayStatus="1" then%>selected<%end if%>>Authorized</option>
				<option value="2" <%if pcv_setPayStatus="2" then%>selected<%end if%>>Paid</option>
			</select>
			&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=302')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>		</td>
	</tr>
	<%'End SDBA%>
	<tr>
		<td colspan="4" class="pcCPspacer">
		<input type="hidden" name="idPayment" value="<%=idPayment%>"> 
		<input type="hidden" name="gwCode" value="<%=gwCode%>">		</td>
	</tr>
	<tr>
		<td colspan="4" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<td align="center" colspan="4">
			<input type="submit" name="Submit" value="Update" class="submit2"> 
			&nbsp;
			<input type="button" name="back" value="Back" onClick="javascript:history.back()" class="ibtnGrey">		</td>
	</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->
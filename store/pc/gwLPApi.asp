<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/ppdstatus.inc"-->
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

'/////////////////////////////////////////////////
' Settings for debug log
'
'/////////////////////////////////////////////////

    ' Declarations for form fields
 
    
    ' transaction object
    Dim LPTxn 


    ' response holders
    Dim R_Time 
    Dim R_Ref 
    Dim R_Approved 
    Dim R_Code 
    Dim R_Authresr 
    Dim R_Error 
    Dim R_OrderNum 
    Dim R_Message 
    Dim R_Score 
    Dim R_TDate 
    Dim R_AVS 
    Dim R_FraudCode 
    Dim R_ESD 
    Dim R_Tax 
    Dim R_Shipping 

    ' Top level LPOrderPart
    Dim order,op 

    Dim dbg
    dbg = True

 
     Function LeadingZero(ByVal InpStr) 
        If Len(InpStr) = 1 Then
            LeadingZero =  ("0" & InpStr)
        Else
            LeadingZero = InpStr
        End If
    End Function

 

    

     Function ParseTag( tag ,  rsp ) 
        Dim sb 
        Dim idxSt, idxEnd 'As Integer
        
        rsp = rsp
        
        sb = "<" & tag & ">"
        idxSt = -1
        idxEnd = -1

        idxSt = InStr(rsp,sb)
        If 0 = idxSt Then
            ParseTag = ""
            Exit Function
        End If

        idxSt = idxSt + Len(sb)
        sb = "</" & tag & ">"
        idxEnd = InStr(idxSt, rsp,sb)

        If 0 = idxEnd Then
           ParseTag = ""
           Exit Function
        End If

        ParseTag = Mid(rsp, idxSt, (idxEnd - idxSt))

    End Function

     Sub Cleanup()

        If Not (LPTxn Is Nothing) Then
            Set LPTxn = Nothing
        End If
        If Not (order Is Nothing) Then
            res = order.removeAllParts()
            order = Nothing
        End If
        If Not (op Is Nothing) Then
            res = op.removeAllParts()
            op = Nothing
        End If

    End Sub
'//Set redirect page to the current file name
session("redirectPage")="gwlpAPI.asp"

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

<% call opendb() 
query="SELECT orders.shipmentDetails, orders.discountDetails, orders.taxAmount FROM orders WHERE (orders.idOrder)="&pcGatewayDataIdOrder&";"
set rs=server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
pshipmentDetails=rs("shipmentDetails")
pdiscountDetails=rs("discountDetails")
pcv_TaxAmount=rs("taxAmount")
set rs=nothing
call closedb()
if NOT isNumeric(pcv_TaxAmount) then
	pcv_TaxAmount=0
end if
'shipping details
pcv_shipping=split(pshipmentDetails,",")
if ubound(pcv_shipping)>1 then
	if NOT isNumeric(pcv_shipping(2)) then
		pcv_Postage=0
	else
		pcv_Postage=pcv_shipping(2)
	end if
else
	pcv_Postage=0
end if

'Check if more then one discount code was utilized
if instr(pdiscountDetails,",") then
	DiscountDetailsArry=split(pdiscountDetails,",")
	intArryCnt=ubound(DiscountDetailsArry)
else
	intArryCnt=0
end if
pTotalDiscountAmount=0
for k=0 to intArryCnt
	if intArryCnt=0 then
		pTempDiscountDetails=pdiscountDetails
	else
		pTempDiscountDetails=DiscountDetailsArry(k)
	end if
	if instr(pTempDiscountDetails,"- ||") then 
		discounts= split(pTempDiscountDetails,"- ||")
		pdiscountDesc=discounts(0)
		pdiscountAmt=trim(discounts(1))
		pIsNumeric=1
		if NOT isNumeric(pdiscountAmt) then
			pdiscountAmt=0
			pIsNumeric=0
		end if
		if (pdiscountAmt>0 OR pdiscountAmt=0) AND pIsNumeric=1 then
			storeAdminEmail=storeAdminEmail & dictLanguage.Item(Session("language")&"_adminMail_9") & pdiscountDesc & vbCrLf
		end if
	Else
		pdiscountAmt=0
	end if
	pTotalDiscountAmount=pTotalDiscountAmount+pdiscountAmt
Next



pcv_SubTotal=ccur(pcBillingTotal)-ccur(pcv_Postage)-ccur(pcv_TaxAmount)
%>
<% '//Set customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer

dim connTemp, rs
call openDb()

'//Retrieve any gateway specific data from database or hard-code the variables

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


configfile = lp_storeName ' Change this to your store number 
if PPD="1" then
	filename="/"&scPcFolder&"/" & scAdminFolderName
else
	filename="../"&scAdminFolderName
end if
filename = Server.MapPath (filename)
keyfile    = filename &"\" &lp_storeName&".pem" ' Change this to the name and location of your certificate file 

if lp_testmode ="YES" Then
	 host 	   = "staging.linkpt.net"
else
	 host 	   = "secure.linkpt.net"
End if 
Const port	   = "1129"

   
set rs=nothing
call closedb()

if request("PaymentSubmitted")="Go" then
	If lp_CVM="1" Then
		reqCVV = request.Form("cvm")			 
		if not isnumeric(reqCVV) or len(reqCVV) < 3 or len(reqCVV) > 4 Then				 
			Response.redirect"msgb.asp?message="&server.URLEncode("<b>Error</b>: "& dictLanguage.Item(Session("language")&"_paymntb_o_7")&"<br><br><a href="""&tempURL&"?psslurl=gwlpAPI.asp&idCustomer="&session("idcustomer")&"&idOrder="&session("GWOrderID")&"""><img src="""&rslayout("back")&""" border=0></a>")
		End If
	End if 

	dim varReply, nStatus, strErrorInfo	
	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	' Create an empty order
	Set order = Server.CreateObject("LpiCom_6_0.LPOrderPart")
	order.setPartName("order")
	' Create an empty part
	Set op = Server.CreateObject("LpiCom_6_0.LPOrderPart")                

	' Build 'orderoptions'
	' For a test, set result to GOOD, DECLINE, or DUPLICATE
	if lp_testmode ="YES" Then
		res=op.put("result", "GOOD")			
	else
		res=op.put("result", "LIVE")			
	End if 
	
	res=op.put("ordertype", ucase(lp_transType))		
	' add 'orderoptions to order
	res=order.addPart("orderoptions", op)
	
	res=op.clear()
		
	if pcv_UsePcOrderID=1 then
		res=op.put("oid",session("GWOrderId"))
	else
		res=op.put("oid","")          
	end if
        
	res=op.put("ip",pcCustIpAddress)
	' add 'merchantinfo to order
	res=order.addPart("transactiondetails", op)

	' Build 'merchantinfo'
	res=op.clear()
	
	res=op.put("configfile",configfile)
	' add 'merchantinfo to order
	res=order.addPart("merchantinfo", op)
	ccnum = request.form("cardnumber") 
	ccexpmonth = request.Form("expmonth")
	ccexpyear = request.Form("expyear")
	' Build 'creditcard'
	res=op.clear()
	res=op.put("cardnumber",ccnum )
	res=op.put("cardexpmonth",ccexpmonth )
	res=op.put("cardexpyear", right(ccexpyear,2))
	if lp_CVM = "1" then
		res=op.put("cvmvalue",left(request.Form("cvm"),4))
		res=op.put("cvmindicator","provided")
	end if
	' add 'creditcard to order
	res=order.addPart("creditcard", op)
	
	public function cleanInput(strRawText,strType)

		if strType="Number" then
			strClean="0123456789."
			bolHighOrder=false
		elseif strType="VendorTxCode" then
			strClean="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_."
			bolHighOrder=false
		else
			strClean=" ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789" & vbCRLF
			bolHighOrder=true
		end if
			
		strCleanedText=""
		iCharPos = 1
	
		do while iCharPos<=len(strRawText)
			'** Only include valid characters **
			chrThisChar=mid(strRawText,iCharPos,1)
	
			if instr(StrClean,chrThisChar)<>0 then 
				strCleanedText=strCleanedText & chrThisChar
			elseif bolHighOrder then
				'** Fix to allow accented characters and most high order bit chars which are harmless **
				if asc(chrThisChar)>=191 then strCleanedText=strCleanedText & chrThisChar
			end if
	
			iCharPos=iCharPos+1
		loop       
	  
		cleanInput = trim(strCleanedText)
		
	end function

	' Build 'billing'
	  res=op.clear()
	  res=op.put("name", cleanInput(pcBillingFirstName, "TEXT") &" " & cleanInput(pcBillingLastname, "TEXT"))
	  res=op.put("company", cleanInput(pcBillingCompany, "TEXT"))
	  res=op.put("address1", cleanInput(pcBillingAddress, "TEXT"))
	  res=op.put("address2", cleanInput(pcBillingAddress2, "TEXT"))
	  if pcBillingStateCode <>"" then
          res=op.put("state",pcBillingStateCode)
	  else
		  res=op.put("state",cleanInput(pcBillingProvince , "TEXT"))
	  end if
	  res=op.put("city",cleanInput(pcBillingcity , "TEXT"))
	  res=op.put("country",pcBillingCountryCode)
	  res=op.put("zip",pcBillingPostalCode)
	  res=op.put("email",cleanInput(pcCustomerEmail, "TEXT"))
	  res=op.put("phone",pcBillingPhone)
	  
 
	' add 'billing to order
	  res=order.addPart("billing", op)
	
	
	' Build 'shipping'
	 res=op.clear()
	 res=op.put("name",cleanInput(pcShippingFullName, "TEXT"))
	 res=op.put("address1", cleanInput(pcshippingAddress, "TEXT"))
	 res=op.put("address2", cleanInput(pcshippingAddress2, "TEXT"))
	 if pcshippingStateCode <>"" then
         res=op.put("state",pcshippingStateCode )
	 else
		 res=op.put("state",cleanInput(pcshippingProvince , "TEXT"))
	 end if
	 res=op.put("city",cleanInput(pcshippingcity, "TEXT"))
	 res=op.put("country",pcshippingCountryCode)
	 res=op.put("zip",pcshippingPostalCode)


	' add 'shipping to order
	res=order.addPart("shipping", op)

	' Build 'payment'
	res=op.clear()
	
	If ccur(pTotalDiscountAmount)=0 then
		res=op.put("subtotal",replace(money(pcv_SubTotal),",",""))
		res=op.put("tax",money(pcv_TaxAmount))
		res=op.put("shipping",money(pcv_Postage))
	End If
	res=op.put("chargetotal",replace(money(pcBillingTotal),",",""))
	
	' add 'payment to order
	res=order.addPart("payment", op)

	' create transaction object	
	Set LPTxn = Server.CreateObject("LpiCom_6_0.LinkPointTxn")
	' Set this one to false if you don't want any logging
	Const fLog = False
	
	' set logging level  
	' 0 - no logging
	' 1 - errors
	' 2 - debug ( errors + addl info )
	' 3 - trace ( full debug mode with call tracing )
	Const logLvl = 0
		  
	' set log file name
	' IMPORTANT: this file must have write access rights 
	'            for IIS' default IUSR_XXXXXX user account.
	'	     Otherwise no logging will take place
	logFile = "LINKLOG.log" ' Change this if you want logging

	if (fLog = True) and ( logLvl > 0 ) Then
        
		Dim res1, resDesc
		resDesc = "ORDID:" & session("gwidorder")
		res1 = res
		if PPD="1" then
			filename2="/"&scPcFolder&"/includes"
		else
			filename2="../includes"
		end if
		logFile = Server.MapPath(filename2) &"\" & logFile
		response.write logFile
		'response.end 
		'Next call return level of accepted logging in 'res1'
		'On error 'res1' contains negative number
		'You can check 'resDesc' to get error description
		'if any
		
		res = LPTxn.setDbgOpts(logFile,logLvl,resDesc,res1)
	End If
        
	' get outgoing XML from 'order' object
	Dim outXml, resp
	
	outXml = order.toXML()
	'response.write keyfile
	'Response.end
	' Call LPTxn
	rsp = LPTxn.send(keyfile, host, port, outXml)

	'response.buffer=true
	'response.write "Please be patient while your order is being processed. This could take up to 2 minutes, depending on your connection and current internet traffic."
	
	'Store transaction data on Session and redirect
	
	Set LPTxn = Nothing
	Set order = Nothing
	Set op    = Nothing
		
	R_Time = ParseTag("r_time", rsp)
	R_Ref = ParseTag("r_ref", rsp)		
	R_Approved = ParseTag("r_approved", rsp)
	R_Code = ParseTag("r_code", rsp)
	R_OrderNum = ParseTag("r_ordernum", rsp)
	R_Message = ParseTag("r_message", rsp)		
	R_Error = ParseTag("r_error", rsp)		
	R_TDate = ParseTag("r_tdate", rsp)
	
	Set LPTxn = Server.CreateObject("LpiCom_6_0.LinkPointTxn")

	if R_Approved = "APPROVED" then 

		session("GWAuthCode")=R_Code
		session("GWTransId")=R_Ref
		session("GWTransType")=lp_transType

		If lCASE(lp_transType)="preauth" Then
		
			call opendb()
			
			Dim pTodaysDate
			pTodaysDate=Date()
			if SQL_Format="1" then
				pTodaysDate=Day(pTodaysDate)&"/"&Month(pTodaysDate)&"/"&Year(pTodaysDate)
			else
				pTodaysDate=Month(pTodaysDate)&"/"&Day(pTodaysDate)&"/"&Year(pTodaysDate)
			end if
			if scDB="Access" then
				tmpStr="#"& pTodaysDate &"#"
			else
				tmpStr="'"& pTodaysDate &"'"
			end if
			
			pcv_SecurityPass = pcs_GetSecureKey
			pcv_SecurityKeyID = pcs_GetKeyID
			
			pCardNumber=ccnum
			ccnum=enDeCrypt(pCardNumber, pcv_SecurityPass)
			err.clear
			query="INSERT INTO pcPay_LinkPointAPI (idOrder, pcPay_LPAPI_ccnum, pcPay_LPAPI_ccexpmonth, pcPay_LPAPI_ccexpyear, pcPay_LPAPI_amount, pcPay_LPAPI_paymentmethod, pcPay_LPAPI_transtype, pcPay_LPAPI_authcode, idCustomer, pcPay_LPAPI_captured, pcPay_LPAPI_AuthorizedDate, pcPay_LPAPI_RTDate, pcSecurityKeyID) VALUES ("&pcGatewayDataIdOrder&",'"&ccnum&"','"&ccexpmonth&"','"&ccexpyear&"', "&pcBillingTotal&", 'LPAPI', '"&lp_transType&"', '"&session("GWTransId")&"', "&pcIdCustomer&", 0," & tmpStr & ", '"&R_TDate&"', "&pcv_SecurityKeyID&");"
			
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)		
			 
			if err.number<>0 then			
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			
			call closedb()
		End If		
		
		Response.redirect "gwReturn.asp?s=true&gw=LinkPointApi"
		Response.end
	 Else	
	  
	   	Response.redirect"msgb.asp?message="&server.URLEncode("<font color="&MType&"><b>Error</b>: "&r_error&"<br><br><a href="""&tempURL&"?psslurl=gwlpAPI.asp&idCustomer="&session("idcustomer")&"&idOrder="&session("GWOrderID")&"""><img src="""&rslayout("back")&""" border=0></a>")
     	Response.end
	 End If

	
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
					<% if pcPay_Cys_TestMode="0" then %>
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
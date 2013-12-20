<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/PPConstants.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="header.asp"-->
<%
'******************************************************************
'// PayPal Payflow Pro Currency
'******************************************************************
pcv_strPFPCurrency = "USD" '// "EUR"

dim pfpresolveTimeout, pfpconnectTimeout, pfpsendTimeout, pfpreceiveTimeout

pfpresolveTimeout	= 10000
pfpconnectTimeout	= 10000
pfpsendTimeout		= 10000
pfpreceiveTimeout	= 60000
'1000ms = 1 sec
%>

<% 'Gateway specific files %>
<div id="pcMain">
<table class="pcMainTable">
	<tr>
		<td><img src="images/checkout_bar_step5.gif" alt=""></td>
	</tr>
	<tr>
		<td class="pcSpacer"></td>
	</tr>
	<tr>
		<td>
			
		<% if session("GWOrderDone")="YES" then
			tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/default.asp"),"//","/")
			tempURL=replace(tempURL,"https:/","https://")
			tempURL=replace(tempURL,"http:/","http://") 
			session("GWOrderDone")=""
			response.redirect tempURL
		end if
		
		dim connTemp, rs
		%>
		<!-- #Include File="pcPay_Cent_XMLFunctions.asp"-->
		<!-- #Include File="pcPay_Cent_Utility.asp"-->

		<% session("redirectPage")="gwpfp.asp" %>
		
		<% Dim pcCustIpAddress
		pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
		If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
		
		dim tempURL
		If scSSL="" OR scSSL="0" Then
			tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
			tempURL=replace(tempURL,"https:/","https://")
			tempURL=replace(tempURL,"http:/","http://") 
		Else
			tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
			tempURL=replace(tempURL,"https:/","https://")
			tempURL=replace(tempURL,"http:/","http://")
		End If
		
		' Get Order ID
		if session("GWOrderId")="" then
			session("GWOrderId")=request("idOrder")
		end if
		
		pcGatewayDataIdOrder=session("GWOrderID")
		%>
		<!--#include file="pcGateWayData.asp"-->
		<%
		session("idCustomer")=pcIdCustomer
		
		If Request.Form("PaymentGWCentinel")="Go" OR request.QueryString("centinel")<>"" Then %>
			
			<!--#include file="pcCentinelInclude.asp"-->
			
			<%
			call opendb()
				
			query="SELECT v_Partner, v_Vendor, v_User, v_Password,pfl_testmode,pfl_transtype,pfl_CSC FROM verisign_pfp WHERE id=1;"
			set rs=Server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
		
			pPARTNER=rs("v_Partner")
			pVENDOR=rs("v_Vendor")
			pUSER=rs("v_User")
			pPWD=rs("v_Password")
			pfl_testmode=rs("pfl_testmode")
			pfl_transtype=rs("pfl_transtype")
			pfl_CSC=rs("pfl_CSC")
			
			if pfl_CSC="YES" then	 
				if not isnumeric(session("reqCVV")) or len(session("reqCVV")) < 3 or len(session("reqCVV")) > 4 Then				 
				  response.redirect "msgb.asp?message="&server.URLEncode("<b>Error</b>: "& dictLanguage.Item(Session("language")&"_paymntb_o_7")&dictLanguage.Item(Session("language")&"_paymntb_c_4")&"<br><br><a href="""&tempURL&"?psslurl="&session("redirectPage")&"&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"&ordertotal="&pcBillingTotal&"""><img src="""&rslayout("back")&""" border=0></a>")
			    End If
			End if 
			 
			set rs=nothing
	
			if pfl_testmode="YES" then
				GatewayHost = "https://pilot-payflowpro.paypal.com"
			else
				GatewayHost = "https://payflowpro.paypal.com"
			end if
			
			pcv_SecurityPass = pcs_GetSecureKey
			pcv_SecurityKeyID = pcs_GetKeyID
			
			call closedb()

			dim pCardNumber, pCardNumber2
			pCardNumber=session("reqCardNumber")
			pCardNumber2=enDeCrypt(pCardNumber, pcv_SecurityPass)

			if pcPay_Cent_Active=1 AND pfl_testmode="YES" then
				pcv_ExpDate="0111"
			else
				pcv_ExpDate= session("reqExpMonth")&session("reqExpYear")
			end if

			' Start out declaring our variables.
			Dim objWinHttp
			Dim strHTML
			Dim parmList
			Dim requestID

			' Need to generate a unique id for the request id
			requestID = generateRequestID()

			pcBillingFirstName=replace(pcBillingFirstName,"&amp;","&")
			pcBillingFirstName=replace(pcBillingFirstName,"&","and")
			pcBillingLastName=replace(pcBillingLastName,"&amp;","&")
			pcBillingLastName=replace(pcBillingLastName,"&","and")
			pcBillingCompany=replace(pcBillingCompany,"&amp;","&")
			pcBillingCompany=replace(pcBillingCompany,"&","and")
			pcShippingCity=replace(pcShippingCity,"&amp;","&")
			pcShippingCity=replace(pcShippingCity,"&","and")
			pcShippingFullName=replace(pcShippingFullName,"&amp;","&")
			pcShippingFullName=replace(pcShippingFullName,"&","and")
			pcShippingAddress=replace(pcShippingAddress,"&amp;","&")
			pcShippingAddress=replace(pcShippingAddress,"&","and")
			pcBillingCity=replace(pcBillingCity,"&amp;","&")
			pcBillingCity=replace(pcBillingCity,"&","and")
			pcBillingAddress=replace(pcBillingAddress,"&amp;","&")
			pcBillingAddress=replace(pcBillingAddress,"&","and")
			pcBillingPostalCode=replace(pcBillingPostalCode,"-","")
			pcBillingPostalCode=replace(pcBillingPostalCode," ","")
			
			if pcShippingFullName<>"" then
				pcShippingNameArry=split(pcShippingFullName, " ")
				if ubound(pcShippingNameArry)>0 then
					pcShippingFirstName=pcShippingNameArry(0)
					if ubound(pcShippingNameArry)>1 then
						 tmpShipFirstName = pcShippingFirstName&" "
						 pcShippingLastName = replace(pcShippingFullName,tmpShipFirstName,"")
					else
						pcShippingLastName=pcShippingNameArry(1)
					end if
				else
					pcShippingFirstName=pcShippingFullName
					pcShippingLastName=pcShippingFullName
				end if
			else
				pcShippingFirstName=pcBillingFirstName
				pcShippingLastName=pcBillingLastName
			end if

			'Build the parameter list

			'This a very, very basic implementation to just how how you can post data. What data
			'you decide to send and how your react to the response is a business decision that you
			'must make.
			parmList = "TENDER=C"
			parmList = parmList & "&ACCT=" & pCardNumber
			parmList = parmList & "&PWD=" & pPWD
			parmList = parmList & "&USER=" & pUSER
			parmList = parmList & "&VENDOR=" & pVENDOR
			parmList = parmList & "&PARTNER=" & pPARTNER
			parmList = parmList & "&EXPDATE=" & pcv_ExpDate
			parmList = parmList & "&AMT=" & money(pcBillingTotal)
			parmList = parmList & "&TRXTYPE=" & pfl_transtype
			if pfl_testmode="YES" then
				parmList=parmList & "&COMMENT1=" & "ASP/COM Test Transaction"
			else
				parmList=parmList & "&COMMENT1=" & "Web Store Transaction"
			end if
			parmList=parmList & "&COMMENT2=" & session("GWOrderId")
			if pfl_CSC="YES" then
				parmList=parmList & "&CVV2=" & session("reqCVV")
			end if
			parmList=parmList & "&FIRSTNAME="&Left(pcBillingFirstName,15)
			parmList=parmList & "&LASTNAME="&Left(pcBillingLastName,15)
			parmList = parmList & "&STREET=" & Left(pcBillingAddress,30)
			parmList = parmList & "&ZIP=" & pcBillingPostalCode
			parmList = parmList & "&CITY="&Left(pcBillingCity,20)
			parmList = parmList & "&BILLTOCOUNTRY="&Left(pcBillingCountryCode,30)
			parmList = parmList & "&CUSTCODE="&Left(pcIdCustomer,4)
			parmList = parmList & "&EMAIL="&Left(pcCustomerEmail,64)
			parmList = parmList & "&SHIPTOCITY="&Left(pcShippingCity,30)
			parmList = parmList & "&SHIPTOFIRSTNAME="&Left(pcShippingFirstName,30)
			parmList = parmList & "&SHIPTOLASTNAME="&Left(pcShippingLastName,30)
			parmList = parmList & "&SHIPTOSTATE="&Left(pcShippingState,10)
			parmList = parmList & "&SHIPTOSTREET="&Left(pcShippingAddress,30)
			parmList = parmList & "&SHIPTOZIP="&Left(pcShippingPostalCode,9)
			parmList = parmList & "&STATE="&Left(pcShippingState,2)
			parmList = parmList & "&SHIPTOCOUNTRY="&Left(pcShippingCountryCode,3)
			parmList = parmList & "&INVNUM=" & session("GWOrderId")
			parmList = parmList & "&PHONENUM="&Left(pcBillingPhone,20)
			
			if pcPay_Cent_Active=1 AND session("Centinel_ECI")<>"" AND pcPay_CardType="YES" Then
				parmList=parmList & "&AUTHENTICATION_STATUS=" & Session("Centinel_PAResStatus")
				parmList=parmList & "&XID["&len(session("Centinel_XID"))&"]=" & Session("Centinel_XID")
				parmList=parmList & "&CAVV["&len(session("Centinel_CAVV"))&"]=" & Session("Centinel_CAVV")
				parmList=parmList & "&ECI["&len(session("Centinel_ECI"))&"]=" & Session("Centinel_ECI")
			end if

			'Open Session
			Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
			objWinHttp.setTimeouts pfpresolveTimeout, pfpconnectTimeout, pfpsendTimeout, pfpreceiveTimeout
			objWinHttp.Open "POST", GatewayHost, False

			objWinHttp.setRequestHeader "Content-Type", "text/namevalue" ' for XML, use text/xml
			objWinHttp.SetRequestHeader "X-VPS-Timeout", "30"
			objWinHttp.SetRequestHeader "X-VPS-Request-ID", requestID

			'Send Parameter List
			objWinHttp.Send parmList

			' Print out the request status:
			'Response.Write "Status: " & objWinHttp.Status & " " & objWinHttp.StatusText & "<br />"

			' Get the text of the response.
			transResponse = objWinHttp.ResponseText

			' Trash our object now that we are finished with it.
			Set objWinHttp = Nothing

			'clear Centinel sessions
			Session("Centinel_Enrolled")=""
			Session("Centinel_ErrorNo")=""
			Session("Centinel_ErrorDesc")=""
			Session("Centinel_PAResStatus")=""
			Session("Centinel_SignatureVerification")=""
			Session("Centinel_ECI")=""
			Session("Centinel_XID")=""
			Session("Centinel_CAVV")=""
			Session("Centinel_ErrorNo")=""
			Session("Centinel_ErrorDesc")=""
			Session("Centinel_TransactionId")=""
			Session("Centinel_ACSURL")=""
			Session("Centinel_PAYLOAD")=""
			pfp_pnref = ShowResponse(transResponse, "PNREF")
			pfp_result = ShowResponse(transResponse, "RESULT")
			pfp_respmsg = ShowResponse(transResponse, "RESPMSG")
			pfp_authcode = ShowResponse(transResponse, "AUTHCODE")
			session("GWTransId")=pfp_pnref
			
			'/////////////////////////////////////////////////////
			'// Create Log of response and save in includes
			'/////////////////////////////////////////////////////
			dim pfpLogging
			pfpLogging=0 'Change to 1 to log
			
			if pfpLogging=1 then
				pCardNumberf4=left(pCardNumber,4)
				pCardNumberl4=right(pCardNumber,4)
				pCardNumberLog=pCardNumberf4&"********"&pCardNumberl4
				
				parmListLog = "TENDER=C"
				parmListLog = parmListLog & "&ACCT=" & pCardNumberLog
				parmListLog = parmListLog & "&PWD=" & pPWD
				parmListLog = parmListLog & "&USER=" & pUSER
				parmListLog = parmListLog & "&VENDOR=" & pVENDOR
				parmListLog = parmListLog & "&PARTNER=" & pPARTNER
				parmListLog = parmListLog & "&EXPDATE=" & pcv_ExpDate
				parmListLog = parmListLog & "&AMT=" & money(pcBillingTotal)
				parmListLog = parmListLog & "&TRXTYPE=" & pfl_transtype
				if pfl_testmode="YES" then
					parmListLog=parmListLog & "&COMMENT1=" & "ASP/COM Test Transaction"
				else
					parmListLog=parmListLog & "&COMMENT1=" & "Web Store Transaction"
				end if
				parmListLog=parmListLog & "&COMMENT2=" & session("GWOrderId")
				if pfl_CSC="YES" then
					parmListLog=parmListLog & "&CVV2=" & session("reqCVV")
				end if
				parmListLog=parmListLog & "&FIRSTNAME="&Left(pcBillingFirstName,15)
				parmListLog=parmListLog & "&LASTNAME="&Left(pcBillingLastName,15)
				parmListLog = parmListLog & "&STREET=" & Left(pcBillingAddress,30)
				parmListLog = parmListLog & "&ZIP=" & pcBillingPostalCode
				parmListLog = parmListLog & "&CITY="&Left(pcBillingCity,20)
				parmListLog = parmListLog & "&BILLTOCOUNTRY="&Left(pcBillingCountryCode,30)
				parmListLog = parmListLog & "&CUSTCODE="&Left(pcIdCustomer,4)
				parmListLog = parmListLog & "&EMAIL="&Left(pcCustomerEmail,64)
				parmListLog = parmListLog & "&SHIPTOCITY="&Left(pcShippingCity,30)
				parmListLog = parmListLog & "&SHIPTOFIRSTNAME="&Left(pcShippingFirstName,30)
				parmListLog = parmListLog & "&SHIPTOLASTNAME="&Left(pcShippingLastName,30)
				parmListLog = parmListLog & "&SHIPTOSTATE="&Left(pcShippingState,10)
				parmListLog = parmListLog & "&SHIPTOSTREET="&Left(pcShippingAddress,30)
				parmListLog = parmListLog & "&SHIPTOZIP="&Left(pcShippingPostalCode,9)
				parmListLog = parmListLog & "&STATE="&Left(pcShippingState,2)
				parmListLog = parmListLog & "&SHIPTOCOUNTRY="&Left(pcShippingCountryCode,3)
				parmListLog = parmListLog & "&INVNUM=" & session("GWOrderId")
				parmListLog = parmListLog & "&PHONENUM="&Left(pcBillingPhone,20)
	
				if PPD="1" then
					pcStrFileName=Server.Mappath ("/"&scPcFolder&"/includes/PFPLOG.txt")
				else
					pcStrFileName=Server.Mappath ("../includes/PFPLOG.txt")
				end if
	
				dim strFileName
				dim fs
				dim OutputFile
				
				'Specify directory and file to store silent post information
				strFileName = pcStrFileName
				Set fs = CreateObject("Scripting.FileSystemObject")
				Set OutputFile = fs.OpenTextFile (strFileName, 8, True)

				OutputFile.WriteLine now()
				OutputFile.WriteLine "=============HEADER======================="
				OutputFile.WriteLine """Content-Type"", ""text/namevalue"""
				OutputFile.WriteLine """X-VPS-Timeout"", ""30"""
				OutputFile.WriteLine """X-VPS-Request-ID"", "&requestID
				OutputFile.WriteLine "=============NVP=================="
				OutputFile.WriteLine "Request from ProductCart: " & parmListLog
				OutputFile.WriteBlankLines(1)
	
	
				OutputFile.WriteLine "Response from PayFlow Pro: " & transResponse
				OutputFile.WriteBlankLines(1)
	   
				OutputFile.WriteLine "PNREF: " & pfp_pnref
				OutputFile.WriteLine "RESULT: " & pfp_result
				OutputFile.WriteLine "RESPMSG: " & pfp_respmsg
				OutputFile.WriteLine "AUTHCODE: " & pfp_authcode
				OutputFile.WriteBlankLines(2)
	
				OutputFile.Close
			end if
			'/////////////////////////////////////////////////////
			'// End - Create Log of response and save in includes
			'/////////////////////////////////////////////////////

			'Response.Write "RESULT = " & ShowResponse(transResponse, "RESULT") & "<br>"
			'Response.Write "PNREF = " & ShowResponse(transResponse, "PNREF") & "<br>"
			'Response.Write "RESPMSG = " & ShowResponse(transResponse, "RESPMSG") & "<br>"
			'Response.Write "AUTHCODE = " & ShowResponse(transResponse, "AUTHCODE") & "<br>"
			'response.End()

			'STOP IF THERE ARE ERRORS PARSING THE RESPONSE USING pfp_getvalue()!!!
			if err.number<>0 then
				query="Error parsing the response from the gateway"
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			
			session("GWAuthCode")=pfp_authcode
			session("GWTransType")=pfl_transtype

			Dim pfp_rd_successurl, pfp_rd_resultfailurl, pfp_rd_avsfailurl, pfp_rd_resultfail, pfp_rd_avsfail

			pfp_rd_successurl="gwReturn.asp?s=true&gw=PFPRO"
			pfp_rd_resultfail="sorry_pfp.asp"
			pfp_rd_avsfail="sorry_pfp.asp"
			pfp_rd_resultfailurl=pfp_rd_resultfail & "?pfpresult=" & pfp_result & "&pfprespmsg=" & pfp_respmsg
			pfp_rd_avsfailurl=pfp_rd_avsfail & "?pfpresult=avsfail&pfprespmsg=Address%20Verification%20Failed."
	
			pfp_rd_resultfailurl="msgb.asp?message="&server.URLEncode("<b>Error&nbsp;"&pfp_result&"</b>: "&pfp_respmsg&"<br><br><a href="""&tempURL&"?psslurl="&session("redirectPage")&"&idOrder="&session("GWOrderID")&"""><img src="""&rslayout("back")&"""></a>")

			If pfp_rd_resultfail <> "" and pfp_result <> "0" Then
				Response.Redirect pfp_rd_resultfailurl
			ElseIf pfp_rd_successurl <> "" and pfp_result="0" Then
				tOID=int(session("GWOrderID"))-scpre
				'save info in pfpOrders
				if pfl_transtype="A" then
					pfp_fullname=pcShippingFirstName&" "&pcShippingLastName
					
					call opendb()
					
					query="INSERT INTO pfporders (idOrder, amt,tender,trxtype,origid,acct,expdate,idCustomer,fullname,street,state,email,zip,captured,pcSecurityKeyID) VALUES ("&tOID&", "&pcBillingTotal&",'C','A','"&pfp_pnref&"','"&pCardNumber2&"','"&session("reqExpMonth")&session("reqExpYear")&"',"&session("idcustomer")&",'"&replace(pfp_fullname,"'","''")&"','"&replace(pcBillingAddress,"'","''")&"','"&pcBillingState&"','"&pcCustomerEmail&"','"&pcBillingPostalCode&"',0,"&pcv_SecurityKeyID&");"
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=connTemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					
					call closeDb()
				end if
				Response.Redirect pfp_rd_successurl
			End If
		ELSE
			call opendb()
			
			query="SELECT pfl_CSC FROM verisign_pfp WHERE id=1;"
			set rs=server.CreateObject("ADODB.Recordset")
			set rs=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			pfp_CSC=rs("pfl_CSC")
			set rs=nothing
			
			call closedb()
			%>

			<form action="gwpfp.asp" method="post" name="form1" id="form1" class="pcForms">
				<input type="hidden" name="PaymentGWCentinel" value="Go">
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
					<% if pcPay_ACH_TestMode=1 then %>
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
						<% if pfp_CSC="YES" then %>
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
						<% end if %>
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></p></td>
						<td><%=money(pcBillingTotal)%></td>
					</tr>
					
					<% if pcPay_Cent_Active=1 then %>
						<script LANGUAGE="JavaScript">
						function popUp(url) {
							popupWin=window.open(url,"win",'toolbar=0,location=0,directories=0,status=1,menubar=1,scrollbars=1,width=570,height=450');
							self.name = "mainWin"; }
						</script>
							
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
						 <tr> 
							<td colspan="2">
							<p><a href='javascript:popUp("pcPay_Cent_mcsc.asp")'><img src='images/pc_mcsc.gif' alt="MasterCard SecureCode - Learn More" border='0' /></a>&nbsp;&nbsp;<a href='javascript:popUp("pcPay_Cent_vbv.asp")'><img src='images/pc_vbv.gif' alt="Verified by Visa - Learn More" border='0'/></a></p>
							<p>Your card may be eligible or enrolled in Verified by Visa&#8482; or MasterCard&reg; SecureCode&#8482; payer authentication programs. After clicking the 'Continue' button, your Card Issuer may prompt you for your payer authentication password to complete your purchase.</p>
								<p>&nbsp;</p>
							</td>
						</tr>
					<% end if %>
						
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr> 
						<td colspan="2" align="center">
							<!--#include file="inc_gatewayButtons.asp"-->
						</td>
					</tr>
				</table>
			</form>
		<% end if %>
		</td>
	</tr>
</table>
</div>
<% 
Function generateRequestID()
	pcKeys_Key_ID= Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now) & session("idcustomer")
    generateRequestID = pcKeys_Key_ID
End Function

Function ShowResponse(transResponse, Param)

     curString = transResponse

     Do while Len(curString) <> 0

          if InStr(curString,"&") Then
               varString = Left(curString, InStr(curString , "&" ) -1)
          else
               varString = curString
          end if

          name = Left(varString, InStr(varString, "=" ) -1)
          value = Right(varString, Len(varString) - (Len(name)+1))

          if name = Param Then
               MyValue = value
               Exit Do
          end if

          if Len(curString) <> Len(varString) Then
               curString = Right(curString, Len(curString) - (Len(varString)+1))
          else
               curString = ""
          end if

     Loop
     ShowResponse = MyValue

End Function
%>
<!--#include file="footer.asp"-->
<%PmAdmin=9%>
<% 'void order
err.clear					
			err.number = 0
			Set objTest = Server.CreateObject("PayPal.Payments.Communication.PayflowNETAPI")			' 
			
			If Err.Number = 0 Then
				Set objTest = Nothing
				pfp_com ="PayflowNETAPI"
			Else			    
			    err.number = 0
			    objTest =Server.CreateObject("PFProCOMControl.PFProCOMControl.1")				  
			end if 
			
			If Err.Number <> 0 Then
				  Response.write "<div class=""pcErrorMessage"" >"
				  Response.write "ERROR: Required Object PayPal.Payments.Communication.PayflowNETAPI Not Found <BR> OR<BR>"
				  Response.write "ERROR: Alternative Required Legacy Object PFProCOMControl.PFProCOMControl.1 Not Found <BR>"
			      Response.write "<a href=""http://wiki.earlyimpact.com/productcart/payflow_pro"">Click Here To Read About How To Configure Pay Flow Pro.</a></div>"
			      Response.end  			
			end if 
	 
			
			
			if pfp_com = "PayflowNETAPI" Then
				Set pfp_client = Server.CreateObject("PayPal.Payments.Communication.PayflowNETAPI")				
			else			
				Set pfp_client=Server.CreateObject("PFProCOMControl.PFProCOMControl.1")		
			end if 
			
idOrder=Request.Form("idOrder"&r)
pOrderStatus=request.Form("orderstatus"&r)
pCheckEmail=request.Form("checkEmail"&r)
pFullName=request.Form("fullname"&r)
pStreet=request.Form("street"&r)
pZip=request.Form("zip"&r)
pState=request.Form("state"&r)
pOrigid=request.Form("origid"&r)
pAcct=request.Form("acct"&r)
pExpdate=request.Form("expdate"&r)
pEmail=request.Form("email"&r)
pIdCustomer=request.Form("idCustomer"&r)

pfp_parm="TRXTYPE=V&TENDER=C"
pfp_parm=pfp_parm & "&USER=" & pUSER
pfp_parm=pfp_parm & "&PWD=" & pPWD
pfp_parm=pfp_parm & "&PARTNER=" & pPARTNER
pfp_parm=pfp_parm & "&VENDOR=" & pVENDOR
pfp_parm=pfp_parm & "&ORIGID=" & pOrigid
pfp_parm=pfp_parm & "&AMT=" & pfpamount
pfp_parm=pfp_parm & "&STREET=" & pStreet
pfp_parm=pfp_parm & "&ZIP=" & pZip

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			  ''' check to see what object we are using and sprep the settings
			''''''''''''''''''''''''''''''''''''''''''''
			if pfp_com = "PayflowNETAPI" Then					
				
				if pfl_testmode="YES" then
					v_URL="pilot-payflowpro.paypal.com"
				else
					v_URL="payflowpro.paypal.com"
				end if		
				' Host port 
				port = 443
				' Transaction time out value in seconds
				timeOut = 45
				
				' Modify the following proxy details if required.
				set proxyAddress = nothing
				proxyPort = 0
				set proxyLogon = nothing
				set proxyPassword = nothing
				
				'Specify whether exception trace is enabled or not. when ON, response will also have entire stack
				'trace of exception if any. Default is OFF. 
				traceEnabled = "OFF"
				
				'DO NOT CHANGE THE FOLLOWING 4 LEGACY PARAMS
				set doNotChangeParam1 = nothing
				set doNotChangeParam2 = nothing
				set doNotChangeParam3 = nothing
				doNotChangeParam4 = true
				
				'Call SetParameters to set the connection parameters
				pfp_client.SetParameters  v_URL, port, timeOut, proxyAddress, proxyPort, proxyLogon, proxyPassword, traceEnabled, doNotChangeParam1, doNotChangeParam2, doNotChangeParam3, doNotChangeParam4
				' The SetParameters object replaces the original CreateContext of the COM SDK:
				reqId = pfp_client.GenerateRequestId()
	
				'Submit the transaction
				pfp_string = pfp_client.SubmitTransaction(pfp_parm,reqId)
				
				else' old object
				
				if pfl_testmode="YES" then
					v_URL="test-payflow.verisign.com"
				else
					v_URL="payflow.verisign.com"
				end if
				
				pfp_ctx=pfp_client.CreateContext(v_URL, 443, 30, "", 0, "", "")
				pfp_string=pfp_client.SubmitTransaction(pfp_ctx, pfp_parm, Len(pfp_parm))				
				pfp_client.DestroyContext (pfp_ctx)
			end if 
			Set pfp_client=Nothing
pfp_pnref=pfp_getvalue("PNREF", pfp_string)
session("pnref")=pfp_pnref
pfp_result=pfp_getvalue("RESULT", pfp_string)
pfp_respmsg=pfp_getvalue("RESPMSG", pfp_string)
pfp_authcode=pfp_getvalue("AUTHCODE", pfp_string)
session("authcode")=pfp_authcode
pfp_avsaddr=pfp_getvalue("AVSADDR", pfp_string)
pfp_avszip=pfp_getvalue("AVSZIP", pfp_string)
pfp_comment2=pfp_getvalue("COMMENT2", pfp_string)
pfp_amt=pfp_getvalue("AMT", pfp_string)

'if success add to success/void
if pfp_result = 0 then
	'CAPTURE NEW TRANSACTION, if no errors
	'Send the request to the PFP processor.
	Set pfp_client=Server.CreateObject("PFProCOMControl.PFProCOMControl.1")
	pfp_parm="TRXTYPE=S&TENDER=C"
	pfp_parm=pfp_parm & "&USER=" & pUSER
	pfp_parm=pfp_parm & "&PWD=" & pPWD
	pfp_parm=pfp_parm & "&PARTNER=" & pPARTNER
	pfp_parm=pfp_parm & "&VENDOR=" & pVENDOR
	pfp_parm=pfp_parm & "&COMMENT1=" & "ASP/COM Transaction"
	pfp_parm=pfp_parm & "&COMMENT2=" & (int(idOrder)+scPre)
	pfp_parm=pfp_parm & "&ACCT=" & pAcct
	pfp_parm=pfp_parm & "&EXPDATE=" & pExpdate
	pfp_parm=pfp_parm & "&AMT=" & curamount
	pfp_parm=pfp_parm & "&STREET=" & pStreet
	pfp_parm=pfp_parm & "&ZIP=" & pZip
	
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			  ''' check to see what object we are using and sprep the settings
			''''''''''''''''''''''''''''''''''''''''''''
			if pfp_com = "PayflowNETAPI" Then					
				
				if pfl_testmode="YES" then
					v_URL="pilot-payflowpro.paypal.com"
				else
					v_URL="payflowpro.paypal.com"
				end if		
				' Host port 
				port = 443
				' Transaction time out value in seconds
				timeOut = 45
				
				' Modify the following proxy details if required.
				set proxyAddress = nothing
				proxyPort = 0
				set proxyLogon = nothing
				set proxyPassword = nothing
				
				'Specify whether exception trace is enabled or not. when ON, response will also have entire stack
				'trace of exception if any. Default is OFF. 
				traceEnabled = "OFF"
				
				'DO NOT CHANGE THE FOLLOWING 4 LEGACY PARAMS
				set doNotChangeParam1 = nothing
				set doNotChangeParam2 = nothing
				set doNotChangeParam3 = nothing
				doNotChangeParam4 = true
				
				'Call SetParameters to set the connection parameters
				pfp_client.SetParameters  v_URL, port, timeOut, proxyAddress, proxyPort, proxyLogon, proxyPassword, traceEnabled, doNotChangeParam1, doNotChangeParam2, doNotChangeParam3, doNotChangeParam4
				' The SetParameters object replaces the original CreateContext of the COM SDK:
				reqId = pfp_client.GenerateRequestId()
	
				'Submit the transaction
				pfp_string = pfp_client.SubmitTransaction(pfp_parm,reqId)
				
				else' old object
				
				if pfl_testmode="YES" then
					v_URL="test-payflow.verisign.com"
				else
					v_URL="payflow.verisign.com"
				end if
				
				pfp_ctx=pfp_client.CreateContext(v_URL, 443, 30, "", 0, "", "")
				pfp_string=pfp_client.SubmitTransaction(pfp_ctx, pfp_parm, Len(pfp_parm))				
				pfp_client.DestroyContext (pfp_ctx)
			end if 
			Set pfp_client=Nothing
end if
err.clear %>

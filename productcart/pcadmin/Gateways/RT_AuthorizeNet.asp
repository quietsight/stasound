<%
'---Start Authorize.Net---
Function gwaEdit()
	call opendb()
	'request gateway variables and insert them into the authorizeNet table
	query= "SELECT x_Login,x_Password FROM authorizeNet where id=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	x_Login2=rs("x_Login")
	'decrypt
	x_Login2=enDeCrypt(x_Login2, scCrypPass)
	x_Password2=rs("x_Password")
	'decrypt
	x_Password2=enDeCrypt(x_Password2, scCrypPass)
	set rs=nothing
	x_Type=request.Form("x_Type")
	x_Login=request.Form("x_Login")
	if x_Login="" then
		x_Login=x_Login2
	end if
	'encrypt
	x_Login=enDeCrypt(x_Login, scCrypPass)
	x_Password=request.Form("x_Password")
	if x_Password="" then
		x_Password=x_Password2
	end if
	'encrypt
	x_Password=enDeCrypt(x_Password, scCrypPass)
	cardTypes=request.Form("cardTypes")
	x_Curcode=request.Form("x_Curcode")
	x_URLMethod="gwAuthorizeAIM.asp"
	x_AIMType=request.Form("x_AIMType")
	x_CVV=request.Form("x_CVV")
	x_eCheck=request.Form("x_eCheck")
	x_secureSource=request.Form("x_secureSource")
	x_eCheckPending=request.Form("x_eCheckPending")
	if x_eCheckPending&""="" then
		x_eCheckPending = 0
	end if
	x_testmode=request.Form("x_testmode")
	if x_testmode="YES" then
		x_testmode="1"
	else
		x_testmode="0"
	end if
	paymentNickName2=replace(request.Form("paymentNickName2"),"'","''")
	if paymentNickName2="" then
		paymentNickName2="Check"
	end if
	
	'check to see if centinel is activated
	pcPay_Cent_Active=request.Form("pcPay_Cent_Active")
	if pcPay_Cent_Active="YES" then
		pcPay_Cent_Active=1
	else
		pcPay_Cent_Active=0
	end if
	pcPay_Cent_TransactionURL=request.Form("pcPay_Cent_TransactionURL")
	pcPay_Cent_MerchantID=request.Form("pcPay_Cent_MerchantID")
	pcPay_Cent_ProcessorId=request.Form("pcPay_Cent_ProcessorId")
	pcPay_Cent_Password=request.Form("pcPay_Cent_Password")
	if pcPay_Cent_TransactionURL="" or pcPay_Cent_MerchantID="" OR	pcPay_Cent_ProcessorId="" OR	pcPay_Cent_Password="" then
		pcPay_Cent_Active=0
		query="UPDATE pcPay_Centinel SET pcPay_Cent_TransactionURL='"&pcPay_Cent_TransactionURL&"',	pcPay_Cent_MerchantID='"&pcPay_Cent_MerchantID&"',	pcPay_Cent_ProcessorId='"&pcPay_Cent_ProcessorId&"', pcPay_Cent_Active=0, pcPay_Cent_Password='"&pcPay_Cent_Password&"' WHERE pcPay_Cent_ID=1;"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=conntemp.execute(query)
		set rs=nothing
	else
		query="UPDATE pcPay_Centinel SET pcPay_Cent_TransactionURL='"&pcPay_Cent_TransactionURL&"',	pcPay_Cent_MerchantID='"&pcPay_Cent_MerchantID&"',	pcPay_Cent_ProcessorId='"&pcPay_Cent_ProcessorId&"', pcPay_Cent_Active="&pcPay_Cent_Active&", pcPay_Cent_Password='"&pcPay_Cent_Password&"' WHERE pcPay_Cent_ID=1;"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=conntemp.execute(query)
		set rs=nothing
	end if
	
	query="UPDATE authorizeNet SET x_Type='"&x_Type&"||"&cardTypes&"',x_Login='"&x_Login&"',x_Password='"&x_Password&"',x_Version='3.1',x_Curcode='"&x_Curcode&"',x_Method='AIM',x_AIMType='"&x_AIMType&"',x_CVV="&x_CVV&",x_testmode="&x_testmode&",x_eCheck="&x_eCheck&",x_secureSource="&x_secureSource&",x_eCheckPending="&x_eCheckPending&" where id=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",sslURL='"&x_URLMethod&"', priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	if x_eCheck="1" then
		query="SELECT * FROM payTypes WHERE gwCode=16"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if rs.eof then
			query="INSERT INTO payTypes (paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES ('AuthorizeCheck','gwAuthorizeCheck.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",16,'"&paymentNickName2&"')"
		else
			query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",paymentNickName='"&paymentNickName2&"' WHERE gwCode=16"
		end if
	else
		query="DELETE FROM payTypes WHERE gwCode=16"
	end if
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	call closedb()
end function

Function gwa()
	varCheck=1
	'request gateway variables and insert them into the authorizeNet table
	x_Curcode=request.Form("x_Curcode")
	if x_Curcode="" then
		x_Curcode="USD"
	end if
	x_Method="AIM"
	x_URLMethod="gwauthorizeAIM.asp"
	x_AIMType="KEY"
	x_CVV=request.Form("x_CVV")
	x_eCheck=request.Form("x_eCheck")
	x_secureSource=request.Form("x_secureSource")
	x_eCheckPending=request.Form("x_eCheckPending")
	if x_eCheckPending&""="" then
		x_eCheckPending = 0
	end if
	x_testmode=request.Form("x_testmode")
	if x_testmode="YES" then
		x_testmode="1"
	else
		x_testmode="0"
	end if
	x_Type=request.Form("x_Type")
	x_Password=request.Form("x_Password")
	'encrypt
	x_Password=enDeCrypt(x_Password, scCrypPass)
	cardTypes=request.Form("cardTypes")
	x_Login=request.Form("x_Login")
	If x_Login="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Authorize.Net as your payment gateway. <b>""Login ID""</b> is a required field.")
	End If
	'encrypt
	x_Login=enDeCrypt(x_Login, scCrypPass)
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
	paymentNickName2=replace(request.Form("paymentNickName2"),"'","''")
	if paymentNickName2="" then
		paymentNickName2="Check"
	end if

	'check to see if centinel is activated
	pcPay_Cent_Active=request.Form("pcPay_Cent_Active")
	if pcPay_Cent_Active="YES" then
		pcPay_Cent_Active=1
		pcPay_Cent_TransactionURL=request.Form("pcPay_Cent_TransactionURL")
		pcPay_Cent_MerchantID=request.Form("pcPay_Cent_MerchantID")
		pcPay_Cent_ProcessorId=request.Form("pcPay_Cent_ProcessorId")
		pcPay_Cent_Password=request.Form("pcPay_Cent_Password")
		if pcPay_Cent_TransactionURL="" or pcPay_Cent_MerchantID="" OR	pcPay_Cent_ProcessorId="" OR	pcPay_Cent_Password="" then
			response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Cardinal Centinel for Authorize.Net. <b>""Tranaction URL, Merchant ID and Process ID""</b> are all required fields.")
		else
			err.clear
			err.number=0
			
			call openDb() 
			
			query="UPDATE pcPay_Centinel SET pcPay_Cent_TransactionURL='"&pcPay_Cent_TransactionURL&"',	pcPay_Cent_MerchantID='"&pcPay_Cent_MerchantID&"',	pcPay_Cent_ProcessorId='"&pcPay_Cent_ProcessorId&"', pcPay_Cent_Active=1, pcPay_Cent_Password='"&pcPay_Cent_Password&"' WHERE pcPay_Cent_ID=1;"
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
	
	query="UPDATE authorizeNet SET x_Type='"&x_Type&"||"&cardTypes&"',x_Login='"&x_Login&"',x_Password='"&x_Password&"',x_Version='3.1',x_Curcode='"&x_Curcode&"',x_Method='AIM',x_AIMType='"&x_AIMType&"',x_CVV="&x_CVV&",x_testmode="&x_testmode&", x_eCheck="&x_eCheck&", x_secureSource="&x_secureSource&",x_eCheckPending="&x_eCheckPending&" WHERE id=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'Authorize.Net','"&x_URLMethod&"',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",1,'"&paymentNickName&"')"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	if x_eCheck="1" then
		query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'Authorize eCheck','gwAuthorizeCheck.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",16,'"&paymentNickName2&"')"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	end if
	
	set rs=nothing
	call closedb()
end function
%>
				
<% if request("gwchoice")="1" then
	if request("mode")="Edit" then
		call opendb()
		query= "SELECT x_Type,x_Login,x_Password,x_Curcode,x_AIMType,x_CVV,x_testmode,x_eCheck,x_secureSource,x_eCheckPending FROM authorizeNet where id=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		x_Type=rs("x_Type")
		x_Login=rs("x_Login")
		'decrypt
		x_Login=enDeCrypt(x_Login, scCrypPass)
		x_Password=rs("x_Password")
		'decrypt
		x_Password=enDeCrypt(x_Password, scCrypPass)
		x_Curcode=rs("x_Curcode")
		x_AIMType=rs("x_AIMType")
		x_CVV=rs("x_CVV")
		x_testmode=rs("x_testmode")
		x_eCheck=rs("x_eCheck")
		x_secureSource=rs("x_secureSource")
		x_eCheckPending=rs("x_eCheckPending")

		x_TypeArray=Split(x_Type,"||")
		x_Type1=x_TypeArray(0)
		M="0"
		V="0"
		A="0"
		D="0"
		if ubound(x_TypeArray)=1 then
			x_Type2=x_TypeArray(1)
			cardTypeArray=split(x_Type2,", ")
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
		end if
				
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=1"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=conntemp.execute(query)
		if rs.eof then
			paymentNickName="Credit Card"
		else
			pcv_processOrder=rs("pcPayTypes_processOrder")
			pcv_setPayStatus=rs("pcPayTypes_setPayStatus")
			priceToAdd=rs("priceToAdd")
			percentageToAdd=rs("percentageToAdd")
			paymentNickName=rs("paymentNickName")
			if percentageToAdd<>"0" then
				priceToAddType="percentage"
			end if
			if priceToAdd<>"0" then
				priceToAddType="price"
			end if
		end if


		query= "SELECT paymentNickName FROM payTypes WHERE gwCode=16"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=conntemp.execute(query)
		if rs.eof then
			paymentNickName2="Check"
		else
			paymentNickName2=rs("paymentNickName")
		end if

		query="SELECT pcPay_Centinel.pcPay_Cent_TransactionURL, pcPay_Centinel.pcPay_Cent_ProcessorId, pcPay_Centinel.pcPay_Cent_MerchantID, pcPay_Centinel.pcPay_Cent_Active, pcPay_Centinel.pcPay_Cent_Password FROM pcPay_Centinel WHERE (((pcPay_Centinel.pcPay_Cent_ID)=1));"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		pcPay_Cent_TransactionURL = rs("pcPay_Cent_TransactionURL")
		pcPay_Cent_ProcessorId = rs("pcPay_Cent_ProcessorId")
		pcPay_Cent_MerchantID = rs("pcPay_Cent_MerchantID")
		pcPay_Cent_Active=rs("pcPay_Cent_Active")
		pcPay_Cent_Password = rs("pcPay_Cent_Password")
		
		set rs=nothing
		call closedb()
		if x_Curcode="" then
			x_Curcode="USD"
		end if
		%>
    
		<input type="hidden" name="mode" value="Edit">
    <% end if %>
	<input type="hidden" name="addGw" value="1">
    <input type="hidden" name="x_AIMType" value="KEY">
    <input type="hidden" name="x_Method" value="AIM">
	<% If x_eCheck=1 Then %>
    <% Else %>
        <input name="paymentNickName2" value="" type="hidden" />
    <% End If %>
    <% if intCentActive<>0 AND request("mode")<>"Edit" then %>
    <% else %>
        <% if trim(pcPay_Cent_TransactionURL)="" then
            pcPay_Cent_TransactionURL="https://centineltest.cardinalcommerce.com/maps/txns.asp"
        end if %>
        <% if pcPay_Cent_MerchantID<>"" then
            pcPay_Cent_MerchantID=replace(pcPay_Cent_MerchantID,"""","&quot;")
        end if %>
        <% if pcPay_Cent_ProcessorID<>"" then
            pcPay_Cent_ProcessorID=replace(pcPay_Cent_ProcessorID,"""","&quot;")
        end if %>
        <% if pcPay_Cent_Password<>"" then
            pcPay_Cent_Password=replace(pcPay_Cent_Password,"""","&quot;")
        end if %>
    <% end if %>

<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/authorizenet.png" width="234" height="57"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
<table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
<tr>
    <td><strong>Enter your Gateway Account information<br>
    <br>
    </strong>ProductCart uses Authorize.Net AIM (Advanced Integration Method / Direct Response) to communicate with the payment gateway. Customers never leave your Web store and credit cards are validated in real-time according to the settings you have set in the Authorize.Net Merchant Center (e.g. address verification on or off). Please note that you must have SSL enabled to process transactions through AIM.<strong><br>
    <br>
    <a href="http://reseller.authorize.net/application.asp?id=220675" target="_blank">Sign Up Now!</a></strong><br />
<br />
</td>
</tr>
<tr>
    <td>
        <div id="CollapsiblePanel1" class="CollapsiblePanel">
            <div class="CollapsiblePanelTab1">
                <table width="100%">
                  <tr>
                    <td width="580" class="pcPanelTitle1">Step 1: Configure Account</td>
                  </tr>
                </table>
            </div>
            <div class="CollapsiblePanelContent">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <% if request("mode")="Edit" then %>
					<% dim A_LoginCnt,A_LoginEnd,A_LoginStart
                    A_LoginCnt=(len(x_Login)-2)
                    A_LoginEnd=right(x_Login,2)
                    A_LoginStart=""
                    for c=1 to A_LoginCnt
                        A_LoginStart=A_LoginStart&"*"
                    next %>
                    <tr> 
                        <td colspan="2">Current Login ID:&nbsp;<%=A_LoginStart&A_LoginEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;Login ID&quot; is only partially shown on this page. If you need to edit your account information, please re-enter your &quot;Login ID&quot; below.</td>
                    </tr>
                <% end if %>
                <tr>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                    <td width="205"> <div align="right">Login ID:</div></td>
                    <td width="574"> <input type="text" name="x_Login" size="30"></td>
                </tr>
                <tr> 
                    <td> <div align="right">Transaction Key:</div></td>
                    <td> <input name="x_Password" type="text" size="30"> </td>
                </tr>
                <tr> 
                    <td> <div align="right">Transaction Type:</div></td>
                    <td> <select name="x_Type">
                            <option value="AUTH_CAPTURE" selected="selected">Sale</option>
                            <option value="AUTH_ONLY" <% if x_Type1="AUTH_ONLY" then %>selected<% end if %>>Authorize Only</option>
                        </select> </td>
                </tr>
                <tr> 
                    <td><div align="right">Currency Code:</div></td>
                    <td><input name="x_Curcode" type="text" value="<%=x_Curcode%>" size="6" maxlength="4"> 
                        <a href="help_auth_codes.asp" target="_blank">Find Codes</a></td>
                </tr>
                <tr> 
                    <td> <div align="right">Require CVV:</div></td>
                    <td> <input type="radio" class="clearBorder" name="x_CVV" value="1" checked>
                        Yes 
                        <input name="x_CVV" type="radio" class="clearBorder" value="0" <% if x_CVV="0" then %>checked<%end if %>>
                        No</td>
                </tr>
                <tr> 
                    <td> <div align="right">Accepted Cards:</div></td>
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
                    <td> <div align="right">SecureSource:</div></td>
                    <td> <input type="radio" class="clearBorder" name="x_secureSource" value="1" checked> Yes 
                        <input name="x_secureSource" type="radio" class="clearBorder" value="0" <% if x_secureSource=0 then%>checked<% end if %>> No (Select Yes if you are a Wells Fargo SecureSource Merchant)</td>
                </tr>
            	<tr>
                	<td colspan="2" class="pcCPspacer"><hr /></td>
                </tr>
                <% If x_eCheck=1 Then %> 
                    <tr> 
                        <td> <div align="right">Accept Checks:</div></td>
                        <td> <input type="radio" class="clearBorder" name="x_eCheck" value="1" checked>
                            Yes 
                            <input name="x_eCheck" type="radio" class="clearBorder" value="0" <% if x_eCheck=0 then%>checked<% end if %>>
                            No</td>
                    </tr>
                    <tr> 
                        <td colspan="2">All eCheck orders are automatically considered &quot;Processed&quot; by ProductCart, as the order amount is always debited to the customer's bank account. If for any reasons you would like eCheck orders to be considered &quot;Pending&quot;, use this option. Should eCheck orders be considered &quot;Pending&quot;? 
                            <input type="radio" class="clearBorder" name="x_eCheckPending" value="1" checked> Yes 
                            <input name="x_eCheckPending" type="radio" class="clearBorder" value="0" <% if x_eCheckPending=0 then%>checked<% end if %>> No</td>
                    </tr>
                <% Else %>
                	<input name="x_eCheck" type="hidden" value="0" /> 
                <% End If %>
            	<tr>
                	<td colspan="2" class="pcCPspacer"><hr /></td>
                </tr>
                <tr> 
                    <td><div align="right"> 
                            <input name="x_testmode" type="checkbox" class="clearBorder" value="YES" <% if x_testmode=1 then%>checked<% end if%>> 
                        </div></td>
                    <td><b>Enable Test Mode </b>(Credit cards will not be charged)</td>
                </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td class="pcSubmenuContent">&nbsp;</td>
                  </tr>
                </table>
            </div>
        </div>
        <div id="CollapsiblePanel2" class="CollapsiblePanel">
            <div class="CollapsiblePanelTab1">
                        <table width="100%">
                        <tr>
                        <td class="pcPanelTitle1">Step 2: You have the option to charge a processing fee for this payment option.</td>
                            <td width="24" class="pcPanelTitle1" align="right"><img src="images/expand.gif" width="19" height="19" alt="Expand Selection" /></td>
                        </tr>
                        </table>
                    </div>
                    <div class="CollapsiblePanelContent">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                                <td nowrap="nowrap">&nbsp;</td>
                                <td class="pcSubmenuContent">&nbsp;</td>
                            </tr>
                            <tr>
                                <td width="7%" nowrap="nowrap"><div align="left">Processing Fee:&nbsp;</div></td>
                                <td>
                                <input type="radio" class="clearBorder" name="priceToAddType" value="price" <% if priceToAddType="price" then%>checked<%end if%>>Flat fee&nbsp;&nbsp; &nbsp;<%=scCurSign%> <input name="priceToAdd" size="6" value="<%=money(priceToAdd)%>"></td>
                            </tr>
                            <tr>
                                <td>&nbsp;</td>
                                <td><input type="radio" class="clearBorder" name="priceToAddType" value="percentage" <% if priceToAddType="percentage" then%>checked<%end if%>>Percentage of Order Total&nbsp;&nbsp; &nbsp; %<input name="percentageToAdd" size="6" value="<%=percentageToAdd%>"></td>
                            </tr>
                            <tr>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                            </tr>
                        </table>
            </div>
        </div>
        <div id="CollapsiblePanel3" class="CollapsiblePanel">
            <div class="CollapsiblePanelTab1">
                <table width="100%">
                  <tr>
                            <td class="pcPanelTitle1">Step 3: You can change the display name that is shown for this payment type. </td>
                            <td width="24" class="pcPanelTitle1" align="right"><img src="images/expand.gif" width="19" height="19" alt="Expand Selection" /></td>
                  </tr>
                </table>
            </div>
            <div class="CollapsiblePanelContent">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td nowrap="nowrap">&nbsp;</td>
                              <td>&nbsp;</td>
                            </tr>
                            <tr> 
                                <td width="10%" nowrap="nowrap"><div align="left">Payment Name:&nbsp;</div></td>
                                        <td width="90%"><input name="paymentNickName" value="<%=paymentNickName%>" size="35" maxlength="255"></td>
                            </tr>
			                <% If x_eCheck=1 Then %> 
			                    <tr> 
			                        <td width="10%" nowrap="nowrap"><div align="left"><strong>eCheck Payment Name</strong>:&nbsp;</div></td>
			                        <td><input name="paymentNickName2" value="<%=paymentNickName2%>" size="35" maxlength="255"></td>
			                    </tr>
			               	<% End If %>
                    <tr>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                </table>
            </div>
        </div>
        <div id="CollapsiblePanel4" class="CollapsiblePanel">
            <div class="CollapsiblePanelTab1">
                <table width="100%">
                  <tr>
                            <td class="pcPanelTitle1">Step 4: Order Processing: Order Status and Payment Status</td>
                            <td width="24" class="pcPanelTitle1" align="right"><img src="images/expand.gif" width="19" height="19" alt="Expand Selection" /></td>
                          </tr>
                        </table>
                    </div>
                    <div class="CollapsiblePanelContent">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td>&nbsp;</td>
                            </tr>
                            <tr>
                                        <td>Process orders when they are placed: <input type="checkbox" class="clearBorder" name="pcv_processOrder" value="1" <%if pcv_processOrder="1" then%>checked<%end if%>>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=301')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
                    </tr>
                    <tr> 
                        <td>When orders are placed, set the payment status to:
                        <select name="pcv_setPayStatus">
                            <option value="3" selected="selected">Default</option>
                                    <option value="0" <%if pcv_setPayStatus="0" then%>selected<%end if%>>Pending</option>
                                    <option value="1" <%if pcv_setPayStatus="1" then%>selected<%end if%>>Authorized</option>
                                    <option value="2" <%if pcv_setPayStatus="2" then%>selected<%end if%>>Paid</option>
                        </select>
                        &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=302')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>					</td>
                    </tr>
                    <tr>
                      <td>&nbsp;</td>
                    </tr>
                </table>
            </div>
        </div>
        <div id="CollapsiblePanel5" class="CollapsiblePanel">
            <div class="CollapsiblePanelTab1">
                <table width="100%">
                  <tr>
                            <td class="pcPanelTitle1">Step 5: Minimize fraud by enabling Centinel by CardinalCommerce</td>
                            <td width="24" class="pcPanelTitle1" align="right"><img src="images/expand.gif" width="19" height="19" alt="Expand Selection" /></td>
                  </tr>
                </table>
            </div>
            <div class="CollapsiblePanelContent">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td colspan="2">&nbsp;</td>
                </tr>
                <tr> 
                    <td colspan="2">Note: Additional charges apply. <a href="http://billing.cardinalcommerce.com/centinel/registration/frame_services.asp?RefId=PRDCTCART" target="_blank">Contact CardinalCommerce for more information &gt;&gt;</a></td>
                </tr>
                <% if intCentActive<>0 AND request("mode")<>"Edit" then %>
                    <tr>
                        <td colspan="2">Centinel has already been activated for this or another payment gateway. To edit its settings or remove it, simply activate this payment gateway and then click on the &quot;Modify&quot; button on the payment options summary page.</td>
                    </tr>
                <% else %>
                    <tr> 
                                    <td width="14%">&nbsp;</td>
                                    <td width="86%"><input name="pcPay_Cent_Active" type="checkbox" class="clearBorder" value="YES" <%if pcPay_Cent_Active=1 then%>checked<%end if%>>
                                    <strong> Enable Centinel for USAePay</strong></td>
                                </tr>
                                <% if trim(pcPay_Cent_TransactionURL)="" then
                                    pcPay_Cent_TransactionURL="https://centineltest.cardinalcommerce.com/maps/txns.asp"
                                end if %>
                                <tr> 
                                    <td nowrap="nowrap"><div align="left">Transaction Url:</div></td>
                                    <td><input name="pcPay_Cent_TransactionURL" size="60" maxlength="255" value="<%=pcPay_Cent_TransactionURL%>"></td>
                                </tr>
                                <% if pcPay_Cent_MerchantID<>"" then
                                    pcPay_Cent_MerchantID=replace(pcPay_Cent_MerchantID,"""","&quot;")
                                end if %>
                                <tr> 
                                    <td><div align="left">Merchant ID: </div></td>
                                    <td><input name="pcPay_Cent_MerchantID" size="35" maxlength="255" value="<%=pcPay_Cent_MerchantID%>"></td>
                                </tr>
                                <% if pcPay_Cent_ProcessorID<>"" then
                                    pcPay_Cent_ProcessorID=replace(pcPay_Cent_ProcessorID,"""","&quot;")
                                end if %>
                                <tr> 
                                    <td><div align="left">Processor ID: </div></td>
                                    <td><input name="pcPay_Cent_ProcessorId" size="35" maxlength="255" value="<%=pcPay_Cent_ProcessorID%>"></td>
                                </tr>
                                <tr> 
                                    <td><div align="left">Password: </div></td>
                                    <td><input name="pcPay_Cent_Password" size="35" maxlength="255" value="<%=pcPay_Cent_Password%>"></td>
                                </tr>
							<% end if %>
                    	</table>
        </div>
              	</div>
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr>
                        <td colspan="2">
                        <% if request("mode")="Edit" then
                            strButtonValue="Save Changes" %>
                            <input type="hidden" name="submitMode" value="Edit">
                        <%  else
                            strButtonValue="Add New Payment Method" %>
                            <input type="hidden" name="submitMode" value="Add Gateway">
                        <% end if %>
                        <input type="submit" value="<%=strButtonValue%>" name="Submit" class="submit2"> 
                        &nbsp;
                        <input type="button" value="Back" onclick="javascript:history.back()">
                        </td>
                    </tr>
				</table>
        <script type="text/javascript">
        <!--
        var CollapsiblePanel1 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel1", {contentIsOpen:true});
        var CollapsiblePanel2 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel2", {contentIsOpen:false});;
        var CollapsiblePanel3 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel3", {contentIsOpen:false});
        var CollapsiblePanel4 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel4", {contentIsOpen:false});
        var CollapsiblePanel5 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel5", {contentIsOpen:false});
        //-->
        </script>
    </td>
</tr>
</table>
<!-- New View End --><% end if %>

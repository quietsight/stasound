<%
'--- Start PayPal Payments Advanced ---
Function gwPPPUKEdit()
	call opendb()

	pcPay_PayPal_TransType=request.Form("pcPay_PayPal_TransType")
	pcPay_PayPal_Username=request.Form("pcPay_PayPal_Username")
	pcPay_PayPal_Subject=request.Form("pcPay_PayPal_Subject")
	pcPay_PayPal_Password=request.Form("pcPay_PayPal_Password")
	pcPay_PayPal_Sandbox=request.Form("pcPay_PayPal_Sandbox")
	pcPay_PayPal_Vendor=request.Form("pcPay_PayPal_Vendor")
	pcPay_PayPal_Partner=request.Form("pcPay_PayPal_Partner")
	pcPay_PayPal_Signature=request.Form("pcPay_PayPal_Signature")					
	pcPay_PayPal_Currency=request.Form("pcPay_PayPal_Currency")		
	pcPay_PayPal_CVC=request.Form("pcPay_PayPal_CVC")	
	pcPay_PayPal_CardTypes=request.Form("pcPay_PayPal_CardTypes")
	PayPalPaymentURL="gwPayPalUK.asp"
	PayPalName="PayPal Website Payments Pro"	
	ppGwcode=53			
	if pcPay_PayPal_Sandbox="YES" then
		pcPay_PayPal_Sandbox=1
	else
		pcPay_PayPal_Sandbox=0
	end if
	if pcPay_PayPal_CVC="" then
		pcPay_PayPal_CVC=0
	end if
	ppAB=1

	query="UPDATE pcPay_PayPal SET pcPay_PayPal_TransType="&pcPay_PayPal_TransType&", pcPay_PayPal_Username='"&pcPay_PayPal_Username&"', pcPay_PayPal_Subject='"&pcPay_PayPal_Subject&"', pcPay_PayPal_Password='"&pcPay_PayPal_Password&"', pcPay_PayPal_Sandbox="&pcPay_PayPal_Sandbox&", pcPay_PayPal_Signature='"&pcPay_PayPal_Signature&"', pcPay_PayPal_Currency='"&pcPay_PayPal_Currency&"', pcPay_PayPal_CVC="&pcPay_PayPal_CVC&", pcPay_PayPal_Vendor='"&pcPay_PayPal_Vendor&"', pcPay_PayPal_Partner='"&pcPay_PayPal_Partner&"', pcPay_PayPal_CardTypes='"&pcPay_PayPal_CardTypes&"' WHERE (((pcPay_PayPal_ID)=1));"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",paymentDesc='"&PayPalName&"',sslUrl='',priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"',pcPayTypes_ppab=1 WHERE gwCode=53"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If

	set rs=nothing
	call closedb()
end function

Function gwPPPUK()
	varCheck=1
	'request gateway variables and insert them into the database
	pcPay_PayPal_TransType=request.Form("pcPay_PayPal_TransType")
	pcPay_PayPal_Username=request.Form("pcPay_PayPal_Username")
	pcPay_PayPal_Subject=request.Form("pcPay_PayPal_Subject")
	pcPay_PayPal_Password=request.Form("pcPay_PayPal_Password")
	pcPay_PayPal_Sandbox=request.Form("pcPay_PayPal_Sandbox")
	pcPay_PayPal_Vendor=request.Form("pcPay_PayPal_Vendor")
	pcPay_PayPal_Partner=request.Form("pcPay_PayPal_Partner")
	pcPay_PayPal_Signature=request.Form("pcPay_PayPal_Signature")					
	pcPay_PayPal_Currency=request.Form("pcPay_PayPal_Currency")		
	pcPay_PayPal_CVC=request.Form("pcPay_PayPal_CVC")	
	pcPay_PayPal_CardTypes=request.Form("pcPay_PayPal_CardTypes")		
	PayPalPaymentURL="gwPayPalUK.asp"
	PayPalName="PayPal Website Payments Pro"	
	ppGwcode=53			
	if pcPay_PayPal_Sandbox="YES" then
		pcPay_PayPal_Sandbox=1
	else
		pcPay_PayPal_Sandbox=0
	end if
	if pcPay_PayPal_CVC="" then
		pcPay_PayPal_CVC=0
	end if
	ppAB=1

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
			
	err.clear
	err.number=0
			
	call openDb() 

	query="UPDATE pcPay_PayPal SET pcPay_PayPal_TransType="&pcPay_PayPal_TransType&", pcPay_PayPal_Username='"&pcPay_PayPal_Username&"', pcPay_PayPal_Subject='"&pcPay_PayPal_Subject&"', pcPay_PayPal_Password='"&pcPay_PayPal_Password&"', pcPay_PayPal_Sandbox="&pcPay_PayPal_Sandbox&", pcPay_PayPal_Signature='"&pcPay_PayPal_Signature&"', pcPay_PayPal_Currency='"&pcPay_PayPal_Currency&"', pcPay_PayPal_CVC="&pcPay_PayPal_CVC&", pcPay_PayPal_Vendor='"&pcPay_PayPal_Vendor&"', pcPay_PayPal_Partner='"&pcPay_PayPal_Partner&"', pcPay_PayPal_CardTypes='"&pcPay_PayPal_CardTypes&"' WHERE (((pcPay_PayPal_ID)=1));"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName, pcPayTypes_ppab) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'PayPal','gwPayPalUK.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",53,'"&paymentNickName&"',1)"
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
end function
%>
				
<% if request("gwchoice")="53" then
	pcConflictIdPayment = 0

	if request("mode")="Edit" then
		call opendb()

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
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=53"
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

		set rs=nothing
		call closedb()
		%>
		<input type="hidden" name="mode" value="Edit">
    <% else
		'//Check if any other PayPal Services are activated.
		call opendb()
		query= "SELECT idPayment, gwCode FROM payTypes WHERE gwCode=3 OR gwCode=46 OR gwCode=80"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=conntemp.execute(query)
		if NOT rs.eof then
			pcConflictIdPayment = rs("idPayment")
			pcConflictID = rs("gwCode")
			select case pcConflictID
				case "3"
					pcConflictDesc = "PayPal Payments Standard"
				case "46"
					pcConflictDesc = "PayPal Payments Pro"
				case "80"
					pcConflictDesc = "PayPal Payments Advanced"
			end select %>
        	<div class="pcCPmessage">
        	  <p>You currently have <strong><%=pcConflictDesc%></strong> active for this store. In order to use <strong>Website Payments Pro</strong> you will need to first disable <strong><%=pcConflictDesc%></strong>.<br />
        	    <br />
        	  </p>
        	  <p><a href="pcConfigurePayment.asp?mode=Del&id=<%=pcConflictIdPayment%>&gwchoice=<%=pcConflictID%>&activate=53">Disable <%=pcConflictDesc%> and enable Website Payments Pro</a></p>
        	  <br />
        	  <p><a href="pcPaymentSelection.asp">Back to payment selection</a><br />
        	    <br />
      	    </p>
            </div>
<% end if
		set rs = nothing
		call closedb()
	end if %>
<% if pcConflictIdPayment = 0 then %>
	<input type="hidden" name="addGw" value="53">
    <div class="pcCPmessageSuccess">
        <% if request("mode")="Edit" then %>
            <p><strong>You're editing Website Payments Pro</strong>
            <br />
            <br />
            <p><strong><a href="pcPaymentSelection.asp">Change Payment Option</a></strong></p>
            <br /></p>
        <% else %>
            <p><strong>You've selected Website Payments Pro</strong><br />
            <br />Fully customize your checkout pages and
            accept credit cards directly on your site. PayPal
            simplifies PCI compliance for you, plus you get
            Virtual Terminal at no added cost.<strong><br>
            <br>
            <a href="https://www.paypal.
            com/webapps/mpp/paypal-payments-pro" target="_blank">Sign Up and Learn More</a></strong>
            </p>
            <br />
            <p>To start accepting payments, please complete the process below. <strong><a href="pcPaymentSelection.asp">Change Payment Option</a></strong></p>
            <br /></p>
        <% end if %>
    </div>
	<table width="100%">
        <tr>
            <td align="left" style="font-size:15px;"><img src="gateways/logos/paypal_logo1.gif" width="253" height="80" /></td>
            <td align="left" style="font-size:15px;">&nbsp;</td>
        </tr>
    </table>
    <br>
    	<table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
        <tr>
            <td>
                <div id="CollapsiblePanel1" class="CollapsiblePanel">
                    <div class="CollapsiblePanelTab1">
                        <table width="100%">
                          <tr>
                            <td width="580" class="pcPanelTitle1"><strong>Configure Account - Website Payments Pro (UK)</strong></td>
                          </tr>
                        </table>
                    </div>
                    <div class="CollapsiblePanelContent">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                    <td colspan="2" class="pcCPspacer"></td>
                                </tr>
								<tr>
								  <td colspan="2" valign="top">
                                    <p>You must have a PayPal business account to use Website Payments Pro. If you don't have an account, <a href="#">sign up for one now</a>.</p><br />
<p>You created this information when you signed
up for Website Payments Pro. Enter it here to connect your account and allow payments. (Note: This is also your login information for <a href="http://manager.paypal.com" target="_blank">PayPal Manager</a>.)</p><br /></td>
						  </tr>
								<tr> 
									<td width="127" valign="top"><div align="right">Partner Name:</div></td>
									<td>
									<input type="text" value="<%=pcPay_PayPal_Partner%>" name="pcPay_PayPal_Partner" size="30" maxlength="250">
									<a href="JavaScript:win('helpOnline.asp?ref=806')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature" /></a></td>
								</tr>
								<tr> 
									<td width="127" valign="top"><div align="right">Merchant Login:</div></td>
									<td>
									<input type="text" value="<%=pcPay_PayPal_Vendor%>" name="pcPay_PayPal_Vendor" size="30" maxlength="250">
									<a href="JavaScript:win('helpOnline.asp?ref=806')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature" /></a><a href="JavaScript:win('helpOnline.asp?ref=805')"></a></td>
								</tr>
								<tr> 
									<td width="127" valign="top"><div align="right">User:</div></td>
									<td><input type="text" value="<%=pcPay_PayPal_Username%>" name="pcPay_PayPal_Username" size="30" maxlength="50">
							        <a href="JavaScript:win('helpOnline.asp?ref=806')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature" /></a><a href="JavaScript:win('helpOnline.asp?ref=805')"></a></td>
								</tr>
								
								<tr> 
									<td valign="top"><div align="right">Password:</div></td>
									<td><input type="text" value="<%=pcPay_PayPal_Password%>" name="pcPay_PayPal_Password" size="30" maxlength="50">
							        <a href="JavaScript:win('helpOnline.asp?ref=806')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature" /></a><a href="JavaScript:win('helpOnline.asp?ref=805')"></a></td>
								</tr>
                                <% 
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
								<tr> 
									<td width="127" valign="top" nowrap><div align="right">Card Types:</div></td>
									<td>
										<% if V="1" then %>
                                            <input name="pcPay_PayPal_CardTypes" type="checkbox" class="clearBorder" id="pcPay_PayPal_CardTypes" value="V" checked> 
                                        <% else %>
                                            <input name="pcPay_PayPal_CardTypes" type="checkbox" class="clearBorder" id="pcPay_PayPal_CardTypes" value="V"> 
                                        <% end if %> Visa 
                                        <% if M="1" then %>
                                            <input name="pcPay_PayPal_CardTypes" type="checkbox" class="clearBorder" id="pcPay_PayPal_CardTypes" value="M" checked> 
                                        <% else %>
                                            <input name="pcPay_PayPal_CardTypes" type="checkbox" class="clearBorder" id="pcPay_PayPal_CardTypes" value="M"> 
                                        <% end if %> MasterCard 
                                        <% if A="1" then %>
                                            <input name="pcPay_PayPal_CardTypes" type="checkbox" class="clearBorder" id="pcPay_PayPal_CardTypes" value="A" checked> 
                                        <% else %>
                                            <input name="pcPay_PayPal_CardTypes" type="checkbox" class="clearBorder" id="pcPay_PayPal_CardTypes" value="A"> 
                                        <% end if %>  American Express 
                                        <% if D="1" then %>
                                            <input name="pcPay_PayPal_CardTypes" type="checkbox" class="clearBorder" id="pcPay_PayPal_CardTypes" value="D" checked> 
                                        <% else %>
                                            <input name="pcPay_PayPal_CardTypes" type="checkbox" class="clearBorder" id="pcPay_PayPal_CardTypes" value="D"> 
                                        <% end if %> 
                                        Maestro
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
                                        &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=805')"></a></td>
								</tr>
								<tr> 
									<td> <div align="right"> 
											<input name="pcPay_PayPal_Sandbox" type="checkbox" class="clearBorder" value="YES" <% if pcPay_PayPal_Sandbox=1 then%>checked<% end if %>>
										</div></td>
									<td><b>Enable Test Mode </b>(Credit cards will not be charged) <a href="JavaScript:win('helpOnline.asp?ref=806')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature" /></a><a href="JavaScript:win('helpOnline.asp?ref=805')"></a>													</td>
								</tr>
								<tr> 
									<td> <div align="right">  
									<input name="pcPay_PayPal_CVC" type="checkbox" class="clearBorder" value="1" <% if pcPay_PayPal_CVC=1 then%>checked<% end if %>>
										</div></td>
									<td><b>Enable Credit Card Security Code</b><a href="JavaScript:win('helpOnline.asp?ref=805')"></a>
							      <div></div>									</td>
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
                        <td width="580" class="pcPanelTitle1"><strong>Step 2: Configure Settings</strong></td>
                        </tr>
                        </table>
                    </div>
                    <div class="CollapsiblePanelContent">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td colspan="2"><p><br />
                              To customize your PayPal Express Checkout transaction experience log into your PayPal account to upload a logo under your Profile</p>
                          </tr>
                          <tr>
                            <td width="18%">&nbsp;</td>
                            <td width="82%" class="pcSubmenuContent">&nbsp;</td>
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
                var CollapsiblePanel2 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel2", {contentIsOpen:true});;
                var CollapsiblePanel3 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel3", {contentIsOpen:false});
                var CollapsiblePanel4 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel4", {contentIsOpen:false});
                //-->
                </script>
            </td>
        </tr>
    </table>
<% end if
end if %>
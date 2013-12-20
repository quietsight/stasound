<%
'--- Start PayFlow Pro ---
Function gwPFPEdit()
	call opendb()
	'request gateway variables and insert them into the verisign_pfp table
	query="SELECT v_User,v_Password FROM verisign_pfp where id=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	v_User2=rs("v_User")
	'decrypt
	v_User2=enDeCrypt(v_User2, scCrypPass)
	v_Password2=rs("v_Password")
	'decrypt
	v_Password2=enDeCrypt(v_Password2, scCrypPass)
	set rs=nothing
	v_User=request.Form("v_User")
	if v_User="" then
		v_User=v_User2
	end if
	'encrypt
	v_User=enDeCrypt(v_User, scCrypPass)
	v_Password=request.Form("v_Password")
	if v_Password="" then
		v_Password=v_Password2
	end if
	'encrypt
	v_Password=enDeCrypt(v_Password, scCrypPass)

	v_Url=request.Form("pfp_Url")
	v_Partner=request.Form("pfp_Partner")
	v_Vendor=request.Form("pfp_Vendor")
	pfl_testmode=request.Form("pfp_testmode")
	pfl_CSC=request.Form("pfp_CSC")
	if pfl_testmode="" then
		pfl_testmode=0
	end if
	pfl_transtype=request.Form("pfp_transtype") 
	
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
	
	query="UPDATE verisign_pfp SET v_Url='"&v_Url&"',v_User='"&v_User&"',v_Partner='"&v_Partner&"',v_Password='"&v_Password &"',v_Vendor='"&v_Vendor&"',pfl_testmode='"&pfl_testmode&"',pfl_transtype='"&pfl_transtype&"',pfl_CSC='"&pfl_CSC&"' where id=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"',pcPayTypes_ppab=0 WHERE gwCode=2"
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

Function gwPFP()
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
		response.redirect "pcConfigurePayment.asp?gwchoice=2&msg="&Server.URLEncode("An error occurred while trying to add Payflow Pro as your payment gateway. <b>""Vendor""</b> is a required field.")
	End If 
			
	'check to see if centinel is activated
	pcPay_Cent_Active=request.Form("pcPay_Cent_Active_pfp")
	
	if pcPay_Cent_Active="YES" then
		pcPay_Cent_Active=1
		pcPay_Cent_TransactionURL=request.Form("pcPay_Cent_TransactionURL_pfp")
		pcPay_Cent_MerchantID=request.Form("pcPay_Cent_MerchantID_pfp")
		pcPay_Cent_ProcessorId=request.Form("pcPay_Cent_ProcessorId_pfp")
		if pcPay_Cent_TransactionURL="" or pcPay_Cent_MerchantID="" OR	pcPay_Cent_ProcessorId="" then
			response.redirect "pcConfigurePayment.asp?gwchoice=2&msg="&Server.URLEncode("An error occurred while trying to add Cardinal Centinel for Authorize.Net. <b>""Tranaction URL, Merchant ID and Process ID""</b> are all required fields.")
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

	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName, pcPayTypes_ppab) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'PayPal-PayFlow-Pro','gwpfp.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",2,'"&paymentNickName&"',0)"
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
				
<% if request("gwchoice")="2" then
	if request("mode")="Edit" then
		call opendb()
		
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
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=2"
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
	<input type="hidden" name="addGw" value="2">
    <div class="pcCPmessageSuccess">
        <% if request("mode")="Edit" then %>
            <p><strong>You're editing </strong><strong>PayPal Payflow Pro</strong>
            <br />
            <br />
            <p><strong><a href="pcPaymentSelection.asp">Change Payment Option</a></strong></p>
            <br /></p>
        <% else %>
            <p><strong>You've selected PayPal Payflow Pro</strong>
            <br /><br />
            </strong>Use your own merchant account and stay in control of your checkout pages with this fully customizable gateway solution. PayPal simplifies PCI compliance for you, if needed.<strong>
            <br /><br />
            <a href="https://merchant.paypal.com/us/cgi-bin/?&amp;cmd=_render-content&amp;content_ID=merchant/payment_gateway&amp;nav=2.1.2&amp;nav=2.0.8" target="_blank">Sign Up and Learn More</a></strong>
            <br /><br />
            To start accepting payments, please complete the process below.
            <strong><a href="pcPaymentSelection.asp">Change Payment Option</a></strong></p>       
       	<% end if %>
    </div>
    <table width="100%">
        <tr>
            <td align="left" style="font-size:15px;"><img src="gateways/logos/payflow_logo.jpg" width="150" height="68"></td>
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
                            <td width="580" class="pcPanelTitle1"><strong>Step 1: Payflow Account Information</strong></td>
                          </tr>
                        </table>
                    </div>
                    <div class="CollapsiblePanelContent">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
							<% if request("mode")="Edit" then %>
                          <% 
                                dim v_UserCnt,v_UserEnd,v_UserStart
                                v_UserCnt=(len(pfp_User)-2)
                                v_UserEnd=right(pfp_User,2)
                                v_UserStart=""
                                for c=1 to v_UserCnt
                                    v_UserStart=v_UserStart&"*"
                                next
                                %>
                                <tr> 
                                    <td>&nbsp;</td>
                                    <td>Current User:&nbsp;<%=v_UserStart&v_UserEnd%></td>
                                </tr>
                                <tr> 
                                    <td>&nbsp;</td>
                                    <td> For security reasons, your &quot;User&quot; is only 
                                        partially shown on this page. The password is not shown. 
                                        If you need to edit your account information, please re-enter 
                                        your &quot;User&quot; and password below.</td>
                                </tr>
							<% else %>
                                <tr>
                                    <td colspan="2">You must have a PayPal Payflow account to use Payflow Pro. If you don't have an account, sign up for one now. Sign up now
                                        <br /><br />
                                        <strong>Enter your PayPal Payflow Information</strong><br />
                                      You created this information when you signed up for PayPal Payflow Pro. Enter it here to connect your account and allow payments. (Note: This is also your login information for <a href="http://manager.paypal.com" target="_blank">PayPal Manager</a>.)<br /><br />
                                    </td>
                                </tr>
							<% end if %>
                            <tr> 
                                <td width="127" valign="top"><div align="right">Partner Name:</div></td>
                              <td><input type="text" value="<%=pfp_Partner%>" name="pfp_Partner" size="24">
                              <a href="JavaScript:win('helpOnline.asp?ref=804')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature" /></a><a href="JavaScript:win('helpOnline.asp?ref=301')"></a></td>
                            </tr>
                            <tr> 
                                <td width="127" valign="top"><div align="right">Merchant Login:</div></td>
                                <td><input type="text" value="<%=pfp_Vendor%>" name="pfp_Vendor" size="24"> <a href="JavaScript:win('helpOnline.asp?ref=804')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature" /></a></td>
                            </tr>
                            <tr> 
                                <td width="127" valign="top"><div align="right">User:</div></td>
                              <td>
<input type="text" value="" name="pfp_User" size="24"> 
                                    <input type="hidden" value="<%=pfp_Url%>" name="pfp_Url" size="24">
                                    <a href="JavaScript:win('helpOnline.asp?ref=804')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature" /></a><a href="JavaScript:win('helpOnline.asp?ref=301')"></a></td>
                            </tr>
								
                            <tr> 
                                <td valign="top"><div align="right">Password:</div></td>
                                <td>
                                    <input type="password" value="" name="pfp_Password" size="24"> <a href="JavaScript:win('helpOnline.asp?ref=804')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature" /></a></td>
                            </tr>
                            <tr> 
                                <td valign="top" nowrap><div align="right">Transaction Type:</div></td>
                                <td><select name="pfp_transtype">
                                <option value="S" selected>Sale</option>
                                    <option value="A"<% if pfp_transtype="A" then
                                    response.write " selected"
                                    end if %>
                                    >Authorize Only</option>
                                    </select></td>
                            </tr>
                            <tr> 
                                <td valign="top" nowrap><div align="right">Require CSC:</div></td>
                                <td><% if pfp_CSC="YES" then %> <input type="radio" class="clearBorder" name="pfp_CSC" value="YES" checked>
                                    Yes 
                                    <input name="pfp_CSC" type="radio" class="clearBorder" value="NO">
                                    No 
                                    <% else %> <input type="radio" class="clearBorder" name="pfp_CSC" value="YES">
                                    Yes 
                                    <input name="pfp_CSC" type="radio" class="clearBorder" value="NO" checked>
                                    No 
                              <% end if %> </td>
                            </tr>
                            <tr> 
                                <td>&nbsp;</td>
                              <td> <% if pfp_testmode="YES" then %> <input type="checkbox" class="clearBorder" name="pfp_testmode" value="YES" checked> 
                                <% else %> <input type="checkbox" class="clearBorder" name="pfp_testmode" value="YES"> 
                                <% end if %> 
                                <b>Enable Test Mode </b>(Credit 
                                cards will not be charged) <a href="JavaScript:win('helpOnline.asp?ref=804')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature" /></a></td>
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
                            <td class="pcPanelTitle1">Step 2: You can change the display name that is shown for this payment type. </td>
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
                <div id="CollapsiblePanel3" class="CollapsiblePanel">
                    <div class="CollapsiblePanelTab1">
                        <table width="100%">
                          <tr>
                            <td class="pcPanelTitle1">Step 3: Order Processing: Order Status and Payment Status</td>
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
                <div  id="CollapsiblePanel4" class="CollapsiblePanel">
                    <div class="CollapsiblePanelTab1">
                        <table width="100%">
                          <tr>
                            <td class="pcPanelTitle1">Step 4: Minimize fraud by enabling Centinel by CardinalCommerce</td>
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
                //-->
                </script>
          	</td>
       	</tr>
   	</table>
<% end if %>
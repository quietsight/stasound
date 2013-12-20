<%
'--- Start Payflow Link ---
Function gwPFLEdit()
	call opendb()
	'request gateway variables and insert them into the verisign_pfp table
	query="SELECT v_User,v_Password,v_Vendor FROM verisign_pfp where id=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	pfl_User2=rs("v_User")
	'decrypt
	pfl_User2=enDeCrypt(pfl_User2, scCrypPass)
	
	pfl_Password2=rs("v_Password")
	'decrypt
	pfl_Password2=enDeCrypt(pfl_Password2, scCrypPass)
	
	pfl_MerchantLogin2=rs("v_Vendor")
	'decrypt
	pfl_MerchantLogin2=enDeCrypt(pfl_MerchantLogin2, scCrypPass)
	
	set rs=nothing
	pfl_User=request.Form("pfl_User")
	if pfl_User="" then
		pfl_User=pfl_User2
	end if
	'encrypt
	pfl_User=enDeCrypt(pfl_User, scCrypPass)
	pfl_Password=request.Form("pfl_Password")
	if pfl_Password="" then
		pfl_Password=pfl_Password2
	end if
	'encrypt
	pfl_Password=enDeCrypt(pfl_Password, scCrypPass)

	pfl_MerchantLogin=request.Form("pfl_MerchantLogin")
	if pfl_MerchantLogin="" then
		pfl_MerchantLogin=pfl_MerchantLogin2
	end if
	'encrypt
	pfl_MerchantLogin=enDeCrypt(pfl_MerchantLogin, scCrypPass)
	
	pfl_Partner=request.Form("pfl_Partner")
	pfl_testmode=request.Form("pfl_testmode")
	pfl_CSC=request.Form("pfl_CSC")
	if pfl_testmode="" then
		pfl_testmode=0
	end if
	pfl_transtype=request.Form("pfl_transtype") 
	
	query="UPDATE verisign_pfp SET v_Url='na',v_User='"&pfl_User&"',v_Partner='"&pfl_Partner&"',v_Password='"&pfl_Password &"',v_Vendor='"&pfl_MerchantLogin&"',pfl_testmode='"&pfl_testmode&"',pfl_transtype='"&pfl_transtype&"',pfl_CSC='"&pfl_CSC&"' where id=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"',pcPayTypes_ppab=0 WHERE gwCode=99"
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

Function gwPFL()
	varCheck=1
	'request gateway variables and insert them into the verisign_pfp table
	pfl_User=request.Form("pfl_User")
	pfl_User=enDeCrypt(pfl_User, scCrypPass)

	pfl_Password=request.Form("pfl_Password")
	pfl_Password=enDeCrypt(pfl_Password, scCrypPass)

	pfl_MerchantLogin=request.Form("pfl_MerchantLogin")
	pfl_MerchantLogin=enDeCrypt(pfl_MerchantLogin, scCrypPass)

	pfl_testmode=request.Form("pfl_testmode")
	pfl_Partner=request.Form("pfl_Partner")
	if pfl_testmode="" then
		pfl_testmode=0
	end if
	pfl_transtype=request.Form("pfl_transtype")
	pfl_CSC=request.Form("pfl_CSC")
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
			
	err.clear
	err.number=0
			
	call openDb() 

	query="UPDATE verisign_pfp SET v_Url='na',v_Type='na',v_User='"&pfl_User&"',v_Partner='"&pfl_Partner&"' ,v_Password='"&pfl_Password &"' ,v_Vendor='"&pfl_MerchantLogin&"',v_Tender='na',pfl_testmode='"&pfl_testmode&"',pfl_transtype='"&pfl_transtype&"',pfl_CSC='"&pfl_CSC&"' WHERE id=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName, pcPayTypes_ppab) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'PayPal-Payflow-Link','gwpflEB.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",99,'"&paymentNickName&"',0)"
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
				
<% if request("gwchoice")="99" then
	if request("mode")="Edit" then
		call opendb()
		
		query= "SELECT v_User,v_Partner,v_Password,v_Vendor,pfl_testmode,pfl_transtype,pfl_CSC FROM verisign_pfp where id=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pfl_User=rs("v_User")
			pfl_User=enDeCrypt(pfl_User, scCrypPass)

		pfl_Partner=rs("v_Partner")
		pfl_Password=rs("v_Password")
			pfl_Password=enDeCrypt(pfl_Password, scCrypPass)

		pfl_MerchantLogin=rs("v_Vendor")
			pfl_MerchantLogin=enDeCrypt(pfl_MerchantLogin, scCrypPass)

		pfl_testmode=rs("pfl_testmode")
		pfl_transtype=rs("pfl_transtype")
		pfl_CSC=rs("pfl_CSC") 
		set rs=nothing
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=99"
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
		
		dim pfl_UserCnt,pfl_UserEnd,pfl_UserStart
		pfl_UserCnt=(len(pfl_User)-2)
		pfl_UserEnd=right(pfl_User,2)
		pfl_UserStart=""
		for c=1 to pfl_UserCnt
			pfl_UserStart=pfl_UserStart&"*"
		next
		
		dim pfl_MLoginCnt,pfl_MLoginEnd,pfl_MLoginStart
		pfl_MLoginCnt=(len(pfl_MerchantLogin)-2)
		pfl_MLoginEnd=right(pfl_MerchantLogin,2)
		pfl_MLoginStart=""
		for c=1 to pfl_MLoginCnt
			pfl_MLoginStart=pfl_MLoginStart&"*"
		next
		%>
    
		<input type="hidden" name="mode" value="Edit">
    <% end if %>
	<input type="hidden" name="addGw" value="99">
    <div class="pcCPmessageSuccess">
  <% if request("mode")="Edit" then %>
            <p>
                <strong>You're editing </strong><strong>PayPal Payflow Link</strong>
                - Embedded Payment Integration<br />
                <br />
                <p><strong><a href="pcPaymentSelection.asp">Change Payment Option</a></strong></p>
                <br />
            </p>
        <% else %>
            <p><strong>You've selected PayPal Payflow Link</strong> - Embedded Payment Integration<br />
            <br />
            </strong>Connect your merchant account with a
            PCI-compliant gateway. Setup is quick and
            customers pay without leaving your site.<br />
            <br />
            <strong> <a href="https://merchant.paypal.com/us/cgi-bin/?&amp;cmd=_render-content&amp;content_ID=merchant/payment_gateway&amp;nav=2.1.2&amp;nav=2.0.8" target="_blank">Sign Up and Learn More</a></strong>
            <br />
            <br />
            To start accepting payments, please complete the process below.
            <strong><a href="pcPaymentSelection.asp">Change Payment Option</a></strong></p>
    	<% end if %>
    </div>
    <br />
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
                            	<tr> 
                                	<td>Current User</td>
                                	<td width="83%">:&nbsp;<%=pfl_UserStart&pfl_UserEnd%></td>
                            	</tr>
                            	<tr> 
                                	<td>Current Merchant</td>
                                	<td width="83%">:&nbsp;<%=pfl_MLoginStart&pfl_MLoginEnd%></td>
                            	</tr>
	                            <tr> 
	                                <td colspan="2"><br />
For security reasons, your &quot;Login&quot; is only 
	                                    partially shown on this page. If you need to edit your 
	                                    account information, please re-enter your &quot;Login&quot; 
	                                    below.</td>
	                            </tr>
							<% else %>
                                <tr><td colspan="2">You must have a PayPal Payflow account to use Payflow Link. If you don't have an account, sign up for one now. Sign up now
                                <br />
                                <br />
                                Enter your PayPal Payflow Information You created this information when you signed up for PayPal Payflow Link. Enter it here to connect your account and allow payments. (Note: This is also your login information for PayPal Manager.)<br /></td></tr>
							<% end if %>
                            <% if pfl_Partner&""="" then
								pfl_Partner="PayPal"
							end if %>
	                            <tr> 
	                                <td width="17%">Partner Name:</td>
	                                <td><input type="text" value="<%=pfl_Partner%>" name="pfl_Partner" size="24">&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=801')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
	                            </tr>
	                            <tr> 
	                                <td width="17%">Merchant Login:</td>
	                                <td><input type="text" value="" name="pfl_MerchantLogin" size="24">&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=801')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
	                            </tr>
	                            <tr>
	                              <td>User:</td>
	                              <td><input type="text" value="" name="pfl_User" size="24" />&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=801')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
	                            </tr>
	                            <tr>
	                              <td>Password:</td>
	                              <td><input type="text" value="" name="pfl_Password" size="24" />&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=801')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
	                            </tr>
	                            <tr> 
                                <td width="17%">Transaction Type:</td>
                                <td><select name="pfl_transtype">
                                <option value="S" selected>Sale</option>
                                        <option value="A"<% if pfl_transtype="A" then
                                        response.write " selected"
                                        end if %>
                                        >Authorize Only</option>
                                    </select></td>
                            </tr>
							<tr> 
                                <td>Enable Test Mode</td>
                                <td><% if pfl_testmode="YES" then %> <input type="checkbox" class="clearBorder" name="pfl_testmode" value="YES" checked> 
                                <% else %> <input type="checkbox" class="clearBorder" name="pfl_testmode" value="YES"> 
                                <% end if %>
                                 <a href="JavaScript:win('helpOnline.asp?ref=801')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature" /></a></td>
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
                        <td width="580" class="pcPanelTitle1"><strong>Step 2: Configure PayPal Payments Advanced</strong></td>
                        </tr>
                        </table>
                    </div>
                    <div class="CollapsiblePanelContent">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td colspan="2"><p>You must adjust these settings to accept payments. To make configuration easier, ProductCart will adjust any additional settings not listed here. 
                              </p>
                            <p>
                                <ol>
                                <li>Log in to <a href="http://manager.paypal.com" target="_blank"><strong>PayPal Manager</strong></a>.</li>
                                <li>Select <strong>Service Settings</strong>.</li>
                                <li>Select <strong>Hosted Checkout Pages</strong>, select <strong>Set up</strong>.</li>
                                <li>Under <strong>Security Options</strong>, please set <strong>Enable Secure Token</strong> to &quot;Yes&quot;.</li>
                            </ol>
                            </p>
                            <p><a href="http://www.youtube.com/watch?v=9lnjtS-iX1A&amp;list=PL3E37BD5019E92D63&amp;index=9&amp;feature=plpp_video" target="_blank">View a tutorial</a> | <a href="https://merchant.paypal.com/us/cgi-bin/?cmd=_render-content&amp;content_ID=merchant/wf_question" target="_blank">PayPal Help</a></p></td>
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
                            <td width="580" class="pcPanelTitle1"><strong>Step 3: You can change the display name that is shown for this payment type. </strong></td>
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
                            <td width="580" class="pcPanelTitle1"><strong>Step 4: Order Processing: Order Status and Payment Status</strong></td>
                          </tr>
                        </table>
                    </div>
                    <div class="CollapsiblePanelContent">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
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
<% end if %>
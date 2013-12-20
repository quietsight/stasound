<%
'--- Start PayPal Payments Advanced ---
Function gwPPAEdit()
	call opendb()
	'request gateway variables and insert them into the pcPay_PayPalAdvanced table
	query="SELECT pcPay_PayPalAd_User,pcPay_PayPalAd_Password, pcPay_PayPalAd_MerchantLogin FROM pcPay_PayPalAdvanced where pcPay_PayPalAd_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	pcPay_PayPalAd_User2=rs("pcPay_PayPalAd_User")
	'decrypt
	pcPay_PayPalAd_User2=enDeCrypt(pcPay_PayPalAd_User2, scCrypPass)
	
	pcPay_PayPalAd_Password2=rs("pcPay_PayPalAd_Password")
	'decrypt
	pcPay_PayPalAd_Password2=enDeCrypt(pcPay_PayPalAd_Password2, scCrypPass)
	
	pcPay_PayPalAd_MerchantLogin2=rs("pcPay_PayPalAd_MerchantLogin")
	'decrypt
	pcPay_PayPalAd_MerchantLogin2=enDeCrypt(pcPay_PayPalAd_MerchantLogin2, scCrypPass)

	set rs=nothing
	
	pcPay_PayPalAd_User=request.Form("pcPay_PayPalAd_User")
	if pcPay_PayPalAd_User="" then
		pcPay_PayPalAd_User=pcPay_PayPalAd_User2
	end if
	'encrypt
	pcPay_PayPalAd_User=enDeCrypt(pcPay_PayPalAd_User, scCrypPass)
	pcPay_PayPalAd_Password=request.Form("pcPay_PayPalAd_Password")
	if pcPay_PayPalAd_Password="" then
		pcPay_PayPalAd_Password=pcPay_PayPalAd_Password2
	end if
	'encrypt
	pcPay_PayPalAd_Password=enDeCrypt(pcPay_PayPalAd_Password, scCrypPass)

	pcPay_PayPalAd_Partner=request.Form("pcPay_PayPalAd_Partner")

	pcPay_PayPalAd_MerchantLogin=request.Form("pcPay_PayPalAd_MerchantLogin")
	if pcPay_PayPalAd_MerchantLogin="" then
		pcPay_PayPalAd_MerchantLogin=pcPay_PayPalAd_MerchantLogin2
	end if
	'encrypt
	pcPay_PayPalAd_MerchantLogin=enDeCrypt(pcPay_PayPalAd_MerchantLogin, scCrypPass)
	
	pcPay_PayPalAd_Sandbox=request.Form("pcPay_PayPalAd_Sandbox")
	pcPay_PayPalAd_CSC=request.Form("pcPay_PayPalAd_CSC")
	if pcPay_PayPalAd_Sandbox="" then
		pcPay_PayPalAd_Sandbox=0
	end if
	pcPay_PayPalAd_TransType=request.Form("pcPay_PayPalAd_TransType") 
	
	query="UPDATE pcPay_PayPalAdvanced SET pcPay_PayPalAd_Partner='"&pcPay_PayPalAd_Partner&"',pcPay_PayPalAd_MerchantLogin='"&pcPay_PayPalAd_MerchantLogin&"',pcPay_PayPalAd_User='"&pcPay_PayPalAd_User&"',pcPay_PayPalAd_Password='"&pcPay_PayPalAd_Password &"',pcPay_PayPalAd_TransType='"&pcPay_PayPalAd_TransType&"',pcPay_PayPalAd_CSC='"&pcPay_PayPalAd_CSC&"',pcPay_PayPalAd_Sandbox='"&pcPay_PayPalAd_Sandbox&"' where pcPay_PayPalAd_ID=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='PayPal',pcPayTypes_ppab=0 WHERE gwCode=80"
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

Function gwPPA()
	varCheck=1
	'request gateway variables and insert them into the pcPay_PayPalAdvanced table
	pcPay_PayPalAd_User=request.Form("pcPay_PayPalAd_User")
	pcPay_PayPalAd_User=enDeCrypt(pcPay_PayPalAd_User, scCrypPass)
	pcPay_PayPalAd_Partner=request.Form("pcPay_PayPalAd_Partner")
	pcPay_PayPalAd_Password=request.Form("pcPay_PayPalAd_Password")
	pcPay_PayPalAd_Password=enDeCrypt(pcPay_PayPalAd_Password, scCrypPass)
	pcPay_PayPalAd_Sandbox=request.Form("pcPay_PayPalAd_Sandbox")
	if pcPay_PayPalAd_Sandbox="" then
		pcPay_PayPalAd_Sandbox=0
	end if
	pcPay_PayPalAd_TransType=request.Form("pcPay_PayPalAd_TransType")
	pcPay_PayPalAd_CSC=request.Form("pcPay_PayPalAd_CSC")
	priceToAddType=request.Form("priceToAddType")
	if priceToAddType="price" then
		priceToAdd=replacecomma(Request("priceToAdd"))
		percentageToAdd="0"
		If priceToAdd="" Then
			priceToAdd="0"
		end if
	else
		priceToAdd="0"
		if priceToAddType="null" then
			percentageToAdd="0"
		else
		percentageToAdd=request.Form("percentageToAdd")
		If percentageToAdd="" Then
			percentageToAdd="0"
		end if
	end if
	end if
	paymentNickName="PayPal"
	pcPay_PayPalAd_MerchantLogin=request.Form("pcPay_PayPalAd_MerchantLogin")
	pcPay_PayPalAd_MerchantLogin=enDeCrypt(pcPay_PayPalAd_MerchantLogin, scCrypPass)
	If pcPay_PayPalAd_MerchantLogin="" then
		session("adminpcPay_PayPalAd_TransType")=pcPay_PayPalAd_TransType
		session("adminpcPay_PayPalAd_User")=pcPay_PayPalAd_User
		session("adminpcPay_PayPalAd_Partner")=pcPay_PayPalAd_Partner
		session("adminpcPay_PayPalAd_Password")=pcPay_PayPalAd_Password
		session("adminpcPay_PayPalAd_Sandbox")=pcPay_PayPalAd_Sandbox
		session("adminpcPay_PayPalAd_CSC")=pcPay_PayPalAd_CSC
		response.redirect "pcConfigurePayment.asp?gwchoice=80&msg="&Server.URLEncode("An error occurred while trying to add Payflow Pro as your payment gateway. <b>""Merchant Login""</b> is a required field.")
	End If 
			
	err.clear
	err.number=0
			
	call openDb() 

	query="UPDATE pcPay_PayPalAdvanced SET pcPay_PayPalAd_Partner='"&pcPay_PayPalAd_Partner&"' ,pcPay_PayPalAd_MerchantLogin='"&pcPay_PayPalAd_MerchantLogin&"',pcPay_PayPalAd_User='"&pcPay_PayPalAd_User&"',pcPay_PayPalAd_Password='"&pcPay_PayPalAd_Password &"' ,pcPay_PayPalAd_TransType='"&pcPay_PayPalAd_TransType&"',pcPay_PayPalAd_CSC='"&pcPay_PayPalAd_CSC&"',pcPay_PayPalAd_Sandbox='"&pcPay_PayPalAd_Sandbox&"' WHERE pcPay_PayPalAd_ID=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName, pcPayTypes_ppab) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'PayPal-Payments-Advanced','gwPPA.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",80,'"&paymentNickName&"',0)"
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
				
<% if request("gwchoice")="80" then
	pcConflictIdPayment = 0

	if request("mode")="Edit" then
		call opendb()

		query= "SELECT pcPay_PayPalAd_ID, pcPay_PayPalAd_Partner, pcPay_PayPalAd_MerchantLogin, pcPay_PayPalAd_User, pcPay_PayPalAd_Password, pcPay_PayPalAd_TransType, pcPay_PayPalAd_CSC, pcPay_PayPalAd_Sandbox FROM pcPay_PayPalAdvanced where pcPay_PayPalAd_ID=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcPay_PayPalAd_Partner=rs("pcPay_PayPalAd_Partner")
		pcPay_PayPalAd_MerchantLogin=rs("pcPay_PayPalAd_MerchantLogin")
		pcPay_PayPalAd_User=rs("pcPay_PayPalAd_User")
		pcPay_PayPalAd_User=enDeCrypt(pcPay_PayPalAd_User, scCrypPass)
		pcPay_PayPalAd_Password=rs("pcPay_PayPalAd_Password")
		pcPay_PayPalAd_TransType=rs("pcPay_PayPalAd_TransType")
		pcPay_PayPalAd_CSC=rs("pcPay_PayPalAd_CSC") 
		pcPay_PayPalAd_Sandbox=rs("pcPay_PayPalAd_Sandbox")
		set rs=nothing
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=80"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=conntemp.execute(query)
		if rs.eof then
			paymentNickName="PayPal"
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
		
		dim ppa_UserCnt,ppa_UserEnd,ppa_UserStart
		ppa_UserCnt=(len(pcPay_PayPalAd_User)-2)
		ppa_UserEnd=right(pcPay_PayPalAd_User,2)
		ppa_UserStart=""
		for c=1 to ppa_UserCnt
			ppa_UserStart=ppa_UserStart&"*"
		next %>
		<input type="hidden" name="mode" value="Edit">
    <% else
		'//Check if any other PayPal Services are activated.
		call opendb()
		query= "SELECT idPayment, gwCode FROM payTypes WHERE gwCode=3 OR gwCode=53 OR gwCode=46"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=conntemp.execute(query)
		if NOT rs.eof then
			pcConflictIdPayment = rs("idPayment")
			pcConflictID = rs("gwCode")
			select case pcConflictID
				case "3"
					pcConflictDesc = "PayPal Payments Standard"
				case "53"
					pcConflictDesc = "Website Payments Pro (UK)"
				case "46"
					pcConflictDesc = "PayPal Payments Pro"
			end select %>
        	<div class="pcCPmessage">
        	  <p>You currently have <strong><%=pcConflictDesc%></strong> active for this store. In order to use <strong>PayPal Payments Advanced</strong> you will need to first disable <strong><%=pcConflictDesc%></strong>.<br />
        	    <br />
        	  </p>
        	  <p><a href="pcConfigurePayment.asp?mode=Del&id=<%=pcConflictIdPayment%>&gwchoice=<%=pcConflictID%>&activate=80">Disable <%=pcConflictDesc%> and enable PayPal Payments Advanced</a></p>
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
	<input type="hidden" name="addGw" value="80">
    <div class="pcCPmessageSuccess">
        <% if request("mode")="Edit" then %>
            <p><strong>You're editing </strong><strong>PayPal Payments Advanced</strong>
            <br />
            <br />
            <p><strong><a href="pcPaymentSelection.asp">Change Payment Option</a></strong></p>
            <br /></p>
        <% else %>
            <p><strong>You've selected PayPal Payments Advanced</strong><br>
            <br>
            The easy way to create a professional checkout experience that lets buyers pay without leaving your site. And PayPal processes credit cards behind the scenes, helping you simplify PCI compliance.<strong><br>
            <br>
            <a href="https://www.paypal.com/webapps/mpp/paypal-payments-advanced" target="_blank">Sign Up and Learn More</a></strong>
      </p>
            <br />
      <p>To start accepting payments, please complete the process below.        <strong><a href="pcPaymentSelection.asp">Change Payment Option</a></strong></p>
            <br /></p>
    	<% end if %>
</div>
	<table width="100%">
        <tr>
            <td align="left" style="font-size:15px;"><img src="gateways/logos/paypal_logo1.gif" width="253" height="80"></td>
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
                            <td width="580" class="pcPanelTitle1"><strong>Step 1: Configure Account - PayPal Payments Advanced</strong></td>
                          </tr>
                        </table>
                    </div>
                    <div class="CollapsiblePanelContent">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
							<% if request("mode")="Edit" then %>
                            	<tr> 
                                	<td>Current User</td>
                                	<td width="83%">:&nbsp;<%=ppa_UserStart&ppa_UserEnd%></td>
                            	</tr>
	                            <tr> 
	                                <td colspan="2"><br />
For security reasons, your &quot;Login&quot; is only 
	                                    partially shown on this page. If you need to edit your 
	                                    account information, please re-enter your &quot;Login&quot; 
	                                    below.</td>
	                            </tr>
							<% else %>
	                            <tr>
	                              <td colspan="2" valign="top">Enter your PayPal Account information
We need this information to work with PayPal
so that payments can be sent to your account.<br /></td>
	                            </tr>
							<% end if %>
                            <% if pcPay_PayPalAd_Partner&""="" then
								pcPay_PayPalAd_Partner = "PayPal"
							end if %>
	                            <tr> 
	                                <td width="17%" align="right">Partner Name:</td>
	                                <td><input type="text" value="<%=pcPay_PayPalAd_Partner%>" name="pcPay_PayPalAd_Partner" size="24">&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=800')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
	                            </tr>
	                            <tr> 
	                                <td width="17%" align="right">Merchant Login:</td>
	                                <td><input type="text" value="" name="pcPay_PayPalAd_MerchantLogin" size="24">&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=800')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
	                            </tr>
	                            <tr>
	                              <td align="right">User:</td>
	                              <td><input type="text" value="" name="pcPay_PayPalAd_User" size="24" autocomplete="off" />&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=800')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
	                            </tr>
	                            <tr>
	                              <td align="right">Password:</td>
	                              <td><input type="text" value="" name="pcPay_PayPalAd_Password" size="24" />&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=800')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
	                            </tr>
	                            <tr> 
                                <td width="17%" align="right">Transaction Type:</td>
                                <td><select name="pcPay_PayPalAd_TransType">
                                <option value="S" selected>Sale</option>
                                        <option value="A"<% if pcPay_PayPalAd_TransType="A" then
                                        response.write " selected"
                                        end if %>
                                        >Authorize Only</option>
                                    </select></td>
                            </tr>
                            <tr> 
                                <td align="right">Require CSC:</td>
                                <td><% if pcPay_PayPalAd_CSC="YES" then %> <input type="radio" class="clearBorder" name="pcPay_PayPalAd_CSC" value="YES" checked>
                                    Yes 
                                    <input name="pcPay_PayPalAd_CSC" type="radio" class="clearBorder" value="NO">
                                    No 
                                    <% else %> <input type="radio" class="clearBorder" name="pcPay_PayPalAd_CSC" value="YES">
                                    Yes 
                                    <input name="pcPay_PayPalAd_CSC" type="radio" class="clearBorder" value="NO" checked>
                                    No 
                              <% end if %> </td>
                            </tr>
							<tr> 
                                <td align="right">Enable Test Mode</td>
                                <td><% if pcPay_PayPalAd_Sandbox="YES" then %> <input type="checkbox" class="clearBorder" name="pcPay_PayPalAd_Sandbox" value="YES" checked> 
                                <% else %> <input type="checkbox" class="clearBorder" name="pcPay_PayPalAd_Sandbox" value="YES"> 
                              <% end if %> <a href="JavaScript:win('helpOnline.asp?ref=800')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature" /></a></td>
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
                            <td width="580" class="pcPanelTitle1"><strong>Step 3: Order Processing: Order Status and Payment Status</strong></td>
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
                //-->
                </script>
            </td>
        </tr>
    </table>
<% end if
end if %>
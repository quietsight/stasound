<%
'---Start Echo---
Function gwechoEdit()
	call opendb()
	'request gateway variables and insert them into the Echo table
	query= "SELECT merchant_echo_id, merchant_pin FROM echo where id=1"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	transaction_type=request.Form("transaction_type")
	merchant_echo_id2=rstemp("merchant_echo_id")
	'decrypt
	merchant_echo_id2=enDeCrypt(merchant_echo_id2, scCrypPass)
	merchant_pin2=rstemp("merchant_pin")
	'decrypt
	merchant_pin2=enDeCrypt(merchant_pin2, scCrypPass)
	merchant_email=request.Form("merchant_email")
	merchant_echo_id=request.Form("merchant_echo_id")
	if merchant_echo_id="" then
		merchant_echo_id=merchant_echo_id2
	end if
	'encrypt
	merchant_echo_id=enDeCrypt(merchant_echo_id, scCrypPass)
	merchant_pin=request.Form("merchant_pin")
	if merchant_pin="" then
		merchant_pin=merchant_pin2
	end if
	'encrypt
	merchant_pin=enDeCrypt(merchant_pin, scCrypPass)
	cnp_security=request.Form("cnp_security")
	
	query="UPDATE echo SET transaction_type='"&transaction_type&"',merchant_echo_id='"&merchant_echo_id&"',merchant_pin='"&merchant_pin&"',merchant_email='"&merchant_email&"',cnp_security="&cnp_security&" WHERE id=1"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=19"
	
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If		
	call closedb()
end function

Function gwecho()
	varCheck=1
	'request gateway variables and insert them into the echo table
	order_type=request.Form("order_type")
	billing_ip_address=request.Form("billing_ip_address")
	merchant_email=request.Form("merchant_email")
	merchant_echo_id=request.Form("merchant_echo_id")
	cnp_security=request.Form("cnp_security")
	If merchant_echo_id="" then
		call closedb()
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Echo as your payment gateway. <b>""Merchant ID""</b> is a required field.")
	End If
	'encrypt
	merchant_echo_id=enDeCrypt(merchant_echo_id, scCrypPass)
				
	merchant_pin=request.Form("merchant_pin")
	If merchant_pin="" then
		call closedb()
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Echo as your payment gateway. <b>""Merchant's PIN""</b> is a required field.")
	End If
	'encrypt
	merchant_pin=enDeCrypt(merchant_pin, scCrypPass)
	transaction_type=request.Form("transaction_type")
	priceToAddType=request.Form("priceToAddType")
	if priceToAddType="price" then
		priceToAdd=replacecomma(Request("priceToAdd"))
		percentageToAdd="0"
	else
		priceToAdd="0"
		percentageToAdd=request.Form("percentageToAdd")
	end if
	If priceToAdd="" Then
		priceToAdd="0"
	end if
	If percentageToAdd="" Then
		percentageToAdd="0"
	end if
	paymentNickName=replace(request.Form("paymentNickName"),"'","''")
	if paymentNickName="" then
		paymentNickName="Credit Card"
	End If
	
	err.clear
	err.number=0
	call openDb() 

	query="UPDATE echo SET transaction_type='"&transaction_type&"', merchant_echo_id='"&merchant_echo_id&"',merchant_pin='"&merchant_pin&"',merchant_email='"&merchant_email&"',cnp_security="&cnp_security&" WHERE id=1"
	set rs=Server.CreateObject("ADODB.Recordset")  
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'Echo','gwecho.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",19,'"&paymentNickName&"')"
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

if request("gwchoice")="19" then
	if request("mode")="Edit" then
		call opendb()
		query= "SELECT transaction_type,merchant_echo_id,merchant_pin,merchant_email,cnp_security FROM echo where id=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		transaction_type=rs("transaction_type")
		merchant_echo_id=rs("merchant_echo_id")
			'decrypt
			merchant_echo_id=enDeCrypt(merchant_echo_id, scCrypPass)
		merchant_pin=rs("merchant_pin")
			'decrypt
			merchant_pin=enDeCrypt(merchant_pin, scCrypPass)
		merchant_email=rs("merchant_email")
		cnp_security=rs("cnp_security")
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=19"
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
    <% end if %>
	<input type="hidden" name="addGw" value="19">
    <table width="100%">
        <tr>
            <td align="left" style="font-size:15px;"><b>Echo Gateway</b> ( <a href="http://www.echo-inc.com/" target="_blank">Web 
                    site</a> )</td>
            <td align="left" style="font-size:15px;">&nbsp;</td>
        </tr>
    </table>
    <br>
    <table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
        <tr>
            <td><strong>Edit your Echo Gateway Account.<br />
              <br />
              NOTE:
            Echo Gateway is no longer supported by ProductCart. If you disable your account you will not be able to reactivate it.</strong><strong><br />
            <br>
            </strong></td>
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
                            <% dim ECHO_MerchantIDCnt,ECHO_MerchantIDEnd,ECHO_MerchantIDStart
                            ECHO_MerchantIDCnt=(len(merchant_echo_ID)-2)
                            ECHO_MerchantIDEnd=right(merchant_echo_ID,2)
                            ECHO_MerchantIDStart=""
                            for c=1 to ECHO_MerchantIDCnt
                            ECHO_MerchantIDStart=ECHO_MerchantIDStart&"*"
                            next %>
                            <tr>
                                <td height="31">Current Merchant ID:&nbsp;<%=ECHO_MerchantIDStart&ECHO_MerchantIDEnd%></td>
                            </tr>
                            <tr> 
                                <% dim ECHO_MerchantPINCnt,ECHO_MerchantPINEnd,ECHO_MerchantPINStart
                                ECHO_MerchantPINCnt=(len(merchant_pin)-2)
                                ECHO_MerchantPINEnd=right(merchant_pin,2)
                                ECHO_MerchantPINStart=""
                                for c=1 to ECHO_MerchantPINCnt
                                ECHO_MerchantPINStart=ECHO_MerchantPINStart&"*"
                                next %>
                                <td height="31" colspan="2">Current Merchant PIN:&nbsp;<%=ECHO_MerchantPINStart&ECHO_MerchantPINEnd%></td>
                            </tr>
                            <tr> 
                                <td colspan="2"> For security reasons, your &quot;Merchant ID&quot; 
                                and &quot;Merchant PIN&quot; are only partially 
                                shown on this page. If you need to edit your account 
                                information, please re-enter your &quot;Merchant 
                                ID&quot; and &quot;Merchant PIN&quot; below.</td>
                            </tr>
                        <% end if %>
                        <tr> 
                            <td width="137"> <div align="right">Merchant ID:</div></td>
                            <td width="476"> <input type="text" name="merchant_echo_id" size="20"></td>
                        </tr>
                        <tr> 
                            <td> <div align="right">Merchant's PIN :</div></td>
                            <td> <input name="merchant_pin" type="password" size="20"></td>
                        </tr>
                        <tr> 
                            <td> <div align="right">Transaction Type:</div></td>
                            <td>
                                <select name="transaction_type">
                                    <option value="AS" selected>Authorize Only</option>
                                    <option value="AV" <% if transaction_type="AV" then%>selected<%end if%>>Authorize 
                                    w/ AVS</option>
                                    <option value="ES" <% if transaction_type="ES" then%>selected<%end if%>>Authorization 
                                    &amp; Deposit</option>
                                    <option value="EV" <% if transaction_type="EV" then%>selected<%end if%>>Authorization 
                                    w/ AVS &amp; Deposit</option>
                                </select> </td>
                        </tr>
                        <tr> 
                            <td> <div align="right">Merchant's Email Address :</div></td>
                            <td> <input name="merchant_email" type="text" size="40" value="<%=merchant_email%>"></td>
                        </tr>
                        <tr> 
                            <td><div align="right">Require CVV:</div></td>
                            <td><input type="radio" class="clearBorder" name="cnp_security" value="1" checked>
                                Yes 
                                <input type="radio" class="clearBorder" name="cnp_security" value="0" <% if cnp_security="0" then %>checked<% end if%>>
                                No</td>
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
                        <td width="580" class="pcPanelTitle1">Step 2: You have the option to charge a processing fee for this payment option.</td>
                        </tr>
                        </table>
                    </div>
                    <div class="CollapsiblePanelContent">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="18%" nowrap><span class="pcSubmenuHeader">Processing fee:</span><br /></td>
                            <td width="82%" class="pcSubmenuContent">
                              <input type="radio" class="clearBorder" name="priceToAddType" value="price" <% if priceToAddType="price" then%>checked<%end if%>>Flat fee&nbsp;&nbsp; &nbsp;<%=scCurSign%> <input name="priceToAdd" size="6" value="<%=money(priceToAdd)%>"></td>
                            </tr>
                          <tr>
                            <td>&nbsp;</td>
                            <td class="pcSubmenuContent"><input type="radio" class="clearBorder" name="priceToAddType" value="percentage" <% if priceToAddType="percentage" then%>checked<%end if%>>
                                Percentage of Order Total&nbsp;&nbsp; &nbsp; % 
                                <input name="percentageToAdd" size="6" value="<%=percentageToAdd%>"></td>
                          </tr>
                          <tr>
                            <td>&nbsp;</td>
                            <td class="pcSubmenuContent">&nbsp;</td>
                          </tr>
                        </table>
                    </div>
                </div>
                <div id="CollapsiblePanel3" class="CollapsiblePanel">
                    <div class="CollapsiblePanelTab1">
                        <table width="100%">
                          <tr>
                            <td width="580" class="pcPanelTitle1">Step 3: You can change the display name that is shown for this payment type. </td>
                          </tr>
                        </table>
                    </div>
                    <div class="CollapsiblePanelContent">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                                <td width="18%"><div align="left"><strong>Payment Name:&nbsp;</strong></div></td>
                                <td width="82%"><input name="paymentNickName" value="<%=paymentNickName%>" size="35" maxlength="255"></td>
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
                            <td width="580" class="pcPanelTitle1">Step 4: Order Processing: Order Status and Payment Status</td>
                          </tr>
                        </table>
                    </div>
                    <div class="CollapsiblePanelContent">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                                <td>&nbsp;</td>
                                <td>Process orders when they are placed: <input type="checkbox" class="clearBorder" name="pcv_processOrder" value="1" <%if pcv_processOrder="1" then%>checked<%end if%>>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=301')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
                            </tr>
                            <tr> 
                                <td>&nbsp;</td>
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
                              <td>&nbsp;</td>
                            </tr>
                            <tr>
                              <td colspan="2"><input type="submit" value="Add Selected Options" name="Submit" class="submit2"> 
                            &nbsp;
                            <input type="button" value="Back" onclick="javascript:history.back()"></td>
                          </tr>
                        </table>
                    </div>
                </div>
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
<!-- New View End --><% end if %>

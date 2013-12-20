<%
'--- Start PaymentExpress ---
Function gwPaymentExpressEdit()
	call opendb()
	'request gateway variables and insert them into the PaymentExpress table
	query="SELECT pcPay_PaymentExpress_Username,pcPay_PaymentExpress_Password FROM pcPay_PaymentExpress Where pcPay_PaymentExpress_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	pcPay_PaymentExpress_Username2=rs("pcPay_PaymentExpress_Username")
	pcPay_PaymentExpress_Password2=rs("pcPay_PaymentExpress_Password")
	
	pcPay_PaymentExpress_Username=request.Form("pcPay_PaymentExpress_Username")
	if pcPay_PaymentExpress_Username="" then
		pcPay_PaymentExpress_Username=pcPay_PaymentExpress_Username2
	end if
	pcPay_PaymentExpress_Password=request.Form("pcPay_PaymentExpress_Password")
	if pcPay_PaymentExpress_Password="" then
		pcPay_PaymentExpress_Password=pcPay_PaymentExpress_Password2
	end if
	set rs=nothing

	pcPay_PaymentExpress_TransType=request.Form("pcPay_PaymentExpress_TransType")
	pcPay_PaymentExpress_TestUsername=request.Form("pcPay_PaymentExpress_TestUsername")
	pcPay_PaymentExpress_ReceiptEmail=request.Form("pcPay_PaymentExpress_ReceiptEmail")			
	pcPay_PaymentExpress_TestMode=request.Form("pcPay_PaymentExpress_TestMode")
	if pcPay_PaymentExpress_TestMode="" then
		pcPay_PaymentExpress_TestMode=0
	else
		pcPay_PaymentExpress_TestMode=1
		if pcPay_PaymentExpress_TestUsername="" then
			response.redirect "techErr.asp?error="&Server.URLEncode("An error occured while trying to modify PaymentExpress settings. <b>""Test Username""</b> is a required field when activating Test Mode.")
		End If
	end if
	pcPay_PaymentExpress_Cvc2=request.Form("pcPay_PaymentExpress_Cvc2")
	if pcPay_PaymentExpress_Cvc2="1" then
		pcPay_PaymentExpress_Cvc2=1
	else
		pcPay_PaymentExpress_Cvc2=0
	end if
	pcPay_PaymentExpress_AVS=request.Form("pcPay_PaymentExpress_AVS")
	if pcPay_PaymentExpress_AVS="YES" then
		pcPay_PaymentExpress_AVS="1"
	else
		pcPay_PaymentExpress_AVS="0"
	end if			
	query="UPDATE pcPay_PaymentExpress SET pcPay_PaymentExpress_TransType='"&pcPay_PaymentExpress_TransType&"', pcPay_PaymentExpress_TestUsername='"&pcPay_PaymentExpress_TestUsername&"', pcPay_PaymentExpress_Username='"&pcPay_PaymentExpress_Username&"', pcPay_PaymentExpress_Password='"&pcPay_PaymentExpress_Password&"', pcPay_PaymentExpress_ReceiptEmail='"&pcPay_PaymentExpress_ReceiptEmail&"', pcPay_PaymentExpress_TestMode="&pcPay_PaymentExpress_TestMode&", pcPay_PaymentExpress_Cvc2="&pcPay_PaymentExpress_Cvc2&", pcPay_PaymentExpress_AVS="&pcPay_PaymentExpress_AVS&" WHERE pcPay_PaymentExpress_ID=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If

	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ", priceToAdd="&priceToAdd&" ,percentageToAdd="&percentageToAdd&" ,paymentNickName='"&paymentNickName&"' WHERE gwCode=47"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	call closedb()
end function

Function gwPaymentExpress()
	varCheck=1
		
	' Test Username
	pcPay_PaymentExpress_TestUsername = request.Form("pcPay_PaymentExpress_TestUsername")
	' Trans Type
	pcPay_PaymentExpress_TransType=request.Form("pcPay_PaymentExpress_TransType")
	' Trans Mode
	pcPay_PaymentExpress_TestMode=request.Form("pcPay_PaymentExpress_TestMode")
	if pcPay_PaymentExpress_TestMode="" then
		pcPay_PaymentExpress_TestMode=0
	else
		pcPay_PaymentExpress_TestMode=1
		if pcPay_PaymentExpress_TestUsername="" then
			response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add PaymentExpress as your payment gateway. <b>""Test Username""</b> is a required field when activating Test Mode.")
		End If
	end if
	' Trans Password
	pcPay_PaymentExpress_Password=request.Form("pcPay_PaymentExpress_Password")
	if pcPay_PaymentExpress_Password="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add PaymentExpress as your payment gateway. <b>""Password""</b> is a required field.")
	End If
	' Trans Username
	pcPay_PaymentExpress_Username=request.Form("pcPay_PaymentExpress_Username")
	if pcPay_PaymentExpress_Username="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add PaymentExpress as your payment gateway. <b>""Username""</b> is a required field.")
	End If
	' Security Code
	pcPay_PaymentExpress_Cvc2=request.Form("pcPay_PaymentExpress_Cvc2")
	if pcPay_PaymentExpress_Cvc2="1" then
		pcPay_PaymentExpress_Cvc2=1
	else
		pcPay_PaymentExpress_Cvc2=0
	end if
	' Email Receipts			
	pcPay_PaymentExpress_ReceiptEmail=request.Form("pcPay_PaymentExpress_ReceiptEmail")
	if pcPay_PaymentExpress_ReceiptEmail="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add PaymentExpress as your payment gateway. <b>""Email""</b> is a required field.")
	End If
	' AVS *** Payment Express does not appear to be compatible at this time, but may in the future
	'pcPay_PaymentExpress_AVS=request.Form("pcPay_PaymentExpress_AVS")
	'if pcPay_PaymentExpress_AVS="YES" then
	'	pcPay_PaymentExpress_AVS="1"
	'else
		pcPay_PaymentExpress_AVS=0
	'end if
	err.clear
	err.number=0
	call openDb()  

	query=        "UPDATE pcPay_PaymentExpress "
	query=query & "SET pcPay_PaymentExpress_TransType='"&pcPay_PaymentExpress_TransType&"', "
	query=query & "pcPay_PaymentExpress_Username='"&pcPay_PaymentExpress_Username&"', " 
	query=query & "pcPay_PaymentExpress_TestUsername='"&pcPay_PaymentExpress_TestUsername&"', "    
	query=query & "pcPay_PaymentExpress_Password='"&pcPay_PaymentExpress_Password&"', "
	query=query & "pcPay_PaymentExpress_ReceiptEmail='"&pcPay_PaymentExpress_ReceiptEmail&"', "
	query=query & "pcPay_PaymentExpress_TestMode="&pcPay_PaymentExpress_TestMode&", "
	query=query & "pcPay_PaymentExpress_Cvc2="&pcPay_PaymentExpress_Cvc2&", "
	query=query & "pcPay_PaymentExpress_AVS="&pcPay_PaymentExpress_AVS&" "
	query=query & "WHERE pcPay_PaymentExpress_ID=1; "
	set rs=Server.CreateObject("ADODB.Recordset")  
	set rs=connTemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

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
	query = ""
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'PaymentExpress','gwPaymentExpress.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",47,'"&paymentNickName&"')"
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

if request("gwchoice")="47" then
	if request("mode")="Edit" then
		call opendb()
		query="SELECT pcPay_PaymentExpress_TransType, pcPay_PaymentExpress_TestUsername, pcPay_PaymentExpress_Username, pcPay_PaymentExpress_Password, pcPay_PaymentExpress_ReceiptEmail,  pcPay_PaymentExpress_TestMode, pcPay_PaymentExpress_AVS, pcPay_PaymentExpress_Cvc2 FROM pcPay_PaymentExpress WHERE pcPay_PaymentExpress_ID=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If 
		pcPay_PaymentExpress_TransType=rs("pcPay_PaymentExpress_TransType") ' auth or sale
		pcPay_PaymentExpress_Username=rs("pcPay_PaymentExpress_Username") ' username
		pcPay_PaymentExpress_TestUsername=rs("pcPay_PaymentExpress_TestUsername") ' username
		pcPay_PaymentExpress_Password=rs("pcPay_PaymentExpress_Password") ' password
		pcPay_PaymentExpress_ReceiptEmail=rs("pcPay_PaymentExpress_ReceiptEmail") ' additional receipt email
		pcPay_PaymentExpress_TestMode=rs("pcPay_PaymentExpress_TestMode")  ' test mode or live mode
		pcPay_PaymentExpress_AVS=rs("pcPay_PaymentExpress_AVS") ' avs "on" or "off"
		pcPay_PaymentExpress_Cvc2=rs("pcPay_PaymentExpress_Cvc2") ' cvc "on" or "off"
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=47"
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
	<input type="hidden" name="addGw" value="47">
<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/paymentexpress.jpg" width="276" height="42"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
<table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
<tr>
    <td><strong><br>
    </strong>DPS is a high growth, innovative global leader in payment technology. Providing PCI DSS compliant payment solutions under the Payment Express brand.<strong><br>
    <br>
    <a href="http://www.paymentexpress.com/Technical_Resources/Ecommerce_NonHosted/PxPost.aspx" target="_blank">Payment Express Website</a></strong><br />
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
					<% dim pcPay_PaymentExpress_UsernameCnt,pcPay_PaymentExpress_UsernameEnd,pcPay_PaymentExpress_UsernameStart
                    pcPay_PaymentExpress_UsernameCnt=(len(pcPay_PaymentExpress_Username)-2)
                    pcPay_PaymentExpress_UsernameEnd=right(pcPay_PaymentExpress_Username,2)
                    pcPay_PaymentExpress_UsernameStart=""
                    for c=1 to pcPay_PaymentExpress_UsernameCnt
                        pcPay_PaymentExpress_UsernameStart=pcPay_PaymentExpress_UsernameStart&"*"
                    next
                    %>
                    
                    <tr>
                        <td height="31" colspan="2">Account Username:&nbsp;<%=pcPay_PaymentExpress_UsernameStart&pcPay_PaymentExpress_UsernameEnd%></td>
                    </tr>
                    <tr>
                        <td colspan="2"> For security reasons, your &quot;Account Username&quot; 
                            is only partially shown on this page. If you need 
                            to edit your account information, please re-enter 
                            your &quot;Account Username&quot; and 
                        &quot;Account Password&quot; below.</td>
                    </tr>
                <% end if %>
                <tr>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                    <td width="24%"><div align="right">Account Username:</div></td>
                    <td width="76%"> <div align="left"><input type="text" value="" name="pcPay_PaymentExpress_Username" size="30"></div></td>
                </tr>
                
                <tr> 
                    <td width="24%"><div align="right">Test Mode Username:</div></td>
                    <td width="76%"> 
                    <div align="left">
                    <input type="text" value="<%=pcPay_PaymentExpress_TestUsername%>" name="pcPay_PaymentExpress_TestUsername" size="30">
                    <font color="#FF0000">&nbsp;&nbsp;*Required if you are activating Test Mode. </font></div></td>
                </tr>
                <tr> 
                    <td width="24%"><div align="right">Account Password:</div></td>
                    <td width="76%"> <div align="left"><input type="text" value="<%=pcPay_PaymentExpress_Password%>" name="pcPay_PaymentExpress_Password" size="30"></div></td>
                </tr>
                <tr> 
                    <td width="24%"><div align="right">Email Address:</div></td>
                    <td width="76%"> <div align="left"><input type="text" value="<%=pcPay_PaymentExpress_ReceiptEmail%>" name="pcPay_PaymentExpress_ReceiptEmail" size="30"></div></td>
                </tr>
    
                <tr> 
                    <td> <div align="right">Transaction Type:</div></td>
                    <td> 
                    <select name="pcPay_PaymentExpress_TransType">
                        <option value="SALE" selected>Sale</option>
                        <option value="AUTH" <%if pcPay_PaymentExpress_TransType="AUTH" then%>selected<%end if%>>Authorize Only</option>
                    </select></td>
                </tr>
                
                <tr> 
                    <td> <div align="right">Require CVV:</div></td>
                    <td> <input type="radio" class="clearBorder" name="pcPay_PaymentExpress_Cvc2" value="1" checked>Yes 
                    <input name="pcPay_PaymentExpress_Cvc2" type="radio" class="clearBorder" value="0" <%if pcPay_PaymentExpress_Cvc2=0 then%>checked<%end if%>>No
                    <font color="#FF0000">&nbsp;&nbsp;*Required if you are accepting Discover cards.</font></td>
                </tr>
    
                <tr> 
                    <td><div align="right"> 
                    <input name="pcPay_PaymentExpress_TestMode" type="checkbox" class="clearBorder" id="pcPay_PaymentExpress_TestMode" value="1" <% if pcPay_PaymentExpress_TestMode=1 then %>checked<% end if %> /></div></td>
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
        var CollapsiblePanel2 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel2", {contentIsOpen:false});;
        var CollapsiblePanel3 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel3", {contentIsOpen:false});
        var CollapsiblePanel4 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel4", {contentIsOpen:false});
        //-->
        </script>
    </td>
</tr>
</table>
<!-- New View End --><% end if %>

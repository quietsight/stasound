<% 

'---Start eProcessingNetwork---
Function gwEPNEdit()
	call opendb()
	'request gateway variables and insert them into the EPN table
	query="SELECT pcPay_EPN_Account FROM pcPay_EPN Where pcPay_EPN_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription)
	end If
	pcPay_EPN_Account2=rs("pcPay_EPN_Account")
	pcPay_EPN_Account2=enDeCrypt(pcPay_EPN_Account2, scCrypPass)
	set rs=nothing
	pcPay_EPN_Account=request.Form("pcPay_EPN_Account")
	if pcPay_EPN_Account="" then
		pcPay_EPN_Account=pcPay_EPN_Account2
	end if
	pcPay_EPN_CVV=request.Form("pcPay_EPN_CVV")
	pcPay_EPN_TestMode=request.Form("pcPay_EPN_TestMode")
	if pcPay_EPN_TestMode="1" then
		pcPay_EPN_TestMode=1
	else
		pcPay_EPN_TestMode=0
	end if
	pcPay_EPN_Account=enDeCrypt(pcPay_EPN_Account, scCrypPass)

	query="UPDATE pcPay_EPN SET pcPay_EPN_Account='"&pcPay_EPN_Account&"',pcPay_EPN_CVV="&pcPay_EPN_CVV&",pcPay_EPN_TestMode="&pcPay_EPN_TestMode&",pcPay_EPN_RestrictKey='"&pcPay_EPN_RestrictKey&"' WHERE pcPay_EPN_ID=1"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription)
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=42"
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

Function gwEPN()
	varCheck=1
	'request gateway variables and insert them into the EPN table
	pcPay_EPN_Account=request.Form("pcPay_EPN_Account")
	pcPay_EPN_CVV=request.Form("pcPay_EPN_CVV")
	pcPay_EPN_testmode=request.Form("pcPay_EPN_testmode")
	if pcPay_EPN_testmode="" then
		pcPay_EPN_testmode="0"
	end if
	pcPay_EPN_RestrictKey=request.Form("pcPay_EPN_RestrictKey")

	If pcPay_EPN_Account="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add eProcessing Network as your payment gateway. <b>""Account Number""</b> is a required field.")
	End If
	'encrypt
	pcPay_EPN_Account=enDeCrypt(pcPay_EPN_Account, scCrypPass)

	If pcPay_EPN_RestrictKey="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add eProcessing Network as your payment gateway. <b>""RestrictKey""</b> is a required field.")
	End If
	'encrypt
	pcPay_EPN_RestrictKey=enDeCrypt(pcPay_EPN_RestrictKey, scCrypPass)
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

	query="UPDATE pcPay_EPN SET pcPay_EPN_Account='"&pcPay_EPN_Account&"',pcPay_EPN_CVV="&pcPay_EPN_CVV&",pcPay_EPN_TestMode=" & pcPay_EPN_TestMode & ",pcPay_EPN_RestrictKey='"&pcPay_EPN_RestrictKey&"' WHERE pcPay_EPN_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'EPN','gwEPN.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",42,'"&paymentNickName&"')"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	call closedb()
end function
%>

<% if request("gwchoice")="42" then
	'//Check if restrict key Field Exists
	on error resume next
	err.clear
	call opendb()
	query="SELECT * FROM pcPay_EPN;"
	set rsChkObj = Server.CreateObject("ADODB.Recordset")
	set rsChkObj = conntemp.execute(query)
	chkRestrictKey = rsChkObj("pcPay_EPN_RestrictKey")
	if err.number<>0 then
		set rsChkObj=nothing
		call closedb()
		response.redirect "upddbEPN.asp?mode=Edit&id=42"
	else
		set rsChkObj=nothing
		call closedb()
	end if

	if request("mode")="Edit" then
		call opendb()
		query="SELECT pcPay_EPN_Account,pcPay_EPN_CVV,pcPay_EPN_testmode,pcPay_EPN_RestrictKey FROM pcPay_EPN Where pcPay_EPN_ID=1;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription)
		end If
		pcPay_EPN_Account=rs("pcPay_EPN_Account")
		pcPay_EPN_Account=enDeCrypt(pcPay_EPN_Account, scCrypPass)
		pcPay_EPN_RestrictKey=enDeCrypt(pcPay_EPN_RestrictKey, scCrypPass)
		pcPay_EPN_CVV=rs("pcPay_EPN_CVV")
		pcPay_EPN_testmode=rs("pcPay_EPN_testmode")
		pcPay_EPN_RestrictKey=rs("pcPay_EPN_RestrictKey")
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=42"
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
	<input type="hidden" name="addGw" value="42">
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/eprocessingnetwork_logo.png" width="240" height="120"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
<table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
<tr>
    <td><strong>eProcessing Network - Transparent Database Engine Template<br>
    <br>
    </strong>Gateway description goes here!<strong><br>
    <br>
    <a href="http://www.eprocessingnetwork.com/" target="_blank">Website</a></strong><br />
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
					<% dim pcPay_EPN_AccountCnt,pcPay_EPN_AccountEnd,pcPay_EPN_AccountStart
					pcPay_EPN_AccountCnt=(len(pcPay_EPN_Account)-2)
					pcPay_EPN_AccountEnd=right(pcPay_EPN_Account,2)
					pcPay_EPN_AccountStart=""
					for c=1 to pcPay_EPN_AccountCnt
					pcPay_EPN_AccountStart=pcPay_EPN_AccountStart&"*"
					next %>
					<tr>
						<td height="31" colspan="2">Current Store ID:&nbsp;<%=pcPay_EPN_AccountStart&pcPay_EPN_AccountEnd%></td>
					</tr>
					<tr>
						<td colspan="2"> For security reasons, your &quot;Account Number&quot;
							is only partially shown on this page. If you need
							to edit your account information, please re-enter
							your &quot;Account Number&quot; below.</td>
					</tr>
				<% end if %>
				<tr>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				  </tr>
				<tr>
					<td width="111"> <div align="right">Account Number:</div></td>
					<td width="360"> <input type="text" name="pcPay_EPN_Account" size="20"></td>
				</tr>
				<tr>
					<td width="111"> <div align="right">RestrictKey:</div></td>
					<td width="360"> <input type="text" name="pcPay_EPN_RestrictKey" size="20"></td>
				</tr>
				<tr>
					<td width="111"> <div align="right">Require CVV:</div></td>
					<td width="360">
						<input type="radio" class="clearBorder" name="pcPay_EPN_CVV" value="1" checked>
						Yes
						<input name="pcPay_EPN_CVV" type="radio" class="clearBorder" value="0" <% if pcPay_EPN_CVV="0" then %>checked<% end if %> />
						No</td>
				</tr>
				<tr>
					<td> <div align="right">
							<input name="pcPay_EPN_TestMode" type="checkbox" class="clearBorder" id="pcPay_EPN_TestMode" value="1" <% if pcPay_EPN_testmode=1 then%>checked<% end if%>>
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
<% end if %>
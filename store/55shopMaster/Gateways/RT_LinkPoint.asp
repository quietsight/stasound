<%
Function gwlpEdit()
	call opendb()
	'request gateway variables and insert them into the linkpoint table
	query= "SELECT storename FROM linkpoint where id=1"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription)
	end If
	storeName2=rs("storename")
	set rs=nothing
	storeName=request.Form("storeName")
	If storeName="" then
		storeName=storeName2
	End If
	lp_transType=request.Form("lp_transType")
	lp_testmode=request.Form("lp_testmode")
	lp_cards=request.Form("lp_cards")
	lp_CVM=request.Form("lp_CVM")
	lp_yourpay=request.Form("lp_yourpay")
	if NOT isNumeric(lp_CVM) OR lp_CVM="" then
		lp_CVM=0
	end if
	query="UPDATE linkpoint SET storeName='"&storeName&"',transType='"&lp_transType&"',lp_testmode='"&lp_testmode&"',lp_cards='"&lp_cards&"',CVM="&lp_CVM&",lp_yourpay='"&lp_yourpay&"' WHERE ID=1"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription)
	end If
	if lp_yourpay = "API" Then
	   sslURL = "GwLPApi.asp"
	else
	   sslURL = "gwlp.asp"
	end if


	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus &",sslURL='"&sslURL& "', priceToAdd="&priceToAdd&" ,percentageToAdd="&percentageToAdd&" ,paymentNickName='"&paymentNickName&"' WHERE gwCode=8"
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

Function gwlp()
	varCheck=1
	'request gateway variables and insert them into the linkpoint table
	storename=request.Form("storename")
	lp_transtype=request.Form("lp_transtype")
	lp_testmode=request.Form("lp_testmode")
	lp_cards=request.Form("lp_cards")
	lp_CVM=request.Form("lp_CVM")
	if NOT isNumeric(lp_CVM) or lp_CVM="" then
		lp_CVM=0
	end if
	lp_yourpay=request.Form("lp_yourpay")

	If storename="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add LinkPoint as your payment gateway. <b>""store name""</b> is a required field.")
	End If
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
	end if

	err.clear
	err.number=0

	call openDb()

	query="UPDATE linkpoint SET storeName='"&storename&"', transType='"&lp_transtype&"',lp_testmode='"&lp_testmode&"',lp_cards='"&lp_cards&"',CVM="&lp_CVM&",lp_yourpay='"&lp_yourpay&"' WHERE id=1"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if lp_yourpay = "API" Then
	   sslURL = "GwLPApi.asp"
	else
	   sslURL = "gwlp.asp"
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'LinkPoint','"&sslURL&"',-1,0,0,9999,0,9999,0,9999,-1,"&priceToAdd&","&percentageToAdd&",8,'"&paymentNickName&"')"
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

<% if request("gwchoice")="8" then
	on error resume next
	err.clear
	call opendb()
	query="select * from pcPay_LinkPointAPI"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if err.number<>0 then
		set rs=nothing
		call closedb()
		response.redirect "upddbLPApi.asp"
	else
		set rs=nothing
		call closedb()
	end if

	isComErr = Cint(0)
	strComErr = Cstr()
	err.clear
	err.number=0
	Set objTest = Server.CreateObject("LpiCom_6_0.LPOrderPart")
	if err.number<>0 then
		isComErr = 1
		strComErr = strComErr & "LpiCom_6_0.LPOrderPart<BR>"
	end if

	Set objTest = Server.CreateObject("LpiCom_6_0.LinkPointTxn")
	if err.number<>0 then
		isComErr = 1
		strComErr = strComErr & "LpiCom_6_0.LinkPointTxn<BR>"
	end if

	if request("mode")="Edit" then
		call opendb()
		query= "SELECT storename, transType, lp_testmode, lp_cards, CVM, lp_yourpay FROM linkpoint where id=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription)
		end If
		storeName=rs("storeName")
		lp_transType=rs("transType")
		lp_testmode=rs("lp_testmode")
		lp_cards=rs("lp_cards")
		lp_CVM=rs("CVM")
		lp_yourpay=rs("lp_yourpay")
		
						
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=8"
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

		lV="0"
		lM="0"
		lA="0"
		lD="0"
		'V, M, A
		cardTypeArray=split(lp_cards,", ")
		for i=0 to ubound(cardTypeArray)
			select case cardTypeArray(i)
			case "V"
				lV="1"
			case "M"
				lM="1"
			case "A"
				lA="1"
			case "D"
				lD="1"
			end select
		next

		dim storeNameCnt,storeNameEnd,storeNameStart
		storeNameCnt=(len(storeName)-2)
		storeNameEnd=right(storeName,2)
		storeNameStart=""
		for c=1 to storeNameCnt
			storeNameStart=storeNameStart&"*"
		next %>
		<input type="hidden" name="mode" value="Edit">
		<%
	end if
	%>
	<input type="hidden" name="addGw" value="8">
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/FirstDatalogo.jpg" width="293" height="112"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
<table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
<tr>
    <td><strong><br>
    </strong>First Data online payment processing solutions deliver the innovation and security you and your customers demand to keep pace in this ever-changing marketplace.<strong><br>
    <br>
    <a href="http://www.linkpoint.com" target="_blank">First Data Website </a></strong><br />
<br />
</td>
</tr>
<tr>
    <td>
        <div id="CollapsiblePanel1" class="CollapsiblePanel">
            <div class="CollapsiblePanelTab1">
                <table width="100%">
                  <tr>
                    <td width="580" class="pcPanelTitle1">Step 1: Select Integration</td>
                  </tr>
                </table>
            </div>
            <div class="CollapsiblePanelContent">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
				  <td colspan="2">&nbsp;</td>
				  </tr>
				<tr>
					<td colspan="2">
					<!--' lp_yourpaY = "" "YES" YOUR PAY, "API" LINK POINT API-->
					  <p>
						  <input name="lp_yourpay" type="radio" class="clearBorder" id="lp_yourpay" value="" <% if lp_yourpay="" then %>checked<% end if %>>
						  I am using <b>FDGG Connect</b> <strong>(LinkPoint Basic)</strong>
					  <br>
						  <input name="lp_yourpay" type="radio" class="clearBorder" id="lp_yourpay" value="YES" <% if lp_yourpay="YES" then %>checked<% end if %>>
						  I am using <b>FDGG  Virtual Terminal</b><strong> (YourPay)</strong>  - <a href="http://www.YourPay.com" target="_blank">Web site</a>
						  <BR>
						<% if isComErr = 1 then %>
						<input type="radio" disabled="disabled" class="clearBorder" /> I am using <strong>FDGG API (LinkPoint API) with Batch Post Authorizations</strong><br />

			&nbsp;<img src="images/red_x.png" width="12" height="12" /><strong> &nbsp;FDGG API (LinkPoint API) with Batch Post Authorizations</strong> cannot be enabled. Errors were found while testing for the required components. These library files are available for download directly from First Data and need to be installed directly on the server.<br />
							<center>
						<strong>Required components for FDGG API with Batch Post Authorizations:</strong><br /><i><%= strComErr %></i></center><br /><br />

						<% else %>
							<input name="lp_yourpay" type="radio" class="clearBorder" id="lp_yourpay" value="API" <% if lp_yourpay="API" then %>checked<% end if %>>
							I am using <strong>FDGG API (LinkPoint API) with Batch Post Authorizations</strong>
							<BR>
							API Keys Directory: productcart/pcadmin/ folder (Copy your sercurity keys file *.pem to this folder before activating this payment gateway)
						<% End If %>
						</td>
				</tr>
				<% if request("mode")="Edit" then %>
					<tr>
						<td colspan="2">Current Store Name:&nbsp;<%=storeNameStart&storeNameEnd %></td>
					</tr>
					<tr>
						<td colspan="2"> For security reasons, your &quot;Store Name&quot; is
							only partially shown on this page. If you need to edit
							your account information, please re-enter your &quot;Store
							Name&quot; below.</td>
					</tr>
			   <% end if %>
				<tr>
					<td width="24%"> <div align="right">Store Name: </div></td>
					<td width="76%"> <div align="left">
							<input type="text" value="" name="storename" size="30">
						</div></td>
				</tr>
				<tr>
					<td> <div align="right">Transaction Type:</div></td>
					<td>
						<select name="lp_transType">
							<option value="sale" selected>Sale</option>
							<option value="preauth" <% if lp_transType="preauth" then %>selected<% end if %>>Authorize Only</option>
						</select>
				   </td>
				</tr>
				<tr>
					<td> <div align="right">Require CVM:</div></td>
					<td> <input type="radio" class="clearBorder" name="lp_CVM" value="1" checked>
						Yes
						<input name="lp_CVM" type="radio" class="clearBorder" value="0" <% if lp_CVM="0" then%>checked<%end if%>>
						No<font color="#FF0000">&nbsp;&nbsp;*Required if you
						are accepting Discover cards.</font></td>
				</tr>
				<tr>
					<td> <div align="right">Accepted Cards:</div></td>
					<td>
						<input name="lp_cards" type="checkbox" class="clearBorder" value="V" <% if lV="1" then%>checked<% end if %>>                        Visa
						<input name="lp_cards" type="checkbox" class="clearBorder" value="M" <% if lM="1" then%>checked<% end if %>>                        Master Card
						<input name="lp_cards" type="checkbox" class="clearBorder" value="A" <% if lA="1" then%>checked<% end if %>>                        American Express
						<input name="lp_cards" type="checkbox" class="clearBorder" value="D" <% if lD="1" then%>checked<% end if %>>                        Discover
					</td>
				</tr>
				<tr>
					<td><div align="right">
							<input name="lp_testmode" type="checkbox" class="clearBorder" value="YES" <% if lp_testmode="YES" then %>checked<% end if %> />
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
<!-- New View End --><% end if %>

<%
'---Start Protx---
Function gwprotxEdit()
	call opendb()
	'request gateway variables and insert them into the protx table
	query="SELECT Protxid,ProtxPassword,ProtxTestmode,CVV,TxType, ProtxCurcode, avs FROM protx WHERE idProtx=1"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	Protxid2=rstemp("Protxid")
	Protxid=request.Form("Protxid")
	if Protxid="" then
		Protxid=Protxid2
	end if
	
	ProtxPassword2=rstemp("ProtxPassword")
	'decrypt
	ProtxPassword2=enDeCrypt(ProtxPassword2, scCrypPass)
	ProtxPassword=request.Form("ProtxPassword")
	if ProtxPassword="" then
		ProtxPassword=ProtxPassword2
	end if
	'encrypt
	ProtxPassword=enDeCrypt(ProtxPassword, scCrypPass)
	
	ProtxCurcode2=rstemp("ProtxCurcode")
	ProtxCurcode=request.Form("ProtxCurcode")
	If ProtxCurcode="" then
		ProtxCurcode=ProtxCurcode2
	end if
	
	ProtxTestmode=request.Form("ProtxTestmode")
	if ProtxTestmode="" then
		ProtxTestmode="0"
	end if
	CVV=request.Form("CVV")
	avs=request.Form("avs")
	if avs="1" then
		avs=1
	else
		avs=0
	end if
	TxType2=rstemp("TxType")
	TxType=request.Form("TxType")
	If TxType="" then
		TxType=TxType2
	end if
	ProtxMethod=request.Form("ProtxMethod")
	if ProtxMethod="DIRECT" then
		ProtxURL="gwprotxVSP.asp"
	else
		ProtxURL="gwprotx.asp"
	end if
	ProtxApply3DSecure=request.Form("ProtxApply3DSecure")
	if ProtxApply3DSecure="" then
		ProtxApply3DSecure=3
	end if

	ProtxCardTypes=request.Form("ProtxCardTypes")
	query="UPDATE protx SET Protxid='"&Protxid&"', ProtxPassword='"&ProtxPassword&"',ProtxTestmode="&ProtxTestmode&",CVV="&CVV&",ProtxCurcode='"&ProtxCurcode&"',TxType='"&TxType&"', avs="&avs&", ProtxMethod='"&ProtxMethod&"', ProtxCardTypes='"&ProtxCardTypes&"', ProtxApply3DSecure="&ProtxApply3DSecure&" WHERE idProtx=1;"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ", sslURL='"&ProtxURL&"', pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=26"

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

Function gwprotx()
	varCheck=1
	'request gateway variables and insert them into the protx table
	Protxid=request.Form("Protxid")
	If Protxid="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Sage Pay as your payment gateway. <b>""Vendor Login Name""</b> is a required field.")
	End If
	
	ProtxMethod=request.Form("ProtxMethod")
	
	ProtxPassword=request.Form("ProtxPassword")
	If ProtxPassword="" AND ProtxMethod="FORM" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Sage Pay as your payment gateway. <b>""Password""</b> is a required field.")
	End If
	'encrypt
	ProtxPassword=enDeCrypt(ProtxPassword, scCrypPass)
	
	CVV=request.Form("CVV")
	if NOT isNumeric(CVV) or CVV="" then
		CVV=0
	end if
	
	ProtxCurcode=request.Form("ProtxCurcode")
	
	avs=request.Form("avs")
	if avs="1" then
		avs=1
	else
		avs=0
	end if
	
	ProtxTestmode=request.Form("ProtxTestmode")
	if ProtxTestmode="" then
		ProtxTestmode="0"
	end if
	
	TxType=request.Form("TxType")
	ProtxCardTypes=request.Form("ProtxCardTypes")

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
	ProtxApply3DSecure=request.Form("ProtxApply3DSecure")
	if ProtxApply3DSecure="" then
		ProtxApply3DSecure=3
	end if
	
	err.clear
	err.number=0
	call openDb() 
	
	query="UPDATE protx SET Protxid='"&Protxid&"',ProtxPassword='"&ProtxPassword&"',CVV="&CVV&",ProtxTestmode="&ProtxTestmode&",ProtxCurcode='"&ProtxCurcode&"', TxType='"&TxType&"', avs="&avs&", ProtxMethod='"&ProtxMethod&"', ProtxCardTypes='"&ProtxCardTypes&"', ProtxApply3DSecure="&ProtxApply3DSecure&" WHERE idProtx=1"
	set rs=Server.CreateObject("ADODB.Recordset")  

	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	if ProtxMethod="DIRECT" then
		ProtxURL="gwprotxVSP.asp"
	else
		ProtxURL="gwprotx.asp"
	end if
	
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'Protx','"&ProtxURL&"',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",26,'"&paymentNickName&"')"
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

if request("gwchoice")="26" then
	if request("mode")="Edit" then
		call opendb()
		query= "SELECT Protxid,ProtxPassword,ProtxTestmode,CVV,TxType, avs, ProtxCurcode,ProtxMethod, ProtxCardTypes, ProtxApply3DSecure FROM protx WHERE idProtx=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		Protxid=rs("Protxid")
		ProtxPassword=rs("ProtxPassword")
			'decrypt
			ProtxPassword=enDeCrypt(ProtxPassword, scCrypPass)
		ProtxTestmode=rs("ProtxTestmode")
		ProtxCurcode=rs("ProtxCurcode")
		CVV=rs("CVV")
		avs=rs("avs")
		TxType=rs("TXType")
		ProtxMethod=rs("ProtxMethod")
		ProtxCardTypes=rs("ProtxCardTypes")
		ProtxApply3DSecure=rs("ProtxApply3DSecure")
		cardTypeArray=split(ProtxCardTypes,", ")
		for i=0 to ubound(cardTypeArray)
			select case cardTypeArray(i)
				case "MC"
					MC="1" 
				case "VISA"
					VISA="1"
				case "DELTA"
					DELTA="1"
				case "AMEX"
					AMEX="1"
				case "SOLO"
					SOLO="1"
				case "MAESTRO"
					MAESTRO="1"
				case "UKE"
					UKE="1"
				case "DC"
					DC="1"
				case "JCB"
					JCB="1"
			end select
		next
		if isNULL(ProtxApply3DSecure) OR ProtxApply3DSecure="" then
			ProtxApply3DSecure=3
		end if
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=26"
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
	<input type="hidden" name="addGw" value="26">
<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/sagepay_logo.jpg" width="313" height="90"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
<table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
<tr>
    <td><strong><br>
    </strong>Security and service is at the heart of Sage Pay's business so that you can depend on us for fast, secure and reliable payments.<strong><br>
    <br>
    <a href="http://www.sagepay.com/" target="_blank">SagePay Website</a></strong><br />
<br />
</td>
</tr>
<tr>
    <td>
        <div id="CollapsiblePanel1" class="CollapsiblePanel">
            <div class="CollapsiblePanelTab1">
                <table width="100%">
                  <tr>
                    <td width="580" class="pcPanelTitle1">Step 1: Configure Account - Sage Pay</td>
                  </tr>
                </table>
            </div>
            <div class="CollapsiblePanelContent">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td align="right" valign="top">&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                    <td width="111" align="right" valign="top"><input name="ProtxMethod" type="radio" value="FORM" <% if ProtxMethod <> "DIRECT" then%>checked<% end if %> onClick="document.getElementById('cvv').style.display='none';" ></td>
                    <td><strong>Form</strong><br>Smaller online merchants often prefer their websites to be managed by third-party hosting companies. Unfortunately, many web management companies do not allow specific components to be installed on their servers, so applications like Sage Pay Server and Sage Pay Direct are not an option. In such circumstances, VSP Form is the ideal system for your website. It operates in a similar manner to the other Payment Service Providers, sending transaction registration information in hidden HTML fields through the customers browser, but unlike many systems, Sage Pay Form encrypts the important information first, so the customer cannot tamper with it.</td>
                </tr>
                <tr> 
                    <td align="right" valign="top"><input type="radio" value="DIRECT" name="ProtxMethod" <% if ProtxMethod="DIRECT" then%>checked<% end if %> onClick="document.getElementById('cvv').style.display='';" ></td>
                    <td valign="top"><b>Direct</b><br>This is the most secure and <u>preferred</u> method of processing payments.&nbsp;<br>Some online merchants wish to take the credit or debit card details on their own site, rather than redirecting the customer to Sage Pay. There are a number of reasons for doing this, for instance, you may want to periodically charge the same card, perform additional fraud checking, or if you already take credit cards you may want an additional means of authorisation. Sage Pay Direct allows you to simply send us the credit card details in a server-to-server, highly-encrypted message, the details are then sent to the merchant bank to obtain authorisation, the Sage Pay server replies immediately, with no redirection or callbacks. </p>
                    </td>
                </tr>
                <% if request("mode")="Edit" then %>
					<% dim ProtxIDCnt,ProtxIDEnd,ProtxIDStart
                    ProtxIDCnt=(len(ProtxID)-2)
                    ProtxIDEnd=right(ProtxID,2)
                    ProtxIDStart=""
                    for c=1 to ProtxIDCnt
                        ProtxIDStart=ProtxIDStart&"*"
                    next %>
                    <tr> 
                        <td height="31" colspan="2">Current Vendor Login Name:&nbsp;<%=ProtxIDStart&ProtxIDEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;Vendor Login Name&quot; are only partially shown on this page. If you need to edit your account information, please re-enter your &quot;Vendor Login Name&quot; and &quot;Password&quot; below.</td>
                    </tr>
                <% end if %>
                <tr>
                    <td> <div align="right">Vendor Login Name:</div></td>
                    <td> <input type="text" name="Protxid" size="20"></td>
                </tr>
                <tr> 
                    <td> <div align="right">Encryption Password:</div></td>
                    <td> <input name="ProtxPassword" type="password" size="20">
                        *Only required for<strong> Sage Pay Form</strong> </td>
                </tr>
                <tr> 
                    <td> <div align="right">Transaction Type:</div></td>
                    <td> <select name="TxType">
                            <option value="PAYMENT" selected>Sale</option>
                            <option value="DEFERRED" <% if TxType="DEFERRED" then%>selected<% end if %>>Deferred</option>
                            <option value="AUTHENTICATE" <% if TxType="AUTHENTICATE" then%>selected<% end if %>>Authenticate Only</option>
                        </select> &nbsp;&nbsp; <input name="avs" type="checkbox" class="clearBorder" value="1" <% if avs=1 then%>checked<% end if %>>
                        Enable AVS Mode <font size=1>(Address Verification Service)</font></td>
                </tr>
                <% if ProtxCurcode="" then
					 ProtxCurcode="GBP"
			 	end if %>
                <tr> 
                    <td> <div align="right">Currency Code:</div></td>
                    <td> <input type="text" name="ProtxCurcode" value="<%=ProtxCurcode%>" size="8"> 
                        <font size=1>(Please ask Protx about the code to be used for your currency)</font> </td>
                </tr>
                <tr id="cvv" style="display:<%if ProtxMethod="DIRECT" then%>table-row<%else%>none<%end if%>;" > 
                    <td> <div align="right">Require CVV:</div></td>
                    <td> <input type="radio" class="clearBorder" name="CVV" value="1" checked>
                        Yes 
                        <input type="radio" class="clearBorder" name="CVV" value="0" <% if CVV="0" then%>checked<% end if %>>
                        No</td>
                </tr>
                <tr>
                  <td>Enable 3DSecure:</td>
                  <td><input type="radio" class="clearBorder" name="ProtxApply3DSecure" value="0" checked="checked" />
Yes
  <input type="radio" class="clearBorder" name="ProtxApply3DSecure" value="3" <% if ProtxApply3DSecure="3" then%>checked<% end if %> />
No&nbsp;&nbsp;*Only available for<strong> Sage Pay Direct </strong>(Your account must be 3D-enabled)</td>
                </tr>
                <tr> 
                    <td><div align="right">Accepted Cards:</div></td>
                    <td>
						<% if VISA="1" then %>
                        <input name="ProtxCardTypes" type="checkbox" class="clearBorder" value="VISA" checked> 
                        <% else %> <input name="ProtxCardTypes" type="checkbox" class="clearBorder" value="VISA"> 
                        <% end if %>
                        VISA 
                        <% if MC="1" then %> <input name="ProtxCardTypes" type="checkbox" class="clearBorder" value="MC" checked> 
                        <% else %> <input name="ProtxCardTypes" type="checkbox" class="clearBorder" value="MC"> 
                        <% end if %>
                        MasterCard 
                        <% if AMEX="1" then %> <input name="ProtxCardTypes" type="checkbox" class="clearBorder" value="AMEX" checked> 
                        <% else %> <input name="ProtxCardTypes" type="checkbox" class="clearBorder" value="AMEX"> 
						<% end if %>
                        American Express 
                        <% if DELTA="1" then %> <input name="ProtxCardTypes" type="checkbox" class="clearBorder" value="DELTA" checked> 
                        <% else %> <input name="ProtxCardTypes" type="checkbox" class="clearBorder" value="DELTA"> 
                        <% end if %>
                        Delta 
                        <% if SOLO="1" then %> <input name="ProtxCardTypes" type="checkbox" class="clearBorder" value="SOLO" checked> 
                        <% else %> <input name="ProtxCardTypes" type="checkbox" class="clearBorder" value="SOLO"> 
                        <% end if %>
                        Solo<BR> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                        <% if MAESTRO="1" then %> <input name="ProtxCardTypes" type="checkbox" class="clearBorder" value="MAESTRO" checked> 
                        <% else %> <input name="ProtxCardTypes" type="checkbox" class="clearBorder" value="MAESTRO"> 
                        <% end if %>
                        Maestro 
                        <% if UKE="1" then %> <input name="ProtxCardTypes" type="checkbox" class="clearBorder" value="UKE" checked> 
						<% else %> <input name="ProtxCardTypes" type="checkbox" class="clearBorder" value="UKE"> 
                        <% end if %>
                        VISA Electron 
                        <% if DC="1" then %> <input name="ProtxCardTypes" type="checkbox" class="clearBorder" value="DC" checked> 
                        <% else %> <input name="ProtxCardTypes" type="checkbox" class="clearBorder" value="DC"> 
                        <% end if %>
                        Diners Club 
                        <% if JCB="1" then %> <input name="ProtxCardTypes" type="checkbox" class="clearBorder" value="JCB" checked> 
                        <% else %> <input name="ProtxCardTypes" type="checkbox" class="clearBorder" value="JCB"> 
                        <% end if %>
                        JCB                	</td>
                </tr>
                <tr> 
                    <td> <div align="right"> 
                            <input name="ProtxTestmode" type="radio" class="clearBorder" value="1" <% if ProtxTestmode="1" then %>checked<% end if %> /></div></td>
                    <td><b>Enable Test Mode </b>(ProductCart Test Mode - Credit cards will not be charged)</td>
                </tr>
                <tr> 
                    <td> <div align="right"> 
                            <input name="ProtxTestmode" type="radio" class="clearBorder" value="2" <% if ProtxTestmode="2" then %>checked<% end if %> /></div></td>
                    <td><b>&quot;Going Live&quot; Test Mode </b>(Protx Test Mode - Credit cards will not be charged)</td>
                </tr>
                <tr> 
                    <td> <div align="right"> 
                            <input name="ProtxTestmode" type="radio" class="clearBorder" value="0" <% if ProtxTestmode="0" then %>checked<% end if %> /></div></td>
                    <td><b>&quot;Live&quot; Mode</b></td>
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

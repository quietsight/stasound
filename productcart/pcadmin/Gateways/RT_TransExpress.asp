<%
'---Start TransFirst Express---
Function gwTXPEdit()
	call opendb()
	'request gateway variables and insert them into the USAePay table
	query="SELECT pcPay_TXPGatewayID, pcPay_TXPRegistrationKey FROM pcPay_TransFirstXP where pcPay_TransFirstXPID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	pcPay_TXPGatewayID=rs("pcPay_TXPGatewayID") 
	pcPay_TXPRegistrationKey=rs("pcPay_TXPRegistrationKey") 

	pcPay_TXPGatewayID2=pcPay_TXPGatewayID
	'decrypt
	pcPay_TXPGatewayID2=enDeCrypt(pcPay_TXPGatewayID2, scCrypPass)
	pcPay_TXPGatewayID=request.Form("pcPay_TXPGatewayID")
	if pcPay_TXPGatewayID="" then
		pcPay_TXPGatewayID=pcPay_TXPGatewayID2
	end if
	'encrypt
	pcPay_TXPGatewayID=enDeCrypt(pcPay_TXPGatewayID, scCrypPass)
	
	pcPay_TXPRegistrationKey2=pcPay_TXPRegistrationKey
	'decrypt
	pcPay_TXPRegistrationKey2=enDeCrypt(pcPay_TXPRegistrationKey2, scCrypPass)
	pcPay_TXPRegistrationKey=request.Form("pcPay_TXPRegistrationKey")
	if pcPay_TXPRegistrationKey="" then
		pcPay_TXPRegistrationKey=pcPay_TXPRegistrationKey2
	end if
	'encrypt
	pcPay_TXPRegistrationKey=enDeCrypt(pcPay_TXPRegistrationKey, scCrypPass)
	
	pcPay_TXPTestMode=request.Form("pcPay_TXPTestMode")
	if pcPay_TXPTestMode="" then
		pcPay_TXPTestMode="0"
	end if
	
	pcPay_TXPTransType=request.Form("pcPay_TXPTransType")
	pcPay_TXPReqCardCode=request.Form("pcPay_TXPReqCardCode")
	pcPay_TXPCardTypes=request.Form("pcPay_TXPCardTypes")
	
	query="UPDATE pcPay_TransFirstXP SET pcPay_TXPGatewayID='"&pcPay_TXPGatewayID&"',pcPay_TXPRegistrationKey='"&pcPay_TXPRegistrationKey&"',pcPay_TXPTransType='" & pcPay_TXPTransType & "',pcPay_TXPTestMode="&pcPay_TXPTestMode&",pcPay_TXPReqCardCode="&pcPay_TXPReqCardCode&",pcPay_TXPCardTypes='"&pcPay_TXPCardTypes&"' WHERE pcPay_TransFirstXPID=1"
	
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=70"
	
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

Function gwTXP()
	varCheck=1
	'request gateway variables and insert them into the CyberSource table
	pcPay_TXPGatewayID=request.Form("pcPay_TXPGatewayID")
	pcPay_TXPRegistrationKey=request.Form("pcPay_TXPRegistrationKey")
	pcPay_TXPTransType=request.Form("pcPay_TXPTransType")
	pcPay_TXPTestMode=request.Form("pcPay_TXPTestMode")
	if pcPay_TXPTestMode="" then
		pcPay_TXPTestMode="0"
	end if
	pcPay_TXPReqCardCode=request.Form("pcPay_TXPReqCardCode")
	pcPay_TXPCardTypes=request.Form("pcPay_TXPCardTypes")

	If pcPay_TXPGatewayID="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add TransFirst Express as your payment gateway. <b>""Gateway ID""</b> is a required field.")
	End If
	'encrypt
	pcPay_TXPGatewayID=enDeCrypt(pcPay_TXPGatewayID, scCrypPass)
	If pcPay_TXPRegistrationKey="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add TransFirst Express as your payment gateway. <b>""Registration Key""</b> is a required field.")
	End If
	'encrypt
	pcPay_TXPRegistrationKey=enDeCrypt(pcPay_TXPRegistrationKey, scCrypPass)

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

	query="UPDATE pcPay_TransFirstXP SET pcPay_TXPGatewayID='"&pcPay_TXPGatewayID&"',pcPay_TXPRegistrationKey='"&pcPay_TXPRegistrationKey&"',pcPay_TXPTransType='" & pcPay_TXPTransType & "',pcPay_TXPTestMode="&pcPay_TXPTestMode&",pcPay_TXPReqCardCode="&pcPay_TXPReqCardCode&",pcPay_TXPCardTypes='"&pcPay_TXPCardTypes&"' WHERE pcPay_TransFirstXPID=1"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'TranFirst Express','gwTransFirstXP.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",70,'"&paymentNickName&"')"
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
'---End TransFirst---
%>
<% if request("gwchoice")="70" then
	'Check if table exist in database, if not we will rediect to database update prompt
	err.clear
	call openDb()
	query="select * from pcPay_TransFirstXP"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		set rs=nothing
		call closedb()
		response.redirect "upddbTXP.asp"
	else
		set rs=nothing
		call closedb()
	end if
	if request("mode")="Edit" then
		call opendb()
		query="SELECT pcPay_TransFirstXP.pcPay_TXPGatewayID, pcPay_TransFirstXP.pcPay_TXPRegistrationKey, pcPay_TransFirstXP.pcPay_TXPTransType, pcPay_TransFirstXP.pcPay_TXPTestMode, pcPay_TransFirstXP.pcPay_TXPReqCardCode, pcPay_TransFirstXP.pcPay_TXPCardTypes FROM pcPay_TransFirstXP WHERE (((pcPay_TransFirstXP.pcPay_TransFirstXPID)=1));"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcPay_TXPGatewayID=rs("pcPay_TXPGatewayID")
		pcPay_TXPGatewayID=enDeCrypt(pcPay_TXPGatewayID, scCrypPass)
		pcPay_TXPRegistrationKey=rs("pcPay_TXPRegistrationKey")
		pcPay_TXPRegistrationKey=enDeCrypt(pcPay_TXPRegistrationKey, scCrypPass)
		pcPay_TXPTransType=rs("pcPay_TXPTransType")
		pcPay_TXPTestMode=rs("pcPay_TXPTestMode")
		pcPay_TXPReqCardCode=rs("pcPay_TXPReqCardCode")
		pcPay_TXPCardTypes=rs("pcPay_TXPCardTypes")

		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=70"
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
	<input type="hidden" name="addGw" value="70">

<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/transfirst_logo.jpg" width="400" height="98"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
<table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
<tr>
    <td><strong><br>
    </strong>Transaction Express, TransFirst's new payment gateway, puts control of your business's payment acceptance into your hands,			wherever you are – in your store, at your office, in your favorite coffee shop.  And no matter how you access the Internet – desktop computer, laptop, tablet or smart phone – you'll have the power to help grow your business right at your fingertips.<strong><br>
    <br>
    <a href="http://www.transactionexpress.com/" target="_blank">Transaction Express Website</a></strong><br />
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
                    <% dim pcPay_TXPGatewayIDCnt,pcPay_TXPGatewayIDEnd,pcPay_TXPGatewayIDStart
                    pcPay_TXPGatewayIDCnt=(len(pcPay_TXPGatewayID)-2)
                    pcPay_TXPGatewayIDEnd=right(pcPay_TXPGatewayID,2)
                    pcPay_TXPGatewayIDStart=""
                    for c=1 to pcPay_TXPGatewayIDCnt
                    pcPay_TXPGatewayIDStart=pcPay_TXPGatewayIDStart&"*"
                    next %>
                    <tr> 
                        <td height="31" colspan="2">Current TransFirst Express Gateway ID:&nbsp;<%=pcPay_TXPGatewayIDStart&pcPay_TXPGatewayIDEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;Gateway ID&quot; 
                            is only partially shown on this page. If you need 
                            to edit your account information, please re-enter 
                            your &quot;Gateway ID&quot; and &quot;Registration Key&quot; 
                            below.</td>
                    </tr>
                <% end if %>
                <tr>
                  <td valign="top">&nbsp;</td>
                  <td valign="top">&nbsp;</td>
                </tr>
                <tr> 
                    <td width="111" valign="top"><div align="right">Gateway ID:</div></td>
                  <td valign="top"> <input type="text" name="pcPay_TXPGatewayID" size="20"></td>
                </tr>
                <tr> 
                    <td valign="top"><div align="right">Registration Key:</div></td>
                    <td valign="top"><input type="text" name="pcPay_TXPRegistrationKey" size="20"></td>
                </tr>
                <tr> 
                    <td valign="top"><div align="right">Transaction Type:</div></td>
                    <td valign="top"> <select name="pcPay_TXPTransType">
                            <option value="0" selected>Authorize Only</option>
                            <option value="1" <%if pcPay_TXPTransType="1" then%>selected<% end if %>>Sale</option>
                        </select> </td>
                </tr>
                <tr> 
                    <td> <div align="right">Require Card Security Code:</div></td>
                    <td> <input type="radio" class="clearBorder" name="pcPay_TXPReqCardCode" value="1" checked>
                        Yes 
                        <input name="pcPay_TXPReqCardCode" type="radio" class="clearBorder" value="0" <%if clng(pcPay_TXPReqCardCode)=0 then%>checked<%end if%>>
                        No</td>
                </tr>
                <tr> 
                    <td> <div align="right">Accepted Cards:</div></td>
                  <td> 
                        <input name="pcPay_TXPCardTypes" type="checkbox" class="clearBorder" value="VISA" <%if (instr(pcPay_TXPCardTypes,"VISA,")>0) or (instr(pcPay_TXPCardTypes,"VISA")>0) then%>checked<%end if%>>
                        Visa 
                        <input name="pcPay_TXPCardTypes" type="checkbox" class="clearBorder" value="MAST" <%if (instr(pcPay_TXPCardTypes,"MAST,")>0) or (instr(pcPay_TXPCardTypes,"MAST")>0) then%>checked<%end if%>>
                        MasterCard 
                      <input name="pcPay_TXPCardTypes" type="checkbox" class="clearBorder" value="AMER" <%if (instr(pcPay_TXPCardTypes,"AMER,")>0) or (instr(pcPay_TXPCardTypes,"AMER")>0) then%>checked<%end if%>>
                        American Express 
                        <input name="pcPay_TXPCardTypes" type="checkbox" class="clearBorder" value="DISC" <%if (instr(pcPay_TXPCardTypes,"DISC,")>0) or (instr(pcPay_TXPCardTypes,"DISC")>0) then%>checked<%end if%>>
                      Discover                      </td>
                </tr>
                <tr> 
                    <td> <div align="right"> 
                            <input name="pcPay_TXPTestMode" type="checkbox" class="clearBorder" id="pcPay_TXPTestMode" value="1" <% if pcPay_TXPTestMode=1 then%>checked<% end if%>> 
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

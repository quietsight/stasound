<%
'---Start PayJunction---
Function gwPayJunctionEdit()
	call opendb()
	'request gateway variables and insert them into the PayJunction table
	query="SELECT pcPay_PJ_MerchantID,pcPay_PJ_MerchantPassword,pcPay_PJ_cardTypes, pcPay_PJ_CVC,pcPay_PJ_TestMode FROM pcPay_PayJunction Where pcPay_PJ_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	pcPay_PJ_MerchantID2=rs("pcPay_PJ_MerchantID")
	pcPay_PJ_MerchantID2=enDeCrypt(pcPay_PJ_MerchantID2, scCrypPass)

	pcPay_PJ_MerchantID=request.Form("pcPay_PJ_MerchantID")
	if pcPay_PJ_MerchantID="" then
		pcPay_PJ_MerchantID=pcPay_PJ_MerchantID2
	end if
	pcPay_PJ_MerchantPassword2=rs("pcPay_PJ_MerchantPassword")
	pcPay_PJ_MerchantPassword2=enDeCrypt(pcPay_PJ_MerchantPassword2, scCrypPass)
	set rs=nothing
	pcPay_PJ_MerchantPassword=request.Form("pcPay_PJ_MerchantPassword")
	if pcPay_PJ_MerchantPassword="" then
		pcPay_PJ_MerchantPassword=pcPay_PJ_MerchantPassword2
	end if
	pcPay_PJ_TransType = request.Form("pcPay_PJ_TransType")
	pcPay_PJ_cardTypes=request.Form("pcPay_PJ_cardTypes")
	pcPay_PJ_CVC= request.Form("pcPay_PJ_CVC")
	pcPay_PJ_TestMode=request.Form("pcPay_PJ_TestMode")
	if pcPay_PJ_TestMode="1" then
		pcPay_PJ_TestMode=1
	else
		pcPay_PJ_TestMode=0
	end if

	pcPay_PJ_MerchantID=enDeCrypt(pcPay_PJ_MerchantID, scCrypPass)
	pcPay_PJ_MerchantPassword=enDeCrypt(pcPay_PJ_MerchantPassword, scCrypPass)
	
	query="UPDATE pcPay_PayJunction SET pcPay_PJ_MerchantID='"&pcPay_PJ_MerchantID&"',pcPay_PJ_MerchantPassword='"&pcPay_PJ_MerchantPassword&"',pcPay_PJ_TransType='"&pcPay_PJ_TransType&"',pcPay_PJ_cardTypes ='"&pcPay_PJ_cardTypes&"',pcPay_PJ_CVC=" & pcPay_PJ_CVC &",pcPay_PJ_TestMode=" & pcPay_PJ_TestMode & " WHERE pcPay_PJ_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=64"
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

Function gwPayJunction()
	varCheck=1
	'request gateway variables and insert them into the PayJunction table
	pcPay_PJ_MerchantID=request.Form("pcPay_PJ_MerchantID")
	pcPay_PJ_MerchantPassword=request.Form("pcPay_PJ_MerchantPassword")
	pcPay_PJ_TransType = request.Form("pcPay_PJ_TransType")
	pcPay_PJ_cardTypes=request.Form("pcPay_PJ_cardTypes")
	pcPay_PJ_CVC= request.Form("pcPay_PJ_CVC")
	pcPay_PJ_TestMode=request.Form("pcPay_PJ_TestMode")
	if pcPay_PJ_TestMode="1" then
		pcPay_PJ_TestMode=1
	else
		pcPay_PJ_TestMode=0
	end if

	If pcPay_PJ_MerchantID="" OR pcPay_PJ_MerchantPassword="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add PayJunction as your payment gateway. <b>""Merchant Login""</b> and <b>""Merchant Password""</b> are required fields.")
	End If
	'encrypt
	pcPay_PJ_MerchantID=enDeCrypt(pcPay_PJ_MerchantID, scCrypPass)
	pcPay_PJ_MerchantPassword=enDeCrypt(pcPay_PJ_MerchantPassword, scCrypPass)
	
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

	query="UPDATE pcPay_PayJunction SET pcPay_PJ_MerchantID='"&pcPay_PJ_MerchantID&"',pcPay_PJ_MerchantPassword='"&pcPay_PJ_MerchantPassword&"',pcPay_PJ_TransType='"&pcPay_PJ_TransType&"',pcPay_PJ_cardTypes ='"&pcPay_PJ_cardTypes&"',pcPay_PJ_CVC=" & pcPay_PJ_CVC &",pcPay_PJ_TestMode=" & pcPay_PJ_TestMode & " WHERE pcPay_PJ_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'PayJunction','gwPayJunction.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",64,'"&paymentNickName&"')"
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

if request("gwchoice")="64" then
	'Check if table exist in database, if not we will rediect to database update prompt
	err.clear
	call openDb()
	query="select * from pcPay_PayJunction"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		set rs=nothing
		call closedb()
		response.redirect "upddbPayJunction.asp"
	else
		set rs=nothing
		call closedb()
	end if
	intDoNotApply = 0
	
	'//Check for required components and that they are working 
	reDim strComponent(0)
	reDim strClass(0,3)
	
	'The component names
	strComponent(0) = "PayJunction"
	
	'The component class names
	strClass(0,0) = "WinHttp.WinHttpRequest.5"
	strClass(0,1) = "WinHttp.WinHttpRequest.5.1"
	strClass(0,3) = "1"	
	
	isComErr = Cint(0)
	strComErr = Cstr("")
	
	For i=0 to UBound(strComponent)
		strErr = IsObjInstalled(i)
		If strErr <> "" Then
			strComErr = strComErr & strErr
			isComErr = 1
		End If
	Next

	if request("mode")="Edit" then
	
	 	call opendb()	
		query="SELECT pcPay_PJ_MerchantID,pcPay_PJ_MerchantPassword,pcPay_PJ_TransType,pcPay_PJ_cardTypes, pcPay_PJ_CVC,pcPay_PJ_TestMode FROM pcPay_PayJunction Where pcPay_PJ_ID=1;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcPay_PJ_MerchantID=rs("pcPay_PJ_MerchantID")
		pcPay_PJ_MerchantID=enDeCrypt(pcPay_PJ_MerchantID, scCrypPass)
		pcPay_PJ_MerchantPassword=rs("pcPay_PJ_MerchantPassword")
		pcPay_PJ_MerchantPassword=enDeCrypt(pcPay_PJ_MerchantPassword, scCrypPass)
		pcPay_PJ_TransType = rs("pcPay_PJ_TransType")
		pcPay_PJ_cardTypes=rs("pcPay_PJ_cardTypes")
		pcPay_PJ_CVC=rs("pcPay_PJ_CVC")
		pcPay_PJ_TestMode=rs("pcPay_PJ_TestMode")
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=64"
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
	<input type="hidden" name="addGw" value="64">
<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/PayJunction_Logo.jpg" width="275" height="96"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
<table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
<tr>
    <td><strong><br>
    </strong>One stop shop. PayJunction will setup both a merchant account and gateway   service for you. 24/7 customer support for accounts at no additional charge. PayJunction is also a PCI Level 1 compliant provider. <strong><br>
    <br>
    <a href="#">Payjunction Website</a></strong><br />
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
                            <%
							if isComErr = 1 then
                                intDoNotApply = 1 %>
                                    <tr>
                                      <td colspan="2">&nbsp;</td>
                                    </tr>
                                    <tr>
                                      <td colspan="2"><img src="images/red_x.png" alt="Unable to add Paradata" width="12" height="12" /> <strong>Pay Junction cannot be activated</strong> - the required third party components are not installed and working on the web server.<br />                	    
                                      <br /></td>
                                    </tr>
                                    <tr>
                                        <td colspan="2">
                                            <center>
                                              <strong>Required components forPay Junction:</strong><br />
                                              <i><%= strComErr %></i><br /><br />
                                       <input type="button" value="Back" onclick="javascript:history.back()"></center> </td>
                                    </tr>
                            <% else %>
                    <% if request("mode")="Edit" then
                        dim pcPay_PJ_MerchantIDCnt,pcPay_PJ_MerchantIDEnd,pcPay_PJ_MerchantIDStart
                        pcPay_PJ_MerchantIDCnt=(len(pcPay_PJ_MerchantID)-2)
                        pcPay_PJ_MerchantIDEnd=right(pcPay_PJ_MerchantID,2)
                        pcPay_PJ_MerchantIDStart=""
                        for c=1 to pcPay_PJ_MerchantIDCnt
                            pcPay_PJ_MerchantIDStart=pcPay_PJ_MerchantIDStart&"*"
                        next
                        %>
                        <tr> 
                            <td height="31" colspan="2">Current Merchant Login:&nbsp;<%=pcPay_PJ_MerchantIDStart&pcPay_PJ_MerchantIDEnd%></td>
                        </tr>
                        <tr> 
                            <td colspan="2"> For security reasons, your &quot;Merchant Login&quot; 
                                is only partially shown on this page. If you need 
                                to edit your Login information, please re-enter 
                                your &quot;Merchant Login&quot; and &quot; Merchant Password&quot; below.</td>
                        </tr>
                    <% end if %>
                    <tr>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr>
                        <td width="111"> <div align="right">Merchant Login:</div></td>
                        <td width="360"> <input type="text" name="pcPay_PJ_MerchantID" size="20"></td>
                    </tr>
                    <tr>
                        <td width="111"> <div align="right">Merchant Password:</div></td>
                        <td width="360"> <input type="text" name="pcPay_PJ_MerchantPassword" size="20"></td>
                    </tr>
                      <tr> 
                        <td> <div align="right">Transaction Type:</div></td>
                        <td> <select name="pcPay_PJ_TransType">
                                <option value="AUTHORIZATION_CAPTURE" selected="selected">Authorization Capture</option>
                                <option value="AUTHORIZATION" <% if pcPay_PJ_TransType="AUTHORIZATION" then %>selected<% end if %>>Authorize Only</option>
                            </select> </td>
                    </tr>
                        <tr> 
                        <td> <div align="right">Accepted Cards:</div></td>
                        <td> <input name="pcPay_PJ_CardTypes" type="checkbox" class="clearBorder" value="VISA" <% if pcPay_PJ_CardTypes="VISA" or (instr(pcPay_PJ_CardTypes,"VISA,")>0) or (instr(pcPay_PJ_CardTypes,", VISA")>0) then%>checked<%end if%>>
                        Visa 
                        <input name="pcPay_PJ_CardTypes" type="checkbox" class="clearBorder" value="MAST" <% if pcPay_PJ_CardTypes="MAST" or (instr(pcPay_PJ_CardTypes,"MAST,")>0) or (instr(pcPay_PJ_CardTypes,", MAST")>0) then%>checked<%end if%>>
                        MasterCard 
                        <input name="pcPay_PJ_CardTypes" type="checkbox" class="clearBorder" value="AMER" <% if pcPay_PJ_CardTypes="AMER" or (instr(pcPay_PJ_CardTypes,"AMER,")>0) or (instr(pcPay_PJ_CardTypes,", AMER")>0) then%>checked<%end if%>>
                        American Express 
                        <input name="pcPay_PJ_CardTypes" type="checkbox" class="clearBorder" value="DISC" <% if pcPay_PJ_CardTypes="DISC" or (instr(pcPay_PJ_CardTypes,"DISC,")>0) or (instr(pcPay_PJ_CardTypes,", DISC")>0) then%>checked<%end if%>>
                        Discover 
                        <input name="pcPay_PJ_CardTypes" type="checkbox" class="clearBorder" value="DINE" <% if pcPay_PJ_CardTypes="DINE" or (instr(pcPay_PJ_CardTypes,"DINE,")>0) or (instr(pcPay_PJ_CardTypes,", DINE")>0) then%>checked<%end if%>>
                        Diner's Club 
                        <input name="pcPay_PJ_CardTypes" type="checkbox" class="clearBorder" value="JCB" <% if pcPay_PJ_CardTypes="JCB" or (instr(pcPay_PJ_CardTypes,"JCB")>0) or (instr(pcPay_ACH_CardTypes,", JCB")>0) then%>checked<%end if%>>
                        JCB
                        </td>
                    </tr>
                        <tr> 
                            <td> <div align="right">Require CVC:</div></td>
                            <td>
                            <input type="radio" class="clearBorder" name="pcPay_PJ_CVC" value="1" checked> Yes 
                            <input name="pcPay_PJ_CVC" type="radio" class="clearBorder" value="0" <%if pcPay_PJ_CVC=0 then%> Checked <%end if %> /> No
                            </td>
                        </tr>
                        <tr> 
                        <td> <div align="right"> 
                                <input name="pcPay_PJ_TestMode" type="checkbox" class="clearBorder" id="pcPay_PJ_TestMode" value="1" <% if pcPay_PJ_TestMode=1 then%>checked<% end if%>> 
                            </div><!--<input type="hidden" name="pcPay_PJ_PayPeriod" value="1">--></td>
                        <td><b>Enable Test Mode </b>(Credit cards will not be charged and cannot be viewed in the Merchant Interface.)</td>
                    </tr>
                            <% end if %>
                  <tr>
                    <td>&nbsp;</td>
                    <td class="pcSubmenuContent">&nbsp;</td>
                  </tr>
                </table>
            </div>
        </div>
                <% if intDoNotApply = 0 then %>
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
                <% end if %>
        <script type="text/javascript">
        <!--
                var CollapsiblePanel1 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel1", {contentIsOpen:true});
                <% if intDoNotApply = 0 then %>
                var CollapsiblePanel2 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel2", {contentIsOpen:false});;
                var CollapsiblePanel3 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel3", {contentIsOpen:false});
                var CollapsiblePanel4 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel4", {contentIsOpen:false});
                <% end if %>
                //-->
        </script>
    </td>
</tr>
</table>
<!-- New View End --><% end if %>

<%
'---Start CyberSource---
Function gwcysEdit()
	call opendb()
	query="SELECT pcPay_Cys_MerchantId, pcPay_Cys_TransType FROM pcPay_CyberSource WHERE pcPay_Cys_ID=1;"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	pcPay_Cys_MerchantId2=rstemp("pcPay_Cys_MerchantId")
	pcPay_Cys_MerchantId2=enDeCrypt(pcPay_Cys_MerchantId2, scCrypPass)
	pcPay_Cys_MerchantId=request.Form("pcPay_Cys_MerchantId")
	If pcPay_Cys_MerchantId="" then
		pcPay_Cys_MerchantId=pcPay_Cys_MerchantId2
	end if
	pcPay_Cys_TransType=request.Form("pcPay_Cys_TransType")
	pcPay_Cys_CardType=request.Form("pcPay_Cys_CardType")
	pcPay_Cys_CVV=request.Form("pcPay_Cys_CVVEnabled")
	pcPay_Cys_TestMode=request.Form("pcPay_Cys_TestMode")
	if pcPay_Cys_TestMode="YES" then
		pcPay_Cys_TestMode="0"
	else
		pcPay_Cys_TestMode="1"
	end if
	pcPay_Cys_eCheck=request.Form("pcPay_Cys_eCheck")
	pcPay_Cys_eCheckPending=request.Form("pcPay_Cys_eCheckPending")
	paymentNickName2=replace(request.Form("paymentNickName2"),"'","''")
	if paymentNickName2="" then
		paymentNickName2="Check"
	end if
	pcPay_Cys_MerchantId=enDeCrypt(pcPay_Cys_MerchantId, scCrypPass)
	
	query="UPDATE pcPay_CyberSource SET pcPay_Cys_MerchantId='"&pcPay_Cys_MerchantId&"',pcPay_Cys_TransType="&pcPay_Cys_TransType&",pcPay_Cys_CardType='"&pcPay_Cys_CardType&"',pcPay_Cys_CVV="&pcPay_Cys_CVV&", pcPay_Cys_TestMode="&pcPay_Cys_TestMode&", pcPay_Cys_eCheck=" & pcPay_Cys_eCheck &",pcPay_Cys_eCheckPending=" &pcPay_Cys_eCheckPending &" WHERE pcPay_Cys_ID=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ", priceToAdd="&priceToAdd&" ,percentageToAdd="&percentageToAdd&" ,paymentNickName='"&paymentNickName&"' WHERE gwCode=32"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	if pcPay_Cys_eCheck="1" then
		query="SELECT * FROM payTypes WHERE gwCode=62"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if rs.eof then
			query="INSERT INTO payTypes (paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES ('CyberSource eCheck','gwCyberSourceEcheck.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",62,'"&paymentNickName2&"')"
		else
			query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",paymentNickName='"&paymentNickName2&"' WHERE gwCode=62"
		end if
	else
		query="DELETE FROM payTypes WHERE gwCode=62"
	end if
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

Function gwcys()
	varCheck=1
	'request gateway variables and insert them into the pcPay_CyberSource table
	pcPay_Cys_MerchantId=request.Form("pcPay_Cys_MerchantId")
	pcPay_Cys_TransType=request.Form("pcPay_Cys_TransType")
	pcPay_Cys_CardType=request.Form("pcPay_Cys_CardType")
	pcPay_Cys_CVV=request.Form("pcPay_Cys_CVVEnabled")
	pcPay_Cys_TestMode=request.Form("pcPay_Cys_TestMode")
	pcPay_Cys_eCheck=request.Form("pcPay_Cys_eCheck")
	pcPay_Cys_eCheckPending=request.Form("pcPay_Cys_eCheckPending")
	if pcPay_Cys_TestMode="YES" then
		pcPay_Cys_TestMode="0"
	else
		pcPay_Cys_TestMode="1"
	end if
	If pcPay_Cys_MerchantId="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add CyberSource as your payment gateway. <b>""Merchant ID""</b> is a required field.")
	End If
	If pcPay_Cys_CardType="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add CyberSource as your payment gateway. Please select atleast one Card type.")
	End If
	'encrypt
	pcPay_Cys_MerchantId=enDeCrypt(pcPay_Cys_MerchantId, scCrypPass)
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
	
	paymentNickName2=replace(request.Form("paymentNickName2"),"'","''")
	if paymentNickName2="" then
		paymentNickName2="Check"
	end if
	
	err.clear
	err.number=0
	call openDb() 

	query="UPDATE pcPay_CyberSource SET pcPay_Cys_MerchantId='"&pcPay_Cys_MerchantId&"',pcPay_Cys_TransType="&pcPay_Cys_TransType&",pcPay_Cys_CardType='"&pcPay_Cys_CardType&"',pcPay_Cys_CVV="&pcPay_Cys_CVV&", pcPay_Cys_TestMode="&pcPay_Cys_TestMode&", pcPay_Cys_eCheck=" & pcPay_Cys_eCheck &",pcPay_Cys_eCheckPending=" &pcPay_Cys_eCheckPending &" WHERE pcPay_Cys_ID=1"
	set rs=Server.CreateObject("ADODB.Recordset")  
	set rs=connTemp.execute(query)
	if err.number<>0 then
		response.write err.description
		response.End()
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'CyberSource','gwCyberSource.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",32,'"&paymentNickName&"')"
	set rs=Server.CreateObject("ADODB.Recordset")  
	set rs=connTemp.execute(query)
	if err.number<>0 then
		response.write err.description
		response.End()
	end if
	if pcPay_Cys_eCheck="1" then
		query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'CyberSource eCheck','gwCyberSourceECheck.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",62,'"&paymentNickName2&"')"
		set rs=Server.CreateObject("ADODB.Recordset")  
		set rs=connTemp.execute(query)		

		if err.number<>0 then
			response.write err.description
			response.End()
		end if
	 End if 
	set rs=nothing
    call closedb()
end function
%>

<% 
'// 62 Check
'// 32
if request("gwchoice")="32" then
	tmp_id=request("id")
	tmp_mode=request("mode")

	'Check to see if fields exists in table, if not, add
	err.clear
	call openDb()
	query="SELECT pcPay_Cys_eCheck FROM pcPay_CyberSource"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		set rs=nothing
		call closedb()
		response.redirect "upddbCyberSource.asp?mode="&tmp_mode&"&id="&tmp_id
	else
		set rs=nothing
		call closedb()
	end if
	
	intDoNotApply = 0

	'//Check for required components and that they are working 
	reDim strComponent(0)
	reDim strClass(0,2)

	'The component names
	strComponent(0) = "CyberSource"
	
	'The component class names
	strClass(0,0) = "CyberSourceWS.MerchantConfig"
	strClass(0,1) = "CyberSourceWS.Hashtable"

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
		call opendb
		query="Select pcPay_Cys_MerchantId,pcPay_Cys_TransType,pcPay_Cys_CardType,pcPay_Cys_CVV, pcPay_Cys_TestMode, pcPay_Cys_eCheck, pcPay_Cys_eCheckPending from pcPay_CyberSource where pcPay_Cys_ID=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcPay_Cys_MerchantId=rs("pcPay_Cys_MerchantId")
		pcPay_Cys_MerchantId=enDeCrypt(pcPay_Cys_MerchantId, scCrypPass)
		pcPay_Cys_TransType=rs("pcPay_Cys_TransType")
		pcPay_Cys_CardType=rs("pcPay_Cys_CardType")
		pcPay_Cys_CVV=rs("pcPay_Cys_CVV")
		pcPay_Cys_TestMode=rs("pcPay_Cys_TestMode")
		pcPay_Cys_eCheck=rs("pcPay_Cys_eCheck")
		pcPay_Cys_eCheckPending=rs("pcPay_Cys_eCheckPending")
		if pcPay_Cys_TestMode="0" then
			pcPay_Cys_TestMode="YES"
		end if
				
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=32"
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
		
		query= "SELECT paymentNickName FROM payTypes WHERE gwCode=62"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=conntemp.execute(query)
		if rs.eof then
			paymentNickName2="Check"
		else
			paymentNickName2=rs("paymentNickName")
		end if
		set rs=nothing
		call closedb()
		%>
		<input type="hidden" name="mode" value="Edit">
    <% end if %>
	<input type="hidden" name="addGw" value="32">

    <!-- New View Start -->
    <table width="100%">
        <tr>
            <td align="left" style="font-size:15px;"><img src="gateways/logos/CyberSource_logo.JPG" width="177" height="52"></td>
            <td align="left" style="font-size:15px;">&nbsp;</td>
        </tr>
    </table>
    <br>
    <table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
        <tr>
            <td><strong>CyberSource - Simple Order API integration<br>
            <br>
            </strong>With CyberSource you can sell online almost anywhere in the world, instantly boosting your customer reach. You can accept the payment types preferred in local markets, transact payments in over 190 countries and deposit funds in major traded currencies, all through a single connection.<strong><br>
            <br>
            <a href="http://www.cybersource.com/support_center/implementation/downloads/simple_order/matrix/">CyberSource Website</a></strong><br />
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
                                  <td colspan="2"><img src="images/red_x.png" alt="Unable to add Paradata" width="12" height="12" /> <strong>CyberSource cannot be activated</strong> - the required third party components are not installed and working on the web server.<br />                	    
                                  <br /></td>
                                </tr>
                                <tr>
                                    <td colspan="2">
                                        <center>
                                          <strong>Required components for CyberSource:</strong><br /><i><%= strComErr %></i><br /><br />
                                    <input type="button" value="Back" onclick="javascript:history.back()"></center></td>
                                </tr>
                            <% else %>
								<% if request("mode")="Edit" then %>
                                    <% dim pcPay_Cys_MerchantIdCnt,pcPay_Cys_MerchantIdEnd,pcPay_Cys_MerchantIdStart
                                    pcPay_Cys_MerchantIdCnt=(len(pcPay_Cys_MerchantId)-2)
                                    pcPay_Cys_MerchantIdEnd=right(pcPay_Cys_MerchantId,2)
                                    pcPay_Cys_MerchantIdStart=""
                                    for c=1 to pcPay_Cys_MerchantIdCnt
                                        pcPay_Cys_MerchantIdStart=pcPay_Cys_MerchantIdStart&"*"
                                    next %>
                                    <tr> 
                                        <td height="31" colspan="2">Current CyberSource Merchant Number:&nbsp;<%=pcPay_Cys_MerchantIdStart&pcPay_Cys_MerchantIdEnd%></td>
                                    </tr>
                                    <tr> 
                                        <td colspan="2"> For security reasons, your &quot;CyberSource 
                                            Merchant ID&quot; is only partially shown on this 
                                            page. If you need to edit your account information, 
                                            please re-enter your &quot;Merchant ID&quot; below.</td>
                                    </tr>
                                <% end if %>
                                <tr>
                                  <td>&nbsp;</td>
                                  <td>&nbsp;</td>
                                </tr>
                                <tr> 
                                    <td> <div align="right">Merchant ID:</div></td>
                                    <td> <input type="text" value="" name="pcPay_Cys_MerchantId" size="24"></td>
                                </tr>
                                <tr> 
                                    <td valign="top"> <div align="right">Keys Directory:</div></td>
                                    <td><b><%=scPcFolder&"/"&scAdminFolderName&"/"%></b> folder <br>
                                        <font size="1"><i>(Copy your sercurity keys file *.p12 
                                        to this folder before active this payment gateway)</i></font></td>
                                </tr>
                                <tr> 
                                    <td nowrap> <div align="right">Transaction Type:</div></td>
                                    <td> <select name="pcPay_Cys_TransType">
                                            <option value="2" selected>Sale</option>
                                            <option value="0" <% if pcPay_Cys_TransType=0 then%>selected<% end if %>>Authorize Only</option>
                                        </select></td>
                                </tr>
                                <tr> 
                                    <td> <div align="right">Currency:</div></td>
                                    <td>USD only </td>
                                </tr>
                                <tr> 
                                    <td> <div align="right">Accepted Cards:</div></td>
                                    <td>
                                    <input name="pcPay_Cys_CardType" type="checkbox" class="clearBorder" id="cys_visa" value="V" <% if pcPay_Cys_CardType="V" or (instr(pcPay_Cys_CardType,"V,")>0) or (instr(pcPay_Cys_CardType,", V")>0) then%>checked<%end if%>>
                                    Visa 
                                    <input name="pcPay_Cys_CardType" type="checkbox" class="clearBorder" id="cys_mc" value="M" <% if pcPay_Cys_CardType="M" or (instr(pcPay_Cys_CardType,"M,")>0) or (instr(pcPay_Cys_CardType,", M")>0) then%>checked<%end if%>>
                                    MasterCard 
                                    <input name="pcPay_Cys_CardType" type="checkbox" class="clearBorder" id="cys_amex" value="A" <% if pcPay_Cys_CardType="A" or (instr(pcPay_Cys_CardType,"A,")>0) or (instr(pcPay_Cys_CardType,", A")>0) then%>checked<%end if%>>
                                    American Express 
                                    <input name="pcPay_Cys_CardType" type="checkbox" class="clearBorder" id="cys_disc" value="D" <% if pcPay_Cys_CardType="D" or (instr(pcPay_Cys_CardType,"D,")>0) or (instr(pcPay_Cys_CardType,", D")>0) then%>checked<%end if%>>
                                    Discover 
                                    <input name="pcPay_Cys_CardType" type="checkbox" class="clearBorder" id="cys_diner" value="E" <% if pcPay_Cys_CardType="E" or (instr(pcPay_Cys_CardType,"E,")>0) or (instr(pcPay_Cys_CardType,", E")>0) then%>checked<%end if%>>
                                    Diners</td>
                                </tr>
                                <tr> 
                                    <td> <div align="right">Require CVV:</div></td>
                                    <td> <input type="radio" class="clearBorder" name="pcPay_Cys_CVVEnabled" value="1" checked>
                                        Yes 
                                        <input type="radio" class="clearBorder" name="pcPay_Cys_CVVEnabled" value="0" <% if pcPay_Cys_CVV=0 then%>checked<% end if %>>
                                        No</td>
                                </tr>
                                 <tr> 
                                    <td> <div align="right">Accept Checks:</div></td>
                                    <td> <input type="radio" class="clearBorder" name="pcPay_Cys_eCheck" value="1" checked>
                                        Yes 
                                        <input name="pcPay_Cys_eCheck" type="radio" class="clearBorder" value="0" <% if pcPay_Cys_eCheck=0 then%>checked<% end if %>>
                                        No</td>
                                </tr>
                                 <tr> 
                                    <td colspan="2">All eCheck orders are automatically considered 
                                        &quot;Processed&quot; by ProductCart, as the order amount 
                                        is always debited to the customer's bank account. If 
                                        for any reasons you would like eCheck orders to be considered 
                                        &quot;Pending&quot;, use this option. Should eCheck 
                                        orders be considered &quot;Pending&quot;? 
                                        <input type="radio" class="clearBorder" name="pcPay_Cys_eCheckPending" value="1" checked> Yes 
                                        <input name="pcPay_Cys_eCheckPending" type="radio" class="clearBorder" value="0" <% if pcPay_Cys_eCheckPending=0 then%>checked<% end if %>> No</td>
                                </tr>
                                <tr> 
                                    <td><div align="right"> 
                                            <input name="pcPay_Cys_TestMode" type="checkbox" class="clearBorder" value="YES" <% if pcPay_Cys_TestMode="YES" then%>checked<% end if%>> 
                                        </div></td>
                                    <td><b>Enable Test Mode </b>(Credit cards will not be charged)</td>
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
                        		<td width="10%" nowrap="nowrap"><div align="left">eCheck&nbsp;&nbsp;Payment Name:&nbsp;</div></td>
                        		<td><input name="paymentNickName2" value="<%=paymentNickName2%>" size="35" maxlength="255"></td>
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

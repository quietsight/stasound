<%
'---Start Virtual Merchant---
Function gwVMPayEdit()
	call opendb()
	'request gateway variables and insert them into the VirtualMerchant table
	query= "SELECT TOP 1 pcPay_VM_MerchantID,pcPay_VM_UserID,pcPay_VM_Pin,pcPay_VM_TransType,pcPay_VM_TestMode,pcPay_VM_CVV2 FROM pcPay_VirtualMerchant;"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		call closedb()
	  response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddPaymentOpt: "&Err.Description) 
	end If
	pcPay_VM_MerchantID2=rstemp("pcPay_VM_MerchantID")
	pcPay_VM_MerchantID=request.Form("pcPay_VM_MerchantID")
	if pcPay_VM_MerchantID="" then
		pcPay_VM_MerchantID=pcPay_VM_MerchantID2
	end if
	
	pcPay_VM_UserID2=rstemp("pcPay_VM_UserID")
	pcPay_VM_UserID=request.Form("pcPay_VM_UserID")
	if pcPay_VM_UserID="" then
		pcPay_VM_UserID=pcPay_VM_UserID2
	end if
	
	pcPay_VM_PIN2=rstemp("pcPay_VM_PIN")
	pcPay_VM_PIN=request.Form("pcPay_VM_PIN")
	if pcPay_VM_PIN="" then
		pcPay_VM_PIN=pcPay_VM_PIN2
	else
		pcPay_VM_PIN=enDeCrypt(pcPay_VM_PIN, scCrypPass)
	end if
	
	pcPay_VM_TestMode=request.Form("pcPay_VM_TestMode")
	if pcPay_VM_TestMode="" then
		pcPay_VM_TestMode=0
	else
		pcPay_VM_TestMode=1
	end if
	
	pcPay_VM_TransType2=rstemp("pcPay_VM_TransType")
	pcPay_VM_TransType=request.Form("pcPay_VM_TransType")
	If pcPay_VM_TransType="" then
		pcPay_VM_TransType=pcPay_VM_TransType2
	end if
	
	pcPay_VM_CVV2=request.Form("pcPay_VM_CVV2")

	query="UPDATE pcPay_VirtualMerchant SET pcPay_VM_MerchantID='"&pcPay_VM_MerchantID&"', pcPay_VM_UserID='"&pcPay_VM_UserID&"', pcPay_VM_PIN='"&pcPay_VM_PIN&"', pcPay_VM_TransType='"&pcPay_VM_TransType&"', pcPay_VM_TestMode="&pcPay_VM_TestMode&", pcPay_VM_CVV2="&pcPay_VM_CVV2&";"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=56"

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

Function gwVMPay()
	varCheck=1
	'request gateway variables and insert them into the SecPay table
	pcPay_VM_MerchantID=request.Form("pcPay_VM_MerchantID")
	pcPay_VM_UserID=request.Form("pcPay_VM_UserID")
	pcPay_VM_PIN=request.Form("pcPay_VM_PIN")
	pcPay_VM_TransType=request.Form("pcPay_VM_TransType")
	pcPay_VM_CVV2=request.Form("pcPay_VM_CVV2")
	pcPay_VM_TestMode=request.Form("pcPay_VM_TestMode")
	if pcPay_VM_TestMode="" then
		pcPay_VM_TestMode=0
	else
		pcPay_VM_TestMode=1
	end if
	
	If pcPay_VM_MerchantID="" then
		call closedb()
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add VirtualMerchant as your payment gateway. <b>""Merchant ID""</b> is a required field.")
	End If
	
	If pcPay_VM_UserID="" then
		call closedb()
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add VirtualMerchant as your payment gateway. <b>""User ID""</b> is a required field.")
	End If
	
	If pcPay_VM_PIN="" then
		call closedb()
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add VirtualMerchant as your payment gateway. <b>""PIN""</b> is a required field.")
	End If
	
	pcPay_VM_PIN=enDeCrypt(pcPay_VM_PIN, scCrypPass)
	
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

				
	query="UPDATE pcPay_VirtualMerchant SET pcPay_VM_MerchantID='"&pcPay_VM_MerchantID&"',pcPay_VM_UserID='"&pcPay_VM_UserID&"',pcPay_VM_PIN='"&pcPay_VM_PIN&"',pcPay_VM_TransType='"&pcPay_VM_TransType&"',pcPay_VM_TestMode="&pcPay_VM_TestMode&",pcPay_VM_CVV2="&pcPay_VM_CVV2 & " WHERE pcPay_VM_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")  
	set rs=connTemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'VirtualMerchant','gwVMPay.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",56,'"&paymentNickName&"')"
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

if request("gwchoice")="56" then
	'Check if table exist in database, if not we will rediect to database update prompt
	err.clear
	call openDb()
	query="select * from pcPay_VirtualMerchant"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		set rs=nothing
		call closedb()
		response.redirect "upddbVirtualMerchant.asp"
	else
		set rs=nothing
		call closedb()
	end if
	if request("mode")="Edit" then
	 	call opendb()
		query= "SELECT pcPay_VM_MerchantID, pcPay_VM_UserID, pcPay_VM_Pin, pcPay_VM_TransType, pcPay_VM_TestMode, pcPay_VM_CVV2 FROM pcPay_VirtualMerchant Where pcPay_VM_ID=1;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcPay_VM_MerchantID=rs("pcPay_VM_MerchantID")
		pcPay_VM_UserID=rs("pcPay_VM_UserID")
		pcPay_VM_Pin=rs("pcPay_VM_Pin")
		pcPay_VM_TransType=rs("pcPay_VM_TransType")
		pcPay_VM_TestMode=rs("pcPay_VM_TestMode")
		pcPay_VM_CVV2=rs("pcPay_VM_CVV2")
				
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=56"
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
	<input type="hidden" name="addGw" value="56">
<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/virtualmerchant_logo.jpg" width="500" height="54"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
<table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
<tr>
    <td><strong><br>
    </strong>Elavon VirtualMerchant (formerly known as Nova ViaKlix) is a complete, hosted payment solution for  e-commerce   transactions<strong><br>
    <br>
    <a href="https://www.myvirtualmerchant.com/" target="_blank">VirtualMerchant Website</a></strong><br />
<br />
</td>
</tr>
<tr>
    <td>
        <div id="CollapsiblePanel1" class="CollapsiblePanel">
            <div class="CollapsiblePanelTab1">
                <table width="100%">
                  <tr>
                    <td width="580" class="pcPanelTitle1">Step 1: Configure Account - Virtual Merchant</td>
                  </tr>
                </table>
            </div>
            <div class="CollapsiblePanelContent">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <% if request("mode")="Edit" then
					pcPay_VM_MCCnt=(len(pcPay_VM_MerchantID)-2)
					pcPay_VM_MCEnd=right(pcPay_VM_MerchantID,2)
					pcPay_VM_MCStart=""
					for c=1 to pcPay_VM_MCCnt
						pcPay_VM_MCStart=pcPay_VM_MCStart&"*"
					next
					pcPay_VM_USCnt=(len(pcPay_VM_UserID)-2)
					pcPay_VM_USEnd=right(pcPay_VM_UserID,2)
					pcPay_VM_USStart=""
					for c=1 to pcPay_VM_USCnt
						pcPay_VM_USStart=pcPay_VM_USStart&"*"
					next
					pcPay_VM_PICnt=(len(pcPay_VM_Pin))
					pcPay_VM_PIEnd=""
					pcPay_VM_PIStart=""
					for c=1 to pcPay_VM_PICnt
						pcPay_VM_PIStart=pcPay_VM_PIStart&"*"
					next %>
                    <tr> 
                        <td height="31" colspan="2">Current Merchant ID:&nbsp;<%=pcPay_VM_MCStart&pcPay_VM_MCEnd%></td>
                    </tr>
                    <tr> 
                        <td height="31" colspan="2">Current User ID:&nbsp;<%=pcPay_VM_USStart&pcPay_VM_USEnd%></td>
                    </tr>
                    <tr> 
                        <td height="31" colspan="2">Current PIN:&nbsp;<%=pcPay_VM_PIStart&pcPay_VM_PIEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your VirtualMerchant &quot;Merchant ID, User ID, PIN&quot; are only partially shown on this page. If you need to edit your account information, please re-enter your information below.</td>
                    </tr>
				<% end if %>
                <tr>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                    <td width="121"> <div align="right">Merchant ID:</div></td>
                    <td width="555"> <input type="text" name="pcPay_VM_MerchantID" size="30"></td>
                </tr>
                <tr> 
                    <td> <div align="right">User ID:</div></td>
                    <td> <input type="text" name="pcPay_VM_UserID" size="30"></td>
                </tr>
                <tr> 
                    <td> <div align="right">PIN:</div></td>
                    <td> <input type="text" name="pcPay_VM_PIN" size="30"></td>
                </tr>
                <tr> 
                    <td> <div align="right">Transaction Type:</div></td>
                    <td> <select name="pcPay_VM_TransType">
                            <option value="CCSALE" selected>Sale</option>
                            <option value="CCAUTHONLY" <% if pcPay_VM_TransType="CCAUTHONLY" then%>selected<% end if %>>Authorize Only</option>
                        </select></td>
                </tr>
                <tr>
                    <td><div align="right">Require CVV2:</div></td>
                    <td><input type="radio" class="clearBorder" name="pcPay_VM_CVV2" value="1" checked>
                    Yes
                    <input name="pcPay_VM_CVV2" type="radio" class="clearBorder" value="0" <% if pcPay_VM_CVV2=0 then%>checked<% end if %>>
                    No</td>
                </tr>
                <tr> 
                    <td> <div align="right"> 
                            <input name="pcPay_VM_TestMode" type="checkbox" class="clearBorder" value="1" <% if pcPay_VM_TestMode=1 then%>checked<% end if %>>
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
<%
'--- Start PxPay ---
Function gwPxPayEdit()
	call opendb()
	'request gateway variables and insert them into the PxPay table
	query="SELECT pcPay_PxPay.pcPay_PxPay_PxPayUserId, pcPay_PxPay.pcPay_PxPay_PxPayTestUserId, pcPay_PxPay.pcPay_PxPay_PxPayKey, pcPay_PxPay.pcPay_PxPay_TxnType, pcPay_PxPay.pcPay_PxPay_TestMode, pcPay_PxPay.pcPay_PxPay_CurrencyInput FROM pcPay_PxPay WHERE (((pcPay_PxPay.pcPay_PxPay_ID)=1));"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	pcPay_PxPay_PxPayUserId2=rs("pcPay_PxPay_PxPayUserId")
	pcPay_PxPay_PxPayKey2=rs("pcPay_PxPay_PxPayKey")
	
	pcPay_PxPay_PxPayUserId=request.Form("pcPay_PxPay_PxPayUserId")
	if pcPay_PxPay_PxPayUserId="" then
		pcPay_PxPay_PxPayUserId=pcPay_PxPay_PxPayUserId2
	end if
	pcPay_PxPay_PxPayKey=request.Form("pcPay_PxPay_PxPayKey")
	if pcPay_PxPay_PxPayKey="" then
		pcPay_PxPay_PxPayKey=pcPay_PxPay_PxPayKey2
	end if
	set rs=nothing

	pcPay_PxPay_TxnType=request.Form("pcPay_PxPay_TxnType")
	pcPay_PxPay_PxPayTestUserId=request.Form("pcPay_PxPay_PxPayTestUserId")
	pcPay_PxPay_TestMode=request.Form("pcPay_PxPay_TestMode")
	if pcPay_PxPay_TestMode="" then
		pcPay_PxPay_TestMode=0
	else
		pcPay_PxPay_TestMode=1
		if pcPay_PxPay_PxPayTestUserId="" then
			response.redirect "techErr.asp?error="&Server.URLEncode("An error occured while trying to modify PaymentExpress settings. <b>""Test Username""</b> is a required field when activating Test Mode.")
		End If
	end if
	pcPay_PxPay_CurrencyInput=request.Form("pcPay_PxPay_CurrencyInput")
 
 	query="UPDATE pcPay_PxPay SET pcPay_PxPay_TxnType='"&pcPay_PxPay_TxnType&"', pcPay_PxPay_PxPayTestUserId='"&pcPay_PxPay_PxPayTestUserId&"', pcPay_PxPay_PxPayUserId='"&pcPay_PxPay_PxPayUserId&"', pcPay_PxPay_PxPayKey='"&pcPay_PxPay_PxPayKey&"',  pcPay_PxPay_TestMode="&pcPay_PxPay_TestMode&", pcPay_PxPay_CurrencyInput='"&pcPay_PxPay_CurrencyInput&"' WHERE pcPay_PxPay_ID=1"

	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If

	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=12"
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

Function gwPxPay()
  
	varCheck=1
	' Test Username
	pcPay_PxPay_PxPayTestUserId = request.Form("pcPay_PxPay_PxPayTestUserId")
	' Trans Type
	pcPay_PxPay_TxnType=request.Form("pcPay_PxPay_TxnType")
	' Trans Mode
	pcPay_PxPay_CurrencyInput=request.Form("pcPay_PxPay_CurrencyInput")
	pcPay_PxPay_TestMode=request.Form("pcPay_PxPay_TestMode")
	if pcPay_PxPay_TestMode="" then
		pcPay_PxPay_TestMode=0
	else
		pcPay_PxPay_TestMode=1
		if pcPay_PxPay_PxPayTestUserId="" then
			response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add PxPay as your payment gateway. <b>""Test Username""</b> is a required field when activating Test Mode.")
		End If
	end if
	' Trans Password
	pcPay_PxPay_PxPayKey=request.Form("pcPay_PxPay_PxPayKey")
	if pcPay_PxPay_PxPayKey="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add PxPay as your payment gateway. <b>""Key""</b> is a required field.")
	End If
	' Trans Username
	pcPay_PxPay_PxPayUserId=request.Form("pcPay_PxPay_PxPayUserId")
	if pcPay_PxPay_PxPayUserId="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add PxPay as your payment gateway. <b>""Username""</b> is a required field.")
	End If
  

	err.clear
	err.number=0	
	call openDb() 
	query= "UPDATE pcPay_PxPay SET  pcPay_PxPay_PxPayUserId='"&pcPay_PxPay_PxPayUserId&"', pcPay_PxPay_PxPayTestUserId='"&pcPay_PxPay_PxPayTestUserId&"', pcPay_PxPay_PxPayKey='"&pcPay_PxPay_PxPayKey&"', pcPay_PxPay_TxnType='"&pcPay_PxPay_TxnType&"', pcPay_PxPay_TestMode="&pcPay_PxPay_TestMode&", pcPay_PxPay_CurrencyInput='"&pcPay_PxPay_CurrencyInput&"'; "
	 
	set rs=Server.CreateObject("ADODB.Recordset")  
	set rs=connTemp.execute(query)
	if err.number <> 0 then
	     Response.write err.description &"<BR>"
		set rs=nothing
		call closedb()
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
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ", 'PxPay','gwPxPay.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",12,'"&paymentNickName&"')"

	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)

	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end if

	set rs=nothing
    call closedb()
end function

if request("gwchoice")="12" then
err.clear
	call openDb()
	query="select * from pcPay_PXPay"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		set rs=nothing
		call closedb()
		response.redirect "upddbPXPay.asp"
	else
		set rs=nothing
		call closedb()
	end if
	if request("mode")="Edit" then
		call opendb()
		query="SELECT pcPay_PxPay.pcPay_PxPay_PxPayUserId, pcPay_PxPay.pcPay_PxPay_PxPayTestUserId, pcPay_PxPay.pcPay_PxPay_PxPayKey, pcPay_PxPay.pcPay_PxPay_TxnType, pcPay_PxPay.pcPay_PxPay_TestMode, pcPay_PxPay.pcPay_PxPay_CurrencyInput FROM pcPay_PxPay WHERE (((pcPay_PxPay.pcPay_PxPay_ID)=1));"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcPay_PxPay_PxPayUserId=rs("pcPay_PxPay_PxPayUserId")
		pcPay_PxPay_PxPayTestUserId=rs("pcPay_PxPay_PxPayTestUserId")
		pcPay_PxPay_PxPayKey=rs("pcPay_PxPay_PxPayKey")
		pcPay_PxPay_TxnType=rs("pcPay_PxPay_TxnType")
		pcPay_PxPay_CurrencyInput=rs("pcPay_PxPay_CurrencyInput")
		pcPay_PxPay_TestMode=rs("pcPay_PxPay_TestMode")
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=12"
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
	<input type="hidden" name="addGw" value="12">
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
    </strong>DPS hosts and manages the payment page in its PCI compliant data center and a SSl Certificate is not required by the merchant.<strong><br>
    <br>
    <a href="http://www.paymentexpress.com/Products/Ecommerce/DPS_Hosted" target="_blank">Payment Express Website</a></strong><br />
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
					<% dim pcPay_PxPay_PxPayUserIdCnt,pcPay_PxPay_PxPayUserIdEnd,pcPay_PxPay_PxPayUserIdStart
                    pcPay_PxPay_PxPayUserIdCnt=(len(pcPay_PxPay_PxPayUserId)-2)
                    pcPay_PxPay_PxPayUserIdEnd=right(pcPay_PxPay_PxPayUserId,2)
                    pcPay_PxPay_PxPayUserIdStart=""
                    for c=1 to pcPay_PxPay_PxPayUserIdCnt
                        pcPay_PxPay_PxPayUserIdStart=pcPay_PxPay_PxPayUserIdStart&"*"
                    next
                    %>
                    
                  <tr class="normal">
                    <td height="31" colspan="2">Account UserId:&nbsp;<%=pcPay_PxPay_PxPayUserIdStart&pcPay_PxPay_PxPayUserIdEnd%></td>
                  </tr>
                  <tr class="normal">
                    <td colspan="2"> For security reasons, your &quot;Account UserId&quot; 
                      is only partially shown on this page. If you need 
                      to edit your account information, please re-enter 
                      your &quot;Account UserId&quot; and &quot;Account 
                      Character Key&quot; below.</td>
                  </tr>
                <% end if %>
                <tr class="normal">
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr class="normal"> 
                    <td width="24%"><div align="right">Account Username:</div></td>
                    <td width="76%"> <div align="left"> 
                    <input type="text" value="" name="pcPay_PxPay_PxPayUserId" size="30">
                  </div></td>
                </tr>
                <tr class="normal"> 
                    <td width="24%"><div align="right">Test Mode Username:</div></td>
                    <td width="76%"> <div align="left"> 
                    <input type="text" value="<%=pcPay_PxPay_PxPayTestUserId%>" name="pcPay_PxPay_PxPayTestUserId" size="30">
                    <font color="#FF0000">&nbsp;&nbsp;*Required if you are activating Test Mode. </font> </div></td>
                </tr>
                <tr class="normal"> 
                    <td width="24%"><div align="right">PX Post Key:</div></td>
                    <td width="76%"> <div align="left"> 
                    <input type="text" value="<%=pcPay_PxPay_PxPayKey%>" name="pcPay_PxPay_PxPayKey" size="30">
                    </div></td>
                </tr>
                <tr class="normal"> 
                    <td> <div align="right">Transaction Type:</div></td>
                    <td> <select name="pcPay_PxPay_TxnType">
                    <option value="Purchase" selected>Purchase</option>
                    <option value="Auth" <%if pcPay_PxPay_TxnType="Auth" then%>selected<%end if%>>Authorization Only</option>
                    </select> </td>
                </tr>
                <tr class="normal"> 
                    <td> <div align="right">Currency Type:</div></td>
                    <td>
                    	<select name="pcPay_PxPay_CurrencyInput">
                            <option value="CAD" <%if pcPay_PxPay_CurrencyInput="CAD" then%>selected<%end if%>>Canadian Dollar</option>
                            <option value="CHF" <%if pcPay_PxPay_CurrencyInput="CHF" then%>selected<%end if%>>Swiss Franc</option>
                            <option value="EUR" <%if pcPay_PxPay_CurrencyInput="EUR" then%>selected<%end if%>>Euro</option>
                            <option value="FRF" <%if pcPay_PxPay_CurrencyInput="FRF" then%>selected<%end if%>>French Franc</option>
                            <option value="GBP" <%if pcPay_PxPay_CurrencyInput="GBP" then%>selected<%end if%>>United Kingdom Pound</option>
                            <option value="HKD" <%if pcPay_PxPay_CurrencyInput="HKD" then%>selected<%end if%>>Hong Kong Dollar</option>
                            <option value="JPY" <%if pcPay_PxPay_CurrencyInput="JPY" then%>selected<%end if%>>Japanese Yen</option>
                            <option value="NZD" <%if pcPay_PxPay_CurrencyInput="NZD" then%>selected<%end if%>>New Zealand Dollar</option>
                            <option value="SGD" <%if pcPay_PxPay_CurrencyInput="SGD" then%>selected<%end if%>>Singapore Dollar</option>
                            <option value="USD" <%if pcPay_PxPay_CurrencyInput="USD" then%>selected<%end if%>>United States Dollar</option>
                            <option value="ZAR" <%if pcPay_PxPay_CurrencyInput="ZAR" then%>selected<%end if%>>Rand</option>
                            <option value="AUD" <%if pcPay_PxPay_CurrencyInput="AUD" then%>selected<%end if%>>Australian Dollar</option>
                            <option value="WST" <%if pcPay_PxPay_CurrencyInput="WST" then%>selected<%end if%>>Samoan Tala</option>
                            <option value="VUV" <%if pcPay_PxPay_CurrencyInput="VUV" then%>selected<%end if%>>Vanuatu Vatu</option>
                            <option value="TOP" <%if pcPay_PxPay_CurrencyInput="TOP" then%>selected<%end if%>>Tongan Pa'anga</option>
                            <option value="SBD" <%if pcPay_PxPay_CurrencyInput="SBD" then%>selected<%end if%>>Solomon Islands Dollar</option>
                            <option value="PNG" <%if pcPay_PxPay_CurrencyInput="PNG" then%>selected<%end if%>>Papua New Guinea Kina</option>
                            <option value="MYR" <%if pcPay_PxPay_CurrencyInput="MYR" then%>selected<%end if%>>Malaysian Ringgit</option>
                            <option value="KWD" <%if pcPay_PxPay_CurrencyInput="KWD" then%>selected<%end if%>>Kuwaiti Dinar</option>
                            <option value="FJD" <%if pcPay_PxPay_CurrencyInput="FJD" then%>selected<%end if%>>Fiji Dollar</option>
                		</select>
                	</td>
                </tr>
                <tr class="normal"> 
                    <td><div align="right"> 
                    <input name="pcPay_PxPay_TestMode" type="checkbox" id="pcPay_PxPay_TestMode" value="1" <% if pcPay_PxPay_TestMode=1 then %>checked<% end if %> />
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

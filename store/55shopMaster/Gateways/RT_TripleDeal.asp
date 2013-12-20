<%
'---Start TripleDeal---
Function gwTripleDealEdit()
	call opendb()
	'request gateway variables and insert them into the TripleDeal table
	query="SELECT pcPay_TD_MerchantName,pcPay_TD_MerchantPassword FROM pcPay_TripleDeal Where pcPay_TD_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	pcPay_TD_MerchantName2=rs("pcPay_TD_MerchantName")
	pcPay_TD_MerchantName2=enDeCrypt(pcPay_TD_MerchantName2, scCrypPass)

	pcPay_TD_MerchantName=request.Form("pcPay_TD_MerchantName")
	if pcPay_TD_MerchantName="" then
		pcPay_TD_MerchantName=pcPay_TD_MerchantName2
	end if
	pcPay_TD_MerchantPassword2=rs("pcPay_TD_MerchantPassword")
	pcPay_TD_MerchantPassword2=enDeCrypt(pcPay_TD_MerchantPassword2, scCrypPass)
	set rs=nothing
	pcPay_TD_MerchantPassword=request.Form("pcPay_TD_MerchantPassword")
	if pcPay_TD_MerchantPassword="" then
		pcPay_TD_MerchantPassword=pcPay_TD_MerchantPassword2
	end if
	pcPay_TD_ClientLang=request.Form("pcPay_TD_ClientLang")
	pcPay_TD_PayPeriod=request.Form("pcPay_TD_PayPeriod")
	if pcPay_TD_PayPeriod="" then
		pcPay_TD_PayPeriod=1
	end if
	pcPay_TD_TestMode=request.Form("pcPay_TD_TestMode")
	if pcPay_TD_TestMode="" then
		pcPay_TD_TestMode=0
	end if

	pcPay_TD_MerchantName=enDeCrypt(pcPay_TD_MerchantName, scCrypPass)
	pcPay_TD_MerchantPassword=enDeCrypt(pcPay_TD_MerchantPassword, scCrypPass)
	
	query="UPDATE pcPay_TripleDeal SET pcPay_TD_MerchantName='"&pcPay_TD_MerchantName&"',pcPay_TD_MerchantPassword='"&pcPay_TD_MerchantPassword&"',pcPay_TD_ClientLang='"&pcPay_TD_ClientLang&"',pcPay_TD_PayPeriod="&pcPay_TD_PayPeriod&",pcPay_TD_TestMode="&pcPay_TD_TestMode&" WHERE pcPay_TD_ID=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=43"
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

Function gwTripleDeal()
	varCheck=1
	'request gateway variables and insert them into the TripleDeal table
	pcPay_TD_MerchantName=request.Form("pcPay_TD_MerchantName")
	pcPay_TD_MerchantPassword=request.Form("pcPay_TD_MerchantPassword")
	pcPay_TD_Profile="standard"
	pcPay_TD_ClientLang=request.Form("pcPay_TD_ClientLang")
	pcPay_TD_PayPeriod=request.Form("pcPay_TD_PayPeriod")
	if pcPay_TD_PayPeriod="" then
		pcPay_TD_PayPeriod=1
	end if
	pcPay_TD_TestMode=request.Form("pcPay_TD_TestMode")
	if pcPay_TD_TestMode="" then
		pcPay_TD_TestMode=0
	end if

	If pcPay_TD_MerchantName="" OR pcPay_TD_MerchantPassword="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add eProcessing Network as your payment gateway. <b>""Merchant Name""</b> and <b>""Merchant Password""</b> are required fields.")
	End If
	'encrypt
	pcPay_TD_MerchantName=enDeCrypt(pcPay_TD_MerchantName, scCrypPass)
	pcPay_TD_MerchantPassword=enDeCrypt(pcPay_TD_MerchantPassword, scCrypPass)
	
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

	query="UPDATE pcPay_TripleDeal SET pcPay_TD_MerchantName='"&pcPay_TD_MerchantName&"',pcPay_TD_MerchantPassword='"&pcPay_TD_MerchantPassword&"', pcPay_TD_ClientLang='"&pcPay_TD_ClientLang&"',pcPay_TD_PayPeriod=" & pcPay_TD_PayPeriod & ",pcPay_TD_TestMode=" & pcPay_TD_TestMode & " WHERE pcPay_TD_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'TD','gwTripleDeal.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",43,'"&paymentNickName&"')"
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

if request("gwchoice")="43" then
	if request("mode")="Edit" then
		call opendb()
		query="SELECT pcPay_TD_MerchantName,pcPay_TD_MerchantPassword,pcPay_TD_ClientLang,pcPay_TD_PayPeriod,pcPay_TD_TestMode FROM pcPay_TripleDeal Where pcPay_TD_ID=1;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcPay_TD_MerchantName=rs("pcPay_TD_MerchantName")
		pcPay_TD_MerchantName=enDeCrypt(pcPay_TD_MerchantName, scCrypPass)
		pcPay_TD_MerchantPassword=rs("pcPay_TD_MerchantPassword")
		pcPay_TD_MerchantPassword=enDeCrypt(pcPay_TD_MerchantPassword, scCrypPass)
		pcPay_TD_ClientLang=rs("pcPay_TD_ClientLang")
		pcPay_TD_PayPeriod=rs("pcPay_TD_PayPeriod")
		pcPay_TD_TestMode=rs("pcPay_TD_TestMode")
				
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=43"
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
	<input type="hidden" name="addGw" value="43">
<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><b>TripleDeal</b> ( <a href="http://www.tripledeal.com/" target="_blank">Web site</a> )</td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
<table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
<tr>
    <td><strong>Edit your TripleDeal Gateway Account.<br />
        <br />
NOTE:
            TripleDeal Gateway is no longer supported by ProductCart. If you disable your account you will not be able to reactivate it.<br>
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
					<% dim pcPay_TD_MerchantNameCnt,pcPay_TD_MerchantNameEnd,pcPay_TD_MerchantNameStart
                    pcPay_TD_MerchantNameCnt=(len(pcPay_TD_MerchantName)-2)
                    pcPay_TD_MerchantNameEnd=right(pcPay_TD_MerchantName,2)
                    pcPay_TD_MerchantNameStart=""
                    for c=1 to pcPay_TD_MerchantNameCnt
                        pcPay_TD_MerchantNameStart=pcPay_TD_MerchantNameStart&"*"
                    next
                    %>
                    <tr> 
                        <td height="31" colspan="2">Current Merchant Name:&nbsp;<%=pcPay_TD_MerchantNameStart&pcPay_TD_MerchantNameEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;Account Number&quot; 
                            is only partially shown on this page. If you need 
                            to edit your account information, please re-enter 
                            your &quot;Merchant Name&quot; and &quot;Merchant 
                            Password&quot; below.</td>
                    </tr>
                <% end if %>
                <tr>
                    <td width="111"> <div align="right">Merchant Name:</div></td>
                    <td width="360"> <input type="text" name="pcPay_TD_MerchantName" size="20"></td>
                </tr>
                <tr>
                    <td width="111"> <div align="right">Merchant Password:</div></td>
                    <td width="360"> <input type="text" name="pcPay_TD_MerchantPassword" size="20"></td>
                </tr>
                <tr> 
                    <td width="111"> <div align="right">Client Language:</div></td>
                    <td width="360">
                        <select name="pcPay_TD_ClientLang">
                            <option value="nl" selected>Dutch (nl)</option>
                            <option value="en" <%if pcPay_TD_ClientLang="en" then%>selected<%end if%>>English (en)</option>
                            <option value="de" <%if pcPay_TD_ClientLang="de" then%>selected<%end if%>>German (de)</option>
                            <option value="fr" <%if pcPay_TD_ClientLang="fr" then%>selected<%end if%>>French (fr)</option>
                            <option value="es" <%if pcPay_TD_ClientLang="es" then%>selected<%end if%>>Spanish (es)</option>
                        </select></td>
                </tr>
                <tr> 
                    <td> <div align="right"> 
                            <input name="pcPay_TD_TestMode" type="checkbox" class="clearBorder" id="pcPay_TD_TestMode" value="1" <% if pcPay_TD_TestMode=1 then %>checked<% end if %> />
                        </div><input type="hidden" name="pcPay_TD_PayPeriod" value="1"></td>
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

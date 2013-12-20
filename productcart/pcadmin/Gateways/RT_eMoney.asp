<%
'---Start eMoney---
Function gweMoneyEdit()
	call opendb()
	'request gateway variables and insert them into the eMoney table
	query="SELECT pcPay_EM_MerchantID,pcPay_EM_MerchantPassword,pcPay_EM_cardTypes, pcPay_EM_CVC,pcPay_EM_TestMode FROM pcPay_Emoney Where pcPay_EM_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	pcPay_EM_MerchantID2=rs("pcPay_EM_MerchantID")
	pcPay_EM_MerchantID2=enDeCrypt(pcPay_EM_MerchantID2, scCrypPass)

	pcPay_EM_MerchantID=request.Form("pcPay_EM_MerchantID")
	if pcPay_EM_MerchantID="" then
		pcPay_EM_MerchantID=pcPay_EM_MerchantID2
	end if
	pcPay_EM_MerchantPassword2=rs("pcPay_EM_MerchantPassword")
	pcPay_EM_MerchantPassword2=enDeCrypt(pcPay_EM_MerchantPassword2, scCrypPass)
	set rs=nothing
	pcPay_EM_MerchantPassword=request.Form("pcPay_EM_MerchantPassword")
	if pcPay_EM_MerchantPassword="" then
		pcPay_EM_MerchantPassword=pcPay_EM_MerchantPassword2
	end if
	pcPay_EM_cardTypes=request.Form("pcPay_EM_cardTypes")
	pcPay_EM_CVC= request.Form("pcPay_EM_CVC")
	pcPay_EM_TestMode=request.Form("pcPay_EM_TestMode")
	if pcPay_EM_TestMode="1" then
		pcPay_EM_TestMode=1
	else
		pcPay_EM_TestMode=0
	end if

	pcPay_EM_MerchantID=enDeCrypt(pcPay_EM_MerchantID, scCrypPass)
	pcPay_EM_MerchantPassword=enDeCrypt(pcPay_EM_MerchantPassword, scCrypPass)
	
	query="UPDATE pcPay_eMoney SET pcPay_EM_MerchantID='"&pcPay_EM_MerchantID&"',pcPay_EM_MerchantPassword='"&pcPay_EM_MerchantPassword&"',pcPay_EM_cardTypes ='"&pcPay_EM_cardTypes&"',pcPay_EM_CVC=" & pcPay_EM_CVC &",pcPay_EM_TestMode=" & pcPay_EM_TestMode & " WHERE pcPay_EM_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=54"
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

Function gweMoney()
	varCheck=1
	'request gateway variables and insert them into the eMoney table
	pcPay_EM_MerchantID=request.Form("pcPay_EM_MerchantID")
	pcPay_EM_MerchantPassword=request.Form("pcPay_EM_MerchantPassword")
	pcPay_EM_cardTypes=request.Form("pcPay_EM_cardTypes")
	pcPay_EM_CVC= request.Form("pcPay_EM_CVC")
	pcPay_EM_TestMode=request.Form("pcPay_EM_TestMode")
	if pcPay_EM_TestMode="1" then
		pcPay_EM_TestMode=1
	else
		pcPay_EM_TestMode=0
	end if

	If pcPay_EM_MerchantID="" OR pcPay_EM_MerchantPassword="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add eProcessing Network as your payment gateway. <b>""Merchant ID""</b> and <b>""Merchant Password""</b> are required fields.")
	End If
	'encrypt
	pcPay_EM_MerchantID=enDeCrypt(pcPay_EM_MerchantID, scCrypPass)
	pcPay_EM_MerchantPassword=enDeCrypt(pcPay_EM_MerchantPassword, scCrypPass)
	
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

	query="UPDATE pcPay_eMoney SET pcPay_EM_MerchantID='"&pcPay_EM_MerchantID&"',pcPay_EM_MerchantPassword='"&pcPay_EM_MerchantPassword&"',pcPay_EM_cardTypes ='"&pcPay_EM_cardTypes&"',pcPay_EM_CVC=" & pcPay_EM_CVC &",pcPay_EM_TestMode=" & pcPay_EM_TestMode & " WHERE pcPay_EM_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'eMoney','gweMoney.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",54,'"&paymentNickName&"')"
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

if request("gwchoice")="54" then
	'Check if table exist in database, if not we will rediect to database update prompt
	err.clear
	call openDb()
	query="select * from pcPay_Emoney"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		set rs=nothing
		call closedb()
		response.redirect "upddbEMoney.asp"
	else
		set rs=nothing
		call closedb()
	end if
	if request("mode")="Edit" then
	 	call opendb()	
		query="SELECT pcPay_EM_MerchantID,pcPay_EM_MerchantPassword,pcPay_EM_cardTypes, pcPay_EM_CVC,pcPay_EM_TestMode FROM pcPay_Emoney Where pcPay_EM_ID=1;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcPay_EM_MerchantID=rs("pcPay_EM_MerchantID")
		pcPay_EM_MerchantID=enDeCrypt(pcPay_EM_MerchantID, scCrypPass)
		pcPay_EM_MerchantPassword=rs("pcPay_EM_MerchantPassword")
		pcPay_EM_MerchantPassword=enDeCrypt(pcPay_EM_MerchantPassword, scCrypPass)
		pcPay_EM_cardTypes=rs("pcPay_EM_cardTypes")
		pcPay_EM_CVC=rs("pcPay_EM_CVC")
		pcPay_EM_TestMode=rs("pcPay_EM_TestMode")
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=54"
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
	<input type="hidden" name="addGw" value="54">
<!-- END eMoney -->
<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/ets_logo.png" width="170" height="45"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
<table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
<tr>
    <td><strong>ETS EMoney Transaction Platform<br>
    <br>
    </strong>ETS developed the EMoney transaction platform to provide merchants with a &quot;one-stop&quot; turnkey solution for merchant processing.<strong><br>
    <br>
    <a href="https://www.etsms.com/emoney.aspx" target="_blank">ETS Website</a></strong><br />
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
                <% if request("mode")="Edit" then
					dim pcPay_EM_MerchantIDCnt,pcPay_EM_MerchantIDEnd,pcPay_EM_MerchantIDStart
					pcPay_EM_MerchantIDCnt=(len(pcPay_EM_MerchantID)-2)
					pcPay_EM_MerchantIDEnd=right(pcPay_EM_MerchantID,2)
					pcPay_EM_MerchantIDStart=""
					for c=1 to pcPay_EM_MerchantIDCnt
						pcPay_EM_MerchantIDStart=pcPay_EM_MerchantIDStart&"*"
					next
					%>
                    <tr> 
                        <td height="31" colspan="2">Current Merchant ID:&nbsp;<%=pcPay_EM_MerchantIDStart&pcPay_EM_MerchantIDEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;Account Number&quot; 
                            is only partially shown on this page. If you need 
                            to edit your account information, please re-enter 
                            your &quot;Merchant ID&quot; and &quot;Merchant 
                            Password&quot; below.</td>
                    </tr>
                <% end if %>
                <tr>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                    <td width="111"> <div align="right">Merchant ID:</div></td>
                    <td width="360"> <input type="text" name="pcPay_EM_MerchantID" size="20"></td>
                </tr>
                <tr>
                    <td width="111"> <div align="right">Merchant Password:</div></td>
                    <td width="360"> <input type="text" name="pcPay_EM_MerchantPassword" size="20"></td>
                </tr>
                    <tr> 
                    <td> <div align="right">Accepted Cards:</div></td>
                    <td> Accepted Cards: 
                    <input name="pcPay_EM_CardTypes" type="checkbox" class="clearBorder" value="VISA" <% if pcPay_EM_CardTypes="VISA" or (instr(pcPay_EM_CardTypes,"VISA,")>0) or (instr(pcPay_EM_CardTypes,", VISA")>0) then%>checked<%end if%>>
                    Visa 
                    <input name="pcPay_EM_CardTypes" type="checkbox" class="clearBorder" value="MAST" <% if pcPay_EM_CardTypes="MAST" or (instr(pcPay_EM_CardTypes,"MAST,")>0) or (instr(pcPay_EM_CardTypes,", MAST")>0) then%>checked<%end if%>>
                    MasterCard 
                    <input name="pcPay_EM_CardTypes" type="checkbox" class="clearBorder" value="AMER" <% if pcPay_EM_CardTypes="AMER" or (instr(pcPay_EM_CardTypes,"AMER,")>0) or (instr(pcPay_EM_CardTypes,", AMER")>0) then%>checked<%end if%>>
                    American Express 
                    <input name="pcPay_EM_CardTypes" type="checkbox" class="clearBorder" value="DISC" <% if pcPay_EM_CardTypes="DISC" or (instr(pcPay_EM_CardTypes,"DISC,")>0) or (instr(pcPay_EM_CardTypes,", DISC")>0) then%>checked<%end if%>>
                    Discover 
                    <input name="pcPay_EM_CardTypes" type="checkbox" class="clearBorder" value="DINE" <% if pcPay_EM_CardTypes="DINE" or (instr(pcPay_EM_CardTypes,"DINE,")>0) or (instr(pcPay_EM_CardTypes,", DINE")>0) then%>checked<%end if%>>
                    Diner's Club 
                    <input name="pcPay_EM_CardTypes" type="checkbox" class="clearBorder" value="JCB" <% if pcPay_EM_CardTypes="JCB" or (instr(pcPay_EM_CardTypes,"JCB")>0) or (instr(pcPay_ACH_CardTypes,", JCB")>0) then%>checked<%end if%>>
                    JCB
                    </td>
                </tr>
                    <tr> 
                        <td> <div align="right">Require CVC:</div></td>
                        <td>
                        <input type="radio" class="clearBorder" name="pcPay_EM_CVC" value="1" checked> Yes 
                        <input name="pcPay_EM_CVC" type="radio" class="clearBorder" value="0" <%if pcPay_EM_CVC=0 then%> Checked <%end if %> /> No
                    	</td>
                    </tr>
                    <tr> 
                    <td> <div align="right"> 
                            <input name="pcPay_EM_TestMode" type="checkbox" class="clearBorder" id="pcPay_EM_TestMode" value="1" <% if pcPay_EM_TestMode=1 then%>checked<% end if%>> 
                        </div><!--<input type="hidden" name="pcPay_EM_PayPeriod" value="1">--></td>
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

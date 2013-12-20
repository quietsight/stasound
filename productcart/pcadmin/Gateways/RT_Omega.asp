<%
'---Start Omega---
Function gwOmegaEdit()
	call opendb()
	'request gateway variables and insert them into the Omega table
	query="SELECT pcPay_OMG_MerchantID,pcPay_OMG_MerchantPassword,pcPay_OMG_TransType,pcPay_OMG_cardTypes, pcPay_OMG_CVC,pcPay_OMG_TestMode FROM pcPay_Omega Where pcPay_OMG_ID=1;"

	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	pcPay_OMG_MerchantID2=rs("pcPay_OMG_MerchantID")
	pcPay_OMG_MerchantID2=enDeCrypt(pcPay_OMG_MerchantID2, scCrypPass)

	pcPay_OMG_MerchantID=request.Form("pcPay_OMG_MerchantID")
	if pcPay_OMG_MerchantID="" then
		pcPay_OMG_MerchantID=pcPay_OMG_MerchantID2
	end if
	pcPay_OMG_MerchantPassword2=rs("pcPay_OMG_MerchantPassword")
	pcPay_OMG_MerchantPassword2=enDeCrypt(pcPay_OMG_MerchantPassword2, scCrypPass)
	set rs=nothing
	pcPay_OMG_MerchantPassword=request.Form("pcPay_OMG_MerchantPassword")
	if pcPay_OMG_MerchantPassword="" then
		pcPay_OMG_MerchantPassword=pcPay_OMG_MerchantPassword2
	end if
	pcPay_OMG_TransType =request.Form("pcPay_OMG_TransType")
	pcPay_OMG_cardTypes=request.Form("pcPay_OMG_cardTypes")
	pcPay_OMG_CVC= request.Form("pcPay_OMG_CVC")
	pcPay_OMG_TestMode=request.Form("pcPay_OMG_TestMode")
	if pcPay_OMG_TestMode="" then
		pcPay_OMG_TestMode=0
	end if
	
	pcPay_OMG_MerchantID=enDeCrypt(pcPay_OMG_MerchantID, scCrypPass)
	pcPay_OMG_MerchantPassword=enDeCrypt(pcPay_OMG_MerchantPassword, scCrypPass)
		
	query="UPDATE pcPay_Omega SET pcPay_OMG_MerchantID='"&pcPay_OMG_MerchantID&"',pcPay_OMG_MerchantPassword='"&pcPay_OMG_MerchantPassword&"',pcPay_OMG_TransType='"&pcPay_OMG_TransType&"',pcPay_OMG_cardTypes ='"&pcPay_OMG_cardTypes&"',pcPay_OMG_CVC=" & pcPay_OMG_CVC &",pcPay_OMG_TestMode=" & pcPay_OMG_TestMode & " WHERE pcPay_OMG_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=59"
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

Function gwOmega()
	varCheck=1
	'request gateway variables and insert them into the Omega table
	pcPay_OMG_MerchantID=request.Form("pcPay_OMG_MerchantID")
	pcPay_OMG_MerchantPassword=request.Form("pcPay_OMG_MerchantPassword")
	pcPay_OMG_TransType =request.Form("pcPay_OMG_TransType")
	pcPay_OMG_cardTypes=request.Form("pcPay_OMG_cardTypes")
	pcPay_OMG_CVC= request.Form("pcPay_OMG_CVC")
	pcPay_OMG_TestMode=request.Form("pcPay_OMG_TestMode")
	if pcPay_OMG_TestMode="" then
		pcPay_OMG_TestMode=0
	end if

	If pcPay_OMG_MerchantID="" OR pcPay_OMG_MerchantPassword="" then
	response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add eProcessing Network as your payment gateway. <b>""Merchant Name""</b> and <b>""Merchant Password""</b> are required fields.")
	End If
	'encrypt
	pcPay_OMG_MerchantID=enDeCrypt(pcPay_OMG_MerchantID, scCrypPass)
	pcPay_OMG_MerchantPassword=enDeCrypt(pcPay_OMG_MerchantPassword, scCrypPass)
	
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
	call opendb()
	query="UPDATE pcPay_Omega SET pcPay_OMG_MerchantID='"&pcPay_OMG_MerchantID&"',pcPay_OMG_MerchantPassword='"&pcPay_OMG_MerchantPassword&"',pcPay_OMG_TransType='"&pcPay_OMG_TransType&"',pcPay_OMG_cardTypes ='"&pcPay_OMG_cardTypes&"',pcPay_OMG_CVC=" & pcPay_OMG_CVC &",pcPay_OMG_TestMode=" & pcPay_OMG_TestMode & " WHERE pcPay_OMG_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'Omega','gwOmega.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",59,'"&paymentNickName&"')"
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

if request("gwchoice")="59" then
	'Check if table exist in database, if not we will rediect to database update prompt
	err.clear
	call openDb()
	query="select * from pcPay_Omega"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		set rs=nothing
		call closedb()
		response.redirect "upddbOmega.asp"
	else
		set rs=nothing
		call closedb()
	end if
	if request("mode")="Edit" then
	 	call opendb()	
		    query="SELECT pcPay_OMG_MerchantID,pcPay_OMG_MerchantPassword,pcPay_OMG_TransType,pcPay_OMG_cardTypes, pcPay_OMG_CVC,pcPay_OMG_TestMode FROM pcPay_Omega Where pcPay_OMG_ID=1;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
			pcPay_OMG_MerchantID=rs("pcPay_OMG_MerchantID")
			pcPay_OMG_MerchantID=enDeCrypt(pcPay_OMG_MerchantID, scCrypPass)
			pcPay_OMG_MerchantPassword=rs("pcPay_OMG_MerchantPassword")
			pcPay_OMG_MerchantPassword=enDeCrypt(pcPay_OMG_MerchantPassword, scCrypPass)
			pcPay_OMG_TransType =rs("pcPay_OMG_TransType")
			pcPay_OMG_cardTypes=rs("pcPay_OMG_cardTypes")
			pcPay_OMG_CVC=rs("pcPay_OMG_CVC")
			pcPay_OMG_TestMode=rs("pcPay_OMG_TestMode")
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=59"
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
	<input type="hidden" name="addGw" value="59">
<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/omega_logo.JPG" width="255" height="84"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
<table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
<tr>
    <td><strong><br>
    </strong>OMEGA understands the special payment needs of your Internet business. Our customized, secure, easy to use payment software virtually eliminates fraud and charge backs and qualifies your transactions at the guaranteed lowest rates. <strong><br>
    <br>
    <a href="http://www.omegap.com/" target="_blank">Omega Website</a></strong><br />
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
					dim pcPay_OMG_MerchantIDCnt,pcPay_OMG_MerchantIDEnd,pcPay_OMG_MerchantIDStart
					pcPay_OMG_MerchantIDCnt=(len(pcPay_OMG_MerchantID)-2)
					pcPay_OMG_MerchantIDEnd=right(pcPay_OMG_MerchantID,2)
					pcPay_OMG_MerchantIDStart=""
					for c=1 to pcPay_OMG_MerchantIDCnt
						pcPay_OMG_MerchantIDStart=pcPay_OMG_MerchantIDStart&"*"
                    next %>
                    <tr> 
                        <td height="31" colspan="2"><div align="Left">Current Merchant Name:&nbsp;<%=pcPay_OMG_MerchantIDStart&pcPay_OMG_MerchantIDEnd%></div></td>
                    </tr>
                    <tr> 
                    <td colspan="2"> For security reasons, your &quot;Merchant Name&quot; 
                        is only partially shown on this page. If you need 
                        to edit your account information, please re-enter 
                        your &quot;Merchant Name&quot; and &quot;Merchant 
                        Password&quot; below.</td>
                    </tr>
                <% end if %>
                <tr>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                    <td width="20%"><div align="right">Merchant Name:</div></td>
                    <td><input type="text" value="" name="pcPay_OMG_MerchantID" size="24"></td>
                </tr>
                <tr> 
                    <td width="20%" height="31"><div align="right">Merchant Password:</div></td>
                    <td height="31"><input type="text" value="" name="pcPay_OMG_MerchantPassword" size="24"></td>
                </tr>
                <tr> 
                    <td valign="top"><div align="right">Transaction Type:</div></td>
                    <td valign="top">
                        <select name="pcPay_OMG_TransType">                                    
                            <option value="auth" selected>Authorize Only</option>
                            <option value="sale" <% if pcPay_OMG_TransType="sale" then%>selected<% end if %>>Sale</option>
                        </select>
                	</td>
                </tr>
                <tr> 
                    <td> <div align="right">Accepted Cards:</div></td>
                    <td> 
                    <input name="pcPay_OMG_CardTypes" type="checkbox" class="clearBorder" value="VISA" <%if pcPay_OMG_CardTypes="VISA" or (instr(pcPay_OMG_CardTypes,"VISA")>0) or (instr(pcPay_OMG_CardTypes,", VISA")>0) then%>checked<%end if%>> Visa 
                    <input name="pcPay_OMG_CardTypes" type="checkbox" class="clearBorder" value="MAST" <%if pcPay_OMG_CardTypes="MAST" or (instr(pcPay_OMG_CardTypes,"MAST")>0) or (instr(pcPay_OMG_CardTypes,", MAST")>0) then%>checked<%end if%>> MasterCard 
                    <input name="pcPay_OMG_CardTypes" type="checkbox" class="clearBorder" value="AMER" <%if pcPay_OMG_CardTypes="AMER" or (instr(pcPay_OMG_CardTypes,"AMER")>0) or (instr(pcPay_OMG_CardTypes,", AMER")>0) then%>checked<%end if%>> American Express 
                    <input name="pcPay_OMG_CardTypes" type="checkbox" class="clearBorder" value="DISC" <%if pcPay_OMG_CardTypes="DISC" or (instr(pcPay_OMG_CardTypes,"DISC")>0) or (instr(pcPay_OMG_CardTypes,", DISC")>0) then%>checked<%end if%>> Discover 
                    <input name="pcPay_OMG_CardTypes" type="checkbox" class="clearBorder" value="DINE" <%if pcPay_OMG_CardTypes="DINE" or (instr(pcPay_OMG_CardTypes,"DINE")>0) or (instr(pcPay_OMG_CardTypes,", DINE")>0) then%>checked<%end if%>> Diner's Club 
                    <input name="pcPay_OMG_CardTypes" type="checkbox" class="clearBorder" value="JCB" <%if pcPay_OMG_CardTypes="JCB" or (instr(pcPay_OMG_CardTypes,"JCB")>0) or (instr(pcPay_ACH_CardTypes,", JCB")>0) then%>checked<%end if%>> JCB</td>
                </tr>
                <tr> 
                    <td> <div align="right">Require CVC:</div></td>
                    <td>
                        <input type="radio" class="clearBorder" name="pcPay_OMG_CVC" value="1" Checked> Yes 
                        <input name="pcPay_OMG_CVC" type="radio" class="clearBorder" value="0" <%if pcPay_OMG_CVC=0 then%>Checked<%end if %>> No
                    </td>
                </tr>
                <tr> 
                    <td><div align="right"> 
                        <input name="pcPay_OMG_testmode" type="checkbox" class="clearBorder" id="pcPay_OMG_testmode" value="1" <% if pcPay_OMG_testmode="1" then%>checked<% end if %>></div></td>
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

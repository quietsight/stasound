<!-- START EMerchant -->
<% 
Function gwEMerchantEdit()
	call opendb()
	'request gateway variables and insert them into the eMerchant table
	query= "SELECT pcPay_eMerch_MerchantID, pcPay_eMerch_PaymentKey FROM pcPay_eMerchant where pcPay_eMerch_ID=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		call closedb()
	  response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddPaymentOpt: "&Err.Description) 
	end If
	pcPay_eMerch_MerchantID2=rs("pcPay_eMerch_MerchantID")
	pcPay_eMerch_MerchantID=request.Form("pcPay_eMerch_MerchantID")
	pcPay_eMerch_MerchantID=enDeCrypt(pcPay_eMerch_MerchantID, scCrypPass)
	If pcPay_eMerch_MerchantID="" then
		pcPay_eMerch_MerchantID=pcPay_eMerch_MerchantID2
	End If
	
	pcPay_eMerch_PaymentKey2=rs("pcPay_eMerch_PaymentKey")
	pcPay_eMerch_PaymentKey=request.Form("pcPay_eMerch_PaymentKey")
	pcPay_eMerch_PaymentKey=enDeCrypt(pcPay_eMerch_PaymentKey, scCrypPass)
	If pcPay_eMerch_PaymentKey="" then
		pcPay_eMerch_PaymentKey=pcPay_eMerch_PaymentKey2
	End If
	set rs=nothing

	pcPay_eMerch_CVD=request.Form("pcPay_eMerch_CVD")
	pcPay_eMerch_CardType=request.Form("pcPay_eMerch_CardType")
	pcPay_eMerch_TransType=request.Form("pcPay_eMerch_TransType")
	if NOT isNumeric(pcPay_eMerch_CVD) OR pcPay_eMerch_CVD="" then
		pcPay_eMerch_CVD=0
	end if
	
	pcPay_eMerch_TestMode=request.Form("pcPay_eMerch_TestMode")
	if NOT isNumeric(pcPay_eMerch_TestMode) OR pcPay_eMerch_TestMode="" then
		pcPay_eMerch_TestMode=0
	end if

	query="UPDATE pcPay_eMerchant SET pcPay_eMerch_MerchantID='"&pcPay_eMerch_MerchantID&"',pcPay_eMerch_PaymentKey='"&pcPay_eMerch_PaymentKey&"',pcPay_eMerch_CVD="&pcPay_eMerch_CVD&",pcPay_eMerch_CardType='"&pcPay_eMerch_CardType&"', pcPay_eMerch_TransType='"&pcPay_eMerch_TransType&"', pcPay_eMerch_TestMode="&pcPay_eMerch_TestMode&" WHERE pcPay_eMerch_ID=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		call closedb()
	  response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddPaymentOpt: "&Err.Description) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ", priceToAdd="&priceToAdd&" ,percentageToAdd="&percentageToAdd&" ,paymentNickName='"&paymentNickName&"' WHERE gwCode=51"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		call closedb()
	  response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddPaymentOpt: "&Err.Description) 
	end If
	call closedb()
end Function

Function gwEMerchant()
	varCheck=1
	'request gateway variables and insert them into the echo table
	pcPay_eMerch_MerchantID=request.Form("pcPay_eMerch_MerchantID")

	If pcPay_eMerch_MerchantID="" then
		call closedb()
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add eMerchant as your payment gateway. <b>""Merchant ID""</b> is a required field.")
	End If
	'encrypt
	pcPay_eMerch_MerchantID=enDeCrypt(pcPay_eMerch_MerchantID, scCrypPass)
				
	pcPay_eMerch_PaymentKey=request.Form("pcPay_eMerch_PaymentKey")
	If pcPay_eMerch_PaymentKey="" then
		call closedb()
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Echo as your payment gateway. <b>""Merchant's PIN""</b> is a required field.")
	End If
	'encrypt
	pcPay_eMerch_PaymentKey=enDeCrypt(pcPay_eMerch_PaymentKey, scCrypPass)
	pcPay_eMerch_CVD=request.Form("pcPay_eMerch_CVD")
	if NOT isNumeric(pcPay_eMerch_CVD) OR pcPay_eMerch_CVD="" then
		pcPay_eMerch_CVD=0
	end if
	pcPay_eMerch_CardType=request.Form("pcPay_eMerch_CardType")
	pcPay_eMerch_TransType=request.Form("pcPay_eMerch_TransType")
	pcPay_eMerch_TestMode=request.Form("pcPay_eMerch_TestMode")
	if NOT isNumeric(pcPay_eMerch_TestMode) OR pcPay_eMerch_TestMode="" then
		pcPay_eMerch_TestMode=0
	end if
	
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
	query="UPDATE pcPay_eMerchant SET pcPay_eMerch_MerchantID='"&pcPay_eMerch_MerchantID&"',pcPay_eMerch_PaymentKey='"&pcPay_eMerch_PaymentKey&"',pcPay_eMerch_CVD="&pcPay_eMerch_CVD&",pcPay_eMerch_CardType='"&pcPay_eMerch_CardType&"', pcPay_eMerch_TransType='"&pcPay_eMerch_TransType&"', pcPay_eMerch_TestMode="&pcPay_eMerch_TestMode&" WHERE pcPay_eMerch_ID=1"
	set rs=Server.CreateObject("ADODB.Recordset")  
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'eMerchant','gweMerchant.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",51,'"&paymentNickName&"')"
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

<% if request("gwchoice")="51" then
	MC="0"
	MD="0"
	SO="0"
	SW="0"
	VI="0"
	VD="0"
	VE="0"
	if request("mode")="Edit" then
		call opendb()
		query="SELECT pcPay_eMerch_MerchantID, pcPay_eMerch_PaymentKey, pcPay_eMerch_CVD, pcPay_eMerch_CardType, pcPay_eMerch_TransType, pcPay_eMerch_TestMode FROM pcPay_eMerchant WHERE pcPay_eMerch_ID=1;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error in modPaymentOpt: "&Err.Description) 
		end If
		pcPay_eMerch_MerchantID=rs("pcPay_eMerch_MerchantID")
		pcPay_eMerch_MerchantID=enDeCrypt(pcPay_eMerch_MerchantID, scCrypPass)
		pcPay_eMerch_PaymentKey=rs("pcPay_eMerch_PaymentKey")
		pcPay_eMerch_CVD=rs("pcPay_eMerch_CVD")
		pcPay_eMerch_CardType=rs("pcPay_eMerch_CardType")
		pcPay_eMerch_TransType=rs("pcPay_eMerch_TransType")
		pcPay_eMerch_TestMode=rs("pcPay_eMerch_TestMode")
		cardTypeArray=split(pcPay_eMerch_CardType,", ")
		for i=0 to ubound(cardTypeArray)
			select case cardTypeArray(i)
			case "MC"
				MC="1" 
			case "MD"
				MD="1" 
			case "SO"
				SO="1" 
			case "SW"
				SW="1" 
			case "VI"
				VI="1" 
			case "VD"
				VD="1" 
			case "VE"
				VE="1" 
			end select
		next
		set rs=nothing
		call closedb()
		%>
    
		<input type="hidden" name="mode" value="Edit">
    <% end if %>
	<input type="hidden" name="addGw" value="51">
    <tr> 
        <td height="21"><b>Fasthosts eMerchant </b> ( <a href="http://www.fasthosts.co.uk" target="_blank">Web site</a> )</td>
    </tr>
    <tr> 
        <td>
            <table width="100%" border="0" cellspacing="0" cellpadding="4">
                <tr> 
                    <th colspan="2">Enter eMerchant settings</th>
                </tr>
                <% if request("mode")="Edit" then %>
                    <tr> 
                        <td width="20%">&nbsp;</td>
                        <% dim pcPay_eMerch_MerchantIDCnt,pcPay_eMerch_MerchantIDEnd,pcPay_eMerch_MerchantIDStart
                        pcPay_eMerch_MerchantIDCnt=(len(pcPay_eMerch_MerchantID)-2)
                        pcPay_eMerch_MerchantIDEnd=right(pcPay_eMerch_MerchantID,2)
                        pcPay_eMerch_MerchantIDStart=""
                        for c=1 to pcPay_eMerch_MerchantIDCnt
                            pcPay_eMerch_MerchantIDStart=pcPay_eMerch_MerchantIDStart&"*"
                        next %>
                        <td>Current Merchant ID:&nbsp;<%=pcPay_eMerch_MerchantIDStart&pcPay_eMerch_MerchantIDEnd %></td>
                    </tr>
                    <tr> 
                        <td>&nbsp;</td>
                        <td> For security reasons, your &quot;Store Name&quot; is 
                            only partially shown on this page. If you need to edit 
                            your account information, please re-enter your &quot;Store 
                            Name&quot; below.</td>
                    </tr>
                <% end if %>
                <tr bgcolor="#FFFFFF"> 
                    <td width="24%"> <div align="right">Merchant ID : </div></td>
                    <td width="76%"> <div align="left"> 
                            <input type="text" value="" name="pcPay_eMerch_MerchantID" size="30">
                        </div></td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td width="24%"> <div align="right">Payment Key : </div></td>
                    <td width="76%"> <div align="left"> 
                            <input type="text" value="" name="pcPay_eMerch_PaymentKey" size="30">
                        </div></td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td> <div align="right">Transaction Type:</div></td>
                    <td> <select name="pcPay_eMerch_TransType">
                            <option value="P" selected>Purchase</option>
                            <option value="A" <% if pcPay_eMerch_TransType="A" then%>selected<% end if %>>Authorize Only</option>
                        </select> </td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td> <div align="right">Require CVD:</div></td>
                    <td> <input type="radio" class="clearBorder" name="pcPay_eMerch_CVD" value="1" checked>
                        Yes 
                        <input name="pcPay_eMerch_CVD" type="radio" class="clearBorder" value="0" <% if pcPay_eMerch_CVD="0" then%>checked<% end if %>>
                        No<font color="#FF0000"></font></td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td> <div align="right">Accepted Cards:</div></td>
                    <td> 
                        <input name="pcPay_eMerch_CardType" type="checkbox" class="clearBorder" value="MC" <% if MC="1" then%>checked<% end if %>>
                        Master Card 
                        <input name="pcPay_eMerch_CardType" type="checkbox" class="clearBorder" value="MD" <% if MD="1" then%>checked<% end if %>>
                        Maestro 
                        <input name="pcPay_eMerch_CardType" type="checkbox" class="clearBorder" value="SO" <% if SO="1" then%>checked<% end if %>>
                        Solo 
                        <input name="pcPay_eMerch_CardType" type="checkbox" class="clearBorder" value="SW" <% if SW="1" then%>checked<% end if %>>
                        Switch/Maestro
                        <input name="pcPay_eMerch_CardType" type="checkbox" class="clearBorder" value="VI" <% if VI="1" then%>checked<% end if %>>
                        Visa
                        <input name="pcPay_eMerch_CardType" type="checkbox" class="clearBorder" value="VD" <% if VD="1" then%>checked<% end if %>>
                        Visa Delta
                        <input name="pcPay_eMerch_CardType" type="checkbox" class="clearBorder" value="VE" <% if VE="1" then%>checked<% end if %>>
                        Visa Electron</td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td><div align="right"> 
                        <input name="pcPay_eMerch_TestMode" type="checkbox" class="clearBorder" value="1" <% if pcPay_eMerch_TestMode=1 then%>checked<% end if%>></div></td>
                    <td><b>Enable Test Mode </b>(Credit cards will not be charged)</td>
                </tr>
                
                <tr> 
                    <th colspan="2">You have the option to charge a processing fee for this payment option.</th>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td width="111"> <div align="right">Processing fee: </div></td>
                    <td> <input type="radio" class="clearBorder" name="priceToAddType" value="price" <% if priceToAddType="price" then%>checked<%end if%>>
                        Flat fee&nbsp;&nbsp; &nbsp;<%=scCurSign%> <input name="priceToAdd" size="6" value="<%=money(priceToAdd)%>">
                    </td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td>&nbsp;</td>
                    <td> <input type="radio" class="clearBorder" name="priceToAddType" value="percentage" <% if priceToAddType="percentage" then%>checked<%end if%>>
                        Percentage of Order Total&nbsp;&nbsp; &nbsp; % 
                        <input name="percentageToAdd" size="6" value="<%=percentageToAdd%>">
                    </td>
                </tr>
                <tr> 
                    <th colspan="2">You can change the display name that is shown for this payment type. </th>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td><div align="right">Payment Name:&nbsp;</div></td>
                    <td><input name="paymentNickName" value="<%=paymentNickName%>" size="35" maxlength="255"></td>
                </tr>
            </table>
    	</td>
    </tr>
<!-- END Emerchant -->

<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="images/pcv4_icon_pg.png" width="48" height="48"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
<table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
<tr>
    <td><strong>Enter your Gateway Account information<br>
    <br>
    </strong>Gateway description goes here!<strong><br>
    <br>
    <a href="#">Sign Up Now! </a></strong></td>
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
                  <tr>
                    <td width="18%" nowrap><span class="pcSubmenuHeader">PayPal ID/E-mail</span><br /></td>
                    <td width="82%" class="pcSubmenuContent"><p>
                      <label for="textfield"></label>
                      <input type="text" name="textfield" id="textfield">
                    </p></td>
                  </tr>
                  <tr>
                    <td>Currency</td>
                    <td class="pcSubmenuContent">
                        <select name="PayPal_Currency">
                            <option value="USD">U.S. Dollars ($)</option>
                            <option value="AUD">Australian Dollars ($)</option>
                            <option value="CAD">Canadian Dollars (C $)</option>
                            <option value="CZK">Czech Koruna</option>
                            <option value="DKK">Danish Krone</option>
                            <option value="EUR">Euros (€)</option>
                            <option value="HKD">Hong Kong Dollar</option>
                            <option value="HUF">Hungarian Forint</option>
                            <option value="ILS">Israeli New Shekel</option>
                            <option value="JPY">Yen (¥)</option>
                            <option value="MXN">Mexican Peso</option> 
                            <option value="NOK">Norwegian Krone</option>
                            <option value="NZD">New Zealand Dollar</option>
                            <option value="PHP">Philippine Peso</option> 
                            <option value="PLN">Polish Zloty</option>
                            <option value="GBP">Pounds Sterling (£)</option>											
                            <option value="SGD">Singapore Dollar</option>
                            <option value="SEK">Swedish Krona</option>
                            <option value="CHF">Swiss Franc</option>     
                            <option value="TWD">Taiwan New Dollar</option>    
                            <option value="THB">Thai Baht</option>
                        </select>
                    </td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td class="pcSubmenuContent"><input name="gwvpfl" type="checkbox" class="clearBorder" value="1" <% if request("gwchoice")="VeriSignLk" then%>Checked<%end if%>> 
                                <a name="GWA"></a>Enable PayPal Payflow Link - <a href="https://www.paypal.com/cgi-bin/webscr?cmd=_profile-comparison" target="_blank">More Information</a></td>
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
                    <td width="82%" class="pcSubmenuContent"><p>
                      <input type="radio" class="clearBorder" name="priceToAddType" value="price">Flat fee&nbsp;&nbsp; &nbsp;$<input name="priceToAdd" size="6" value="0.00">
                    </p></td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td class="pcSubmenuContent"><input type="radio" class="clearBorder" name="priceToAddType" value="percentage">Percentage of Order Total&nbsp;&nbsp; &nbsp; %<input name="percentageToAdd" size="6" value="0"></td>
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
                        <td width="82%"><input name="paymentNickName" value="Credit Card" size="35" maxlength="255"></td>
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
                        <td>Process orders when they are placed: <input type="checkbox" class="clearBorder" name="pcv_processOrder" value="1">&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=301')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
                    </tr>
                    <tr> 
                        <td>&nbsp;</td>
                        <td>When orders are placed, set the payment status to:
                        <select name="pcv_setPayStatus">
                            <option value="3" selected="selected">Default</option>
                            <option value="0">Pending</option>
                            <option value="1">Authorized</option>
                            <option value="2">Paid</option>
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

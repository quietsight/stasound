<%
'---Start viaKLIX---
Function gwklixEdit()
	call opendb()
	'request gateway variables and insert them into the Klix table
	query="SELECT ssl_merchant_id,ssl_pin,testmode,CVV,ssl_user_id FROM klix WHERE idKlix=1"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	ssl_merchant_id2=rstemp("ssl_merchant_id")
	ssl_merchant_id=request.Form("ssl_merchant_id")
	if ssl_merchant_id="" then
		ssl_merchant_id=ssl_merchant_id2
	end if
	ssl_pin2=rstemp("ssl_pin")
	'decrypt
	ssl_pin2=enDeCrypt(ssl_pin2, scCrypPass)
	ssl_pin=request.Form("ssl_pin")
	if ssl_pin="" then
		ssl_pin=ssl_pin2
	end if
	'encrypt
	ssl_pin=enDeCrypt(ssl_pin, scCrypPass)
	ssl_user_id2=rstemp("ssl_user_id")
	ssl_user_id=request.Form("ssl_user_id")
	if ssl_user_id="" then
		ssl_user_id=ssl_user_id2
	end if
	CVV=request.Form("CVV")
	testmode=request.Form("testmode")
	query="UPDATE klix SET ssl_merchant_id='"&ssl_merchant_id&"',ssl_pin='"&ssl_pin&"',CVV="&CVV&",ssl_user_id='"&ssl_user_id&"' WHERE idKlix=1"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="&priceToAdd&" , percentageToAdd="&percentageToAdd&", paymentNickName='"&paymentNickName&"' WHERE gwCode=23"

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

Function gwklix()
	varCheck=1
	'request gateway variables and insert them into the klix table
	ssl_merchant_id=request.Form("ssl_merchant_id")
	ssl_user_id=request.Form("ssl_user_id")
	ssl_pin=request.Form("ssl_pin")
	'encrypt
	ssl_pin=enDeCrypt(ssl_pin, scCrypPass)
	CVV=request.Form("CVV")
	testmode=request.Form("ssl_testmode")
	if testmode="YES" then
		testmode="1"
	else
		testmode="0"
	end if
	if NOT isNumeric(CVV) or CVV="" then
		CVV="0"
	end if
	If ssl_merchant_id="" OR ssl_user_id="" then
		call closedb()
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add viaKlix as your payment gateway. <b>""Merchant ID""</b> and <b>""User ID""</b> are required fields.")
	End If
	If ssl_pin="" then
		call closedb()
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add viaKlix as your payment gateway. <b>""PIN""</b> is a required field.")
	End If
	'encrypt
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

	query="UPDATE klix SET ssl_merchant_id='"&ssl_merchant_id&"',ssl_pin='"&ssl_pin&"',CVV="&CVV&", testmode="&testmode&",ssl_user_id='"&ssl_user_id&"' WHERE idKlix=1"
	set rs=Server.CreateObject("ADODB.Recordset")  

	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'viaKLIX','gwklix.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",23,'"&paymentNickName&"')"
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

if request("gwchoice")="23" then
	if request("mode")="Edit" then
		call opendb()
		query= "SELECT ssl_merchant_id,ssl_pin,CVV,testmode,ssl_user_id FROM klix WHERE idKlix=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		ssl_merchant_id=rs("ssl_merchant_id")
		ssl_pin=rs("ssl_pin")
			'decrypt
			ssl_pin=enDeCrypt(ssl_pin, scCrypPass)
		CVV=rs("CVV")
		testmode=rs("testmode")
		ssl_user_id=rs("ssl_user_id")
		set rs=nothing
		call closedb()
		%>
		<input type="hidden" name="mode" value="Edit">
    <% end if %>
	<input type="hidden" name="addGw" value="23">
    <tr> 
        <td height="21"><b>viaKLIX</b> ( <a href="https://www2.viaklix.com/Admin/main.asp" target="_blank">Web site</a> )</td>
    </tr>
    <tr> 
        <td>
            <table width="100%" border="0" cellspacing="0" cellpadding="4">
                <tr> 
                    <th colspan="2">Enter viaKLIX settings</th>
                </tr>
                <% if request("mode")="Edit" then %>
					<% dim ssl_merchant_idCnt,ssl_merchant_idEnd,ssl_merchant_idStart,ssl_user_idCnt,ssl_user_idEnd,ssl_user_idStart
                    ssl_merchant_idCnt=(len(ssl_merchant_id)-2)
                    ssl_merchant_idEnd=right(ssl_merchant_id,2)
                    ssl_merchant_idStart=""
                    for c=1 to ssl_merchant_idCnt
                        ssl_merchant_idStart=ssl_merchant_idStart&"*"
                    next
                    ssl_user_idCnt=(len(ssl_user_id)-2)
                    ssl_user_idEnd=right(ssl_user_id,2)
                    ssl_user_idStart=""
                    for c=1 to ssl_user_idCnt
                        ssl_user_idStart=ssl_user_idStart&"*"
                    next %>
                    <tr> 
                        <td colspan="2">Current Merchant ID:&nbsp;<%=ssl_merchant_idStart&ssl_merchant_idEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2">Current User ID:&nbsp;<%=ssl_user_idStart&ssl_user_idEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;Merchant ID&quot; 
                            and &quot;User ID&quot; are only partially shown 
                            on this page. If you need to edit your account information, 
                            please re-enter your &quot;Merchant ID&quot; and 
                            &quot;User ID&quot; below.</td>
                    </tr>
                <% end if %>
                <tr> 
                    <td><div align="right">Merchant ID:</div></td>
                    <td width="360"> <input type="text" name="ssl_merchant_id" size="20"></td>
                </tr>
                <tr> 
                    <td><div align="right">User ID:</div></td>
                    <td width="360"> <input type="text" name="ssl_user_id" size="20"></td>
                </tr>
                <tr> 
                    <td><div align="right">PIN # :</div></td>
                    <td><input name="ssl_pin" type="password" size="20"> </td>
                </tr>
                <tr> 
                    <td><div align="right">Require CVV:</div></td>
                    <td><input type="radio" class="clearBorder" name="CVV" value="1" checked>
                        Yes 
                        <input type="radio" class="clearBorder" name="CVV" value="0" <% if CVV="0" then %>checked<% end if %>>
                        No</td>
                </tr>
                <tr> 
                    <th colspan="2">You have the option to charge a processing fee for this payment option.</th>
                </tr>
                <tr> 
                    <td width="111"> <div align="right">Processing fee: </div></td>
                    <td> <input type="radio" class="clearBorder" name="priceToAddType" value="price" <% if priceToAddType="price" then%>checked<%end if%>>
                        Flat fee&nbsp;&nbsp; &nbsp;<%=scCurSign%> <input name="priceToAdd" size="6" value="<%=money(priceToAdd)%>">
                    </td>
                </tr>
                <tr> 
                    <td>&nbsp;</td>
                    <td> <input type="radio" class="clearBorder" name="priceToAddType" value="percentage" <% if priceToAddType="percentage" then%>checked<%end if%>>
                        Percentage of Order Total&nbsp;&nbsp; &nbsp; % 
                        <input name="percentageToAdd" size="6" value="<%=percentageToAdd%>">
                    </td>
                </tr>
 <tr> 
                    <th colspan="2">You can change the display name that is shown for this payment type. </th>
                </tr>
                <tr> 
                    <td><div align="right">Payment Name:&nbsp;</div></td>
                    <td><input name="paymentNickName" value="<%=paymentNickName%>" size="35" maxlength="255"></td>
                </tr>
            </table>
        </td>
    </tr>
<!-- END viaKLIX -->

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

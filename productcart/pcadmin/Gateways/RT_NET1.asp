<%
'---Start NETOne---
Function gwNETOneEdit()
	call opendb()
	'request gateway variables and insert them into the USAePay table
	query="SELECT pcPay_NETOne.pcPay_NETOne_Mid, pcPay_NETOne.pcPay_NETOne_Mkey, pcPay_NETOne.pcPay_NETOne_TCode, pcPay_NETOne.pcPay_NETOne_CVV, pcPay_NETOne.pcPay_NETOne_CardTypes FROM pcPay_NETOne WHERE (((pcPay_NETOne.pcPay_NETOne_Id)=1));"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	pcPay_NETOne_Mid=rs("pcPay_NETOne_Mid")
	pcPay_NETOne_Mkey=rs("pcPay_NETOne_Mkey")
	
	pcPay_NETOne_Mid2=pcPay_NETOne_Mid
	'decrypt
	pcPay_NETOne_Mid2=enDeCrypt(pcPay_NETOne_Mid2, scCrypPass)
	pcPay_NETOne_Mid=request.Form("pcPay_NETOne_Mid")
	if pcPay_NETOne_Mid="" then
		pcPay_NETOne_Mid=pcPay_NETOne_Mid2
	end if
	'encrypt
	pcPay_NETOne_Mid=enDeCrypt(pcPay_NETOne_Mid, scCrypPass)

	pcPay_NETOne_Mkey2=pcPay_NETOne_Mkey
	'decrypt
	pcPay_NETOne_Mkey2=enDeCrypt(pcPay_NETOne_Mkey2, scCrypPass)
	pcPay_NETOne_Mkey=request.Form("pcPay_NETOne_Mkey")
	if pcPay_NETOne_Mkey="" then
		pcPay_NETOne_Mkey=pcPay_NETOne_Mkey2
	end if
	'encrypt
	pcPay_NETOne_Mkey=enDeCrypt(pcPay_NETOne_Mkey, scCrypPass)
	
	pcPay_NETOne_TCode=request.Form("pcPay_NETOne_TCode")
	pcPay_NETOne_CVV=request.Form("pcPay_NETOne_CVV")
	pcPay_NETOne_CardTypes=request.Form("pcPay_NETOne_CardTypes")

	query="UPDATE pcPay_NETOne SET pcPay_NETOne_Mid='"&pcPay_NETOne_Mid&"',pcPay_NETOne_Mkey='"&pcPay_NETOne_Mkey&"',pcPay_NETOne_TCode='" & pcPay_NETOne_TCode & "',pcPay_NETOne_CVV="&pcPay_NETOne_CVV&",pcPay_NETOne_CardTypes='"&pcPay_NETOne_CardTypes&"' WHERE pcPay_NETOne_Id=1"

	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=40"

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

Function gwNETOne()
	varCheck=1
	'request gateway variables and insert them into the NETOne table

	pcPay_NETOne_Mid=request.Form("pcPay_NETOne_Mid")
	pcPay_NETOne_Mkey=request.Form("pcPay_NETOne_Mkey")
	pcPay_NETOne_Tcode=request.Form("pcPay_NETOne_Tcode")
	pcPay_NETOne_CVV=request.Form("pcPay_NETOne_CVV")
	pcPay_NETOne_CardTypes=request.Form("pcPay_NETOne_CardTypes")

	If pcPay_NETOne_Mid="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add NET1 as your payment gateway. <b>""Merchant ID""</b> is a required field.")
	End If
	'encrypt
	pcPay_NETOne_Mid=enDeCrypt(pcPay_NETOne_Mid, scCrypPass)
	If pcPay_NETOne_Mkey="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add NET1 as your payment gateway. <b>""Merchant Key""</b> is a required field.")
	End If
	'encrypt
	pcPay_NETOne_Mkey=enDeCrypt(pcPay_NETOne_Mkey, scCrypPass)

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

	query="UPDATE pcPay_NETOne SET pcPay_NETOne_Mid='"&pcPay_NETOne_Mid&"',pcPay_NETOne_Mkey='"&pcPay_NETOne_Mkey&"',pcPay_NETOne_Tcode='" & pcPay_NETOne_Tcode & "',pcPay_NETOne_CVV="&pcPay_NETOne_CVV&",pcPay_NETOne_CardTypes='"&pcPay_NETOne_CardTypes&"' WHERE pcPay_NETOne_Id=1"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'NETOne','gwNETOne.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",40,'"&paymentNickName&"')"
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
%>                

<% if request("gwchoice")="40" then
	if request("mode")="Edit" then
		call opendb()
		query="SELECT pcPay_NETOne.pcPay_NETOne_MID, pcPay_NETOne.pcPay_NETOne_Mkey, pcPay_NETOne.pcPay_NETOne_Tcode, pcPay_NETOne.pcPay_NETOne_CVV, pcPay_NETOne.pcPay_NETOne_CardTypes FROM pcPay_NETOne WHERE (((pcPay_NETOne.pcPay_NETOne_ID)=1));"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcPay_NETOne_Mid=rs("pcPay_NETOne_Mid")
		pcPay_NETOne_Mid=enDeCrypt(pcPay_NETOne_Mid, scCrypPass)
		pcPay_NETOne_Mkey=rs("pcPay_NETOne_Mkey")
		pcPay_NETOne_Mkey=enDeCrypt(pcPay_NETOne_Mkey, scCrypPass)
		pcPay_NETOne_TCode=rs("pcPay_NETOne_TCode")
		pcPay_NETOne_CVV=rs("pcPay_NETOne_CVV")
		pcPay_NETOne_CardTypes=rs("pcPay_NETOne_CardTypes")
		set rs=nothing
		call closedb()
		%>
    
		<input type="hidden" name="mode" value="Edit">
    <% end if %>
	<input type="hidden" name="addGw" value="40">
    <tr> 
        <td height="21"><b>NET1 Payment Services</b> ( <a href="http://www.eftsecure.com/" target="_blank">Web site</a> )</td>
    </tr>
    <tr> 
        <td>
        	<table width="100%" border="0" cellspacing="0" cellpadding="4">
                <tr> 
                    <th colspan="2">Enter NET1 settings</th>
                </tr>
                <% if request("mode")="Edit" then %>
					<% dim pcPay_NETOne_MIDCnt,pcPay_NETOne_MIDEnd,pcPay_NETOne_MIDStart
                    pcPay_NETOne_MIDCnt=(len(pcPay_NETOne_MID)-2)
                    pcPay_NETOne_MIDEnd=right(pcPay_NETOne_MID,2)
                    pcPay_NETOne_MIDStart=""
                    for c=1 to pcPay_NETOne_MIDCnt
                        pcPay_NETOne_MIDStart=pcPay_NETOne_MIDStart&"*"
                    next %>
                    <tr> 
                        <td height="31" colspan="2">Current ACH Direct Merchant ID:&nbsp;<%=pcPay_NETOne_MIDStart&pcPay_NETOne_MIDEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;Merchant ID&quot; 
                            is only partially shown on this page. If you need 
                            to edit your account information, please re-enter 
                            your &quot;Merchant ID&quot; and &quot;Password&quot; 
                            below.</td>
                    </tr>
                <% end if %>
                <tr> 
                    <td valign="top"><div align="right">Merchant ID:</div></td>
                    <td valign="top"> <input type="text" name="pcPay_NETOne_Mid" size="20">
                        12 Digit Merchant Identification Number</td>
                </tr>
                <tr> 
                    <td valign="top"><div align="right">Merchant Key:</div></td>
                    <td valign="top"><input type="text" name="pcPay_NETOne_Mkey" size="20">
                        12 Digit Merchant Key</td>
                </tr>
                <tr> 
                    <td valign="top"><div align="right">Transaction Type:</div></td>
                    <td valign="top"> <select name="pcPay_NETOne_Tcode">
                            <option value="01" selected>Sale</option>
                            <option value="02" <%if pcPay_NETOne_TCode="02" then%>selected<%end if%>>Authorize Only</option>
                        </select> </td>
                </tr>
                <tr> 
                    <td> <div align="right">Require CVV:</div></td>
                    <td> <input type="radio" class="clearBorder" name="pcPay_NETOne_CVV" value="1" checked>
                        Yes 
                        <input name="pcPay_NETOne_CVV" type="radio" class="clearBorder" value="0" <%if clng(pcPay_NETOne_CVV)=0 then%>checked<%end if%>>
                        No</td>
                </tr>
                <tr> 
                    <td> <div align="right">Accepted Cards:</div></td>
                    <td>
                    	<input name="pcPay_NETOne_CardTypes" type="checkbox" class="clearBorder" value="VISA" <% if pcPay_NETOne_CardTypes="VISA" or (instr(pcPay_NETOne_CardTypes,"VISA,")>0) or (instr(pcPay_NETOne_CardTypes,", VISA")>0) then%>checked<%end if%>> Visa 
                        <input name="pcPay_NETOne_CardTypes" type="checkbox" class="clearBorder" value="MAST" <% if pcPay_NETOne_CardTypes="MAST" or (instr(pcPay_NETOne_CardTypes,"MAST,")>0) or (instr(pcPay_NETOne_CardTypes,", MAST")>0) then%>checked<%end if%>> MasterCard 
                        <input name="pcPay_NETOne_CardTypes" type="checkbox" class="clearBorder" value="AMER" <% if pcPay_NETOne_CardTypes="AMER" or (instr(pcPay_NETOne_CardTypes,"AMER,")>0) or (instr(pcPay_NETOne_CardTypes,", AMER")>0) then%>checked<%end if%>> American Express 
                        <input name="pcPay_NETOne_CardTypes" type="checkbox" class="clearBorder" value="DISC" <% if pcPay_NETOne_CardTypes="DISC" or (instr(pcPay_NETOne_CardTypes,"DISC,")>0) or (instr(pcPay_NETOne_CardTypes,", DISC")>0) then%>checked<%end if%>> Discover</td>
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
<!-- End of NETOne -->
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
    <td><strong>Sage - formerly Net1<br>
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

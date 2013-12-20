<%
'--- Start ParaData ---
Function gwParaDataEdit()
	call opendb()
	'request gateway variables and insert them into the ParaData table
	query="SELECT pcPay_ParaData_Key FROM pcPay_ParaData Where pcPay_ParaData_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	pcPay_ParaData_Key2=rs("pcPay_ParaData_Key")

	pcPay_ParaData_Key=request.Form("pcPay_ParaData_Key")
	if pcPay_ParaData_Key="" then
		pcPay_ParaData_Key=pcPay_ParaData_Key2
	end if
	set rs=nothing
	
	pcPay_ParaData_TransType=request.Form("pcPay_ParaData_TransType")
	
	pcPay_ParaData_TestMode=request.Form("pcPay_ParaData_TestMode")
	if pcPay_ParaData_TestMode="" then
		pcPay_ParaData_TestMode="0"
	end if

	pcPay_ParaData_CVC=request.Form("pcPay_ParaData_CVC")
	if pcPay_ParaData_CVC="1" then
		pcPay_ParaData_CVC=1
	else
		pcPay_ParaData_CVC=0
	end if

	pcPay_ParaData_AVS=request.Form("pcPay_ParaData_AVS")
	if pcPay_ParaData_AVS="YES" then
		pcPay_ParaData_AVS="1"
	else
		pcPay_ParaData_AVS="0"
	end if
	
	query="UPDATE pcPay_ParaData SET pcPay_ParaData_TransType='"&pcPay_ParaData_TransType&"', pcPay_ParaData_Key='"&pcPay_ParaData_Key&"', pcPay_ParaData_TestMode="&pcPay_ParaData_TestMode&", pcPay_ParaData_CVC="&pcPay_ParaData_CVC&", pcPay_ParaData_AVS="&pcPay_ParaData_AVS&" WHERE pcPay_ParaData_ID=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If

	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ", priceToAdd="&priceToAdd&" ,percentageToAdd="&percentageToAdd&" ,paymentNickName='"&paymentNickName&"' WHERE gwCode=45"
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

Function gwParaData()
	varCheck=1
	pcPay_ParaData_TransType=request.Form("pcPay_ParaData_TransType")

	pcPay_ParaData_TestMode=request.Form("pcPay_ParaData_TestMode")
	if pcPay_ParaData_TestMode="" then
		pcPay_ParaData_TestMode="0"
	end if

	pcPay_ParaData_Key=request.Form("pcPay_ParaData_Key")
	if pcPay_ParaData_Key="" AND pcPay_ParaData_TestMode="0" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Paradata as your payment gateway. <b>""Transaction Key""</b> is a required field.")
	End If

	pcPay_ParaData_CVC=request.Form("pcPay_ParaData_CVC")
	if pcPay_ParaData_CVC="1" then
		pcPay_ParaData_CVC=1
	else
		pcPay_ParaData_CVC=0
	end if

	pcPay_ParaData_AVS=request.Form("pcPay_ParaData_AVS")
	if pcPay_ParaData_AVS="YES" then
		pcPay_ParaData_AVS="1"
	else
		pcPay_ParaData_AVS="0"
	end if

	err.clear
	err.number=0
	call openDb()   

	query="UPDATE pcPay_ParaData SET pcPay_ParaData_TransType='"&pcPay_ParaData_TransType&"', pcPay_ParaData_Key='"&pcPay_ParaData_Key&"', pcPay_ParaData_TestMode="&pcPay_ParaData_TestMode&", pcPay_ParaData_CVC="&pcPay_ParaData_CVC&", pcPay_ParaData_AVS="&pcPay_ParaData_AVS&" WHERE pcPay_ParaData_ID=1;"

	set rs=Server.CreateObject("ADODB.Recordset")  
	set rs=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
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
	query = ""
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'Paradata','gwParaData.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",45,'"&paymentNickName&"')"
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

if request("gwchoice")="45" then
	'//Check for required components and that they are working 
	reDim strComponent(0)
	reDim strClass(0,1)

	'The component names
	strComponent(0) = "ParaData"
	
	'The component class names
	strClass(0,0) = "Paygateway.EClient.1"
	
	isComErr = Cint(0)
	strComErr = Cstr()
	
	For i=0 to UBound(strComponent)
		strErr = IsObjInstalled(i)
		If strErr <> "" Then
			strComErr = strComErr & strErr
			isComErr = 1
		End If
	Next

	if request("mode")="Edit" then
		call opendb()
		query="SELECT pcPay_ParaData_TransType, pcPay_ParaData_Key, pcPay_ParaData_TestMode, pcPay_ParaData_AVS, pcPay_ParaData_CVC FROM pcPay_ParaData WHERE pcPay_ParaData_ID=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If 
		pcPay_ParaData_TransType=rs("pcPay_ParaData_TransType") ' auth or sale
		pcPay_ParaData_Key=rs("pcPay_ParaData_Key") ' private key
		pcPay_ParaData_TestMode=rs("pcPay_ParaData_TestMode")  ' test mode or live mode
		pcPay_ParaData_AVS=rs("pcPay_ParaData_AVS") ' avs "on" or "off"
		pcPay_ParaData_CVC=rs("pcPay_ParaData_CVC") ' cvc "on" or "off"
		set rs=nothing
		call closedb()
		%>
    
		<input type="hidden" name="mode" value="Edit">
    <% end if %>
	<input type="hidden" name="addGw" value="45">
    <tr> 
        <td height="21"><b>Paradata</b> ( <a href="http://www.ParaData.com" target="_blank">Web site</a> )</td>
    </tr>
    <tr> 
	    <td>
         	<% if isComErr = 1 then
			   intDoNotApply = 1 %>
                <table width="100%" border="0" cellspacing="0" cellpadding="4">
                	<tr>
                	  <td><img src="images/red_x.png" alt="Unable to add Paradata" width="12" height="12" /> <strong>Paradata cannot be activated</strong> - the required third party components are not installed and working on the web server.<br />                	    <br /></td>
              	  </tr>
                	<tr>
                    	<td>
                        	<center>
                        	<strong>Required components for Paradata:</strong><br />
                       	  <i><%= strComErr %></i><br /><br />
                        	<input type="button" value="Back" onclick="javascript:history.back()"></center></td>
                  	</tr>
              	</table>
<% else %>
    			<table width="100%" border="0" cellspacing="0" cellpadding="4">
                <tr> 
                    <th colspan="2">Enter Paradata settings</th>
                </tr>
                <% if request("mode")="Edit" then %>
					<% dim pcPay_ParaData_KeyCnt,pcPay_ParaData_KeyEnd,pcPay_ParaData_KeyStart
                    pcPay_ParaData_KeyCnt=(len(pcPay_ParaData_Key)-2)
                    pcPay_ParaData_KeyEnd=right(pcPay_ParaData_Key,2)
                    pcPay_ParaData_KeyStart=""
                    for c=1 to pcPay_ParaData_KeyCnt
                        pcPay_ParaData_KeyStart=pcPay_ParaData_KeyStart&"*"
                    next
                    %>
                    <tr> 
                        <td height="31" colspan="2">Current Account Token:&nbsp;<%=pcPay_ParaData_KeyStart&pcPay_ParaData_KeyEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;Account Token&quot; 
                            is only partially shown on this page. If you need 
                            to edit your account information, please re-enter 
                            your &quot;Account Token&quot; below.</td>
                    </tr>
                <% end if %>
                <tr> 
                    <td width="24%"><div align="right">Account Token:</div></td>
                    <td width="76%"> <div align="left"><input type="text" value="" name="pcPay_ParaData_Key" size="30"></div></td>
                </tr>
                <tr> 
                    <td> <div align="right">Transaction Type:</div></td>
                    <td> 
                    <select name="pcPay_ParaData_TransType">
                    <option value="SALE" selected>Sale</option>
                    <option value="AUTH" <% if pcPay_ParaData_TransType="AUTH" then%>selected<% end if %>>Authorize Only</option>
                    </select></td>
                </tr>
                
                <tr> 
                    <td> <div align="right">Require CVV:</div></td>
                    <td> <input type="radio" class="clearBorder" name="pcPay_ParaData_CVC" value="1" checked>Yes 
                    <input name="pcPay_ParaData_CVC" type="radio" class="clearBorder" value="0" <% if pcPay_ParaData_CVC=0 then%>checked<% end if %>>No
                    <font color="#FF0000">&nbsp;&nbsp;*Required if you are accepting Discover cards.</font></td>
                </tr>
                
                <tr> 
                    <td><div align="right"> 
                    <input name="pcPay_ParaData_TestMode" type="checkbox" class="clearBorder" id="pcPay_ParaData_TestMode" value="1" <% if pcPay_ParaData_TestMode=1 then %>checked<% end if %> />
                    </div></td>
                    <td><b>Enable Test Mode </b>(Credit cards will not be charged)</td>
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
			<% end if %>
        </td>
    </tr>
<!-- END ParaData -->

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

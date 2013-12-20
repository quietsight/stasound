<%
'--- Start PayPal Payments Advanced ---
Function gwPPEdit()
	call opendb()

	PayPalPaymentURL="gwpp.asp"
	PayPalName="PayPal"
	PayPal_Email=request.Form("PayPal_Email")
	pcPay_PayPal_Sandbox=request.Form("PayPal_Sandbox")
	if PayPal_Email = "" then 
		response.redirect"pcConfigurePayment.asp?mode=Edit&id=133&gwchoice=3&msg=You must enter your PayPal email address to activate PayPal Payments Standard."
	end if
	PayPal_Currency=request.Form("PayPal_Currency")
	ppGwcode=3
	if pcPay_PayPal_Sandbox="YES" then
		pcPay_PayPal_Sandbox=1
	else
		pcPay_PayPal_Sandbox=0
	end if
	if pcPay_PayPal_CVC="" then
		pcPay_PayPal_CVC=0
	end if

	query="UPDATE paypal SET Pay_To='"&PayPal_Email&"', URL='https://www.paypal.com/cgi-bin/webscr', PP_Currency='"&PayPal_Currency&"', PP_Sandbox="&pcPay_PayPal_Sandbox&" WHERE ID=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='PayPal',pcPayTypes_ppab=0 WHERE gwCode=3"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If

	set rs=nothing
	call closedb()
end function

Function gwPP()
	varCheck=1
	'request gateway variables and insert them into the pcPay_PayPalAdvanced table
	PayPalPaymentURL="gwpp.asp"
	PayPalName="PayPal"
	PayPal_Email=request.Form("PayPal_Email")
	pcPay_PayPal_Sandbox=request.Form("PayPal_Sandbox")
	if PayPal_Email = "" then 
		response.redirect "pcConfigurePayment.asp?gwchoice=3&msg=You must enter your PayPal email address to activate PayPal Payments Standard."
	end if
	PayPal_Currency=request.Form("PayPal_Currency")
	ppGwcode=3
	if pcPay_PayPal_Sandbox="YES" then
		pcPay_PayPal_Sandbox=1
	else
		pcPay_PayPal_Sandbox=0
	end if
	if pcPay_PayPal_CVC="" then
		pcPay_PayPal_CVC=0
	end if

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
	paymentNickName="PayPal"
			
	err.clear
	err.number=0
			
	call openDb() 

	query="UPDATE paypal SET Pay_To='"&PayPal_Email&"', URL='https://www.paypal.com/cgi-bin/webscr', PP_Currency='"&PayPal_Currency&"', PP_Sandbox="&pcPay_PayPal_Sandbox&" WHERE ID=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName, pcPayTypes_ppab) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'PayPal','gwpp.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",3,'PayPal',0)"
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
				
<% if request("gwchoice")="3" then
	pcConflictIdPayment = 0

	if request("mode")="Edit" then
		call opendb()

		query= "SELECT Pay_To, PP_Currency, PP_Sandbox FROM paypal WHERE ID=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		PayPal_Email=rs("Pay_To")
		PayPal_Currency=rs("PP_Currency")
		PayPal_Sandbox=rs("PP_Sandbox")
		set rs=nothing
		
		query= "SELECT idPayment, pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=3"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=conntemp.execute(query)
		if rs.eof then
			paymentNickName="PayPal"
		else
			pcv_idPayment=rs("idPayment")
			pcv_processOrder=rs("pcPayTypes_processOrder")
			pcv_setPayStatus=rs("pcPayTypes_setPayStatus")
			priceToAdd=rs("priceToAdd")
			percentageToAdd=rs("percentageToAdd")
			paymentNickName="PayPal"
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
    <% else
		'//Check if any other PayPal Services are activated.
		call opendb()
		query= "SELECT idPayment, gwCode FROM payTypes WHERE gwCode=80 OR gwCode=53 OR gwCode=46"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=conntemp.execute(query)
		if NOT rs.eof then
			pcConflictIdPayment = rs("idPayment")
			pcConflictID = rs("gwCode")
			select case pcConflictID
				case "80"
					pcConflictDesc = "PayPal Payments Advanced"
				case "53"
					pcConflictDesc = "Website Payments Pro (UK)"
				case "46"
					pcConflictDesc = "PayPal Payments Pro"
			end select %>
        	<div class="pcCPmessage">
        	  <p>You currently have <strong><%=pcConflictDesc%></strong> active for this store. In order to use <strong>PayPal Payments Standard</strong> you will need to disable <strong><%=pcConflictDesc%></strong>.<br />
        	    <br />
        	  </p>
        	  <p><a href="pcConfigurePayment.asp?mode=Del&id=<%=pcConflictIdPayment%>&gwchoice=<%=pcConflictID%>&activate=3">Disable <%=pcConflictDesc%> and enable PayPal Payments Standard</a></p>
        	  <br />
        	  <p><a href="pcPaymentSelection.asp">Back to payment selection</a><br />
        	    <br />
      	    </p>
            </div>
		<% end if
		set rs = nothing
		call closedb()
	end if %>
    <% if pcConflictIdPayment = 0 then %>
	<input type="hidden" name="addGw" value="3">
    <div class="pcCPmessageSuccess">
        <% if request("mode")="Edit" then %>
            <p><strong>You're editing PayPal Payments Standard</strong>
            <br />
            <br />
        	<p><strong><a href="pcPaymentSelection.asp">Change Payment Option</a></strong></p>
            <br /></p>
        <% else %>
            <p><strong>You've selected PayPal Payments Standard</strong>
            <br />
            Accept credit cards quickly and securely. Buyers are sent to PayPal to pay, and then return to your site when finished. Setup is easy, there are no monthly charges, and buyers don't need a PayPal account. <strong><br>
            <br>
            <a href="https://www.paypal.com/webapps/mpp/paypal-payments-standard" target="_blank">Sign Up and Learn More</a></strong></p>
            <br />
            <p>To start accepting payments, please complete the process below.
            <strong><a href="pcPaymentSelection.asp">Change Payment Option</a></strong></p>
            <br /></p>
        <% end if %>
    </div>
	<table width="100%">
        <tr>
            <td align="left" style="font-size:15px;"><img src="gateways/logos/paypal_logo1.gif" width="253" height="80"></td>
            <td align="left" style="font-size:15px;">&nbsp;</td>
        </tr>
    </table>
    <br>
    	<table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
        <tr>
            <td>
                <div id="CollapsiblePanel1" class="CollapsiblePanel">
                    <div class="CollapsiblePanelTab1">
                        <table width="100%">
                          <tr>
                            <td width="580" class="pcPanelTitle1"><strong>Step 1: Configure Settings - PayPal Payments Standard</strong></td>
                          </tr>
                        </table>
                    </div>
                    <div class="CollapsiblePanelContent">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                                <td colspan="2" class="pcCPspacer"></td>
                            </tr>
                            <tr>
                              <td colspan="2" valign="top">We need this information to work with PayPal so that payments can be sent to your account. You must have a PayPal business account to use PayPal Payments Standard. If you don't have an account, sign up for one now. <a href="https://www.paypal.com/webapps/mpp/paypal-payments-standard" target="_blank">Sign up now</a><br />                                <br /></td>
                            </tr>
                            <tr> 
                                <td width="127" nowrap="nowrap"><div align="right">PayPal  ID/Email: </div></td>
                                <td><input type="text" value="<%=PayPal_Email%>" name="PayPal_Email" size="30" maxlength="50" autocomplete='off'/> 

                                <a href="JavaScript:win('helpOnline.asp?ref=802')"> <img src="images/pcv3_infoIcon.gif" alt="More information on this feature" border="0" /></a></td>
                            </tr>
                            <tr> 
                                <td><div align="right">Currency:</div></td>
                                <td>
                                    <select name="PayPal_Currency">
                                        <option value="USD" <% if PayPal_Currency="USD" then%>selected<% end if %>>U.S. Dollars ($)</option>
                                        <option value="AUD" <% if PayPal_Currency="AUD" then%>selected<% end if %>>Australian Dollars ($)</option>
                                        <option value="CAD" <% if PayPal_Currency="CAD" then%>selected<% end if %>>Canadian Dollars (C $)</option>
                                        <option value="CZK" <% if PayPal_Currency="CZK" then%>selected<% end if %>>Czech Koruna</option>
                                        <option value="DKK" <% if PayPal_Currency="DKK" then%>selected<% end if %>>Danish Krone</option>
                                        <option value="EUR" <% if PayPal_Currency="EUR" then%>selected<% end if %>>Euros (€)</option>
                                        <option value="HKD" <% if PayPal_Currency="HKD" then%>selected<% end if %>>Hong Kong Dollar</option>
                                        <option value="HUF" <% if PayPal_Currency="HUF" then%>selected<% end if %>>Hungarian Forint</option>
                                        <option value="ILS" <% if PayPal_Currency="ILS" then%>selected<% end if %>>Israeli New Shekel</option>
                                        <option value="JPY" <% if PayPal_Currency="JPY" then%>selected<% end if %>>Yen (¥)</option>
                                        <option value="MXN" <% if PayPal_Currency="MXN" then%>selected<% end if %>>Mexican Peso</option> 
                                        <option value="NOK" <% if PayPal_Currency="NOK" then%>selected<% end if %>>Norwegian Krone</option>
                                        <option value="NZD" <% if PayPal_Currency="NZD" then%>selected<% end if %>>New Zealand Dollar</option>
                                        <option value="PHP" <% if PayPal_Currency="PHP" then%>selected<% end if %>>Philippine Peso</option> 
                                        <option value="PLN" <% if PayPal_Currency="PLN" then%>selected<% end if %>>Polish Zloty</option>
                                        <option value="GBP" <% if PayPal_Currency="GBP" then%>selected<% end if %>>Pounds Sterling (£)</option>											
                                        <option value="SGD" <% if PayPal_Currency="SGD" then%>selected<% end if %>>Singapore Dollar</option>
                                        <option value="SEK" <% if PayPal_Currency="SEK" then%>selected<% end if %>>Swedish Krona</option>                                   
                                        <option value="CHF" <% if PayPal_Currency="CHF" then%>selected<% end if %>>Swiss Franc</option>     
                                        <option value="TWD" <% if PayPal_Currency="TWD" then%>selected<% end if %>>Taiwan New Dollar</option>    
                                        <option value="THB" <% if PayPal_Currency="THB" then%>selected<% end if %>>Thai Baht</option>   
                                    </select>						

                                    </td>
                            </tr>
                            <tr> 
                                <td> <div align="right"> 
                                <input name="PayPal_Sandbox" type="checkbox" class="clearBorder" value="YES" <% if PayPal_Sandbox=1 then%>checked<% end if %>>
                                </div></td>
                                <td><b>Enable Test Mode </b>(Credit cards will not be charged) <a href="JavaScript:win('helpOnline.asp?ref=802')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature" border="0" /></a></td>
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
                            <td width="580" class="pcPanelTitle1"><strong>Step 2: Order Processing: Order Status and Payment Status</strong></td>
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
                //-->
                </script>
            </td>
        </tr>
    </table>
	<% end if
end if %>
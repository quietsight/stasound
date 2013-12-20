<% dim pcv_BeagleNotAvailable
pcv_BeagleNotAvailable=1

'//Check if Beagle Field Exists
on error resume next
err.clear
call opendb()
query="SELECT * FROM eWay;"
set rstemp=Server.CreateObject("ADODB.Recordset")     
set rstemp=conntemp.execute(query)
eWay_BeagleActive=rstemp("eWayBeagleActive")
if err.number<>0 then
	pcv_BeagleNotAvailable=0
end if
set rstemp=nothing
call closedb()

'---Start eWay---
Function gwEwayEdit()
	call opendb()
	
	pcv_BeagleNotAvailable=1
	
	query="SELECT * FROM eWay;"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	eWay_BeagleActive=rstemp("eWayBeagleActive")
	if err.number<>0 then
		pcv_BeagleNotAvailable=0
	end if

	'request gateway variables and insert them into the eWay table
	query="SELECT eWayCustomerId, eWayPostMethod FROM eWay WHERE eWayID=1;"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	eWayCustomerId2=rstemp("eWayCustomerId")
	eWayCustomerId=request.Form("eWayCustomerId")
	If eWayCustomerId="" then
		eWayCustomerId=eWayCustomerId2
	end if
	eWayPostMethod2=rstemp("eWayPostMethod")
	eWayPostMethod=request.Form("eWayPostMethod")
	If eWayPostMethod="" then
		eWayPostMethod=eWayPostMethod2
	end if
	eWayTestmode=request.Form("eWayTestmode")
	if eWayTestmode="YES" then
		eWayTestmode="1"
	else
		eWayTestmode="0"
	end if
	eWay_CVV = request.form("eWay_CVV")
	if pcv_BeagleNotAvailable=1 then
		eWay_BeagleActive = request.form("eWay_BeagleActive")
	end if
	query="UPDATE eWay SET eWayCustomerId='"&eWayCustomerId&"', eWayPostMethod='"&eWayPostMethod&"', eWayTestmode="&eWayTestmode&",eWayCVV=" & eWay_CVV
	if pcv_BeagleNotAvailable=1 then
		query=query &", eWayBeagleActive=" & eWay_BeagleActive
	end if
	query=query &" WHERE eWayID=1;"

	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=31"

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

Function gwEway()
	call opendb()
	pcv_BeagleNotAvailable=1
	
	query="SELECT * FROM eWay;"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	eWay_BeagleActive=rstemp("eWayBeagleActive")
	if err.number<>0 then
		pcv_BeagleNotAvailable=0
	end if

	varCheck=1
	'request gateway variables and insert them into the eWay table
	eWayCustomerId=request.Form("eWayCustomerId")
	eWayPostMethod=request.Form("eWayPostMethod")
	eWayTestmode=request.Form("eWayTestmode")
	eWay_CVV = request.form("eWay_CVV")
	eWay_BeagleActive = request.form("eWay_BeagleActive")
	if eWayTestmode="YES" then
		eWayTestmode="1"
	else
		eWayTestmode="0"
	end if
	If eWayCustomerId="" then
		call closedb()
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add eWay as your payment gateway. <b>""Customer ID""</b> is a required field.")
	End If
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

	query="UPDATE eWay SET eWayCustomerId='"&eWayCustomerId&"',eWayPostMethod='"&eWayPostMethod&"',eWayTestmode="&eWayTestmode&",eWayCVV=" & eWay_CVV
	if pcv_BeagleNotAvailable=1 then
		query=query &",eWayBeagleActive=" & eWay_BeagleActive
	end if
	query=query &" WHERE eWayID=1"
	set rs=Server.CreateObject("ADODB.Recordset")  
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'eWay','gwEway.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",31,'"&paymentNickName&"')"
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

if request("gwchoice")="31" then
	tmp_id=request("id")
	tmp_mode=request("mode")

	'Check to see if fields exists in table, if not, add
	err.clear
	call openDb()
	query="SELECT eWayCVV FROM eWay"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		set rs=nothing
		call closedb()
		response.redirect "upddbEway.asp?mode="&tmp_mode&"&id="&tmp_id
	else
		set rs=nothing
		call closedb()
	end if
	
	if request("mode")="Edit" then
        call opendb()
		query="SELECT eWayCustomerId, eWayPostMethod, eWayTestmode, eWayCVV"
		if pcv_BeagleNotAvailable=1 then
			query=query&", eWayBeagleActive"
		end if
		query=query&" FROM eWay WHERE eWayID=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		eWayCustomerId=rs("eWayCustomerId")
		eWayPostMethod=rs("eWayPostMethod")
		eWayTestmode=rs("eWayTestmode")
		eWay_CVV = rs("eWayCVV")
		if pcv_BeagleNotAvailable=1 then
			eWay_BeagleActive = rs("eWayBeagleActive")
		end if
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=31"
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
	<input type="hidden" name="addGw" value="31">
<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/eway_logo.png" width="180" height="83"></td>
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
    <a href="http://www.eway.com.au/" target="_blank">eWay Website</a></strong><br />
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
                <tr>
                  <td align="right" valign="top">&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                    <td align="right" valign="top"><input name="eWayPostMethod" type="radio" class="clearBorder" value="SHARED" checked></td>
                    <td><strong>Shared Payment</strong><b><br>
                </b>Process credit card payments via eWAY's own secure 
                Shared Payment Page in real time. You POST purchase 
                information from your web site to the eWAY secured site.</td>
                </tr>
                <tr> 
                    <td align="right" valign="top"><input name="eWayPostMethod" type="radio" class="clearBorder" value="XML" <% if eWayPostMethod="XML" then%>checked<% end if %>></td>
                    <td valign="top"><b>XML Payment<font color="#FF0000"> 
                        </font></b><em>(recommended)</em><b><br>
                        </b>Process credit card payments directly through your 
                        own website in real time. Using the eWAY XML Solution, 
                        your web site appears as the payment gateway, with the 
                        transactions POSTed in the background. </td>
                </tr>
                <% if request("mode")="Edit" then %>
					<% dim eWayCustomerIdCnt,eWayCustomerIdEnd,eWayCustomerIdStart
                    eWayCustomerIdCnt=(len(eWayCustomerId)-2)
                    eWayCustomerIdEnd=right(eWayCustomerId,2)
                    eWayCustomerIdStart=""
                    for c=1 to eWayCustomerIdCnt
                        eWayCustomerIdStart=eWayCustomerIdStart&"*"
                    next %>
                    <tr> 
                        <td height="31" colspan="2">Current InternetSecure Merchant Number:&nbsp;<%=eWayCustomerIdStart&eWayCustomerIdEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;eWay Customer 
                            ID&quot; is only partially shown on this page. If 
                            you need to edit your account information, please 
                            re-enter your &quot;Customer ID&quot; below.</td>
                    </tr>
                <% end if %>
                <tr> 
                    <td> <div align="right">Customer ID:</div></td>
                    <td width="1203"> <input type="text" name="eWayCustomerId" size="20"></td>
                </tr>
                <tr> 
                    <td> <div align="right"> 
                            <input name="eWayTestmode" type="checkbox" class="clearBorder" value="YES" <% if eWayTestmode=1 then%>checked<% end if%>> 
                        </div></td>
                    <td><b>Enable Test Mode </b>(Credit cards will not be charged)</td>
                </tr>
				<TR>
				  <td><div align="right">Real-Time CVN:</div></td>
			      <td>
				        <input type="radio" class="clearBorder" name="eWay_CVV" value="1" <%if eWay_CVV=1 then%> Checked <%end if %> /> Yes 
                        <input name="eWay_CVV" type="radio" class="clearBorder" value="0" <%if eWay_CVV=0 then%> Checked <%end if %> /> 
                       No
(Real-Time CVN is an optional eWay subscription) </td>
				</TR>
                <% if pcv_BeagleNotAvailable=1 then %> 
                    <TR>
                      <td nowrap="nowrap"><div align="right">Beagle (Geo-IP Anti Fraud):</div></td>
                      <td>
                            <input type="radio" class="clearBorder" name="eWay_BeagleActive" value="1" <%if eWay_BeagleActive=1 then%> Checked <%end if %> /> Yes 
                            <input name="eWay_BeagleActive" type="radio" class="clearBorder" value="0" <%if eWay_BeagleActive=0 then%> Checked <%end if %> /> 
                           No
    (Beagle Fraud Prevention is an optional eWay subscription) </td>
                    </TR> 
                <% else %>
                    <TR>
                      <td nowrap="nowrap" valign="top"><div align="right">Beagle (Geo-IP Anti Fraud):</div></td>
                      <td>
                        <input name="eWay_BeagleActive" type="hidden" value="0" />
                        Beagle Fraud Prevention is an optional eWay subscription. This feature is not available in ProductCart until you update your database. <a href="upddbEway.asp">Click here to update your database now.</a></td>
                    </TR> 
                <% end if %>
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
                        <td class="pcPanelTitle1"><strong>Step 2: You have the option to charge a processing fee for this payment option.</strong></td>
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

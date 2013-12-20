<%
'---Start FastCharge---
Function gwfastchargeEdit()
	call opendb()
	query="SELECT pcPay_FAC_ATSID FROM pcPay_FastCharge WHERE pcPay_FAC_Id=1;"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	pcPay_FAC_ATSID2=rstemp("pcPay_FAC_ATSID")
	pcPay_FAC_ATSID=request.Form("pcPay_FAC_ATSID")
	If pcPay_FAC_ATSID="" then
		pcPay_FAC_ATSID=pcPay_FAC_ATSID2
	end if
	pcPay_FAC_ATSID=enDeCrypt(pcPay_FAC_ATSID, scCrypPass)
	pcPay_FAC_TransType=request.Form("pcPay_FAC_TransType")
	if pcPay_FAC_TransType="" then
	pcPay_FAC_TransType="0"
	end if
	pcPay_FAC_CVVEnabled=request.Form("pcPay_FAC_CVVEnabled")
	if pcPay_FAC_CVVEnabled="" then
	pcPay_FAC_CVVEnabled="0"
	end if
	pcPay_FAC_Checking=request.Form("pcPay_FAC_Checking")
	pcPay_FAC_CheckPending=request.Form("pcPay_FAC_CheckPending")
	If pcPay_FAC_ATSID="" then
		call closedb()
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Fast Charge as your payment gateway. <b>""ATSID""</b> is a required field.")
	End If
	
	paymentNickName2=replace(request.Form("paymentNickName2"),"'","''")
	if paymentNickName2="" then
		paymentNickName2="eCheck"
	end if
	
	query="UPDATE pcPay_FastCharge SET pcPay_FAC_ATSID='"&pcPay_FAC_ATSID&"',pcPay_FAC_TransType=" & pcPay_FAC_TransType & ",pcPay_FAC_CVV="&pcPay_FAC_CVVEnabled&",pcPay_FAC_Checking="&pcPay_FAC_Checking&",pcPay_FAC_CheckPending="&pcPay_FAC_CheckPending&" WHERE pcPay_FAC_Id=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ", priceToAdd="&priceToAdd&" ,percentageToAdd="&percentageToAdd&" ,paymentNickName='"&paymentNickName&"' WHERE gwCode=37"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	if pcPay_FAC_Checking="1" then
	
		query="SELECT paymentDesc FROM payTypes where gwCode=38"
		set rs=conntemp.execute(query)
		
		if not rs.eof then
			query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ", priceToAdd="&priceToAdd&" ,percentageToAdd="&percentageToAdd&" ,paymentNickName='"&paymentNickName2&"' WHERE gwCode=38"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=conntemp.execute(query)
		else
			query="INSERT INTO payTypes (paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES ('FastCharge eCheck','gwFastChargeCheck.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",38,'"&paymentNickName2&"')"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=conntemp.execute(query)
		end if
	
	else
	
		query="DELETE FROM payTypes where gwCode=38"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=conntemp.execute(query)
	
	end if
	
	set rs=nothing
	call closedb()
end function

Function gwfastcharge()
	varCheck=1

	'request gateway variables and insert them into the fasttransact table
	pcPay_FAC_ATSID=request.Form("pcPay_FAC_ATSID")
	pcPay_FAC_ATSID=enDeCrypt(pcPay_FAC_ATSID, scCrypPass)

	pcPay_FAC_TransType=request.Form("pcPay_FAC_TransType")
	if pcPay_FAC_TransType="" then
		pcPay_FAC_TransType="0"
	end if
	pcPay_FAC_CVVEnabled=request.Form("pcPay_FAC_CVVEnabled")
	if pcPay_FAC_CVVEnabled="" then
		pcPay_FAC_CVVEnabled="0"
	end if
	pcPay_FAC_Checking=request.Form("pcPay_FAC_Checking")
	pcPay_FAC_CheckPending=request.Form("pcPay_FAC_CheckPending")

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
	paymentNickName2=replace(request.Form("paymentNickName2"),"'","''")
	if paymentNickName2="" then
		paymentNickName2="Check"
	end if

	call opendb()
	query="UPDATE pcPay_FastCharge SET pcPay_FAC_ATSID='"&pcPay_FAC_ATSID&"',pcPay_FAC_TransType=" & pcPay_FAC_TransType & ",pcPay_FAC_CVV="&pcPay_FAC_CVVEnabled&",pcPay_FAC_Checking="&pcPay_FAC_Checking&",pcPay_FAC_CheckPending="&pcPay_FAC_CheckPending&" WHERE pcPay_FAC_Id=1;"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'FastCharge','gwfastcharge.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",37,'"&paymentNickName&"')"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	set rs=nothing
	
	if pcPay_FAC_Checking="1" then
		query="INSERT INTO payTypes (pcPayTypes_processOrder, pcPayTypes_setPayStatus, paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'FastCharge eCheck','gwFastChargeCheck.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",38,'"&paymentNickName2&"')"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	end if

	call closedb()
end function

'37 = credit
'38 = checks

if request("gwchoice")="37" then
	intDoNotApply = 0
	
	'//Check for required components and that they are working 
	reDim strComponent(0)
	reDim strClass(0,1)

	'The component names
	strComponent(0) = "FastCharge"
	
	'The component class names
	strClass(0,0) = "ATS.SecurePost"
	
	isComErr = Cint(0)
	strComErr = Cstr("")
	
	For i=0 to UBound(strComponent)
		strErr = IsObjInstalled(i)
		If strErr <> "" Then
			strComErr = strComErr & strErr
			isComErr = 1
		End If
	Next

	if request("mode")="Edit" then
		call opendb()
		query="Select pcPay_FAC_ATSID,pcPay_FAC_TransType,pcPay_FAC_CVV,pcPay_FAC_Checking,pcPay_FAC_CheckPending from pcPay_FastCharge where pcPay_FAC_Id=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcPay_FAC_ATSID=rs("pcPay_FAC_ATSID")
		pcPay_FAC_TransType=rs("pcPay_FAC_TransType")
		pcPay_FAC_ATSID=enDeCrypt(pcPay_FAC_ATSID, scCrypPass)
		pcPay_FAC_CVVEnabled=rs("pcPay_FAC_CVV")
		pcPay_FAC_Checking=rs("pcPay_FAC_Checking")
		pcPay_FAC_CheckPending=rs("pcPay_FAC_CheckPending")
		set rs=nothing
				
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=37"
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


		query= "SELECT paymentNickName FROM payTypes WHERE gwCode=38"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=conntemp.execute(query)
		if rs.eof then
			paymentNickName2="Check"
		else
			paymentNickName2=rs("paymentNickName")
		end if
		set rs=nothing
		call closedb()
		%>
    
		<input type="hidden" name="mode" value="Edit">
    <% end if %>
	<input type="hidden" name="addGw" value="37">
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/FastCharge.JPG" width="221" height="71"></td>
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
    <a href="http://www.fastcharge.com/">FastCharge Website</a></strong><br />
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
                            <%
							if isComErr = 1 then
                                intDoNotApply = 1 %>
                                    <tr>
                                      <td colspan="2">&nbsp;</td>
                                    </tr>
                                    <tr>
                                      <td colspan="2"><img src="images/red_x.png" alt="Unable to add Paradata" width="12" height="12" /> <strong>FastCharge cannot be activated</strong> - the required third party components are not installed and working on the web server.<br />                	    
                                      <br /></td>
                                    </tr>
                                    <tr>
                                        <td colspan="2">
                                            <center>
                                            <strong>Required components for FastCharge:</strong><br />
                                          <i><%= strComErr %></i><br /><br />
                                        <input type="button" value="Back" onclick="javascript:history.back()"></center></td>
                                    </tr>
                            <% else %>
                <% if request("mode")="Edit" then %>
					<% dim pcPay_FAC_ATSIDCnt,pcPay_FAC_ATSIDEnd,pcPay_FAC_ATSIDStart
                    pcPay_FAC_ATSIDCnt=(len(pcPay_FAC_ATSID)-2)
                    pcPay_FAC_ATSIDEnd=right(pcPay_FAC_ATSID,2)
                    pcPay_FAC_ATSIDStart=""
                    for c=1 to pcPay_FAC_ATSIDCnt
                        pcPay_FAC_ATSIDStart=pcPay_FAC_ATSIDStart&"*"
                    next %>
                    <tr> 
                        <td height="31" colspan="2">Current FastCharge ATSID:&nbsp;<%=pcPay_FAC_ATSIDStart&pcPay_FAC_ATSIDEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;ATSID&quot; 
                            is only partially shown on this page. If you need 
                            to edit your account information, please re-enter 
                            your &quot;ATSID&quot; below.</td>
                    </tr>
                <% end if %>
                <tr>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                    <td width="21%"> <div align="right">ATSID :</div></td>
                    <td width="79%"> <input type="text" name="pcPay_FAC_ATSID" size="20"> 
                      (Online Commerce Suite Account ID) </td>
                </tr>
                
                <tr> 
                    <td nowrap="nowrap"> <div align="right">Transaction Type:</div></td>
                    <td> <select name="pcPay_FAC_TransType">
                            <option value="1" selected>Sale</option>
                            <option value="0" <%if pcPay_FAC_TransType="0" then%>selected<%end if%>>Authorize Only</option>
                    </select>  </td>
                </tr>
                <tr> 
                    <td> <div align="right">Require CVV:</div></td>
                    <td> <input type="radio" class="clearBorder" name="pcPay_FAC_CVVEnabled" value="1" checked>
                        Yes 
                        <input type="radio" class="clearBorder" name="pcPay_FAC_CVVEnabled" value="0" <%if pcPay_FAC_CVVEnabled="0" then%>checked<%end if%>>
                        No</td>
                </tr>
                <tr> 
                    <td><div align="right">Accept Checks:</div></td>
                    <td><input type="radio" class="clearBorder" name="pcPay_FAC_Checking" value="1" checked>
                        Yes 
                        <input name="pcPay_FAC_Checking" type="radio" class="clearBorder" value="0" <% if pcPay_FAC_Checking="0" then%>checked<%end if%>>
                        No</td>
                </tr>
                <tr> 
                    <td colspan="2">All electronic check orders are considered 
                        &quot;Processed&quot;, as the order amount is always 
                        debited to the customer's bank account. If for any 
                        reasons you would like electronic check orders to 
                        be considered &quot;Pending&quot;, use this option. 
                        Should electronic check orders be considered &quot;Pending&quot;? 
                        <input type="radio" class="clearBorder" name="pcPay_FAC_CheckPending" value="1" checked>
                        Yes 
                        <input name="pcPay_FAC_CheckPending" type="radio" class="clearBorder" value="0" <%if pcPay_FAC_CheckPending="0" then%>checked<%end if%>>
                        No</td>
                </tr>
                            <% end if %>
                  <tr>
                    <td>&nbsp;</td>
                    <td class="pcSubmenuContent">&nbsp;</td>
                  </tr>
                </table>
            </div>
        </div>
                <% if intDoNotApply = 0 then %>
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
                        		<td width="10%" nowrap="nowrap"><div align="left">eCheck&nbsp;&nbsp;Payment Name:&nbsp;</div></td>
                        		<td><input name="paymentNickName2" value="<%=paymentNickName2%>" size="35" maxlength="255"></td>
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
                <% end if %>
        <script type="text/javascript">
        <!--
                var CollapsiblePanel1 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel1", {contentIsOpen:true});
                <% if intDoNotApply = 0 then %>
                var CollapsiblePanel2 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel2", {contentIsOpen:false});;
                var CollapsiblePanel3 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel3", {contentIsOpen:false});
                var CollapsiblePanel4 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel4", {contentIsOpen:false});
                <% end if %>
                //-->
        </script>
    </td>
</tr>
</table>
<!-- New View End --><% end if %>

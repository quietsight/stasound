<%
'---Start TCLink---
Function gwtclinkEdit()
	call opendb()
	'request gateway variables and insert them into the TCLink table
	query="SELECT TCLinkid,TCLinkPassword,TCTestmode,CVV,TranType,TCLinkCheck, TCLinkCheckPending, TCLinkecheck, TCCurcode, avs FROM tclink WHERE idTCLink=1"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	TCLinkPassword2=rstemp("TCLinkPassword")
	'decrypt
	TCLinkPassword2=enDeCrypt(TCLinkPassword2, scCrypPass)
	TCLinkid=request.Form("TCLinkid")
	TCTestmode=request.Form("TCTestmode")
	if TCTestmode="" then
		TCTestmode="0"
	end if

	TCCurcode2=rstemp("TCCurcode")
	TCCurcode=request.Form("TCCurcode")
	If TCCurcode="" then
		TCCurcode=TCCurcode2
	end if
	
	TCLinkPassword=request.Form("TCLinkPassword")
	if TCLinkPassword="" then
		TCLinkPassword=TCLinkPassword2
	end if
	'encrypt
	TCLinkPassword=enDeCrypt(TCLinkPassword, scCrypPass)
	CVV=request.Form("CVV")
	TCLinkAVS=request.Form("TCLinkAVS")
	if TCLinkAVS="1" then
		TCLinkAVS=1
	else
		TCLinkAVS=0
	end if
	TranType2=rstemp("TranType")
	TranType=request.Form("TranType")
	If TranType="" then
		TranType=TranType2
	end if
	
	TCLinkCheck=request.Form("TCLinkCheck")
	TCLinkCheckPending=request.Form("TCLinkCheckPending")
	paymentNickName2=replace(request.Form("paymentNickName2"),"'","''")
	query="UPDATE tclink SET TCLinkid='"&TCLinkid&"',TCLinkPassword='"&TCLinkPassword&"',TCTestmode="&TCTestmode&",CVV="&CVV&",TCCurcode='"&TCCurcode&"',TranType='"&TranType&"',TCLinkCheck="&TCLinkCheck&",TCLinkCheckPending="&TCLinkCheckPending&",avs="&TCLinkAVS&" WHERE idTCLink=1;"
	
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",sslURL='gwtclink.asp', priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=24"

	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
		if TCLinkCheck="1" then
		query="SELECT * FROM payTypes WHERE gwCode=25"
		set rstemp=server.CreateObject("ADODB.RecordSet")
		set rstemp=conntemp.execute(query)
		if rstemp.eof then
			query="INSERT INTO payTypes (paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES ('TCLink eCheck','gwtclinkCheck.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",25,'"&paymentNickName2&"')"
		else
			query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",paymentNickName='"&paymentNickName2&"' WHERE gwCode=25"
		end if
	else
		query="DELETE FROM payTypes WHERE gwCode=25"
	end if
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

Function gwtclink()
	varCheck=1
	'request gateway variables and insert them into the tclink table
	TCLinkid=request.Form("TCLinkid")
	TCLinkPassword=request.Form("TCLinkPassword")
	'encrypt
	TCLinkPassword=enDeCrypt(TCLinkPassword, scCrypPass)
	CVV=request.Form("CVV")
	TCLinkCheck=request.Form("TCLinkCheck")
	TCLinkCheckPending=request.Form("TCLinkCheckPending")
	TCLinkeheck=request.Form("TCLinkeheck")
	TCCurcode=request.Form("TCCurcode")
	TCLinkAVS=request.Form("TCLinkAVS")
	TCTestmode=request.Form("TCTestmode")
	if TCTestmode="" then
		TCTestmode="0"
	end if
	TranType=request.Form("TranType")
	
	if TCLinkAVS="" then
		TCLinkAVS=0
	else
		TCLinkAVS=1
	end if

	if NOT isNumeric(CVV) or CVV="" then
		CVV=0
	end if
	If TCLinkid="" then
		call closedb()
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Concord as your payment gateway. <b>""CustID""</b> is a required field.")
	End If
	TCLinkPassword=request.Form("TCLinkPassword")
	If TCLinkPassword="" then
		call closedb()
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Concord as your payment gateway. <b>""Password""</b> is a required field.")
	End If
	'encrypt
	TCLinkPassword=enDeCrypt(TCLinkPassword, scCrypPass)
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
	end if
	paymentNickName2=replace(request.Form("paymentNickName2"),"'","''")
	if paymentNickName2="" then
		paymentNickName2="Check"
	End If
	
	err.clear
	err.number=0
	call openDb() 

	query="UPDATE tclink SET TCLinkid='"&TCLinkid&"',TCLinkPassword='"&TCLinkPassword&"',CVV="&CVV&",TCTestmode="&TCTestmode&",TCCurcode='"&TCCurcode&"', TranType='"&TranType&"', TCLinkCheck="&TCLinkCheck&", TCLinkCheckPending="&TCLinkCheckPending&",TCLinkecheck='"&paymentNickName2&"',avs="&TCLinkAVS&" WHERE idTCLink=1"
	
	set rs=Server.CreateObject("ADODB.Recordset")  

	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'TCLink','gwtclink.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",24,'"&paymentNickName&"')"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if TCLinkCheck="1" then
		query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'TCLink eCheck','gwtclinkCheck.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","& percentageToAdd &",25,'"&paymentNickName2&"')"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	end if
	set rs=nothing
	call closedb()
end function

if request("gwchoice")="24" then
	if request("mode")="Edit" then
		call opendb()
		query= "SELECT TCLinkid,TCLinkPassword,TCTestmode,CVV,TranType,TCLinkCheck, TCLinkCheckPending, TCLinkecheck, avs, TCCurcode FROM tclink WHERE idTCLink=1"

		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		TCLinkid=rs("TCLinkid")
		TCLinkPassword=rs("TCLinkPassword")
		'decrypt
		TCLinkPassword=enDeCrypt(TCLinkPassword, scCrypPass)
		TCTestmode=rs("TCTestmode")
		TCCurcode=rs("TCCurcode")
		CVV=rs("CVV")
		TCLinkAVS=rs("avs")
		TranType=rs("TranType")
		TCLinkCheck=rs("TCLinkCheck")
		TCLinkCheckPending=rs("TCLinkCheckPending")
		TCLinkecheck=rs("TCLinkecheck")
		set rs=nothing
		
		query= "SELECT paymentNickName FROM payTypes WHERE gwCode=25"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=conntemp.execute(query)
		if rs.eof then
			paymentNickName2="Check"
		else
			paymentNickName2=rs("paymentNickName")
		end if

		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=24"
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
	<input type="hidden" name="addGw" value="24">
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/trustcommerce_logo.jpg" width="564" height="90"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
<table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
<tr>
    <td><strong><br>
    </strong>TrustCommerce payment acceptance solutions serve card not present environments using the TCLink integration<strong><br>
    <br>
    <a href="http://www.trustcommerce.com/" target="_blank">TrustCommerce Website</a></strong><br />
<br />
</td>
</tr>
<tr>
    <td>
        <div id="CollapsiblePanel1" class="CollapsiblePanel">
            <div class="CollapsiblePanelTab1">
                <table width="100%">
                  <tr>
                    <td width="580" class="pcPanelTitle1">Step 1: Configure Account - TrustCommerce TCLink</td>
                  </tr>
                </table>
            </div>
            <div class="CollapsiblePanelContent">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <% if request("mode")="Edit" then %>
					<% dim TCLinkIDCnt,TCLinkIDEnd,TCLinkIDStart
                    TCLinkIDCnt=(len(TCLinkID)-2)
                    TCLinkIDEnd=right(TCLinkID,2)
                    TCLinkIDStart=""
                    for c=1 to TCLinkIDCnt
                        TCLinkIDStart=TCLinkIDStart&"*"
                    next %>
                    <tr> 
                        <td height="31" colspan="2">Current Customer ID:&nbsp;<%=TCLinkIDStart&TCLinkIDEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;Customer ID&quot; 
                            are only partially shown on this page. If you need 
                            to edit your account information, please re-enter 
                            your &quot;Customer ID&quot; and &quot;Password&quot; 
                            below.</td>
                    </tr>
                <% end if %>
                <tr>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                    <td width="111"><div align="right">Customer ID:</div></td>
                    <td width="360"> <input type="text" name="TCLinkid" size="20"></td>
                </tr>
                <tr> 
                    <td><div align="right">Password:</div></td>
                    <td><input name="TCLinkPassword" type="password" size="20"></td>
                </tr>
                <tr> 
                    <td> <div align="right">Transaction Type:</div></td>
                    <td> <select name="TranType">
                            <option value="sale" selected>Sale</option>
                            <option value="preauth" <% if TranType="preauth" then%>selected<% end if %>>Authorize Only</option>
                        </select> &nbsp;&nbsp; <input name="TCLinkAVS" type="checkbox" class="clearBorder" value="1" <% if TCLinkAVS=1 then %>checked<% end if %>> Enable AVS Mode <font size=1>(Address Verification Service)</font></td>
                </tr>
                <tr> 
                    <td> <div align="right">Currency Code:</div></td>
                    <td> <input type="text" name="TCCurcode" size="8" value="<%=TCCurcode%>" /> 
                        <font size=1>(Outside the US: Ask TrustCommerce about the code to be used for your currency)</font> </td>
                </tr>
                <tr> 
                    <td><div align="right">Require CVV:</div></td>
                    <td><input type="radio" class="clearBorder" name="CVV" value="1" checked>
                        Yes 
                        <input type="radio" class="clearBorder" name="CVV" value="0" <% if CVV=0 then%>checked<% end if %>>
                        No</td>
                </tr>
                <tr> 
                    <td> <div align="right"> 
                            <input name="TCTestmode" type="checkbox" class="clearBorder" value="1" <% if TCTestmode="1" then %>checked<% end if %> />
                        </div></td>
                    <td><b>Enable Test Mode </b>(Credit cards will not be charged)</td>
                </tr>
                <tr> 
                    <td><div align="right">Accept Checks:</div></td>
                    <td> <input type="radio" class="clearBorder" name="TCLinkCheck" value="1" checked>
                        Yes 
                        <input name="TCLinkCheck" type="radio" class="clearBorder" value="0" <% if TCLinkCheck="0" then%>checked<% end if %>>
                        No</td>
                </tr>
                <tr> 
                    <td colspan="2">All electronic check orders are considered 
                        &quot;Processed&quot;, as the order amount is always 
                        debited to the customer's bank account. If for any reasons 
                        you would like electronic check orders to be considered 
                        &quot;Pending&quot;, use this option. Should electronic 
                        check orders be considered &quot;Pending&quot;? 
                        <input type="radio" class="clearBorder" name="TCLinkCheckPending" value="1" checked>
                        Yes 
                        <input name="TCLinkCheckPending" type="radio" class="clearBorder" value="0" <% if TCLinkCheckPending="0" then%>checked<% end if %>>
                        No</td>
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

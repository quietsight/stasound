<%
'---Start ACHDirect---
Function gwACHEdit()
	call opendb()
	'request gateway variables and insert them into the USAePay table
	query="SELECT pcPay_ACHDirect.pcPay_ACH_MerchantID, pcPay_ACHDirect.pcPay_ACH_PWD, pcPay_ACHDirect.pcPay_ACH_TransType, pcPay_ACHDirect.pcPay_ACH_TestMode, pcPay_ACHDirect.pcPay_ACH_CVV, pcPay_ACHDirect.pcPay_ACH_CardTypes FROM pcPay_ACHDirect WHERE (((pcPay_ACHDirect.pcPay_ACH_Id)=1));"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	pcPay_ACH_MerchantID=rs("pcPay_ACH_MerchantID")
	pcPay_ACH_PWD=rs("pcPay_ACH_PWD")
	
	pcPay_ACH_MerchantID2=pcPay_ACH_MerchantID
	'decrypt
	pcPay_ACH_MerchantID2=enDeCrypt(pcPay_ACH_MerchantID2, scCrypPass)
	pcPay_ACH_MerchantID=request.Form("pcPay_ACH_MerchantID")
	if pcPay_ACH_MerchantID="" then
		pcPay_ACH_MerchantID=pcPay_ACH_MerchantID2
	end if
	'encrypt
	pcPay_ACH_MerchantID=enDeCrypt(pcPay_ACH_MerchantID, scCrypPass)
	
	pcPay_ACH_PWD2=pcPay_ACH_PWD
	'decrypt
	pcPay_ACH_PWD2=enDeCrypt(pcPay_ACH_PWD2, scCrypPass)
	pcPay_ACH_PWD=request.Form("pcPay_ACH_PWD")
	if pcPay_ACH_PWD="" then
		pcPay_ACH_PWD=pcPay_ACH_PWD2
	end if
	'encrypt
	pcPay_ACH_PWD=enDeCrypt(pcPay_ACH_PWD, scCrypPass)
	
	pcPay_ACH_TestMode=request.Form("pcPay_ACH_TestMode")
	if pcPay_ACH_TestMode="" then
		pcPay_ACH_TestMode="0"
	end if
	
	pcPay_ACH_TransType=request.Form("pcPay_ACH_TransType")
	pcPay_ACH_CVV=request.Form("pcPay_ACH_CVV")
	pcPay_ACH_CardTypes=request.Form("pcPay_ACH_CardTypes")
	
	query="UPDATE pcPay_ACHDirect SET pcPay_ACH_MerchantID='"&pcPay_ACH_MerchantID&"',pcPay_ACH_PWD='"&pcPay_ACH_PWD&"',pcPay_ACH_TransType='" & pcPay_ACH_TransType & "',pcPay_ACH_TestMode="&pcPay_ACH_TestMode&",pcPay_ACH_CVV="&pcPay_ACH_CVV&",pcPay_ACH_CardTypes='"&pcPay_ACH_CardTypes&"' WHERE pcPay_ACH_Id=1"
	
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=39"
	
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

Function gwACH()
	varCheck=1
	'request gateway variables and insert them into the CyberSource table
	pcPay_ACH_MerchantID=request.Form("pcPay_ACH_MerchantID")
	pcPay_ACH_PWD=request.Form("pcPay_ACH_PWD")
	pcPay_ACH_TransType=request.Form("pcPay_ACH_TransType")
	pcPay_ACH_TestMode=request.Form("pcPay_ACH_TestMode")
	if pcPay_ACH_TestMode="" then
		pcPay_ACH_TestMode="0"
	end if
	pcPay_ACH_CVV=request.Form("pcPay_ACH_CVV")
	pcPay_ACH_CardTypes=request.Form("pcPay_ACH_CardTypes")

	If pcPay_ACH_MerchantID="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add ACH Direct as your payment gateway. <b>""Merchant ID""</b> is a required field.")
	End If
	'encrypt
	pcPay_ACH_MerchantID=enDeCrypt(pcPay_ACH_MerchantID, scCrypPass)
	If pcPay_ACH_PWD="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add ACH Direct as your payment gateway. <b>""Password""</b> is a required field.")
	End If
	'encrypt
	pcPay_ACH_PWD=enDeCrypt(pcPay_ACH_PWD, scCrypPass)

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

	query="UPDATE pcPay_ACHDirect SET pcPay_ACH_MerchantID='"&pcPay_ACH_MerchantID&"',pcPay_ACH_PWD='"&pcPay_ACH_PWD&"',pcPay_ACH_TransType='" & pcPay_ACH_TransType & "',pcPay_ACH_TestMode="&pcPay_ACH_TestMode&",pcPay_ACH_CVV="&pcPay_ACH_CVV&",pcPay_ACH_CardTypes='"&pcPay_ACH_CardTypes&"' WHERE pcPay_ACH_Id=1"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'ACHDirect','gwACH.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",39,'"&paymentNickName&"')"
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
<% if request("gwchoice")="39" then
	intDoNotApply = 0
	
	'//Check for required components and that they are working 
	reDim strComponent(0)
	reDim strClass(0,1)

	'The component names
	strComponent(0) = "ACHDirect"
	
	'The component class names
	strClass(0,0) = "SendPmt.clsSendPmt"
	
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
		query="SELECT pcPay_ACHDirect.pcPay_ACH_MerchantID, pcPay_ACHDirect.pcPay_ACH_PWD, pcPay_ACHDirect.pcPay_ACH_TransType, pcPay_ACHDirect.pcPay_ACH_TestMode, pcPay_ACHDirect.pcPay_ACH_CVV, pcPay_ACHDirect.pcPay_ACH_CardTypes FROM pcPay_ACHDirect WHERE (((pcPay_ACHDirect.pcPay_ACH_Id)=1));"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcPay_ACH_MerchantID=rs("pcPay_ACH_MerchantID")
		pcPay_ACH_MerchantId=enDeCrypt(pcPay_ACH_MerchantId, scCrypPass)
		pcPay_ACH_PWD=rs("pcPay_ACH_PWD")
		pcPay_ACH_PWD=enDeCrypt(pcPay_ACH_PWD, scCrypPass)
		pcPay_ACH_TransType=rs("pcPay_ACH_TransType")
		pcPay_ACH_TestMode=rs("pcPay_ACH_TestMode")
		pcPay_ACH_CVV=rs("pcPay_ACH_CVV")
		pcPay_ACH_CardTypes=rs("pcPay_ACH_CardTypes")
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=39"
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
	<input type="hidden" name="addGw" value="39">
    <table width="100%">
        <tr>
            <td align="left" style="font-size:15px;"><img src="gateways/logos/ach.gif" width="211" height="101"></td>
            <td align="left" style="font-size:15px;">&nbsp;</td>
        </tr>
    </table>
    <br>
    <table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
        <tr>
            <td><strong>ACH Direct's Payments<em>Gateway</em><sup>TM</sup> Advanced Gateway Interface integration<br>
            <br>
            </strong>ACH Direct's Payments<em>Gateway</em> is a  high-capacity modular payment processing platform designed for  maximum flexibility and availability.<strong><br>
            <br>
          <a href="http://www.achdirect.com/" target="_blank">ACH Direct Website</a></strong><br />
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
                                  <td width="16%" colspan="2"><img src="images/red_x.png" alt="Unable to add Paradata" width="12" height="12" /> <strong>Payments<em>Gateway</em> cannot be activated</strong> - the required third party components are not installed and working on the web server.<br />                	    
                                  <br /></td>
                                </tr>
                                <tr>
                                    <td colspan="2">
                                        <center>
                                          <strong>Required components for Payments<em>Gateway</em>:</strong><br /><i><%= strComErr %></i><br /><br />
                                   <input type="button" value="Back" onclick="javascript:history.back()"></center> </td>
                                </tr>
                            <% else %>
                                <% if request("mode")="Edit" then %>
                                    <% dim pcPay_ACH_MerchantIdCnt,pcPay_ACH_MerchantIdEnd,pcPay_ACH_MerchantIdStart
                                    pcPay_ACH_MerchantIdCnt=(len(pcPay_ACH_MerchantId)-2)
                                    pcPay_ACH_MerchantIdEnd=right(pcPay_ACH_MerchantId,2)
                                    pcPay_ACH_MerchantIdStart=""
                                    for c=1 to pcPay_ACH_MerchantIdCnt
                                        pcPay_ACH_MerchantIdStart=pcPay_ACH_MerchantIdStart&"*"
                                    next %>
                                    <tr> 
                                        <td height="31" colspan="2">Current Merchant ID:&nbsp;<%=pcPay_ACH_MerchantIdStart&pcPay_ACH_MerchantIdEnd%></td>
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
                                  <td valign="top">&nbsp;</td>
                                  <td valign="top">&nbsp;</td>
                                </tr>
                                <tr> 
                                    <td valign="top"><div align="right">Merchant ID:</div></td>
                                    <td width="84%" valign="top"> <input type="text" name="pcPay_ACH_MerchantID" size="20">
                                        Merchant's six digit ID code</td>
                                </tr>
                                <tr> 
                                    <td valign="top"><div align="right">Password:</div></td>
                                    <td valign="top"><input type="text" name="pcPay_ACH_PWD" size="20"></td>
                                </tr>
                                <tr> 
                                    <td valign="top"><div align="right">Transaction Type:</div></td>
                                    <td valign="top"> <select name="pcPay_ACH_TransType">
                                            <option value="AUTH ONLY" selected>Authorize Only</option>
                                            <option value="SALE" <%if pcPay_ACH_TransType="SALE" then%>selected<% end if %>>Sale</option>
                                        </select> </td>
                                </tr>
                                <tr> 
                                    <td> <div align="right">Require CVV:</div></td>
                                    <td>
                                    `<input type="radio" class="clearBorder" name="pcPay_ACH_CVV" value="1" checked> Yes
                                    <input name="pcPay_ACH_CVV" type="radio" class="clearBorder" value="0" <%if clng(pcPay_ACH_CVV)=0 then%>checked<%end if%>> No</td>
                                </tr>
                                <tr> 
                                    <td> <div align="right">Accepted Cards:</div></td>
                                    <td> 
                                        <input name="pcPay_ACH_CardTypes" type="checkbox" class="clearBorder" value="VISA" <%if (instr(pcPay_ACH_CardTypes,"VISA,")>0) or (instr(pcPay_ACH_CardTypes,"VISA")>0) then%>checked<%end if%>>
                                        Visa 
                                        <input name="pcPay_ACH_CardTypes" type="checkbox" class="clearBorder" value="MAST" <%if (instr(pcPay_ACH_CardTypes,"MAST,")>0) or (instr(pcPay_ACH_CardTypes,"MAST")>0) then%>checked<%end if%>>
                                        MasterCard 
                                        <input name="pcPay_ACH_CardTypes" type="checkbox" class="clearBorder" value="AMER" <%if (instr(pcPay_ACH_CardTypes,"AMER,")>0) or (instr(pcPay_ACH_CardTypes,"AMER")>0) then%>checked<%end if%>>
                                        American Express 
                                        <input name="pcPay_ACH_CardTypes" type="checkbox" class="clearBorder" value="DISC" <%if (instr(pcPay_ACH_CardTypes,"DISC,")>0) or (instr(pcPay_ACH_CardTypes,"DISC")>0) then%>checked<%end if%>>
                                        Discover 
                                        <input name="pcPay_ACH_CardTypes" type="checkbox" class="clearBorder" value="DINE" <%if (instr(pcPay_ACH_CardTypes,"DINE,")>0) or (instr(pcPay_ACH_CardTypes,"DINE")>0) then%>checked<%end if%>>
                                        Diner's Club 
                                        <input name="pcPay_ACH_CardTypes" type="checkbox" class="clearBorder" value="JCB" <%if (instr(pcPay_ACH_CardTypes,"JCB,")>0) or (instr(pcPay_ACH_CardTypes,"JCB")>0) then%>checked<%end if%>>
                                        JCB</td>
                                </tr>
                                <tr> 
                                    <td> <div align="right"> 
                                            <input name="pcPay_ACH_TestMode" type="checkbox" class="clearBorder" id="pcPay_ACH_TestMode" value="1" <% if pcPay_ACH_TestMode=1 then%>checked<% end if%>> 
                                        </div></td>
                                    <td><b>Enable Test Mode </b>(Credit cards will not be charged)</td>
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

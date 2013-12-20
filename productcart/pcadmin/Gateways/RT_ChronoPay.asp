<%
'---Start ChronoPay---
Function gwCPEdit()
	call opendb()
	'request gateway variables and insert them into the ChronoPay table
	query= "SELECT CP_ProdID FROM pcPay_Chronopay WHERE CP_id=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	CP_ProdID2=rs("CP_ProdID")
	set rs=nothing
	CP_ProdID=request.Form("CP_ProdID")
	If CP_ProdID="" then
		CP_ProdID=CP_ProdID2
	end if
	CP_Currency=request.Form("CP_Currency")
	CP_testmode=request.Form("CP_testmode")
	query="UPDATE pcPay_Chronopay SET CP_ProdID='"&CP_ProdID&"', CP_Currency='"&CP_Currency&"',CP_testmode='"&CP_testmode&"' WHERE CP_id=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=52"
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

Function gwCP()
	varCheck=1
	'request gateway variables and insert them into the Chronopay table
	CP_ProdID=request.form("CP_ProdID")
	CP_Currency=request.Form("CP_Currency")
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
	CP_testmode=request.Form("CP_testmode")
	paymentNickName=replace(request.Form("paymentNickName"),"'","''")
	if paymentNickName="" then
		paymentNickName="Credit Card"
	End If
	
	err.clear
	err.number=0
	call openDb() 

	query="UPDATE pcPay_Chronopay SET CP_ProdID='"&CP_ProdID&"',CP_Currency='"&CP_Currency&"',CP_testmode='"&cp_testmode&"' WHERE CP_id=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'ChronoPay','gwChronoPay.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",52,'"&paymentNickName&"')"
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

if request("gwchoice")="52" then
	if request("mode")="Edit" then
		call opendb()
		query= "SELECT CP_ProdID,CP_Currency,CP_testmode FROM pcPay_Chronopay WHERE CP_id=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		CP_ProdID=rs("CP_ProdID")
		CP_Currency=rs("CP_Currency")
		CP_testmode=rs("CP_testmode")
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=52"
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
	<input type="hidden" name="addGw" value="52">
    <!-- New View Start -->
    <table width="100%">
    <tr>
        <td align="left" style="font-size:15px;"><img src="gateways/logos/chronopay_logo.JPG" width="214" height="55"></td>
        <td align="left" style="font-size:15px;">&nbsp;</td>
    </tr>
    </table>
    <br>
    <table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
    <tr>
        <td><strong>ChronoPay<br>
        <br>
        </strong>Accept payments for any goods and services in the internet by bank cards (VISA, MasterCard, Maestro, American Express) and the most popular payment systems in many different countries all over the world via the unique ChronoPay's interface.<strong><br>
        <br>
        <a href="http://www.chronopay.com" target="_blank">ChronoPay Website</a></strong><br />
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
                    <% if request("mode")="Edit" then %>
                        <% dim CP_ProdIDCnt,CP_ProdIDEnd,CP_ProdIDStart
                        CP_ProdIDCnt=(len(CP_ProdID)-2)
                        CP_ProdIDEnd=right(CP_ProdID,2)
                        CP_ProdIDStart=""
                        for c=1 to CP_ProdIDCnt
                            CP_ProdIDStart=CP_ProdIDStart&"*"
                        next %>
                        <tr> 
                            <td colspan="2">Current Product ID:&nbsp;<%=CP_ProdIDStart&CP_ProdIDEnd%></td>
                        </tr>
                        <tr> 
                            <td colspan="2"> For security reasons, your &quot;Product 
                                ID&quot; is only partially shown on this page. If 
                                you need to edit your account information, please 
                                re-enter your &quot;Product ID&quot; below.</td>
                        </tr>
                        <tr>
                          <td align="right">&nbsp;</td>
                          <td>&nbsp;</td>
                        </tr>
                        <tr> 
                            <td width="7%" align="right" nowrap="nowrap">Change Product ID:</td>
                            <td> 
                                <input type="text" value="" name="CP_ProdID" size="24"></td>
                        </tr>
                    <% else %>
                    <tr>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr> 
                        <td> <div align="right">Product ID:</div></td>
                        <td> <input type="text" name="CP_ProdID" size="20"></td>
                    </tr>
                    <% end if %>
                    <tr> 
                        <td> <div align="right">Currency:</div></td>
                        <td>
                            <select name="CP_Currency">
                                <option value="CAD" selected>Canadian Dollars (C $)</option>
                                <option value="EUR" <% if CP_Currency="EUR" then%>selected<% end if %>>Euros 
                                (&euro;)</option>
                                <option value="GBP" <% if CP_Currency="GBP" then%>selected<% end if %>>Pounds 
                                Sterling (&pound;)</option>
                                <option value="JPY" <% if CP_Currency="JPY" then%>selected<% end if %>>Yen 
                                (&yen;)</option>
                                <option value="USD" <% if CP_Currency="USD" then%>selected<% end if %>>U.S. 
                                Dollars ($)</option>
                            </select>
                         </td>
                    </tr>
                    <tr> 
                        <td> <div align="right"> 
                                <input name="cp_testmode" type="checkbox" class="clearBorder" value="YES" <% if CP_testmode="YES" then%>checked<% end if%>> 
                            </div></td>
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
<% end if %>
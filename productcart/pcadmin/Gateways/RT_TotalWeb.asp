<%
'---Start TotalWeb---
Function gwTotalWebEdit()
	call opendb()
	'request gateway variables and insert them into the TotalWeb table
	query="SELECT pcPay_TW_MerchantID,pcPay_TW_CurCode,pcPay_TW_TestMode FROM pcPay_TotalWeb Where pcPay_TW_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If

	pcPay_TW_MerchantID2=rs("pcPay_TW_MerchantID")
	pcPay_TW_MerchantID2=enDeCrypt(pcPay_TW_MerchantID2, scCrypPass)

	pcPay_TW_MerchantID=request.Form("pcPay_TW_MerchantID")
	if pcPay_TW_MerchantID="" then
		pcPay_TW_MerchantID=pcPay_TW_MerchantID2
	end if
	 pcPay_TW_MerchantID=enDeCrypt(pcPay_TW_MerchantID, scCrypPass)
	set rs=nothing
		
	pcPay_TW_CurCode = request.form("pcPay_TW_CurCode")			
	pcPay_TW_TestMode = request.form("pcPay_TW_TestMode")
	if pcPay_TW_TestMode&"" = "" Then
		pcPay_TW_TestMode = 0
	end if
	
	
	query="UPDATE pcPay_TotalWeb SET pcPay_TW_MerchantID='"&pcPay_TW_MerchantID&"', pcPay_TW_CurCode ='" & pcPay_TW_CurCode &"',pcPay_TW_TestMode="&pcPay_TW_TestMode&" WHERE pcPay_TW_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=63"
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

Function gwTotalWeb()
	varCheck=1
	'request gateway variables and insert them into the TotalWeb table
	pcPay_TW_MerchantID=request.Form("pcPay_TW_MerchantID")	
	pcPay_TW_CurCode = request.form("pcPay_TW_CurCode")			
	pcPay_TW_TestMode = request.form("pcPay_TW_TestMode")
	If pcPay_TW_MerchantID=""  then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add TotalWeb as your payment gateway. <b>""Client ID""</b> are required fields.")
	End If
	'encrypt
	pcPay_TW_MerchantID=enDeCrypt(pcPay_TW_MerchantID, scCrypPass)
	
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

	
	query="UPDATE pcPay_TotalWeb SET pcPay_TW_MerchantID='"&pcPay_TW_MerchantID&"', pcPay_TW_CurCode ='" & pcPay_TW_CurCode &"',pcPay_TW_TestMode="&pcPay_TW_TestMode&" WHERE pcPay_TW_ID=1;"
	'Response.write query 
	'Response.end
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'TotalWeb Solutions','gwTotalWeb.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",63,'"&paymentNickName&"')"
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

if request("gwchoice")="63" then
	'Check if table exist in database, if not we will rediect to database update prompt
	err.clear
	call openDb()
	query="select * from pcPay_TotalWeb"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		set rs=nothing
		call closedb()
		response.redirect "upddbTotalWeb.asp"
	else
		set rs=nothing
		call closedb()
	end if
	if request("mode")="Edit" then
	 	call opendb()	
		query="SELECT pcPay_TW_MerchantID,pcPay_TW_CurCode,pcPay_TW_TestMode FROM pcPay_TotalWeb Where pcPay_TW_ID=1;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcPay_TW_MerchantID=rs("pcPay_TW_MerchantID")
		pcPay_TW_MerchantID=enDeCrypt(pcPay_TW_MerchantID, scCrypPass)
		pcPay_TW_CurCode = rs("pcPay_TW_CurCode")
		pcPay_TW_TestMode = rs("pcPay_TW_TestMode")
				
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=63"
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
	<input type="hidden" name="addGw" value="63">

<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/tws_logo_rgb.gif" width="176" height="52"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
<table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
<tr>
    <td><strong><br>
    </strong>Provides a valuable e-commerce solution for small business or enterprise to             accept online payments                around the clock.<strong><br>
    <br>
    <a href="http://payments.totalwebsolutions.com/" target="_blank">TotalWeb Solutions Website</a></strong><br />
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
                <% if request("mode")="Edit" then
					dim pcPay_TW_MerchantIDCnt,pcPay_TW_MerchantIDEnd,pcPay_TW_MerchantIDStart
					pcPay_TW_MerchantIDCnt=(len(pcPay_TW_MerchantID)-2)
					pcPay_TW_MerchantIDEnd=right(pcPay_TW_MerchantID,2)
					pcPay_TW_MerchantIDStart=""
					for c=1 to pcPay_TW_MerchantIDCnt
					pcPay_TW_MerchantIDStart=pcPay_TW_MerchantIDStart&"*"
					next
					
					%>
					<tr> 
						<td height="31">&nbsp;</td>
						<td height="31">Current Client ID:&nbsp;<%=pcPay_TW_MerchantIDStart&pcPay_TW_MerchantIDEnd%></td>
					</tr>
					<tr> 
						<td>&nbsp;</td>
						<td> For security reasons, your &quot;Client ID &quot; 
							is only partially shown on this page. If you need 
							to edit your account information, please re-enter 
							your &quot;Client ID&quot; below.</td>
					</tr>
                <% end if %>
                <tr>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                  <td width="111"> <div align="right">Client&nbsp;ID:</div></td>
                  <td width="328"> <input type="text" name="pcPay_TW_MerchantID" size="20"> </td>
                </tr>
                  <tr> 
                    <td> <div align="right">Currency: </div></td>
                    <td> <select name="pcPay_TW_CurCode">
                            <option value="036" <% if pcPay_TW_CurCode="036" then%>selected<% end if %> >Australian Dollar</option>
                            <option value="124" <% if pcPay_TW_CurCode="124" then%>selected<% end if %>>Canadian Dollar</option>
                            <option value="208" <% if pcPay_TW_CurCode="208" then%>selected<% end if %>>Danish Kone</option>                                      
                            <option value="344" <% if pcPay_TW_CurCode="344" then%>selected<% end if %>>Hong Kong Dollar</option>
                            <option value="376" <% if pcPay_TW_CurCode="376" then%>selected<% end if %>>Israeli Shekel</option>
                            <option value="392" <% if pcPay_TW_CurCode="392" then%>selected<% end if %>>Japanese Yen</option>
							<option value="410" <% if pcPay_TW_CurCode="410" then%>selected<% end if %> >Korean Won</option>
                            <option value="578" <% if pcPay_TW_CurCode="578" then%>selected<% end if %>>Norwegian Krone</option>
                            <option value="826" <% if pcPay_TW_CurCode="826" then%>selected<% end if %>>Pound Sterling</option>                                      
                            <option value="682" <% if pcPay_TW_CurCode="682" then%>selected<% end if %>>Saudi Arabian Rya</option>
                            <option value="752" <% if pcPay_TW_CurCode="752" then%>selected<% end if %>>Swedish Krone</option>
                            <option value="756" <% if pcPay_TW_CurCode="756" then%>selected<% end if %>>Swiss Franc</option>
							<option value="840" <% if pcPay_TW_CurCode="840" then%>selected<% end if %>>US Dollar</option>
                            <option value="978" <% if pcPay_TW_CurCode="978" then%>selected<% end if %>>Euro</option>
                        </select> </td>
                </tr>
				 <tr> 
                    <td> <div align="right"> 
                            <input name="pcPay_TW_TestMode" type="checkbox" class="clearBorder" id="pcPay_TW_TestMode" value="1" <% if pcPay_TW_TestMode=1 then%>checked<%end if %>>
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
<!-- New View End --><% end if %>

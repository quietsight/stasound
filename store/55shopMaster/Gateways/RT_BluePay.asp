<%
'---Start BluePay---
Function gwBluePayEdit()
	call opendb()
	'request gateway variables and insert them into the BluePay table
	query="SELECT BPMerchant,BPTestmode,BPTransType,BPSECRET_KEY FROM BluePay WHERE idBluePay=1"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	BPMerchant2=rstemp("BPMerchant")
	BPMerchant=request.Form("BPMerchant")
	if BPMerchant="" then
		BPMerchant=BPMerchant2
	end if
	BPTestmode=request.Form("BPTestmode")
	if BPTestmode="YES" then
		BPTestmode="1"
	else
		BPTestmode="0"
	end if
	BPTransType2=rstemp("BPTransType")
	BPTransType=request.Form("BPTransType")
	If BPTransType="" then
		BPTransType=BPTransType2
	end if
	BPSECRET_KEY2=rstemp("BPSECRET_KEY")
	BPSECRET_KEY=request.Form("BPSECRET_KEY")
	if BPSECRET_KEY="" then
		BPSECRET_KEY=BPSECRET_KEY2
	end if
	
	query="UPDATE BluePay SET BPMerchant='"&BPMerchant&"',BPTestmode="&BPTestmode&",BPTransType='"&BPTransType&"',BPSECRET_KEY='"&BPSECRET_KEY&"' WHERE idBluePay=1;"
	
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=29"

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

Function gwBluePay()
	varCheck=1
	'request gateway variables and insert them into the BluePay table
	BPMerchant=request.Form("BPMerchant")
	BPTestmode=request.Form("BPTestmode")
	if BPTestmode="YES" then
		BPTestmode="1"
	else
		BPTestmode="0"
	end if
	BPTransType=request.Form("BPTransType")
	BPSECRET_KEY=request.Form("BPSECRET_KEY")
	
	If BPMerchant="" then
		call closedb()
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add BluePay as your payment gateway. <b>""Account ID""</b> is a required field.")
	End If
	
	If BPSECRET_KEY="" then
		call closedb()
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add BluePay as your payment gateway. <b>""Secret Key""</b> is a required field.")
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

	query="UPDATE BluePay SET BPMerchant='"&BPMerchant&"', BPTestmode="&BPTestmode&", BPSECRET_KEY='"&BPSECRET_KEY&"', BPTransType='"&BPTransType&"' WHERE idBluePay=1"
	
	set rs=Server.CreateObject("ADODB.Recordset")  

	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'BluePay','gwBluePay.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",29,'"&paymentNickName&"')"

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

if request("gwchoice")="29" then
	if request("mode")="Edit" then
		call opendb()
		query= "SELECT BPMerchant,BPTestmode,BPTransType,BPSECRET_KEY FROM BluePay WHERE idBluePay=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		BPMerchant=rs("BPMerchant")
		BPTestmode=rs("BPTestmode")
		BPTransType=rs("BPTransType")
		BPSECRET_KEY=rs("BPSECRET_KEY")
				
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=29"
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
	<input type="hidden" name="addGw" value="29">
    <!-- New View Start -->
    <table width="100%">
        <tr>
            <td align="left" style="font-size:15px;"><img src="gateways/logos/bluepay.JPG" width="347" height="92"></td>
            <td align="left" style="font-size:15px;">&nbsp;</td>
        </tr>
    </table>
    <br>
    <table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
        <tr>
            <td><strong>BluePay - Integrated using BluePay 2.0 Post <br>
            <br>
            </strong>An All-In-One Credit Card Processing Company<strong><br>
            <br>
            <a href="http://www.BluePay.com/">BluePay Website</a></strong><br />
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
                            <% dim BPMerchantCnt,BPMerchantEnd,BPMerchantStart
                            BPMerchantCnt=(len(BPMerchant)-2)
                            BPMerchantEnd=right(BPMerchant,2)
                            BPMerchantStart=""
                            for c=1 to BPMerchantCnt
                                BPMerchantStart=BPMerchantStart&"*"
                            next %>
                            <tr> 
                                <td height="31" colspan="2">Current BluePay Merchant Code:&nbsp;<%=BPMerchantStart&BPMerchantEnd%></td>
                            </tr>
                            <tr> 
                                <td colspan="2"> For security reasons, your &quot;BluePay Account 
                                    ID&quot; is only partially shown on this page. If 
                                    you need to edit your account information, please 
                                    re-enter your &quot;Account ID&quot; below.</td>
                            </tr>
                            <tr> 
                                <td width="7%" nowrap="nowrap"> <div align="right">Change Account ID:</div></td>
                                <td><input type="text" value="" name="BPMerchant" size="24"></td>
                            </tr>
                            <tr> 
                                <td> <div align="right">Change Secret Key:</div></td>
                                <td><input type="Password" value="" name="BPSECRET_KEY" size="24">	
                                </td>
                            </tr>
                        <% else %>
                            <tr>
                              <td>&nbsp;</td>
                              <td>&nbsp;</td>
                            </tr>
                            <tr width="7%"> 
                                <td> <div align="right">Account ID:</div></td>
                                <td> <input type="text" name="BPMerchant" size="30"></td>
                            </tr>
                            <tr> 
                                <td> <div align="right">Secret Key:</div></td>
                                <td> <input type="text" name="BPSECRET_KEY" size="30"></td>
                            </tr>
                       <% end if %>
                        <tr> 
                            <td width="7%" nowrap="nowrap"> <div align="right">Transaction Type:</div></td>
                            <td> <select name="BPTransType">
                                    <option value="AUTH" selected>Authorize Only</option>
                                    <option value="SALE" <% if BPTransType="Sale" then%>selected<% end if %>>Sale</option>
                                </select></td>
                        </tr>
                        <tr> 
                            <td> <div align="right"> 
                                    <input name="BPTestmode" type="checkbox" class="clearBorder" value="YES" <% if BPTestmode=1 then%>checked<% end if%>> 
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

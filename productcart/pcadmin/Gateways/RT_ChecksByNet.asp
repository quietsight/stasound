<%
'---Start CheckByNet by CrossCheck---
Function gwCBNEdit()
	call opendb()
	query="SELECT pcPay_CBN_merchant,pcPay_CBN_test,pcPay_CBN_status FROM pcPay_CBN WHERE pcPay_CBN_id=1"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	pcvPay_CBN_merchant2=rs("pcPay_CBN_merchant")
	pcvPay_CBN_test=rs("pcPay_CBN_test")
	pcvPay_CBN_status=rs("pcPay_CBN_status")

	pcPay_CBN_Merchant=request.Form("pcvPay_CBN_Merchant")
		if pcPay_CBN_Merchant = "" then
			pcPay_CBN_Merchant = pcvPay_CBN_merchant2
		end if
	pcPay_CBN_test=request.Form("pcvPay_CBN_Test")
	pcPay_CBN_status=request.Form("pcvPay_CBN_status")
		
	query="UPDATE pcPay_CBN SET pcPay_CBN_Merchant='"&pcPay_CBN_Merchant&"',pcPay_CBN_test="&pcPay_CBN_Test&",pcPay_CBN_status="&pcPay_CBN_status&" WHERE pcPay_CBN_id=1;"
		
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=33"

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

Function gwCBN()
	varCheck=1
	pcPay_CBN_Merchant=request.Form("pcvPay_CBN_Merchant")
	pcPay_CBN_Test=request.Form("pcvPay_CBN_Test")
	pcPay_CBN_Status=request.Form("pcvPay_CBN_Status")

	If pcPay_CBN_Merchant="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add ChecksByNet as your payment gateway. <b>""Merchant ID""</b> is a required field.")
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
		paymentNickName="Electronic Check"
	End If
	
	err.clear
	err.number=0
	call openDb() 

	query="UPDATE pcPay_CBN SET pcPay_CBN_merchant='"&pcPay_CBN_Merchant&"',pcPay_CBN_test="&pcPay_CBN_Test&",pcPay_CBN_status="&pcPay_CBN_Status&" WHERE pcPay_CBN_id=1"

	set rs=Server.CreateObject("ADODB.Recordset")  
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'ChecksByNet','gwCBN.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",33,'"&paymentNickName&"')"

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

if request("gwchoice")="33" then
	if request("mode")="Edit" then
		call opendb()
		query="SELECT pcPay_CBN_merchant,pcPay_CBN_test,pcPay_CBN_status FROM pcPay_CBN WHERE pcPay_CBN_id=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcvPay_CBN_merchant=rs("pcPay_CBN_merchant")
		pcvPay_CBN_test=rs("pcPay_CBN_test")
		pcvPay_CBN_status=rs("pcPay_CBN_status")
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=33"
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
	<input type="hidden" name="addGw" value="33">
    <!-- New View Start -->
    <table width="100%">
    <tr>
        <td align="left" style="font-size:15px;"><img src="gateways/logos/crosscheck_logo.jpg" width="159" height="117"></td>
        <td align="left" style="font-size:15px;">&nbsp;</td>
    </tr>
    </table>
    <br>
    <table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
    <tr>
        <td><strong>CrossCheck - ChecksByNet<br>
        <br>
        </strong>
    ChecksByNet is a<strong> check guarantee service</strong> for   both retail and internet-based businesses with NO setup or application fees, and   NO monthly minimums - available only for United States of America   checking accounts with U.S. funds.
          <strong><br>
        <br>
        <a href="http://www.cross-check.com/" target="_blank">CrossCheck Website</a></strong><br />
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
                    <%if request("mode")="Edit" then %>
                        <% dim pcvPay_CBN_merchantCnt,pcvPay_CBN_merchantEnd,pcvPay_CBN_merchantStart
                        pcvPay_CBN_merchantCnt=(len(pcvPay_CBN_merchant)-2)
                        pcvPay_CBN_merchantEnd=right(pcvPay_CBN_merchant,2)
                        pcvPay_CBN_merchantStart=""
                        for c=1 to pcvPay_CBN_merchantCnt
                            pcvPay_CBN_merchantStart=pcvPay_CBN_merchantStart&"*"
                        next %>
                        <tr> 
                            <td height="31" colspan="2">Current ChecksByNet Merchant ID:&nbsp;<%=pcvPay_CBN_merchantStart&pcvPay_CBN_merchantEnd%></td>
                        </tr>
                        <tr> 
                            <td colspan="2"> For security reasons, your &quot;ChecksByNet 
                                merchant ID&quot; is only partially shown on this 
                                page. If you need to edit your account information, 
                                please re-enter your &quot;Merchant ID&quot; below.</td>
                        </tr>
                        <tr> 
                            <td width="20%" height="31" align="right">Change Merchant Number: </td>
                          <td height="31"><input type="text" value="" name="pcvPay_CBN_merchant" size="20"></td>
                        </tr>
                    <% else %>
                        <tr> 
                            <td align="right">Merchant Number:</td>
                            <td width="431"> <input type="text" name="pcvPay_CBN_Merchant" size="20">								</td>
                        </tr>
                    <% end if %>
                    <tr> 
                        <td width="20%" height="31" align="right">Enable Test Mode? </td>
                        <td><input type="radio" class="clearBorder" name="pcvPay_CBN_test" value="1" Checked> Yes 
                            <input type="radio" class="clearBorder" name="pcvPay_CBN_test" value="0" <% if pcvPay_CBN_test = 0 then %>Checked<% end if %>> No</td>
                    </tr>
                    <tr> 
                        <td colspan="2">
                        When orders are submitted, should they be considered &quot;Pending&quot; or &quot;Processed&quot;? 
                        <input type="radio" class="clearBorder" name="pcvPay_CBN_status" value="1" Checked> Pending 
                        <input name="pcvPay_CBN_status" type="radio" class="clearBorder" value="0" <% if pcvPay_CBN_status = 0 then %>Checked<% end if %>> Processed</td>
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
                            <td width="82%" class="pcSubmenuContent">
                              <input type="radio" class="clearBorder" name="priceToAddType" value="price" <% if priceToAddType="price" then%>checked<%end if%>>Flat fee&nbsp;&nbsp; &nbsp;<%=scCurSign%> <input name="priceToAdd" size="6" value="<%=money(priceToAdd)%>"></td>
                      </tr>
                      <tr>
                        <td>&nbsp;</td>
                            <td class="pcSubmenuContent"><input type="radio" class="clearBorder" name="priceToAddType" value="percentage" <% if priceToAddType="percentage" then%>checked<%end if%>>
                                Percentage of Order Total&nbsp;&nbsp; &nbsp; % 
                                <input name="percentageToAdd" size="6" value="<%=percentageToAdd%>"></td>
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
                                <td width="82%"><input name="paymentNickName" value="<%=paymentNickName%>" size="35" maxlength="255"></td>
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
                                <td>Process orders when they are placed: <input type="checkbox" class="clearBorder" name="pcv_processOrder" value="1" <%if pcv_processOrder="1" then%>checked<%end if%>>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=301')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
                    </tr>
                    <tr> 
                        <td>&nbsp;</td>
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
                          <td>&nbsp;</td>
                        </tr>
                        <tr>
                              <td colspan="2">
								<% if request("mode")="Edit" then
                                    strButtonValue="Save Changes"
								%>
                                <input type="hidden" name="submitMode" value="Edit">
                                <%
                                else
                                    strButtonValue="Add New Payment Method"
								%>
                                <input type="hidden" name="submitMode" value="Add Gateway">
                                <%
                                end if
                                %>
                                <input type="submit" value="<%=strButtonValue%>" name="Submit" class="submit2"> 
                                &nbsp;
                                <input type="button" value="Back" onclick="javascript:history.back()">

                              </td>
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
        

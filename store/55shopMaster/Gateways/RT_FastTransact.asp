<%
'---Start FastTransAct---
Function gwfastEdit()
	call opendb()
	'request gateway variables and insert them into the fasttransact table
	query= "SELECT AccountID, SiteTag FROM fasttransact WHERE id=1;"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	AccountID2=rstemp("AccountID")
	SiteTag2=rstemp("SiteTag")
	AccountID=request.Form("AccountID")
	If AccountID="" then
		AccountID=AccountID2
	end if
	SiteTag=request.Form("SiteTag")
	if SiteTag="" then
		SiteTag=SiteTag2
	end if
	tran_type=request.Form("tran_type")
	card_types=request.Form("card_types")
	CVV2=request.Form("CVV2")
	query="UPDATE fasttransact SET AccountID='"&AccountID&"',SiteTag='"&SiteTag&"',tran_type='"&tran_type&"',card_types='"&card_types&"',CVV2="&CVV2&" WHERE id=1;"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=15"
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

Function gwfast()
	varCheck=1
	'request gateway variables and insert them into the fasttransact table
	AccountId=request.form("AccountID")
	SiteTag=request.form("SiteTag")
	tran_type=request.Form("tran_type")
	card_types=request.Form("card_types")
	CVV2=request.Form("CVV2")
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
	
	call opendb()
	
	query="UPDATE fasttransact SET AccountID='"&AccountID&"',SiteTag='"&SiteTag&"',tran_type='"&tran_type&"',card_types='"&card_types&"',CVV2="&CVV2&" WHERE id=1;"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'FastTransact','gwfast.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",15,'"&paymentNickName&"')"
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

if request("gwchoice")="15" then
	if request("mode")="Edit" then
		call opendb()
		query= "SELECT AccountID, SiteTag, tran_type, card_types, CVV2 FROM fasttransact WHERE id=1;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		AccountID=rs("AccountID")
		SiteTag=rs("SiteTag")
		tran_type=rs("tran_type")
		card_types=rs("card_types")
		CVV2=rs("CVV2")
		
		cardTypeArray=split(card_types,", ")
		for i=0 to ubound(cardTypeArray)
			select case cardTypeArray(i)
				case "M"
					M="1" 
				case "V"
					V="1"
				case "D"
					D="1"
				case "A"
					A="1"
			end select
		next
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=15"
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
	<input type="hidden" name="addGw" value="15">
    <tr> 
        <td height="21"><b>Fast Transact</b> ( <a href="http://www.fasttransact.com/" target="_blank">Web site</a> )</td>
    </tr>
    <tr> 
        <td>
        	<table width="100%" border="0" cellspacing="0" cellpadding="4">
                <tr> 
                    <th colspan="2">Enter Fast Transact  settings</th>
                </tr>
                <% if request("mode")="Edit" then %>
					<% dim FT_AccountIDCnt,FT_AccountIDEnd,FT_AccountIDStart
                    FT_AccountIDCnt=(len(AccountID)-2)
                    FT_AccountIDEnd=right(AccountID,2)
                    FT_AccountIDStart=""
                    for c=1 to FT_AccountIDCnt
                        FT_AccountIDStart=FT_AccountIDStart&"*"
                    next %>
                    <tr> 
                        <td height="31" colspan="2">Current Account ID:&nbsp;<%=FT_AccountIDStart&FT_AccountIDEnd%></td>
                    </tr>
                    <% dim FT_SiteTagCnt,FT_SiteTagEnd,FT_SiteTagStart
                    FT_SiteTagCnt=(len(SiteTag)-2)
                    FT_SiteTagEnd=right(SiteTag,2)
                    FT_SiteTagStart=""
                    for c=1 to FT_SiteTagCnt
                        FT_SiteTagStart=FT_SiteTagStart&"*"
                    next %>
                    <tr> 
                        <td height="31" colspan="2">Current Site Tag:&nbsp;<%=FT_SiteTagStart&FT_SiteTagEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;Account ID&quot; 
                            and &quot;Site Tag&quot; is only partially shown 
                            on this page. If you need to edit your account information, 
                            please re-enter your &quot;Account ID&quot; and 
                            &quot;Site Tag&quot; below.</td>
                    </tr>
                <% end if %>
                <tr bgcolor="#FFFFFF"> 
                    <td> <div align="right">Account ID :</div></td>
                    <td> <input type="text" name="AccountID" size="20"> </td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td> <div align="right">Site Tag :</div></td>
                    <td width="440"> <input name="SiteTag" type="text" size="30"></td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td> <div align="right">Transaction Type:</div></td>
                    <td> <select name="tran_type">
                            <option value="SALE" selected>Sale</option>
                            <option value="PREAUTH" <% if tran_type="PREAUTH" then %>selected<% end if %>>Authorize Only</option>
                        </select> </td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td> <div align="right">Require CVV:</div></td>
                    <td> <input type="radio" class="clearBorder" name="CVV2" value="1" checked>
                        Yes 
                        <input type="radio" class="clearBorder" name="CVV2" value="0" <% if CVV2="0" then %>checked<% end if %>>
                        No</td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td> <div align="right">Accepted Cards:</div></td>
                    <td>
                    <% if V="1" then %> <input name="card_types" type="checkbox" class="clearBorder" value="V" checked> 
					<% else %> <input name="card_types" type="checkbox" class="clearBorder" value="V"> 
                    <% end if %>
                    Visa 
                    <% if M="1" then %> <input name="card_types" type="checkbox" class="clearBorder" value="M" checked> 
                    <% else %> <input name="card_types" type="checkbox" class="clearBorder" value="M"> 
                    <% end if %>
                    MasterCard 
                    <% if A="1" then %> <input name="card_types" type="checkbox" class="clearBorder" value="A" checked> 
                    <% else %> <input name="card_types" type="checkbox" class="clearBorder" value="A"> 
                    <% end if %>
                    American Express 
                    <% if D="1" then %> <input name="card_types" type="checkbox" class="clearBorder" value="D" checked> 
                    <% else %> <input name="card_types" type="checkbox" class="clearBorder" value="D"> 
                    <% end if %>
                    Discover
                    </td>
                </tr>
                <tr> 
                    <th colspan="2">You have the option to charge a processing fee for this payment option.</th>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td width="111"> <div align="right">Processing fee: </div></td>
                    <td> <input type="radio" class="clearBorder" name="priceToAddType" value="price" <% if priceToAddType="price" then%>checked<%end if%>>
                        Flat fee&nbsp;&nbsp; &nbsp;<%=scCurSign%> <input name="priceToAdd" size="6" value="<%=money(priceToAdd)%>">
                    </td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td>&nbsp;</td>
                    <td> <input type="radio" class="clearBorder" name="priceToAddType" value="percentage" <% if priceToAddType="percentage" then%>checked<%end if%>>
                        Percentage of Order Total&nbsp;&nbsp; &nbsp; % 
                        <input name="percentageToAdd" size="6" value="<%=percentageToAdd%>">
                    </td>
                </tr>
                <tr> 
                    <th colspan="2">You can change the display name that is shown for this payment type. </th>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td><div align="right">Payment Name:&nbsp;</div></td>
                    <td><input name="paymentNickName" value="<%=paymentNickName%>" size="35" maxlength="255"></td>
                </tr>
            </table>
        </td>
    </tr>
<!-- END fasttransact -->

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

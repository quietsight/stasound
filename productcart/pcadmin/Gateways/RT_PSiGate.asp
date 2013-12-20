<%		
'---Start PsiGate---
Function gwpsiEdit()
	call opendb()
	'request gateway variables and insert them into the PSIGate table
	pcv_PSIPOST=request.Form("PSIPOST")
	select case pcv_PSIPOST
		case "DLL"
			psiFile="gwPSI.asp"
			psi_post="DLL"
			Config_File_Name=request.form("Config_File_Name")
			Config_File_Name_Full=request.Form("Config_File_Name_Full")
			Host=request.form("Host")
			Port=request.form("Port")
			Userid=request.form("Userid")
			psi_TransType=request.form("pMode")
			psi_testmode=request.Form("psi_testmode")
		case "HTML"
			query= "SELECT Userid FROM PSIGate WHERE id=1"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=conntemp.execute(query)
			if err.number <> 0 then
				strErrorDescription=err.description
				set rs=nothing
				call closeDb()
				response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
			end If
			Userid2=rs("Userid")
			set rs=nothing
			psiFile="gwPSI_H.asp"
			psi_post="HTML"
			Config_File_Name="NA"
			Config_File_Name_Full="NA"
			Host="NA"
			Port="NA"
			Userid=request.form("psi_merchantID")
			if Userid="" then
				Userid=Userid2
			end if
			psi_TransType=request.form("ModeH")
			psi_testmode=request.form("psi_testmodeH")
		case "XML"
			query= "SELECT Config_File_Name, Userid FROM PSIGate WHERE id=1"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=conntemp.execute(query)
			if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
			end If
			Config_File_Name2=rs("Config_File_Name")
			Userid2=rs("Userid")
			set rs=nothing
			psiFile="gwPSI_XML.asp"
			psi_post="XML"
			Config_File_Name_Full="NA"
			Host="NA"
			Port="NA"
			Userid=request.form("psi_XMLStoreID")
			if Userid="" then
				Userid=Userid2
			end if
			Config_File_Name=request.form("psi_XMLPassPhrase")
			if Config_File_Name="" then
				Config_File_Name=Config_File_Name2
			end if
			psi_TransType=request.form("psi_XMLTransType")
			psi_testmode=request.form("psi_XMLTestmode")
	end select
	query="UPDATE PSIGate SET Config_File_Name='"&Config_File_Name&"',Config_File_Name_Full='"&Config_File_Name_Full&"',Host='"&Host&"',Port='"&Port&"',Userid='"&Userid&"',Mode="&psi_TransType&",psi_post='"&psi_post&"',psi_testmode='"&psi_testmode&"' where id=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",sslURL='"&psiFile&"',priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=4"
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

Function gwpsi()
	varCheck=1
	'request gateway variables and insert them into the PSIGate table
	pcv_PSIPostType=request.form("PSIPOST")
	select case pcv_PSIPostType
	 case "DLL"
		psiFile="gwPSI.asp"
		psi_post="DLL"
		Config_File_Name=request.form("Config_File_Name")
		Config_File_Name_Full=request.Form("Config_File_Name_Full")
		Host=request.form("Host")
		Port=request.form("Port")
		Userid=request.form("Userid")
		tMode=request.form("pMode")
		psi_testmode=request.Form("psi_testmode")
		if psi_testmode="" then
			psi_testmode="NO"
		end if
	case "HTML"
		psiFile="gwPSI_H.asp"
		psi_post="HTML"
		Config_File_Name=request.form("Config_File_Name")
		Config_File_Name_Full=request.Form("Config_File_Name_Full")
		Host=request.form("Host")
		Port=request.form("Port")
		Userid=request.form("psi_merchantID")
		tMode=request.form("ModeH")
		psi_testmode=request.Form("psi_testmodeH")
		if psi_testmode="" then
			psi_testmode="NO"
		end if
	case "XML"
		psiFile="gwPSI_XML.asp"
		psi_post="XML"
		Userid=request.form("psi_XMLStoreID")
		Config_File_Name=request.form("psi_XMLPassPhrase") 'use the old dll variable name for the XML PassPhrase
		Config_File_Name_Full="NA"
		Host="NA"
		Port="NA"
		tMode=request.form("psi_XMLTransType")
		psi_testmode=request.Form("psi_XMLTestmode")
		if psi_testmode="" then
			psi_testmode="NO"
		end if
		
	end select

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
	End If
	
	err.clear
	err.number=0
	call openDb() 

	query="UPDATE PSIGate SET Config_File_Name='"&Config_File_Name&"',Config_File_Name_Full='"&Config_File_Name_Full&"',Host='"&Host&"',Port='"&Port&"',Userid='"&Userid&"',Mode='"&tMode&"',psi_post='"&psi_post&"',psi_testmode='"&psi_testmode&"' WHERE id=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'PSIGate','"&psiFile&"',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",4,'"&paymentNickName&"')"
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

if request("gwchoice")="4" then
	if request("mode")="Edit" then
		call opendb()
		query= "SELECT Config_File_Name,Config_File_Name_Full,Host,Port,Userid,Mode,psi_post,psi_testmode FROM PSIGate WHERE id=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		Config_File_Name=rs("Config_File_Name")
		Config_File_Name_Full=rs("Config_File_Name_Full")
		Host=rs("Host")
		Port=rs("Port")
		Userid=rs("Userid")
		psi_TransType=rs("Mode")
		psi_post=rs("psi_post")
		psi_testmode=rs("psi_testmode")
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=4"
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
		call closedb() %>
		<input type="hidden" name="mode" value="Edit">
    <% end if %>
	<input type="hidden" name="addGw" value="4">
<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/psigate_logo.gif" width="265" height="64"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
<table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
<tr>
    <td><strong><br>
    </strong>PSiGate makes it easy to get your business online and accepting credit card and debit transactions quickly and securely.<strong><br>
    <br>
    <a href="http://www.psigate.com/" target="_blank">PSiGate Website</a></strong><br />
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
                  <td colspan="2">&nbsp;</td>
                </tr>
                <tr> 
                    <td colspan="2"><input name="PSIPOST" type="radio" class="clearBorder" value="XML" checked>
                        Use PSiGate XML Posting - No server component required.</td>
                </tr>
                <% if request("mode")="Edit" AND psi_post="XML" then
					dim UserXMLIDCnt,UserXMLIDEnd,UserXMLIDStart
					UserXMLIDCnt=(len(UserID)-2)
					UserXMLIDEnd=right(UserID,2)
					UserXMLIDStart=""
					for c=1 to UserXMLIDCnt
						UserXMLIDStart=UserXMLIDStart&"*"
					next %>
                    <tr> 
                        <td colspan="2">Current XML Store ID:&nbsp;<%=UserXMLIDStart&UserXMLIDEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;XML Store ID&quot; 
                            is only partially shown on this page and your Pass Phrase has been hidden. If you need 
                            to edit your account information, please re-enter 
                            your &quot;Store ID&quot; or &quot;Pass Phrase&quot; below.</td>
                    </tr>
                <% end if %>
                <tr>
                    <td width="111"><div align="right">Merchant ID:</div></td>
                    <td><input name="psi_XMLStoreID" type="text" value="" size="30"></td>
                </tr>
                <tr> 
                    <td><div align="right">Passphrase:</div></td>
                    <td><input name="psi_XMLPassPhrase" type="text" value="" size="30"></td>
                </tr>
                <tr> 
                    <td><div align="right">Transaction Type:</div></td>
                    <td>
                    	<select name="psi_XMLTransType">
                            <option value="0" selected>Sale</option>
                            <option value="1" <% if psi_TransType="1" then %>selected<% end if %>>Authorize Only</option>
                        </select>
                    </td>
                </tr>
                <tr> 
                    <td><div align="right"><input name="psi_XMLTestmode" type="checkbox" class="clearBorder" value="YES" <% if psi_testmode="YES" then %>checked<% end if %> /></div></td>
                    <td><b>Enable Test Mode </b>(Credit cards will not be charged)</td>
                </tr>
                <tr> 
                    <td colspan="2"><hr size="1" noshade></td>
                </tr>
                <tr> 
                    <td colspan="2"><input name="PSIPOST" type="radio" class="clearBorder" value="HTML" <% if psi_post="HTML" then %>checked<% end if %>> Use PSiGate HTML Posting - No server component required.</td>
                </tr>
				<% if request("mode")="Edit" AND psi_post="HTML" then
                    dim UserIDCnt,UserIDEnd,UserIDStart
                    UserIDCnt=(len(UserID)-2)
                    UserIDEnd=right(UserID,2)
                    UserIDStart=""
                    for c=1 to UserIDCnt
                        UserIDStart=UserIDStart&"*"
                    next %>
                    <tr> 
                        <td colspan="2">Current Merchant ID:&nbsp;<%=UserIDStart&UserIDEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;Merchant ID&quot; 
                            is only partially shown on this page. If you need 
                            to edit your account information, please re-enter 
                            your &quot;Merchant ID&quot; below.</td>
                    </tr>
                <% end if %>
                <tr> 
                    <td> <div align="right">Merchant ID:</div></td>
                    <td> <input name="psi_merchantID" type="text" value="" size="30"></td>
                </tr>
                <tr> 
                    <td> <div align="right">Transaction Type:</div></td>
                    <td>
                    	<select name="ModeH">
                            <option value="0" selected>Sale</option>
                            <option value="1" <% if psi_TransType="1" then %>selected<% end if %>>Authorize Only</option>
                        </select>
                	</td>
                </tr>
                <%
				'//Check for required components and that they are working 
				reDim strComponent(0)
				reDim strClass(0,1)
			
				'The component names
				strComponent(0) = "PSIGate"
				
				'The component class names
				strClass(0,0) = "MyServer.PsiGate"
			
				isComErr = Cint(0)
				strComErr = Cstr()
				
				For i=0 to UBound(strComponent)
					strErr = IsObjInstalled(i)
					If strErr <> "" Then
						strComErr = strComErr & strErr
						isComErr = 1
					End If
				Next
				%>
                <tr> 
                    <td><div align="right"><input name="psi_testmodeH" type="checkbox" class="clearBorder" value="YES" <% if psi_testmode="YES" then %>checked<% end if %> /></div></td>
                    <td><b>Enable Test Mode </b>(Credit cards will not be charged)</td>
                </tr>
                <tr> 
                    <td colspan="2"><hr size="1" noshade></td>
                </tr>
                <% if psi_post="DLL" then %>
                <% else
					Config_File_Name=""
					Config_File_Name_Full=""
					Host=""
					Port=""
					Userid=""
					psi_TransType=""
					psi_post=rs("psi_post")
					psi_testmode=""
				end if %>
                <tr>
                	<td colspan="2">
                	<%
					if isComErr = 1 then %>
                    <input type="radio" class="clearBorder" disabled="disabled" />Use PSiGate Transaction DLL - Requires the PSiGate server component.<br /><br />
						&nbsp;<img src="images/red_x.png" width="12" height="12" /><strong> &nbsp;PSiGate Transaction DLL</strong> cannot be enabled. Errors were found while testing for the required components. These library files are available for download directly from PSiGate and need to be installed directly on the server.<br /> 
							<center><strong><br />
				      Required components for PSiGate Transaction DLL:</strong><br />
				      <i><%= strComErr %></i></center><br />
                 	<% else%>
                    <input type="radio" class="clearBorder" name="PSIPOST" value="DLL" <% if psi_post="DLL" then %>checked<% end if %>>Use PSiGate Transaction DLL - Requires the PSiGate server component.
                    <% end if %>
                    </td>
                	
                </tr>
                <% if isComErr = 1 then
			   		intDoNotApply = 0
				else %>
                    <tr> 
                        <td><div align="right">Configuration File Name:</div></td>
                        <td><input type="text" name="Config_File_Name" size="30" value="<%=Config_File_Name%>"></td>
                    </tr>
                    <tr> 
                        <td> <div align="right">Certificate Path/Name:</div></td>
                        <td> <input type="text" name="Config_File_Name_Full" size="35" value="<%=Config_File_Name_Full%>"></td>
                    </tr>
                    <tr> 
                        <td> <div align="right">Host:</div></td>
                        <td> <input type="text" name="Host" size="24" value="<%=Host%>"></td>
                    </tr>
                    <tr> 
                        <td> <div align="right">Port:</div></td>
                        <td> <input type="text" name="Port" size="15" value="<%=Port%>"></td>
                    </tr>
                    <input type="hidden" name="Userid" size="24" value="<%=Userid%>">
                    <tr> 
                        <td> <div align="right">Transaction Type:</div></td>
                        <td> <select name="pMode">
                                <option value="0" selected>Sale</option>
                               <option value="1" <% if psi_TransType="1" then %>selected<% end if %>>Authorize Only</option>
                            </select> </td>
                    </tr>
                    <tr> 
                        <td> <div align="right"> 
                                <input name="psi_testmode" type="checkbox" class="clearBorder" value="YES" <% if psi_testmode="YES" then %>checked<% end if %> />
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

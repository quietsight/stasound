<%
'---Start USAePay---
Function gwuepEdit()
	call opendb()
	query="SELECT pcPay_Uep_SourceKey, pcPay_Uep_TransType, pcPay_Uep_TestMode, pcPay_Uep_Checking, pcPay_Uep_CheckPending FROM pcPay_USAePay WHERE pcPay_Uep_Id=1"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	pcPay_Uep_SourceKey2=rstemp("pcPay_Uep_SourceKey")
	pcPay_Uep_SourceKey2=enDeCrypt(pcPay_Uep_SourceKey2, scCrypPass)
	pcPay_Uep_SourceKey=request.Form("pcPay_Uep_SourceKey")
	if pcPay_Uep_SourceKey="" then
		pcPay_Uep_SourceKey=pcPay_Uep_SourceKey2
	end if
	pcPay_Uep_TestMode=request.Form("pcPay_Uep_TestMode")
	if pcPay_Uep_TestMode="" then
		pcPay_Uep_TestMode="0"
	end if
	pcPay_Uep_TransType2=rstemp("pcPay_Uep_TransType")
	pcPay_Uep_TransType=request.Form("pcPay_Uep_TransType")
	if pcPay_Uep_TransType="" then
		pcPay_Uep_TransType=pcPay_Uep_TransType2
	end if
	pcPay_Uep_Checking=request.Form("pcPay_Uep_Checking")
	pcPay_Uep_CheckPending=request.Form("pcPay_Uep_CheckPending")

	paymentNickName2=replace(request.Form("paymentNickName2"),"'","''")
	if paymentNickName2="" then
		paymentNickName2="eCheck"
	end if
	'encrypt
	pcPay_Uep_SourceKey=enDeCrypt(pcPay_Uep_SourceKey, scCrypPass)
	
	'check to see if centinel is activated
	pcPay_Cent_Active=request.Form("pcPay_Cent_Active")
	if pcPay_Cent_Active="YES" then
		pcPay_Cent_Active=1
	else
		pcPay_Cent_Active=0
	end if
	pcPay_Cent_TransactionURL=request.Form("pcPay_Cent_TransactionURL")
	pcPay_Cent_MerchantID=request.Form("pcPay_Cent_MerchantID")
	pcPay_Cent_ProcessorId=request.Form("pcPay_Cent_ProcessorId")
	if pcPay_Cent_TransactionURL="" or pcPay_Cent_MerchantID="" OR	pcPay_Cent_ProcessorId="" then
		pcPay_Cent_Active=0
		query="UPDATE pcPay_Centinel SET pcPay_Cent_TransactionURL='"&pcPay_Cent_TransactionURL&"',	pcPay_Cent_MerchantID='"&pcPay_Cent_MerchantID&"',	pcPay_Cent_ProcessorId='"&pcPay_Cent_ProcessorId&"', pcPay_Cent_Active=0 WHERE pcPay_Cent_ID=1;"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
	else
		query="UPDATE pcPay_Centinel SET pcPay_Cent_TransactionURL='"&pcPay_Cent_TransactionURL&"',	pcPay_Cent_MerchantID='"&pcPay_Cent_MerchantID&"',	pcPay_Cent_ProcessorId='"&pcPay_Cent_ProcessorId&"', pcPay_Cent_Active="&pcPay_Cent_Active&" WHERE pcPay_Cent_ID=1;"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
	end if
	
	query="UPDATE pcPay_USAePay SET pcPay_Uep_SourceKey='"&pcPay_Uep_SourceKey&"',pcPay_Uep_TransType=" & pcPay_Uep_TransType & ",pcPay_Uep_TestMode="&pcPay_Uep_TestMode&",pcPay_Uep_Checking="&pcPay_Uep_Checking&",pcPay_Uep_CheckPending="&pcPay_Uep_CheckPending&" WHERE pcPay_Uep_Id=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	query="UPDATE payTypes SET  pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ", priceToAdd="&priceToAdd&" ,percentageToAdd="&percentageToAdd&" ,paymentNickName='"&paymentNickName&"' WHERE gwCode=35"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	if pcPay_Uep_Checking="1" then
	
		query="SELECT paymentDesc FROM payTypes where gwCode=36"
		set rs=conntemp.execute(query)
		
		if not rs.eof then
			query="UPDATE payTypes SET priceToAdd="&priceToAdd&" ,percentageToAdd="&percentageToAdd&" ,paymentNickName='"&paymentNickName2&"' WHERE gwCode=36"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=conntemp.execute(query)
		else
			query="INSERT INTO payTypes (paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES ('USAePay eCheck','gwUSAePayCheck.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",36,'"&paymentNickName2&"')"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=conntemp.execute(query)
		end if
	
	else
	
		query="DELETE FROM payTypes where gwCode=36"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=conntemp.execute(query)
	
	end if
	
	set rs=nothing
	call closedb()
end function

Function gwuep()
	varCheck=1
	'request gateway variables and insert them into the CyberSource table
	pcPay_Uep_SourceKey=request.Form("pcPay_Uep_SourceKey")
	pcPay_Uep_TransType=request.Form("pcPay_Uep_TransType")
	if pcPay_Uep_TransType="" then
		pcPay_Uep_TransType="0"
	end if
	pcPay_Uep_TestMode=request.Form("pcPay_Uep_TestMode")
	if pcPay_Uep_TestMode="" then
		pcPay_Uep_TestMode="0"
	end if
	pcPay_Uep_Checking=request.Form("pcPay_Uep_Checking")
	pcPay_Uep_CheckPending=request.Form("pcPay_Uep_CheckPending")
	If pcPay_Uep_SourceKey="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add USAePay as your payment gateway. <b>""Source Key""</b> is a required field.")
	End If
	'encrypt
	pcPay_Uep_SourceKey=enDeCrypt(pcPay_Uep_SourceKey, scCrypPass)
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
	paymentNickName2=replace(request.Form("paymentNickName2"),"'","''")
	if paymentNickName2="" then
		paymentNickName2="eCheck"
	end if
	
	paymentNickName=replace(request.Form("paymentNickName"),"'","''")
	if paymentNickName="" then
		paymentNickName="Credit Card"
	End If
	
	'check to see if centinel is activated
	pcPay_Cent_Active=request.Form("pcPay_Cent_Active_uep")
	if pcPay_Cent_Active="YES" then
		pcPay_Cent_Active=1
		pcPay_Cent_TransactionURL=request.Form("pcPay_Cent_TransactionURL_uep")
		pcPay_Cent_MerchantID=request.Form("pcPay_Cent_MerchantID_uep")
		pcPay_Cent_ProcessorId=request.Form("pcPay_Cent_ProcessorId_uep")
		if pcPay_Cent_TransactionURL="" or pcPay_Cent_MerchantID="" OR	pcPay_Cent_ProcessorId="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Cardinal Centinel for Authorize.Net. <b>""Tranaction URL, Merchant ID and Process ID""</b> are all required fields.")
		else
			err.clear
			err.number=0
			call openDb() 
			query="UPDATE pcPay_Centinel SET pcPay_Cent_TransactionURL='"&pcPay_Cent_TransactionURL&"',	pcPay_Cent_MerchantID='"&pcPay_Cent_MerchantID&"',	pcPay_Cent_ProcessorId='"&pcPay_Cent_ProcessorId&"', pcPay_Cent_Active=1 WHERE pcPay_Cent_ID=1;"
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
		end if
	end if
	
	err.clear
	err.number=0
	
	call openDb() 

	query="UPDATE pcPay_USAePay SET pcPay_Uep_SourceKey='"&pcPay_Uep_SourceKey&"',pcPay_Uep_TransType=" & pcPay_Uep_TransType & ",pcPay_Uep_TestMode="&pcPay_Uep_TestMode&",pcPay_Uep_Checking="&pcPay_Uep_Checking&",pcPay_Uep_CheckPending="&pcPay_Uep_CheckPending&" WHERE pcPay_Uep_Id=1"
	set rs=Server.CreateObject("ADODB.Recordset")  
	   
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'USAePay','gwUSAePay.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",35,'"&paymentNickName&"')"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if pcPay_Uep_Checking=1 then
		query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'USAePay eCheck','gwUSAePayCheck.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",36,'"&paymentNickName2&"')"
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

'35 = Credit
'36 = Checks

if request("gwchoice")="35" then
	intDoNotApply = 0

	'//Check for required components and that they are working 
	reDim strComponent(0)
	reDim strClass(0,1)

	'The component names
	strComponent(0) = "USAePay"
	
	'The component class names
	strClass(0,0) = "USAePayXChargeCom2.XChargeCom2"

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
		query="Select pcPay_Uep_SourceKey,pcPay_Uep_TransType,pcPay_Uep_TestMode,pcPay_Uep_Checking,pcPay_Uep_CheckPending from pcPay_USAePay where pcPay_Uep_Id=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcPay_Uep_SourceKey=rs("pcPay_Uep_SourceKey")
		pcPay_Uep_TransType=rs("pcPay_Uep_TransType")
		pcPay_Uep_SourceKey=enDeCrypt(pcPay_Uep_SourceKey, scCrypPass)
		pcPay_Uep_TestMode=rs("pcPay_Uep_TestMode")
		pcPay_Uep_Checking=rs("pcPay_Uep_Checking")
		pcPay_Uep_CheckPending=rs("pcPay_Uep_CheckPending")
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=35"
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

		query="SELECT pcPay_Centinel.pcPay_Cent_TransactionURL, pcPay_Centinel.pcPay_Cent_ProcessorId, pcPay_Centinel.pcPay_Cent_MerchantID, pcPay_Centinel.pcPay_Cent_Active, pcPay_Centinel.pcPay_Cent_Password FROM pcPay_Centinel WHERE (((pcPay_Centinel.pcPay_Cent_ID)=1));"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		pcPay_Cent_TransactionURL = rs("pcPay_Cent_TransactionURL")
		pcPay_Cent_ProcessorId = rs("pcPay_Cent_ProcessorId")
		pcPay_Cent_MerchantID = rs("pcPay_Cent_MerchantID")
		pcPay_Cent_Active = rs("pcPay_Cent_Active")
		pcPay_Cent_Password = rs("pcPay_Cent_Password")
		set rs=nothing
		
		'//Nickname for checks
		query= "SELECT paymentNickName FROM payTypes WHERE gwCode=36"
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
	<input type="hidden" name="addGw" value="35">
<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/usaepay_logo.jpg" width="157" height="90"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
<table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
<tr>
    <td><strong><br>
    </strong>Since 1998 the USA ePay team has been making it possible for businesses to   accept credit cards easily, safely and securely.<strong><br>
    <br>
    <a href="http://www.usaepay.com/" target="_blank">USAePay Website
    </a></strong><br />
    <br /></td>
</tr>
<tr>
    <td>
        <div id="CollapsiblePanel1" class="CollapsiblePanel">
            <div class="CollapsiblePanelTab1">
                <table width="100%">
                  <tr>
                    <td width="580" class="pcPanelTitle1">Step 1: Configure Account - USAePay</td>
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
                          <td colspan="2"><img src="images/red_x.png" alt="Unable to add USAePay" width="12" height="12" /> <strong>USAePay cannot be activated</strong> - the required third party components are not installed and working on the web server.<br />                	    
                          <br /></td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <center>
                              <strong>Required components for <b>USAePay</b>:</strong><br /><i><%= strComErr %></i>
                              <br /><br /><input type="button" value="Back" onclick="javascript:history.back()"></center>
                              </td>
                        </tr>
					<% else %>
						<% if request("mode")="Edit" then %>
							<% dim pcPay_Uep_SourceKeyCnt,pcPay_Uep_SourceKeyEnd,pcPay_Uep_SourceKeyStart
                            pcPay_Uep_SourceKeyCnt=(len(pcPay_Uep_SourceKey)-2)
                            pcPay_Uep_SourceKeyEnd=right(pcPay_Uep_SourceKey,2)
                            pcPay_Uep_SourceKeyStart=""
                            for c=1 to pcPay_Uep_SourceKeyCnt
                                pcPay_Uep_SourceKeyStart=pcPay_Uep_SourceKeyStart&"*"
                            next %>
                            <tr> 
                                <td height="31" colspan="2">Current USAePay Source Key:&nbsp;<%=pcPay_Uep_SourceKeyStart&pcPay_Uep_SourceKeyEnd%></td>
                            </tr>
                            <tr> 
                                <td colspan="2"> For security reasons, your &quot;Source Key&quot; 
                                    is only partially shown on this page. If you need 
                                    to edit your account information, please re-enter 
                                    your &quot;Source Key&quot; below.</td>
                            </tr>
						<% end if %>
                        <tr>
                          <td valign="top">&nbsp;</td>
                          <td valign="top">&nbsp;</td>
                        </tr>
                        <tr> 
                            <td width="111" valign="top"><div align="right">Source Key:</div></td>
                            <td valign="top"> <input type="text" name="pcPay_Uep_SourceKey" size="20"> 
                                (Generated by the Merchant Console at www.usaepay.com)</td>
                        </tr>
                        <tr> 
                            <td valign="top"><div align="right">Transaction Type:</div></td>
                            <td valign="top"> <select name="pcPay_Uep_TransType">
                                    <option value="1" selected>Sale</option>
                                    <option value="0" <%if pcPay_Uep_TransType="0" then%>selected<%end if%>>Authorize Only</option>
                                </select> </td>
                        </tr>
                        <tr> 
                            <td> <div align="right"> 
                                    <input name="pcPay_Uep_TestMode" type="checkbox" class="clearBorder" id="pcPay_Uep_TestMode" value="1" <% if pcPay_Uep_TestMode="1" then %>checked<% end if %> />
                                </div></td>
                            <td><b>Enable Test Mode </b>(Credit cards will not be charged)</td>
                        </tr>
                        <tr> 
                            <td><div align="right">Accept Checks:</div></td>
                            <td> <input type="radio" class="clearBorder" name="pcPay_Uep_Checking" value="1" checked>
                                Yes 
                                <input name="pcPay_Uep_Checking" type="radio" class="clearBorder" value="0" <% if pcPay_Uep_Checking=0 then%>checked<% end if %>>
                                No</td>
                        </tr>
                        <tr> 
                            <td colspan="2">All electronic check orders are considered 
                                &quot;Processed&quot;, as the order amount is always 
                                debited to the customer's bank account. If for any reasons 
                                you would like electronic check orders to be considered 
                                &quot;Pending&quot;, use this option. Should electronic 
                                check orders be considered &quot;Pending&quot;? 
                                <input type="radio" class="clearBorder" name="pcPay_Uep_CheckPending" value="1" checked>
                                Yes 
                                <input name="pcPay_Uep_CheckPending" type="radio" class="clearBorder" value="0" <% if pcPay_Uep_CheckPending=0 then%>checked<%end if %>>
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
			                        <td width="10%" nowrap="nowrap"><div align="left"><strong>eCheck Payment Name</strong>:&nbsp;</div></td>
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
        <div id="CollapsiblePanel5" class="CollapsiblePanel">
            <div class="CollapsiblePanelTab1">
                <table width="100%">
                  <tr>
                            <td class="pcPanelTitle1">Step 5: Minimize fraud by enabling Centinel by CardinalCommerce</td>
                            <td width="24" class="pcPanelTitle1" align="right"><img src="images/expand.gif" width="19" height="19" alt="Expand Selection" /></td>
                  </tr>
                </table>
            </div>
            <div class="CollapsiblePanelContent">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td colspan="2">&nbsp;</td>
                </tr>
                <tr> 
                    <td colspan="2">Note: Additional charges apply. <a href="http://billing.cardinalcommerce.com/centinel/registration/frame_services.asp?RefId=PRDCTCART" target="_blank">Contact CardinalCommerce for more information &gt;&gt;</a></td>
                </tr>
                <% if intCentActive<>0 AND request("mode")<>"Edit" then %>
                    <tr>
                        <td colspan="2">Centinel has already been activated for this or another payment gateway. To edit its settings or remove it, simply activate this payment gateway and then click on the &quot;Modify&quot; button on the payment options summary page.</td>
                    </tr>
                <% else %>
                    <tr> 
                                    <td width="14%">&nbsp;</td>
                                    <td width="86%"><input name="pcPay_Cent_Active" type="checkbox" class="clearBorder" value="YES" <%if pcPay_Cent_Active=1 then%>checked<%end if%>>
                                    <strong> Enable Centinel for USAePay</strong></td>
                                </tr>
                                <% if trim(pcPay_Cent_TransactionURL)="" then
                                    pcPay_Cent_TransactionURL="https://centineltest.cardinalcommerce.com/maps/txns.asp"
                                end if %>
                                <tr> 
                                    <td nowrap="nowrap"><div align="left">Transaction Url:</div></td>
                                    <td><input name="pcPay_Cent_TransactionURL" size="60" maxlength="255" value="<%=pcPay_Cent_TransactionURL%>"></td>
                                </tr>
                                <% if pcPay_Cent_MerchantID<>"" then
                                    pcPay_Cent_MerchantID=replace(pcPay_Cent_MerchantID,"""","&quot;")
                                end if %>
                                <tr> 
                                    <td><div align="left">Merchant ID: </div></td>
                                    <td><input name="pcPay_Cent_MerchantID" size="35" maxlength="255" value="<%=pcPay_Cent_MerchantID%>"></td>
                                </tr>
                                <% if pcPay_Cent_ProcessorID<>"" then
                                    pcPay_Cent_ProcessorID=replace(pcPay_Cent_ProcessorID,"""","&quot;")
                                end if %>
                                <tr> 
                                    <td><div align="left">Processor ID: </div></td>
                                    <td><input name="pcPay_Cent_ProcessorId" size="35" maxlength="255" value="<%=pcPay_Cent_ProcessorID%>"></td>
                                </tr>
                                <tr> 
                                    <td><div align="left">Password: </div></td>
                                    <td><input name="pcPay_Cent_Password" size="35" maxlength="255" value="<%=pcPay_Cent_Password%>"></td>
                                </tr>
							<% end if %>
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
		var CollapsiblePanel5 = new Spry.Widget.CollapsiblePanel("CollapsiblePanel5", {contentIsOpen:false});
        <% end if %>
        //-->
        </script>
    </td>
</tr>
</table>
<!-- New View End --><% end if %>

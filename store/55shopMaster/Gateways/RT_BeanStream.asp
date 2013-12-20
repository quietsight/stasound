<%
'---Start BeanStream---
Function gwBeanStreamEdit()
	call opendb()
	'request gateway variables and insert them into the beanstream table
	query="SELECT pcPay_BS_MerchantID, pcPay_BS_MerchantPassword, pcPay_BS_TransType,pcPay_BS_VBV, pcPay_BS_Interac, pcPay_BS_cardTypes, pcPay_BS_CVC, pcPay_BS_AccountID, pcPay_BS_TestMode FROM pcPay_BeanStream Where pcPay_BS_ID=1;"

	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	pcPay_BS_MerchantID2=rs("pcPay_BS_MerchantID")
	pcPay_BS_MerchantID2=enDeCrypt(pcPay_BS_MerchantID2, scCrypPass)

	pcPay_BS_MerchantID=request.Form("pcPay_BS_MerchantID")
	if pcPay_BS_MerchantID="" then
		pcPay_BS_MerchantID=pcPay_BS_MerchantID2
	end if
	pcPay_BS_MerchantPassword2=rs("pcPay_BS_MerchantPassword")
	pcPay_BS_MerchantPassword2=enDeCrypt(pcPay_BS_MerchantPassword2, scCrypPass)
	pcPay_BS_MerchantPassword=request.Form("pcPay_BS_MerchantPassword")
	if pcPay_BS_MerchantPassword="" then
		pcPay_BS_MerchantPassword=pcPay_BS_MerchantPassword2
	end if
	
	pcPay_BS_AccountID2= rs("pcPay_BS_AccountID")
	pcPay_BS_AccountID2=enDeCrypt(pcPay_BS_AccountID2, scCrypPass)
	pcPay_BS_AccountID=request.Form("pcPay_BS_AccountID")
	if pcPay_BS_AccountID="" then
		pcPay_BS_AccountID=pcPay_BS_AccountID2
	end if
	set rs=nothing
	
	pcPay_BS_TransType = request.form("pcPay_BS_TransType")
	pcPay_BS_VBV = request.form("pcPay_BS_VBV")
	pcPay_BS_Interac = request.form("pcPay_BS_Interac")		
	pcPay_BS_cardTypes=request.Form("pcPay_BS_cardTypes")
	pcPay_BS_CVC= request.Form("pcPay_BS_CVC")
	pcPay_BS_TestMode=request.Form("pcPay_BS_TestMode")
	if pcPay_BS_TestMode="YES" then
		pcPay_BS_TestMode=1
	else
		pcPay_BS_TestMode=0
	end if

	pcPay_BS_MerchantID=enDeCrypt(pcPay_BS_MerchantID, scCrypPass)
	pcPay_BS_MerchantPassword=enDeCrypt(pcPay_BS_MerchantPassword, scCrypPass)
	pcPay_BS_AccountID=enDeCrypt(pcPay_BS_AccountID, scCrypPass)
	
	query="UPDATE pcPay_BeanStream SET pcPay_BS_MerchantID='"&pcPay_BS_MerchantID&"',pcPay_BS_MerchantPassword='"&pcPay_BS_MerchantPassword&"',pcPay_BS_TransType ='" &pcPay_BS_TransType &"',pcPay_BS_VBV =" & pcPay_BS_VBV &", pcPay_BS_Interac =" & pcPay_BS_Interac &",pcPay_BS_cardTypes ='"&pcPay_BS_cardTypes&"',pcPay_BS_CVC=" & pcPay_BS_CVC &",pcPay_BS_AccountID='" & pcPay_BS_AccountID&"',pcPay_BS_TestMode=" & pcPay_BS_TestMode & " WHERE pcPay_BS_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=57"
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

Function gwBeanStream()
	varCheck=1
	'request gateway variables and insert them into the Beanstream table
	pcPay_BS_MerchantID=request.Form("pcPay_BS_MerchantID")
	pcPay_BS_AccountID = request.form("pcPay_BS_AccountID")
	pcPay_BS_MerchantPassword=request.Form("pcPay_BS_MerchantPassword")
	pcPay_BS_TransType = request.form("pcPay_BS_TransType")
	pcPay_BS_VBV = request.form("pcPay_BS_VBV")
	pcPay_BS_Interac = request.form("pcPay_BS_Interac")			
	pcPay_BS_cardTypes=request.Form("pcPay_BS_cardTypes")
	pcPay_BS_CVC= request.Form("pcPay_BS_CVC")
	pcPay_BS_TestMode=request.Form("pcPay_BS_TestMode")
	if pcPay_BS_TestMode="" then
		pcPay_BS_TestMode=0
	end if

	If pcPay_BS_MerchantID="" OR pcPay_BS_MerchantPassword="" or pcPay_BS_AccountID ="" then
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add BeanStream Network as your payment gateway. <b>""Merchant ID""</b>, <b>""Account Login""</b> and <b>""Account Password""</b> are required fields.")
	End If
	'encrypt
	pcPay_BS_MerchantID=enDeCrypt(pcPay_BS_MerchantID, scCrypPass)
	pcPay_BS_AccountID=enDeCrypt(pcPay_BS_AccountID, scCrypPass)
	pcPay_BS_MerchantPassword=enDeCrypt(pcPay_BS_MerchantPassword, scCrypPass)
	
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
	call opendb()
	query="UPDATE pcPay_BeanStream SET pcPay_BS_MerchantID='"&pcPay_BS_MerchantID&"',pcPay_BS_MerchantPassword='"&pcPay_BS_MerchantPassword&"',pcPay_BS_TransType ='" &pcPay_BS_TransType&"',pcPay_BS_VBV=" & pcPay_BS_VBV &",pcPay_BS_Interac=" & pcPay_BS_Interac & ",pcPay_BS_cardTypes ='"&pcPay_BS_cardTypes&"',pcPay_BS_CVC=" & pcPay_BS_CVC &",pcPay_BS_AccountID='" & pcPay_BS_AccountID&"',pcPay_BS_TestMode=" & pcPay_BS_TestMode & " WHERE pcPay_BS_ID=1;"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'Beanstream','gwBeanStream.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",57,'"&paymentNickName&"')"
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

if request("gwchoice")="57" then
	'Check if table exist in database, if not we will rediect to database update prompt
	err.clear
	call openDb()
	query="select * from pcPay_BeanStream"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=connTemp.execute(query)
	if err.number<>0 then
		set rs=nothing
		call closedb()
		response.redirect "upddbBeanStream.asp"
	else
		set rs=nothing
		call closedb()
	end if

	if request("mode")="Edit" then
	 	call opendb()	
		query="SELECT pcPay_BS_MerchantID, pcPay_BS_MerchantPassword, pcPay_BS_TransType, pcPay_BS_VBV, pcPay_BS_Interac, pcPay_BS_cardTypes, pcPay_BS_CVC, pcPay_BS_AccountID, pcPay_BS_TestMode FROM pcPay_BeanStream Where pcPay_BS_ID=1;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		pcPay_BS_MerchantID=rs("pcPay_BS_MerchantID")
		pcPay_BS_MerchantID=enDeCrypt(pcPay_BS_MerchantID, scCrypPass)
		pcPay_BS_MerchantPassword=rs("pcPay_BS_MerchantPassword")
		pcPay_BS_MerchantPassword=enDeCrypt(pcPay_BS_MerchantPassword, scCrypPass)
		pcPay_BS_TransType = rs("pcPay_BS_TransType")
		pcPay_BS_VBV = rs("pcPay_BS_VBV")
		pcPay_BS_Interac = rs("pcPay_BS_Interac")
		pcPay_BS_cardTypes=rs("pcPay_BS_cardTypes")
		pcPay_BS_CVC=rs("pcPay_BS_CVC")
		pcPay_BS_AccountID = rs("pcPay_BS_AccountID")
		pcPay_BS_AccountID=enDeCrypt(pcPay_BS_AccountID, scCrypPass)
		pcPay_BS_TestMode=rs("pcPay_BS_TestMode")
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=57"
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
	<input type="hidden" name="addGw" value="57">
	<script language="javascript" >
    // function to check and change transaction type to sale if interac online is turned on
       function doInterac(intOnline){
        if(document.getElementById("tranType").options[1].selected == true) {
         if (confirm('In order to offer Interac Online the Tranaction Type field must be set to "Sale". Do you want to offer Interact Online and set the Tranaction Type field to "Sale"?')){			 
              intOnline.checked=true;
              document.getElementById("tranType").options[0].selected = true;		
          }else{				 
              intOnline.checked=false;
              document.getElementById("InteracNo").checked=true;
          }
        }		   
       }
       
    // function to check and change interec online to off if previously turned on and trans action type chnage to auth only.
        function checkInterac(intOnline){
        if(document.getElementById("InteracYes").checked==true && document.getElementById("tranType").options[1].selected == true){
        if(confirm('In order to offer Interac Online the Tranaction Type field must be set to "Sale". Changing the Tranaction Type field to "Authorize Only" will disable Interac Online. Do you want to change the Tranaction Type field to "Authorize Only" and disable Interact Online now?')){			 
              document.getElementById("InteracNo").checked=true;
              intOnline.options[1].selected = true;		
          }else{				 
              document.getElementById("InteracYes").checked=true;
              intOnline.options[0].selected = true;
          }
        }		   
       }
       
    </script>

<!-- END BeanStream -->

<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/beanstreamlogo.jpg" width="315" height="80"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
<table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
<tr>
    <td><strong>Beanstream electronic payment processing - Direct post API<br>
    <br>
    </strong>Are you looking for a simple, cost-effective way to streamline credit card acceptance?  Beanstream provides multiple, flexible options to help you simplify billing processes, manage online payments and more. <strong><br>
    <br>
    <a href="http://www.beanstream.com/site/ca/index.html">Beanstream Website</a></strong><br />
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
					dim pcPay_BS_MerchantIDCnt,pcPay_BS_MerchantIDEnd,pcPay_BS_MerchantIDStart
					pcPay_BS_MerchantIDCnt=(len(pcPay_BS_MerchantID)-2)
					pcPay_BS_MerchantIDEnd=right(pcPay_BS_MerchantID,2)
					pcPay_BS_MerchantIDStart=""
					for c=1 to pcPay_BS_MerchantIDCnt
						pcPay_BS_MerchantIDStart=pcPay_BS_MerchantIDStart&"*"
					next
					
					dim pcPay_BS_AccountIDCnt,pcPay_BS_AccountIDEnd,pcPay_BS_AccountIDStart
					pcPay_BS_AccountIDCnt=(len(pcPay_BS_AccountID)-2)
					pcPay_BS_AccountIDEnd=right(pcPay_BS_AccountID,2)
					pcPay_BS_AccountIDStart=""
					for c=1 to pcPay_BS_AccountIDCnt
						pcPay_BS_AccountIDStart=pcPay_BS_AccountIDStart&"*"
					next
					%>
					<tr> 
						<td height="31" colspan="2">Current Merchant ID:&nbsp;<%=pcPay_BS_MerchantIDStart&pcPay_BS_MerchantIDEnd%></td>
					</tr>
					<tr> 
						<td height="31" colspan="2">Current Account Login:&nbsp;<%=pcPay_BS_AccountIDStart&pcPay_BS_AccountIDEnd%></td>
					</tr>
					<tr> 
						<td colspan="2"> For security reasons, your &quot;Account Login and your Merchant ID&quot; 
							is only partially shown on this page. If you need 
							to edit your account information, please re-enter 
							your &quot;Merchant ID&quot;,&quot;Account Login&quot; and &quot;Account 
							Password&quot; below.</td>
					</tr>
                <% end if %>
                <tr>
                  <td width="93"> <div align="right">Merchant&nbsp;ID:</div></td>
                  <td width="328"> <input type="text" name="pcPay_BS_MerchantID" size="20">								</td>
                <tr>
                  <td width="93"> <div align="right">Account Login:</div></td>
                  <td width="328"> <input type="text" name="pcPay_BS_AccountID" size="20">								</td>
                </tr>
                <tr>
                  <td width="93"> <div align="right">Account Password:</div></td>
                  <td width="328"> <input type="text" name="pcPay_BS_MerchantPassword" size="20">								</td>
                </tr>
                <tr> 
                    <td colspan="2"><div align="left">
                      Create a new user from the Beanstream control panel. This will be your "Account Login" and "Account Password",to &ldquo;
You can create a login and password for ProductCart to use, by going to your Beanstream Order Settings page, beneath the heading &ldquo;Use username/password validation against transaction&rdquo;.
                      </div></td>
                </tr>
                <tr> 
                    <td> <div align="right">Transaction Type:</div></td>
                    <td> <select name="pcPay_BS_TransType" id="tranType" onChange="checkInterac(this)" >
                            <option value="P" selected>Sale</option>
                            <option value="PA" <% if pcPay_BS_TransType="PA" then%>selected<% end if %>>Authorize Only</option>
                        </select> </td>
                </tr>
                    <tr> 
                    <td> <div align="right">Accepted Cards:</div></td>
                    <td>
                        <input name="pcPay_BS_CardTypes" type="checkbox" class="clearBorder" value="VISA" <%if pcPay_BS_CardTypes="VISA" or (instr(pcPay_BS_CardTypes,"VISA,")>0) or (instr(pcPay_BS_CardTypes,", VISA")>0) then%>checked<%end if%>>
                        Visa 
                        <input name="pcPay_BS_CardTypes" type="checkbox" class="clearBorder" value="MAST" <%if pcPay_BS_CardTypes="MAST" or (instr(pcPay_BS_CardTypes,"MAST,")>0) or (instr(pcPay_BS_CardTypes,", MAST")>0) then%>checked<%end if%>>
                        MasterCard 
                        <input name="pcPay_BS_CardTypes" type="checkbox" class="clearBorder" value="AMER" <%if pcPay_BS_CardTypes="AMER" or (instr(pcPay_BS_CardTypes,"AMER,")>0) or (instr(pcPay_BS_CardTypes,", AMER")>0) then%>checked<%end if%>>
                        American Express 
                        <input name="pcPay_BS_CardTypes" type="checkbox" class="clearBorder" value="DINE" <%if pcPay_BS_CardTypes="DINE" or (instr(pcPay_BS_CardTypes,"DINE,")>0) or (instr(pcPay_BS_CardTypes,", DINE")>0) then%>checked<%end if%>>
                        Diner's Club 
                        <input name="pcPay_BS_CardTypes" type="checkbox" class="clearBorder" value="DISC" <%if pcPay_BS_CardTypes="DISC" or (instr(pcPay_BS_CardTypes,"DISC,")>0) or (instr(pcPay_BS_CardTypes,", DISC")>0) then%>checked<%end if%>>
                        Discover Card
                        <BR/><input name="pcPay_BS_cardTypes" type="checkbox" class="clearBorder" value="SEARS" <%if pcPay_BS_CardTypes="SEARS" or (instr(pcPay_BS_CardTypes,"SEARS,")>0) or (instr(pcPay_BS_CardTypes,", SEARS")>0) then%>checked<%end if%> >
                        Sears Card
                        </td>
                </tr>                           
                <tr> 
                    <td> <div align="right">Require CVC:</div></td>
                    <td> <input type="radio" class="clearBorder" name="pcPay_BS_CVC" value="1" checked>
                        Yes 
                        <input name="pcPay_BS_CVC" type="radio" class="clearBorder" value="0" <% if pcPay_BS_CVC="0" then%>checked<% end if %>>
                        No<font color="#FF0000"></font></td>
                </tr>
                 <tr> 
                    <td> <div align="right">Verify By Visa:</div></td>
                    <td> <input type="radio" class="clearBorder" name="pcPay_BS_VBV" value="1" checked>
                        Yes 
                        <input name="pcPay_BS_VBV" type="radio" class="clearBorder" value="0" <% if pcPay_BS_VBV="0" then%>checked<% end if %>>
                        No<font color="#FF0000"></font></td>
                </tr>
                 <tr> 
                    <td> <div align="right">Interac&copy; Online:</div></td>
                    <td> <input type="radio" class="clearBorder" name="pcPay_BS_Interac" value="1" onClick="doInterac(this)"  id="InteracYes" checked> Yes 
                        <input name="pcPay_BS_Interac" type="radio" class="clearBorder" value="0" id="InteracNo" <% if pcPay_BS_Interac="0" then%>checked<% end if %>>
                        No<font color="#FF0000"></font></td>
                </tr>
                <tr> 
                <td> <div align="right"> 
                        <input name="pcPay_BS_TestMode" type="checkbox" class="clearBorder" id="pcPay_BS_TestMode" value="1" <% if pcPay_BS_TestMode="1" then%>checked<% end if %>></div></td>
                <td><b>Enable Test Mode </b>(Credit cards will not be charged)</td>
                </tr>                  <tr>
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

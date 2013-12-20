<%
'---Start Concord---
Function gwconcordEdit()
	call opendb()
	'request gateway variables and insert them into the Concord table
	query= "SELECT StoreID,StoreKey,testmode,Curcode,CVV,MethodName FROM concord where idConcord=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	StoreID=request.Form("StoreID")
	StoreKey2=rs("StoreKey")
	'decrypt
	StoreKey2=enDeCrypt(StoreKey2, scCrypPass)
	StoreKey=request.Form("StoreKey")
	if StoreKey="" then
		StoreKey=StoreKey2
	end if
	'encrypt
	StoreKey=enDeCrypt(StoreKey, scCrypPass)
	CVV=request.Form("CVV")
	Curcode=request.Form("Curcode")
	testmode=request.Form("testmode")
	if testmode&""="" then
		testmode=0
	end if
	MethodName=request.Form("MethodName")
	query="UPDATE concord SET StoreID='"&StoreID&"',StoreKey='"&StoreKey&"',CVV='"&CVV&"',testmode="&testmode&",MethodName='"&MethodName&"' WHERE idConcord=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=22"

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

Function gwconcord()
	varCheck=1
	'request gateway variables and insert them into the concord table
	StoreID=request.Form("StoreID")
	StoreKey=request.Form("StoreKey")
	'encrypt
	StoreKey=enDeCrypt(StoreKey, scCrypPass)
	CVV=request.Form("CVV")
	Curcode=request.Form("Curcode")
	testmode=request.Form("testmode")
	MethodName=request.Form("MethodName")
	if testmode="YES" then
		testmode="1"
	else
		testmode="0"
	end if
	if NOT isNumeric(CVV) or CVV="" then
		CVV=0
	end if
	If StoreID="" then
		call closedb()
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Concord as your payment gateway. <b>""Store ID""</b> is a required field.")
	End If
	StoreKey=request.Form("StoreKey")
	If StoreKey="" then
		call closedb()
		response.redirect "AddRTPaymentOpt.asp?msg="&Server.URLEncode("An error occured while trying to add Concord as your payment gateway. <b>""Password""</b> is a required field.")
	End If
	'encrypt
	StoreKey=enDeCrypt(StoreKey, scCrypPass)
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

	query="UPDATE concord SET StoreID='"&StoreID&"',StoreKey='"&StoreKey&"',CVV="&CVV&",testmode="&testmode&",Curcode='"&Curcode&"', MethodName='"&MethodName&"' WHERE idConcord=1"
	set rs=Server.CreateObject("ADODB.Recordset")  

	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'Concord','gwconcord.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",22,'"&paymentNickName&"')"
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

if request("gwchoice")="22" then
	if request("mode")="Edit" then
		call opendb()
		query= "SELECT StoreID,StoreKey,Curcode,testmode,CVV,MethodName FROM concord WHERE idConcord=1"

		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error in modPaymentOpt: "&Err.Description) 
		end If
		StoreID=rs("StoreID")
		StoreKey=rs("StoreKey")
			'decrypt
			StoreKey=enDeCrypt(StoreKey, scCrypPass)
		CVV=rs("CVV")
		Curcode=rs("Curcode")
		testmode=rs("testmode")
		MethodName=rs("MethodName")
		set rs=nothing
		call closedb()
		%>
    
		<input type="hidden" name="mode" value="Edit">
    <% end if %>
	<input type="hidden" name="addGw" value="22">
    <tr> 
        <td height="21"><b>Concord EFSnet</b></td>
    </tr>
    <tr> 
        <td>
        	<table width="100%" border="0" cellspacing="0" cellpadding="4">
                <tr> 
                    <th colspan="2">Enter Concord EFSnet settings</th>
                </tr>
                <% if request("mode")="Edit" then %>
                
                	<% dim StoreIDCnt,StoreIDEnd,StoreIDStart
					StoreIDCnt=(len(StoreID)-2)
					StoreIDEnd=right(StoreID,2)
					StoreIDStart=""
					for c=1 to StoreIDCnt
						StoreIDStart=StoreIDStart&"*"
					next %>
                    <tr> 
                        <td height="31" colspan="2">Current StoreID:&nbsp;<%=StoreIDStart&StoreIDEnd%></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> For security reasons, your &quot;Store ID&quot; 
                        is only partially shown on this page. If you need 
                        to edit your account information, please re-enter 
                        your &quot;Store ID&quot; below.</td>
                    </tr>
                <% end if %>
                <tr bgcolor="#FFFFFF"> 
                    <td> <div align="right">Store ID:</div></td>
                    <td width="360"> <input type="text" name="StoreID" size="24"></td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td> <div align="right">Change StoreKey:</div></td>
                    <td width="360"> <input type="Password" value="" name="StoreKey" size="24"></td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td> <div align="right">Transaction Type:</div></td>
                    <td width="360">
                    	<select name="MethodName">
                            <option value="Authorize">Authorize Only</option>
                            <option value="Charge">Sale (Authorize and Settle)</option>
                        </select>
                	</td>
                </tr>
<tr bgcolor="#FFFFFF"> 
                    <td><div align="right">Currency Code:</div></td>
                    <td><input name="Curcode" type="hidden" id="Curcode" value="<%=Curcode%>"></td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td> <div align="right">Require CVV:</div></td>
                    <td> <input type="radio" class="clearBorder" name="CVV" value="1" checked>
                        Yes 
                        <input name="CVV" type="radio" class="clearBorder" value="0" <% if  CVV="0" then %>checked<%end if %>>
                        No</td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td> <div align="right"> 
                            <input name="testmode" type="checkbox" class="clearBorder" value="1" <% if testmode=1 then%>checked<% end if %>>
                        </div></td>
                    <td><b>Enable Test Mode </b>(Credit cards will not be charged)</td>
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
<% end if %>
<!-- END CONCORD -->

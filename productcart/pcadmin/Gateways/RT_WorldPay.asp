<%
'---Start WorldPay---
Function gwwpEdit()
	call opendb()
	'request gateway variables and insert them into the WorldPay table
	query= "SELECT WP_instID FROM WorldPay WHERE wp_id=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	WP_instID2=rs("WP_instID")
	set rs=nothing
	WP_instID=request.Form("WP_instID")
	If WP_instID="" then
		WP_instID=WP_instID2
	end if
	WP_Currency=request.Form("WP_Currency")
	WP_testmode=request.Form("WP_testmode")
	query="UPDATE WorldPay SET WP_instID='"&WP_instID&"', WP_Currency='"&WP_Currency&"',wp_testmode='"&WP_testmode&"' WHERE wp_id=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"' WHERE gwCode=10"
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

Function gwwp()
	varCheck=1
	'request gateway variables and insert them into the WorldPay table
	WP_instID=request.form("WP_instID")
	WP_Currency=request.Form("WP_Currency")
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
	WP_testmode=request.Form("wp_testmode")
	paymentNickName=replace(request.Form("paymentNickName"),"'","''")
	if paymentNickName="" then
		paymentNickName="Credit Card"
	End If
	
	err.clear
	err.number=0
	call openDb() 

	query="UPDATE WorldPay SET WP_instID='"&WP_instID&"',WP_Currency='"&WP_Currency&"',WP_testmode='"&wp_testmode&"' WHERE wp_id=1"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	query="INSERT INTO payTypes (pcPayTypes_processOrder,pcPayTypes_setPayStatus,paymentDesc, sslURL, active, Cbtob, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,paymentNickName) VALUES (" & pcv_processOrder & "," & pcv_setPayStatus & ",'WorldPay','gwwp.asp',-1,0,0,9999,0,9999,0,9999,-1,"& priceToAdd &","&percentageToAdd&",10,'"&paymentNickName&"')"
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
'---End WorldPay---
%>
<% if request("gwchoice")="10" then
	if request("mode")="Edit" then
		call opendb()
		query= "SELECT WP_instID,WP_Currency,wp_testmode FROM WorldPay WHERE wp_id=1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		WP_instID=rs("WP_instID")
		WP_Currency=rs("WP_Currency")
		wp_testmode=rs("WP_testmode")
		
		query= "SELECT pcPayTypes_processOrder, pcPayTypes_setPayStatus, priceToAdd, percentageToAdd, paymentNickName FROM payTypes WHERE gwCode=10"
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
	<input type="hidden" name="addGw" value="10">
<!-- New View Start -->
<table width="100%">
<tr>
    <td align="left" style="font-size:15px;"><img src="gateways/logos/worldpay_logo.jpg" width="260" height="79"></td>
    <td align="left" style="font-size:15px;">&nbsp;</td>
</tr>
</table>
<br>
<table width="100%" style="background-color:#FFF;border: solid 1px #CCC;">
<tr>
    <td><strong>Enter your Gateway Account information<br>
    <br>
    </strong>WorldPay's Hosted Payment Page service is secure, provides WorldPay with
required information to perform active fraud risk assessment, and is the fastest way
to get up and running with on-line payments.
 <strong><br>
    <br>
    <a href="http://www.worldpay.com" target="_blank">WorldPay Website</a></strong><br />
<br />
</td>
</tr>
<tr>
    <td>
        <div id="CollapsiblePanel1" class="CollapsiblePanel">
            <div class="CollapsiblePanelTab1">
                <table width="100%">
                  <tr>
                    <td width="580" class="pcPanelTitle1">Step 1: Configure Account - WorldPay Hosted Payment Page (HTML Redirect Service)</td>
                  </tr>
                </table>
            </div>
            <div class="CollapsiblePanelContent">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <% if request("mode")="Edit" then %>
				<% dim WP_instIDCnt,WP_instIDEnd,WP_instIDStart
                WP_instIDCnt=(len(WP_instID)-2)
                WP_instIDEnd=right(WP_instID,2)
                WP_instIDStart=""
                for c=1 to WP_instIDCnt
                    WP_instIDStart=WP_instIDStart&"*"
                next %>
                <tr> 
                    <td colspan="2">Current Installation ID:&nbsp;<%=WP_instIDStart&WP_instIDEnd%></td>
                </tr>
                <tr> 
                    <td colspan="2"> For security reasons, your &quot;Installation 
                        ID&quot; is only partially shown on this page. If 
                        you need to edit your account information, please 
                        re-enter your &quot;Installation ID&quot; below.</td>
                </tr>
			<% end if %>
            <tr>
              <td nowrap="nowrap">&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr> 
                <td width="10%" nowrap="nowrap"><div align="left">Installation ID:</div></td>
                <td width="90%"> <input type="text" name="WP_instID" size="20"></td>
            </tr>
            <tr> 
                <td> <div align="left">Currency:</div></td>
                <td>
                	<select name="WP_Currency">
                        <option value="AFA" selected>Afghani</option>
                        <option value="ALL" <% if WP_Currency="ALL" then%>selected<% end if %>>Lek</option>
                        <option value="DZD" <% if WP_Currency="DZD" then%>selected<% end if %>>Algerian Dinar</option>
                        <option value="AON" <% if WP_Currency="AON" then%>selected<% end if %>>New Kwanza</option>
                        <option value="ARS" <% if WP_Currency="ARS" then%>selected<% end if %>>Argentine Peso</option>
                        <option value="AWG" <% if WP_Currency="AWG" then%>selected<% end if %>>Aruban Guilder</option>
                        <option value="AUD" <% if WP_Currency="AUD" then%>selected<% end if %>>Australian Dollar</option>
                        <option value="EUR" <% if WP_Currency="EUR" then%>selected<% end if %>>Euro</option>
                        <option value="BSD" <% if WP_Currency="BSD" then%>selected<% end if %>>Bahamian Dollar</option>
                        <option value="BHD" <% if WP_Currency="BHD" then%>selected<% end if %>>Bahraini Dinar</option>
                        <option value="BDT" <% if WP_Currency="BDT" then%>selected<% end if %>>Taka</option>
                        <option value="BBD" <% if WP_Currency="BBD" then%>selected<% end if %>>Barbados Dollar</option>
                        <option value="EUR" <% if WP_Currency="EUR" then%>selected<% end if %>>Euro</option>
                        <option value="BZD" <% if WP_Currency="BZD" then%>selected<% end if %>>Belize Dollar</option>
                        <option value="BMD" <% if WP_Currency="BMD" then%>selected<% end if %>>Bermudian Dollar</option>
                        <option value="BOB" <% if WP_Currency="BOB" then%>selected<% end if %>>Boliviano</option>
                        <option value="BAD" <% if WP_Currency="BAD" then%>selected<% end if %>>Bosnian Dinar</option>
                        <option value="BWP" <% if WP_Currency="BWP" then%>selected<% end if %>>Pula</option>
                        <option value="BRL" <% if WP_Currency="BRL" then%>selected<% end if %>>Real</option>
                        <option value="BND" <% if WP_Currency="BND" then%>selected<% end if %>>Brunei Dollar</option>
                        <option value="BGL" <% if WP_Currency="BGL" then%>selected<% end if %>>Lev</option>
                        <option value="XOF" <% if WP_Currency="XOF" then%>selected<% end if %>>CFA Franc BCEAO</option>
                        <option value="BIF" <% if WP_Currency="BIF" then%>selected<% end if %>>Burundi Franc</option>
                        <option value="KHR" <% if WP_Currency="KHR" then%>selected<% end if %>>Cambodia Riel</option>
                        <option value="XAF" <% if WP_Currency="XAF" then%>selected<% end if %>>CFA Franc BEAC</option>
                        <option value="CAD" <% if WP_Currency="CAD" then%>selected<% end if %>>Canadian Dollar</option>
                        <option value="CVE" <% if WP_Currency="CVE" then%>selected<% end if %>>Cape Verde Escudo</option>
                        <option value="KYD" <% if WP_Currency="KYD" then%>selected<% end if %>>Cayman Islands Dollar</option>
                        <option value="CLP" <% if WP_Currency="CLP" then%>selected<% end if %>>Chilean Peso</option>
                        <option value="CNY" <% if WP_Currency="CNY" then%>selected<% end if %>>Yuan Renminbi</option>
                        <option value="COP" <% if WP_Currency="COP" then%>selected<% end if %>>Colombian Peso</option>
                        <option value="KMF" <% if WP_Currency="KMF" then%>selected<% end if %>>Comoro Franc</option>
                        <option value="CRC" <% if WP_Currency="CRC" then%>selected<% end if %>>Costa Rican Colon</option>
                        <option value="HRK" <% if WP_Currency="HRK" then%>selected<% end if %>>Croatian Kuna</option>
                        <option value="CUP" <% if WP_Currency="CUP" then%>selected<% end if %>>Cuban Peso</option>
                        <option value="CYP" <% if WP_Currency="CYP" then%>selected<% end if %>>Cyprus Pound</option>
                        <option value="CZK" <% if WP_Currency="CZK" then%>selected<% end if %>>Czech Koruna</option>
                        <option value="DKK" <% if WP_Currency="DKK" then%>selected<% end if %>>Danish Krone</option>
                        <option value="DJF" <% if WP_Currency="DJF" then%>selected<% end if %>>Djibouti Franc</option>
                        <option value="XCD" <% if WP_Currency="XCD" then%>selected<% end if %>>East Caribbean Dollar</option>
                        <option value="DOP" <% if WP_Currency="DOP" then%>selected<% end if %>>Dominican Peso</option>
                        <option value="TPE" <% if WP_Currency="TPE" then%>selected<% end if %>>Timor Escudo</option>
                        <option value="ECS" <% if WP_Currency="ECS" then%>selected<% end if %>>Ecuador Sucre</option>
                        <option value="EGP" <% if WP_Currency="EGP" then%>selected<% end if %>>Egyptian Pound</option>
                        <option value="SVC" <% if WP_Currency="SVC" then%>selected<% end if %>>El Salvador Colon</option>
                        <option value="EEK" <% if WP_Currency="EEK" then%>selected<% end if %>>Kroon</option>
                        <option value="ETB" <% if WP_Currency="ETB" then%>selected<% end if %>>Ethiopian Birr</option>
                        <option value="EUR" <% if WP_Currency="EUR" then%>selected<% end if %>>Euro</option>
                        <option value="FKP" <% if WP_Currency="FKP" then%>selected<% end if %>>Falkland Islands Pound</option>
                        <option value="FJD" <% if WP_Currency="FJD" then%>selected<% end if %>>Fiji Dollar</option>
                        <option value="EUR" <% if WP_Currency="EUR" then%>selected<% end if %>>Euro</option>
                        <option value="EUR" <% if WP_Currency="EUR" then%>selected<% end if %>>Euro</option>
                        <option value="XPF" <% if WP_Currency="XPF" then%>selected<% end if %>>CFP Franc</option>
                        <option value="GMD" <% if WP_Currency="GMD" then%>selected<% end if %>>Dalasi</option>
                        <option value="EUR" <% if WP_Currency="EUR" then%>selected<% end if %>>Euro</option>
                        <option value="GHC" <% if WP_Currency="GHC" then%>selected<% end if %>>Cedi</option>
                        <option value="GIP" <% if WP_Currency="GIP" then%>selected<% end if %>>Gibraltar Pound</option>
                        <option value="EUR" <% if WP_Currency="EUR" then%>selected<% end if %>>Euro</option>
                        <option value="GTQ" <% if WP_Currency="GTQ" then%>selected<% end if %>>Quetzal</option>
                        <option value="GNF" <% if WP_Currency="GNF" then%>selected<% end if %>>Guinea Franc</option>
                        <option value="GWP" <% if WP_Currency="GWP" then%>selected<% end if %>>Guinea - Bissau Peso</option>
                        <option value="GYD" <% if WP_Currency="GYD" then%>selected<% end if %>>Guyana Dollar</option>
                        <option value="HTG" <% if WP_Currency="HTG" then%>selected<% end if %>>Gourde</option>
                        <option value="HNL" <% if WP_Currency="HNL" then%>selected<% end if %>>Lempira</option>
                        <option value="HKD" <% if WP_Currency="HKD" then%>selected<% end if %>>Hong Kong Dollar</option>
                        <option value="HUF" <% if WP_Currency="HUF" then%>selected<% end if %>>Forint</option>
                        <option value="ISK" <% if WP_Currency="ISK" then%>selected<% end if %>>Iceland Krona</option>
                        <option value="INR" <% if WP_Currency="INR" then%>selected<% end if %>>Indian Rupee</option>
                        <option value="IDR" <% if WP_Currency="IDR" then%>selected<% end if %>>Rupiah</option>
                        <option value="IRR" <% if WP_Currency="IRR" then%>selected<% end if %>>Iranian Rial</option>
                        <option value="IQD" <% if WP_Currency="IQD" then%>selected<% end if %>>Iraqi Dinar</option>
                        <option value="EUR" <% if WP_Currency="EUR" then%>selected<% end if %>>Euro</option>
                        <option value="ILS" <% if WP_Currency="ILS" then%>selected<% end if %>>Shekel</option>
                        <option value="EUR" <% if WP_Currency="EUR" then%>selected<% end if %>>Euro</option>
                        <option value="JMD" <% if WP_Currency="JMD" then%>selected<% end if %>>Jamaican Dollar</option>
                        <option value="JPY" <% if WP_Currency="JPY" then%>selected<% end if %>>Yen</option>
                        <option value="JOD" <% if WP_Currency="JOD" then%>selected<% end if %>>Jordanian Dinar</option>
                        <option value="KZT" <% if WP_Currency="KZT" then%>selected<% end if %>>Tenge</option>
                        <option value="KES" <% if WP_Currency="KES" then%>selected<% end if %>>Kenyan Shilling</option>
                        <option value="KRW" <% if WP_Currency="KRW" then%>selected<% end if %>>Won</option>
                        <option value="KPW" <% if WP_Currency="KPW" then%>selected<% end if %>>North Korean Won</option>
                        <option value="KWD" <% if WP_Currency="KWD" then%>selected<% end if %>>Kuwaiti Dinar</option>
                        <option value="KGS" <% if WP_Currency="KGS" then%>selected<% end if %>>Som</option>
                        <option value="LAK" <% if WP_Currency="LAK" then%>selected<% end if %>>Kip</option>
                        <option value="LVL" <% if WP_Currency="LVL" then%>selected<% end if %>>Latvian Lats</option>
                        <option value="LBP" <% if WP_Currency="LBP" then%>selected<% end if %>>Lebanese Pound</option>
                        <option value="LSL" <% if WP_Currency="LSL" then%>selected<% end if %>>Loti</option>
                        <option value="LRD" <% if WP_Currency="LRD" then%>selected<% end if %>>Liberian Dollar</option>
                        <option value="LYD" <% if WP_Currency="LYD" then%>selected<% end if %>>Libyan Dinar</option>
                        <option value="LTL" <% if WP_Currency="LTL" then%>selected<% end if %>>Lithuanian Litas</option>
                        <option value="EUR" <% if WP_Currency="EUR" then%>selected<% end if %>>Euro</option>
                        <option value="MOP" <% if WP_Currency="MOP" then%>selected<% end if %>>Pataca</option>
                        <option value="MKD" <% if WP_Currency="MKD" then%>selected<% end if %>>Denar</option>
                        <option value="MGF" <% if WP_Currency="MGF" then%>selected<% end if %>>Malagasy Franc</option>
                        <option value="MWK" <% if WP_Currency="MWK" then%>selected<% end if %>>Kwacha</option>
                        <option value="MYR" <% if WP_Currency="MYR" then%>selected<% end if %>>Malaysian Ringitt</option>
                        <option value="MVR" <% if WP_Currency="MVR" then%>selected<% end if %>>Rufiyaa</option>
                        <option value="MTL" <% if WP_Currency="MTL" then%>selected<% end if %>>Maltese Lira</option>
                        <option value="MRO" <% if WP_Currency="MRO" then%>selected<% end if %>>Ouguiya</option>
                        <option value="MUR" <% if WP_Currency="MUR" then%>selected<% end if %>>Mauritius Rupee</option>
                        <option value="MXN" <% if WP_Currency="MXN" then%>selected<% end if %>>Mexico Peso</option>
                        <option value="MNT" <% if WP_Currency="MNT" then%>selected<% end if %>>Mongolia Tugrik</option>
                        <option value="MAD" <% if WP_Currency="MAD" then%>selected<% end if %>>Moroccan Dirham</option>
                        <option value="MZM" <% if WP_Currency="MZM" then%>selected<% end if %>>Metical</option>
                        <option value="MMK" <% if WP_Currency="MMK" then%>selected<% end if %>>Myanmar Kyat</option>
                        <option value="NAD" <% if WP_Currency="NAD" then%>selected<% end if %>>Namibian Dollar</option>
                        <option value="NPR" <% if WP_Currency="NPR" then%>selected<% end if %>>Nepalese Rupee</option>
                        <option value="ANG" <% if WP_Currency="ANG" then%>selected<% end if %>>Netherlands Antilles Guilder</option>
                        <option value="EUR" <% if WP_Currency="EUR" then%>selected<% end if %>>Euro</option>
                        <option value="NZD" <% if WP_Currency="NZD" then%>selected<% end if %>>New Zealand Dollar</option>
                        <option value="NIO" <% if WP_Currency="NIO" then%>selected<% end if %>>Cordoba Oro</option>
                        <option value="NGN" <% if WP_Currency="NGN" then%>selected<% end if %>>Naira</option>
                        <option value="NOK" <% if WP_Currency="NOK" then%>selected<% end if %>>Norwegian Krone</option>
                        <option value="OMR" <% if WP_Currency="OMR" then%>selected<% end if %>>Rial Omani</option>
                        <option value="PKR" <% if WP_Currency="PKR" then%>selected<% end if %>>Pakistan Rupee</option>
                        <option value="PAB" <% if WP_Currency="PAB" then%>selected<% end if %>>Balboa</option>
                        <option value="PGK" <% if WP_Currency="PGK" then%>selected<% end if %>>New Guinea Kina</option>
                        <option value="PYG" <% if WP_Currency="PYG" then%>selected<% end if %>>Guarani</option>
                        <option value="PEN" <% if WP_Currency="PEN" then%>selected<% end if %>>Nuevo Sol</option>
                        <option value="PHP" <% if WP_Currency="PHP" then%>selected<% end if %>>Philippine Peso</option>
                        <option value="PLN" <% if WP_Currency="PLN" then%>selected<% end if %>>New Zloty</option>
                        <option value="EUR" <% if WP_Currency="EUR" then%>selected<% end if %>>Euro</option>
                        <option value="QAR" <% if WP_Currency="QAR" then%>selected<% end if %>>Qatari Rial</option>
                        <option value="ROL" <% if WP_Currency="ROL" then%>selected<% end if %>>Leu</option>
                        <option value="RUB" <% if WP_Currency="RUB" then%>selected<% end if %>>Russian Ruble</option>
                        <option value="RWF" <% if WP_Currency="RWF" then%>selected<% end if %>>Rwanda Franc</option>
                        <option value="WST" <% if WP_Currency="WST" then%>selected<% end if %>>Tala</option>
                        <option value="STD" <% if WP_Currency="STD" then%>selected<% end if %>>Dobra</option>
                        <option value="SAR" <% if WP_Currency="SAR" then%>selected<% end if %>>Saudi Riyal</option>
                        <option value="SCR" <% if WP_Currency="SCR" then%>selected<% end if %>>Seychelles Rupee</option>
                        <option value="SLL" <% if WP_Currency="SLL" then%>selected<% end if %>>Leone</option>
                        <option value="SGD" <% if WP_Currency="SGD" then%>selected<% end if %>>Singapore Dollar</option>
                        <option value="SKK" <% if WP_Currency="SKK" then%>selected<% end if %>>Slovak Koruna</option>
                        <option value="SIT" <% if WP_Currency="SIT" then%>selected<% end if %>>Tolar</option>
                        <option value="SBD" <% if WP_Currency="SBD" then%>selected<% end if %>>Solomon Islands Dollar</option>
                        <option value="SOS" <% if WP_Currency="SOS" then%>selected<% end if %>>Somalia Shilling</option>
                        <option value="ZAR" <% if WP_Currency="ZAR" then%>selected<% end if %>>Rand</option>
                        <option value="EUR" <% if WP_Currency="EUR" then%>selected<% end if %>>Euro</option>
                        <option value="LKR" <% if WP_Currency="LKR" then%>selected<% end if %>>Sri Lanka Rupee</option>
                        <option value="SHP" <% if WP_Currency="SHP" then%>selected<% end if %>>St Helena Pound</option>
                        <option value="SDP" <% if WP_Currency="SDP" then%>selected<% end if %>>Sudanese Pound</option>
                        <option value="SRG" <% if WP_Currency="SRG" then%>selected<% end if %>>Suriname Guilder</option>
                        <option value="SZL" <% if WP_Currency="SZL" then%>selected<% end if %>>Swaziland Lilangeni</option>
                        <option value="SEK" <% if WP_Currency="SEK" then%>selected<% end if %>>Sweden Krona</option>
                        <option value="CHF" <% if WP_Currency="CHF" then%>selected<% end if %>>Swiss Franc</option>
                        <option value="SYP" <% if WP_Currency="SYP" then%>selected<% end if %>>Syrian Pound</option>
                        <option value="TWD" <% if WP_Currency="TWD" then%>selected<% end if %>>New Taiwan Dollar</option>
                        <option value="TJR" <% if WP_Currency="TJR" then%>selected<% end if %>>Tajik Ruble</option>
                        <option value="TZS" <% if WP_Currency="TZS" then%>selected<% end if %>>Tanzanian Shilling</option>
                        <option value="THB" <% if WP_Currency="THB" then%>selected<% end if %>>Baht</option>
                        <option value="TOP" <% if WP_Currency="TOP" then%>selected<% end if %>>Tonga Pa'anga</option>
                        <option value="TTD" <% if WP_Currency="TTD" then%>selected<% end if %>>Trinidad & Tobago Dollar</option>
                        <option value="TND" <% if WP_Currency="TND" then%>selected<% end if %>>Tunisian Dinar</option>
                        <option value="TRY" <% if WP_Currency="TRY" then%>selected<% end if %>>Turkish Lira</option>
                        <option value="UGX" <% if WP_Currency="UGX" then%>selected<% end if %>>Uganda Shilling</option>
                        <option value="UAH" <% if WP_Currency="UAH" then%>selected<% end if %>>Ukrainian Hryvnia</option>
                        <option value="AED" <% if WP_Currency="AED" then%>selected<% end if %>>United Arab Emirates Dirham</option>
                        <option value="GBP" <% if WP_Currency="GBP" then%>selected<% end if %>>Pounds Sterling</option>
                        <option value="USD" <% if WP_Currency="USD" then%>selected<% end if %>>US Dollar</option>
                        <option value="UYU" <% if WP_Currency="UYU" then%>selected<% end if %>>Uruguayan Peso</option>
                        <option value="VUV" <% if WP_Currency="VUV" then%>selected<% end if %>>Vanuatu Vatu</option>
                        <option value="VEB" <% if WP_Currency="VEB" then%>selected<% end if %>>Venezuela Bolivar</option>
                        <option value="VND" <% if WP_Currency="VND" then%>selected<% end if %>>Viet Nam Dong</option>
                        <option value="YER" <% if WP_Currency="YER" then%>selected<% end if %>>Yemeni Rial</option>
                        <option value="YUM" <% if WP_Currency="YUM" then%>selected<% end if %>>Yugoslavian New Dinar</option>
                        <option value="ZRN" <% if WP_Currency="ZRN" then%>selected<% end if %>>New Zaire</option>
                        <option value="ZMK" <% if WP_Currency="ZMK" then%>selected<% end if %>>Zambian Kwacha</option>
                        <option value="ZWD" <% if WP_Currency="ZWD" then%>selected<% end if %>>Zimbabwe Dollar</option>
					</select>
               </td>
            </tr>
            <tr> 
                <td> <div align="right"> 
                        <input name="wp_testmode" type="checkbox" class="clearBorder" value="YES" <% if wp_testmode="YES" then %>checked<% end if %> />
                    </div></td>
                <td>Enable Test Mode (Credit cards will not be charged)</td>
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
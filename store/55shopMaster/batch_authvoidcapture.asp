<%PmAdmin=9%>
<% 'void order
err.clear
data = "x_login=" & x_Login & ""   
data = data & "&x_tran_key=" & x_Password & ""
data = data & "&x_first_name=" & fName  & ""
data = data & "&x_last_name=" & lName  & ""
data = data & "&x_address=" & address  & ""
data = data & "&x_zip=" & zip  & ""
data = data & "&x_card_num=" & ccnum2  & ""
data = data & "&x_exp_date=" & ccexp & ""
data = data & "&x_amount=" & authamount  & ""
data = data & "&x_cust_id="
data = data & "&x_invoice_num=" & Request.Form("idOrder"&r)  & ""
data = data & "&x_description="
data = data & "&x_trans_id=" & gwTransId  & ""
data = data & "&x_delim_data=True"
data = data & "&x_delim_char=,"
data = data & "&x_relay_response=False"    
data = data & "&x_version=3.1"
data = data & "&x_method=CC"
data = data & "&x_type=VOID"
if x_testmode="1" then
	data = data & "&x_test_request=True"
end if  
data = data & "&x_email_customer=False"
set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
xml.open "POST", "https://secure.authorize.net/gateway/transact.dll?" _
		 & data & "", false
xml.send ""
authnetStatus = xml.Status
authnetVal = xml.responseText
set xml = nothing

authnetResults = split(authnetVal, ",", -1)
x_response_code           = authnetResults(0)
x_response_subcode        = authnetResults(1)
x_response_reason_code    = authnetResults(2)
x_response_reason_text    = authnetResults(3)
x_auth_code               = authnetResults(4)    
x_avs_code                = authnetResults(5)
x_trans_id                = authnetResults(6)   
x_invoice_num             = authnetResults(7)
x_description             = authnetResults(8)
x_amount                  = authnetResults(9)
x_method                  = authnetResults(10)
x_type                    = authnetResults(11)
x_cust_id                 = authnetResults(12)
x_first_name              = authnetResults(13)
x_last_name               = authnetResults(14)
x_company                 = authnetResults(15)
x_address                 = authnetResults(16)
x_city                    = authnetResults(17)
x_state                   = authnetResults(18)
x_zip                     = authnetResults(19)
x_country                 = authnetResults(20)
x_phone                   = authnetResults(21)
x_fax                     = authnetResults(22)
x_email                   = authnetResults(23)
x_ship_to_first_name      = authnetResults(24)
x_ship_to_last_name       = authnetResults(25)
x_ship_to_company         = authnetResults(26)
x_ship_to_address         = authnetResults(27)
x_ship_to_city            = authnetResults(28)
x_ship_to_state           = authnetResults(29)
x_ship_to_zip             = authnetResults(30)
x_ship_to_country         = authnetResults(31)
x_tax                     = authnetResults(32)
x_duty                    = authnetResults(33)
x_freight                 = authnetResults(34)
x_tax_exempt              = authnetResults(35)
x_po_num                  = authnetResults(36)
x_md5_hash                = authnetResults(37)
x_card_code               = authnetResults(38)
	
ordnum=(int(x_invoice_num)-scpre)

'if success add to success/void
if x_response_code = 1 then
	'CAPTURE NEW TRANSACTION, if no errors
	'Send the request to the Authorize.NET processor.
	data="x_Version=3.1"
	data=data & "&x_Delim_Data=True"
	if x_testmode="1" then
		data=data & "&x_Test_Request=True"
	else
		data=data & "&x_Test_Request=False"
	end if
	data=data & "&x_Login=" & x_Login
	data=data & "&x_tran_key=" & x_Password
	data=data & "&x_Amount=" & curamount
	data=data & "&x_Card_Num=" & ccnum2
	data=data & "&x_Exp_Date=" & ccexp
	data=data & "&x_relay_response=FALSE"
	data=data & "&x_Type=AUTH_CAPTURE"
	data=data & "&x_Currency_Code=" & x_Curcode
	'check these fields
	data=data & "&x_Description=" & replace(scCompanyName,",","-") & " Order: " & Request.Form("idOrder"&r)
	data=data & "&x_Invoice_Num=" & Request.Form("idOrder"&r)
	data=data & "&x_Cust_ID=" & idCustomer
	data=data & "&x_First_Name=" & fName
	data=data & "&x_Last_Name=" & lName
	billingAddress=address
	if address2<>"" then
		billingAddress=billingAddress&" "&address2
	end if
	data=data & "&x_Address=" & billingAddress
	data=data & "&x_City=" & City
	data=data & "&x_State=" & stateCode
	data=data & "&x_Zip=" & zip
	data=data & "&x_Country=" & countryCode
	data=data & "&x_Phone=" & phone
	data=data & "&x_EMail=" & email
	if shippingFullName<>"" then
		shippingNameArray=split(shippingFullName," ")
		tshippingFirstName=shippingNameArray(0)
		if ubound(shippingNameArray)>0 then
			tshippingLastName=shippingNameArray(1)
		else
			tshippingLastName=""
		end if
	else
		tshippingFirstName=""
		tshippingLastName=""
	end if
	data=data & "&x_Ship_To_First_Name=" & tshippingFirstName
	data=data & "&x_Ship_To_Last_Name=" & tshippingLastName
	data=data & "&x_Ship_To_Address=" & shippingAddress
	data=data & "&x_Ship_To_City=" & shippingCity
	data=data & "&x_Ship_To_State=" & shippingStateCode
	data=data & "&x_Ship_To_Zip=" & shippingZip
	data=data & "&x_Ship_To_Country=" & shippingCountryCode
	'Send the transaction info as part of the querystring
	set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
	xml.open "POST", "https://secure.authorize.net/gateway/transact.dll?"& data & "", false

	xml.send ""
	strStatus = xml.Status
	'store the response
	authnetVal = xml.responseText
	Set xml = Nothing
end if
err.clear %>

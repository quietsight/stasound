<%
if session("idCustomer")="" OR session("idCustomer")="0" then
	if isNumeric(pcv_strIdPayment) AND pcv_strIdPayment&""<>"" then
		query="SELECT gwCode FROM paytypes WHERE idPayment="&pcv_strIdPayment&";"
		set rs=Server.CreateObject("ADODB.Recordset")     
		set rs=connTemp.execute(query)
		if rs.eof then
			set rs=nothing
			call closedb()
			response.redirect "msg.asp?message=211" 
		end if
	
		pcv_intGwCode = rs("gwCode")
	
		Dim intCheckReferer
		intCheckReferer = 0
	
		select case pcv_intGwCode
			case "2", "3", "9", "53", "46", "4", "5", "8", "10", "11", "13" , "15", "16", "1", "19", "23", "22", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "35", "36", "37", "38", "39", "40", "42", "43", "44", "45", "47", "48", "49", "51", "52", "12", "54", "55", "56", "57", "58", "59", "60", "61", "62", "63" , "64" , "65" , "66", "7", "6"
				intCheckReferer = 1
		end select
		
		if intCheckReferer = 1 then
			referer = Request.ServerVariables("http_referer")
			if instr(referer, "?") then
				refererAry=split(referer, "?")
				referer = refererAry(0)
			end if
	
			if scSSL="1" then
				SSLTempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwsubmit.asp"),"//","/")
			end if
			TempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwsubmit.asp"),"//","/")
			
			TempURL=replace(TempURL,"//","/")
			TempURL=replace(TempURL,"http:/","http://")
			SSLTempURL=replace(SSLTempURL,"//","/")
			SSLTempURL=replace(SSLTempURL,"https:/","https://")
	
			if (scSSL="1" AND lcase(referer) = lcase(SSLTempURL)) OR (lcase(referer) = lcase(TempURL)) then
			else
				call closedb()
				response.redirect "msg.asp?message=211"  
			end if
		end if
	else
		referer = Request.ServerVariables("http_referer")
		if instr(referer, "?") then
			refererAry=split(referer, "?")
			referer = refererAry(0)
		end if
	
		if scSSL="1" then
			SSLTempURL=replace((scSslURL&"/"&scPcFolder&"/pc/pcModifyBillingInfo.asp"),"//","/")
		end if
		TempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/pcModifyBillingInfo.asp"),"//","/")
		
		TempURL=replace(TempURL,"//","/")
		TempURL=replace(TempURL,"http:/","http://")
		SSLTempURL=replace(SSLTempURL,"//","/")
		SSLTempURL=replace(SSLTempURL,"https:/","https://")
	
		if (scSSL="1" AND lcase(referer) = lcase(SSLTempURL)) OR (lcase(referer) = lcase(TempURL)) then
		else
			call closedb()
			response.redirect "msg.asp?message=211"
		end if
	end if
else
	'//Verify Session with idORder
	query="SELECT idCustomer From orders WHERE orders.idOrder="&pcGatewayDataIdOrder&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if NOT rs.eof then
		pcv_tempID=rs("idCustomer")
	end if
	set rs=nothing
	
	if isNumeric(pcv_tempID) AND pcv_tempID<>session("idCustomer") then
		call closedb()
		response.redirect "msg.asp?message=211"     
	end if
end if	%>
<% 'This file will contain all the processing info
'We are directed to this page to allow the customer to add any more info that may be required to checkout, such as catering/delivery information.%>
<%
if request.ServerVariables("CONTENT_LENGTH") > 0 then
	pord_OrderName=request.form("ord_OrderName")	
	if pord_OrderName<>"" then
	else
		pord_OrderName="No Name"
	end if
	session("pord_OrderName")=pord_OrderName
	
	If request.form("DF1") <> "" Then
		if scDateFrmt="DD/MM/YY" then
			expDateArray=split(request.form("DF1"),"/")
			session("DF1")=(expDateArray(1)&"/"&expDateArray(0)&"/"&expDateArray(2))
		else
			session("DF1")= month(request.form("DF1")) & "/" & day(request.form("DF1")) & "/" & year(request.form("DF1"))
		end if
	Else
		session("DF1")=""
	End If
	
	if session("DF1")<>"" then
		If Not IsDate(session("DF1")) then
			response.redirect "login.asp?msg="&Server.URLEncode(dictLanguage.Item(Session("language")&"_catering_2"))
		end if
		'Past Years
		If year(session("DF1"))<year(Date()) then
			response.redirect "login.asp?msg="&Server.URLEncode(dictLanguage.Item(Session("language")&"_catering_2"))
		end if
		'Same Year, Past Months
		If (year(session("DF1"))=year(Date())) and (month(session("DF1"))<month(Date())) then
			response.redirect "login.asp?msg="&Server.URLEncode(dictLanguage.Item(Session("language")&"_catering_2"))
		end if
		'Same Year, Same Month, Past Days
		If (year(session("DF1"))=year(Date())) and (month(session("DF1"))=month(Date())) and (day(session("DF1"))<day(date())) then
			response.redirect "login.asp?msg="&Server.URLEncode(dictLanguage.Item(Session("language")&"_catering_2"))
		end if
	end if
	
	session("TF1")=request.form("TF1")
	
	if session("TF1")<>"" then
		If Not IsDate(session("TF1")) then
			response.redirect "login.asp?msg="&Server.URLEncode(dictLanguage.Item(Session("language")&"_catering_2"))
		else
			'Current Date
			If (year(session("DF1"))=year(Date())) and (month(session("DF1"))=month(Date())) and (day(session("DF1"))=day(date())) then
				TF2=CDate(session("TF1"))
				if TF2-time()<=0 then
					response.redirect "login.asp?msg="&Server.URLEncode(dictLanguage.Item(Session("language")&"_catering_2"))
				end if
			end if
		end if
	end if
	
	If DTCheck="1" then
		if session("DF1")<>"" then
			DF2=CDate(session("DF1"))
			if DF2-Date()<=0 then
				response.redirect "login.asp?msg="&Server.URLEncode(dictLanguage.Item(Session("language")&"_catering_3"))
			else
				if (DF2-Date()=1) then
					if session("TF1")<>"" then
						TF2=CDate(session("TF1"))
						if TF2<time() then
							response.redirect "login.asp?msg="&Server.URLEncode(dictLanguage.Item(Session("language")&"_catering_3"))
						end if
					end if
				end if
			end if	
		else
			if session("TF1")<>"" then
				TF2=CDate(session("TF1"))
				if TF2-time()<24 then
					response.redirect "login.asp?msg="&Server.URLEncode(dictLanguage.Item(Session("language")&"_catering_3"))
				end if
			end if
		end if
	end if
end if

If request("msg") = "" Then
	if session("DF1")<>"" and IsDate(session("DF1")) Then
		call opendb()
		if SQL_Format="1" then
			expDateArray=split(session("DF1"),"/")
			tmpDFDate=(expDateArray(1)&"/"&expDateArray(0)&"/"&expDateArray(2))
		else
			tmpDFDate=session("DF1")
		end if
		query="SELECT [blackout_message] from blackout WHERE blackout_date="
		if scDB="SQL" then
			query=query&"'" & tmpDFDate  & "'"
		else
			query=query&"#" & tmpDFDate  & "#"
		end if	
		set rsQ=conntemp.execute(query)
		icounter = 0
		if not rsQ.eof then
			icounter = icounter + 1
			blackoutmessage = rsQ(0)
		end if
		set rsQ=nothing
		if icounter > 0 then
			call closeDb()
			response.redirect "login.asp?msg=" & blackoutmessage & Server.URLEncode(dictLanguage.Item(Session("language")&"_catering_5"))
		end if
		call closeDb()
	End If
End If

If DeliveryZip = "1" Then
	call opendb()
	query="SELECT * from zipcodevalidation WHERE zipcode='" &pcStrShippingPostalCode& "'"
	set rsZipCodeObj=server.CreateObject("ADODB.RecordSet")
	set rsZipCodeObj=conntemp.execute(query)
	if rsZipCodeObj.eof then
		set rsZipCodeObj=nothing
		call closeDb()
		response.redirect "login.asp?msg="&Server.Urlencode(dictLanguage.Item(Session("language")&"_Custmoda_23"))
	end if
	call closeDb()	
End If
%>

<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/validation.asp"--> 
<!--#include file="../includes/securitysettings.asp" -->
<!--#include file="../includes/emailsettings.asp"--> 
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/MailUpFunctions.asp"-->
<!--#include file="pcStartSession.asp" -->
<!--#include file="opc_contentType.asp" -->

<% On Error Resume Next

Call SetContentType()

'MailUp-S

Dim MaxRequestTime,StopHTTPRequests

'maximum seconds for each HTTP request time
MaxRequestTime=5

StopHTTPRequests=0

'MailUp-E

'MAILUP-S

	tmp_setup=0
	pcMailUpSett_APIUser=""
	pcMailUpSett_APIPassword=""
	pcMailUpSett_URL=""

	call opendb()
	query="SELECT pcMailUpSett_APIUser,pcMailUpSett_APIPassword,pcMailUpSett_URL,pcMailUpSett_AutoReg,pcMailUpSett_RegSuccess FROM pcMailUpSettings;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcMailUpSett_APIUser=rs("pcMailUpSett_APIUser")
		session("SF_MU_APIUser")=pcMailUpSett_APIUser
		pcMailUpSett_APIPassword=enDeCrypt(rs("pcMailUpSett_APIPassword"), scCrypPass)
		session("SF_MU_APIPassword")=pcMailUpSett_APIPassword
		pcMailUpSett_URL=rs("pcMailUpSett_URL")
		session("SF_MU_URL")=pcMailUpSett_URL
		tmp_Auto=rs("pcMailUpSett_AutoReg")
		if IsNull(tmp_Auto) or tmp_Auto="" then
			tmp_Auto=0
		end if
		session("SF_MU_Auto")=tmp_Auto
		tmp_setup=rs("pcMailUpSett_RegSuccess")
		if IsNull(tmp_setup) or tmp_setup="" then
			tmp_setup=0
		end if
		session("SF_MU_Setup")=tmp_setup
	end if
	set rs=nothing
	call closedb()

'MAILUP-E

call openDb()

Dim pcv_strCatcher
pcv_strCatcher = Session("pcCartIndex")
If pcv_strCatcher=0 Then
	pcv_strCatcher=""		
	response.clear
	Call SetContentType()
	response.write "SECURITY"
	response.End
End If

Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")

%><!--#include file="../includes/pcServerSideValidation.asp" --><%

Function generatePassword(passwordLength)
	Dim sDefaultChars
	Dim iCounter
	Dim sMyPassword
	Dim iPickedChar
	Dim iDefaultCharactersLength
	Dim iPasswordLength
	
	sDefaultChars="abcdefghijklmnopqrstuvxyzABCDEFGHIJKLMNOPQRSTUVXYZ0123456789"
	iPasswordLength=passwordLength
	iDefaultCharactersLength = Len(sDefaultChars) 
	Randomize
	For iCounter = 1 To iPasswordLength
		iPickedChar = Int((iDefaultCharactersLength * Rnd) + 1) 
		sMyPassword = sMyPassword & Mid(sDefaultChars,iPickedChar,1)
	Next 
	generatePassword = sMyPassword
End Function

pcErrMsg=""

'//////////////////////////////////////////////////////////////////////////
'// START: VALIDATE BILLING
'//////////////////////////////////////////////////////////////////////////

pcStrBillingFirstName=URLDecode(getUserInput(request("billfname"),50))
pcStrBillingLastName=URLDecode(getUserInput(request("billlname"),50))
pcStrBillingCompany=URLDecode(getUserInput(request("billcompany"),150))
pcStrBillingVATID=URLDecode(getUserInput(request("billVATID"),150))
pcStrBillingSSN=URLDecode(getUserInput(request("billSSN"),150))
pcStrBillingPhone=URLDecode(getUserInput(request("billphone"),20))
pcStrCustomerEmail=URLDecode(getUserInput(request("billemail"),150))
pcStrBillingAddress=URLDecode(getUserInput(request("billaddr"),255))
pcStrBillingPostalCode=URLDecode(getUserInput(request("billzip"),10))
pcStrBillingStateCode=URLDecode(getUserInput(request("billstate"),4))
pcStrBillingProvince=URLDecode(getUserInput(request("billprovince"),150))
pcStrBillingCity=URLDecode(getUserInput(request("billcity"),50))
pcStrBillingCountryCode=URLDecode(getUserInput(request("billcountry"),4))
pcStrBillingAddress2=URLDecode(getUserInput(request("billaddr2"),255))
pcStrBillingFax=URLDecode(getUserInput(request("billfax"),20))
pcIntShippingResidential=URLDecode(getUserInput(request("pcAddressType"),0))

if pcIntShippingResidential<>"" then
	if not IsNumeric(pcIntShippingResidential) then
		pcIntShippingResidential="1"
	end if
end if
pcStrNewPass1=""
if scGuestCheckoutOpt=2 then
	pcStrNewPass1=URLDecode(getUserInput(request("billpass"),250))
	pcStrNewPass2=URLDecode(getUserInput(request("billrepass"),250))
end if

'Check the PostalCode Length for United States
If pcStrBillingCountryCode="US" Then
	if len(pcStrBillingPostalCode)<5 then
		response.clear
		Call SetContentType()
		response.Write("ZIPLENGTH")
		response.End()
	end if
End If

if pcStrBillingFirstName="" then
	pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_58")&"</li>"
end if
if pcStrBillingLastName="" then
	pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_59")&"</li>"
end if
if pcStrBillingAddress="" then
	pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_60")&"</li>"
end if
if pcStrBillingCity="" then
	pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_61")&"</li>"
end if
if pcStrBillingCountryCode="" then
	pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_62")&"</li>"
end if
if (pcStrBillingCountryCode="US") OR (pcStrBillingCountryCode="CA") then
	if pcStrBillingStateCode="" then
		pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_63")&"</li>"
	end if
	if pcStrBillingPostalCode="" then
		pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_64")&"</li>"
	end if
end if
if session("idCustomer")="" OR session("idCustomer")=0 then
	if pcStrCustomerEmail="" then
		pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_65")&"</li>"
	else
		pcStrCustomerEmail=replace(pcStrCustomerEmail," ","")
		if instr(pcStrCustomerEmail,"@")=0 or instr(pcStrCustomerEmail,".")=0 then 
			pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_66")&"</li>"
		end if
	end if
	
	if scGuestCheckoutOpt=2 then
		if pcErrMsg="" then
			query="SELECT idCustomer FROM Customers WHERE [email] like '" & pcStrCustomerEmail & "';"
			set rs=connTemp.execute(query)
			if not rs.eof then
				pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_5a")&"</li>"
			end if
			set rs=nothing
		end if
	
		if pcStrNewPass1="" then
			pcErrMsg=pcErrMsg & "<li>" & dictLanguage.Item(Session("language")&"_opc_createacc_5") & "</li>"
		end if

		if pcStrNewPass2="" then
			pcErrMsg=pcErrMsg & "<li>" & dictLanguage.Item(Session("language")&"_opc_createacc_4") & "</li>"
		end if

		if pcStrNewPass1<>pcStrNewPass2 then
			pcErrMsg=pcErrMsg & "<li>" & dictLanguage.Item(Session("language")&"_opc_createacc_3") & "</li>"
		end if
	end if
	
end if
if session("idCustomer")>"0" then
	query="SELECT idCustomer FROM Customers WHERE idcustomer=" & session("idCustomer") & ";"
	set rs=connTemp.execute(query)
	if rs.eof then
		pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_67")&"</li>"
	end if
end if

'MAILUP-S
	IF session("SF_MU_Setup")="1" THEN
	tmp_DontShowMailUp=1
	Session("pcSFpcNewsListCount")=""
	tmpNewsListCount=getUserInput(request("newslistcount"),0)
	if tmpNewsListCount<>"" then
		tmp_DontShowMailUp=0
		Session("pcSFpcNewsListCount")=tmpNewsListCount
		Session("pcSFCRecvNews")="0"
		For j=0 to tmpNewsListCount
			Session("pcSFpcNewsList" & j)=getUserInput(request("newslist" & j),0)
			if Session("pcSFpcNewsList" & j)<>"" then
				Session("pcSFCRecvNews")="1"
			end if
		Next
	end if
	ELSE
		tmpRecvNews=URLDecode(getUserInput(request("CRecvNews"),0))
		if tmpRecvNews="" then
			tmpRecvNews=0
		end if
		if not IsNumeric(tmpRecvNews) then
			tmpRecvNews=0
		end if
		Session("pcSFCRecvNews")=tmpRecvNews
	END IF
'MAILUP-E

'//////////////////////////////////////////////////////////////////////////
'// END: VALIDATE BILLING
'//////////////////////////////////////////////////////////////////////////



'//////////////////////////////////////////////////////////////////////////
'// START: UPDATE BILLING
'//////////////////////////////////////////////////////////////////////////
if pcErrMsg="" then

	tmpIDRefer=URLDecode(getUserInput(request("IDRefer"),0))
	if tmpIDRefer<>"" then
		if not IsNumeric(tmpIDRefer) then
			tmpIDRefer=0
		end if
	else
		if session("idCustomer")="0" OR session("idCustomer")="" then
			tmpIDRefer=0
		end if
	end if
	Session("pcSFIDrefer")=tmpIDRefer
		
	if session("idCustomer")>"0" then

		if session("CustomerGuest") = "" OR isNULL(session("CustomerGuest")) then
			session("CustomerGuest") = 0
		end if
		query="UPDATE Customers SET [name]='" & pcStrBillingFirstName & "',lastName='" & pcStrBillingLastName & "',customerCompany='" & pcStrBillingCompany & "', pcCust_VATID='" & pcStrBillingVATID & "', pcCust_SSN='" & pcStrBillingSSN & "', phone='" & pcStrBillingPhone & "',address='" & pcStrBillingAddress & "',zip='" & pcStrBillingPostalCode & "',stateCode='" & pcStrBillingStateCode & "',state='" & pcStrBillingProvince & "',city='" & pcStrBillingCity & "',countryCode='" & pcStrBillingCountryCode & "',address2='" & pcStrBillingAddress2 & "',fax='" & pcStrBillingFax & "'"
		if session("CustomerGuest")>"0" then
			query=query & ",email='" & pcStrCustomerEmail & "'"
		end if
		query=query & " WHERE idcustomer=" & session("idCustomer") & " AND pcCust_Guest=" & session("CustomerGuest") & ";"
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closeDb()
			response.clear
			Call SetContentType()
			response.write "ERROR"
			response.End
		end if
		set rs=nothing
		OKmsg="OK"
		
	else
		if pcStrNewPass1<>"" then
			pcPassword=pcStrNewPass1
			tmpCustomerGuest=0
		else
			pcPassword=generatePassword(10)
			tmpCustomerGuest=1
		end if
		if scGuestCheckoutOpt=0 OR scGuestCheckoutOpt="" OR scGuestCheckoutOpt=1 then
			tmpCustomerGuest=1
		end if
		pcPassword=enDeCrypt(pcPassword, scCrypPass)
		if tmpCustomerGuest=0 then
		query="SELECT idCustomer FROM Customers WHERE [email] like '" & pcStrCustomerEmail & "' AND pcCust_Guest<>1;"
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closeDb()
			response.clear
			Call SetContentType()
			response.write "ERROR"
			response.End
		end if
		if not rs.eof then
			tmpCustomerGuest=2
		end if
		set rs=nothing
		end if

		pcv_dateCustomerRegistration=Date()
		if SQL_Format="1" then
			pcv_dateCustomerRegistration=Day(pcv_dateCustomerRegistration)&"/"&Month(pcv_dateCustomerRegistration)&"/"&Year(pcv_dateCustomerRegistration)
		else
			pcv_dateCustomerRegistration=Month(pcv_dateCustomerRegistration)&"/"&Day(pcv_dateCustomerRegistration)&"/"&Year(pcv_dateCustomerRegistration)
		end if
		
		query="INSERT INTO Customers (Password,customerType,email,[name],lastName,customerCompany,pcCust_VATID,pcCust_SSN,phone,address,zip,stateCode,state,city,countryCode,address2,fax,pcCust_DateCreated,pcCust_Guest) VALUES ('" & pcPassword & "',0,'" & pcStrCustomerEmail & "','" & pcStrBillingFirstName & "','" & pcStrBillingLastName & "','" & pcStrBillingCompany & "','" & pcStrBillingVATID & "','" & pcStrBillingSSN & "','" & pcStrBillingPhone & "','" & pcStrBillingAddress & "','" & pcStrBillingPostalCode & "','" & pcStrBillingStateCode & "','" & pcStrBillingProvince & "','" & pcStrBillingCity & "','" & pcStrBillingCountryCode & "','" & pcStrBillingAddress2 & "','" & pcStrBillingFax & "','" & pcv_dateCustomerRegistration & "'," & tmpCustomerGuest & ");"
		set rsBA=connTemp.execute(query)		
		query="SELECT TOP 1 idCustomer,pcCust_Guest,[name], lastName, email FROM Customers WHERE email like '" & pcStrCustomerEmail & "' ORDER BY idCustomer DESC;"
		set rsBA=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsBA=nothing
			call closeDb()
			response.clear
			Call SetContentType()
			response.write "ERROR"
			response.End
		end if
		if not rsBA.eof then
			session("idCustomer")=rsBA("idCustomer")
			session("CustomerGuest")=rsBA("pcCust_Guest")
			session("CustomerType")=0
			pcStrBillingFirstName = rsBA("name")
			pcStrBillingLastName = rsBA("lastName")
			pcStrCustomerEmail = rsBA("email")
			'// Send New Customer Emails
			pcv_strNoticeNewCust="1" '// Send to Admin
			If session("CustomerGuest")="0" Then
				pcv_strNewCustEmail="1" '// Send to Customer
			End If
			%> <!--#include file="adminNewCustEmail.asp"--> <%
		else
			pcErrMsg=dictLanguage.Item(Session("language")&"_opc_68")
		end if
		set rsBA=nothing
		OKmsg="NEW"
	end if
	if pcErrMsg="" then
		if session("OPCstep")<"2" then
			session("OPCstep")=2
		end if
	end if
end if
if pcErrMsg<>"" then
	pcErrMsg=dictLanguage.Item(Session("language")&"_opc_69")&"<br><ul>" & pcErrMsg & "</ul>"
	response.write pcErrMsg
else
	call opendb()
	
	'MAILUP-S
	query="SELECT [email] FROM customers WHERE idCustomer=" & Session("idCustomer") & ";"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcStrCustomerEmail2=rs("email")
	end if
	set rs=nothing
	
	MUResult=1
	tmpNewsListCount=Session("pcSFpcNewsListCount")
	if tmpNewsListCount<>"" then
		For j=0 to tmpNewsListCount
			if Session("pcSFpcNewsList" & j)<>"" then
				query="SELECT pcMailUpLists_ListID,pcMailUpLists_ListGuid FROM pcMailUpLists WHERE pcMailUpLists_ID=" & Session("pcSFpcNewsList" & j) & ";"
				set rs=connTemp.execute(query)
				ListID=rs("pcMailUpLists_ListID")
				ListGuid=rs("pcMailUpLists_ListGuid")
				tmpMUResult=UpdUserReg(Session("idCustomer"),pcStrCustomerEmail2,ListID,ListGuid,session("SF_MU_URL"),session("SF_MU_Auto"))
				if tmpMUResult=0 then
					MUResult=0
				end if
		end if
	Next
	query="SELECT pcMailUpLists_ID FROM pcMailUpSubs WHERE idCustomer=" & Session("idCustomer") & ";"
	set rs=connTemp.execute(query)
	if not rs.eof then
		tmpArr=rs.getRows()
		intCount=ubound(tmpArr,2)
		For j=0 to intCount
			tmpRmv=1
			For k=0 to tmpNewsListCount
				if Session("pcSFpcNewsList" & k)<>"" then
					if Clng(Session("pcSFpcNewsList" & k))=Clng(tmpArr(0,j)) then
						tmpRmv=0
						exit for
					end if
				end if
			Next
			if tmpRmv=1 then
				query="SELECT pcMailUpLists_ListID,pcMailUpLists_ListGuid FROM pcMailUpLists WHERE pcMailUpLists_ID=" & tmpArr(0,j) & ";"
				set rs=connTemp.execute(query)
				ListID=rs("pcMailUpLists_ListID")
				ListGuid=rs("pcMailUpLists_ListGuid")
				tmpMUResult=UnsubUser(Session("idCustomer"),pcStrCustomerEmail2,ListID,ListGuid,session("SF_MU_URL"),session("SF_MU_Auto"))
				if tmpMUResult=0 then
					MUResult=0
				end if
			end if
		Next
	end if
	set rs=nothing
	end if
	if err.number<>0 then
		err.number=0
		err.description=""
	end if
	'MAILUP-E
	
	'MAILUP-S: MailUp Lists, show it for new customer and when existing customers edit their account
	strNewMailUpArea=""
	IF (session("SF_MU_Setup")="1" AND Session("idCustomer")<>0) OR (session("SF_MU_Setup")="1" AND Session("idCustomer")=0 AND ((NewsCheckout="1") OR (NewsReg="1"))) THEN
		query="SELECT pcMailUpLists_ID,pcMailUpLists_ListID,pcMailUpLists_ListGuid,pcMailUpLists_ListName,pcMailUpLists_ListDesc,0 FROM pcMailUpLists WHERE pcMailUpLists_Active>0 and pcMailUpLists_Removed=0;"
		set rs=connTemp.execute(query)
		if not rs.eof then
			pcv_TurnMUOn=1
			strNewMailUpArea="<table>"
			tmpArr=rs.getRows()
			set rs=nothing
			intCount=ubound(tmpArr,2)
			tmpNListChecked=0
			pcv_MUSynError=0
			if Session("idCustomer")<>0 then
			'Synchronizing
				For j=0 to intCount
					tmpResult=CheckUserStatus(Session("idCustomer"),pcStrCustomerEmail2,tmpArr(1,j),tmpArr(2,j),session("SF_MU_URL"),session("SF_MU_Auto"))
					if tmpResult="-1" then
						pcv_MUSynError=1
						exit for
					else
						if tmpResult="2" then
							query="SELECT pcMailUpSubs_ID FROM pcMailUpSubs WHERE idCustomer=" & Session("idCustomer") & " AND pcMailUpLists_ID=" & tmpArr(0,j) & ";"
							set rs=connTemp.execute(query)
							dtTodaysDate=Date()
							if SQL_Format="1" then
								dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
							else
								dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
							end if
							if not rs.eof then
								if scDB="SQL" then
									query="UPDATE pcMailUpSubs SET idCustomer=" & Session("idCustomer") & ",pcMailUpLists_ID=" & tmpArr(0,j) & ",pcMailUpSubs_LastSave='" & dtTodaysDate & "',pcMailUpSubs_SyncNeeded=0,pcMailUpSubs_Optout=0 WHERE idCustomer=" & Session("idCustomer") & " AND pcMailUpLists_ID=" & tmpArr(0,j) & ";"
								else
									query="UPDATE pcMailUpSubs SET idCustomer=" & Session("idCustomer") & ",pcMailUpLists_ID=" & tmpArr(0,j) & ",pcMailUpSubs_LastSave=#" & dtTodaysDate & "#,pcMailUpSubs_SyncNeeded=0,pcMailUpSubs_Optout=0 WHERE idCustomer=" & Session("idCustomer") & " AND pcMailUpLists_ID=" & tmpArr(0,j) & ";"
								end if
							else
							if scDB="SQL" then
								query="INSERT INTO pcMailUpSubs (idCustomer,pcMailUpLists_ID,pcMailUpSubs_LastSave,pcMailUpSubs_SyncNeeded,pcMailUpSubs_Optout) VALUES (" & Session("idCustomer") & "," & tmpArr(0,j) & ",'" & dtTodaysDate & "',0,0);"
							else
								query="INSERT INTO pcMailUpSubs (idCustomer,pcMailUpLists_ID,pcMailUpSubs_LastSave,pcMailUpSubs_SyncNeeded,pcMailUpSubs_Optout) VALUES (" & Session("idCustomer") & "," & tmpArr(0,j) & ",#" & dtTodaysDate & "#,0,0);"
							end if
						end if
						set rs=nothing
						set rs=connTemp.execute(query)
						set rs=nothing
					end if
					if tmpResult="4" then
						tmpArr(5,j)=4
					end if
						if tmpResult="1" or tmpResult="3" then
							query="DELETE FROM pcMailUpSubs WHERE idCustomer=" & Session("idCustomer") & " AND pcMailUpLists_ID=" & tmpArr(0,j) & ";"
							set rs=connTemp.execute(query)
							set rs=nothing
						end if
					end if
				Next
				For j=0 to intCount
					query="SELECT idcustomer FROM pcMailUpSubs WHERE idcustomer=" & Session("idCustomer") & " AND pcMailUpLists_ID=" & tmpArr(0,j) & " AND pcMailUpSubs_Optout=0;"
					set rs=connTemp.execute(query)
					tmpOptedIn=0
					if not rs.eof then
						tmpOptedIn=1
						tmpNListChecked=1
					end if
					set rs=nothing
					if tmpArr(5,j)<>4 then
						tmpArr(5,j)=tmpOptedIn
					end if
				Next
			end if
			if pcv_MUSynError=1 then
				strNewMailUpArea=strNewMailUpArea & "<tr><td colspan=""4""><div class=""pcErrorMessage"">" & dictLanguage.Item(Session("language")&"_MailUp_SynNote1") & "</div></td></tr>"
			end if
	
			strNewMailUpArea=strNewMailUpArea & "<tr><td colspan=""4"" class=""pcSpacer""><script>tmpNListChecked=" & tmpNListChecked & ";</script><input type=""hidden"" name=""newslistcount"" value=""" & intCount & """></td></tr>"
			strNewMailUpArea=strNewMailUpArea & "<tr><td colspan=""4"">" & dictLanguage.Item(Session("language")&"_MailUp_RegisterLabel") & "</td></tr>"
	
			For j=0 to intCount
				strNewMailUpArea=strNewMailUpArea & "<tr> <td align=""right"" valign=""top""><input type=""checkbox"" onclick=""javascript: tmpNListChecked=1;"" value=""" & tmpArr(0,j) & """ name=""newslist" & j & """ "
				if tmpArr(5,j)="1" or tmpArr(5,j)="4" or (Session("idCustomer")=0 AND Session("pcSFpcNewsList" & j)&""=tmpArr(0,j)&"") then
					strNewMailUpArea=strNewMailUpArea & "checked"
				end if
				strNewMailUpArea=strNewMailUpArea & " class=""clearBorder"" /></td>"
				strNewMailUpArea=strNewMailUpArea & "<td valign=""top"" colspan=""3""><b>" & tmpArr(3,j) & "</b>"
				if tmpArr(5,j)="4" then
					strNewMailUpArea=strNewMailUpArea & "&nbsp;(<span class=""pcTextMessage"">" & dictLanguage.Item(Session("language")&"_MailUp_RegisterLabel1") & "</span>)"
				end if
				if tmpArr(4,j)<>"" then
					strNewMailUpArea=strNewMailUpArea & "<br>" & tmpArr(4,j)
				end if
				if tmpArr(5,j)="4" then
					strNewMailUpArea=strNewMailUpArea & "<div class=""pcErrorMessage"">" & dictLanguage.Item(Session("language")&"_MailUp_RegisterLabel2a") & "<a href=""javascript:newWindow('mu_subscribe.asp?listid=" & tmpArr(1,j) & "','window1');"">" & dictLanguage.Item(Session("language")&"_MailUp_RegisterLabel2b") & "</a>" & dictLanguage.Item(Session("language")&"_MailUp_RegisterLabel2c") & "</div>"
				end if
				strNewMailUpArea=strNewMailUpArea & "</td></tr>"
			Next
			strNewMailUpArea=strNewMailUpArea & "</table>"
		end if
		set rs=nothing	
		'End If MailUp Lists
	END IF
	'MAILUP-E
	
	'Start Special Customer Fields
	tmpCustCFList=""
	pcSFCustFieldsExist=""
	
	query="SELECT pcCField_ID, pcCField_Name, pcCField_FieldType, pcCField_Value, pcCField_Length, pcCField_Maximum, pcCField_Required, pcCField_PricingCategories, pcCField_ShowOnReg, pcCField_ShowOnCheckout,'',pcCField_Description,0 FROM pcCustomerFields ORDER BY pcCField_Name ASC;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closeDb()
		response.clear
		Call SetContentType()
		response.write "ERROR"
		response.End
	end if
	if not rs.eof then
		pcSFCustFieldsExist="YES"
		tmpCustCFList=rs.GetRows()
	end if
	set rs=nothing

	if pcSFCustFieldsExist="YES" AND Session("idCustomer")<>0 then
	pcArr=tmpCustCFList
	For k=0 to ubound(pcArr,2)
		pcArr(10,k)=""
		query="SELECT pcCFV_Value FROM pcCustomerFieldsValues WHERE idcustomer=" & Session("idCustomer") & " AND pcCField_ID=" & pcArr(0,k) & ";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closeDb()
			response.clear
			Call SetContentType()
			response.write "ERROR"
			response.End
		end if
		if not rs.eof then
			pcArr(10,k)=rs("pcCFV_Value")
		end if
		set rs=nothing
	Next
	tmpCustCFList=pcArr
	end if

	if pcSFCustFieldsExist="YES" then
		pcArr=tmpCustCFList
		For k=0 to ubound(pcArr,2)						
			pcv_ShowField=0
			if pcArr(9,k)="1" then
				pcv_ShowField=1
			end if
			if (pcv_ShowField=1) AND (pcArr(7,k)="1") then
			if session("idCustomer")>"0" then
				query="SELECT pcCustFieldsPricingCats.idcustomerCategory FROM pcCustFieldsPricingCats INNER JOIN Customers ON (pcCustFieldsPricingCats.pcCField_ID=" & pcArr(0,k) & " AND pcCustFieldsPricingCats.idCustomerCategory=customers.idCustomerCategory) WHERE customers.idcustomer=" & session("idCustomer")
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=conntemp.execute(query)	
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closeDb()
					response.clear
					Call SetContentType()
					response.write "ERROR"
					response.End
				end if											
				if NOT rs.eof then
					pcv_ShowField=1
				else
					pcv_ShowField=0
				end if
				set rs=nothing
			else
				pcv_ShowField=0
			end if
			end if	
			pcArr(12,k)=pcv_ShowField
		Next
		tmpCustCFList=pcArr
	end if

	if pcSFCustFieldsExist="YES" then
	pcArr=tmpCustCFList
						
	For k=0 to ubound(pcArr,2)
		pcv_ShowField=pcArr(12,k)
		if pcv_ShowField=1 then
			if pcArr(5,k)>"0" then
				pcArr(10,k)=URLDecode(getUserInput(request("custfield" & pcArr(0,k)),pcArr(5,k)))
			else
				pcArr(10,k)=URLDecode(getUserInput(request("custfield" & pcArr(0,k)),0))
			end if
			if pcArr(10,k)<>"" then
				query="SELECT pcCField_ID FROM pcCustomerFieldsValues WHERE idcustomer=" & session("idCustomer") & " AND pcCField_ID=" & pcArr(0,k) & ";"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closeDb()
					response.clear
					Call SetContentType()
					response.write "ERROR"
					response.End
				end if
				if NOT rs.eof then
					query="UPDATE pcCustomerFieldsValues SET pcCFV_Value='" & pcArr(10,k) & "' WHERE idcustomer=" & session("idCustomer") & " AND pcCField_ID=" & pcArr(0,k) & ";"
				else
					query="INSERT INTO pcCustomerFieldsValues (idcustomer,pcCField_ID,pcCFV_Value) VALUES (" & session("idCustomer") & "," & pcArr(0,k) & ",'" & pcArr(10,k) & "');"
				end if
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closeDb()
					response.clear
					Call SetContentType()
					response.write "ERROR"
					response.End
				end if
				set rs=nothing
			else
				query="DELETE FROM pcCustomerFieldsValues WHERE idcustomer=" & session("idCustomer") & " AND pcCField_ID=" & pcArr(0,k) & ";"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closeDb()
					response.clear
					Call SetContentType()
					response.write "ERROR"
					response.End
				end if
				set rs=nothing
			end if
		end if
	Next
	
	end if
	
	tmpStr=""
	
	tmpIDRefer=URLDecode(getUserInput(request("IDRefer"),0))
	if tmpIDRefer<>"" then
		if not IsNumeric(tmpIDRefer) then
			tmpIDRefer=0
		end if
		tmpStr="IDRefer=" & tmpIDRefer
	end if

	'MailUp-S
	if ((session("SF_MU_Setup")="1") AND (tmp_DontShowMailUp=0) AND (NewsCheckout="1")) OR ((session("SF_MU_Setup")<>"1") AND (AllowNews="1") AND (NewsCheckout="1")) then
	'MailUp-E
		if tmpStr<>"" then
			tmpStr=tmpStr & ","
		end if
		tmpRecvNews=URLDecode(getUserInput(request("CRecvNews"),0))
		if tmpRecvNews="" then
			tmpRecvNews=0
		end if
		if not IsNumeric(tmpRecvNews) then
			tmpRecvNews=0
		end if
		tmpStr=tmpStr & "RecvNews=" & tmpRecvNews
		Session("pcSFCRecvNews")=tmpRecvNews
	end if
	
	if Session("pcCustomerTermsAgreed")="1" then
		if tmpStr<>"" then
			tmpStr=tmpStr & ","
		end if
		tmpStr=tmpStr & "pcCust_AgreeTerms=1"
	end if
	
	if tmpStr<>"" then
		query="UPDATE Customers SET " & tmpStr & " WHERE idCustomer=" & session("idCustomer") & ";"
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closeDb()
			response.clear
			Call SetContentType()
			response.write "ERROR"
			response.End
		end if
		set rs=nothing
	end if
	%>
	<!--#include file="DBsv.asp"-->
    <%
	call opendb
	query="UPDATE pcCustomerSessions SET pcCustSession_BillingStateCode='"&pcStrBillingStateCode&"', pcCustSession_BillingCity='"&pcStrBillingCity&"', pcCustSession_BillingProvince='"&pcStrBillingProvince&"', pcCustSession_BillingPostalCode='"&pcStrBillingPostalCode&"', pcCustSession_BillingCountryCode='"&pcStrBillingCountryCode&"', pcCustSession_ShippingResidential='"&pcIntShippingResidential&"' WHERE (((idDbSession)="&session("pcSFIdDbSession")&") AND ((randomKey)="&session("pcSFRandomKey")&"));"
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closeDb()
		response.clear
		Call SetContentType()
		response.write "ERROR"
		response.End
	end if
	set rs=nothing
	%>
	<%
	call closedb()
	if pcErrMsg<>"" then
		pcErrMsg=dictLanguage.Item(Session("language")&"_opc_69")&"<br><ul>" & pcErrMsg & "</ul>"
		response.write pcErrMsg
	end if
	'MailUp-S
	IF session("SF_MU_Setup")="1" THEN
		response.write OKmsg & "|s|" & strNewMailUpArea
	ELSE
		response.write OKmsg
	END IF
	'MailUp-E
	Session("CurrentPanel") = "Ship"
end if
'//////////////////////////////////////////////////////////////////////////
'// END: UPDATE BILLING
'//////////////////////////////////////////////////////////////////////////

conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing

call closeDb()
response.End()
%>


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
<!--#include file="pcStartSession.asp" -->
<!--#include file="opc_contentType.asp" -->
<% On Error Resume Next

Call SetContentType()

pcShipOpt=URLDecode(getUserInput(request("ShipArrOpts"),5))

if session("idCustomer")=0 OR session("idCustomer")="" OR pcShipOpt="" then
	response.clear
	Call SetContentType()
	response.write "SECURITY"
	response.End
end if

call openDb()

session("pcShipOpt")=pcShipOpt

pcErrMsg=""

'//////////////////////////////////////////////////////////////////////////
'// START: VALIDATE SHIPPING
'//////////////////////////////////////////////////////////////////////////

pcStrShippingNickName=URLDecode(getUserInput(request("shipnickname"),50))
pcStrShippingFirstName=URLDecode(getUserInput(request("shipfname"),50))
pcStrShippingLastName=URLDecode(getUserInput(request("shiplname"),50))
pcStrShippingCompany=URLDecode(getUserInput(request("shipcompany"),150))
pcStrShippingPhone=URLDecode(getUserInput(request("shipphone"),20))
pcStrShippingEmail=URLDecode(getUserInput(request("shipemail"),150))
pcStrShippingAddress=URLDecode(getUserInput(request("shipaddr"),255))
if pcStrShippingNickName="" then
	pcStrShippingNickName=pcStrShippingAddress
end if
pcStrShippingPostalCode=URLDecode(getUserInput(request("shipzip"),10))
pcStrShippingStateCode=URLDecode(getUserInput(request("shipstate"),10))
pcStrShippingProvince=URLDecode(getUserInput(request("shipprovince"),150))
pcStrShippingCity=URLDecode(getUserInput(request("shipcity"),150))
pcStrShippingCountryCode=URLDecode(getUserInput(request("shipcountry"),10))
pcStrShippingAddress2=URLDecode(getUserInput(request("shipaddr2"),255))
pcStrShippingFax=URLDecode(getUserInput(request("shipfax"),20))
pcIntShippingResidential=URLDecode(getUserInput(request("pcAddressType"),0))

if pcIntShippingResidential<>"" then
	if not IsNumeric(pcIntShippingResidential) then
		pcIntShippingResidential="1"
	end if
	session(pcShipOpt)=pcIntShippingResidential
end if

if pcStrShippingEmail<>"" then
	pcStrShippingEmail=replace(pcStrShippingEmail," ","")
	if instr(pcStrShippingEmail,"@")=0 or instr(pcStrShippingEmail,".")=0 then
		pcErrMsg=pcErrMsg & "<li>Email Address is not valid</li>"
	end if
end if

if (pcShipOpt<>"-1") AND (pcShipOpt<>"-2") then
	if pcStrShippingFirstName="" then
		pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_58")&"</li>"
	end if
	if pcStrShippingLastName="" then
		pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_59")&"</li>"
	end if
	if pcStrShippingAddress="" then
		pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_72")&"</li>"
	end if
	if pcStrShippingCity="" then
		pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_73")&"</li>"
	end if
	if pcStrShippingCountryCode="" then
		pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_74")&"</li>"
	end if
	if (pcStrShippingCountryCode="US") OR (pcStrShippingCountryCode="CA") then
		if pcStrShippingStateCode="" then
			pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_75")&"</li>"
		end if
		if pcStrShippingPostalCode="" then
			pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_76")&"</li>"
		end if
	end if
end if

pcSFTF1=URLDecode(getUserInput(request("TF1"),0))
pcSFDF1=URLDecode(getUserInput(request("DF1"),0))

if (TFShow="1" AND TFReq="1") AND pcSFTF1="" then
	pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_77")&"</li>"
end if

if (DFShow="1" AND DFReq="1") AND pcSFDF1="" then
	pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_78")&"</li>"
else
	if DFShow="1" then
		if pcSFDF1<>"" then
			if not IsDate(pcSFDF1) then
				pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_79")&"</li>"
			else
				if scDateFrmt="DD/MM/YY" then
					expDateArray=split(pcSFDF1,"/")
					pcSFDF1=(expDateArray(1)&"/"&expDateArray(0)&"/"&expDateArray(2)) '// Normalize
				end if
				if SQL_Format="1" then
					expDateArray=split(pcSFDF1,"/")
					tmpDFDate=(expDateArray(1)&"/"&expDateArray(0)&"/"&expDateArray(2))
				else
					tmpDFDate=pcSFDF1
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
				IF icounter > 0 THEN
					pcErrMsg=pcErrMsg & "<li>" & blackoutmessage & dictLanguage.Item(Session("language")&"_catering_5") & "</li>"
				ELSE
					if CDate(pcSFDF1)<Date() then
						pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_80")&"</li>"
					else
						if pcSFTF1<>"" then
							EnteredHour=split(pcSFTF1,":")
							if Instr(EnteredHour(1),"PM")>0 then
								EnteredHour(0)=Cint(EnteredHour(0))+12
							end if
						end if
						if DTCheck=1 then
							if pcSFTF1="" then
								if (CDate(pcSFDF1)<=Date()) then
									pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_81")&"</li>"
								end if
							else
								if ( (CDate(pcSFDF1)<=Date()) OR ( (CDate(pcSFDF1)=Date()+1) AND (CInt(EnteredHour(0)) < Cint(Hour(Now())) ) ) ) then
									pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_81")&"</li>"
								end if
							end if
						end if
					end if
				END IF
			end if
		end if
	end if
end if
'//////////////////////////////////////////////////////////////////////////
'// END: VALIDATE SHIPPING
'//////////////////////////////////////////////////////////////////////////

OKmsg=""

'//////////////////////////////////////////////////////////////////////////
'// START: UPDATE SHIPPING
'//////////////////////////////////////////////////////////////////////////
if pcErrMsg="" then
	
	'// Use Same Shipping As Billing
	if pcShipOpt="-1" then
		
		query="SELECT [name], lastName, email, phone, fax, customerCompany, address,address2, city, state, stateCode, zip, countryCode FROM customers WHERE idcustomer="& session("idCustomer") &" AND pcCust_Guest=" & session("CustomerGuest") & ";"
		set rs=server.CreateObject("ADODB.RecordSet")
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
		if not rs.eof then
			pcStrShippingNickName=""
			pcStrShippingFirstName=rs("name")
			pcStrShippingLastName=rs("lastName")
			pcStrShippingCompany=rs("customerCompany")
			pcStrShippingPhone=rs("phone")
			pcStrShippingEmail=rs("email")
			pcStrShippingAddress=rs("address")
			pcStrShippingPostalCode=rs("zip")
			pcStrShippingStateCode=rs("stateCode")
			pcStrShippingProvince=rs("state")
			pcStrShippingCity=rs("city")
			pcStrShippingCountryCode=rs("countryCode")
			pcStrShippingAddress2=rs("address2")
			pcStrShippingFax=rs("fax")
			
			pcStrShippingNickName=getUserInput(pcStrShippingNickName,0)
			pcStrShippingFirstName=getUserInput(pcStrShippingFirstName,0)
			pcStrShippingLastName=getUserInput(pcStrShippingLastName,0)
			pcStrShippingCompany=getUserInput(pcStrShippingCompany,0)
			pcStrShippingPhone=getUserInput(pcStrShippingPhone,0)
			pcStrShippingEmail=getUserInput(pcStrShippingEmail,0)
			pcStrShippingAddress=getUserInput(pcStrShippingAddress,0)
			pcStrShippingPostalCode=getUserInput(pcStrShippingPostalCode,0)
			pcStrShippingStateCode=getUserInput(pcStrShippingStateCode,0)
			pcStrShippingProvince=getUserInput(pcStrShippingProvince,0)
			pcStrShippingCity=getUserInput(pcStrShippingCity,0)
			pcStrShippingCountryCode=getUserInput(pcStrShippingCountryCode,0)
			pcStrShippingAddress2=getUserInput(pcStrShippingAddress2,0)
			pcStrShippingFax=getUserInput(pcStrShippingFax,0)
			pcIntShippingResidential=getUserInput(pcIntShippingResidential,0)
			
			OKmsg="OK"
		end if
		set rs=nothing
		
	else '// if pcShipOpt="-1" then


		if pcShipOpt="-2" then
			
			query="select pcEv_IDCustomer, pcEv_Delivery, pcEv_MyAddr, pcEv_HideAddress from PcEvents where pcEv_IDEvent=" & session("Cust_IDEvent")
			set rstemp=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rstemp=nothing
				call closeDb()
				response.clear
				Call SetContentType()
				response.write "ERROR"
				response.End
			end if
			if not rstemp.eof then
				gIDCustomer=rstemp("pcEv_IDCustomer")
				gDelivery=rstemp("pcEv_Delivery")
				if gDelivery<>"" then
				else
					gDelivery=0
				end if
				gMyAddr=rstemp("pcEv_MyAddr")
				if gMyAddr<>"" then
				else
					gMyAddr=0
				end if
				if gDelivery="1" then
					GRTest=1
				end if
				gHideAddress=rstemp("pcEv_HideAddress")
				if IsNull(gHideAddress) OR gHideAddress="" then
					gHideAddress=0
				end if
			end if
			set rstemp=nothing
			
			query="select name,lastname from Customers where IDCustomer=" & gIDCustomer
			set rstemp=connTemp.execute(query)	
			if err.number<>0 then
				call LogErrorToDatabase()
				set rstemp=nothing
				call closeDb()
				response.clear
				Call SetContentType()
				response.write "ERROR"
				response.End
			end if							
			if not rstemp.eof then
				ggg_FName=rstemp("name")
				ggg_LName=rstemp("lastname")
			end if							
			set rstemp=nothing
								
			ggg_idRecipient=0
									
			query="SELECT Address,Address2,City,State,Statecode,Zip,CountryCode,customerCompany, shippingAddress, shippingCity, shippingState, shippingStateCode, shippingZip, shippingCountryCode, shippingCompany, shippingAddress2,phone,email,fax,shippingPhone,shippingEmail,shippingFax,pcCust_Residential FROM customers WHERE idCustomer=" & gIDCustomer
			set rstemp=conntemp.execute(query)	
			if err.number<>0 then
				call LogErrorToDatabase()
				set rstemp=nothing
				call closeDb()
				response.clear
				Call SetContentType()
				response.write "ERROR"
				response.End
			end if									
			if not rstemp.eof then
				if rstemp("shippingAddress")<>"" then
					ggg_idRecipient=0
					ggg_NickName="Gift Registrant''s Address"
					ggg_shipAdd=rstemp("shippingAddress")
					ggg_shipZip=rstemp("shippingZip")
					ggg_shipState=rstemp("shippingState")
					ggg_shipStateCode=rstemp("shippingStateCode") 
					ggg_shipCity=rstemp("shippingCity")
					ggg_shipCountryCode=rstemp("shippingCountryCode")
					ggg_shipCom=rstemp("shippingCompany")
					ggg_shipAdd2=rstemp("shippingAddress2")
					ggg_Phone=rstemp("shippingPhone")
					ggg_Email=rstemp("shippingEmail")
					ggg_Fax=rstemp("shippingFax")
					ggg_Residential=rstemp("pcCust_Residential")
					OKmsg="OK"
				else
					ggg_idRecipient=0
					ggg_NickName="Gift Registrant''s Address"
					ggg_shipAdd=rstemp("Address")
					ggg_shipAdd2=rstemp("Address2")
					ggg_shipZip=rstemp("Zip")
					ggg_shipState=rstemp("State")
					ggg_shipStateCode=rstemp("StateCode") 
					ggg_shipCity=rstemp("City")
					ggg_shipCountryCode=rstemp("CountryCode")
					ggg_shipCom=rstemp("customerCompany")
					ggg_Phone=rstemp("phone")
					ggg_Email=rstemp("email")
					ggg_Fax=rstemp("fax")
					ggg_Residential=rstemp("pcCust_Residential")
					OKmsg="OK"
				end if	
			end if			
			set rstemp=nothing
			
			IF gMyAddr<>0 THEN
			
				query="select idRecipient,recipient_NickName,recipient_FirstName,recipient_LastName,recipient_Company,recipient_Address,recipient_Address2,recipient_City,recipient_StateCode,recipient_State,recipient_Zip,recipient_CountryCode,recipient_Phone,recipient_Fax,recipient_Email,recipient_Residential from recipients where idRecipient=" & gMyAddr
				set rstemp=connTemp.execute(query)	
				if err.number<>0 then
					call LogErrorToDatabase()
					set rstemp=nothing
					call closeDb()
					response.clear
					Call SetContentType()
					response.write "ERROR"
					response.End
				end if			
				if not rstemp.eof then
					ggg_idRecipient=rstemp("idRecipient")
					ggg_NickName=rstemp("recipient_NickName")
					ggg_FName=rstemp("recipient_FirstName")
					ggg_LName=rstemp("recipient_LastName")
					ggg_shipCom=rstemp("recipient_Company")
					ggg_shipAdd=rstemp("recipient_Address")
					ggg_shipAdd2=rstemp("recipient_Address2")
					ggg_shipCity=rstemp("recipient_City")
					ggg_shipStateCode=rstemp("recipient_StateCode")
					ggg_shipState=rstemp("recipient_State")
					ggg_shipZip=rstemp("recipient_Zip")
					ggg_shipCountryCode=rstemp("recipient_CountryCode")
					ggg_Phone=rstemp("recipient_Phone")
					ggg_Email=rstemp("recipient_Email")
					ggg_Fax=rstemp("recipient_Fax")
					ggg_Residential=rstemp("recipient_Residential")
					OKmsg="OK"
				end if				
				set rstemp=nothing
			
			END IF
			
			pcStrShippingNickName=getUserInput(ggg_NickName,0)
			pcStrShippingFirstName=getUserInput(ggg_FName,0)
			pcStrShippingLastName=getUserInput(ggg_LName,0)
			pcStrShippingCompany=getUserInput(ggg_shipCom,0)
			pcStrShippingPhone=getUserInput(ggg_Phone,0)
			pcStrShippingEmail=getUserInput(ggg_Email,0)
			pcStrShippingAddress=getUserInput(ggg_shipAdd,0)
			pcStrShippingPostalCode=getUserInput(ggg_shipZip,0)
			pcStrShippingStateCode=getUserInput(ggg_shipStateCode,0)
			pcStrShippingProvince=getUserInput(ggg_shipState,0)
			pcStrShippingCity=getUserInput(ggg_shipCity,0)
			pcStrShippingCountryCode=getUserInput(ggg_shipCountryCode,0)
			pcStrShippingAddress2=getUserInput(ggg_shipAdd2,0)
			pcStrShippingFax=getUserInput(ggg_Fax,0)
			pcIntShippingResidential=getUserInput(ggg_Residential,0)
		
		else '// if pcShipOpt="-2" then
		
			if pcShipOpt="0" then
				query="UPDATE Customers SET shippingCompany='" & pcStrShippingCompany & "',shippingAddress='" & pcStrShippingAddress & "',shippingAddress2='" & pcStrShippingAddress2 & "',shippingCity='" & pcStrShippingCity & "',shippingState='" & pcStrShippingProvince & "',shippingStateCode='" & pcStrShippingStateCode & "',shippingZip='" & pcStrShippingPostalCode & "',shippingCountryCode='" & pcStrShippingCountryCode & "',shippingPhone='" & pcStrShippingPhone & "',shippingEmail='" & pcStrShippingEmail & "',shippingFax='" & pcStrShippingFax & "'"
				if pcIntShippingResidential<>"" then
					query=query & ",pcCust_Residential=" & pcIntShippingResidential
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
				if pcShipOpt<>"" AND IsNumeric(pcShipOpt) then
					query="UPDATE Recipients SET recipient_FullName='" & pcStrShippingFirstName & " " & pcStrShippingLastName & "',recipient_NickName='" & pcStrShippingNickName & "',recipient_FirstName='" & pcStrShippingFirstName & "',recipient_LastName='" & pcStrShippingLastName & "',recipient_Email='" & pcStrShippingEmail & "',recipient_Phone='" & pcStrShippingPhone & "',recipient_Fax='" & pcStrShippingFax & "',recipient_Company='" & pcStrShippingCompany & "', recipient_Address='" & pcStrShippingAddress & "',recipient_Address2='" & pcStrShippingAddress2 & "',recipient_City='" & pcStrShippingCity & "',recipient_State='" & pcStrShippingProvince & "',recipient_StateCode='" & pcStrShippingStateCode & "',recipient_Zip='" & pcStrShippingPostalCode & "',recipient_CountryCode='" & pcStrShippingCountryCode & "'"
					if pcIntShippingResidential<>"" then
						query=query & ",Recipient_Residential=" & pcIntShippingResidential
					end if
					query=query & " WHERE idcustomer=" & session("idCustomer") & " AND idRecipient=" & pcShipOpt & ";"
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
					if pcShipOpt="ADD" then
						
						pcv_strNickNameUsed="0"
						query="SELECT TOP 1 idRecipient FROM Recipients WHERE idcustomer=" & session("idCustomer") & " AND recipient_NickName='" & pcStrShippingNickName & "' ORDER BY idRecipient DESC;"
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
							pcv_strNickNameUsed="1"
							pcShipOpt=rs("idRecipient")
						end if
						set rs=nothing						
						
						If pcv_strNickNameUsed="1" Then
						
							'// Update Recipient
							query="UPDATE Recipients SET recipient_FullName='" & pcStrShippingFirstName & " " & pcStrShippingLastName & "',recipient_NickName='" & pcStrShippingNickName & "',recipient_FirstName='" & pcStrShippingFirstName & "',recipient_LastName='" & pcStrShippingLastName & "',recipient_Email='" & pcStrShippingEmail & "',recipient_Phone='" & pcStrShippingPhone & "',recipient_Fax='" & pcStrShippingFax & "',recipient_Company='" & pcStrShippingCompany & "', recipient_Address='" & pcStrShippingAddress & "',recipient_Address2='" & pcStrShippingAddress2 & "',recipient_City='" & pcStrShippingCity & "',recipient_State='" & pcStrShippingProvince & "',recipient_StateCode='" & pcStrShippingStateCode & "',recipient_Zip='" & pcStrShippingPostalCode & "',recipient_CountryCode='" & pcStrShippingCountryCode & "'"
							if pcIntShippingResidential<>"" then
								query=query & ",Recipient_Residential=" & pcIntShippingResidential
							end if
							query=query & " WHERE idcustomer=" & session("idCustomer") & " AND idRecipient=" & pcShipOpt & ";"
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

							
						Else
						
							'// Add Recipient
							tmp1=""
							tmp2=""
							if pcIntShippingResidential<>"" then
								tmp1=",Recipient_Residential"
								tmp2="," & pcIntShippingResidential
							end if
							query="INSERT INTO Recipients (idcustomer,recipient_FullName,recipient_NickName,recipient_FirstName,recipient_LastName,recipient_Email,recipient_Phone,recipient_Fax,recipient_Company,recipient_Address,recipient_Address2,recipient_City,recipient_State,recipient_StateCode,recipient_Zip,recipient_CountryCode" & tmp1 & ") VALUES (" & session("idCustomer") & ",'" & pcStrShippingFirstName & " " & pcStrShippingLastName & "','" & pcStrShippingNickName & "','" & pcStrShippingFirstName & "','" & pcStrShippingLastName & "','" & pcStrShippingEmail & "','" & pcStrShippingPhone & "','" & pcStrShippingFax & "','" & pcStrShippingCompany & "','" & pcStrShippingAddress & "','" & pcStrShippingAddress2 & "','" & pcStrShippingCity & "','" & pcStrShippingProvince & "','" & pcStrShippingStateCode & "','" & pcStrShippingPostalCode & "','" & pcStrShippingCountryCode & "'" & tmp2 & ");"
							set rs=connTemp.execute(query)
							set rs=nothing
							OKmsg="OK"
							query="SELECT TOP 1 idRecipient FROM Recipients WHERE idcustomer=" & session("idCustomer") & " ORDER BY idRecipient DESC;"
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
								pcShipOpt=rs("idRecipient")
								session("pcShipOpt")=pcShipOpt
							end if
							set rs=nothing
						
						End If
						
					end if
				end if
			end if
			
		end if '// if pcShipOpt="-2" then
		
	end if '// if pcShipOpt="-1" then

	'Check the PostalCode Length for United States
	If pcStrShippingCountryCode="US" Then
		if len(pcStrShippingPostalCode)<5 then
			response.clear
			Call SetContentType()
			response.Write("ZIPLENGTH")
			response.End()
		end if
	End If
	
	if OKmsg="OK" then
		%>
		<!--#include file="DBsv.asp"-->
		<%
		call opendb()
		if Clng(pcShipOpt)>=0 then
			pcShipOptA="1"
		else
			if (pcShipOpt="-1") OR ((pcShipOpt="-2") AND (gHideAddress="1")) then
				pcShipOptA="0"
			else
				pcShipOptA="1"
			end if
		end if
		query="UPDATE pcCustomerSessions SET pcCustSession_ShowShipAddr=" & pcShipOptA & ", idCustomer="&session("idCustomer")&", pcCustSession_ShippingNickName='"&pcStrShippingNickName&"', pcCustSession_ShippingFirstName='"&pcStrShippingFirstName&"', pcCustSession_ShippingLastName='"&pcStrShippingLastName&"', pcCustSession_ShippingCompany='"&pcStrShippingCompany&"', pcCustSession_ShippingPhone='"&pcStrShippingPhone&"', pcCustSession_ShippingAddress='"&pcStrShippingAddress&"', pcCustSession_ShippingPostalCode='"&pcStrShippingPostalCode&"', pcCustSession_ShippingStateCode='"&pcStrShippingStateCode&"', pcCustSession_ShippingProvince='"&pcStrShippingProvince&"', pcCustSession_ShippingCity='"&pcStrShippingCity&"', pcCustSession_ShippingCountryCode='"&pcStrShippingCountryCode&"', pcCustSession_ShippingAddress2='"&pcStrShippingAddress2&"', pcCustSession_ShippingResidential='"&pcIntShippingResidential&"', pcCustSession_ShippingFax='"&pcStrShippingFax&"', pcCustSession_ShippingEmail='"&pcStrShippingEmail&"', pcCustSession_TF1='"&pcSFTF1&"', pcCustSession_DF1='"&pcSFDF1&"' WHERE (((idDbSession)="&session("pcSFIdDbSession")&") AND ((randomKey)="&session("pcSFRandomKey")&"));"
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
	else '// if OKmsg="OK" then
		pcErrMsg=dictLanguage.Item(Session("language")&"_opc_83")
	end if '// if OKmsg="OK" then
	
end if
'//////////////////////////////////////////////////////////////////////////
'// END: UPDATE SHIPPING
'//////////////////////////////////////////////////////////////////////////

If DeliveryZip = "1" Then
	call opendb()
	query="SELECT * from zipcodevalidation WHERE zipcode='" &pcStrShippingPostalCode& "'"
	set rsZipCodeObj=server.CreateObject("ADODB.RecordSet")
	set rsZipCodeObj=conntemp.execute(query)
	if rsZipCodeObj.eof then
		set rsZipCodeObj=nothing
		call closeDb()
		pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_Custmoda_23")&"</li>"
	end if
	call closeDb()	
End If

if pcErrMsg="" then
	session("OPCstep")=3
	call opendb()
end if

if pcErrMsg<>"" then
	pcErrMsg=dictLanguage.Item(Session("language")&"_opc_84")&"<br><ul>" & pcErrMsg & "</ul>"
	response.write pcErrMsg
else
	response.write OKmsg & "|*|" & pcShipOpt
	Session("CurrentPanel") = "ShipMethod"
end if
call closeDb()
response.End()
%>



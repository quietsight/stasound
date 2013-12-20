
<% if ptaxCanada="1" then

	'// Clear any existing sessions
	iTaxRateCnt=session("SFTaxZoneRateCnt")
	if isNumeric(iTaxRateCnt) then
		for iTCnt=1 to iTaxRateCnt
			session("SFTaxZoneRateName"&iTCnt)=""
			session("SFTaxZoneRateRate"&iTCnt)=""
			session("SFTaxZoneRateApplyToSH"&iTCnt)=""
			session("SFTaxZoneRateTaxable"&iTCnt)=""
			session("SFTaxZoneRateID"&iTCnt)=""
		next
		session("SFTaxZoneRateCnt")=0
		session("taxCnt")=0
	end if

	if session("ExpressCheckoutPayment")="YES" then
		
		'// Express Checkout billing address is anonymous, use shipping address 
		query="SELECT  pcCustomerSessions.pcCustSession_ShippingCity, pcCustomerSessions.pcCustSession_ShippingStateCode, pcCustomerSessions.pcCustSession_ShippingProvince, pcCustomerSessions.pcCustSession_ShippingPostalCode, pcCustomerSessions.pcCustSession_ShippingCountryCode FROM pcCustomerSessions WHERE (((pcCustomerSessions.idDbSession)="&session("pcSFIdDbSession")&") AND ((pcCustomerSessions.randomKey)="&session("pcSFRandomKey")&") AND ((pcCustomerSessions.idCustomer)="&session("idCustomer")&")) ORDER BY idDbSession DESC;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)		
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if		
		if NOT rs.eof then
			pcStrBillingCity=rs("pcCustSession_ShippingCity")
			pcStrBillingStateCode=rs("pcCustSession_ShippingStateCode")
			pcStrBillingProvince=rs("pcCustSession_ShippingProvince")
			pcStrBillingPostalCode=rs("pcCustSession_ShippingPostalCode")
			pcStrBillingCountryCode=rs("pcCustSession_ShippingCountryCode")
		end if	
		set rs=nothing	

	end if
	
	pcStrBillingState=pcStrBillingStateCode
	if len(pcStrBillingStateCode)<1 then
		pcStrBillingState=pcStrBillingProvince
	end if
	
	pcTaxZoneState=pcStrBillingState
	pcTaxZoneCity=pcStrBillingCity
	pcTaxZoneCountryCode=pcStrBillingCountryCode
	pcTaxZonePostalCode=pcStrBillingPostalCode

	if ptaxshippingaddress="1" AND (pcStrShippingCountryCode<>"" OR pcStrShippingPostalCode<>"") then
		pcStrShippingState=pcStrShippingStateCode
		if len(pcStrShippingStateCode)<1 then
			pcStrShippingState=pcStrShippingProvince
		end if
		If pcStrShippingState&""<>"" then
			pcTaxZoneState=pcStrShippingState
			pcTaxZoneCity=pcStrShippingCity
			pcTaxZoneCountryCode=pcStrShippingCountryCode
			pcTaxZonePostalCode=pcStrShippingPostalCode
		End If
	end if

	'// Find Zone Canada
	query="SELECT pcTaxZone_ID FROM pcTaxZones WHERE pcTaxZone_CountryCode='"& pcTaxZoneCountryCode &"' AND pcTaxZone_Province='"& pcTaxZoneState &"';"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	if NOT rs.eof then
		pcIntTaxZone_ID=rs("pcTaxZone_ID")
	else
		pcIntTaxZone_ID=0
	end if
	set rs=nothing

	if pcIntTaxZone_ID<>0 then

		query="SELECT pcTaxZoneRates.pcTaxZoneRate_ID, pcTaxZoneRates.pcTaxZonerate_LocalZone, pcTaxZoneRates.pcTaxZoneRate_Name, pcTaxZoneRates.pcTaxZoneRate_Rate, pcTaxZoneRates.pcTaxZoneRate_ApplyToSH, pcTaxZoneRates.pcTaxZoneRate_Taxable, pcTaxGroups.pcTaxZone_ID FROM pcTaxGroups INNER JOIN (pcTaxZonesGroups INNER JOIN pcTaxZoneRates ON pcTaxZonesGroups.pcTaxZoneRate_ID = pcTaxZoneRates.pcTaxZoneRate_ID) ON pcTaxGroups.pcTaxZoneDesc_ID = pcTaxZonesGroups.pcTaxZoneDesc_ID WHERE (((pcTaxGroups.pcTaxZone_ID)="& pcIntTaxZone_ID &")) ORDER BY pcTaxZoneRates.pcTaxZoneRate_Order;"		
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		
		dim iTaxRateCnt
		iTaxRateCnt=0
		
		do until rs.eof
			intSkipZoneRate=0
			intTempTaxZoneRateID=rs("pcTaxZoneRate_ID")

			'// Check if customer is exempt
			query="SELECT * FROM pcTaxEptCust WHERE idCustomer="& Session("idCustomer") &" AND pcTaxZoneRate_ID="& intTempTaxZoneRateID &";"
			set rsCustCheckObj=server.CreateObject("ADODB.RecordSet")
			set rsCustCheckObj=conntemp.execute(query)
			if NOT rsCustCheckObj.eof then
				intSkipZoneRate=1
			end if
			set rsCustCheckObj=nothing
			
			if (rs("pcTaxZonerate_LocalZone"))<>0 AND (ucase(pcTaxZoneState)<>ucase(scShipFromState)) AND intSkipZoneRate=0 then
				'// make sure zone is local
				intSkipZoneRate=1
			end if

			if intSkipZoneRate=0 then
				iTaxRateCnt=iTaxRateCnt+1
				session("SFTaxZoneRateName"&iTaxRateCnt)=rs("pcTaxZoneRate_Name")
				session("SFTaxZoneRateRate"&iTaxRateCnt)=rs("pcTaxZoneRate_Rate")
				session("SFTaxZoneRateApplyToSH"&iTaxRateCnt)=rs("pcTaxZoneRate_ApplyToSH")
				session("SFTaxZoneRateTaxable"&iTaxRateCnt)=rs("pcTaxZoneRate_Taxable")
				session("SFTaxZoneRateID"&iTaxRateCnt)=intTempTaxZoneRateID
			end if
			rs.moveNext
		loop
		session("SFTaxZoneRateCnt")=iTaxRateCnt
		set rs=nothing
	end if

end if

pcIntDebug=0
if pcIntDebug=1 then
	response.write "Tax Zone ID in Database: "&pcIntTaxZone_ID&"<BR>"
	response.write "Number of Tax Rates Found in Zone: "&session("SFTaxZoneRateCnt")&"<BR>"
	if session("SFTaxZoneRateCnt")="" then
		session("SFTaxZoneRateCnt")=0
	end if
	if session("SFTaxZoneRateCnt")>0 then
		for u=1 to session("SFTaxZoneRateCnt")
			response.write "<HR>Rate Name"&u&": "&session("SFTaxZoneRateName"&u)&"<BR>"
			response.write "Rate "&u&": "&session("SFTaxZoneRateRate"&u)&"<BR>"
			response.write "Taxable"&u&": "&session("SFTaxZoneRateTaxable"&u)&"<BR>"
		next
	end if
	response.end
end if
%>
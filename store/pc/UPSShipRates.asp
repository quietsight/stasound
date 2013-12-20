<% if ups_active=true or ups_active="-1" then
	Dim iUPSFlag
	iUPSFlag=0
	iUPSActive=1
	'//UPS Rates
	ups_postdata=""
	ups_postdata="<?xml version=""1.0""?>"
	ups_postdata=ups_postdata&"<AccessRequest xml:lang=""en-US"">"
	ups_postdata=ups_postdata&"<AccessLicenseNumber>"&ups_license_key&"</AccessLicenseNumber>"
	ups_postdata=ups_postdata&"<UserId>"&ups_userid&"</UserId>"
	ups_postdata=ups_postdata&"<Password>"&ups_password&"</Password>"
	ups_postdata=ups_postdata&"</AccessRequest>"
	ups_postdata=ups_postdata&"<?xml version=""1.0""?>"
	ups_postdata=ups_postdata&"<RatingServiceSelectionRequest xml:lang=""en-US"">"
	ups_postdata=ups_postdata&"<Request>"
	ups_postdata=ups_postdata&"<TransactionReference>"
	ups_postdata=ups_postdata&"<CustomerContext>Rating and Service</CustomerContext>"
	ups_postdata=ups_postdata&"<XpciVersion>1.0001</XpciVersion>"
	ups_postdata=ups_postdata&"</TransactionReference>"
	ups_postdata=ups_postdata&"<RequestAction>rate</RequestAction>"
	ups_postdata=ups_postdata&"<RequestOption>shop</RequestOption>"
	ups_postdata=ups_postdata&"</Request>"
	ups_postdata=ups_postdata&"<PickupType>"
	ups_postdata=ups_postdata&"<Code>"&UPS_PICKUP_TYPE&"</Code>"
	ups_postdata=ups_postdata&"</PickupType>"
	if UPS_CLASSIFICATION_TYPE<>"" then
		ups_postdata=ups_postdata&"<CustomerClassification>"
		ups_postdata=ups_postdata&"<Code>"&UPS_CLASSIFICATION_TYPE&"</Code>"
		ups_postdata=ups_postdata&"</CustomerClassification>"
	end if
	ups_postdata=ups_postdata&"<Shipment>"
	ups_postdata=ups_postdata&"<Shipper>"
	if pcv_UseNegotiatedRates=1 then
		if pcv_UPSShipperNumber<>"" then
			ups_postdata=ups_postdata&"<ShipperNumber>"&pcv_UPSShipperNumber&"</ShipperNumber>"
		end if
	end if
	ups_postdata=ups_postdata&"<Address>"
	ups_postdata=ups_postdata&"<City>"&UPS_ShipFromCity&"</City>"
	ups_postdata=ups_postdata&"<StateProvinceCode>"&UPS_ShipFromState&"</StateProvinceCode>"
	ups_postdata=ups_postdata&"<PostalCode>"&UPS_ShipFromPostalCode&"</PostalCode>"
	ups_postdata=ups_postdata&"<CountryCode>"&UPS_ShipFromPostalCountry&"</CountryCode>"
	ups_postdata=ups_postdata&"</Address>"
	ups_postdata=ups_postdata&"</Shipper>"
	ups_postdata=ups_postdata&"<ShipTo>"
	ups_postdata=ups_postdata&"<Address>"
	ups_postdata=ups_postdata&"<City>"&Universal_destination_city&"</City>"
	ups_postdata=ups_postdata&"<StateProvinceCode>"&Universal_destination_provOrState&"</StateProvinceCode>"
	ups_destination_postal=replace(Universal_destination_postal, " ","")
	ups_destination_postal=replace(ups_destination_postal,"-","")
	ups_postdata=ups_postdata&"<PostalCode>"&ups_destination_postal&"</PostalCode>"
	ups_postdata=ups_postdata&"<CountryCode>"&Universal_destination_country&"</CountryCode>"
	If pResidentialShipping<>"0" then
		ups_postdata=ups_postdata&"<ResidentialAddress>1</ResidentialAddress>"
	else
		ups_postdata=ups_postdata&"<ResidentialAddress>0</ResidentialAddress>"
	end if
	ups_postdata=ups_postdata&"</Address>"
	ups_postdata=ups_postdata&"</ShipTo>"
	for q=1 to pcv_intPackageNum
		ups_postdata=ups_postdata&"<Package>"
		ups_postdata=ups_postdata&"<PackagingType>"
		ups_postdata=ups_postdata&"<Code>"&UPS_PACKAGE_TYPE&"</Code>"
		ups_postdata=ups_postdata&"<Description>Package</Description>"
		ups_postdata=ups_postdata&"</PackagingType>"
		ups_postdata=ups_postdata&"<Description>Rate Shopping</Description>"
		ups_postdata=ups_postdata&"<Dimensions>"
		pUPS_DIM_UNIT=ucase(UPS_DIM_UNIT)
		if q>1 then
			pcv_intOSheight=UPS_HEIGHT
			pcv_intOSwidth=UPS_WIDTH
			pcv_intOSlength=UPS_LENGTH
		end if
		if scShipFromWeightUnit="KGS" AND pUPS_DIM_UNIT="IN" then
			pUPS_DIM_UNIT="CM"
			pcv_intOSlength=pcv_intOSlength*2.54
			pcv_intOSwidth=pcv_intOSwidth*2.54
			pcv_intOSheight=pcv_intOSheight*2.54
		end if
		if scShipFromWeightUnit="LBS" AND pUPS_DIM_UNIT="CM" then
			pUPS_DIM_UNIT="IN"
			pcv_intOSlength=pcv_intOSlength/2.54
			pcv_intOSwidth=pcv_intOSwidth/2.54
			pcv_intOSheight=pcv_intOSheight/2.54
		end if
		ups_postdata=ups_postdata&"<UnitOfMeasurement><Code>"&pUPS_DIM_UNIT&"</Code></UnitOfMeasurement>"
		ups_postdata=ups_postdata&"<Length>"&pc_dimensions(session("UPSPackLength"&q))&"</Length>" 'Between 1 and 108.00
		ups_postdata=ups_postdata&"<Width>"&pc_dimensions(session("UPSPackWidth"&q))&"</Width>" 'Between 1 and 108.00
		ups_postdata=ups_postdata&"<Height>"&pc_dimensions(session("UPSPackHeight"&q))&"</Height>" 'Between 1 and 108.00
		ups_postdata=ups_postdata&"</Dimensions>"
		ups_postdata=ups_postdata&"<PackageWeight>"
		ups_postdata=ups_postdata&"<UnitOfMeasurement>"
		if scShipFromWeightUnit="KGS" then
			ups_postdata=ups_postdata&"<Code>KGS</Code>"
		else
			ups_postdata=ups_postdata&"<Code>LBS</Code>"
		end if
		ups_postdata=ups_postdata&"</UnitOfMeasurement>"
		ups_postdata=ups_postdata&"<Weight>"&pc_dimensions(session("UPSPackWeight"&q))&"</Weight>" '0.1 to 150.0
		ups_postdata=ups_postdata&"</PackageWeight>"
		ups_postdata=ups_postdata&"<OversizePackage>0</OversizePackage>"
		ups_postdata=ups_postdata&"<PackageServiceOptions>"
		ups_postdata=ups_postdata&"<InsuredValue>"
		ups_postdata=ups_postdata&"<CurrencyCode>USD</CurrencyCode>"

		pcv_TempPackPrice=session("UPSPackPrice"&q)			
		If pcv_TempPackPrice="" Then
			pcv_TempPackPrice="100.00"
		End If
		
		ups_postdata=ups_postdata&"<MonetaryValue>"&replace(money(pcv_TempPackPrice),",","")&"</MonetaryValue>"
		ups_postdata=ups_postdata&"</InsuredValue>"
		ups_postdata=ups_postdata&"</PackageServiceOptions>"
		ups_postdata=ups_postdata&"</Package>"
	next
	if pcv_UseNegotiatedRates=1 then
		ups_postdata=ups_postdata&"<RateInformation>"
			ups_postdata=ups_postdata&"<NegotiatedRatesIndicator/>"
		ups_postdata=ups_postdata&"</RateInformation>"
	end if
	ups_postdata=ups_postdata&"</Shipment>"
	ups_postdata=ups_postdata&"</RatingServiceSelectionRequest>"
	
	'get URL to post to
	ups_URL="https://www.ups.com/ups.app/xml/Rate"

	Set srvUPSXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
	srvUPSXmlHttp.open "POST", ups_URL, false
	srvUPSXmlHttp.send(ups_postdata)
	UPS_result = srvUPSXmlHttp.responseText

	Set UPSXMLdoc = server.CreateObject("Msxml2.DOMDocument"&scXML)
	UPSXMLDoc.async = false 
	if UPSXMLDOC.loadXML(UPS_result) then ' if loading from a string
		set objLst = UPSXMLDOC.getElementsByTagName("RatedShipment") 
		for i = 0 to (objLst.length - 1)
			varFlag=0
			for j=0 to ((objLst.item(i).childNodes.length)-1)
				If objLst.item(i).childNodes(j).nodeName="Service" then
					serviceVar=objLst.item(i).childNodes(j).text
					select case serviceVar
					case "01"
						if pcv_UPSCanadaOrigin=0 then
							availableShipStr=availableShipStr&"|?|UPS|01|"&pServiceCodeString01
						else
							availableShipStr=availableShipStr&"|?|UPS|01|"&pServiceCodeString01
						end if
						varFlag=1
						iUPSFlag=1
					case "02"
						if pcv_UPSCanadaOrigin=0 then
							availableShipStr=availableShipStr&"|?|UPS|02|"&pServiceCodeString02
						else
							availableShipStr=availableShipStr&"|?|UPS|02|"&pServiceCodeString02
						end if
						varFlag=1
						iUPSFlag=1
					case "03"
						availableShipStr=availableShipStr&"|?|UPS|03|"&pServiceCodeString03
						varFlag=1
						iUPSFlag=1
					case "07"
						availableShipStr=availableShipStr&"|?|UPS|07|"&pServiceCodeString07
						varFlag=1
						iUPSFlag=1
					case "08"
						availableShipStr=availableShipStr&"|?|UPS|08|"&pServiceCodeString08
						varFlag=1
						iUPSFlag=1
					case "11"
						availableShipStr=availableShipStr&"|?|UPS|11|"&pServiceCodeString11
						varFlag=1
						iUPSFlag=1
					case "12"
						availableShipStr=availableShipStr&"|?|UPS|12|"&pServiceCodeString12
						varFlag=1
						iUPSFlag=1
					case "13"
						if pcv_UPSCanadaOrigin=0 then
							availableShipStr=availableShipStr&"|?|UPS|13|"&pServiceCodeString13
						else
							availableShipStr=availableShipStr&"|?|UPS|13|"&pServiceCodeString13
						end if
						varFlag=1
						iUPSFlag=1
					case "14"
						if pcv_UPSCanadaOrigin=0 then
							availableShipStr=availableShipStr&"|?|UPS|14|"&pServiceCodeString14
						else
							availableShipStr=availableShipStr&"|?|UPS|14|"&pServiceCodeString14
						end if
						varFlag=1
						iUPSFlag=1
					case "54"
						if pcv_UPSCanadaOrigin=0 then
							availableShipStr=availableShipStr&"|?|UPS|54|"&pServiceCodeString54
						else
							availableShipStr=availableShipStr&"|?|UPS|54|"&pServiceCodeString54
						end if							
						varFlag=1
						iUPSFlag=1
					case "59"
						availableShipStr=availableShipStr&"|?|UPS|59|"&pServiceCodeString59
						varFlag=1
						iUPSFlag=1
					case "65"
						availableShipStr=availableShipStr&"|?|UPS|65|"&pServiceCodeString65
						varFlag=1
						iUPSFlag=1
					end select
				End if
				
				'// Get Monetary Value
				If objLst.item(i).childNodes(j).nodeName="TotalCharges" then
					for k=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
						if objLst.item(i).childNodes(j).childNodes(k).nodeName="MonetaryValue" then
							availableShipStr=availableShipStr&"|"&objLst.item(i).childNodes(j).childNodes(k).text
						end if
					next
				End if

				if pcv_UseNegotiatedRates=1 then
					If objLst.item(i).childNodes(j).nodeName="NegotiatedRates" then
						for k=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
							if objLst.item(i).childNodes(j).childNodes(k).childNodes(0).childNodes(1).nodeName="MonetaryValue" then
								availableShipStr=availableShipStr&"|"&objLst.item(i).childNodes(j).childNodes(k).childNodes(0).childNodes(1).text
							else
								availableShipStr=availableShipStr&"|NONE"
							end if
						next
					End if
				end if
								
				If objLst.item(i).childNodes(j).nodeName="GuaranteedDaysToDelivery" AND varFlag=1 then
					if objLst.item(i).childNodes(j).text="1" then
						availableShipStr=availableShipStr&"|Next Day"
					else
						if objLst.item(i).childNodes(j).text<>"" then
							availableShipStr=availableShipStr&"|"&objLst.item(i).childNodes(j).text&" Days"
						else
							availableShipStr=availableShipStr&"|NA"
						end if
					end if
				End If
				If objLst.item(i).childNodes(j).nodeName="ScheduledDeliveryTime" AND varFlag=1 then
					If objLst.item(i).childNodes(j).text<>"" then
						availableShipStr=availableShipStr&" by "&objLst.item(i).childNodes(j).text
					end if
				End If
			next
			
		next 
	end if
end if 'if ups is active
%>
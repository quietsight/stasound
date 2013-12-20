<%
pcPageName = "ShipRates.asp"

pcv_strMethodNameWS = "RateRequest"
pcv_strMethodReplyWS = "RateResponse"
CustomerTransactionIdentifier = "ProductCart_Rates"
pcv_strEnvironment = FEDEXWS_Environment


'// FEDEX CREDENTIALS
call opendb()
query = "SELECT ShipmentTypes.userID, ShipmentTypes.password, ShipmentTypes.AccessLicense, ShipmentTypes.FedExKey, ShipmentTypes.FedExPwd "
query = query & "FROM ShipmentTypes "
query = query & "WHERE (((ShipmentTypes.idShipment)=9));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if NOT rs.eof then
	FedExAccountNumber=rs("userID")
	FedExMeterNumber=rs("password")
	pcv_strEnvironment=rs("AccessLicense")
	FedExkey=rs("FedExKey")
	FedExPassword=rs("FedExPwd")
end if
set rs=nothing

if (FedEXWS_active=true or FedExWS_active="-1") AND FedEXWS_AccountNumber<>"" then

	iFedExWSActive=1
	dim arryFedExWSService
	dim arryFedExWSRate
	dim arrFedExWSDeliveryDate
	arryFedExWSService=""
	arryFedExWSRate=""
	arrFedExWSDeliveryDate=""
	arryFedExWSRate2 = ""

	pcv_TmpListRate = FEDEXWS_LISTRATE
	pcv_TmpSaturdayDelivery = 0

	'// Override List Rates for International addresses
	If Universal_destination_country<>"US" Then
		pcv_TmpListRate = "0"
	End If

	'// FedEx EXPRESS RATES
	set objFedExWSClass = New pcFedExWSClass

	fedex_postdataWS=""
	FEDEXWS_result=""

	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'// START: FedEx Express
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////

	objFedExWSClass.NewXMLSubscription pcv_strMethodNameWS, FedExkey, FedExPassword, pcv_strVersion, "rate"

		objFedExWSClass.WriteParent "ClientDetail", pcv_strVersion, ""
			objFedExWSClass.AddNewNode "AccountNumber", pcv_strVersion, FedExAccountNumber
			objFedExWSClass.AddNewNode "MeterNumber", pcv_strVersion, FedExMeterNumber
			objFedExWSClass.AddNewNode "ClientProductId", pcv_strVersion, "EIPC"
			objFedExWSClass.AddNewNode "ClientProductVersion", pcv_strVersion, "3424"
		objFedExWSClass.WriteParent "ClientDetail", pcv_strVersion, "/"

		'// Transaction ID
		objFedExWSClass.WriteParent "TransactionDetail", pcv_strVersion, ""
			objFedExWSClass.AddNewNode "CustomerTransactionId", pcv_strVersion, "IE-NRF-004"
		objFedExWSClass.WriteParent "TransactionDetail", pcv_strVersion, "/"

		objFedExWSClass.WriteParent "Version", pcv_strVersion, ""
			objFedExWSClass.AddNewNode "ServiceId", pcv_strVersion, "crs"
			objFedExWSClass.AddNewNode "Major", pcv_strVersion, "9"
			objFedExWSClass.AddNewNode "Intermediate", pcv_strVersion, "0"
			objFedExWSClass.AddNewNode "Minor", pcv_strVersion, "0"
		objFedExWSClass.WriteParent "Version", pcv_strVersion, "/"
		objFedExWSClass.AddNewNode "ReturnTransitAndCommit", pcv_strVersion, "1"
		objFedExWSClass.WriteSingleParent "CarrierCodes", pcv_strVersion, "FDXE"

		objFedExWSClass.WriteParent "RequestedShipment", pcv_strVersion, ""

			'// Ship Date
			pcAddDay = 	FEDEXWS_ADDDAY
			if pcAddDay&""="" then
				pcAddDay = 0
			end if

			pcFedExShipDate = DateAdd("D", pcAddDay, Date)
			pcWeekDay = WeekDay(pcFedExShipDate)
			if pcWeekDay = "1" then
				'Add a day to the current date
				pcFedExShipDate = DateAdd("D", 1, Date)
			end if
			'if saturday pickup=0 then we need to shift a saturday date to monday
			if pcWeekDay = "7" AND FEDEXWS_SATURDAYPICKUP = "0" then
				'Add 2 days to the current date
				pcFedExShipDate = DateAdd("D", 2, Date)
			end if
			FwsYear = Year(pcFedExShipDate)
			FwsMonth = Month(pcFedExShipDate)
			if len(FwsMonth)=1 Then
				FwsMonth = "0"&FwsMonth
			end if
			FwsDay = Day(pcFedExShipDate)
			if len(FwsDay)=1 Then
				FwsDay = "0"&FwsDay
			end if

			pcFedExFormatedShipDate = FwsYear&"-"&FwsMonth&"-"&FwsDay&"T14:33:57+05:30"

			objFedExWSClass.AddNewNode "ShipTimestamp", pcv_strVersion, pcFedExFormatedShipDate
			'//Identifies the date and time the package is tendered to FedEx. Both the date and time portions of the string are expected to be used. The date should not be a past date or a date more than 10 days in the future. The time is the local time of the shipment based on the shipper's time zone. The date component must be in the format: YYYY-MM-DD (e.g. 2006-06-26). The time component must be in the format: HH:MM:SS using a 24 hour clock (e.g. 11:00 a.m. is 11:00:00, whereas 5:00 p.m. is 17:00:00). The date and time parts are separated by the letter T (e.g. 2006-06-26T17:00:00). There is also a UTC offset component indicating the number of hours/mainutes from UTC (e.g 2006-06-26T17:00:00-0400 is defined form June 26, 2006 5:00 pm Eastern Time).</xs:documentation>

			objFedExWSClass.AddNewNode "ServiceType", pcv_strVersion, ""
			objFedExWSClass.AddNewNode "DropoffType", pcv_strVersion, FEDEXWS_DROPOFF_TYPE
			objFedExWSClass.AddNewNode "PackagingType", pcv_strVersion, FEDEXWS_FEDEX_PACKAGE
			objFedExWSClass.WriteParent "Shipper", pcv_strVersion, ""

				objFedExWSClass.WriteParent "Contact", pcv_strVersion, ""
					objFedExWSClass.AddNewNode "PersonName", pcv_strVersion, scOriginPersonName
					objFedExWSClass.AddNewNode "PhoneNumber", pcv_strVersion, scOriginPhoneNumber
				objFedExWSClass.WriteParent "Contact", pcv_strVersion, "/"

				objFedExWSClass.WriteParent "Address", pcv_strVersion, ""
					objFedExWSClass.AddNewNode "StreetLines", pcv_strVersion, scShipFromAddress1
					objFedExWSClass.AddNewNode "City", pcv_strVersion, scShipFromCity
					objFedExWSClass.AddNewNode "StateOrProvinceCode", pcv_strVersion, scShipFromState
					objFedExWSClass.AddNewNode "PostalCode", pcv_strVersion, scShipFromPostalCode
					objFedExWSClass.AddNewNode "CountryCode", pcv_strVersion, scShipFromPostalCountry
				objFedExWSClass.WriteParent "Address", pcv_strVersion, "/"


			objFedExWSClass.WriteParent "Shipper", pcv_strVersion, "/"

			objFedExWSClass.WriteParent "Recipient", pcv_strVersion, ""
				objFedExWSClass.WriteParent "Contact", pcv_strVersion, ""
					objFedExWSClass.AddNewNode "PersonName", pcv_strVersion, "RECIPIENT"
				objFedExWSClass.WriteParent "Contact", pcv_strVersion, "/"

				objFedExWSClass.WriteParent "Address", pcv_strVersion, ""
					'objFedExWSClass.AddNewNode "StreetLines", pcv_strVersion, ""
					'objFedExWSClass.AddNewNode "City", pcv_strVersion, Universal_destination_city
					if Universal_destination_country="US" OR Universal_destination_country="CA" then
						objFedExWSClass.AddNewNode "StateOrProvinceCode", pcv_strVersion, Universal_destination_provOrState
					end if
					objFedExWSClass.AddNewNode "PostalCode", pcv_strVersion, Universal_destination_postal
					objFedExWSClass.AddNewNode "CountryCode", pcv_strVersion, Universal_destination_country
					if pResidentialShipping="-1" or pResidentialShipping="1" then
						objFedExWSClass.AddNewNode "Residential", pcv_strVersion, "true"
					else
						objFedExWSClass.AddNewNode "Residential", pcv_strVersion, "false"
					end if
				objFedExWSClass.WriteParent "Address", pcv_strVersion, "/"

			objFedExWSClass.WriteParent "Recipient", pcv_strVersion, "/"

			objFedExWSClass.WriteParent "ShippingChargesPayment", pcv_strVersion, ""
				objFedExWSClass.AddNewNode "PaymentType", pcv_strVersion, "SENDER"
			objFedExWSClass.WriteParent "ShippingChargesPayment", pcv_strVersion, "/"

			If pcv_TmpSaturdayDelivery<>"0" Then
				objFedExWSClass.WriteParent "SpecialServicesRequested", pcv_strVersion, ""
					'// Saturday Delivery
					objFedExWSClass.AddNewNode "SpecialServiceTypes", pcv_strVersion, "SATURDAY_DELIVERY"
					'// Saturday Pickup
					'objFedExWSClass.AddNewNode "SpecialServiceTypes", pcv_strVersion, "SATURDAY_PICKUP"
				objFedExWSClass.WriteParent "SpecialServicesRequested", pcv_strVersion, "/"
			End If

			if pcv_TmpListRate = "-1" OR pcv_TmpListRate = -1 then
				pcv_strRateRequestType = "LIST"
			elseif pcv_TmpListRate = "-2" OR pcv_TmpListRate = -2 then
				pcv_strRateRequestType = "PREFERRED"
			else
				pcv_strRateRequestType = "ACCOUNT"
			end if
			objFedExWSClass.AddNewNode "RateRequestTypes", pcv_strVersion, pcv_strRateRequestType
			objFedExWSClass.AddNewNode "PackageCount", pcv_strVersion, pcv_intPackageNum
			objFedExWSClass.AddNewNode "PackageDetail", pcv_strVersion, "INDIVIDUAL_PACKAGES"


			for q=1 to pcv_intPackageNum
				'//FEDEXWSWEIGHTCHANGE///////////////////////////////////////
				session("FedEXWSPackWeight"&q)
				tmpPounds = session("FEDEXWSPackPounds"&q)
				tmpOunces = session("FEDEXWSPackOunces"&q)
				if NOT isNumeric(tmpPounds) then
					tmpPounds = 0
				end if
				if NOT isNumeric(tmpOunces) then
					tmpOunces = 0
				end if
				tmpOuncesDec = Cint(tmpOunces)/16

				objFedExWSClass.WriteParent "RequestedPackageLineItems", pcv_strVersion, ""

					objFedExWSClass.AddNewNode "SequenceNumber", pcv_strVersion, q

					pcv_TempPackPrice=session("FedEXWSPackPrice"&q)
					If pcv_TempPackPrice="" Then
						pcv_TempPackPrice="100.00"
					End If

					objFedExWSClass.WriteParent "InsuredValue", pcv_strVersion, ""
						objFedExWSClass.AddNewNode "Currency", pcv_strVersion, "USD"
						objFedExWSClass.AddNewNode "Amount", pcv_strVersion, replace(money(pcv_TempPackPrice),",","")
					objFedExWSClass.WriteParent "InsuredValue", pcv_strVersion, "/"

					objFedExWSClass.WriteParent "Weight", pcv_strVersion, ""
						If scShipFromWeightUnit="LBS" Then
							objFedExWSClass.WriteSingleParent "Units", pcv_strVersion, "LB"
						Else
							objFedExWSClass.WriteSingleParent "Units", pcv_strVersion, "KG"
						End If
						objFedExWSClass.WriteSingleParent "Value", pcv_strVersion, round((tmpPounds + tmpOuncesDec),1)
					objFedExWSClass.WriteParent "Weight", pcv_strVersion, "/"

					if ((FEDEXWS_FEDEX_PACKAGE="YOUR_PACKAGING") AND (session("FedEXWSPackLength"&q)<>"" AND session("FedEXWSPackWidth"&q)<>"" AND session("FedEXWSPackHeight"&q)<>"")) then
						pcv_strDimUnit = FEDEXWS_DIM_UNIT
						if pcv_strDimUnit="" then
							pcv_strDimUnit = "IN"
						end if
						objFedExWSClass.WriteParent "Dimensions", pcv_strVersion, ""
							objFedExWSClass.AddNewNode "Length", pcv_strVersion, Int(session("FedEXWSPackLength"&q))
							objFedExWSClass.AddNewNode "Width", pcv_strVersion, Int(session("FedEXWSPackWidth"&q))
							objFedExWSClass.AddNewNode "Height", pcv_strVersion, Int(session("FedEXWSPackHeight"&q))
							objFedExWSClass.AddNewNode "Units", pcv_strVersion, FEDEXWS_DIM_UNIT
						objFedExWSClass.WriteParent "Dimensions", pcv_strVersion, "/"
					end if

				objFedExWSClass.WriteParent "RequestedPackageLineItems", pcv_strVersion, "/"

			next

		objFedExWSClass.WriteParent "RequestedShipment", pcv_strVersion, "/"

	objFedExWSClass.EndXMLTransaction pcv_strMethodNameWS, pcv_strVersion

'--------------------------------------------------------------------------------------------------------


	'// Print out our newly formed request xml
	'response.Clear()
	'response.ContentType="text/xml"
	'response.Write(fedex_postdataWS)
	'response.End()

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Send Our Transaction.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	call objFedExWSClass.SendXMLRequest(fedex_postdataWS, pcv_strEnvironment)

	'response.Clear()
	'response.ContentType="text/xml"
	'response.Write(FEDEXWS_result)
	'response.End()

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Load Our Response.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	call objFedExWSClass.LoadXMLResults(FEDEXWS_result)


	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Check for errors from FedEx.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcv_strErrorMsgWS = objFedExWSClass.ReadResponseNode("//v9:Notifications", "v9:Severity")
	if pcv_strErrorMsgWS<>"SUCCESS" AND  pcv_strErrorMsgWS<>"NOTE" AND pcv_strErrorMsgWS<>"WARNING"  then
		pcv_strErrorMsgWS = objFedExWSClass.ReadResponseNode("//v9:Notifications", "v9:Message")
	else
		pcv_strErrorMsgWS = ""
	end if

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Redirect with a Message OR complete some task.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	scHideEstimateDeliveryTimes = "0"
	If NOT len(pcv_strErrorMsgWS)>0 Then
		'// Generate FedEx Arrays
		if pcv_TmpListRate = "-1" OR pcv_TmpListRate = -1 then
			arryFedExWSService = arryFedExWSService & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:ServiceType")
			if scHideEstimateDeliveryTimes="1" then
				arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & ""
				arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & ""
			else
				arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:DeliveryDayOfWeek")
				arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:DeliveryTimestamp")
			end if
			arryFedExWSRate = arryFedExWSRate & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:RatedShipmentDetails/v9:ShipmentRateDetail/v9:TotalNetFedExCharge/v9:Amount")
			Set objOutputXMLDocss = Server.CreateObject("Microsoft.XMLDOM")
			objOutputXMLDocss.loadXML FEDEXWS_result
			Set Nodes = objOutputXMLDocss.selectNodes("//RateReplyDetails")

			set objLst=objOutputXMLDocss.getElementsByTagName("v9:RateReplyDetails")

			tempGetRate = 0
			for i = 0 to (objLst.length - 1)
				for j=0 to ((objLst.item(i).childNodes.length)-1)
					If objLst.item(i).childNodes(j).nodeName="v9:RatedShipmentDetails" then
						for m=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="v9:ShipmentRateDetail" then
								for n=0 to ((objLst.item(i).childNodes(j).childNodes(m).childNodes.length)-1)
									If objLst.item(i).childNodes(j).childNodes(m).childNodes(n).nodeName="v9:RateType" then
										if objLst.item(i).childNodes(j).childNodes(m).childNodes(n).text = "RATED_LIST_PACKAGE" Then
											tempGetRate = 1
										end if
									End If
									If tempGetRate = 1 AND objLst.item(i).childNodes(j).childNodes(m).childNodes(n).nodeName="v9:TotalNetFedExCharge" Then
										for p=0 to ((objLst.item(i).childNodes(j).childNodes(m).childNodes(n).childNodes.length)-1)
											if objLst.item(i).childNodes(j).childNodes(m).childNodes(n).childNodes(p).nodeName = "v9:Amount" Then
												arryFedExWSRate2 = arryFedExWSRate2 & objLst.item(i).childNodes(j).childNodes(m).childNodes(n).childNodes(p).text &","
												tempGetRate = 0
											end if
										next
									End If
								next
							end if
						next
					End if
				next
			next
				arryFedExWSRate = arryFedExWSRate2
		else
			arryFedExWSService = arryFedExWSService & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:ServiceType")
			if scHideEstimateDeliveryTimes="1" then
				arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & ""
				arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & ""
			else
				arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:DeliveryDayOfWeek")
				arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:DeliveryTimestamp")
			end if
			arryFedExWSRate = arryFedExWSRate & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:RatedShipmentDetails/v9:ShipmentRateDetail/v9:TotalNetFedExCharge/v9:Amount")
		end if
	End If

	set objFedExWSClass = nothing
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'// END: FedEx Express
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////



	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'// START: FedEx Smart Post
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	set objFedExWSClass = New pcFedExWSClass

	fedex_postdataWS=""
	FEDEXWS_result=""

	objFedExWSClass.NewXMLSubscription pcv_strMethodNameWS, FedExkey, FedExPassword, pcv_strVersion, "rate"

		objFedExWSClass.WriteParent "ClientDetail", pcv_strVersion, ""
			objFedExWSClass.AddNewNode "AccountNumber", pcv_strVersion, FedExAccountNumber
			objFedExWSClass.AddNewNode "MeterNumber", pcv_strVersion, FedExMeterNumber
			objFedExWSClass.AddNewNode "ClientProductId", pcv_strVersion, "EIPC"
			objFedExWSClass.AddNewNode "ClientProductVersion", pcv_strVersion, "3424"
		objFedExWSClass.WriteParent "ClientDetail", pcv_strVersion, "/"

		'// Transaction ID
		objFedExWSClass.WriteParent "TransactionDetail", pcv_strVersion, ""
			objFedExWSClass.AddNewNode "CustomerTransactionId", pcv_strVersion, "DOM-GND-001"
		objFedExWSClass.WriteParent "TransactionDetail", pcv_strVersion, "/"

		objFedExWSClass.WriteParent "Version", pcv_strVersion, ""
			objFedExWSClass.AddNewNode "ServiceId", pcv_strVersion, "crs"
			objFedExWSClass.AddNewNode "Major", pcv_strVersion, "9"
			objFedExWSClass.AddNewNode "Intermediate", pcv_strVersion, "0"
			objFedExWSClass.AddNewNode "Minor", pcv_strVersion, "0"
		objFedExWSClass.WriteParent "Version", pcv_strVersion, "/"
		objFedExWSClass.AddNewNode "ReturnTransitAndCommit", pcv_strVersion, "1"
		objFedExWSClass.WriteParent "RequestedShipment", pcv_strVersion, ""

			'// Ship Date
			pcAddDay = 	FEDEXWS_ADDDAY
			if pcAddDay&""="" then
				pcAddDay = 0
			end if

			pcFedExShipDate = DateAdd("D", pcAddDay, Date)
			pcWeekDay = WeekDay(pcFedExShipDate)
			if pcWeekDay = "1" then
				'Add a day to the current date
				pcFedExShipDate = DateAdd("D", 1, Date)
			end if
			'if saturday pickup=0 then we need to shift a saturday date to monday
			if pcWeekDay = "7" AND FEDEXWS_SATURDAYPICKUP = "0" then
				'Add 2 days to the current date
				pcFedExShipDate = DateAdd("D", 2, Date)
			end if
			FwsYear = Year(pcFedExShipDate)
			FwsMonth = Month(pcFedExShipDate)
			if len(FwsMonth)=1 Then
				FwsMonth = "0"&FwsMonth
			end if
			FwsDay = Day(pcFedExShipDate)
			if len(FwsDay)=1 Then
				FwsDay = "0"&FwsDay
			end if

			pcFedExFormatedShipDate = FwsYear&"-"&FwsMonth&"-"&FwsDay&"T14:33:57+05:30"

			objFedExWSClass.AddNewNode "ShipTimestamp", pcv_strVersion, pcFedExFormatedShipDate

			objFedExWSClass.AddNewNode "DropoffType", pcv_strVersion, FEDEXWS_DROPOFF_TYPE
			objFedExWSClass.AddNewNode "ServiceType", pcv_strVersion, "SMART_POST"
			objFedExWSClass.AddNewNode "PackagingType", pcv_strVersion, FEDEXWS_FEDEX_PACKAGE
			objFedExWSClass.WriteParent "Shipper", pcv_strVersion, ""
				'objFedExWSClass.WriteParent "Contact", pcv_strVersion, ""
					'objFedExWSClass.AddNewNode "PersonName", pcv_strVersion, scOriginPersonName
					'objFedExWSClass.AddNewNode "PhoneNumber", pcv_strVersion, scOriginPhoneNumber
				'objFedExWSClass.WriteParent "Contact", pcv_strVersion, "/"

				objFedExWSClass.WriteParent "Address", pcv_strVersion, ""
					'objFedExWSClass.AddNewNode "StreetLines", pcv_strVersion, scShipFromAddress1
					'objFedExWSClass.AddNewNode "City", pcv_strVersion, scShipFromCity
					'objFedExWSClass.AddNewNode "StateOrProvinceCode", pcv_strVersion, scShipFromState
					objFedExWSClass.AddNewNode "PostalCode", pcv_strVersion, scShipFromPostalCode
					objFedExWSClass.AddNewNode "CountryCode", pcv_strVersion, scShipFromPostalCountry
				objFedExWSClass.WriteParent "Address", pcv_strVersion, "/"
			objFedExWSClass.WriteParent "Shipper", pcv_strVersion, "/"

			objFedExWSClass.WriteParent "Recipient", pcv_strVersion, ""
				'objFedExWSClass.WriteParent "Contact", pcv_strVersion, ""
					'objFedExWSClass.AddNewNode "PersonName", pcv_strVersion, "RECIPIENT"
				'objFedExWSClass.WriteParent "Contact", pcv_strVersion, "/"

				objFedExWSClass.WriteParent "Address", pcv_strVersion, ""
					'objFedExWSClass.AddNewNode "StreetLines", pcv_strVersion, ""
					'objFedExWSClass.AddNewNode "City", pcv_strVersion, Universal_destination_city
					if Universal_destination_country="US" OR Universal_destination_country="CA" then
						'objFedExWSClass.AddNewNode "StateOrProvinceCode", pcv_strVersion, Universal_destination_provOrState
					end if
					objFedExWSClass.AddNewNode "PostalCode", pcv_strVersion, Universal_destination_postal
					objFedExWSClass.AddNewNode "CountryCode", pcv_strVersion, Universal_destination_country
					if pResidentialShipping="-1" or pResidentialShipping="1" then
						objFedExWSClass.AddNewNode "Residential", pcv_strVersion, "true"
					else
						objFedExWSClass.AddNewNode "Residential", pcv_strVersion, "false"
					end if
				objFedExWSClass.WriteParent "Address", pcv_strVersion, "/"

			objFedExWSClass.WriteParent "Recipient", pcv_strVersion, "/"
			objFedExWSClass.WriteParent "SmartPostDetail", pcv_strVersion, ""
				objFedExWSClass.AddNewNode "Indicia", pcv_strVersion, "PARCEL_SELECT"
				objFedExWSClass.AddNewNode "HubId", pcv_strVersion, FDXWS_SMHUBID
				objFedExWSClass.AddNewNode "CustomerManifestId", pcv_strVersion, "String"
			objFedExWSClass.WriteParent "SmartPostDetail", pcv_strVersion, "/"

			if pcv_TmpListRate = "-1" OR pcv_TmpListRate = -1 then
				pcv_strRateRequestType = "LIST"
			else
				pcv_strRateRequestType = "ACCOUNT"
			end if
			objFedExWSClass.AddNewNode "RateRequestTypes", pcv_strVersion, pcv_strRateRequestType
			objFedExWSClass.AddNewNode "PackageCount", pcv_strVersion, pcv_intPackageNum
			objFedExWSClass.AddNewNode "PackageDetail", pcv_strVersion, "INDIVIDUAL_PACKAGES"
			
			for q=1 to pcv_intPackageNum
				'//FEDEXWSWEIGHTCHANGE///////////////////////////////////////
				session("FedEXWSPackWeight"&q)
				tmpPounds = session("FEDEXWSPackPounds"&q)
				tmpOunces = session("FEDEXWSPackOunces"&q)
				if NOT isNumeric(tmpPounds) then
					tmpPounds = 0
				end if
				if NOT isNumeric(tmpOunces) then
					tmpOunces = 0
				end if
				tmpOuncesDec = Cint(tmpOunces)/16
			
				objFedExWSClass.WriteParent "RequestedPackageLineItems", pcv_strVersion, ""
				objFedExWSClass.AddNewNode "SequenceNumber", pcv_strVersion, q
				objFedExWSClass.WriteParent "Weight", pcv_strVersion, ""
				If scShipFromWeightUnit="LBS" Then
					objFedExWSClass.WriteSingleParent "Units", pcv_strVersion, "LB"
				Else
					objFedExWSClass.WriteSingleParent "Units", pcv_strVersion, "KG"
				End If
				objFedExWSClass.WriteSingleParent "Value", pcv_strVersion, tmpPounds + tmpOuncesDec
				objFedExWSClass.WriteParent "Weight", pcv_strVersion, "/"
				if ((FEDEXWS_FEDEX_PACKAGE="YOUR_PACKAGING") AND (session("FedEXWSPackLength"&q)<>"" AND session("FedEXWSPackWidth"&q)<>"" AND session("FedEXWSPackHeight"&q)<>"")) then
					pcv_strDimUnit = FEDEXWS_DIM_UNIT
					if pcv_strDimUnit="" then
						pcv_strDimUnit = "IN"
					end if
					objFedExWSClass.WriteParent "Dimensions", pcv_strVersion, ""
						objFedExWSClass.AddNewNode "Length", pcv_strVersion, Int(session("FedEXWSPackLength"&q))
						objFedExWSClass.AddNewNode "Width", pcv_strVersion, Int(session("FedEXWSPackWidth"&q))
						objFedExWSClass.AddNewNode "Height", pcv_strVersion, Int(session("FedEXWSPackHeight"&q))
						objFedExWSClass.AddNewNode "Units", pcv_strVersion, FEDEXWS_DIM_UNIT
					objFedExWSClass.WriteParent "Dimensions", pcv_strVersion, "/"
				end if

				objFedExWSClass.WriteParent "RequestedPackageLineItems", pcv_strVersion, "/"

			next

		objFedExWSClass.WriteParent "RequestedShipment", pcv_strVersion, "/"

	objFedExWSClass.EndXMLTransaction pcv_strMethodNameWS, pcv_strVersion

	'--------------------------------------------------------------------------------------------------------


	'// Print out our newly formed request xml
	'response.Clear()
	'response.ContentType="text/xml"
	'response.Write(fedex_postdataWS)
	'response.End()

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Send Our Transaction.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	call objFedExWSClass.SendXMLRequest(fedex_postdataWS, pcv_strEnvironment)

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Load Our Response.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// Print out our newly formed request xml
	'response.Clear()
	'response.ContentType="text/xml"
	'response.Write(FEDEXWS_result)
	'response.End()
	call objFedExWSClass.LoadXMLResults(FEDEXWS_result)

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Check for errors from FedEx.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'call objFedExWSClass.XMLResponseVerify(ErrPageName)
	pcv_strErrorMsgWS = objFedExWSClass.ReadResponseNode("//v9:Notifications", "v9:Severity")
	if pcv_strErrorMsgWS<>"SUCCESS" AND  pcv_strErrorMsgWS<>"NOTE" AND pcv_strErrorMsgWS<>"WARNING"  then
		pcv_strErrorMsgWS = objFedExWSClass.ReadResponseNode("//v9:Notifications", "v9:Message")
	else
		pcv_strErrorMsgWS = ""
	end if

	'/////////////////////////////////////////////////////////////
	'// BASELINE LOGGING
	'/////////////////////////////////////////////////////////////
	'// Log our Transaction
	'call objFedExWSClass.pcs_LogTransaction(fedex_postdataWS, pcv_strMethodNameWS&"_SP_"&q&".in", true)
	'// Log our Response
	'call objFedExWSClass.pcs_LogTransaction(FEDEXWS_result, pcv_strMethodNameWS&"_SP_"&q&".out", true)

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Redirect with a Message OR complete some task.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	scHideEstimateDeliveryTimes = "0"
	If NOT len(pcv_strErrorMsgWS)>0 Then
		'// Generate FedEx Arrays
		if pcv_TmpListRate = "-1" OR pcv_TmpListRate = -1 then
			arryFedExWSService = arryFedExWSService & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:ServiceType")
			if scHideEstimateDeliveryTimes="1" then
				arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & ""
				arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & ""
			else
				pcv_minimumDeliverDays = objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:CommitDetails/v9:TransitTime") 'TWO_DAYS
				pcv_maximumDeliverDays = objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:CommitDetails/v9:MaximumTransitTime") 'EIGHT_DAYS
				arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & "SPTransitTime:"&objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:TransitTime")
				arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & "SPTransitTime:"&objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:CommitDetails/v9:MaximumTransitTime")
			end if
			arryFedExWSRate = arryFedExWSRate & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:RatedShipmentDetails/v9:ShipmentRateDetail/v9:TotalNetFedExCharge/v9:Amount")
			' Parse the XML document.
			Set objOutputXMLDocss = Server.CreateObject("Microsoft.XMLDOM")
			objOutputXMLDocss.loadXML FEDEXWS_result
			Set Nodes = objOutputXMLDocss.selectNodes("//RateReplyDetails")
			set objLst=objOutputXMLDocss.getElementsByTagName("v9:RateReplyDetails")
			tempGetRate = 0
			for i = 0 to (objLst.length - 1)
				for j=0 to ((objLst.item(i).childNodes.length)-1)
					If objLst.item(i).childNodes(j).nodeName="v9:RatedShipmentDetails" then
						for m=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="v9:ShipmentRateDetail" then
								for n=0 to ((objLst.item(i).childNodes(j).childNodes(m).childNodes.length)-1)
									If objLst.item(i).childNodes(j).childNodes(m).childNodes(n).nodeName="v9:RateType" then
										if objLst.item(i).childNodes(j).childNodes(m).childNodes(n).text = "RATED_LIST_PACKAGE" Then
											tempGetRate = 1
										end if
										if tempGetRate = 0 then
											if objLst.item(i).childNodes(j).childNodes(m).childNodes(n).text = "PAYOR_LIST_PACKAGE" Then
												tempGetRate = 1
											end if
										end if
									End If
									If tempGetRate = 1 AND objLst.item(i).childNodes(j).childNodes(m).childNodes(n).nodeName="v9:TotalNetFedExCharge" Then
										for p=0 to ((objLst.item(i).childNodes(j).childNodes(m).childNodes(n).childNodes.length)-1)
											if objLst.item(i).childNodes(j).childNodes(m).childNodes(n).childNodes(p).nodeName = "v9:Amount" Then
												arryFedExWSRate2 = arryFedExWSRate2 & objLst.item(i).childNodes(j).childNodes(m).childNodes(n).childNodes(p).text &","
												tempGetRate = 0
											end if
										next
									End If
								next
							end if
						next
					End if
				next
			next
			arryFedExWSRate = arryFedExWSRate2
		else
			arryFedExWSService = arryFedExWSService & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:ServiceType")
			if scHideEstimateDeliveryTimes="1" then
				arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & ""
				arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & ""
			else
				arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:DeliveryDayOfWeek")
				arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:DeliveryTimestamp")
			end if
			arryFedExWSRate = arryFedExWSRate & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:RatedShipmentDetails/v9:ShipmentRateDetail/v9:TotalNetFedExCharge/v9:Amount")
		end if
	End If
	set objFedExWSClass = nothing
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'// END: FedEx SmartPost
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////

	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'// START: FedEx Ground
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	set objFedExWSClass = New pcFedExWSClass

	fedex_postdataWS=""
	FEDEXWS_result=""

	objFedExWSClass.NewXMLSubscription pcv_strMethodNameWS, FedExkey, FedExPassword, pcv_strVersion, "rate"

		objFedExWSClass.WriteParent "ClientDetail", pcv_strVersion, ""
			objFedExWSClass.AddNewNode "AccountNumber", pcv_strVersion, FedExAccountNumber
			objFedExWSClass.AddNewNode "MeterNumber", pcv_strVersion, FedExMeterNumber
			objFedExWSClass.AddNewNode "ClientProductId", pcv_strVersion, "EIPC"
			objFedExWSClass.AddNewNode "ClientProductVersion", pcv_strVersion, "3424"
		objFedExWSClass.WriteParent "ClientDetail", pcv_strVersion, "/"

		'// Transaction ID
		objFedExWSClass.WriteParent "TransactionDetail", pcv_strVersion, ""
			objFedExWSClass.AddNewNode "CustomerTransactionId", pcv_strVersion, "DOM-GND-001"
		objFedExWSClass.WriteParent "TransactionDetail", pcv_strVersion, "/"

		objFedExWSClass.WriteParent "Version", pcv_strVersion, ""
			objFedExWSClass.AddNewNode "ServiceId", pcv_strVersion, "crs"
			objFedExWSClass.AddNewNode "Major", pcv_strVersion, "9"
			objFedExWSClass.AddNewNode "Intermediate", pcv_strVersion, "0"
			objFedExWSClass.AddNewNode "Minor", pcv_strVersion, "0"
		objFedExWSClass.WriteParent "Version", pcv_strVersion, "/"
		objFedExWSClass.AddNewNode "ReturnTransitAndCommit", pcv_strVersion, "1"
		objFedExWSClass.WriteSingleParent "CarrierCodes", pcv_strVersion, "FDXG"

		objFedExWSClass.WriteParent "RequestedShipment", pcv_strVersion, ""

			'// Ship Date
			pcAddDay = 	FEDEXWS_ADDDAY
			if pcAddDay&""="" then
				pcAddDay = 0
			end if

			pcFedExShipDate = DateAdd("D", pcAddDay, Date)
			pcWeekDay = WeekDay(pcFedExShipDate)
			if pcWeekDay = "1" then
				'Add a day to the current date
				pcFedExShipDate = DateAdd("D", 1, Date)
			end if
			'if saturday pickup=0 then we need to shift a saturday date to monday
			if pcWeekDay = "7" AND FEDEXWS_SATURDAYPICKUP = "0" then
				'Add 2 days to the current date
				pcFedExShipDate = DateAdd("D", 2, Date)
			end if
			FwsYear = Year(pcFedExShipDate)
			FwsMonth = Month(pcFedExShipDate)
			if len(FwsMonth)=1 Then
				FwsMonth = "0"&FwsMonth
			end if
			FwsDay = Day(pcFedExShipDate)
			if len(FwsDay)=1 Then
				FwsDay = "0"&FwsDay
			end if

			pcFedExFormatedShipDate = FwsYear&"-"&FwsMonth&"-"&FwsDay&"T14:33:57+05:30"

			objFedExWSClass.AddNewNode "ShipTimestamp", pcv_strVersion, pcFedExFormatedShipDate

			objFedExWSClass.AddNewNode "ServiceType", pcv_strVersion, ""
			objFedExWSClass.AddNewNode "DropoffType", pcv_strVersion, FEDEXWS_DROPOFF_TYPE
			objFedExWSClass.AddNewNode "PackagingType", pcv_strVersion, FEDEXWS_FEDEX_PACKAGE
			objFedExWSClass.WriteParent "Shipper", pcv_strVersion, ""
				objFedExWSClass.WriteParent "Contact", pcv_strVersion, ""
					objFedExWSClass.AddNewNode "PersonName", pcv_strVersion, scOriginPersonName
					objFedExWSClass.AddNewNode "PhoneNumber", pcv_strVersion, scOriginPhoneNumber
				objFedExWSClass.WriteParent "Contact", pcv_strVersion, "/"

				objFedExWSClass.WriteParent "Address", pcv_strVersion, ""
					objFedExWSClass.AddNewNode "StreetLines", pcv_strVersion, scShipFromAddress1
					objFedExWSClass.AddNewNode "City", pcv_strVersion, scShipFromCity
					objFedExWSClass.AddNewNode "StateOrProvinceCode", pcv_strVersion, scShipFromState
					objFedExWSClass.AddNewNode "PostalCode", pcv_strVersion, scShipFromPostalCode
					objFedExWSClass.AddNewNode "CountryCode", pcv_strVersion, scShipFromPostalCountry
				objFedExWSClass.WriteParent "Address", pcv_strVersion, "/"


			objFedExWSClass.WriteParent "Shipper", pcv_strVersion, "/"

			objFedExWSClass.WriteParent "Recipient", pcv_strVersion, ""
				objFedExWSClass.WriteParent "Contact", pcv_strVersion, ""
					objFedExWSClass.AddNewNode "PersonName", pcv_strVersion, "RECIPIENT"
				objFedExWSClass.WriteParent "Contact", pcv_strVersion, "/"

				objFedExWSClass.WriteParent "Address", pcv_strVersion, ""
					'objFedExWSClass.AddNewNode "StreetLines", pcv_strVersion, ""
					'objFedExWSClass.AddNewNode "City", pcv_strVersion, Universal_destination_city
					if Universal_destination_country="US" OR Universal_destination_country="CA" then
						objFedExWSClass.AddNewNode "StateOrProvinceCode", pcv_strVersion, Universal_destination_provOrState
					end if
					objFedExWSClass.AddNewNode "PostalCode", pcv_strVersion, Universal_destination_postal
					objFedExWSClass.AddNewNode "CountryCode", pcv_strVersion, Universal_destination_country
					if pResidentialShipping="-1" or pResidentialShipping="1" then
						objFedExWSClass.AddNewNode "Residential", pcv_strVersion, "true"
					else
						objFedExWSClass.AddNewNode "Residential", pcv_strVersion, "false"
					end if
				objFedExWSClass.WriteParent "Address", pcv_strVersion, "/"

			objFedExWSClass.WriteParent "Recipient", pcv_strVersion, "/"

			objFedExWSClass.WriteParent "ShippingChargesPayment", pcv_strVersion, ""
				objFedExWSClass.AddNewNode "PaymentType", pcv_strVersion, "SENDER"
			objFedExWSClass.WriteParent "ShippingChargesPayment", pcv_strVersion, "/"

			'// ShipDate
			'objFedExWSClass.AddNewNode "ShipDate", pcv_strVersion, "2011-08-03T14:33:57+05:30"
			'objFedExWSClass.WriteParent "ShipDate", pcv_strVersion, ""
			'	objFedExWSClass.AddNewNode "ShipDate", pcv_strVersion, "2011-08-03"
			'objFedExWSClass.WriteParent "ShipDate", pcv_strVersion, "/"

			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// Start: FDXE Special Services
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			If pcv_TmpSaturdayDelivery<>"0" Then
				objFedExWSClass.WriteParent "SpecialServicesRequested", pcv_strVersion, ""
					'// Saturday Delivery
					objFedExWSClass.AddNewNode "SpecialServiceTypes", pcv_strVersion, "SATURDAY_DELIVERY"
					'// Saturday Pickup
					'objFedExWSClass.AddNewNode "SpecialServiceTypes", pcv_strVersion, "SATURDAY_PICKUP"
				objFedExWSClass.WriteParent "SpecialServicesRequested", pcv_strVersion, "/"
			End If
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// End: FDXE Special Services
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// Start: FDXE International
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'objFedExWSClass.WriteParent "CustomsClearanceDetail", pcv_strVersion, ""

			'// Custom Value
			'objFedExWSClass.WriteParent "CustomsValue", pcv_strVersion, ""
			'		objFedExWSClass.AddNewNode "Currency", pcv_strVersion, "USD"
			'		objFedExWSClass.AddNewNode "Amount", pcv_strVersion, "250"
			'objFedExWSClass.WriteParent "CustomsValue", pcv_strVersion, "/"

			'objFedExWSClass.WriteParent "CustomsClearanceDetail", pcv_strVersion, "/"
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// End: FDXE International
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

			if pcv_TmpListRate = "-1" OR pcv_TmpListRate = -1 then
				pcv_strRateRequestType = "LIST"
			elseif pcv_TmpListRate = "-2" OR pcv_TmpListRate = -2 then
				pcv_strRateRequestType = "PREFERRED"
			else
				pcv_strRateRequestType = "ACCOUNT"
			end if
			
			objFedExWSClass.AddNewNode "RateRequestTypes", pcv_strVersion, pcv_strRateRequestType
			objFedExWSClass.AddNewNode "PackageCount", pcv_strVersion, pcv_intPackageNum
			objFedExWSClass.AddNewNode "PackageDetail", pcv_strVersion, "INDIVIDUAL_PACKAGES"

			for q=1 to pcv_intPackageNum

				'//FEDEXWSWEIGHTCHANGE///////////////////////////////////////
				session("FedEXWSPackWeight"&q)
				tmpPounds = session("FEDEXWSPackPounds"&q)
				tmpOunces = session("FEDEXWSPackOunces"&q)
				if NOT isNumeric(tmpPounds) then
					tmpPounds = 0
				end if
				if NOT isNumeric(tmpOunces) then
					tmpOunces = 0
				end if
				tmpOuncesDec = Cint(tmpOunces)/16

				objFedExWSClass.WriteParent "RequestedPackageLineItems", pcv_strVersion, ""

				objFedExWSClass.AddNewNode "SequenceNumber", pcv_strVersion, q

				pcv_TempPackPrice=session("FedEXWSPackPrice"&q)
				If pcv_TempPackPrice="" Then
					pcv_TempPackPrice="100.00"
				End If

				objFedExWSClass.WriteParent "InsuredValue", pcv_strVersion, ""
					objFedExWSClass.AddNewNode "Currency", pcv_strVersion, "USD"
					objFedExWSClass.AddNewNode "Amount", pcv_strVersion, replace(money(pcv_TempPackPrice),",","")
				objFedExWSClass.WriteParent "InsuredValue", pcv_strVersion, "/"

				objFedExWSClass.WriteParent "Weight", pcv_strVersion, ""
					If scShipFromWeightUnit="LBS" Then
						objFedExWSClass.WriteSingleParent "Units", pcv_strVersion, "LB"
					Else
						objFedExWSClass.WriteSingleParent "Units", pcv_strVersion, "KG"
					End If
					objFedExWSClass.WriteSingleParent "Value", pcv_strVersion, tmpPounds + tmpOuncesDec
				objFedExWSClass.WriteParent "Weight", pcv_strVersion, "/"

				if ((FEDEXWS_FEDEX_PACKAGE="YOUR_PACKAGING") AND (session("FedEXWSPackLength"&q)<>"" AND session("FedEXWSPackWidth"&q)<>"" AND session("FedEXWSPackHeight"&q)<>"")) then
					pcv_strDimUnit = FEDEXWS_DIM_UNIT
					if pcv_strDimUnit="" then
						pcv_strDimUnit = "IN"
					end if
					objFedExWSClass.WriteParent "Dimensions", pcv_strVersion, ""
						objFedExWSClass.AddNewNode "Length", pcv_strVersion, Int(session("FedEXWSPackLength"&q))
						objFedExWSClass.AddNewNode "Width", pcv_strVersion, Int(session("FedEXWSPackWidth"&q))
						objFedExWSClass.AddNewNode "Height", pcv_strVersion, Int(session("FedEXWSPackHeight"&q))
						objFedExWSClass.AddNewNode "Units", pcv_strVersion, FEDEXWS_DIM_UNIT
					objFedExWSClass.WriteParent "Dimensions", pcv_strVersion, "/"
				end if

				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// Start: FDXG Special Services (Package Level)
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				objFedExWSClass.WriteParent "SpecialServicesRequested", pcv_strVersion, ""


						'// NON_STANDARD_CONTAINER
						'objFedExWSClass.AddNewNode "SpecialServiceTypes", pcv_strVersion, "NON_STANDARD_CONTAINER"

						'// COD
						'objFedExWSClass.AddNewNode "SpecialServiceTypes", pcv_strVersion, "COD"
						'objFedExWSClass.WriteParent "CodDetail", pcv_strVersion, ""
							'objFedExWSClass.WriteParent "CodCollectionAmount", pcv_strVersion, ""
							'	objFedExWSClass.AddNewNode "Currency", pcv_strVersion, "USD"
							'	objFedExWSClass.AddNewNode "Amount", pcv_strVersion, "100"
							'objFedExWSClass.WriteParent "CodCollectionAmount", pcv_strVersion, "/"
						'	objFedExWSClass.AddNewNode "CollectionType", pcv_strVersion, "CASH"
						'objFedExWSClass.WriteParent "CodDetail", pcv_strVersion, "/"


						'// Signature Option!
						'objFedExWSClass.AddNewNode "SpecialServiceTypes", pcv_strVersion, "SIGNATURE_OPTION"
						'objFedExWSClass.WriteParent "SignatureOptionDetail", pcv_strVersion, ""
						'	objFedExWSClass.AddNewNode "OptionType", pcv_strVersion, "NO_SIGNATURE_REQUIRED" '// "DIRECT" '// "INDIRECT" '// "ADULT"
						'objFedExWSClass.WriteParent "SignatureOptionDetail", pcv_strVersion, "/"


					objFedExWSClass.WriteParent "SpecialServicesRequested", pcv_strVersion, "/"
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// End: FDXG Special Services (Package Level)
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

					'if pResidentialShipping="-1" or pResidentialShipping="1" then
					'	objFedExWSClass.WriteParent "SpecialServicesRequested", pcv_strVersion, ""
					'		objFedExWSClass.AddNewNode "SpecialServiceTypes", pcv_strVersion, ""
					'	objFedExWSClass.WriteParent "SpecialServicesRequested", pcv_strVersion, "/"
					'end if

				objFedExWSClass.WriteParent "RequestedPackageLineItems", pcv_strVersion, "/"

			next

		objFedExWSClass.WriteParent "RequestedShipment", pcv_strVersion, "/"

	objFedExWSClass.EndXMLTransaction pcv_strMethodNameWS, pcv_strVersion


	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'// END: FedEx Ground
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////


	'// Print out our newly formed request xml
	'response.Clear()
	'response.ContentType="text/xml"
	'response.Write(fedex_postdataWS)
	'response.End()

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Send Our Transaction.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	call objFedExWSClass.SendXMLRequest(fedex_postdataWS, pcv_strEnvironment)
	'// Print out our response
	'response.Clear()
	'response.ContentType="text/xml"
	'response.Write(FEDEXWS_result)
	'response.End()

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Load Our Response.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	call objFedExWSClass.LoadXMLResults(FEDEXWS_result)

	pcv_strErrorMsgWS = cSTR("")

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Check for errors from FedEx.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'call objFedExWSClass.XMLResponseVerify(ErrPageName)
	pcv_strErrorMsgWS = objFedExWSClass.ReadResponseNode("//v9:Notifications", "v9:Severity")
	if pcv_strErrorMsgWS<>"SUCCESS" AND  pcv_strErrorMsgWS<>"NOTE" AND pcv_strErrorMsgWS<>"WARNING"  then
		pcv_strErrorMsgWS = objFedExWSClass.ReadResponseNode("//v9:Notifications", "v9:Message")
	else
		pcv_strErrorMsgWS = ""
	end if

	'/////////////////////////////////////////////////////////////
	'// BASELINE LOGGING
	'/////////////////////////////////////////////////////////////
	'// Log our Transaction
	'call objFedExWSClass.pcs_LogTransaction(fedex_postdataWS, pcv_strMethodNameWS&"_G"&q&".in", true)
	'// Log our Response
	'call objFedExWSClass.pcs_LogTransaction(FEDEXWS_result, pcv_strMethodNameWS&"_G"&q&".out", true)

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Redirect with a Message OR complete some task.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



	If NOT len(pcv_strErrorMsgWS)>0 Then
		'// Generate FedEx Arrays
		if pcv_TmpListRate = "-1" OR pcv_TmpListRate = -1 then
			arryFedExWSService = arryFedExWSService & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:ServiceType")
			if scHideEstimateDeliveryTimes="1" then
				arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & ""
				arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & ""
			else
				arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & "TransitTime:"&objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:TransitTime")
				arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & ""
			end if
			arryFedExWSRate = arryFedExWSRate & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:RatedShipmentDetails/v9:ShipmentRateDetail/v9:TotalNetFedExCharge/v9:Amount")
			' Parse the XML document.
			Set objOutputXMLDocss = Server.CreateObject("Microsoft.XMLDOM")
			objOutputXMLDocss.loadXML FEDEXWS_result
			Set Nodes = objOutputXMLDocss.selectNodes("//RateReplyDetails")
			set objLst=objOutputXMLDocss.getElementsByTagName("v9:RateReplyDetails")
			tempGetRate = 0
			for i = 0 to (objLst.length - 1)
				for j=0 to ((objLst.item(i).childNodes.length)-1)
					If objLst.item(i).childNodes(j).nodeName="v9:RatedShipmentDetails" then
						for m=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="v9:ShipmentRateDetail" then
								for n=0 to ((objLst.item(i).childNodes(j).childNodes(m).childNodes.length)-1)
									If objLst.item(i).childNodes(j).childNodes(m).childNodes(n).nodeName="v9:RateType" then
										if objLst.item(i).childNodes(j).childNodes(m).childNodes(n).text = "PAYOR_LIST_PACKAGE" Then
											tempGetRate = 1
										end if
									End If
									If tempGetRate = 1 AND objLst.item(i).childNodes(j).childNodes(m).childNodes(n).nodeName="v9:TotalNetFedExCharge" Then
										for p=0 to ((objLst.item(i).childNodes(j).childNodes(m).childNodes(n).childNodes.length)-1)
											if objLst.item(i).childNodes(j).childNodes(m).childNodes(n).childNodes(p).nodeName = "v9:Amount" Then
												arryFedExWSRate2 = arryFedExWSRate2 & objLst.item(i).childNodes(j).childNodes(m).childNodes(n).childNodes(p).text &","
												tempGetRate = 0
											end if
										next
									End If
								next
							end if
						next
					End if
				next
			next
			arryFedExWSRate = arryFedExWSRate2
		else
			arryFedExWSService = arryFedExWSService & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:ServiceType")
			if scHideEstimateDeliveryTimes="1" then
				arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & ""
				arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & ""
			else
				arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & "TransitTime:"&objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:TransitTime")
				arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & ""
			end if
			arryFedExWSRate = arryFedExWSRate & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:RatedShipmentDetails/v9:ShipmentRateDetail/v9:TotalNetFedExCharge/v9:Amount")
		end if
	End If

	set objFedExWSClass = nothing
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'// END: FedEx Ground
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////

	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'// START: FedEx Frieght
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	set objFedExWSClass = New pcFedExWSClass

	fedex_postdataWS=""
	FEDEXWS_result=""

	objFedExWSClass.NewXMLSubscription pcv_strMethodNameWS, FedExkey, FedExPassword, pcv_strVersion, "rate"

		objFedExWSClass.WriteParent "ClientDetail", pcv_strVersion, ""
			objFedExWSClass.AddNewNode "AccountNumber", pcv_strVersion, FedExAccountNumber
			objFedExWSClass.AddNewNode "MeterNumber", pcv_strVersion, FedExMeterNumber
			objFedExWSClass.AddNewNode "ClientProductId", pcv_strVersion, "EIPC"
			objFedExWSClass.AddNewNode "ClientProductVersion", pcv_strVersion, "3424"
		objFedExWSClass.WriteParent "ClientDetail", pcv_strVersion, "/"

		'// Transaction ID
		objFedExWSClass.WriteParent "TransactionDetail", pcv_strVersion, ""
			objFedExWSClass.AddNewNode "CustomerTransactionId", pcv_strVersion, "FR-001"
		objFedExWSClass.WriteParent "TransactionDetail", pcv_strVersion, "/"

		objFedExWSClass.WriteParent "Version", pcv_strVersion, ""
			objFedExWSClass.AddNewNode "ServiceId", pcv_strVersion, "crs"
			objFedExWSClass.AddNewNode "Major", pcv_strVersion, "9"
			objFedExWSClass.AddNewNode "Intermediate", pcv_strVersion, "0"
			objFedExWSClass.AddNewNode "Minor", pcv_strVersion, "0"
		objFedExWSClass.WriteParent "Version", pcv_strVersion, "/"
		objFedExWSClass.AddNewNode "ReturnTransitAndCommit", pcv_strVersion, "1"
		objFedExWSClass.WriteSingleParent "CarrierCodes", pcv_strVersion, "FXFR"

		objFedExWSClass.WriteParent "RequestedShipment", pcv_strVersion, ""

			'// Ship Date
			pcAddDay = 	FEDEXWS_ADDDAY
			if pcAddDay&""="" then
				pcAddDay = 0
			end if

			pcFedExShipDate = DateAdd("D", pcAddDay, Date)
			pcWeekDay = WeekDay(pcFedExShipDate)
			if pcWeekDay = "1" then
				'Add a day to the current date
				pcFedExShipDate = DateAdd("D", 1, Date)
			end if
			'if saturday pickup=0 then we need to shift a saturday date to monday
			if pcWeekDay = "7" AND FEDEXWS_SATURDAYPICKUP = "0" then
				'Add 2 days to the current date
				pcFedExShipDate = DateAdd("D", 2, Date)
			end if
			FwsYear = Year(pcFedExShipDate)
			FwsMonth = Month(pcFedExShipDate)
			if len(FwsMonth)=1 Then
				FwsMonth = "0"&FwsMonth
			end if
			FwsDay = Day(pcFedExShipDate)
			if len(FwsDay)=1 Then
				FwsDay = "0"&FwsDay
			end if

			pcFedExFormatedShipDate = FwsYear&"-"&FwsMonth&"-"&FwsDay&"T14:33:57+05:30"

			objFedExWSClass.AddNewNode "ShipTimestamp", pcv_strVersion, pcFedExFormatedShipDate

			objFedExWSClass.AddNewNode "ServiceType", pcv_strVersion, ""
			objFedExWSClass.AddNewNode "DropoffType", pcv_strVersion, FEDEXWS_DROPOFF_TYPE
			objFedExWSClass.AddNewNode "PackagingType", pcv_strVersion, FEDEXWS_FEDEX_PACKAGE
			objFedExWSClass.WriteParent "Shipper", pcv_strVersion, ""
				objFedExWSClass.WriteParent "Contact", pcv_strVersion, ""
					objFedExWSClass.AddNewNode "PersonName", pcv_strVersion, scOriginPersonName
					objFedExWSClass.AddNewNode "PhoneNumber", pcv_strVersion, scOriginPhoneNumber
				objFedExWSClass.WriteParent "Contact", pcv_strVersion, "/"

				objFedExWSClass.WriteParent "Address", pcv_strVersion, ""
					objFedExWSClass.AddNewNode "StreetLines", pcv_strVersion, scShipFromAddress1
					objFedExWSClass.AddNewNode "City", pcv_strVersion, scShipFromCity
					objFedExWSClass.AddNewNode "StateOrProvinceCode", pcv_strVersion, scShipFromState
					objFedExWSClass.AddNewNode "PostalCode", pcv_strVersion, scShipFromPostalCode
					objFedExWSClass.AddNewNode "CountryCode", pcv_strVersion, scShipFromPostalCountry
				objFedExWSClass.WriteParent "Address", pcv_strVersion, "/"


			objFedExWSClass.WriteParent "Shipper", pcv_strVersion, "/"

			objFedExWSClass.WriteParent "Recipient", pcv_strVersion, ""
				objFedExWSClass.WriteParent "Contact", pcv_strVersion, ""
					objFedExWSClass.AddNewNode "PersonName", pcv_strVersion, "RECIPIENT"
				objFedExWSClass.WriteParent "Contact", pcv_strVersion, "/"

				objFedExWSClass.WriteParent "Address", pcv_strVersion, ""
					'objFedExWSClass.AddNewNode "StreetLines", pcv_strVersion, ""
					'objFedExWSClass.AddNewNode "City", pcv_strVersion, Universal_destination_city
					if Universal_destination_country="US" OR Universal_destination_country="CA" then
						objFedExWSClass.AddNewNode "StateOrProvinceCode", pcv_strVersion, Universal_destination_provOrState
					end if
					objFedExWSClass.AddNewNode "PostalCode", pcv_strVersion, Universal_destination_postal
					objFedExWSClass.AddNewNode "CountryCode", pcv_strVersion, Universal_destination_country
					if pResidentialShipping="-1" or pResidentialShipping="1" then
						objFedExWSClass.AddNewNode "Residential", pcv_strVersion, "true"
					else
						objFedExWSClass.AddNewNode "Residential", pcv_strVersion, "false"
					end if
				objFedExWSClass.WriteParent "Address", pcv_strVersion, "/"

			objFedExWSClass.WriteParent "Recipient", pcv_strVersion, "/"

			objFedExWSClass.WriteParent "ShippingChargesPayment", pcv_strVersion, ""
				objFedExWSClass.AddNewNode "PaymentType", pcv_strVersion, "SENDER"
			objFedExWSClass.WriteParent "ShippingChargesPayment", pcv_strVersion, "/"

			'// ShipDate
			'objFedExWSClass.AddNewNode "ShipDate", pcv_strVersion, "2011-08-03T14:33:57+05:30"
			'objFedExWSClass.WriteParent "ShipDate", pcv_strVersion, ""
			'	objFedExWSClass.AddNewNode "ShipDate", pcv_strVersion, "2011-08-03"
			'objFedExWSClass.WriteParent "ShipDate", pcv_strVersion, "/"

			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// Start: FDXE Special Services
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			If pcv_TmpSaturdayDelivery<>"0" Then
				objFedExWSClass.WriteParent "SpecialServicesRequested", pcv_strVersion, ""
					'// Saturday Delivery
					objFedExWSClass.AddNewNode "SpecialServiceTypes", pcv_strVersion, "SATURDAY_DELIVERY"
					'// Saturday Pickup
					'objFedExWSClass.AddNewNode "SpecialServiceTypes", pcv_strVersion, "SATURDAY_PICKUP"
				objFedExWSClass.WriteParent "SpecialServicesRequested", pcv_strVersion, "/"
			End If
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// End: FDXE Special Services
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// Start: FDXE International
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'objFedExWSClass.WriteParent "CustomsClearanceDetail", pcv_strVersion, ""

			'// Custom Value
			'objFedExWSClass.WriteParent "CustomsValue", pcv_strVersion, ""
			'		objFedExWSClass.AddNewNode "Currency", pcv_strVersion, "USD"
			'		objFedExWSClass.AddNewNode "Amount", pcv_strVersion, "250"
			'objFedExWSClass.WriteParent "CustomsValue", pcv_strVersion, "/"

			'objFedExWSClass.WriteParent "CustomsClearanceDetail", pcv_strVersion, "/"
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// End: FDXE International
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


			if pcv_TmpListRate = "-1" OR pcv_TmpListRate = -1 then
				pcv_strRateRequestType = "LIST"
			else
				pcv_strRateRequestType = "ACCOUNT"
			end if
			objFedExWSClass.AddNewNode "RateRequestTypes", pcv_strVersion, pcv_strRateRequestType
			objFedExWSClass.AddNewNode "PackageCount", pcv_strVersion, pcv_intPackageNum
			objFedExWSClass.AddNewNode "PackageDetail", pcv_strVersion, "INDIVIDUAL_PACKAGES"

			for q=1 to pcv_intPackageNum
				'//FEDEXWSWEIGHTCHANGE///////////////////////////////////////
				session("FedEXWSPackWeight"&q)
				tmpPounds = session("FEDEXWSPackPounds"&q)
				tmpOunces = session("FEDEXWSPackOunces"&q)
				if NOT isNumeric(tmpPounds) then
					tmpPounds = 0
				end if
				if NOT isNumeric(tmpOunces) then
					tmpOunces = 0
				end if
				tmpOuncesDec = Cint(tmpOunces)/16
				objFedExWSClass.WriteParent "RequestedPackageLineItems", pcv_strVersion, ""

				objFedExWSClass.AddNewNode "SequenceNumber", pcv_strVersion, q

				pcv_TempPackPrice=session("FedEXWSPackPrice"&q)
				If pcv_TempPackPrice="" Then
					pcv_TempPackPrice="100.00"
				End If

				objFedExWSClass.WriteParent "InsuredValue", pcv_strVersion, ""
					objFedExWSClass.AddNewNode "Currency", pcv_strVersion, "USD"
					objFedExWSClass.AddNewNode "Amount", pcv_strVersion, replace(money(pcv_TempPackPrice),",","")
				objFedExWSClass.WriteParent "InsuredValue", pcv_strVersion, "/"

				objFedExWSClass.WriteParent "Weight", pcv_strVersion, ""
					If scShipFromWeightUnit="LBS" Then
						objFedExWSClass.WriteSingleParent "Units", pcv_strVersion, "LB"
					Else
						objFedExWSClass.WriteSingleParent "Units", pcv_strVersion, "KG"
					End If
					objFedExWSClass.WriteSingleParent "Value", pcv_strVersion, tmpPounds + tmpOuncesDec
				objFedExWSClass.WriteParent "Weight", pcv_strVersion, "/"

				if ((FEDEXWS_FEDEX_PACKAGE="YOUR_PACKAGING") AND (session("FedEXWSPackLength"&q)<>"" AND session("FedEXWSPackWidth"&q)<>"" AND session("FedEXWSPackHeight"&q)<>"")) then
					pcv_strDimUnit = FEDEXWS_DIM_UNIT
					if pcv_strDimUnit="" then
						pcv_strDimUnit = "IN"
					end if
					objFedExWSClass.WriteParent "Dimensions", pcv_strVersion, ""
						objFedExWSClass.AddNewNode "Length", pcv_strVersion, Int(session("FedEXWSPackLength"&q))
						objFedExWSClass.AddNewNode "Width", pcv_strVersion, Int(session("FedEXWSPackWidth"&q))
						objFedExWSClass.AddNewNode "Height", pcv_strVersion, Int(session("FedEXWSPackHeight"&q))
						objFedExWSClass.AddNewNode "Units", pcv_strVersion, FEDEXWS_DIM_UNIT
					objFedExWSClass.WriteParent "Dimensions", pcv_strVersion, "/"
				end if

				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// Start: FDXG Special Services (Package Level)
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				objFedExWSClass.WriteParent "SpecialServicesRequested", pcv_strVersion, ""


						'// NON_STANDARD_CONTAINER
						'objFedExWSClass.AddNewNode "SpecialServiceTypes", pcv_strVersion, "NON_STANDARD_CONTAINER"

						'// COD
						'objFedExWSClass.AddNewNode "SpecialServiceTypes", pcv_strVersion, "COD"
						'objFedExWSClass.WriteParent "CodDetail", pcv_strVersion, ""
							'objFedExWSClass.WriteParent "CodCollectionAmount", pcv_strVersion, ""
							'	objFedExWSClass.AddNewNode "Currency", pcv_strVersion, "USD"
							'	objFedExWSClass.AddNewNode "Amount", pcv_strVersion, "100"
							'objFedExWSClass.WriteParent "CodCollectionAmount", pcv_strVersion, "/"
						'	objFedExWSClass.AddNewNode "CollectionType", pcv_strVersion, "CASH"
						'objFedExWSClass.WriteParent "CodDetail", pcv_strVersion, "/"


						'// Signature Option!
						'objFedExWSClass.AddNewNode "SpecialServiceTypes", pcv_strVersion, "SIGNATURE_OPTION"
						'objFedExWSClass.WriteParent "SignatureOptionDetail", pcv_strVersion, ""
						'	objFedExWSClass.AddNewNode "OptionType", pcv_strVersion, "NO_SIGNATURE_REQUIRED" '// "DIRECT" '// "INDIRECT" '// "ADULT"
						'objFedExWSClass.WriteParent "SignatureOptionDetail", pcv_strVersion, "/"


					objFedExWSClass.WriteParent "SpecialServicesRequested", pcv_strVersion, "/"
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					'// End: FDXG Special Services (Package Level)
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

					'if pResidentialShipping="-1" or pResidentialShipping="1" then
					'	objFedExWSClass.WriteParent "SpecialServicesRequested", pcv_strVersion, ""
					'		objFedExWSClass.AddNewNode "SpecialServiceTypes", pcv_strVersion, ""
					'	objFedExWSClass.WriteParent "SpecialServicesRequested", pcv_strVersion, "/"
					'end if

				objFedExWSClass.WriteParent "RequestedPackageLineItems", pcv_strVersion, "/"

			next

		objFedExWSClass.WriteParent "RequestedShipment", pcv_strVersion, "/"

	objFedExWSClass.EndXMLTransaction pcv_strMethodNameWS, pcv_strVersion


	'// Print out our newly formed request xml
	'response.Clear()
	'response.ContentType="text/xml"
	'response.Write(fedex_postdataWS)
	'response.End()

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Send Our Transaction.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	call objFedExWSClass.SendXMLRequest(fedex_postdataWS, pcv_strEnvironment)
	'// Print out our response
	'response.Clear()
	'response.ContentType="text/xml"
	'response.Write(FEDEXWS_result)
	'response.End()

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Load Our Response.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	call objFedExWSClass.LoadXMLResults(FEDEXWS_result)

	pcv_strErrorMsgWS = cSTR("")

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Check for errors from FedEx.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'call objFedExWSClass.XMLResponseVerify(ErrPageName)
	pcv_strErrorMsgWS = objFedExWSClass.ReadResponseNode("//v9:Notifications", "v9:Severity")
	if pcv_strErrorMsgWS<>"SUCCESS" AND  pcv_strErrorMsgWS<>"NOTE" AND pcv_strErrorMsgWS<>"WARNING"  then
		pcv_strErrorMsgWS = objFedExWSClass.ReadResponseNode("//v9:Notifications", "v9:Message")
	else
		pcv_strErrorMsgWS = ""
	end if

	'/////////////////////////////////////////////////////////////
	'// BASELINE LOGGING
	'/////////////////////////////////////////////////////////////
	'// Log our Transaction
	'call objFedExWSClass.pcs_LogTransaction(fedex_postdataWS, pcv_strMethodNameWS&"_G"&q&".in", true)
	'// Log our Response
	'call objFedExWSClass.pcs_LogTransaction(FEDEXWS_result, pcv_strMethodNameWS&"_G"&q&".out", true)

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Redirect with a Message OR complete some task.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



	If NOT len(pcv_strErrorMsgWS)>0 Then
		'// Generate FedEx Arrays
		if pcv_TmpListRate = "-1" OR pcv_TmpListRate = -1 then
			arryFedExWSService = arryFedExWSService & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:ServiceType")
			if scHideEstimateDeliveryTimes="1" then
				arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & ""
				arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & ""
			else
				arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & "TransitTime:"&objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:TransitTime")
				arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & ""
			end if
			arryFedExWSRate = arryFedExWSRate & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:RatedShipmentDetails/v9:ShipmentRateDetail/v9:TotalNetFedExCharge/v9:Amount")
			' Parse the XML document.
			Set objOutputXMLDocss = Server.CreateObject("Microsoft.XMLDOM")
			objOutputXMLDocss.loadXML FEDEXWS_result
			Set Nodes = objOutputXMLDocss.selectNodes("//RateReplyDetails")
			set objLst=objOutputXMLDocss.getElementsByTagName("v9:RateReplyDetails")
			tempGetRate = 0
			for i = 0 to (objLst.length - 1)
				for j=0 to ((objLst.item(i).childNodes.length)-1)
					If objLst.item(i).childNodes(j).nodeName="v9:RatedShipmentDetails" then
						for m=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="v9:ShipmentRateDetail" then
								for n=0 to ((objLst.item(i).childNodes(j).childNodes(m).childNodes.length)-1)
									If objLst.item(i).childNodes(j).childNodes(m).childNodes(n).nodeName="v9:RateType" then
										if objLst.item(i).childNodes(j).childNodes(m).childNodes(n).text = "PAYOR_LIST_PACKAGE" Then
											tempGetRate = 1
										end if
									End If
									If tempGetRate = 1 AND objLst.item(i).childNodes(j).childNodes(m).childNodes(n).nodeName="v9:TotalNetFedExCharge" Then
										for p=0 to ((objLst.item(i).childNodes(j).childNodes(m).childNodes(n).childNodes.length)-1)
											if objLst.item(i).childNodes(j).childNodes(m).childNodes(n).childNodes(p).nodeName = "v9:Amount" Then
												arryFedExWSRate2 = arryFedExWSRate2 & objLst.item(i).childNodes(j).childNodes(m).childNodes(n).childNodes(p).text &","
												tempGetRate = 0
											end if
										next
									End If
								next
							end if
						next
					End if
				next
			next
			arryFedExWSRate = arryFedExWSRate2
		else
			arryFedExWSService = arryFedExWSService & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:ServiceType")
			if scHideEstimateDeliveryTimes="1" then
				arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & ""
				arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & ""
			else
				arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & "TransitTime:"&objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:TransitTime")
				arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & ""
			end if
			arryFedExWSRate = arryFedExWSRate & objFedExWSClass.ReadResponseasArray("//v9:RateReplyDetails", "v9:RatedShipmentDetails/v9:ShipmentRateDetail/v9:TotalNetFedExCharge/v9:Amount")
		end if
	End If

	set objFedExWSClass = nothing
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'// END: FedEx Frieght
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////



	' trim the last comma if there is one
	'xStringLength = len(ReadResponseasArray)
	'if xStringLength>0 then
	'	ReadResponseasArray = left(ReadResponseasArray,(xStringLength-1))
	'end if

	'Split Arrays
	dim intRateIndexWS
	dim pcFedExWSMultiArry(15,5)
	for z=0 to 15
		pcFedExWSMultiArry(z,1)=0
	next

	pcStrTempFedExService=split(arryFedExWSService,",")
	pcStrTempFexExRate=split(arryFedExWSRate,",")
	pcStrTempFedExDeliveryDate=split(arrFedExWSDeliveryDate,",")
	pcStrTempFedExDeliveryTime=split(arrFedExWSDeliveryTime,",")

	'EUROPE_FIRST_INTERNATIONAL_PRIORITY
	'FEDEX_FREIGHT
	'FEDEX_NATIONAL_FREIGHT
	'INTERNATIONAL_DISTRIBUTION_FREIGHT

	for t=0 to (ubound(pcStrTempFedExService)-1)
		select case pcStrTempFedExService(t)
			case "PRIORITY_OVERNIGHT"
				intRateIndexWS=0
				pcFedExWSMultiArry(intRateIndexWS,2)="FedEx Priority Overnight<sup>&reg;</sup>"
				pcFedExWSMultiArry(intRateIndexWS,3)="PRIORITY_OVERNIGHT"
				pcFedExWSMultiArry(intRateIndexWS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSMultiArry(intRateIndexWS,5)=pcStrTempFedExDeliveryTime(t)
			case "STANDARD_OVERNIGHT"
				intRateIndexWS=1
				pcFedExWSMultiArry(intRateIndexWS,2)="FedEx Standard Overnight<sup>&reg;</sup>"
				pcFedExWSMultiArry(intRateIndexWS,3)="STANDARD_OVERNIGHT"
				pcFedExWSMultiArry(intRateIndexWS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSMultiArry(intRateIndexWS,5)=pcStrTempFedExDeliveryTime(t)
			case "FIRST_OVERNIGHT"
				intRateIndexWS=2
				pcFedExWSMultiArry(intRateIndexWS,2)="FedEx First Overnight<sup>&reg;</sup>"
				pcFedExWSMultiArry(intRateIndexWS,3)="FIRST_OVERNIGHT"
				pcFedExWSMultiArry(intRateIndexWS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSMultiArry(intRateIndexWS,5)=pcStrTempFedExDeliveryTime(t)
			case "FEDEX_2_DAY"
				intRateIndexWS=3
				pcFedExWSMultiArry(intRateIndexWS,2)="FedEx 2Day<sup>&reg;</sup>"
				pcFedExWSMultiArry(intRateIndexWS,3)="FEDEX_2_DAY"
				pcFedExWSMultiArry(intRateIndexWS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSMultiArry(intRateIndexWS,5)=pcStrTempFedExDeliveryTime(t)
			case "FEDEX_EXPRESS_SAVER"
				intRateIndexWS=4
				pcFedExWSMultiArry(intRateIndexWS,2)="FedEx Express Saver<sup>&reg;</sup>"
				pcFedExWSMultiArry(intRateIndexWS,3)="FEDEX_EXPRESS_SAVER"
				pcFedExWSMultiArry(intRateIndexWS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSMultiArry(intRateIndexWS,5)=pcStrTempFedExDeliveryTime(t)
			case "INTERNATIONAL_PRIORITY"
				intRateIndexWS=5
				pcFedExWSMultiArry(intRateIndexWS,2)="FedEx International Priority<sup>&reg;</sup>"
				pcFedExWSMultiArry(intRateIndexWS,3)="INTERNATIONAL_PRIORITY"
				pcFedExWSMultiArry(intRateIndexWS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSMultiArry(intRateIndexWS,5)=pcStrTempFedExDeliveryTime(t)
			case "INTERNATIONAL_ECONOMY"
				intRateIndexWS=6
				pcFedExWSMultiArry(intRateIndexWS,2)="FedEx International Economy<sup>&reg;</sup>"
				pcFedExWSMultiArry(intRateIndexWS,3)="INTERNATIONAL_ECONOMY"
				pcFedExWSMultiArry(intRateIndexWS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSMultiArry(intRateIndexWS,5)=pcStrTempFedExDeliveryTime(t)
			case "INTERNATIONAL_FIRST"
				intRateIndexWS=7
				pcFedExWSMultiArry(intRateIndexWS,2)="FedEx International First<sup>&reg;</sup>"
				pcFedExWSMultiArry(intRateIndexWS,3)="INTERNATIONAL_FIRST"
				pcFedExWSMultiArry(intRateIndexWS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSMultiArry(intRateIndexWS,5)=pcStrTempFedExDeliveryTime(t)
			case "FEDEX_1_DAY_FREIGHT"
				intRateIndexWS=8
				pcFedExWSMultiArry(intRateIndexWS,2)="FedEx 1Day<sup>&reg;</sup> Freight"
				pcFedExWSMultiArry(intRateIndexWS,3)="FEDEX_1_DAY_FREIGHT"
				pcFedExWSMultiArry(intRateIndexWS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSMultiArry(intRateIndexWS,5)=pcStrTempFedExDeliveryTime(t)
			case "FEDEX_2_DAY_FREIGHT"
				intRateIndexWS=9
				pcFedExWSMultiArry(intRateIndexWS,2)="FedEx 2Day<sup>&reg;</sup> Freight"
				pcFedExWSMultiArry(intRateIndexWS,3)="FEDEX_2_DAY_FREIGHT"
				pcFedExWSMultiArry(intRateIndexWS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSMultiArry(intRateIndexWS,5)=pcStrTempFedExDeliveryTime(t)
			case "FEDEX_3_DAY_FREIGHT"
				intRateIndexWS=10
				pcFedExWSMultiArry(intRateIndexWS,2)="FedEx 3Day<sup>&reg;</sup> Freight"
				pcFedExWSMultiArry(intRateIndexWS,3)="FEDEX_3_DAY_FREIGHT"
				pcFedExWSMultiArry(intRateIndexWS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSMultiArry(intRateIndexWS,5)=pcStrTempFedExDeliveryTime(t)
			case "FEDEX_GROUND"
				If Universal_destination_country="US" Then
					intRateIndexWS=11
					pcFedExWSMultiArry(intRateIndexWS,2)="FedEx Ground<sup>&reg;</sup>"
					pcFedExWSMultiArry(intRateIndexWS,3)="FEDEX_GROUND"
					pcFedExWSMultiArry(intRateIndexWS,4)=pcStrTempFedExDeliveryDate(t)
					pcFedExWSMultiArry(intRateIndexWS,5)=pcStrTempFedExDeliveryTime(t)
				End If
			case "GROUND_HOME_DELIVERY"
				If Universal_destination_country="US" Then
					intRateIndexWS=12
					pcFedExWSMultiArry(intRateIndexWS,2)="FedEx Home Delivery<sup>&reg;</sup>"
					pcFedExWSMultiArry(intRateIndexWS,3)="GROUND_HOME_DELIVERY"
					pcFedExWSMultiArry(intRateIndexWS,4)=pcStrTempFedExDeliveryDate(t)
					pcFedExWSMultiArry(intRateIndexWS,5)=pcStrTempFedExDeliveryTime(t)
				End If
			case "INTERNATIONAL_PRIORITY_FREIGHT"
				intRateIndexWS=13
				pcFedExWSMultiArry(intRateIndexWS,2)="FedEx International Priority<sup>&reg;</sup> Freight"
				pcFedExWSMultiArry(intRateIndexWS,3)="INTERNATIONAL_PRIORITY_FREIGHT"
				pcFedExWSMultiArry(intRateIndexWS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSMultiArry(intRateIndexWS,5)=pcStrTempFedExDeliveryTime(t)
			case "INTERNATIONAL_ECONOMY_FREIGHT"
				intRateIndexWS=14
				pcFedExWSMultiArry(intRateIndexWS,2)="FedEx International Economy<sup>&reg;</sup> Freight"
				pcFedExWSMultiArry(intRateIndexWS,3)="INTERNATIONAL_ECONOMY_FREIGHT"
				pcFedExWSMultiArry(intRateIndexWS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSMultiArry(intRateIndexWS,5)=pcStrTempFedExDeliveryTime(t)
			case "SMART_POST"
				intRateIndexWS=15
				pcFedExWSMultiArry(intRateIndexWS,2)="FedEx SmartPost<sup>&reg;</sup>"
				pcFedExWSMultiArry(intRateIndexWS,3)="SMART_POST"
				pcFedExWSMultiArry(intRateIndexWS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSMultiArry(intRateIndexWS,5)=pcStrTempFedExDeliveryTime(t)
		end select
		tempRate=pcFedExWSMultiArry(intRateIndexWS,1)
		pcFedExWSMultiArry(intRateIndexWS,1)=cdbl(tempRate)+cdbl(pcStrTempFexExRate(t))
	next

	for z=0 to 15
		if pcFedExWSMultiArry(z,1)>0 then
			intNoTime = 0
			pcv_strFormattedDate = ""

			tmpDeliveryDayOfWeek = pcFedExWSMultiArry(z,4)

			If instr(tmpDeliveryDayOfWeek,"TransitTime") Then
				If instr(tmpDeliveryDayOfWeek,"SPTransitTime") Then
					tmpDeliveryDayOfWeek = replace(tmpDeliveryDayOfWeek, "SPTransitTime:","")
					intNoTime = 2
					select case tmpDeliveryDayOfWeek
						case "EIGHTEEN_DAYS"
							tmpDeliveryDayOfWeek = "18"
						case "EIGHT_DAYS"
							tmpDeliveryDayOfWeek = "8"
						case "ELEVEN_DAYS"
							tmpDeliveryDayOfWeek = "11"
						case "FIFTEEN_DAYS"
							tmpDeliveryDayOfWeek = "15"
						case "FIVE_DAYS"
							tmpDeliveryDayOfWeek = "5"
						case "FOURTEEN_DAYS"
							tmpDeliveryDayOfWeek = "14"
						case "FOUR_DAYS"
							tmpDeliveryDayOfWeek = "4"
						case "NINETEEN_DAYS"
							tmpDeliveryDayOfWeek = "19"
						case "NINE_DAYS"
							tmpDeliveryDayOfWeek = "9"
						case "ONE_DAY"
							tmpDeliveryDayOfWeek = "1"
						case "SEVENTEEN_DAYS"
							tmpDeliveryDayOfWeek = "17"
						case "SEVEN_DAYS"
							tmpDeliveryDayOfWeek = "7"
						case "SIXTEEN_DAYS"
							tmpDeliveryDayOfWeek = "16"
						case "SIX_DAYS"
							tmpDeliveryDayOfWeek = "6"
						case "TEN_DAYS"
							tmpDeliveryDayOfWeek = "10"
						case "THIRTEEN_DAYS"
							tmpDeliveryDayOfWeek = "13"
						case "THREE_DAYS"
							tmpDeliveryDayOfWeek = "3"
						case "TWELVE_DAYS"
							tmpDeliveryDayOfWeek = "12"
						case "TWENTY_DAYS"
							tmpDeliveryDayOfWeek = "20"
						case "TWO_DAYS"
							tmpDeliveryDayOfWeek = "2"
						case else
							tmpDeliveryDayOfWeek = "Unknown"
					end select
					tmpDeliveryBetween = pcFedExWSMultiArry(z,5)
					tmpDeliveryBetween = replace(tmpDeliveryBetween, "SPTransitTime:","")
					select case tmpDeliveryBetween
						case "EIGHTEEN_DAYS"
							tmpDeliveryBetween = "18"
						case "EIGHT_DAYS"
							tmpDeliveryBetween = "8"
						case "ELEVEN_DAYS"
							tmpDeliveryBetween = "11"
						case "FIFTEEN_DAYS"
							tmpDeliveryBetween = "15"
						case "FIVE_DAYS"
							tmpDeliveryBetween = "5"
						case "FOURTEEN_DAYS"
							tmpDeliveryBetween = "14"
						case "FOUR_DAYS"
							tmpDeliveryBetween = "4"
						case "NINETEEN_DAYS"
							tmpDeliveryBetween = "19"
						case "NINE_DAYS"
							tmpDeliveryBetween = "9"
						case "ONE_DAY"
							tmpDeliveryBetween = "1"
						case "SEVENTEEN_DAYS"
							tmpDeliveryBetween = "17"
						case "SEVEN_DAYS"
							tmpDeliveryBetween = "7"
						case "SIXTEEN_DAYS"
							tmpDeliveryBetween = "16"
						case "SIX_DAYS"
							tmpDeliveryBetween = "6"
						case "TEN_DAYS"
							tmpDeliveryBetween = "10"
						case "THIRTEEN_DAYS"
							tmpDeliveryBetween = "13"
						case "THREE_DAYS"
							tmpDeliveryBetween = "3"
						case "TWELVE_DAYS"
							tmpDeliveryBetween = "12"
						case "TWENTY_DAYS"
							tmpDeliveryBetween = "20"
						case "TWO_DAYS"
							tmpDeliveryBetween = "2"
						case else
							tmpDeliveryBetween = "Unknown"
					end select
				Else
				tmpDeliveryDayOfWeek = replace(tmpDeliveryDayOfWeek, "TransitTime:","")
				intNoTime = 1
				select case tmpDeliveryDayOfWeek
					case "EIGHTEEN_DAYS"
						tmpDeliveryDayOfWeek = "18"
					case "EIGHT_DAYS"
						tmpDeliveryDayOfWeek = "8"
					case "ELEVEN_DAYS"
						tmpDeliveryDayOfWeek = "11"
					case "FIFTEEN_DAYS"
						tmpDeliveryDayOfWeek = "15"
					case "FIVE_DAYS"
						tmpDeliveryDayOfWeek = "5"
					case "FOURTEEN_DAYS"
						tmpDeliveryDayOfWeek = "14"
					case "FOUR_DAYS"
						tmpDeliveryDayOfWeek = "4"
					case "NINETEEN_DAYS"
						tmpDeliveryDayOfWeek = "19"
					case "NINE_DAYS"
						tmpDeliveryDayOfWeek = "9"
					case "ONE_DAY"
						tmpDeliveryDayOfWeek = "1"
					case "SEVENTEEN_DAYS"
						tmpDeliveryDayOfWeek = "17"
					case "SEVEN_DAYS"
						tmpDeliveryDayOfWeek = "7"
					case "SIXTEEN_DAYS"
						tmpDeliveryDayOfWeek = "16"
					case "SIX_DAYS"
						tmpDeliveryDayOfWeek = "6"
					case "TEN_DAYS"
						tmpDeliveryDayOfWeek = "10"
					case "THIRTEEN_DAYS"
						tmpDeliveryDayOfWeek = "13"
					case "THREE_DAYS"
						tmpDeliveryDayOfWeek = "3"
					case "TWELVE_DAYS"
						tmpDeliveryDayOfWeek = "12"
					case "TWENTY_DAYS"
						tmpDeliveryDayOfWeek = "20"
					case "TWO_DAYS"
						tmpDeliveryDayOfWeek = "2"
					case else
						tmpDeliveryDayOfWeek = "Unknown"
				end select
				End If
			End If

			If intNoTime = 0 Then
				select case tmpDeliveryDayOfWeek
					case "MON"
						tmpDeliveryDayOfWeek = "Monday"
					case "TUE"
						tmpDeliveryDayOfWeek = "Tuesday"
					case "WED"
						tmpDeliveryDayOfWeek = "Wednesday"
					case "THU"
						tmpDeliveryDayOfWeek = "Thursday"
					case "FRI"
						tmpDeliveryDayOfWeek = "Friday"
					case "SAT"
						tmpDeliveryDayOfWeek = "Saturday"
				end select

				tmpDeliveryTimestamp = pcFedExWSMultiArry(z,5)
				arrDeliveryTimestamp = split(tmpDeliveryTimestamp, "T")
				tmpDeliveryTime = arrDeliveryTimestamp(1)
				arrTimeFormat = split(tmpDeliveryTime,":")
				tmpTimeHour = Cint(arrTimeFormat(0))
				tmpTimeMinutes = arrTimeFormat(1)
				tmpTimeSeconds = arrTimeFormat(2)
				'//Format hour and check for AM/PM
				if tmpTimeHour < 12 then
					tmpAMPM = "AM"
					tmpHour = Cint(tmpTimeHour)
				else
					tmpAMPM = "PM"
					tmpHour = Cint(tmpTimeHour) - Cint(12)
				end if
				tmpDeliveryDate = arrDeliveryTimestamp(0)
				arrDeliveryDate = split(tmpDeliveryDate,"-")
				tmpDeliveryDay = arrDeliveryDate(2)
				tmpDeliveryMonth = arrDeliveryDate(1)
				select case tmpDeliveryMonth
					case "01"
						tmpDeliveryMonth = "January"
					case "02"
						tmpDeliveryMonth = "February"
					case "03"
						tmpDeliveryMonth = "March"
					case "04"
						tmpDeliveryMonth = "April"
					case "05"
						tmpDeliveryMonth = "May"
					case "06"
						tmpDeliveryMonth = "June"
					case "07"
						tmpDeliveryMonth = "July"
					case "08"
						tmpDeliveryMonth = "August"
					case "09"
						tmpDeliveryMonth = "September"
					case "10"
						tmpDeliveryMonth = "October"
					case "11"
						tmpDeliveryMonth = "November"
					case "12"
						tmpDeliveryMonth = "December"
				end select
				tmpDeliveryYear = arrDeliveryDate(0)
			End If
			If intNoTime = 1 or intNoTime = 2 Then
				if tmpDeliveryDayOfWeek = "Unknown" then
					pcv_strFormattedDate = tmpDeliveryDayOfWeek
				else
					pcAddDay = 	FEDEXWS_ADDDAY
					if pcAddDay&""="" then
						pcAddDay = 0
					end if
					DatePlus = Date()+Cint(pcAddDay)
					pcv_strFormattedDate = FormatDateTime(DateAdd("D", Cint(tmpDeliveryDayOfWeek), DatePlus), 1)&" 7:00 PM"
					if intNoTime = 2 then
						tmpFromSPDate1 = FormatDateTime(DateAdd("D", Cint(tmpDeliveryDayOfWeek), DatePlus), 1)
						if instr(tmpFromSPDate1, "Sunday") then
							tmpDeliveryDayOfWeek = int(tmpDeliveryDayOfWeek)+1
							tmpFromSPDate1 = FormatDateTime(DateAdd("D", Cint(tmpDeliveryDayOfWeek), DatePlus), 1)
						end if
						tmpToSPDate1 = FormatDateTime(DateAdd("D", Cint(tmpDeliveryBetween), DatePlus), 1)
						
						if instr(tmpToSPDate1, "Sunday") then
							tmpDeliveryBetween = int(tmpDeliveryBetween)+1
							tmpToSPDate1 = FormatDateTime(DateAdd("D", Cint(tmpDeliveryBetween), DatePlus), 1)
						end if
					
						pcv_strFormattedDate = "Between: "&tmpFromSPDate1&" 7:00 PM AND "&tmpToSPDate1&" 7:00 PM"
					end if
				end if
			Else
				pcv_strFormattedDate = tmpDeliveryDayOfWeek&", "&tmpDeliveryMonth&" "&tmpDeliveryDay&", "&tmpDeliveryYear&" "&tmpHour&":"&tmpTimeMinutes&" "&tmpAMPM
			End If

			availableShipStr=availableShipStr&"|?|FedExWS|"&pcFedExWSMultiArry(z,3)&"|"&pcFedExWSMultiArry(z,2)&"|"&pcFedExWSMultiArry(z,1)&"|"&pcv_strFormattedDate
			iFedExWSFlag=1

		end if
	next
end if 'if fedex is active
%>
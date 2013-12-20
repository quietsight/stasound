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
call closedb()

if (FedEXWS_active=true or FedExWS_active="-1") AND FedEXWS_AccountNumber<>"" then

	iFedExWSActive=1
	arryFedExWSService=""
	arryFedExWSRate=""
	arrFedExWSDeliveryDate=""
	arryFedExWSRate2 = ""

	pcv_TmpListRate = FEDEXWS_LISTRATE
	pcv_TmpSaturdayDelivery = FEDEXWS_SATURDAYDELIVERY
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
			pcAddDay = 	FEDEXWS_ADDDAY
			if pcAddDay&""="" then
				pcAddDay = 0
			end if

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
						objFedExWSClass.AddNewNode "Residential", pcv_strVersion, "1"
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
				pcv_strRateRequestType = "MULTIWEIGHT"
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

	'Split Arrays
	dim intRateIndexWSS
	dim pcFedExWSSMultiArry(15,5)
	for z=0 to 15
		pcFedExWSSMultiArry(z,1)=0
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
				intRateIndexWSS=0
				pcFedExWSSMultiArry(intRateIndexWSS,2)="FedEx Priority Overnight<sup>&reg;</sup> (Saturday Delivery)"
				pcFedExWSSMultiArry(intRateIndexWSS,3)="PRIORITY_OVERNIGHT"
				pcFedExWSSMultiArry(intRateIndexWSS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSSMultiArry(intRateIndexWSS,5)=pcStrTempFedExDeliveryTime(t)
			case "STANDARD_OVERNIGHT"
				intRateIndexWSS=1
				pcFedExWSSMultiArry(intRateIndexWSS,2)="FedEx Standard Overnight<sup>&reg;</sup> (Saturday Delivery)"
				pcFedExWSSMultiArry(intRateIndexWSS,3)="STANDARD_OVERNIGHT"
				pcFedExWSSMultiArry(intRateIndexWSS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSSMultiArry(intRateIndexWSS,5)=pcStrTempFedExDeliveryTime(t)
			case "FIRST_OVERNIGHT"
				intRateIndexWSS=2
				pcFedExWSSMultiArry(intRateIndexWSS,2)="FedEx First Overnight<sup>&reg;</sup> (Saturday Delivery)"
				pcFedExWSSMultiArry(intRateIndexWSS,3)="FIRST_OVERNIGHT"
				pcFedExWSSMultiArry(intRateIndexWSS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSSMultiArry(intRateIndexWSS,5)=pcStrTempFedExDeliveryTime(t)
			case "FEDEX_2_DAY"
				intRateIndexWSS=3
				pcFedExWSSMultiArry(intRateIndexWSS,2)="FedEx 2Day<sup>&reg;</sup> (Saturday Delivery)"
				pcFedExWSSMultiArry(intRateIndexWSS,3)="FEDEX_2_DAY"
				pcFedExWSSMultiArry(intRateIndexWSS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSSMultiArry(intRateIndexWSS,5)=pcStrTempFedExDeliveryTime(t)
			case "FEDEX_EXPRESS_SAVER"
				intRateIndexWSS=4
				pcFedExWSSMultiArry(intRateIndexWSS,2)="FedEx Express Saver<sup>&reg;</sup> (Saturday Delivery)"
				pcFedExWSSMultiArry(intRateIndexWSS,3)="FEDEX_EXPRESS_SAVER"
				pcFedExWSSMultiArry(intRateIndexWSS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSSMultiArry(intRateIndexWSS,5)=pcStrTempFedExDeliveryTime(t)
			case "INTERNATIONAL_PRIORITY"
				intRateIndexWSS=5
				pcFedExWSSMultiArry(intRateIndexWSS,2)="FedEx International Priority<sup>&reg;</sup> (Saturday Delivery)"
				pcFedExWSSMultiArry(intRateIndexWSS,3)="INTERNATIONAL_PRIORITY"
				pcFedExWSSMultiArry(intRateIndexWSS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSSMultiArry(intRateIndexWSS,5)=pcStrTempFedExDeliveryTime(t)
			case "INTERNATIONAL_ECONOMY"
				intRateIndexWSS=6
				pcFedExWSSMultiArry(intRateIndexWSS,2)="FedEx International Economy<sup>&reg;</sup> (Saturday Delivery)"
				pcFedExWSSMultiArry(intRateIndexWSS,3)="INTERNATIONAL_ECONOMY"
				pcFedExWSSMultiArry(intRateIndexWSS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSSMultiArry(intRateIndexWSS,5)=pcStrTempFedExDeliveryTime(t)
			case "INTERNATIONAL_FIRST"
				intRateIndexWSS=7
				pcFedExWSSMultiArry(intRateIndexWSS,2)="FedEx International First<sup>&reg;</sup> (Saturday Delivery)"
				pcFedExWSSMultiArry(intRateIndexWSS,3)="INTERNATIONAL_FIRST"
				pcFedExWSSMultiArry(intRateIndexWSS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSSMultiArry(intRateIndexWSS,5)=pcStrTempFedExDeliveryTime(t)
			case "FEDEX_1_DAY_FREIGHT"
				intRateIndexWSS=8
				pcFedExWSSMultiArry(intRateIndexWSS,2)="FedEx 1Day<sup>&reg;</sup> Freight (Saturday Delivery)"
				pcFedExWSSMultiArry(intRateIndexWSS,3)="FEDEX_1_DAY_FREIGHT"
				pcFedExWSSMultiArry(intRateIndexWSS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSSMultiArry(intRateIndexWSS,5)=pcStrTempFedExDeliveryTime(t)
			case "FEDEX_2_DAY_FREIGHT"
				intRateIndexWSS=9
				pcFedExWSSMultiArry(intRateIndexWSS,2)="FedEx 2Day<sup>&reg;</sup> Freight (Saturday Delivery)"
				pcFedExWSSMultiArry(intRateIndexWSS,3)="FEDEX_2_DAY_FREIGHT"
				pcFedExWSSMultiArry(intRateIndexWSS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSSMultiArry(intRateIndexWSS,5)=pcStrTempFedExDeliveryTime(t)
			case "FEDEX_3_DAY_FREIGHT"
				intRateIndexWSS=10
				pcFedExWSSMultiArry(intRateIndexWSS,2)="FedEx 3Day<sup>&reg;</sup> Freight (Saturday Delivery)"
				pcFedExWSSMultiArry(intRateIndexWSS,3)="FEDEX_3_DAY_FREIGHT"
				pcFedExWSSMultiArry(intRateIndexWSS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSSMultiArry(intRateIndexWSS,5)=pcStrTempFedExDeliveryTime(t)
			case "FEDEX_GROUND"
				If Universal_destination_country="US" Then
					intRateIndexWSS=11
					pcFedExWSSMultiArry(intRateIndexWSS,2)="FedEx Ground<sup>&reg;</sup> (Saturday Delivery)"
					pcFedExWSSMultiArry(intRateIndexWSS,3)="FEDEX_GROUND"
					pcFedExWSSMultiArry(intRateIndexWSS,4)=pcStrTempFedExDeliveryDate(t)
					pcFedExWSSMultiArry(intRateIndexWSS,5)=pcStrTempFedExDeliveryTime(t)
				End If
			case "GROUND_HOME_DELIVERY"
				If Universal_destination_country="US" Then
					intRateIndexWSS=12
					pcFedExWSSMultiArry(intRateIndexWSS,2)="FedEx Home Delivery<sup>&reg;</sup> (Saturday Delivery)"
					pcFedExWSSMultiArry(intRateIndexWSS,3)="GROUND_HOME_DELIVERY"
					pcFedExWSSMultiArry(intRateIndexWSS,4)=pcStrTempFedExDeliveryDate(t)
					pcFedExWSSMultiArry(intRateIndexWSS,5)=pcStrTempFedExDeliveryTime(t)
				End If
			case "INTERNATIONAL_PRIORITY_FREIGHT"
				intRateIndexWSS=13
				pcFedExWSSMultiArry(intRateIndexWSS,2)="FedEx International Priority<sup>&reg;</sup> Freight (Saturday Delivery)"
				pcFedExWSSMultiArry(intRateIndexWSS,3)="INTERNATIONAL_PRIORITY_FREIGHT"
				pcFedExWSSMultiArry(intRateIndexWSS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSSMultiArry(intRateIndexWSS,5)=pcStrTempFedExDeliveryTime(t)
			case "INTERNATIONAL_ECONOMY_FREIGHT"
				intRateIndexWSS=14
				pcFedExWSSMultiArry(intRateIndexWSS,2)="FedEx International Economy<sup>&reg;</sup> Freight (Saturday Delivery)"
				pcFedExWSSMultiArry(intRateIndexWSS,3)="INTERNATIONAL_ECONOMY_FREIGHT"
				pcFedExWSSMultiArry(intRateIndexWSS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSSMultiArry(intRateIndexWSS,5)=pcStrTempFedExDeliveryTime(t)
			case "SMART_POST"
				intRateIndexWSS=15
				pcFedExWSSMultiArry(intRateIndexWSS,2)="FedEx SmartPost<sup>&reg;</sup> (Saturday Delivery)"
				pcFedExWSSMultiArry(intRateIndexWSS,3)="SMART_POST"
				pcFedExWSSMultiArry(intRateIndexWSS,4)=pcStrTempFedExDeliveryDate(t)
				pcFedExWSSMultiArry(intRateIndexWSS,5)=pcStrTempFedExDeliveryTime(t)
		end select
		tempRate=pcFedExWSSMultiArry(intRateIndexWSS,1)
		pcFedExWSSMultiArry(intRateIndexWSS,1)=cdbl(tempRate)+cdbl(pcStrTempFexExRate(t))
	next

	for z=0 to 15
		if pcFedExWSSMultiArry(z,1)>0 then
			pcv_strFormattedDate = ""

			tmpDeliveryDayOfWeek = pcFedExWSSMultiArry(z,4)
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

			tmpDeliveryTimestamp = pcFedExWSSMultiArry(z,5)
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
				tmpHour = Cint(tmpTimeHour) - 12
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

			pcv_strFormattedDate = tmpDeliveryDayOfWeek&", "&tmpDeliveryMonth&" "&tmpDeliveryDay&", "&tmpDeliveryYear&" "&tmpTimeHour&":"&tmpTimeMinutes&" "&tmpAMPM

			availableShipStr=availableShipStr&"|?|FedExWS|"&pcFedExWSSMultiArry(z,3)&"|"&pcFedExWSSMultiArry(z,2)&"|"&pcFedExWSSMultiArry(z,1)&"|"&pcv_strFormattedDate
			iFedExWSFlag=1

		end if
	next
end if 'if fedex is active
%>
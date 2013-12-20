<%
pcv_strVersion = "9"
FedExVersion = "9"
CSPTurnOn = 1

'For live
FedExWSURL =  "https://gateway.fedex.com:443/web-services"

'/////////////////////////////////////
'// Start building the class here
'/////////////////////////////////////
Class pcFedExWSClass

	private sub Class_Initialize()
		on error resume next
		Set srvFEDEXWSXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
		Set objOutputXMLDocWS = Server.CreateObject("Microsoft.XMLDOM")
		Set objFedExStream = Server.CreateObject("ADODB.Stream")
		Set objFEDEXXmlDoc = server.createobject("Msxml2.DOMDocument"&scXML)
		objFEDEXXmlDoc.async = False
		objFEDEXXmlDoc.validateOnParse = False
		if err.number>0 then
			err.clear
		end if
	end sub



	private sub Class_Terminate()
		'// clean it all up
		Set srvFEDEXWSXmlHttp = nothing
		Set objOutputXMLDocWS = nothing
		Set objFEDEXXmlDoc = nothing
		Set objFedExStream = nothing
	end sub



	Public Sub AddNewNode(NameOfNode, FedExVersion, ValueOfNode)
		if len(ValueOfNode)>0 then
			fedex_postdataWS=fedex_postdataWS&"<v"&FedExVersion&":"&NameOfNode&">"&ValueOfNode&"</v"&FedExVersion&":"&NameOfNode&">"&vbcrlf
		end if
	End Sub



	Public Sub WriteParent(NameOfParent, FedExVersion, isClosing)
		fedex_postdataWS=fedex_postdataWS&"<"&isClosing&"v"&FedExVersion&":"&NameOfParent&">"&vbcrlf
	End Sub



	Public Sub WriteSingleParent(NameOfParent, FedExVersion, ValueOfParent)
		if len(ValueOfParent)>0 then
			fedex_postdataWS=fedex_postdataWS&"<v"&FedExVersion&":"&NameOfParent&">"&ValueOfParent&"</v"&FedExVersion&":"&NameOfParent&">"&vbcrlf
		end if
	End Sub

	'ship v9
	Public Sub AddNewNodeAlt(NameOfNode, ValueOfNode)
		if len(ValueOfNode)>0 then
			fedex_postdataWS=fedex_postdataWS&"<"&NameOfNode&">"&ValueOfNode&"</"&NameOfNode&">"&vbcrlf
		end if
	End Sub

	Public Sub WriteParentAlt(NameOfParent, isClosing)
		fedex_postdataWS=fedex_postdataWS&"<"&isClosing&NameOfParent&">"&vbcrlf
	End Sub

	Public Sub WriteSingleParentAlt(NameOfParent, ValueOfParent)
		if len(ValueOfParent)>0 then
			fedex_postdataWS=fedex_postdataWS&"<"&NameOfParent&">"&ValueOfParent&"</"&NameOfParent&">"&vbcrlf
		end if
	End Sub

	Public Sub NewXMLTransaction(NameOfMethod, FedEX_AccountNumber, FedEX_MeterNumber, FedEX_CarrierCode, CustomerTransactionIdentifier)
		fedex_postdataWS=""
		fedex_postdataWS=fedex_postdataWS&"<?xml version=""1.0"" encoding=""UTF-8"" ?>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<"&NameOfMethod&" xmlns:api=""http://www.fedex.com/fsmapi"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:noNamespaceSchemaLocation="""&NameOfMethod&".xsd"">"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<RequestHeader>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<CustomerTransactionIdentifier>"&CustomerTransactionIdentifier&"</CustomerTransactionIdentifier>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<AccountNumber>"&FedEX_AccountNumber&"</AccountNumber>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<MeterNumber>"&FedEX_MeterNumber&"</MeterNumber>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<CarrierCode>"&FedEX_CarrierCode&"</CarrierCode>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"</RequestHeader>"&vbcrlf
	End Sub



	Public Sub NewXMLCapture(NameOfMethod, FedEX_AccountNumber, FedEX_MeterNumber)
		fedex_postdataWS=""
		fedex_postdataWS=fedex_postdataWS&"<?xml version=""1.0"" encoding=""UTF-8"" ?>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<"&NameOfMethod&" xmlns:api=""http://www.fedex.com/fsmapi"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:noNamespaceSchemaLocation="""&NameOfMethod&".xsd"">"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<RequestHeader>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<AccountNumber>"&FedEX_AccountNumber&"</AccountNumber>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<MeterNumber>"&FedEX_MeterNumber&"</MeterNumber>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"</RequestHeader>"&vbcrlf
	End Sub

	'http://fedex.com/ws/rate/v9

	Public Sub NewXMLSubscription(NameOfMethod, FedEX_Key, FedEX_Password, FedExVersion, FedExName)
		fedex_postdataWS=""
		fedex_postdataWS=fedex_postdataWS&"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:v"&FedExVersion&"=""http://fedex.com/ws/"&FedExName&"/v"&FedExVersion&""">"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<soapenv:Header/>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<soapenv:Body>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<v"&FedExVersion&":"&NameOfMethod&">"&vbcrlf

		fedex_postdataWS=fedex_postdataWS&"<v"&FedExVersion&":WebAuthenticationDetail>"&vbcrlf
		If CSPTurnOn = 1 Then
			fedex_postdataWS=fedex_postdataWS&"<v"&FedExVersion&":CspCredential>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<v"&FedExVersion&":Key>CPTi545ATGa1CD89</v"&FedExVersion&":Key>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<v"&FedExVersion&":Password>8BB07q2XIIOFyNJeJQHMLv094</v"&FedExVersion&":Password>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"</v"&FedExVersion&":CspCredential>"&vbcrlf
		End If
		If len(FedEX_Key)>0 AND len(FedEX_Password)> 0 Then
			fedex_postdataWS=fedex_postdataWS&"<v"&FedExVersion&":UserCredential>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<v"&FedExVersion&":Key>"&FedEX_Key&"</v"&FedExVersion&":Key>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<v"&FedExVersion&":Password>"&FedEX_Password&"</v"&FedExVersion&":Password>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"</v"&FedExVersion&":UserCredential>"&vbcrlf
		End If

		fedex_postdataWS=fedex_postdataWS&"</v"&FedExVersion&":WebAuthenticationDetail>"&vbcrlf

	End Sub

	Public Sub NewXMLLabelWS(NameOfMethod, FedExkey, FedExPassword, FedExAccountNumber, FedExMeterNumber, FedExVersion, FedExName)
		fedex_postdataWS=""
		fedex_postdataWS=fedex_postdataWS&"<?xml version=""1.0"" encoding=""UTF-8"" ?>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:v9=""http://fedex.com/ws/ship/v9"">"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<soapenv:Body>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<v9:"&NameOfMethod&">"&vbcrlf

		fedex_postdataWS=fedex_postdataWS&"<v9:WebAuthenticationDetail>"&vbcrlf
		If CSPTurnOn = 1 Then
			fedex_postdataWS=fedex_postdataWS&"<v9:CspCredential>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"<v9:Key>CPTi545ATGa1CD89</v9:Key>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"<v9:Password>8BB07q2XIIOFyNJeJQHMLv094</v9:Password>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"</v9:CspCredential>"&vbcrlf
		End If
		fedex_postdataWS=fedex_postdataWS&"<v9:UserCredential>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<v9:Key>" & FedExkey & "</v9:Key>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<v9:Password>" & FedExPassword & "</v9:Password>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"</v9:UserCredential>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"</v9:WebAuthenticationDetail>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<v9:ClientDetail>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<v9:AccountNumber>"&FedExAccountNumber&"</v9:AccountNumber>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<v9:MeterNumber>"&FedExMeterNumber&"</v9:MeterNumber>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<v9:ClientProductId>EIPC</v9:ClientProductId>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<v9:ClientProductVersion>3424</v9:ClientProductVersion>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"</v9:ClientDetail>"&vbcrlf
	End Sub

	Public Sub NewXMLLabel(TrackingNumber, EncodedLabelString, FileType, FilePreFix)
		GraphicXML="<Base64Data xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64"" FileName="""&FilePreFix&TrackingNumber&"."&FileType&""">"&EncodedLabelString&"</Base64Data>"
	End Sub

	Public Sub SaveBinaryLabel ()
		objFedExStream.Type = 1
		objFedExStream.Open

		objFedExStream.Write objFEDEXXmlDoc.selectSingleNode("/Base64Data").nodeTypedValue
			err.clear
		strFileName = objFEDEXXmlDoc.selectSingleNode("/Base64Data/@FileName").nodeTypedValue
		'Save the binary stream to the file and overwrite if it already exists in folder
		objFedExStream.SaveToFile server.MapPath("FedExLabels\"&strFileName),2
		objFedExStream.Close()
	End Sub

	Public Sub EndXMLTransaction(NameOfMethod, FedExVersion)
		fedex_postdataWS=fedex_postdataWS&"</v"&FedExVersion&":"&NameOfMethod&">"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"</soapenv:Body>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"</soapenv:Envelope>"&vbcrlf
	End Sub

	Public Sub EndXMLTransactionAlt(NameOfMethod)
		fedex_postdataWS=fedex_postdataWS&"</"&NameOfMethod&">"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"</soapenv:Body>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"</soapenv:Envelope>"&vbcrlf
	End Sub

	Public Sub SendXMLRequest(XMLstring, Environment)
		srvFEDEXWSXmlHttp.open "POST", FedExWSURL, false
		srvFEDEXWSXmlHttp.send(XMLstring)
		FEDEXWS_result = srvFEDEXWSXmlHttp.responseText
	End Sub

	Public Sub SendXMLShipRequest(XMLstring)
		srvFEDEXWSXmlHttp.open "POST", FedExWSURL, false
		srvFEDEXWSXmlHttp.send(XMLstring)
		FEDEXWS_result = srvFEDEXWSXmlHttp.responseText
	End Sub

	Public Sub LoadXMLResults(FEDEXWS_result)
		objOutputXMLDocWS.loadXML FEDEXWS_result
	End Sub


	Public Sub LoadXMLLabel(FEDEXWS_result)
		objFEDEXXmlDoc.loadXML FEDEXWS_result
	End Sub


	Public Sub XMLResponseVerify(ErrPageName)
		on error resume next
		pcv_strErrorMsgWS = objFedExWSClass.ReadResponseNode("//Error", "Message")
		pcv_strErrorCodeReturn = objFedExWSClass.ReadResponseNode("//Error", "Code")
		if len(pcv_strErrorMsgWS)>0 then
			%>
			<!--#include file="pcFedExClassRules.asp" -->
			<%
		end if
	End Sub

	Public Sub XMLResponseVerifyCustom(ErrPageName)
		pcv_strErrorMsgWS = objFedExWSClass.ReadResponseNode("//v9:ProcessShipmentReply", "v9:Notifications/v9:Message")
	End Sub

	Public Function ReadResponseNode(NameOfNode, ValueOfNode)
		on error resume next
		Set Nodes = objOutputXMLDocWS.selectNodes(NameOfNode)
		For Each Node In Nodes

			pcv_strTempValue=Node.selectSingleNode(ValueOfNode).Text
		Next
		ReadResponseNode = pcv_strTempValue
	End Function



	Public Function ReadResponseParent(NameOfNode, ValueOfNode)
		on error resume next
		Set Nodes = objOutputXMLDocWS.selectNodes("//"&NameOfNode)
		For Each Node In Nodes
			pcv_strTempValue=Node.selectSingleNode(ValueOfNode).Text
		Next
		ReadResponseParent = pcv_strTempValue
	End Function


	Public Function ReadResponseasArray(NameOfNode, ValueOfNode)
		on error resume next
		Set Nodes = objOutputXMLDocWS.selectNodes(NameOfNode)
		For Each Node In Nodes
			pcv_strTempValue=Node.selectSingleNode(ValueOfNode).Text
			if pcv_strTempValue="" then
				pcv_strTempValue=" "
			end if
			arryFedExTmp=arryFedExTmp&pcv_strTempValue&","
		Next
		ReadResponseasArray = arryFedExTmp

	End Function



	Public Function pcf_FedExEnabled()
		on error resume next
		pcf_FedExEnabled=false
		query="SELECT ShipmentTypes.active FROM ShipmentTypes WHERE (((ShipmentTypes.idShipment)=9));"
		set rsTmp=server.CreateObject("ADODB.RecordSet")
		set rsTmp=conntemp.execute(query)
		if NOT rsTmp.eof then
			dim FedEX_active
			FedEX_active=rsTmp("active")
			if FedEX_active=true or FedEx_active="-1" then
				pcf_FedExEnabled=true
			end if
		end if
		set rsTmp=nothing
	End Function


	Public Function pcf_FedExPackages(ido)
		on error resume next
		pcf_FedExPackages=false
		query = 		"SELECT pcPackageInfo.idOrder "
		query = query & "FROM pcPackageInfo "
		query = query & "WHERE pcPackageInfo.idOrder=" & ido &" "
		set rsTmp=server.CreateObject("ADODB.RecordSet")
		set rsTmp=conntemp.execute(query)
		if NOT rsTmp.eof then
			FedEX_idOrder=rsTmp("idOrder")
			if FedEX_idOrder=cint(ido) then
				pcf_FedExPackages=true
			end if
		end if
		set rsTmp=nothing
	End Function



	Public Function pcf_FedExSPOD(ido)
		on error resume next
		pcf_FedExSPOD=false
		query = 		"SELECT pcPackageInfo.pcPackageInfo_FDXSPODFlag "
		query = query & "FROM pcPackageInfo "
		query = query & "WHERE pcPackageInfo.pcPackageInfo_ID=" & ido &" "
		set rsTmp=server.CreateObject("ADODB.RecordSet")
		set rsTmp=conntemp.execute(query)
		if NOT rsTmp.eof then
			if rsTmp("pcPackageInfo_FDXSPODFlag") = 1 then
				pcf_FedExSPOD=true
			end if
		end if
		set rsTmp=nothing
	End Function



	Public Function pcf_FedExTrimArray(tmpArray)
		on error resume next
		'// Trim the last comma if there is one
		xStringLength = len(tmpArray)
		if xStringLength>0 then
			pcf_FedExTrimArray = left(tmpArray,(xStringLength-1))
		end if
	End Function


	Public Function pcf_FedExDateFormat(FedExDate)
		on error resume next
		FedExDay=Day(FedExDate)
		FedExMonth=Month(FedExDate)
		FedExYear= Year(FedExDate)
		pcf_FedExDateFormat=FedExYear&"-"&Right(Cstr(FedExMonth + 100),2)&"-"&Right(Cstr(FedExDay + 100),2)
	End Function





	Public Sub pcs_LogTransaction(FedExData, LogFileName, LoggingEnabled)
		on error resume next
		Dim PageName, findit, fs, f
		Set fs=server.CreateObject("Scripting.FileSystemObject")
		If LoggingEnabled = true Then
			Err.number=0

			findit=Server.MapPath("FedExLabels/"&LogFileName)
			if (fs.FileExists(findit))=True OR (fs.FileExists(findit))="True" then
				Set f=fs.GetFile(findit)
				if Err.number=0 then
					f.Delete
				end if
			end if

			if Err.number=0 then
				Set f=fs.OpenTextFile(findit, 2, True)
				f.Write FedExData
				f.Close
			end if

		End If
		Set fs=nothing
		Set f=nothing
	End Sub


end class
'/////////////////////////////////////
'// End building the class here
'/////////////////////////////////////

pcf_FedExWriteLegalDisclaimers = "FedEx service marks are owned by Federal Express Corporation and are used by permission"


%>
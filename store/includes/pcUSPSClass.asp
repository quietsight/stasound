
<%
'/////////////////////////////////////
'// Start building the class here
'/////////////////////////////////////
Class pcUSPSClass 

	private sub Class_Initialize() 
		'// open all object that we will need
		'// define all parameter will use
		'USPS_URL="https://gatewaybeta.USPS.com/GatewayDC"
		Set srvUSPSXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
		Set objOutputXMLDoc = Server.CreateObject("Microsoft.XMLDOM")
		Set objUSPSStream = Server.CreateObject("ADODB.Stream")
		Set objUSPSXmlDoc = server.createobject("Msxml2.DOMDocument"&scXML)
		objUSPSXmlDoc.async = False
		objUSPSXmlDoc.validateOnParse = False
	end sub 
	
	
	
	private sub Class_Terminate() 
		'// clean it all up
		Set srvUSPSXmlHttp = nothing
		Set objOutputXMLDoc = nothing
		Set objUSPSXmlDoc = nothing
		Set objUSPSStream = nothing
	end sub 
	
	
	
	Public Sub AddNewNode(NameOfNode, ValueOfNode, RequiredNode)
		if len(ValueOfNode)>0 OR RequiredNode=1 then
			usps_postdata=usps_postdata&"<"&NameOfNode&">"&ValueOfNode&"</"&NameOfNode&">"&vbcrlf
		end if
	End Sub
	
	Public Sub WriteParent(NameOfParent, isClosing)
		usps_postdata=usps_postdata&"<"&isClosing&""&NameOfParent&">"&vbcrlf
	End Sub
	
	Public Sub WriteEmptyParent(NameOfParent, isClosing)
		usps_postdata=usps_postdata&"<"&NameOfParent&""&isClosing&">"&vbcrlf
	End Sub
	
	Public Sub WriteSingleParent(NameOfParent, ValueOfParent)
		if len(ValueOfParent)>0 then
			usps_postdata=usps_postdata&"<"&NameOfParent&">"&ValueOfParent&"</"&NameOfParent&">"&vbcrlf
		end if
	End Sub

	Public Sub NewXMLTransaction(USPS_APIType, USPS_RequestType, USPS_UserID)
		usps_postdata=""
		usps_postdata=usps_postdata&usps_server&"?API="&USPS_APIType&"&XML="
		usps_postdata=usps_postdata&"<"&USPS_RequestType&" USERID="&chr(34)&USPS_UserID&chr(34)&">"
	End Sub
	
	Public Sub NewXMLLabel(TrackingNumber, EncodedLabelString, FileType, FilePreFix)
		GraphicXML="<Base64Data xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64"" FileName="""&FilePreFix&TrackingNumber&"."&FileType&""">"&EncodedLabelString&"</Base64Data>"
	End Sub
	
	Public Sub SaveBinaryLabel ()
		objUSPSStream.Type = 1
		objUSPSStream.Open
		
		objUSPSStream.Write objUSPSXmlDoc.selectSingleNode("/Base64Data").nodeTypedValue 
			err.clear
		strFileName = objUSPSXmlDoc.selectSingleNode("/Base64Data/@FileName").nodeTypedValue 
		'Save the binary stream to the file and overwrite if it already exists in folder
		objUSPSStream.SaveToFile server.MapPath("USPSLabels\"&strFileName),2
		objUSPSStream.Close()
	End Sub
	
	Public Sub EndXMLTransaction(NameOfMethod)
		usps_postdata=usps_postdata&"</"&NameOfMethod&">"&vbcrlf
	End Sub
	
	
	Public Sub SendXMLRequest(XMLstring, strURL)
		srvUSPSXmlHttp.open "GET", strURL&XMLstring, false
		srvUSPSXmlHttp.send
		USPS_result = srvUSPSXmlHttp.responseText	
		if err>0 then
			'// handle error
		end if
	End Sub
	
	Public Sub LoadXMLResults(USPS_result)
		objOutputXMLDoc.loadXML USPS_result
	End Sub
	
	Public Sub LoadXMLLabel(USPS_result)
		objUSPSXmlDoc.loadXML USPS_result
	End Sub
		
	Public Sub XMLResponseVerify(ErrPageName)
	
		strErrorNumber = objUSPSClass.ReadResponseNode("//Error", "Number") 
		strErrorSource = objUSPSClass.ReadResponseNode("//Error", "Source")
		strErrorDescription = objUSPSClass.ReadResponseNode("//Error", "Description")
		strErrorHelpFile = objUSPSClass.ReadResponseNode("//Error", "HelpFile")
		strErrorHelpContext = objUSPSClass.ReadResponseNode("//Error", "HelpContext")
		
		if len(strErrorNumber)>0 then
			response.redirect ErrPageName & "?LabelMode="&pcv_LabelMode&"&msg=There was an error processing your request.<br>" & strErrorDescription & "<br>USPS Error Code: "&strErrorNumber
		else
			pcv_strErrorMsg=""
		end if
	End Sub
	
	Public Sub XMLResponseVerifyCustom(ErrPageName)
		pcv_strErrorMsg = objUSPSClass.ReadResponseNode("//Error", "ErrorDescription")
	End Sub
	
	Public Function ReadResponseNode(NameOfNode, ValueOfNode)
		on error resume next	
		Set Nodes = objOutputXMLDoc.selectNodes(NameOfNode)
		For Each Node In Nodes
			pcv_strTempValue=Node.selectSingleNode(ValueOfNode).Text  			
		Next
		ReadResponseNode = pcv_strTempValue
	End Function

	Public Function ReadTrackingNode(NameOfNode, ValueOfNode)
		intNodeCnt=0
		Set Nodes = objOutputXMLDoc.selectNodes(NameOfNode)
		For Each Node In Nodes
			pcv_strTempValue=Node.selectSingleNode(ValueOfNode).Text  
			if len(pcv_strTempValue)>1 then
				intNodeCnt=intNodeCnt+1
			end if			
		Next
		ReadTrackingNode = pcv_strTempValue
	End Function

	Public Function ReadResponseParent(NameOfNode, ValueOfNode)	
		on error resume next	
		Set Nodes = objOutputXMLDoc.selectNodes("//"&NameOfNode)	
		For Each Node In Nodes
			pcv_strTempValue=Node.selectSingleNode(ValueOfNode).Text  			
		Next
		ReadResponseParent = pcv_strTempValue
	End Function		
	
	Public Function ReadResponseasArray(NameOfNode, ValueOfNode)	
		on error resume next	
		Set Nodes = objOutputXMLDoc.selectNodes(NameOfNode)	
		For Each Node In Nodes
			pcv_strTempValue=Node.selectSingleNode(ValueOfNode).Text 
			if pcv_strTempValue="" then
				pcv_strTempValue=" "
			end if
			arryUSPSTmp=arryUSPSTmp&pcv_strTempValue&"," 			
		Next
		ReadResponseasArray = arryUSPSTmp

	End Function
	
	Public Function pcf_USPSEnabled()	
		on error resume next
		pcf_USPSEnabled=false	
		query="SELECT ShipmentTypes.active FROM ShipmentTypes WHERE (((ShipmentTypes.idShipment)=4));"
		set rsTmp=server.CreateObject("ADODB.RecordSet")
		set rsTmp=conntemp.execute(query)
		if NOT rsTmp.eof then
			dim USPS_active
			USPS_active=rsTmp("active")
			if USPS_active=true or USPS_active="-1" then
				pcf_USPSEnabled=true
			end if
		end if 
		set rsTmp=nothing		
	End Function
	
	Public Function pcf_USPSURLActive()	
		on error resume next
		pcf_USPSURLActive=true	
		query="SELECT ShipmentTypes.AccessLicense FROM ShipmentTypes WHERE (((ShipmentTypes.idShipment)=4));"
		set rsTmp=server.CreateObject("ADODB.RecordSet")
		set rsTmp=conntemp.execute(query)
		if NOT rsTmp.eof then
			dim USPS_URL
			USPS_URL=rsTmp("AccessLicense")
			if isNull(USPS_URL) or USPS_URL="" then
				pcf_USPSURLActive=false
			end if
		end if 
		set rsTmp=nothing		
	End Function
	
	Public Function pcf_USPSPackages(ido)	
		on error resume next
		pcf_USPSPackages=false			
		query = 		"SELECT pcPackageInfo.idOrder "
		query = query & "FROM pcPackageInfo "
		query = query & "WHERE pcPackageInfo.idOrder=" & ido &" "	
		set rsTmp=server.CreateObject("ADODB.RecordSet")
		set rsTmp=conntemp.execute(query)
		if NOT rsTmp.eof then
			USPS_idOrder=rsTmp("idOrder")
			if USPS_idOrder=cint(ido) then
				pcf_USPSPackages=true
			end if
		end if 
		set rsTmp=nothing
	End Function
	
	
	Public Function pcf_USPSTrimArray(tmpArray)	
		on error resume next
		'// Trim the last comma if there is one
		xStringLength = len(tmpArray)
		if xStringLength>0 then
			pcf_USPSTrimArray = left(tmpArray,(xStringLength-1))
		end if			
	End Function
	
	
	Public Function pcf_USPSDateFormat(USPSDate)
		on error resume next
		USPSDay=Day(USPSDate)
		USPSMonth=Month(USPSDate)
		USPSYear= Year(USPSDate)
		pcf_USPSDateFormat=USPSYear&"-"&Right(Cstr(USPSMonth + 100),2)&"-"&Right(Cstr(USPSDay + 100),2)
	End Function
	
	Public Sub pcs_LogTransaction(USPSData, LogFileName, LoggingEnabled)
		Dim PageName, findit, fs, f
		Set fs=server.CreateObject("Scripting.FileSystemObject")				
		If LoggingEnabled = true Then			
			Err.number=0	
			
			findit=Server.MapPath("USPSLabels/"&LogFileName)	
			if (fs.FileExists(findit))=true then	
				Set f=fs.GetFile(findit)				
				if Err.number=0 then
					f.Delete
				end if
			end if
			
			if Err.number=0 then
				Set f=fs.OpenTextFile(findit, 2, True)
				f.Write USPSData
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

pcf_USPSWriteLegalDisclaimersText = "USPS, THE USPS SHIELD TRADEMARK, THE USPS READY MARK, <br />THE USPS ONLINE TOOLS MARK AND THE COLOR BROWN ARE TRADEMARKS OF<br />UNITED PARCEL SERVICE OF AMERICA, INC. ALL RIGHTS RESERVED."

pcf_USPSWriteLegalDisclaimers = "<table><tr><td width='58' valign='top' bgcolor='#FFFFFF'><div align='right'><img src='../USPSLicense/LOGO_S2.jpg' width='45' height='50' /></div></td><td width='457' valign='top' bgcolor='#FFFFFF'><div align='center'><br />USPS, THE USPS SHIELD TRADEMARK, THE USPS READY MARK, <br />THE USPS ONLINE TOOLS MARK AND THE COLOR BROWN ARE TRADEMARKS OF<br />UNITED PARCEL SERVICE OF AMERICA, INC. ALL RIGHTS RESERVED.</div></td></tr></table>"
%>
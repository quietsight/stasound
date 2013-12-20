<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Shipping Wizard" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/languages_ship.asp" --> 
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/UPSconstants.asp"-->
<!--#include file="../includes/pcUPSClass.asp"-->
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/sendmail.asp" -->
<!--#include file="../includes/emailsettings.asp" -->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../pc/pcPay_GoogleCheckout_Global.asp"-->
<!--#include file="../includes/GoogleCheckout_APIFunctions.asp"-->
<!--#include file="../pc/pcPay_GoogleCheckout_Handler.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcShipTestModes.asp" -->
<% '//Check for error messages
dim msg
msg=request("msg")
if msg<>"" then %>
<table class="pcCPcontent">
	<tr>
		<th colspan="2">UPS OnLine&reg; Tools Shipment Confirmation</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
    <tr>
		<td colspan="2"><div class="pcCPmessage">
        	<% if UPS_TESTMODE="1" then %>
        	UPS Shipping Wizard is currently running in Test Mode <br />
            <% end if %>
        <img src="images/pcadmin_note.gif" alt="Alert" width="20" height="20"> <%=msg%>
    </div></td>
    </tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
</table>
<%
else
	'// PAGE NAME
	pcPageName="pcUPSConfirmLabel.asp"
	ErrPageName="pcUPSConfirmLabel.asp"
	
	dim query, conntemp, rstemp
	Dim objUPSXmlDoc, objUPSStream, GraphicXML
	Dim UPS_postdata, objUPSClass, objOutputXMLDoc, srvUPSXmlHttp, UPS_result, UPS_URL, pcv_strErrorMsg, pcv_strAction
	
	'// Define objects used to create and send Google Checkout Order Processing API requests
	Dim xmlRequest
	Dim xmlResponse
	Dim attrGoogleOrderNumber
	Dim elemAmount 
	Dim elemReason
	Dim elemComment
	Dim elemCarrier
	Dim elemTrackingNumber
	Dim elemMessage
	Dim elemSendEmail
	Dim elemMerchantOrderNumber
	Dim transmitResponse
	
	shipmentTotal=Cdbl(0)
	
	
	call openDb()
	
	'//UPS Variables
	query="SELECT active, userID, [password], AccessLicense FROM ShipmentTypes WHERE idshipment=3"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	ups_active=rs("active")
	ups_userid=trim(rs("userID"))
	ups_password=trim(rs("password"))
	ups_license_key=trim(rs("AccessLicense"))
	set rs=nothing
	call closedb()
	
	Dim iUPSFlag
	iUPSFlag=0
	iUPSActive=1
	'// SET THE UPS OBJECT
	set objUPSClass = New pcUPSClass
	'//UPS Rates
	objUPSClass.NewXMLTransaction ups_license_key, ups_userid, ups_password
	
	objUPSClass.NewXMLShipmentAcceptRequest session("UPSShippingShipmentCustomerContext"), session("UPSShippingShipmentDigest")
	
	'//Clear illegal ampersand characters from XML
	UPS_postdata=replace(UPS_postdata, "&", "and")
	UPS_postdata=replace(UPS_postdata, "andamp;", "and")

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Send Our Transaction.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	if UPS_TESTMODE="1" then
		UPS_URL="https://wwwcie.ups.com/ups.app/xml/ShipAccept"
	else
		UPS_URL="https://www.ups.com/ups.app/xml/ShipAccept"
	end if
	call objUPSClass.SendXMLRequest(UPS_postdata, UPS_URL)
	
	'// Print out our response
	'response.write UPS_result
	'response.end
							
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Load Our Response.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	call objUPSClass.LoadXMLResults(UPS_result)
							
							
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Check for errors from UPS.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	
	'//SOME ERROR CHECKING HERE							
	call objUPSClass.XMLResponseVerify(ErrPageName)
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Redirect with a Message OR complete some task.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	if NOT len(pcv_strErrorMsg)>0 then
							
	
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Set Our Response Data to Local.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~							
		ShipmentCustomerContext = objUPSClass.ReadResponseNode("//TransactionReference", "CustomerContext")
		session("UPSShippingShipmentCustomerContext")=ShipmentCustomerContext
	
		pcArryExists=0
		if instr(ShipmentCustomerContext,",") then
			pcArryCustomerContext=split(ShipmentCustomerContext,",")
			pcArryExists=1
		end if
	
		intLabelCnt=0
		dim LNodes, intLNode, LNode
		'LabelImage
		Set Nodes = objOutputXMLDoc.selectNodes("//PackageResults")
		intNodeCnt=0
		For Each Node In Nodes
			if len(Node.selectSingleNode("TrackingNumber").text)>1 then
				intNodeCnt=intNodeCnt+1
				strTrackingNumber=Node.selectSingleNode("TrackingNumber").Text
			end if
	
			Set LNodes = objOutputXMLDoc.selectNodes("//LabelImage")
			intLNode=0
			intHTMLExist=0
			For Each LNode In LNodes
				intLNode=intLNode+1
				if intLNode=intNodeCnt then	
					strLabelImageFormat=LNode.selectSingleNode("LabelImageFormat").Text
					strGraphicImage=LNode.selectSingleNode("GraphicImage").Text
					if instr(UPS_result, "HTMLImage") then
						intHTMLExist=1
						strHTMLImage=LNode.selectSingleNode("HTMLImage").Text
					end if
				end if
			Next
			
			if strLabelImageFormat="EPL" then
				GraphicXML="<Base64Data xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64"" FileName=""label"&strTrackingNumber&".txt"">"&strGraphicImage&"</Base64Data>"
			else
				strLabelImageFormat="GIF"
				GraphicXML="<Base64Data xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64"" FileName=""label"&strTrackingNumber&".gif"">"&strGraphicImage&"</Base64Data>"
			end if
		
			Dim objXMLDoc 
			Dim objStream
			Dim strFileName
		
			'Create MSXML DOMDocument Object
			Set objXMLDoc = Server.CreateObject("MSXML2.DOMDocument"&scXML)
			objXMLDoc.async = False
			objXMLDoc.validateOnParse = False
			
			'And load it from the request stream
			If objXMLDoc.loadXML (GraphicXML) Then
	
				'Use ADO stream to save the binary data
				Set objStream = Server.CreateObject("ADODB.Stream")
				objStream.Type = 1
				objStream.Open
	
				'The nodeTypedValue automagically converts Base64 data to binary data
				'Write that binary data to the stream
				objStream.Write objXMLDoc.selectSingleNode("/Base64Data").nodeTypedValue 
				
				'Get the FileName attribute's value
				strFileName = objXMLDoc.selectSingleNode("/Base64Data/@FileName").nodeTypedValue 
				
				'on error resume next
				err.clear
				'Save the binary stream to the file
				objStream.SaveToFile server.mappath("UPSLabels\" & strFileName), 2
				if err.number<>0 then
					'response.write "This label has already been saved..."
				end if
				objStream.Close()
				
				Set objStream = Nothing
			Else
				'Failed to load the document
				response.write "<br><br>Failed to load doc..."
			End If	
			
			if intHTMLExist=1 then
				'Create XML
				HTMLXML="<Base64Data xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64"" FileName=""label"&strTrackingNumber&".html"">"&strHTMLImage&"</Base64Data>"
		
				'Create MSXML DOMDocument Object
				Set objXMLDoc = Server.CreateObject("MSXML2.DOMDocument"&scXML)
				objXMLDoc.async = False
				objXMLDoc.validateOnParse = False
		
				'And load it from the request stream
				If objXMLDoc.loadXML (HTMLXML) Then
		
					'Use ADO stream to save the binary data
					Set objStream = Server.CreateObject("ADODB.Stream")
					objStream.Type = 1
					objStream.Open
			
					'The nodeTypedValue automagically converts Base64 data to binary data
					'Write that binary data to the stream
					objStream.Write objXMLDoc.selectSingleNode("/Base64Data").nodeTypedValue 
		
					'Get the FileName attribute's value
					strFileName2 = objXMLDoc.selectSingleNode("/Base64Data/@FileName").nodeTypedValue 
					'on error resume next
					err.clear
					'Save the binary stream to the file
					objStream.SaveToFile server.mappath("UPSLabels\" & strFileName2), 2
					if err.number<>0 then
						'response.write "This label has already been saved..."
					end if
					objStream.Close()
					
					Set objStream = Nothing
				Else
					'Failed to load the document
					response.write "<br><br>Failed to load doc..."
				End If
			end if
				
			'save trackingNumber in the database
			intLabelCnt=intLabelCnt+1
			if pcArryExists=1 then
				pcTempPackageInfo_ID=pcArryCustomerContext(intLabelCnt-1)
				'pcArryCustomerContext=split(ShipmentCustomerContext,",")
			else
				pcTempPackageInfo_ID=ShipmentCustomerContext
			end if
		
			call opendb()
			'//////////////////////////////////////////////////////////////
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' START: SAVE PACKAGES
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
			pcv_PackageCount = session("pcAdminPackageCount")
			if pcv_PackageCount="" then
				pcv_PackageCount=1
			end if
	
			'For pcv_xCounter = 1 to pcv_PackageCount
			'// Get Our Required Data
			select case Session("pcAdminUPSServiceCode")
				case "01"
					pcv_method="UPS Next Day Air"
				case "02"
					pcv_method="UPS 2nd Day Air"
				case "03"
					pcv_method="UPS Ground"
				case "07"
					pcv_method="UPS Worldwide Express"
				case "08"
					pcv_method="UPS Worldwide Expedited"
				case "11"
					pcv_method="UPS Standard To Canada"
				case "12"
					pcv_method="UPS 3 Day Select"
				case "13"
					pcv_method="UPS Next Day Air Saver"
				case "14"
					pcv_method="UPS Next Day Air"
				case "54"
					pcv_method="UPS Worldwide Express Plus"
				case "59"
					pcv_method="UPS 2nd Day Air A.M."
				case "65"
					pcv_method="UPS Express Saver"
			end select
			pcv_method = "UPS: " & pcv_method

			pcv_AdmComments=""
						
			'// Fix quotes on comments
			if pcv_AdmComments<>"" then
				pcv_AdmComments=replace(pcv_AdmComments,"'","''")
			end if
								
			'// Insert Details into Package Info
			'TO DO: 
			pcv_intOrderID=Session("pcAdminOrderID")
			
			dim dtShippedDate
			dtShippedDate=Date()
			if pcv_shippedDate<>"" then									
				'dtShippedDate=objFedExClass.pcf_FedExDateFormat(dtShippedDate)
				if SQL_Format="1" then
					dtShippedDate=(day(dtShippedDate)&"/"&month(dtShippedDate)&"/"&year(dtShippedDate))
				else
					dtShippedDate=(month(dtShippedDate)&"/"&day(dtShippedDate)&"/"&year(dtShippedDate))
				end if
			end if
				
			if scDB="Access" then
				pcInsertDate="#"
			else
				pcInsertDate="'"
			end if	
			query="INSERT INTO pcPackageInfo (idOrder, pcPackageInfo_ShipMethod, pcPackageInfo_ShippedDate, pcPackageInfo_TrackingNumber, pcPackageInfo_PackageNumber, pcPackageInfo_UPSLabelFormat) " 
			query=query&"VALUES (" & pcv_intOrderID & ", '" & pcv_method & "', "&pcInsertDate&dtShippedDate&pcInsertDate&", '" & strTrackingNumber & "', '" & intLabelCnt & "', '"&strLabelImageFormat&"');"
			set rs=connTemp.execute(query)
			set rs=nothing
							
			'// Re-Query for the ID
			query="SELECT pcPackageInfo_ID FROM pcPackageInfo WHERE idorder=" & pcv_intOrderID & " ORDER by pcPackageInfo_ID DESC;"
			set rs=connTemp.execute(query)
			pcv_PackageID=rs("pcPackageInfo_ID")
			set rs=nothing
			qry_ID=pcv_intOrderID
			
			query="SELECT orders.pcOrd_GoogleIDOrder FROM orders WHERE idOrder="& qry_ID
			set rs=server.CreateObject("ADODB.RecordSet")
			Set rs=conntemp.execute(query)
			if Not rs.eof then
				pcv_strGoogleIDOrder = rs("pcOrd_GoogleIDOrder") '// determine if this is a google order
			end if
			set rs=nothing
						
			'// Do a Full Update
			query="UPDATE pcPackageInfo "
			query=query&"SET pcPackageInfo_PackageNumber="&intNodeCnt&", "
			query=query&"pcPackageInfo_PackageWeight=" & Session("pcAdminPackageWeight"&intNodeCnt) & ", "						
			query=query&"pcPackageInfo_ShipToName='" & Session("pcAdminShipToCompanyName") & "', "
			query=query&"pcPackageInfo_ShipToAddress1='" & Session("pcAdminShipToAddressLine1") & "', "
			query=query&"pcPackageInfo_ShipToAddress2='" & Session("pcAdminShipToAddressLine2") & "', "						
			query=query&"pcPackageInfo_ShipToCity='" & Session("pcAdminShipToCity") & "', "
			query=query&"pcPackageInfo_ShipToStateCode='" & Session("pcAdminShipToStateOrProvinceCode") & "', "
			query=query&"pcPackageInfo_ShipToZip='" & Session("pcAdminShipToPostalCode") & "', "
			query=query&"pcPackageInfo_ShipToCountry='" & Session("pcAdminShipToCountryCode") & "', "
			query=query&"pcPackageInfo_ShipToPhone='" & Session("pcAdminShipToPhoneNumber") & "', "
			query=query&"pcPackageInfo_ShipToEmail='" & Session("pcAdminShipToEmailAddress") & "', "
			IF Session("pcAdminResidentialDelivery")="" then
				Session("pcAdminResidentialDelivery")=0
			end if				
			query=query&"pcPackageInfo_ShipToResidential=" & Session("pcAdminResidentialDelivery") & ", "
			query=query&"pcPackageInfo_PackageDescription='" & pcv_strPackagingDescription & "', "					
			query=query&"pcPackageInfo_ShipFromCompanyName='" & Session("pcAdminShipFromCompanyName") & "', "
			query=query&"pcPackageInfo_ShipFromAttentionName='" & Session("pcAdminShipFromAttentionName") & "', "
			query=query&"pcPackageInfo_ShipFromPhoneNumber='" & Session("pcAdminShipFromPhoneNumber") & "', "
			query=query&"pcPackageInfo_ShipFromAddress1='" & Session("pcAdminShipFromAddressLine1") & "', "
			query=query&"pcPackageInfo_ShipFromAddress2='" & Session("pcAdminShipFromAddressLine2") & "', "
			query=query&"pcPackageInfo_ShipFromCity='" & Session("pcAdminShipFromCity") & "', "
			query=query&"pcPackageInfo_ShipFromStateProvinceCode='" & Session("pcAdminShipFromStateOrProvinceCode") & "', "
			query=query&"pcPackageInfo_ShipFromPostalCode='" & Session("pcAdminShipFromPostalCode") & "', "
			query=query&"pcPackageInfo_ShipFromCountryCode='" & Session("pcAdminShipFromCountryCode") & "', "						
			query=query&"pcPackageInfo_UPSServiceCode='" & Session("pcAdminUPSServiceCode") & "', "
			query=query&"pcPackageInfo_UPSPackageType='" & Session("pcAdminPackageTypeCode"&intNodeCnt) & "', "						
			query=query&"pcPackageInfo_PackageInsuredValue='" & Session("pcAdminInsuredValue"&intNodeCnt) & "', "						
			query=query&"pcPackageInfo_PackageLength='" & Session("pcAdminLength"&intNodeCnt) & "', "
			query=query&"pcPackageInfo_PackageWidth='" & Session("pcAdminWidth"&intNodeCnt) & "', "
			query=query&"pcPackageInfo_PackageHeight='" & Session("pcAdminHeight"&intNodeCnt) & "', "						
			query=query&"pcPackageInfo_MethodFlag=2 "			
			query=query&"WHERE pcPackageInfo_ID=" & pcv_PackageID & " ;"
			set rstemp=connTemp.execute(query)
			set rs=nothing
							
			'// Delete the old comments
			query="DELETE FROM pcAdminComments WHERE idorder=" & qry_ID & " AND pcACom_ComType=2 AND pcPackageInfo_ID=" & pcv_PackageID & ";"
			set rstemp=connTemp.execute(query)
			'// Add the new comments
			query=		"INSERT INTO pcAdminComments (idorder,pcACom_ComType,pcACom_Comments,pcDropShipper_ID,pcACom_IsSupplier,pcPackageInfo_ID) "
			query=query&"VALUES (" & qry_ID & ",2,'" & pcv_AdmComments & "',0,0," & pcv_PackageID & ");"
			set rstemp=connTemp.execute(query)
			
			if trim(Session("pcAdminPrdList"&intNodeCnt))<>"" then
				pcA=split(Session("pcAdminPrdList"&intNodeCnt),",")
				For i=lbound(pcA) to ubound(pcA)
					if trim(pcA(i)<>"") then
						query="UPDATE ProductsOrdered SET pcPrdOrd_Shipped=1, pcPackageInfo_ID=" & pcv_PackageID & " WHERE (idorder=" & qry_ID & " AND idProductOrdered=" & pcA(i) & ");"
						set rs=connTemp.execute(query)
						set rs=nothing
					end if
				Next
			else
				query="UPDATE ProductsOrdered SET pcPackageInfo_ID=" & pcv_PackageID & " WHERE idorder=" & qry_ID & " AND pcPrdOrd_Shipped=0 AND pcDropShipper_ID=0;"
				set rsQ=connTemp.execute(query)
				set rsQ=nothing
			end if
			
			pcv_SendCust="1"
			pcv_SendAdmin="0"
			pcv_LastShip="0"
				
			query="SELECT ProductsOrdered.pcPrdOrd_Shipped FROM ProductsOrdered INNER JOIN Orders ON (ProductsOrdered.idorder=Orders.idorder AND ProductsOrdered.pcPrdOrd_Shipped=0) WHERE Orders.idorder=" & qry_ID & " AND Orders.orderstatus<>4;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
			
			if not rs.eof then
				pcv_LastShip="0"
			else
				pcv_LastShip="1"
			end if
			set rs=nothing
						
			if trim(Session("pcAdminPrdList"&intNodeCnt))<>"" then
				if pcv_LastShip="1" then
					query="UPDATE Orders SET orderStatus=4 WHERE idorder=" & qry_ID & ";"
				else
					query="UPDATE Orders SET orderStatus=7 WHERE idorder=" & qry_ID & ";"
				end if
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				set rs=nothing
			else
				query="UPDATE Orders SET orderStatus=4 WHERE idorder=" & qry_ID & ";"
			end if
			
			If pcv_LastShip="1" Then
				'// Perform a Google Action	
				pcv_strGoogleMethod = "mark" ' // Marks the order shipped at Google 			
				%> <!--#include file="../includes/GoogleCheckout_OrderManagement.asp"--> <%
			End If
			%>
			<!--#include file="../pc/inc_PartShipEmail.asp"-->	
			<%							
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' END: SAVE PACKAGES
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
			%>
			<form action="pcUPSConfirmLabel.asp" method="get">
			<table width="200" border="0">
				<tr>
					<td><p><img src="upslabels/<%=strFileName%>" width="241" height="164"></p>
						<p><a href="upslabels/<%=strFileName2%>" target="_blank">Print Label <%=intNodeCnt%></a></p></td>
					</tr>
			</table>
			</form>
			
		<% Next %>
        <!--#include file="UPS_ClearSessions.asp"-->	
        <%
		response.redirect "UPS_ManageShipmentsResults.asp?id="&qry_ID
	end if 'if ups is active
	
	'// DESTROY THE UPS OBJECT
	set objUPSClass = nothing
end if
%>
<!--#include file="AdminFooter.asp"-->
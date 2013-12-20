<%
'--------------------------------------------------------------
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<%
on error resume next

dim resolveTimeout, connectTimeout, sendTimeout, receiveTimeout

resolveTimeout	= 5000
connectTimeout	= 5000
sendTimeout		= 5000
receiveTimeout	= 10000
'1000ms = 1 sec

'UPS DEBUGGING VARIABLES
'******************************************************************
'// Debug UPS post and reply
'// To turn UPS debugging on, change the value of pcv_UPSDebug=1
'******************************************************************
pcv_UPSDebug=0

'******************************************************************
'// Log UPS reply
'// To turn UPS logging on, change the value of pcv_UPS_Logging=1
'******************************************************************
pcv_UPS_Logging=0

'UPS CANADA ONLY VARIABLES
'******************************************************************
'// Use Canada as the Ship From Origin
'// To set Canada as the Ship From Origin,
'// change the value of pcv_UPSCanadaOrigin=1
'// You MUST also run the Script "upddbUPSShipOrigin.asp" that
'// is located in your ProductCart Control Panel Folder
'******************************************************************
dim pcv_UPSCanadaOrigin
pcv_UPSCanadaOrigin=0


'U.S.P.S. OPTIONAL VARIABLES
'******************************************************************
'// USPS Value of Content for International Rates Only
'// If specified, it is used to compute Insurance fee
'// (if insurance is available for service and destination) and
'// indemnity coverage.
'// To turn this variable on, change the value to "1"
'//
'// For Example:
'// pcv_UseValueOfContents=1

'******************************************************************
pcv_UseValueOfContents=1
'******************************************************************
if pcv_UseValueOfContents=1 then
	pcv_ValueOfContents=pSubTotal
end if

'Set variables from Constants UPS
pcv_UseNegotiatedRates=UPS_USENEGOTIATEDRATES
pcv_UPSShipperNumber=UPS_SHIPPERNUM
pcv_InsuredValue=UPS_INSUREDVALUE
pcv_UseDynamicInsuredValue=UPS_DYNAMICINSUREDVALUE

'Set variables from Constants FEDEX
pcv_InsuredValue_FDX=FDX_INSUREDVALUE '// SD
pcv_UseDynamicInsuredValue_FDX=FDX_DYNAMICINSUREDVALUE '// SD

'Set variables from Constants FEDEX WS
pcv_InsuredValue_FDXWS=FDXWS_INSUREDVALUE '// WS
pcv_UseDynamicInsuredValue_FDXWS=FDXWS_DYNAMICINSUREDVALUE '// WS

if pcv_UseDynamicInsuredValue="1" then
	pcv_InsuredValue=pSubTotal
end if

if pcv_UseDynamicInsuredValue_FDX="1" then
	pcv_InsuredValue_FDX=Cdbl(calculateShipCartTotal(pcCartArray, ppcCartIndex))
end if

if pcv_UseDynamicInsuredValue_FDXWS="1" then
	pcv_InsuredValue_FDXWS=Cdbl(calculateShipCartTotal(pcCartArray, ppcCartIndex))
end if

iFedExFlag=0 '// SD
iFedExWSFlag=0 '// WS
iCPFlag=0
iUSPSFlag=0
iCustomFlag=0
strOptionShipmentType=""
strTabShipmentType=""

pcv_intPackageNum=0
pcv_intTotPackageNum=0

dim pcv_EMWeightLimit, pcv_PMWeightLimit,pcv_EM_Null,pcv_PM_Null

pcv_EMWeightLimit=USPS_EM_FREWeightLimit
if NOT isNumeric(pcv_EMWeightLimit) or pcv_EMWeightLimit="" then
	pcv_EMWeightLimit=0
end if
pcv_PMWeightLimit=USPS_PM_FREWeightLimit
if NOT isNumeric(pcv_PMWeightLimit) or pcv_PMWeightLimit="" then
	pcv_PMWeightLimit=0
end if
pcv_EM_Null=0
pcv_PM_Null=0


'// Check if any products are labeled as oversize for UPS & FedEX & USPS
Dim pcv_intOSCheck, pcv_intOSStatus, pcv_arrOSCheckArray, pcv_arrOSArray
if pcv_EOSC="" then
	pcv_intOSCheck=oversizecheck(pcCartArray, ppcCartIndex)
else
	pcv_intOSCheck=eoversizecheck(request("idOrder"))
end if
pcv_intOSStatus=0

'  If products are oversize, double check to be sure values exists
if pcv_intOSCheck<>"" then
	pcv_arrOSCheckArray=split(pcv_intOSCheck,",")
	for i=0 to Ubound(pcv_arrOSCheckArray)-1
		pcv_arrOSArray=split(pcv_arrOSCheckArray(i),"|||")
		if pcv_arrOSArray(0)>pcv_intOSStatus then
			pcv_intOSStatus=1
		end if
	next
end if

strPackageInfo=""
dim intPackageCnt, intWeightCnt
intPackageCnt=0
intWeightCnt=0
dim pcv_intOSwidth, pcv_intOSheight, pcv_intOSlength, intOSstatus
'=====================================================================
' For each oversized package, get height, width, length and weight
' keep a running package count
'---------------------------------------------------------------------

if pcv_intOSStatus<>0 then '// There are OS packages
	'keep track of BTO/OS Items
	for i=0 to Ubound(pcv_arrOSCheckArray)-1  ' loop through OS packages
		intOSweight=0
		pcv_arrOSArray=split(pcv_arrOSCheckArray(i),"|||")
		if pcv_arrOSArray(0)>pcv_intOSStatus then
			pcv_arrOSArray2=pcv_arrOSArray(1)
			pcv_strOSString=split(pcv_arrOSArray2,"||")
			if ubound(pcv_strOSString)=-1 then
				pcv_intOSheight=UPS_HEIGHT
				pcv_intOSwidth=UPS_WIDTH
				pcv_intOSlength=UPS_LENGTH
				pcv_intOSStatus=0
			else
				intPackageCnt=intPackageCnt+1
				strPackageInfo=strPackageInfo&"<tr><td>Oversized Package "&intPackageCnt&"</td>"
				pcv_intOSwidth=pcv_strOSString(0)
				session("UPSPackWidth"&intPackageCnt)=pcv_strOSString(0)
				session("FEDEXPackWidth"&intPackageCnt)=pcv_strOSString(0) '// SD
				session("FEDEXWSPackWidth"&intPackageCnt)=pcv_strOSString(0) '// WS
				session("CPPackWidth"&intPackageCnt)=pcv_strOSString(0)
				session("USPSPackWidth"&intPackageCnt)=pcv_strOSString(0)
				pcv_intOSheight=pcv_strOSString(1)
				session("UPSPackHeight"&intPackageCnt)=pcv_strOSString(1)
				session("FEDEXPackHeight"&intPackageCnt)=pcv_strOSString(1) '// SD
				session("FEDEXWSPackHeight"&intPackageCnt)=pcv_strOSString(1) '// WS
				session("CPPackHeight"&intPackageCnt)=pcv_strOSString(1)
				session("USPSPackHeight"&intPackageCnt)=pcv_strOSString(1)
				pcv_intOSlength=pcv_strOSString(2)
				session("UPSPackLength"&intPackageCnt)=pcv_strOSString(2)
				session("FEDEXPackLength"&intPackageCnt)=pcv_strOSString(2) '// SD
				session("FEDEXWSPackLength"&intPackageCnt)=pcv_strOSString(2) '// WS
				session("CPPackLength"&intPackageCnt)=pcv_strOSString(2)
				session("USPSPackLength"&intPackageCnt)=pcv_strOSString(2)
				pcv_intOSPrice=pcv_strOSString(6)

				'// Price of OverSized Package UPS
				if pcv_UseDynamicInsuredValue="1" then
					session("UPSPackPrice"&intPackageCnt)=pcv_intOSPrice
					'// subtract the price of this OS package from the subtotal if dynamic insured value is used in cart.
					pcv_InsuredValue=ccur(pcv_InsuredValue)-cdbl(pcv_intOSPrice)
				else
					session("UPSPackPrice"&intPackageCnt)=UPS_INSUREDVALUE
				end if

				'// Price of OverSized Package FedEX SD
				if pcv_UseDynamicInsuredValue_FDX="1" then
					session("FEDEXPackPrice"&intPackageCnt)=pcv_intOSPrice
					'// subtract the price of this OS package from the subtotal if dynamic insured value is used in cart.
					pcv_InsuredValue_FDX=ccur(pcv_InsuredValue_FDX)-cdbl(pcv_intOSPrice)
				else
					session("FEDEXPackPrice"&intPackageCnt)=FDX_INSUREDVALUE
				end if

				'// Price of OverSized Package FedEX WS
				if pcv_UseDynamicInsuredValue_FDXWS="1" then
					session("FEDEXWSPackPrice"&intPackageCnt)=pcv_intOSPrice
					'// subtract the price of this OS package from the subtotal if dynamic insured value is used in cart.
					pcv_InsuredValue_FDXWS=ccur(pcv_InsuredValue_FDXWS)-cdbl(pcv_intOSPrice)
				else
					session("FEDEXWSPackPrice"&intPackageCnt)=FDXWS_INSUREDVALUE
				end if

				intOSweight=pcv_strOSString(5)
				if pcv_EMWeightLimit<>0 AND intOSweight>Clng((pcv_EMWeightLimit*16)) then
					pcv_EM_Null=1
				end if
				if pcv_PMWeightLimit<>0 AND Clng(intOSweight)>Clng((pcv_PMWeightLimit*16)) then
					pcv_PM_Null=1
				end if
				intWeightCnt=intWeightCnt+intOSweight
				if scShipFromWeightUnit="KGS" then
					intOSintPounds=int(intOSweight/1000)
					intOSounces=intOSweight-(intOSintPounds*1000)
				else
					intOSintPounds=Int(intOSweight/16) 'intPounds used for USPS
					intOSounces=intOSweight-(intOSintPounds*16) 'intUniversalOunces used for USPS
				end if
				if scShipFromWeightUnit="KGS" then
					session("USPSPackPounds"&intPackageCnt)=intOSintKilos
					session("USPSPackOunces"&intPackageCnt)=intOSgrams
					session("BasicPackPounds"&intPackageCnt)=intOSintKilos
					session("BasicPackOunces"&intPackageCnt)=intOSgrams
				else
					session("USPSPackPounds"&intPackageCnt)=intOSintPounds
					session("USPSPackOunces"&intPackageCnt)=intOSounces
					session("BasicPackPounds"&intPackageCnt)=intOSintPounds
					session("BasicPackOunces"&intPackageCnt)=intOSounces
				end if
				intMPackageWeight=intOSintPounds
				if intMPackageWeight<1 AND intOSounces<1 then
					intMPackageWeight=0
				end if
				if intMPackageWeight<1 AND intOSounces>0 then 'if total weight is less then a pound, make UPS/FedEX weight 1 pound
					intMPackageWeight=1
				else  'total weight is not less then a pound and ounces exist, round weight up one more pound.
					If intMPackageWeight>0 AND intOSounces>0 then
						intMPackageWeight=(intMPackageWeight+1)
					End if
				end if
				strPackageInfo=strPackageInfo&"<td>Weight "&intMPackageWeight&"</td></tr>"
				pcv_intTotPackageNum=pcv_intTotPackageNum+1
				session("UPSPackWeight"&intPackageCnt)=intMPackageWeight
				session("FEDEXPackWeight"&intPackageCnt)=intMPackageWeight
				session("FEDEXWSPackWeight"&intPackageCnt)=intMPackageWeight
				session("CPPackWeight"&intPackageCnt)=intMPackageWeight
			end if
		end if
	next '// End loop through OS packages
	dim intOSpackageCnt
	intOSpackageCnt=intPackageCnt
else '// There are OS packages
	'no oversized packages
	pcv_intOSStatus=0
end if '// There are OS packages



'=====================================================================
intCustomShipWeight=intUniversalWeight
pShipWeight=intUniversalWeight-intWeightCnt
'// No oversized items were in cart, packagecount at 1
if pcv_intOSStatus=0 then
	intPackageCnt=0
end if
if pShipWeight>0 then 'Weight > 0
	if scShipFromWeightUnit="KGS" then
		intPounds=Int(pShipWeight/1000)
		intUniversalOunces=pShipWeight-(intPounds*1000) 'intUniversalOunces used for USPS
	else
		intPounds=Int(pShipWeight/16) 'intPounds used for USPS
		intUniversalOunces=pShipWeight-(intPounds*16) 'intUniversalOunces used for USPS
	end if
	intUniversalWeight=intPounds
	if intUniversalWeight<1 AND intUniversalOunces<1 then
		intUniversalWeight=0
	end if
	if intUniversalWeight<1 AND intUniversalOunces>0 then 'if total weight is less then a pound, make UPS/FedEX weight 1 pound
		intUniversalWeight=1
	else  'total weight is not less then a pound and ounces exist, round weight up one more pound.
		If intUniversalWeight>0 AND intUniversalOunces>0 then
			intUniversalWeight=(intUniversalWeight+1)
		End if
	end if
	'=====================================================================
	' check to see if there is a weight limit set for packages >0
	'---------------------------------------------------------------------
	if int(scPackageWeightLimit)<>0 then '// There is a package Weight limit set
		'see how many package this should be if over the limit
		if int(intUniversalWeight)>int(scPackageWeightLimit) then '// There are more package after OS

			intTempPackageNum=(intUniversalWeight/int(scPackageWeightLimit))
			pcv_intPackageNum=int(intUniversalWeight/int(scPackageWeightLimit))
			if intTempPackageNum>pcv_intPackageNum then
				pcv_intPackageNum=pcv_intPackageNum+1
			end if
			if pcv_UseDynamicInsuredValue="1" then
				pcv_TempInsuredValue=cdbl(pcv_InsuredValue)/Cint(pcv_intPackageNum)
			else
				pcv_TempInsuredValue=cdbl(pcv_InsuredValue)
			end if
			if pcv_UseDynamicInsuredValue_FDX="1" then
				pcv_TempInsuredValue_FDX=cdbl(pcv_InsuredValue_FDX)/Cint(pcv_intPackageNum)
			else
				pcv_TempInsuredValue_FDX=cdbl(pcv_InsuredValue_FDX)
			end if
			if pcv_UseDynamicInsuredValue_FDXWS="1" then
				pcv_TempInsuredValue_FDXWS=cdbl(pcv_InsuredValue_FDXWS)/Cint(pcv_intPackageNum)
			else
				pcv_TempInsuredValue_FDXWS=cdbl(pcv_InsuredValue_FDXWS)
			end if
			for r=1 to (pcv_intPackageNum-1)
				intPackageCnt=intPackageCnt+1
				strPackageInfo=strPackageInfo&"<tr><td>Package "&intPackageCnt&"</td>"
				strPackageInfo=strPackageInfo&"<td>Weight "&scPackageWeightLimit&"</td></tr>"
				pcv_intTotPackageNum=pcv_intTotPackageNum+1
				if ups_active=true or ups_active="-1" then
					session("UPSPackWidth"&intPackageCnt)=UPS_WIDTH
					session("UPSPackHeight"&intPackageCnt)=UPS_HEIGHT
					session("UPSPackLength"&intPackageCnt)=UPS_LENGTH
					session("UPSPackWeight"&intPackageCnt)=scPackageWeightLimit
					session("UPSPackPrice"&intPackageCnt)=pcv_TempInsuredValue
				end if
				if FedEX_active=true or FedEx_active="-1" then
					session("FedEXPackWidth"&intPackageCnt)=FEDEX_WIDTH
					session("FedEXPackHeight"&intPackageCnt)=FEDEX_HEIGHT
					session("FedEXPackLength"&intPackageCnt)=FEDEX_LENGTH
					session("FedEXPackWeight"&intPackageCnt)=scPackageWeightLimit
					session("FedEXPackPrice"&intPackageCnt)=pcv_TempInsuredValue_FDX
				end if
				if FedEXWS_active=true or FedExWS_active="-1" then
					session("FedEXWSPackWidth"&intPackageCnt)=FEDEXWS_WIDTH
					session("FedEXWSPackHeight"&intPackageCnt)=FEDEXWS_HEIGHT
					session("FedEXWSPackLength"&intPackageCnt)=FEDEXWS_LENGTH
					session("FedEXWSPackWeight"&intPackageCnt)=scPackageWeightLimit
					session("FedEXWSPackPrice"&intPackageCnt)=pcv_TempInsuredValue_FDXWS
				end if
				if USPS_active=true or USPS_active="-1" then
					session("USPSPackWidth"&intPackageCnt)=USPS_WIDTH
					session("USPSPackHeight"&intPackageCnt)=USPS_HEIGHT
					session("USPSPackLength"&intPackageCnt)=USPS_LENGTH
					session("USPSPackPounds"&intPackageCnt)=scPackageWeightLimit
					session("USPSPackOunces"&intPackageCnt)=0
					if pcv_EMWeightLimit<>0 AND scPackageWeightLimit>Clng(pcv_EMWeightLimit) then
						pcv_EM_Null=1
					end if
					if pcv_PMWeightLimit<>0 AND scPackageWeightLimit>Clng(pcv_PMWeightLimit) then
						pcv_PM_Null=1
					end if
				end if
				If CP_active=true or CP_active="-1" then
					session("CPPackWidth"&intPackageCnt)=CP_Width
					session("CPPackHeight"&intPackageCnt)=CP_Height
					session("CPPackLength"&intPackageCnt)=CP_Length
					session("CPPackWeight"&intPackageCnt)=scPackageWeightLimit
				end if
				session("BasicPackPounds"&intPackageCnt)=intPounds
				session("BasicPackOunces"&intPackageCnt)=intUniversalOunces
			next
			'last package
			intLastPackageWeight=int(intUniversalWeight-((pcv_intPackageNum-1)*scPackageWeightLimit))
			intPackageCnt=intPackageCnt+1
			strPackageInfo=strPackageInfo&"<tr><td>Package "&intPackageCnt&"</td>"
			strPackageInfo=strPackageInfo&"<td>Weight "&intLastPackageWeight&"lb. "&intUniversalOunces&"oz.</td></tr>"
			pcv_intTotPackageNum=pcv_intTotPackageNum+1
			if ups_active=true or ups_active="-1" then
				session("UPSPackWidth"&intPackageCnt)=UPS_WIDTH
				session("UPSPackHeight"&intPackageCnt)=UPS_HEIGHT
				session("UPSPackLength"&intPackageCnt)=UPS_LENGTH
				session("UPSPackWeight"&intPackageCnt)=intLastPackageWeight
				session("UPSPackPrice"&intPackageCnt)=pcv_TempInsuredValue
			end if
			if FedEX_active=true or FedEx_active="-1" then
				session("FedEXPackWidth"&intPackageCnt)=FEDEX_WIDTH
				session("FedEXPackHeight"&intPackageCnt)=FEDEX_HEIGHT
				session("FedEXPackLength"&intPackageCnt)=FEDEX_LENGTH
				session("FedEXPackWeight"&intPackageCnt)=intLastPackageWeight
				session("FedEXPackPrice"&intPackageCnt)=pcv_TempInsuredValue_FDX
			end if
			if FedEXWS_active=true or FedExWS_active="-1" then
				session("FedEXWSPackWidth"&intPackageCnt)=FEDEXWS_WIDTH
				session("FedEXWSPackHeight"&intPackageCnt)=FEDEXWS_HEIGHT
				session("FedEXWSPackLength"&intPackageCnt)=FEDEXWS_LENGTH
				session("FedEXWSPackWeight"&intPackageCnt)=intLastPackageWeight
				session("FedEXWSPackPrice"&intPackageCnt)=pcv_TempInsuredValue_FDXWS
			end if
			If CP_active=true or CP_active="-1" then
				session("CPPackWidth"&intPackageCnt)=CP_Width
				session("CPPackHeight"&intPackageCnt)=CP_Height
				session("CPPackLength"&intPackageCnt)=CP_Length
				session("CPPackWeight"&intPackageCnt)=intLastPackageWeight
			end if
			if USPS_active=true or USPS_active="-1" then
				session("USPSPackWidth"&intPackageCnt)=USPS_WIDTH
				session("USPSPackHeight"&intPackageCnt)=USPS_HEIGHT
				session("USPSPackLength"&intPackageCnt)=USPS_LENGTH
				session("USPSPackPounds"&intPackageCnt)=intLastPackageWeight
				session("USPSPackOunces"&intPackageCnt)=intUniversalOunces
			end if
			session("BasicPackPounds"&intPackageCnt)=intPounds
			session("BasicPackOunces"&intPackageCnt)=intUniversalOunces
		else 'There are more package after OS
			intPackageCnt=intPackageCnt+1
			strPackageInfo=strPackageInfo&"<tr><td>Package "&intPackageCnt&"</td>"
			strPackageInfo=strPackageInfo&"<td>Weight "&intPounds&"lb. "&intUniversalOunces&"oz.</td></tr>"
			pcv_intTotPackageNum=pcv_intTotPackageNum+1
			if ups_active=true or ups_active="-1" then
				session("UPSPackWidth"&intPackageCnt)=UPS_WIDTH
				session("UPSPackHeight"&intPackageCnt)=UPS_HEIGHT
				session("UPSPackLength"&intPackageCnt)=UPS_LENGTH
				session("UPSPackWeight"&intPackageCnt)=intUniversalWeight
				session("UPSPackPrice"&intPackageCnt)=cdbl(pcv_InsuredValue)
			end if
			if FedEX_active=true or FedEx_active="-1" then
				session("FedEXPackWidth"&intPackageCnt)=FEDEX_WIDTH
				session("FedEXPackHeight"&intPackageCnt)=FEDEX_HEIGHT
				session("FedEXPackLength"&intPackageCnt)=FEDEX_LENGTH
				session("FedEXPackWeight"&intPackageCnt)=intUniversalWeight
				session("FedEXPackPrice"&intPackageCnt)=cdbl(pcv_InsuredValue_FDX)
			end if
			if FedEXWS_active=true or FedExWS_active="-1" then
				session("FedEXWSPackWidth"&intPackageCnt)=FEDEXWS_WIDTH
				session("FedEXWSPackHeight"&intPackageCnt)=FEDEXWS_HEIGHT
				session("FedEXWSPackLength"&intPackageCnt)=FEDEXWS_LENGTH
				session("FedEXWSPackWeight"&intPackageCnt)=intUniversalWeight
				session("FedEXWSPackPrice"&intPackageCnt)=cdbl(pcv_InsuredValue_FDXWS)
			end if
			If CP_active=true or CP_active="-1" then
				session("CPPackWidth"&intPackageCnt)=CP_Width
				session("CPPackHeight"&intPackageCnt)=CP_Height
				session("CPPackLength"&intPackageCnt)=CP_Length
				session("CPPackWeight"&intPackageCnt)=intUniversalWeight
			end if
			if USPS_active=true or USPS_active="-1" then
				session("USPSPackWidth"&intPackageCnt)=USPS_WIDTH
				session("USPSPackHeight"&intPackageCnt)=USPS_HEIGHT
				session("USPSPackLength"&intPackageCnt)=USPS_LENGTH
				session("USPSPackPounds"&intPackageCnt)=intPounds
				session("USPSPackOunces"&intPackageCnt)=intUniversalOunces
				session("UPSPackPrice"&intPackageCnt)=cdbl(pcv_InsuredValue)
				if pcv_EMWeightLimit<>0 AND pShipWeight>Clng((pcv_EMWeightLimit*16)) then
					pcv_EM_Null=1
				end if
				if pcv_PMWeightLimit<>0 AND pShipWeight>Clng((pcv_PMWeightLimit*16)) then
					pcv_PM_Null=1
				end if
			end if
			session("BasicPackPounds"&intPackageCnt)=intPounds
			session("BasicPackOunces"&intPackageCnt)=intUniversalOunces
		end if 'There are more package after OS
	else 'There is a package Weight limit set
		'no weight limit set
		intPackageCnt=intPackageCnt+1
		strPackageInfo=strPackageInfo&"<tr><td>Package "&intPackageCnt&"</td>"
		strPackageInfo=strPackageInfo&"<td>Weight "&intPounds&"lb. "&intUniversalOunces&"oz.</td></tr>"
		pcv_intTotPackageNum=pcv_intTotPackageNum+1
		if ups_active=true or ups_active="-1" then
			session("UPSPackWidth"&intPackageCnt)=UPS_WIDTH
			session("UPSPackHeight"&intPackageCnt)=UPS_HEIGHT
			session("UPSPackLength"&intPackageCnt)=UPS_LENGTH
			session("UPSPackWeight"&intPackageCnt)=intUniversalWeight
			session("UPSPackPrice"&intPackageCnt)=cdbl(pcv_InsuredValue)
		end if
		if FedEX_active=true or FedEx_active="-1" then
			session("FedEXPackWidth"&intPackageCnt)=FEDEX_WIDTH
			session("FedEXPackHeight"&intPackageCnt)=FEDEX_HEIGHT
			session("FedEXPackLength"&intPackageCnt)=FEDEX_LENGTH
			session("FedEXPackWeight"&intPackageCnt)=intUniversalWeight
			session("FedEXPackPrice"&intPackageCnt)=cdbl(pcv_InsuredValue_FDX)
		end if
		if FedEXWS_active=true or FedExWS_active="-1" then
			session("FedEXWSPackWidth"&intPackageCnt)=FEDEXWS_WIDTH
			session("FedEXWSPackHeight"&intPackageCnt)=FEDEXWS_HEIGHT
			session("FedEXWSPackLength"&intPackageCnt)=FEDEXWS_LENGTH
			session("FedEXWSPackWeight"&intPackageCnt)=intUniversalWeight
			session("FedEXWSPackPrice"&intPackageCnt)=cdbl(pcv_InsuredValue_FDXWS)
		end if
		If CP_active=true or CP_active="-1" then
			session("CPPackWidth"&intPackageCnt)=CP_Width
			session("CPPackHeight"&intPackageCnt)=CP_Height
			session("CPPackLength"&intPackageCnt)=CP_Length
			session("CPPackWeight"&intPackageCnt)=intUniversalWeight
		end if
		if USPS_active=true or USPS_active="-1" then
			session("USPSPackWidth"&intPackageCnt)=USPS_WIDTH
			session("USPSPackHeight"&intPackageCnt)=USPS_HEIGHT
			session("USPSPackLength"&intPackageCnt)=USPS_LENGTH
			session("USPSPackPounds"&intPackageCnt)=intPounds
			session("USPSPackOunces"&intPackageCnt)=intUniversalOunces
			if pcv_EMWeightLimit<>0 AND pShipWeight>Clng((pcv_EMWeightLimit*16)) then
				pcv_EM_Null=1
			end if
			if pcv_PMWeightLimit<>0 AND pShipWeight>Clng((pcv_PMWeightLimit*16)) then
				pcv_PM_Null=1
			end if
		end if
		session("BasicPackPounds"&intPackageCnt)=intPounds
		session("BasicPackOunces"&intPackageCnt)=intUniversalOunces
	end if '// There is a package Weight limit set
end if '// Weight > 0
'=====================================================================

pcv_intPackageNum=intPackageCnt

'string
availableShipStr=""
dim iUPSActive, iFedExActive, iFedExWSActive, iUSPSActive, iCPActive
iUPSActive=0
iFedExActive=0
iFedExWSActive=0
iUSPSActive=0
iCPActive=0
UPS_ShipFromCity = scShipFromCity
UPS_ShipFromState = scShipFromState
UPS_ShipFromPostalCode = scShipFromPostalCode
UPS_ShipFromPostalCountry = scShipFromPostalCountry



'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' START: FEDEX RATES
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Dim pcv_strAccountName, pcv_strMeterNumber, pcv_strCarrierCode, pcv_strMethodName, pcv_strMethodReply
'Dim fedex_postdata, objFedExClass, objOutputXMLDoc, srvFEDEXXmlHttp, FEDEX_result, FEDEX_URL, pcv_strErrorMsg

pcv_strMethodName = "FDXRateAvailableServicesRequest"
pcv_strMethodReply = "FDXRateAvailableServicesReply"
CustomerTransactionIdentifier = "ProductCart_Rates"
pcv_strEnvironment = FEDEX_Environment

if (FedEX_active=true or FedEx_active="-1") AND FedEX_AccountNumber<>"" then

	iFedExActive=1
	dim arryFedExService
	dim arryFedExRate
	dim arrFedExDeliveryDate
	arryFedExService=""
	arryFedExRate=""
	arrFedExDeliveryDate=""

	pcv_TmpListRate = FEDEX_LISTRATE

	'// Override List Rates for International addresses
	If Universal_destination_country<>"US" Then
		pcv_TmpListRate = "0"
	End If

	'// FedEx EXPRESS RATES
	set objFedExClass = nothing
	for q=1 to pcv_intPackageNum

		set objFedExClass = New pcFedExClass

		'// Break Point
		'BreakPoint logFilename, "Break Point | pcPay_GoogleCheckout_ShipRates | Just Before FedEx - Line 343", "", err.description

		fedex_postdata=""
		FEDEX_result=""

		objFedExClass.NewXMLTransaction pcv_strMethodName, FedEX_AccountNumber, FedEX_MeterNumber, "FDXE", CustomerTransactionIdentifier

			objFedExClass.WriteSingleParent "ReturnShipmentIndicator", "NONRETURN"
			objFedExClass.WriteSingleParent "DropoffType", FEDEX_DROPOFF_TYPE
			objFedExClass.WriteSingleParent "Packaging", FEDEX_FEDEX_PACKAGE
			objFedExClass.WriteSingleParent "WeightUnits", scShipFromWeightUnit
			objFedExClass.WriteSingleParent "Weight", session("FedEXPackWeight"&q) & ".0"
			if pcv_TmpListRate = "-1" OR pcv_TmpListRate = -1 then
				objFedExClass.WriteSingleParent "ListRate", 1
			end if
			'// Origin Address
			objFedExClass.WriteParent "OriginAddress", ""
				objFedExClass.AddNewNode "StateOrProvinceCode", scShipFromState
				objFedExClass.AddNewNode "PostalCode", scShipFromPostalCode
				objFedExClass.AddNewNode "CountryCode", scShipFromPostalCountry
			objFedExClass.WriteParent "OriginAddress", "/"

			'// Destination Address
			objFedExClass.WriteParent "DestinationAddress", ""
				if Universal_destination_country="US" OR Universal_destination_country="CA" then
					objFedExClass.AddNewNode "StateOrProvinceCode", Universal_destination_provOrState
				end if
				objFedExClass.AddNewNode "PostalCode", Universal_destination_postal
				objFedExClass.AddNewNode "CountryCode", Universal_destination_country
			objFedExClass.WriteParent "DestinationAddress", "/"

			'// Payment Type
			objFedExClass.WriteParent "Payment", ""
				objFedExClass.AddNewNode "PayorType", "SENDER"
			objFedExClass.WriteParent "Payment", "/"

			'// Dims
			if ((FEDEX_FEDEX_PACKAGE="YOURPACKAGING") AND (session("FedEXPackLength"&q)<>"" AND session("FedEXPackWidth"&q)<>"" AND session("FedEXPackHeight"&q)<>"")) then
				pcv_strDimUnit = FEDEX_DIM_UNIT
				if pcv_strDimUnit="" then
					pcv_strDimUnit = "IN"
				end if
				objFedExClass.WriteParent "Dimensions", ""
					objFedExClass.AddNewNode "Length", Int(session("FedEXPackLength"&q))
					objFedExClass.AddNewNode "Width", Int(session("FedEXPackWidth"&q))
					objFedExClass.AddNewNode "Height", Int(session("FedEXPackHeight"&q))
					objFedExClass.AddNewNode "Units", FEDEX_DIM_UNIT
				objFedExClass.WriteParent "Dimensions", "/"
			end if

			objFedExClass.WriteParent "DeclaredValue", ""
				objFedExClass.AddNewNode "Value", pcv_InsuredValue
				objFedExClass.AddNewNode "CurrencyCode", "USD"
			objFedExClass.WriteParent "DeclaredValue", "/"

			if pResidentialShipping="-1" or pResidentialShipping="1" then
				objFedExClass.WriteParent "SpecialServices", ""
					objFedExClass.AddNewNode "ResidentialDelivery", 1 '// 1 or 0, should come from variable
				objFedExClass.WriteParent "SpecialServices", "/"
			end if

			objFedExClass.AddNewNode "PackageCount", "1"

		objFedExClass.EndXMLTransaction pcv_strMethodName
		'fedex_postdata=Session("fedex_postdata")
		' Get the URL to post to

		'// Print out our newly formed request xml
		'response.write fedex_postdata
		'response.end

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Send Our Transaction.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		call objFedExClass.SendXMLRequest(fedex_postdata, pcv_strEnvironment)
		'FEDEX_result=Session("FEDEX_result")
		'// Print out our response
		'response.write FEDEX_result
		'response.end

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Load Our Response.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		call objFedExClass.LoadXMLResults(FEDEX_result)
		'logMessage logFilename, "fedex out " & FEDEX_result & " > " & err.description

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Check for errors from FedEx.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		pcv_strErrorMsg = objFedExClass.ReadResponseNode("//Error", "Message")

		'/////////////////////////////////////////////////////////////
		'// BASELINE LOGGING
		'/////////////////////////////////////////////////////////////
		'// Log our Transaction
		'call objFedExClass.pcs_LogTransaction(fedex_postdata, pcv_strMethodName&"_"&q&".in", true)
		'// Log our Response
		'call objFedExClass.pcs_LogTransaction(FEDEX_result, pcv_strMethodName&"_"&q&".out", true)

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Redirect with a Message OR complete some task.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		If NOT len(pcv_strErrorMsg)>0 Then
			'// Generate FedEx Arrays
			if pcv_TmpListRate = "-1" OR pcv_TmpListRate = -1 then
				arryFedExService = arryFedExService & objFedExClass.ReadResponseasArray("//Entry", "Service")
				arrFedExDeliveryDate = arrFedExDeliveryDate & objFedExClass.ReadResponseasArray("//Entry", "DeliveryDate")
				arryFedExRate = arryFedExRate & objFedExClass.ReadResponseasArray("//ListCharges", "NetCharge")
			else
				arryFedExService = arryFedExService & objFedExClass.ReadResponseasArray("//Entry", "Service")
				arrFedExDeliveryDate = arrFedExDeliveryDate & objFedExClass.ReadResponseasArray("//Entry", "DeliveryDate")
				arryFedExRate = arryFedExRate & objFedExClass.ReadResponseasArray("//DiscountedCharges", "NetCharge")
			end if
			'response.Write(arrFedExDeliveryDate)
			'response.end
		End If
		set objFedExClass = nothing
	next


	for q=1 to pcv_intPackageNum

		set objFedExClass = New pcFedExClass

		fedex_postdata=""
		FEDEX_result=""

		objFedExClass.NewXMLTransaction pcv_strMethodName, FedEX_AccountNumber, FedEX_MeterNumber, "FDXG", CustomerTransactionIdentifier

		objFedExClass.WriteSingleParent "ReturnShipmentIndicator", "NONRETURN"
		objFedExClass.WriteSingleParent "DropoffType", FEDEX_DROPOFF_TYPE
		objFedExClass.WriteSingleParent "Packaging", "YOURPACKAGING"
		objFedExClass.WriteSingleParent "WeightUnits", scShipFromWeightUnit
		objFedExClass.WriteSingleParent "Weight", session("FedEXPackWeight"&q) & ".0"
		if pcv_TmpListRate = "-1" OR pcv_TmpListRate = -1 then
			objFedExClass.WriteSingleParent "ListRate", 1
		end if

		'// Origin Address
		objFedExClass.WriteParent "OriginAddress", ""
			objFedExClass.AddNewNode "StateOrProvinceCode", scShipFromState
			objFedExClass.AddNewNode "PostalCode", scShipFromPostalCode
			objFedExClass.AddNewNode "CountryCode", scShipFromPostalCountry
		objFedExClass.WriteParent "OriginAddress", "/"

		'// Destination Address
		objFedExClass.WriteParent "DestinationAddress", ""
			if Universal_destination_country="US" OR Universal_destination_country="CA" then
				objFedExClass.AddNewNode "StateOrProvinceCode", Universal_destination_provOrState
			end if
			objFedExClass.AddNewNode "PostalCode", Universal_destination_postal
			objFedExClass.AddNewNode "CountryCode", Universal_destination_country
		objFedExClass.WriteParent "DestinationAddress", "/"

		'// Payment Type
		objFedExClass.WriteParent "Payment", ""
			objFedExClass.AddNewNode "PayorType", "SENDER"
		objFedExClass.WriteParent "Payment", "/"

		'// Dims
		if ((FEDEX_FEDEX_PACKAGE="YOURPACKAGING") AND (session("FedEXPackLength"&q)<>"" AND session("FedEXPackWidth"&q)<>"" AND session("FedEXPackHeight"&q)<>"")) then
			pcv_strDimUnit = FEDEX_DIM_UNIT
			if pcv_strDimUnit="" then
				pcv_strDimUnit = "IN"
			end if
			objFedExClass.WriteParent "Dimensions", ""
				objFedExClass.AddNewNode "Length", Int(session("FedEXPackLength"&q))
				objFedExClass.AddNewNode "Width", Int(session("FedEXPackWidth"&q))
				objFedExClass.AddNewNode "Height", Int(session("FedEXPackHeight"&q))
				objFedExClass.AddNewNode "Units", pcv_strDimUnit
			objFedExClass.WriteParent "Dimensions", "/"
		end if

		objFedExClass.WriteParent "DeclaredValue", ""
			objFedExClass.AddNewNode "Value", pcv_InsuredValue
			objFedExClass.AddNewNode "CurrencyCode", "USD"
		objFedExClass.WriteParent "DeclaredValue", "/"

		if pResidentialShipping="-1" or pResidentialShipping="1" then
			objFedExClass.WriteParent "SpecialServices", ""
				objFedExClass.AddNewNode "ResidentialDelivery", 1 '// 1 or 0, should come from variable
			objFedExClass.WriteParent "SpecialServices", "/"
		end if

		'objFedExClass.WriteParent "HomeDelivery", ""
		'	objFedExClass.AddNewNode "Type", "EVENING"
		'objFedExClass.WriteParent "HomeDelivery", "/"

		objFedExClass.AddNewNode "PackageCount", "1"

		objFedExClass.EndXMLTransaction pcv_strMethodName

		' Get the URL to post to

		'// Print out our newly formed request xml
		'response.write fedex_postdata
		'response.end

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Send Our Transaction.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		call objFedExClass.SendXMLRequest(fedex_postdata, pcv_strEnvironment)
		'// Print out our response

		'response.write FEDEX_result
		'response.end

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Load Our Response.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		call objFedExClass.LoadXMLResults(FEDEX_result)


		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Check for errors from FedEx.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		pcv_strErrorMsg = objFedExClass.ReadResponseNode("//Error", "Message")


		'/////////////////////////////////////////////////////////////
		'// BASELINE LOGGING
		'/////////////////////////////////////////////////////////////
		'// Log our Transaction
		'call objFedExClass.pcs_LogTransaction(fedex_postdata, pcv_strMethodName&"_G"&q&".in", true)
		'// Log our Response
		'call objFedExClass.pcs_LogTransaction(FEDEX_result, pcv_strMethodName&"_G"&q&".out", true)

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Redirect with a Message OR complete some task.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		If NOT len(pcv_strErrorMsg)>0 Then
			'// Generate FedEx Arrays
			if pcv_TmpListRate = "-1" OR pcv_TmpListRate = -1 then
				arryFedExService = arryFedExService & objFedExClass.ReadResponseasArray("//Entry", "Service")
				arrFedExDeliveryDate = arrFedExDeliveryDate & objFedExClass.ReadResponseasArray("//Entry", "DeliveryDate")
				arryFedExRate = arryFedExRate & objFedExClass.ReadResponseasArray("//ListCharges", "NetCharge")
			else
				arryFedExService = arryFedExService & objFedExClass.ReadResponseasArray("//Entry", "Service")
				arrFedExDeliveryDate = arrFedExDeliveryDate & objFedExClass.ReadResponseasArray("//Entry", "DeliveryDate")
				arryFedExRate = arryFedExRate & objFedExClass.ReadResponseasArray("//DiscountedCharges", "NetCharge")
			end if
			'response.Write(arrFedExDeliveryDate)
			'response.end
		End If

		set objFedExClass = nothing

	next

	' trim the last comma if there is one
	'xStringLength = len(ReadResponseasArray)
	'if xStringLength>0 then
	'	ReadResponseasArray = left(ReadResponseasArray,(xStringLength-1))
	'end if

	'Split Arrays
	dim intRateIndex
	if isArray(pcFedExMultiArry) <> True then
		dim pcFedExMultiArry(14,4)
	end if
	for z=0 to 14
		pcFedExMultiArry(z,1)=0
	next

	pcStrTempFedExService=split(arryFedExService,",")
	pcStrTempFexExRate=split(arryFedExRate,",")
	pcStrTempFedExDeliveryDate=split(arrFedExDeliveryDate,",")

	for t=0 to (ubound(pcStrTempFedExService)-1)
		select case pcStrTempFedExService(t)
			case "PRIORITYOVERNIGHT"
				intRateIndex=0
				pcFedExMultiArry(intRateIndex,2)="FedEx Priority Overnight<sup>&reg;</sup>"
				pcFedExMultiArry(intRateIndex,3)="PRIORITYOVERNIGHT"
				pcFedExMultiArry(intRateIndex,4)=pcStrTempFedExDeliveryDate(t)
			case "STANDARDOVERNIGHT"
				intRateIndex=1
				pcFedExMultiArry(intRateIndex,2)="FedEx Standard Overnight<sup>&reg;</sup>"
				pcFedExMultiArry(intRateIndex,3)="STANDARDOVERNIGHT"
				pcFedExMultiArry(intRateIndex,4)=pcStrTempFedExDeliveryDate(t)
			case "FIRSTOVERNIGHT"
				intRateIndex=2
				pcFedExMultiArry(intRateIndex,2)="FedEx First Overnight<sup>&reg;</sup>"
				pcFedExMultiArry(intRateIndex,3)="FIRSTOVERNIGHT"
				pcFedExMultiArry(intRateIndex,4)=pcStrTempFedExDeliveryDate(t)
			case "FEDEX2DAY"
				intRateIndex=3
				pcFedExMultiArry(intRateIndex,2)="FedEx 2Day<sup>&reg;</sup>"
				pcFedExMultiArry(intRateIndex,3)="FEDEX2DAY"
				pcFedExMultiArry(intRateIndex,4)=pcStrTempFedExDeliveryDate(t)
			case "FEDEXEXPRESSSAVER"
				intRateIndex=4
				pcFedExMultiArry(intRateIndex,2)="FedEx Express Saver<sup>&reg;</sup>"
				pcFedExMultiArry(intRateIndex,3)="FEDEXEXPRESSSAVER"
				pcFedExMultiArry(intRateIndex,4)=pcStrTempFedExDeliveryDate(t)
			case "INTERNATIONALPRIORITY"
				intRateIndex=5
				pcFedExMultiArry(intRateIndex,2)="FedEx International Priority<sup>&reg;</sup>"
				pcFedExMultiArry(intRateIndex,3)="INTERNATIONALPRIORITY"
				pcFedExMultiArry(intRateIndex,4)=pcStrTempFedExDeliveryDate(t)
			case "INTERNATIONALECONOMY"
				intRateIndex=6
				pcFedExMultiArry(intRateIndex,2)="FedEx International Economy<sup>&reg;</sup>"
				pcFedExMultiArry(intRateIndex,3)="INTERNATIONALECONOMY"
				pcFedExMultiArry(intRateIndex,4)=pcStrTempFedExDeliveryDate(t)
			case "INTERNATIONALFIRST"
				intRateIndex=7
				pcFedExMultiArry(intRateIndex,2)="FedEx International First<sup>&reg;</sup>"
				pcFedExMultiArry(intRateIndex,3)="INTERNATIONALFIRST"
				pcFedExMultiArry(intRateIndex,4)=pcStrTempFedExDeliveryDate(t)
			case "FEDEX1DAYFREIGHT"
				intRateIndex=8
				pcFedExMultiArry(intRateIndex,2)="FedEx 1Day<sup>&reg;</sup> Freight"
				pcFedExMultiArry(intRateIndex,3)="FEDEX1DAYFREIGHT"
				pcFedExMultiArry(intRateIndex,4)=pcStrTempFedExDeliveryDate(t)
			case "FEDEX2DAYFREIGHT"
				intRateIndex=9
				pcFedExMultiArry(intRateIndex,2)="FedEx 2Day<sup>&reg;</sup> Freight"
				pcFedExMultiArry(intRateIndex,3)="FEDEX2DAYFREIGHT"
				pcFedExMultiArry(intRateIndex,4)=pcStrTempFedExDeliveryDate(t)
			case "FEDEX3DAYFREIGHT"
				intRateIndex=10
				pcFedExMultiArry(intRateIndex,2)="FedEx 3Day<sup>&reg;</sup> Freight"
				pcFedExMultiArry(intRateIndex,3)="FEDEX3DAYFREIGHT"
				pcFedExMultiArry(intRateIndex,4)=pcStrTempFedExDeliveryDate(t)
			case "FEDEXGROUND"
				intRateIndex=11
				pcFedExMultiArry(intRateIndex,2)="FedEx Ground<sup>&reg;</sup>"
				pcFedExMultiArry(intRateIndex,3)="FEDEXGROUND"
				pcFedExMultiArry(intRateIndex,4)=pcStrTempFedExDeliveryDate(t)
			case "GROUNDHOMEDELIVERY"
				intRateIndex=12
				pcFedExMultiArry(intRateIndex,2)="FedEx Home Delivery<sup>&reg;</sup>"
				pcFedExMultiArry(intRateIndex,3)="GROUNDHOMEDELIVERY"
				pcFedExMultiArry(intRateIndex,4)=pcStrTempFedExDeliveryDate(t)
			case "INTERNATIONALPRIORITYFREIGHT"
				intRateIndex=13
				pcFedExMultiArry(intRateIndex,2)="FedEx International Priority<sup>&reg;</sup> Freight"
				pcFedExMultiArry(intRateIndex,3)="INTERNATIONALPRIORITYFREIGHT"
				pcFedExMultiArry(intRateIndex,4)=pcStrTempFedExDeliveryDate(t)
			case "INTERNATIONALECONOMYFREIGHT"
				intRateIndex=14
				pcFedExMultiArry(intRateIndex,2)="FedEx International Economy<sup>&reg;</sup> Freight"
				pcFedExMultiArry(intRateIndex,3)="INTERNATIONALECONOMYFREIGHT"
				pcFedExMultiArry(intRateIndex,4)=pcStrTempFedExDeliveryDate(t)
		end select
		tempRate=pcFedExMultiArry(intRateIndex,1)
		pcFedExMultiArry(intRateIndex,1)=cdbl(tempRate)+cdbl(pcStrTempFexExRate(t))
	next

	for z=0 to 14
		if pcFedExMultiArry(z,1)>0 then
			pcv_strFormattedDate = ""
			pcv_strFormattedDate = pcFedExMultiArry(z,4)
			if pcv_strFormattedDate = "" then
				pcv_strFormattedDate="NA"
			else
				pcv_strFormattedDate = pcv_strFormattedDate 'showdateFrmt(pcv_strFormattedDate)
			end if
			if pcFedExMultiArry(z,3)="FEDEXGROUND" OR pcFedExMultiArry(z,3)="GROUNDHOMEDELIVERY" then
				if pResidentialShipping="-1" then pResidentialShipping="1"
				if pResidentialShipping="1" AND pcFedExMultiArry(z,3)="GROUNDHOMEDELIVERY" then
					availableShipStr=availableShipStr&"|?|FedEX|"&pcFedExMultiArry(z,3)&"|"&pcFedExMultiArry(z,2)&"|"&pcFedExMultiArry(z,1)&"|"&pcv_strFormattedDate
				iFedExFlag=1
				end if
				if pResidentialShipping="0" AND pcFedExMultiArry(z,3)="FEDEXGROUND" then
					availableShipStr=availableShipStr&"|?|FedEX|"&pcFedExMultiArry(z,3)&"|"&pcFedExMultiArry(z,2)&"|"&pcFedExMultiArry(z,1)&"|"&pcv_strFormattedDate
				iFedExFlag=1
				end if
			else
				availableShipStr=availableShipStr&"|?|FedEX|"&pcFedExMultiArry(z,3)&"|"&pcFedExMultiArry(z,2)&"|"&pcFedExMultiArry(z,1)&"|"&pcv_strFormattedDate
				iFedExFlag=1
			end if
		end if
	next
end if 'if fedex is active

'// Break Point
'BreakPoint logFilename, "Break Point | pcPay_GoogleCheckout_ShipRates | Just After FedEx - Line 713", "", err.description
err.clear
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' END: FEX RATES
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
if iFedExFlag=1 then
		strDefaultProvider="FEDEX"
		iShipmentTypeCnt=iShipmentTypeCnt+1
		strOptionShipmentType=strOptionShipmentType&"<option value=FedEx>"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_m")&"</option>"
end if
%>
<%
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' START: FEDEX WS RATES
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////

%>
<!--#include file="FedExWebServices.asp"-->
<%

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' END: FEDEX WS RATES
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
if iFedExWSFlag=1 then
	strDefaultProvider="FEDEXWS"
	iShipmentTypeCnt=iShipmentTypeCnt+1
	strOptionShipmentType=strOptionShipmentType&"<option value=FedExWS>"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_m")&"</option>"
end if
%>
<%
'//////////////////////////////////////////
'// START UPS RATES
'//////////////////////////////////////////
if ups_active=true or ups_active="-1" then
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
		if pcv_UseDynamicInsuredValue=1 then
			pcv_TempPackPrice=session("UPSPackPrice"&q)
		else
			pcv_TempPackPrice="100.00"
		end if
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
						availableShipStr=availableShipStr&"|?|UPS|01|"&"UPS Next Day Air&reg;"
						varFlag=1
						iUPSFlag=1
					case "02"
						availableShipStr=availableShipStr&"|?|UPS|02|"&"UPS 2nd Day Air&reg;"
						varFlag=1
						iUPSFlag=1
					case "03"
						availableShipStr=availableShipStr&"|?|UPS|03|"&"UPS Ground"
						varFlag=1
						iUPSFlag=1
					case "07"
						availableShipStr=availableShipStr&"|?|UPS|07|"&"UPS Worldwide Express<sup>SM</sup>"
						varFlag=1
						iUPSFlag=1
					case "08"
						availableShipStr=availableShipStr&"|?|UPS|08|"&"UPS Worldwide Expedited<sup>SM</sup>"
						varFlag=1
						iUPSFlag=1
					case "11"
						availableShipStr=availableShipStr&"|?|UPS|11|"&"UPS Standard To Canada"
						varFlag=1
						iUPSFlag=1
					case "12"
						availableShipStr=availableShipStr&"|?|UPS|12|"&"UPS 3 Day Select<sup>SM</sup>"
						varFlag=1
						iUPSFlag=1
					case "13"
						availableShipStr=availableShipStr&"|?|UPS|13|"&"UPS Next Day Air Saver&reg;"
						varFlag=1
						iUPSFlag=1
					case "14"
						availableShipStr=availableShipStr&"|?|UPS|14|"&"UPS Next Day Air&reg; Early A.M.&reg;"
						varFlag=1
						iUPSFlag=1
					case "54"
						availableShipStr=availableShipStr&"|?|UPS|54|"&"UPS Worldwide Express Plus<sup>SM</sup>"
						varFlag=1
						iUPSFlag=1
					case "59"
						availableShipStr=availableShipStr&"|?|UPS|59|"&"UPS 2nd Day Air A.M.&reg;"
						varFlag=1
						iUPSFlag=1
					case "65"
						availableShipStr=availableShipStr&"|?|UPS|65|"&"UPS Express Saver<sup>SM</sup>"
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
'//////////////////////////////////////////
'// END UPS RATES
'//////////////////////////////////////////
%>
<%
if iUPSFlag=1 then
	strDefaultProvider="UPS"
	iShipmentTypeCnt=iShipmentTypeCnt+1
	strOptionShipmentType=strOptionShipmentType&"<option value=UPS>"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_k")&"</option>"
end if

If usps_active=true or usps_active="-1" then
	iUSPSActive=1
	Dim USPS_PackageSize
	'check that all packages can go through USPS
	dim intUSPSnoShpmnt
	intUSPSnoShpmnt=0
	session("BMP")="Y"
	USPS_DWeightOver =""
	USPS_PackageSizeC=""
	for q=1 to pcv_intPackageNum
		'// If any one side is greater then 12" package is labeled as "LARGE"
		If USPS_PackageSizeC="" AND ((Cint(session("USPSPackLength"&q))>12) OR (Cint(session("USPSPackWidth"&q))>12) OR (Cint(session("USPSPackHeight"&q))>12)) Then
			USPS_PackageSizeC="LARGE"
		Else
			USPS_PackageSize=(Cint(session("USPSPackLength"&q)) + ((Cint(session("USPSPackWidth"&q))*2)+(Cint(session("USPSPackHeight"&q))*2)))
			If USPS_PackageSize<85 then
				USPS_PackageSizeC="REGULAR"
			End if
			If USPS_PackageSize>85 AND USPS_PackageSize<108 AND intPounds<15 then
				USPS_PackageSizeC="LARGE"
				if (Cint(Cint(session("USPSPackLength"&q))) * (Cint(Cint(session("USPSPackWidth"&q))) * Cint(Cint(session("USPSPackHeight"&q)))))>1728 then
					USPS_DWeightOver = "YES"
				end if
			End if
			If USPS_PackageSizeC="" AND USPS_PackageSize<131 then
				USPS_PackageSizeC="OVERSIZE"
			End if
			if USPS_PackageSizeC="" OR USPS_PackageSizeC="REGULAR" then
				if (Cint(Cint(session("USPSPackLength"&q))) * (Cint(Cint(session("USPSPackWidth"&q))) * Cint(Cint(session("USPSPackHeight"&q)))))>1728 then
					USPS_PackageSizeC="LARGE"
					USPS_DWeightOver = "YES"
				end if
			end if
		End If
	next
	If USPS_PackageSizeC="" then
		intUSPSnoShpmnt=1
	end if
	IF session("USPSPackPounds"&q)>15 THEN
		session("BMP")="N"
	end if

	If intUSPSnoShpmnt=0 then
		'//USPS RATES - Domestic
		If Universal_destination_country="US" then
			'parse +4 from the zip code
			if len(Universal_destination_postal)>5 then
				Universal_destination_postal=left(Universal_destination_postal,5)
			end if
			usps_postdata=""
			usps_postdata=usps_postdata&usps_server&"?API=RateV4&XML="

			usps_postdata=usps_postdata&"<RateV4Request%20USERID="&chr(34)&usps_userid&chr(34)&">"
			for q=1 to pcv_intPackageNum
				iNum=q-1
				USPS_PackageSizeC=""

				'// If any one side is greater then 12" package is labeled as "LARGE"
				If (Cint(session("USPSPackLength"&q))>12) OR (Cint(session("USPSPackWidth"&q))>12) OR (Cint(session("USPSPackHeight"&q))>12) Then
					USPS_PackageSizeC="LARGE"
					USPS_DWeightOver = "YES"
				Else
					USPS_PackageSize=(Cint(session("USPSPackLength"&q)) + ((Cint(session("USPSPackWidth"&q))*2)+(Cint(session("USPSPackHeight"&q))*2)))
					If USPS_PackageSize<85 then
						USPS_PackageSizeC="REGULAR"
					End if
					If USPS_PackageSize>85 AND USPS_PackageSize<108 then
						USPS_PackageSizeC="LARGE"
						if (Cint(Cint(session("USPSPackLength"&q))) * (Cint(Cint(session("USPSPackWidth"&q))) * Cint(Cint(session("USPSPackHeight"&q)))))>1728 then
							USPS_DWeightOver = "YES"
						end if
					End if
					If USPS_PackageSizeC="" AND USPS_PackageSize<131 then
						USPS_PackageSizeC="OVERSIZE"
					End if
					if USPS_PackageSizeC="" OR USPS_PackageSizeC="REGULAR" then
						if (Cint(Cint(session("USPSPackLength"&q))) * (Cint(Cint(session("USPSPackWidth"&q))) * Cint(Cint(session("USPSPackHeight"&q)))))>1728 then
							USPS_PackageSizeC="LARGE"
							USPS_DWeightOver = "YES"
						end if
					end if
				End If

				usps_postdata=usps_postdata&"<Package%20ID="&chr(34)&iNum&chr(34)&">"
				usps_postdata=usps_postdata&"<Service>All</Service>"
				usps_postdata=usps_postdata&"<ZipOrigination>"&scShipFromPostalCode&"</ZipOrigination>"
				usps_postdata=usps_postdata&"<ZipDestination>"&Universal_destination_postal&"</ZipDestination>"
				usps_postdata=usps_postdata&"<Pounds>"&session("USPSPackPounds"&q)&"</Pounds>"
				usps_postdata=usps_postdata&"<Ounces>"&round(session("USPSPackOunces"&q))&"</Ounces>"
				if USPS_DWeightOver = "YES" then
					usps_postdata=usps_postdata&"<Container>RECTANGULAR</Container>"
					usps_postdata=usps_postdata&"<Size>"&USPS_PackageSizeC&"</Size>"
					usps_postdata=usps_postdata&"<Width>"&Cint(session("USPSPackWidth"&q))&"</Width>"
					usps_postdata=usps_postdata&"<Length>"&Cint(session("USPSPackLength"&q))&"</Length>"
					usps_postdata=usps_postdata&"<Height>"&Cint(session("USPSPackHeight"&q))&"</Height>"
					usps_postdata=usps_postdata&"<Girth>"&USPS_PackageSize&"</Girth>"
				else
					usps_postdata=usps_postdata&"<Container>VARIABLE</Container>"
					usps_postdata=usps_postdata&"<Size>"&USPS_PackageSizeC&"</Size>"
				end if
				IF USPS_PackageSizeC="LARGE" THEN
					'Check if Machinable or not
					if Cint(session("USPSPackLength"&q))<3 OR Cint(session("USPSPackLength"&q))>34 OR Cint(session("USPSPackWidth"&q))<3 OR Cint(session("USPSPackWidth"&q))>17 OR Cint(session("USPSPackHeight"&q))>17 OR USPS_DWeightOver = "YES" then
						usps_postdata=usps_postdata&"<Machinable>False</Machinable>"
					else
						usps_postdata=usps_postdata&"<Machinable>TRUE</Machinable>"
					end if
				else
					usps_postdata=usps_postdata&"<Machinable>TRUE</Machinable>"
				END IF
				usps_postdata=usps_postdata&"</Package>"
			next
			usps_postdata=usps_postdata&"</RateV4Request>"
			err.clear
			Set srvUSPS2XmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
			srvUSPS2XmlHttp.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
			srvUSPS2XmlHttp.open "GET", usps_postdata, false
			srvUSPS2XmlHttp.send
			USPS2_result = srvUSPS2XmlHttp.responseText

			' Parse the XML document.
			Set objOutputXMLDoc = Server.CreateObject("Microsoft.XMLDOM")
			objOutputXMLDoc.loadXML srvUSPS2XmlHttp.responseText

			Set Nodes = objOutputXMLDoc.selectNodes("//Postage")

			iUSPSEMFlag=0
			iUSPSPMFlag=0
			iUSPSFCFlag=0
			iUSPSPPFlag=0
			iUSPSBPMFlag=0
			iUSPSMMFlag=0
			iUSPSLMFlag=0
			iUSPSEMRate=0
			iUSPSPMRate=0
			iUSPSFCRate=0
			iUSPSPPRate=0
			iUSPSBPMRate=0
			iUSPSMMRate=0
			iUSPSLMRate=0
			iUSPSEMCnt=0
			iUSPSPMCnt=0
			iUSPSFCCnt=0
			iUSPSPPCnt=0
			iUSPSBPMCnt=0
			iUSPSMMCnt=0
			iUSPSLMCnt=0
			iUSPSEMFlagAdded=0
			iUSPSPMFlagAdded=0
			iUSPSFCFlagAdded=0
			iUSPSPPFlagAdded=0
			iUSPSBPMFlagAdded=0
			iUSPSMMFlagAdded=0
			iUSPSLMFlagAdded=0

			USPSErrorDetect1=0

			set objLst=objOutputXMLDoc.getElementsByTagName("Package")
			for i = 0 to (objLst.length - 1)
				for j=0 to ((objLst.item(i).childNodes.length)-1)
					If objLst.item(i).childNodes(j).nodeName="Error" then
						USPSErrorDetect1=1
						for m=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="Number" then
								strUSPSError = objLst.item(i).childNodes(j).childNodes(m).text
							end if
						next
					End if
				next
			next

			for i = 0 to (objLst.length - 1)
				for j=0 to ((objLst.item(i).childNodes.length)-1)
					If objLst.item(i).childNodes(j).nodeName="Postage" then
						intCLASSID=objLst.item(i).childNodes(j).getAttribute("CLASSID")
						for m=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="MailService" then
								strMailService = objLst.item(i).childNodes(j).childNodes(m).text
							end if
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="Rate" then
								strRate = objLst.item(i).childNodes(j).childNodes(m).text
							end if

						next

						if USPSErrorDetect1=0 then

							select case intCLASSID

							case "1"
								if ucase(USPS_PM_PACKAGE)="NONE" OR (pcv_PM_Null=1 AND USPS_PM_FREOption="NONE") then
									iUSPSPMFlag=1
									iUSPSPMCnt=iUSPSPMCnt+1
									iUSPSFlag=1
									if isNumeric(strRate) then
										strRate=cdbl(strRate)
									end if
									iUSPSPMRate=iUSPSPMRate+strRate
								end if
							case "3"
								if ucase(USPS_EM_PACKAGE)="NONE" OR (pcv_EM_Null=1 AND USPS_EM_FREOption="1") then
									iUSPSEMFlag=1
									iUSPSEMCnt=iUSPSEMCnt+1
									iUSPSFlag=1
									if isNumeric(strRate) then
										strRate=cdbl(strRate)
									end if
									iUSPSEMRate=iUSPSEMRate+strRate
								end if
							case "0"
								If instr(strMailService, "Parcel") Then
									if iUSPSFCFlag=0 then
										iUSPSFCFlag=1
										iUSPSFCCnt=iUSPSFCCnt+1
										iUSPSFlag=1
										if isNumeric(strRate) then
											strRate=cdbl(strRate)
										end if
										iUSPSFCRate=iUSPSFCRate+strRate
									end if
								End If
							case "4"
									iUSPSPPFlag=1
									iUSPSPPCnt=iUSPSPPCnt+1
									iUSPSFlag=1
									if isNumeric(strRate) then
										strRate=cdbl(strRate)
									end if
									iUSPSPPRate=iUSPSPPRate+strRate
							case "Bound Printed Matter"
								if session("BMP")="Y" then
									iUSPSBPMFlag=1
									iUSPSBPMCnt=iUSPSBPMCnt+1
									iUSPSFlag=1
									if isNumeric(strRate) then
										strRate=cdbl(strRate)
									end if
									iUSPSBPMRate=iUSPSBPMRate+strRate
								end if
							case "6"
									iUSPSMMFlag=1
									iUSPSMMCnt=iUSPSMMCnt+1
									iUSPSFlag=1
									if isNumeric(strRate) then
										strRate=cdbl(strRate)
									end if
									iUSPSMMRate=iUSPSMMRate+strRate
							case "7"
									iUSPSLMFlag=1
									iUSPSLMCnt=iUSPSLMCnt+1
									iUSPSFlag=1
									if isNumeric(strRate) then
										strRate=cdbl(strRate)
									end if
									iUSPSLMRate=iUSPSLMRate+strRate
							end select
							'Priority Mail
							if iUSPSPMCnt<pcv_intPackageNum then
								iUSPSPMFlag=0
							end if
							if iUSPSPMFlag=1 AND iUSPSPMFlagAdded = 0 then
								availableShipStr=availableShipStr&"|?|USPS|9901|"&"Priority Mail <sup>&reg;</sup>|"&iUSPSPMRate&"|NA|"
								iUSPSPMFlagAdded = 1
							end if
							'Express Mail
							if iUSPSEMCnt<pcv_intPackageNum then
								iUSPSEMFlag=0
							end if
							if iUSPSEMFlag=1 AND iUSPSEMFlagAdded = 0 then
								availableShipStr=availableShipStr&"|?|USPS|9902|"&"Express Mail <sup>&reg;</sup>|"&iUSPSEMRate&"|NA|"
								iUSPSEMFlagAdded=1
							end if
							'First Class Mail
							if iUSPSFCCnt<pcv_intPackageNum then
								iUSPSFCFlag=0
							end if
							if iUSPSFCFlag=1  AND iUSPSFCFlagAdded = 0 then
								availableShipStr=availableShipStr&"|?|USPS|9904|"&"First-Class Mail <sup>&reg;</sup>|"&iUSPSFCRate&"|NA|"
								iUSPSFCFlagAdded =1
							end if
							'Standard Post
							if iUSPSPPCnt<pcv_intPackageNum then
								iUSPSPPFlag=0
							end if
							if iUSPSPPFlag=1  AND iUSPSPPFlagAdded = 0 then
								availableShipStr=availableShipStr&"|?|USPS|9903|"&"Standard Post <sup>&reg;</sup>|"&iUSPSPPRate&"|NA|"
								iUSPSPPFlagAdded = 1
							end if
							'Bound Printed Matter
							if iUSPSBPMCnt<pcv_intPackageNum then
								iUSPSBPMFlag=0
							end if
							if iUSPSBPMFlag=1  AND iUSPSBPMFlagAdded = 0 then
								availableShipStr=availableShipStr&"|?|USPS|9915|"&"Bound Printed Matter <sup>&reg;</sup>|"&iUSPSBPMRate&"|NA|"
								iUSPSBPMFlagAdded = 1
							end if
							'Media Mail
							if iUSPSMMCnt<pcv_intPackageNum then
								iUSPSMMFlag=0
							end if
							if iUSPSMMFlag=1  AND iUSPSMMFlagAdded = 0 then
								availableShipStr=availableShipStr&"|?|USPS|9916|"&"Media Mail <sup>&reg;</sup>|"&iUSPSMMRate&"|NA|"
								iUSPSMMFlagAdded = 1
							end if
							'Library Mail
							if iUSPSLMCnt<pcv_intPackageNum then
								iUSPSLMFlag=0
							end if
							if iUSPSLMFlag=1  AND iUSPSLMFlagAdded = 0 then
								availableShipStr=availableShipStr&"|?|USPS|9917|"&"Library Mail <sup>&reg;</sup>|"&iUSPSLMRate&"|NA|"
								iUSPSLMFlagAdded = 1
							end if

						end if
					End If
				Next
			Next


		end if

		'//USPS RATES - Express
		if ucase(USPS_EM_PACKAGE)<>"NONE" then

			'If envelope isn't valid due to weight, check if Your Package is allowed
			if (pcv_EM_Null=1 AND USPS_EM_FREOption="1") OR pcv_EM_Null=0 then

				If Universal_destination_country="US" then
					usps_postdata=""
					usps_postdata=usps_postdata&usps_server&"?API=RateV4&XML="
					usps_postdata=usps_postdata&"<RateV4Request%20USERID="&chr(34)&usps_userid&chr(34)&">"

					for q=1 to pcv_intPackageNum
						pcv_EM_Package=""
						iNum=q-1
						USPS_PackageSizeC=""
						'//If any one side is greater then 12" package is labeled as "LARGE"
						If (Cint(session("USPSPackLength"&q))>12) OR (Cint(session("USPSPackWidth"&q))>12) OR (Cint(session("USPSPackHeight"&q))>12) Then
							USPS_PackageSizeC="LARGE"
						Else
							USPS_PackageSize=(Cint(session("USPSPackLength"&q)) + ((Cint(session("USPSPackWidth"&q))*2)+(Cint(session("USPSPackHeight"&q))*2)))
							If USPS_PackageSize<85 then
								USPS_PackageSizeC="REGULAR"
							End if
							If USPS_PackageSize>85 AND USPS_PackageSize<108 AND intPounds<15 then
								USPS_PackageSizeC="LARGE"
							End if
							If USPS_PackageSizeC="" AND USPS_PackageSize<131 then
								USPS_PackageSizeC="OVERSIZE"
							End if
							pcv_EM_Package=USPS_EM_PACKAGE
							if USPS_PackageSizeC="LARGE" OR USPS_PackageSizeC="OVERSIZE" then
								pcv_EM_Package=""
							end if
						End If
						usps_postdata=usps_postdata&"<Package%20ID="&chr(34)&iNum&chr(34)&">"
						usps_postdata=usps_postdata&"<Service>Express</Service>"
						usps_postdata=usps_postdata&"<ZipOrigination>"&scShipFromPostalCode&"</ZipOrigination>"
						usps_postdata=usps_postdata&"<ZipDestination>"&Universal_destination_postal&"</ZipDestination>"
						usps_postdata=usps_postdata&"<Pounds>"&session("USPSPackPounds"&q)&"</Pounds>"
						usps_postdata=usps_postdata&"<Ounces>"&round(session("USPSPackOunces"&q))&"</Ounces>"
						'// If FRE is the default, check for weight limit and alternate container
						if pcv_EM_Null=1 AND USPS_EM_FREOption="1" then
							pcv_EM_Package="NONE"
						end if
						usps_postdata=usps_postdata&"<Container>"&pcv_EM_Package&"</Container>"
						usps_postdata=usps_postdata&"<Size>"&USPS_PackageSizeC&"</Size>"
						usps_postdata=usps_postdata&"</Package>"

					next

					usps_postdata=usps_postdata&"</RateV4Request>"

					Set srvUSPS2XmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
					srvUSPS2XmlHttp.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
					srvUSPS2XmlHttp.open "GET", usps_postdata, false
					srvUSPS2XmlHttp.send
					USPS2_result = srvUSPS2XmlHttp.responseText

					' Parse the XML document.
					err.clear
					Set objOutputXMLDoc = Server.CreateObject("Microsoft.XMLDOM")
					objOutputXMLDoc.loadXML USPS2_result

					intUSPSPostage=0
					session("EMService")=""

					set objLst=objOutputXMLDoc.getElementsByTagName("Package")
					for i = 0 to (objLst.length - 1)
						USPS_TempSize=""
						for j=0 to ((objLst.item(i).childNodes.length)-1)
							If objLst.item(i).childNodes(j).nodeName="Size" then
								USPS_TempSize=objLst.item(i).childNodes(j).Text
							End if
							If objLst.item(i).childNodes(j).nodeName="Postage" then
								for m=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
									If objLst.item(i).childNodes(j).childNodes(m).nodeName="MailService" then
										strMailService = objLst.item(i).childNodes(j).childNodes(m).text
									end if
									If objLst.item(i).childNodes(j).childNodes(m).nodeName="Rate" then
										strRate = objLst.item(i).childNodes(j).childNodes(m).text
									end if
								next
							End if
						next
						pcv_EM_MailService=""
						if instr(strMailService, "(") then
							arrMailService=split(strMailService,"(")
							strMailService=arrMailService(0)
						end if

						if instr(strMailService, "Express Mail") then
							strMailService = replace(strMailService,"&amp;lt;","<")
							strMailService = replace(strMailService,"&amp;gt;",">")
							strMailService = replace(strMailService,"&amp;","&")
							strMailService = replace(strMailService,"&lt;","<")
							strMailService = replace(strMailService,"&gt;",">")

							pcv_EM_MailService="USPS "&strMailService
							session("USPSshipStr")="|?|USPS|9902|"&"X|X|X|"
							if isNumeric(strRate) then
								strRate=cdbl(strRate)
							end if
							intUSPSPostage=intUSPSPostage+strRate
							iUSPSFlag=1
						end if

						if USPS_TempSize="LARGE" then
							session("EMService")="LARGE"
						end if
					next

					if session("EMService")="LARGE" then
						pcv_EM_MailService="USPS Express Mail <sup>&reg;</sup>"
					end if
					availableShipStr=availableShipStr&replace(session("USPSshipStr"),"X|X|X|", pcv_EM_MailService)&"|"&intUSPSPostage&"|NA|"
					session("USPSshipStr")=""

				end if
			end if
		end if

		'Priority Mail
		If Universal_destination_country="US" AND iUSPSPMFlag=0 then
			pcv_PMOption=USPS_PM_FREOption

			if isNull(pcv_PMOption) or pcv_PMOption="" then
				pcv_PMOption="0"
			end if

			if (pcv_PM_Null=1 AND pcv_PMOption<>"0" AND pcv_PMOption<>"NONE") OR pcv_PM_Null=0 then

				usps_postdata=""
				usps_postdata=usps_postdata&usps_server&"?API=RateV4&XML="
				usps_postdata=usps_postdata&"<RateV4Request%20USERID="&chr(34)&usps_userid&chr(34)&">"
				for q=1 to pcv_intPackageNum
					iNum=q-1
					USPS_PackageSizeC=""
					'//If any one side is greater then 12" package is labeled as "LARGE"
					If (Cint(session("USPSPackLength"&q))>12) OR (Cint(session("USPSPackWidth"&q))>12) OR (Cint(session("USPSPackHeight"&q))>12) Then
						USPS_PackageSizeC="LARGE"
					Else
						USPS_PackageSize=(Cint(session("USPSPackLength"&q)) + ((Cint(session("USPSPackWidth"&q))*2)+(Cint(session("USPSPackHeight"&q))*2)))
						If USPS_PackageSize<85 then
							USPS_PackageSizeC="REGULAR"
						End if
						If USPS_PackageSize>85 AND USPS_PackageSize<108 AND intPounds<15 then
							USPS_PackageSizeC="LARGE"
						End if
						If USPS_PackageSizeC="" AND USPS_PackageSize<131 then
							USPS_PackageSizeC="OVERSIZE"
						End if
					End If
					'//Eliminate small box if any side is over 3 inchdes
					dim intshowSmallBox, intShowMedBox
					intshowSmallBox = 1
					intShowMedBox = 1

					LengthIsCal = 0
					A = Cint(session("USPSPackLength"&q))
					B = Cint(session("USPSPackWidth"&q))
					C = Cint(session("USPSPackHeight"&q))
					IntLongestLength = Cint(0)
					IntMidLength = Cint(0)
					IntShortestLength = Cint(0)

					If A=>B AND A=>C Then
						'A is the longest
						IntLongestLength = A
						LengthIsCal = 1
						If B=>C Then
							'B is the mid
							IntMidLength = B
							'C is the shortest
							IntShortestLength = C
						Else
							'C is the mid
							IntMidLength = C
							'B is the shortest
							IntShortestLength = B
						End If
					End If

					If (B=>A AND B=>C) AND (LengthIsCal = 0) Then
						'B is the longest
						IntLongestLength = B
						LengthIsCal = 1
						If A=>C Then
							'A is the mid
							IntMidLength = A
							'C is the shortest
							IntShortestLength = C
						Else
							'C is the mid
							IntMidLength = C
							'A is the shortest
							IntShortestLength = A
						End If
					End If

					If (C=>A AND C=>B) AND (LengthIsCal = 0) Then
						'C is the longest
						IntLongestLength = C
						LengthIsCal = 1
						If B=>A Then
							'B is the mid
							IntMidLength = B
							'A is the shortest
							IntShortestLength = A
						Else
							'A is the mid
							IntMidLength = A
							'B is the shortest
							IntShortestLength = B
						End If
					End If

					If IntShortestLength=>5.50 Then
						tUSPS_PM_PACKAGE = "NONE"
						pcv_PM_Null=0
					Else
						tUSPS_PM_PACKAGE=USPS_PM_PACKAGE
					End If

					If tUSPS_PM_PACKAGE <> "NONE" AND IntShortestLength=>1.5 Then
						intShowSmallBox = 0
						pcv_PMOption="Flat Rate Box1"
					End If
					If pcv_PMOption="Flat Rate Box1" AND IntShortestLength=>3.5 Then
						intShowMedBox = 0
						pcv_PMOption="Flat Rate Box2"
					End If

					usps_postdata=usps_postdata&"<Package%20ID="&chr(34)&iNum&chr(34)&">"
					usps_postdata=usps_postdata&"<Service>PRIORITY</Service>"
					usps_postdata=usps_postdata&"<ZipOrigination>"&scShipFromPostalCode&"</ZipOrigination>"
					usps_postdata=usps_postdata&"<ZipDestination>"&Universal_destination_postal&"</ZipDestination>"
					usps_postdata=usps_postdata&"<Pounds>"&session("USPSPackPounds"&q)&"</Pounds>"
					usps_postdata=usps_postdata&"<Ounces>"&round(session("USPSPackOunces"&q))&"</Ounces>"

					'// If FRE is the default, check for weight limit and alternate container
					if pcv_PM_Null=1 AND pcv_PMOption<>"0" then
						if pcv_PMOption="Flat Rate Box" AND intshowSmallBox = 1 then
							pcv_PMOption = "Sm Flat Rate Box"
						end if
						if pcv_PMOption="Flat Rate Box1" AND intShowMedBox = 1 then
							pcv_PMOption = "Md Flat Rate Box"
						end if
						if pcv_PMOption="Flat Rate Box2" then
							pcv_PMOption = "Lg Flat Rate Box"
						end if
						tUSPS_PM_PACKAGE=pcv_PMOption
					end if

					'private const USPS_PM_FREOption="0"
					if USPS_PM_PACKAGE<>"Flat Rate Envelope" then
						if tUSPS_PM_PACKAGE <> "NONE" then
							'pcv_PMOption=USPS_PM_PACKAGE
							if pcv_PMOption="Flat Rate Box" AND intshowSmallBox = 1 then
								pcv_PMOption = "Sm Flat Rate Box"
							end if
							if pcv_PMOption="Flat Rate Box1" AND intShowMedBox = 1 then
								pcv_PMOption = "Md Flat Rate Box"
							end if
							if pcv_PMOption="Flat Rate Box2" then
								pcv_PMOption = "Lg Flat Rate Box"
							end if
							tUSPS_PM_PACKAGE=pcv_PMOption
						else
							if ucase(tUSPS_PM_PACKAGE)="NONE" then
								'check for Priority Mail totals from previous
								if iUSPSPMRate=0 then
									tUSPS_PM_PACKAGE="VARIABLE"
								end if
							end if
						end if
					end if

					IF USPS_PackageSizeC="LARGE" then
						tUSPS_PM_PACKAGE="RECTANGULAR"
					end if
					usps_postdata=usps_postdata&"<Container>"&tUSPS_PM_PACKAGE&"</Container>"
					usps_postdata=usps_postdata&"<Size>"&USPS_PackageSizeC&"</Size>"
					IF USPS_PackageSizeC="LARGE" THEN
						usps_postdata=usps_postdata&"<Width>"&session("USPSPackWidth"&q)&"</Width>"
						usps_postdata=usps_postdata&"<Length>"&session("USPSPackLength"&q)&"</Length>"
						usps_postdata=usps_postdata&"<Height>"&session("USPSPackHeight"&q)&"</Height>"
						usps_postdata=usps_postdata&"<Girth>"&USPS_PackageSize&"</Girth>"
					END IF
					usps_postdata=usps_postdata&"</Package>"
				next

				usps_postdata=usps_postdata&"</RateV4Request>"

				intUSPSPostage=0
				session("PMService")=""
				strMailService=""
				Set srvUSPS2XmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
				srvUSPS2XmlHttp.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
				srvUSPS2XmlHttp.open "GET", usps_postdata, false
				srvUSPS2XmlHttp.send
				USPS2_result = srvUSPS2XmlHttp.responseText

				' Parse the XML document.
				Set objOutputXMLDoc = Server.CreateObject("Microsoft.XMLDOM")
				objOutputXMLDoc.loadXML srvUSPS2XmlHttp.responseText

				set objLst=objOutputXMLDoc.getElementsByTagName("Package")
				for i = 0 to (objLst.length - 1)
					USPS_TempSize=""
					for j=0 to ((objLst.item(i).childNodes.length)-1)
						If objLst.item(i).childNodes(j).nodeName="Size" then
							USPS_TempSize=objLst.item(i).childNodes(j).Text
						End if
						If objLst.item(i).childNodes(j).nodeName="Postage" then
							for m=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
								If objLst.item(i).childNodes(j).childNodes(m).nodeName="MailService" then
									strMailService = objLst.item(i).childNodes(j).childNodes(m).text
								end if
								If objLst.item(i).childNodes(j).childNodes(m).nodeName="Rate" then
									strRate = objLst.item(i).childNodes(j).childNodes(m).text
								end if
							next
						End if
					next

					pcv_PM_MailService=""
					if instr(strMailService, "(") then
						arrMailService=split(strMailService,"(")
						tstrMailService=arrMailService(0)
					end if
					intHasPriority = Cint(0)
					if instr(strMailService, "Priority Mail Flat-Rate Envelope") OR strMailService="Priority Mail Flat Rate Envelope" OR instr(strMailService, "Flat Rate Envelope") then
						strMailService = replace(strMailService,"&amp;lt;","<")
						strMailService = replace(strMailService,"&amp;gt;",">")
						strMailService = replace(strMailService,"&amp;","&")
						strMailService = replace(strMailService,"&lt;","<")
						strMailService = replace(strMailService,"&gt;",">")
						pcv_PM_MailService="USPS "&strMailService
						session("USPSshipStr")="|?|USPS|9901|"&"X|X|X|"
						if isNumeric(strRate) then
							strRate=cdbl(strRate)
						end if
						intUSPSPostage=intUSPSPostage+strRate
						iUSPSFlag=1
						intHasPriority = 1
					end if

					if intHasPriority = 0 AND (instr(strMailService, "Priority Mail Flat-Rate Box") OR instr(strMailService, "Priority Mail Regular Flat-Rate Boxes") OR instr(strMailService, "Priority Mail Regular/Medium Flat-Rate Boxes") OR strMailService="Priority Mail Small Flat Rate Box"  OR strMailService="Priority Mail Medium Flat Rate Box"  OR strMailService="Priority Mail Large Flat Rate Box" OR instr(strMailService,"Flat Rate Box")) then
						strMailService = replace(strMailService,"&amp;lt;","<")
						strMailService = replace(strMailService,"&amp;gt;",">")
						strMailService = replace(strMailService,"&amp;","&")
						strMailService = replace(strMailService,"&lt;","<")
						strMailService = replace(strMailService,"&gt;",">")
						pcv_PM_MailService="USPS "&strMailService
						session("USPSshipStr")="|?|USPS|9901|"&"X|X|X|"
						if isNumeric(strRate) then
							strRate=cdbl(strRate)
						end if
						intUSPSPostage=intUSPSPostage+strRate
						iUSPSFlag=1
						intHasPriority = 1
					end if

					if intHasPriority = 0 AND (USPS_TempSize="LARGE") then
						session("PMService")="LARGE"
						session("USPSshipStr")="|?|USPS|9901|"&"X|X|X|"
						if isNumeric(strRate) then
							strRate=cdbl(strRate)
						end if
						intUSPSPostage=intUSPSPostage+strRate
						iUSPSFlag=1
						intHasPriority = 1
					end if

					if intHasPriority = 0 AND instr(strMailService, "Priority Mail") Then
						'Priority Mail&amp;lt;sup&amp;gt;&amp;amp;reg;&amp;lt;/sup&amp;gt;
						strMailService = replace(strMailService,"&amp;lt;","<")
						strMailService = replace(strMailService,"&amp;gt;",">")
						strMailService = replace(strMailService,"&amp;","&")
						strMailService = replace(strMailService,"&lt;","<")
						strMailService = replace(strMailService,"&gt;",">")
						pcv_PM_MailService="USPS "&strMailService
						session("USPSshipStr")="|?|USPS|9901|"&"X|X|X|"
						if isNumeric(strRate) then
							strRate=cdbl(strRate)
						end if
						intUSPSPostage=intUSPSPostage+strRate
						iUSPSFlag=1
						intHasPriority = 1
					end if
					intHasPriority = 0
				next
				if iUSPSPMRate<>0 then
					intUSPSPostage=intUSPSPostage+iUSPSPMRate
				end if
				if session("PMService")="LARGE" then
					pcv_PM_MailService="USPS Priority Mail <sup>&reg;</sup>"
				end if
				availableShipStr=availableShipStr&replace(session("USPSshipStr"),"X|X|X|", pcv_PM_MailService)&"|"&intUSPSPostage&"|NA|"
				session("USPSshipStr")=""
			end if
		end if
	end if 'size and weight are ok


	err.number=0

	'//USPS RATES - International
	USPS_destination_country=USPSCountry(Universal_destination_country)
	'// Gather post
	session("USPS_ShowGlobalRates")=""
	session("USPS_ShowExpressRates")=""
	session("USPS_ShowPriorityRates")=""
	session("USPS_ShowFirstClassRates")=""

	usps_postdata=""
	usps_postdata=usps_postdata&usps_server&"?API=IntlRate&XML="

	usps_postdata=usps_postdata&"<IntlRateRequest%20USERID="&chr(34)&usps_userid&chr(34)&">"
	for q=1 to pcv_intPackageNum
		'////////////////////////////////////
		'// Check Package Sizes for services
		'////////////////////////////////////
		'/ Get Dimensional Weight for Global Express
		pcv_USPS_Length=Cint(session("USPSPackLength"&q))
		pcv_USPS_Width=Cint(session("USPSPackWidth"&q))
		pcv_USPS_Height=Cint(session("USPSPackHeight"&q))
		pcv_USPS_DimWeight=((pcv_USPS_Length+pcv_USPS_Width+pcv_USPS_Height)/166)

		pcv_Decval = Mid(pcv_USPS_DimWeight, InStr(1, pcv_USPS_DimWeight, ".") + 1)
		pcv_DimWeightRound = CDbl(pcv_USPS_DimWeight)
		If pcv_Decval >= 0 Then
			 pcv_DimWeightRound = CInt(pcv_USPS_DimWeight)
			 pcv_DimWeightRound = pcv_DimWeightRound + 1
		End If

		if pcv_dimWeightRound>session("USPSPackPounds"&q) then
			'// Uncomment the following two line to use Dimensional Weight for USPS Global Express International Packages
			'session("USPSPackPounds"&q)=pcv_DimWeightRound
			'session("USPSPackOunces"&q)=0
		end if

		iNum=q-1
		usps_postdata=usps_postdata&"<Package%20ID="&chr(34)&iNum&chr(34)&">"
		usps_postdata=usps_postdata&"<Pounds>"&session("USPSPackPounds"&q)&"</Pounds>"
		usps_postdata=usps_postdata&"<Ounces>"&round(session("USPSPackOunces"&q))&"</Ounces>"
		usps_postdata=usps_postdata&"<MailType>Package</MailType>"
		'usps_postdata=usps_postdata&"<MailType>envelope</MailType>"
		if pcv_UseValueOfContents=1 then
			usps_postdata=usps_postdata&"<ValueOfContents>"&pcv_ValueOfContents&"</ValueOfContents>"
		end if
		usps_postdata=usps_postdata&"<Country>"&USPS_destination_country&"</Country>"
		usps_postdata=usps_postdata&"</Package>"

		'if weight is over 70 lbs for any package, we do not show rates for Global Express
		if session("USPSPackPounds"&q)>70 then
			session("USPS_ShowGlobalRates")="NO"
		end if
		if pcv_USPS_Length>46 OR pcv_USPS_Width>46 OR pcv_USPS_Height>46 then
			session("USPS_ShowGlobalRates")="NO"
		end if
		'If Express demension of one side exceeds 36, don't show rates
		if pcv_USPS_Length>36 OR pcv_USPS_Width>36 OR pcv_USPS_Height>36 then
			session("USPS_ShowExpressRates")="NO"
		end if
		'If Priority demension of one side exceeds 36, don't show rates
		if pcv_USPS_Length>60 OR pcv_USPS_Width>60 OR pcv_USPS_Height>60 then
			session("USPS_ShowPriorityRates")="NO"
		end if
		'if First Class is over 4 pounds
		if session("USPSPackPounds"&q)>4 then
			session("USPS_ShowFirstClassRates")="NO"
		end if
		if pcv_USPS_Length>24 OR pcv_USPS_Width>24 OR pcv_USPS_Height>24 then
			session("USPS_ShowFirstClassRates")="NO"
		end if
	next

	usps_postdata=usps_postdata&"</IntlRateRequest>"
	Set srvUSPSINTXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
	srvUSPSINTXmlHttp.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
	srvUSPSINTXmlHttp.open "GET", usps_postdata, false
	srvUSPSINTXmlHttp.send

	USPSINT_result = srvUSPSINTXmlHttp.responseText
	Set USPSINTXMLdoc = server.CreateObject("Msxml2.DOMDocument"&scXML)
	USPSINTXMLDoc.async = false

	if USPSINTXMLDoc.loadXML(USPSINT_result) then ' if loading from a string
		USPSA=0
		USPSA_1=Ccur(0)
		USPSA_2=""
		USPSB=0
		USPSB_1=Ccur(0)
		USPSB_2=""
		USPSC=0
		USPSC_1=Ccur(0)
		USPSC_2=""
		USPSD=0
		USPSD_1=Ccur(0)
		USPSD_2=""
		USPSE=0
		USPSE_1=Ccur(0)
		USPSE_2=""
		USPSF=0
		USPSF_1=Ccur(0)
		USPSF_2=""
		USPSG=0
		USPSG_1=Ccur(0)
		USPSG_2=""
		USPSH=0
		USPSH_1=Ccur(0)
		USPSH_2=""
		USPSI=0
		USPSI_1=Ccur(0)
		USPSI_2=""
		USPSJ=0
		USPSJ_1=Ccur(0)
		USPSJ_2=""


		set objLst=USPSINTXMLDoc.getElementsByTagName("Package")

		for i = 0 to (objLst.length - 1)
			for j=0 to ((objLst.item(i).childNodes.length)-1)
				If objLst.item(i).childNodes(j).nodeName="Service" then
						intCLASSID=objLst.item(i).childNodes(j).getAttribute("ID")
						usps_int_1="0"
						usps_int_2="0"
						usps_int_3="0"
						for m=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="SvcCommitments" then
								usps_int_2 = objLst.item(i).childNodes(j).childNodes(m).text
							end if
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="Postage" then
								usps_int_1 = objLst.item(i).childNodes(j).childNodes(m).text
							end if
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="Insurance" then
								usps_int_3 = objLst.item(i).childNodes(j).childNodes(m).text
							end if
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="SvcDescription" then
								serviceVar = objLst.item(i).childNodes(j).childNodes(m).text
							end if
						Next

						select case intCLASSID

						case "4", "Global Express Guaranteed", "Global Express Guaranteed (GXG)"
							if session("USPS_ShowGlobalRates")="" then
								USPSA=1
								USPSA_1=USPSA_1+ccur(usps_int_1)
								USPSA_2=Trim(replace(usps_int_2,CHR(10),""))
								if isNumeric(usps_int_3) then
									usps_int_3=cdbl(usps_int_3)
									USPSA_1=USPSA_1+usps_int_3
								end if
								iUSPSFlag=1
							end if
						case "6", "Global Express Guaranteed Non-Document Rectangular"
							if session("USPS_ShowGlobalRates")="" then
								USPSB=1
								USPSB_1=USPSB_1+ccur(usps_int_1)
								USPSB_2=Trim(replace(usps_int_2,CHR(10),""))
								if isNumeric(usps_int_3) then
									usps_int_3=cdbl(usps_int_3)
									USPSB_1=USPSB_1+usps_int_3
								end if
								iUSPSFlag=1
							end if
						case "7", "Global Express Guaranteed Non-Document Non-Rectangular"
							if session("USPS_ShowGlobalRates")="" then
								USPSC=1
								USPSC_1=USPSC_1+ccur(usps_int_1)
								USPSC_2=Trim(replace(usps_int_2,CHR(10),""))
								if isNumeric(usps_int_3) then
									usps_int_3=cdbl(usps_int_3)
									USPSC_1=USPSC_1+usps_int_3
								end if
								iUSPSFlag=1
							end if
						case "1", "Express Mail International (EMS)", "Express Mail International"
							if session("USPS_ShowExpressRates")="" then
								USPSD=1
								USPSD_1=USPSD_1+ccur(usps_int_1)
								USPSD_2=Trim(replace(usps_int_2,CHR(10),""))
								if isNumeric(usps_int_3) then
									usps_int_3=cdbl(usps_int_3)
									USPSD_1=USPSD_1+usps_int_3
								end if
								iUSPSFlag=1
							end if
						case "10", "Express Mail International (EMS) Flat Rate Envelope", "Express Mail International Flat Rate Envelope"
							if session("USPS_ShowExpressRates")="" then
								USPSE=1
								USPSE_1=USPSE_1+ccur(usps_int_1)
								USPSE_2=Trim(replace(usps_int_2,CHR(10),""))
								if isNumeric(usps_int_3) then
									usps_int_3=cdbl(usps_int_3)
									USPSE_1=USPSE_1+usps_int_3
								end if
								iUSPSFlag=1
							end if
						case "2", "Priority Mail International"
							if session("USPS_ShowPriorityRates")="" then
								USPSF=1
								USPSF_1=USPSF_1+ccur(usps_int_1)
								USPSF_2=Trim(replace(usps_int_2,CHR(10),""))
								if isNumeric(usps_int_3) then
									usps_int_3=cdbl(usps_int_3)
									USPSF_1=USPSF_1+usps_int_3
								end if
								iUSPSFlag=1
							end if
						case "8", "Priority Mail International Flat Rate Envelope"
							if session("USPS_ShowPriorityRates")="" then
								USPSG=1
								USPSG_1=USPSG_1+ccur(usps_int_1)
								USPSG_2=Trim(replace(usps_int_2,CHR(10),""))
								if isNumeric(usps_int_3) then
									usps_int_3=cdbl(usps_int_3)
									USPSG_1=USPSG_1+usps_int_3
								end if
								iUSPSFlag=1
							end if
						case "9", "Priority Mail International Medium Flat Rate Box"
							if session("USPS_ShowPriorityRates")="" then
								USPSH=1
								USPSH_1=USPSH_1+ccur(usps_int_1)
								USPSH_2=Trim(replace(usps_int_2,CHR(10),""))
								if isNumeric(usps_int_3) then
									usps_int_3=cdbl(usps_int_3)
									USPSH_1=USPSH_1+usps_int_3
								end if
								iUSPSFlag=1
							end if
						case "15", "First-Class Mail International", "First Class Mail International Package", "First-Class Mail International Package"
							if session("USPS_ShowFirstClassRates")="" then
								USPSI=1
								USPSI_1=USPSI_1+ccur(usps_int_1)
								USPSI_2=Trim(replace(usps_int_2,CHR(10),""))
								if isNumeric(usps_int_3) then
									usps_int_3=cdbl(usps_int_3)
									USPSI_1=USPSI_1+usps_int_3
								end if
								iUSPSFlag=1
							end if
						end select

				End If
			Next
		Next

		'//COMPILE availableShipStr
		if USPSA=1 then
			availableShipStr=availableShipStr&"|?|USPS|9914|"&"Global Express Guaranteed<sup>&reg;</sup>|"&USPSA_1&"|"&Trim(replace(USPSA_2,CHR(10),""))
		end if
		if USPSB=1 then
			availableShipStr=availableShipStr&"|?|USPS|9905|"&"Global Express Guaranteed<sup>&reg;</sup> Non-Document Rectangular|"&USPSB_1&"|"&Trim(replace(USPSB_2,CHR(10),""))
		end if
		if USPSC=1 then
			availableShipStr=availableShipStr&"|?|USPS|9910|"&"Global Express Guaranteed<sup>&reg;</sup> Non-Document Non-Rectangular|"&USPSC_1&"|"&Trim(replace(USPSC_2,CHR(10),""))
		end if
		if USPSD=1 then
			availableShipStr=availableShipStr&"|?|USPS|9906|"&"Express Mail<sup>&reg;</sup> International (EMS)|"&USPSD_1&"|"&Trim(replace(USPSD_2,CHR(10),""))
		end if
		if USPSE=1 then
			availableShipStr=availableShipStr&"|?|USPS|9911|"&"Express Mail<sup>&reg;</sup> International (EMS) Flat Rate Envelope|"&USPSE_1&"|"&Trim(replace(USPSE_2,CHR(10),""))
		end if
		if USPSF=1 then
			availableShipStr=availableShipStr&"|?|USPS|9907|"&"Priority Mail<sup>&reg;</sup> International|"&USPSF_1&"|"&Trim(replace(USPSF_2,CHR(10),""))
		end if
		if USPSG=1 then
			availableShipStr=availableShipStr&"|?|USPS|9908|"&"Priority Mail<sup>&reg;</sup> International Flat Rate Envelope|"&USPSG_1&"|"&Trim(replace(USPSG_2,CHR(10),""))
		end if
		if USPSH=1 then
			availableShipStr=availableShipStr&"|?|USPS|9909|"&"Priority Mail<sup>&reg;</sup> International Flat Rate Box|"&USPSH_1&"|"&Trim(replace(USPSH_2,CHR(10),""))
		end if
		if USPSI=1 then
			availableShipStr=availableShipStr&"|?|USPS|9912|"&"First-Class Mail<sup>&reg;</sup> International|"&USPSI_1&"|"&Trim(replace(USPSI_2,CHR(10),""))
		end if

	end if
end if 'if usps is active

if iUSPSFlag=1 then
	strDefaultProvider="USPS"
	iShipmentTypeCnt=iShipmentTypeCnt+1
	strOptionShipmentType=strOptionShipmentType&"<option value=USPS>"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_l")&"</option>"
end if

If CP_active=true or CP_active="-1" then
	iCPActive=1
	'// Canada Post
	'// Compile xml
	CP_postdata=""
	CP_postdata=CP_postdata&"<?xml version=""1.0"" ?>"
	CP_postdata=CP_postdata&"<eparcel>"
	CP_postdata=CP_postdata&"<language>en</language>" '// Prefered language for the
	CP_postdata=CP_postdata&"<ratesAndServicesRequest>" '// Merchant Identification assigned by Canada Post
	CP_postdata=CP_postdata&"<merchantCPCID>"&CP_userid&"</merchantCPCID>"
	CP_postdata=CP_postdata&"<lineItems>"
	for q=1 to pcv_intPackageNum
		CP_postdata=CP_postdata&"<item>"
		CP_postdata=CP_postdata&"<quantity>1</quantity>"
		CP_postdata=CP_postdata&"<weight>"&session("CPPackWeight"&q)&"</weight>"
		CP_postdata=CP_postdata&"<length>"&session("CPPackLength"&q)&"</length>"
		CP_postdata=CP_postdata&"<width>"&session("CPPackWidth"&q)&"</width>"
		CP_postdata=CP_postdata&"<height>"&session("CPPackHeight"&q)&"</height>"
		CP_postdata=CP_postdata&"<description>My Item #"&q&"</description>"
		CP_postdata=CP_postdata&"</item>"
	next
	CP_postdata=CP_postdata&"</lineItems>"
	CP_postdata=CP_postdata&"<city>"&Universal_destination_city&"</city>"
	CP_postdata=CP_postdata&"<provOrState>"&Universal_destination_provOrState&"</provOrState>"
	CP_postdata=CP_postdata&"<country>"&Universal_destination_country&"</country>"
	CP_postdata=CP_postdata&"<postalCode>"&Universal_destination_postal&"</postalCode>"
	CP_postdata=CP_postdata&"</ratesAndServicesRequest>"
	CP_postdata=CP_postdata&"</eparcel>"

	Set srvCPXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
	srvCPXmlHttp.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
	srvCPXmlHttp.open "POST", CP_server, false
	srvCPXmlHttp.send(CP_postdata)

	CP_result = srvCPXmlHttp.responseText
	Set CPXMLdoc = server.CreateObject("Msxml2.DOMDocument"&scXML)
	CPXMLDoc.async = false
	if CPXMLDoc.loadXML(CP_result) then '//  If loading from a string
		set objLst = CPXMLDoc.getElementsByTagName("product")
		for i = 0 to (objLst.length - 1)
			varFlag=0
			CP_ID=objLst.item(i).getAttribute("id")
			for j=0 to ((objLst.item(i).childNodes.length)-1)
				If objLst.item(i).childNodes(j).nodeName="name" then
					serviceVar=objLst.item(i).childNodes(j).text
					select case CP_ID

					case "1010"
						availableShipStr=availableShipStr&"|?|CP|1010|"&"Canada Post - REGULAR"
						varFlag=1
						iCPFlag=1
					case "1020"
						availableShipStr=availableShipStr&"|?|CP|1020|"&"Canada Post - EXPEDITED"
						varFlag=1
						iCPFlag=1
					case "1030"
						availableShipStr=availableShipStr&"|?|CP|1030|"&"Canada Post - XPRESSPOST"
						varFlag=1
						iCPFlag=1
					case "1040"
						availableShipStr=availableShipStr&"|?|CP|1040|"&"Canada Post - PRIORITY COURIER"
						varFlag=1
						iCPFlag=1
					case "1120"
						availableShipStr=availableShipStr&"|?|CP|1120|"&"Canada Post - EXPEDITED EVENING"
						varFlag=1
						iCPFlag=1
					case "1130"
						availableShipStr=availableShipStr&"|?|CP|1130|"&"Canada Post - XPRESSPOST EVENING"
						varFlag=1
						iCPFlag=1
					case "1220"
						availableShipStr=availableShipStr&"|?|CP|1220|"&"Canada Post - EXPEDITED SATURDAY"
						varFlag=1
						iCPFlag=1
					case "1230"
						availableShipStr=availableShipStr&"|?|CP|1230|"&"Canada Post - XPRESSPOST SATURDAY"
						varFlag=1
						iCPFlag=1
					case "2010"
						availableShipStr=availableShipStr&"|?|CP|2010|"&"Canada Post - SURFACE US"
						varFlag=1
						iCPFlag=1
					case "2020"
						availableShipStr=availableShipStr&"|?|CP|2020|"&"Canada Post - AIR US"
						varFlag=1
						iCPFlag=1
					case "2030"
						availableShipStr=availableShipStr&"|?|CP|2030|"&"Canada Post - XPRESSPOST US"
						varFlag=1
						iCPFlag=1
					case "2040"
						availableShipStr=availableShipStr&"|?|CP|2040|"&"Canada Post - PUROLATOR US"
						varFlag=1
						iCPFlag=1
					case "2050"
						availableShipStr=availableShipStr&"|?|CP|2050|"&"Canada Post - PUROPAK US"
						varFlag=1
						iCPFlag=1
					case "3010"
						availableShipStr=availableShipStr&"|?|CP|3010|"&"Canada Post - SURFACE INTERNATIONAL"
						varFlag=1
						iCPFlag=1
					case "3020"
						availableShipStr=availableShipStr&"|?|CP|3020|"&"Canada Post - AIR INTERNATIONAL"
						varFlag=1
						iCPFlag=1
					case "3040"
						availableShipStr=availableShipStr&"|?|CP|3040|"&"Canada Post - PUROLATOR INTERNATIONAL"
						varFlag=1
						iCPFlag=1
					case "3050"
						availableShipStr=availableShipStr&"|?|CP|3050|"&"Canada Post - PUROPAK INTERNATIONAL"
						varFlag=1
						iCPFlag=1
					case "2005"
						availableShipStr=availableShipStr&"|?|CP|2005|"&"Canada Post - SMALL PACKETS SURFACE US"
						varFlag=1
						iCPFlag=1
					case "2015"
						availableShipStr=availableShipStr&"|?|CP|2015|"&"Canada Post - SMALL PACKETS AIR US"
						varFlag=1
						iCPFlag=1
					case "2025"
						availableShipStr=availableShipStr&"|?|CP|2025|"&"Canada Post - EXPEDITED US COMMERCIAL"
						varFlag=1
						iCPFlag=1
					case "3005"
						availableShipStr=availableShipStr&"|?|CP|3005|"&"Canada Post - SMALL PACKETS SURFACE INTERNATIONAL"
						varFlag=1
						iCPFlag=1
					case "3015"
						availableShipStr=availableShipStr&"|?|CP|3015|"&"Canada Post - SMALL PACKETS AIR INTERNATIONAL"
						varFlag=1
						iCPFlag=1
					case "3025"
						availableShipStr=availableShipStr&"|?|CP|3025|"&"Canada Post - XPRESSPOST INTERNATIONAL INTERNATIONAL"
						varFlag=1
						iCPFlag=1

					end select


				End if
				If objLst.item(i).childNodes(j).nodeName="rate" AND varFlag=1 then
					availableShipStr=availableShipStr&"|"&objLst.item(i).childNodes(j).text
				End if
				If objLst.item(i).childNodes(j).nodeName="shippingDate" AND varFlag=1 then
					shippingDate=objLst.item(i).childNodes(j).text
					if shippingDate<>"" then
						shippingDateArry=split(shippingDate,"-")
						shippingDateMonth=shippingDateArry(1)
						shippingDateYear=shippingDateArry(0)
						shippingDateDay=shippingDateArry(2)
						shippingDateFrmt=(shippingDateMonth&"/"&shippingDateDay&"/"&shippingDateYear)
					end if
					'availableShipStr=availableShipStr&"|"&objLst.item(i).childNodes(j).text&" Days"
				End if
				If objLst.item(i).childNodes(j).nodeName="deliveryDate" AND varFlag=1 then
					deliveryDate=objLst.item(i).childNodes(j).text
					if deliveryDate<>"" then
						deliveryDateArry=split(deliveryDate,"-")
						deliveryDateMonth=deliveryDateArry(1)
						deliveryDateYear=deliveryDateArry(0)
						deliveryDateDay=deliveryDateArry(2)
						deliveryDateFrmt=(deliveryDateMonth&"/"&deliveryDateDay&"/"&deliveryDateYear)
					end if
					DeliveryDays=DateDiff("d",shippingDateFrmt,deliveryDateFrmt)
					availableShipStr=availableShipStr&"|"&DeliveryDays&" Days"
				End if
			next
		next
	end if
end if '//  If canada post is active

if iCPFlag=1 then
	strDefaultProvider="CP"
	iShipmentTypeCnt=iShipmentTypeCnt+1
	strOptionShipmentType=strOptionShipmentType&"<option value=CP>"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_n")&"</option>"
end if

'// Custom Rates
ERR.NUMBER=0
mySQL="SELECT idFlatShiptype,WQP,FlatShipTypeDesc,FlatShipTypeDelivery,startIncrement FROM FlatShipTypes"
set rstemp=conntemp.execute(mySQL)

Do until rstemp.eof
	ifound=0
	idTemp=rstemp("idFlatShiptype")
	VarWQP=trim(rstemp("WQP"))

	If len(VarWQP)>0 Then

	  select case VarWQP
	  case "W"
		  mySQL2="SELECT shippingPrice FROM FlatShipTypeRules WHERE idFlatShipType="& idTemp &" AND quantityTo>=" &intCustomShipWeight& " AND quantityFrom<=" &intCustomShipWeight
	  case "Q"
		  mySQL2="SELECT shippingPrice FROM FlatShipTypeRules WHERE idFlatShipType="& idTemp &" AND quantityTo>=" &pCartShipQuantity & " AND quantityFrom<=" & pCartShipQuantity
	  case "P"
		  mySQL2="SELECT shippingPrice FROM FlatShipTypeRules WHERE idFlatShipType="& idTemp &" AND quantityTo>=" & pShipSubTotal & " AND quantityFrom<=" & pShipSubTotal
	  case "O"
		  mySQL2="SELECT shippingPrice FROM FlatShipTypeRules WHERE idFlatShipType="& idTemp &" AND quantityTo>=" & pShipSubTotal & " AND quantityFrom<=" & pShipSubTotal
	  case "I"
		  if pCartShipQuantity=1 then
			  pCartShipQuantity2=2
		  else
			  pCartShipQuantity2=pCartShipQuantity
		  end if
		  mySQL2="SELECT shippingPrice, quantityTo FROM FlatShipTypeRules WHERE idFlatShipType="& idTemp &" AND quantityFrom<=" & pCartShipQuantity2
	  end select

	  set rsShipObj=conntemp.execute(mySQL2)

	if NOT rsShipObj.eof then
		ifound=1
		tempShipPrice=rsShipObj("shippingPrice")
		availableShipStr=availableShipStr&"|?|CUSTOM|C"&idTemp&"|"&rstemp("FlatShipTypeDesc")
		iCustomFlag=1

		'// Calculate shipping price for I and O
		if VarWQP="O" then
			'// Shipping price is the percentage
			tempPercentage=tempShipPrice
			tempShipPrice=((tempPercentage/100)*pShipSubTotal)
		end if

		if VarWQP="I" then
			dim iRegPrice, iAddRegPrice
			mySQL3="SELECT startIncrement FROM FlatShipTypes WHERE startIncrement>0 AND idFlatShipType="&idTemp
			set rsIncretObj=conntemp.execute(mySQL3)
			TempShipPrice=rsIncretObj("startIncrement")
			IShipCnt=pCartShipQuantity
			iTempCompleted=0
			query="SELECT quantityFrom, quantityTo, shippingPrice FROM FlatShipTypeRules WHERE (((FlatShipTypeRules.idFlatshipType)="&idTemp&")) ORDER BY FlatShipTypeRules.num;"
			set rsIncretObj=conntemp.execute(query)
			Do until rsIncretObj.eof or iTempCompleted=1
				iQuantityFrom=rsIncretObj("quantityFrom")
				iQuantityTo=rsIncretObj("quantityTo")
				AddPrice=rsIncretObj("shippingPrice")

				if ccur(IShipCnt) - ccur(iQuantityTo) => 0 then
					TierCnt = (ccur(iQuantityTo) - ccur(iQuantityFrom))+1
					TempShipPrice = TempShipPrice + (ccur(TierCnt) * ccur(AddPrice))
				else
					if IShipCnt=>ccur(iQuantityFrom) then
						TierCnt = (ccur(IShipCnt) - ccur(iQuantityFrom))+1
						TempShipPrice = TempShipPrice + (TierCnt * ccur(AddPrice))
						iTempCompleted=1
					else
						iTempCompleted=1
					end if
				end if
				rsIncretObj.moveNext
			loop
		  end if

		  availableShipStr=availableShipStr&"|"&tempShipPrice

		  CustomShipDelivery=rstemp("FlatShipTypeDelivery")
		  if CustomShipDelivery="" then
			  availableShipStr=availableShipStr&"|NA"
		  else
			  availableShipStr=availableShipStr&"|"&rstemp("FlatShipTypeDelivery")
		  end if
		else
			query1="SELECT idshipservice FROM shipService WHERE serviceCode like 'C" & idTemp & "' AND serviceFree<>0 AND serviceFreeOverAmt<" & pShipSubTotal & ";"
			set rsShipObj=conntemp.execute(query1)
			if NOT rsShipObj.eof then
				ifound=1
				tempShipPrice=0
				availableShipStr=availableShipStr&"|?|CUSTOM|C"&idTemp&"|"&rstemp("FlatShipTypeDesc")
				iCustomFlag=1
			end if
		end if

	End If '// If len(VarWQP)>0 Then
	rstemp.moveNext
loop

if iCustomFlag=1 then
	strDefaultProvider="CUSTOM"
	iShipmentTypeCnt=iShipmentTypeCnt+1
	strOptionShipmentType=strOptionShipmentType&"<option value=CUSTOM>"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_j")&"</option>"
end if


if instr(strOptionShipmentType, scDefaultProvider) AND scDefaultProvider<>"" AND iShipmentTypeCnt>1 then
	strDefaultProvider=scDefaultProvider
	TempDefaultProvider = scDefaultProvider
else
	If instr(strOptionShipmentType, "UPS") Then
		strDefaultProvider="UPS"
		TempDefaultProvider = strDefaultProvider
	End If
end if

if pcv_intTotPackageNum="1" then
	Dim tmpList
	tmpList="*****"
	Dim tmpCount,tmpCount1
	tmpCount1=0
	tmpCount=0
	pcCartArray=Session("pcCartSession")
	pcCartIndex=Session("pcCartIndex")

	for f=1 to pcCartIndex
		tmp_idproduct=pcCartArray(f,0)
		query="SELECT products.pcDropShipper_ID,pcDropShippersSuppliers.pcDS_IsDropShipper FROM products,pcDropShippersSuppliers WHERE products.idproduct=" & tmp_idproduct & " AND products.pcProd_IsDropShipped=1 AND pcDropShippersSuppliers.idproduct=products.idproduct;"
		set rs=connTemp.execute(query)
		if not rs.eof then
			if Instr(tmpList,"*****" & rs("pcDropShipper_ID") & "**" & rs("pcDS_IsDropShipper") & "*****")=0 then
				tmpList=tmpList & rs("pcDropShipper_ID") & "**" & rs("pcDS_IsDropShipper") & "*****"
				tmpCount=tmpCount+1
			end if
		else
			tmpCount1=tmpCount1+1
		end if
		set rs=nothing
	next
	if tmpCount1>0 then
		pcv_intTotPackageNum=pcv_intTotPackageNum+tmpCount
	else
		pcv_intTotPackageNum=tmpCount
	end if
end if

'// Kill Sessions
for q=1 to pcv_intPackageNum
	'session("UPSPackWidth"&q)=""
	session("FEDEXPackWidth"&q)="" '// SD
	session("FEDEXWSPackWidth"&q)="" '// WS
	session("CPPackWidth"&q)=""
	session("UPSPackHeight"&q)=""
	session("FEDEXPackHeight"&q)="" '// SD
	session("FEDEXWSPackHeight"&q)="" '// WS
	session("CPPackHeight"&q)=""
	session("UPSPackLength"&q)=""
	session("FEDEXPackLength"&q)="" '// SD
	session("FEDEXWSPackLength"&q)="" '// WS
	session("CPPackLength"&q)=""
	'session("UPSPackWeight"&q)=""
	session("UPSPackPrice"&q)=""
	session("FEDEXPackPrice"&q)="" '// SD
	session("FEDEXWSPackWeight"&q)="" '// WS
	session("FEDEXPackPrice"&q)="" '// SD
	session("FEDEXWSPackWeight"&q)="" '// WS
	session("CPPackWeight"&q)=""
	session("USPSPackWidth"&q)=""
	session("USPSPackHeight"&q)=""
	session("USPSPackLength"&q)=""
	session("USPSPackPounds"&q)=""
	session("USPSPackOunces"&q)=""
next
%>
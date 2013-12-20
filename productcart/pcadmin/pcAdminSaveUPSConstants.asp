<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/stringfunctions.asp"-->
<% 
'/////////////////////////////////////////////////////
'// Write all changes to Settings.asp file
'/////////////////////////////////////////////////////
Dim objFS
Dim objFile

Set objFS = Server.CreateObject ("Scripting.FileSystemObject")
if PPD="1" then
	pcStrFileName=Server.Mappath ("/"&scPcFolder&"/includes/UPSconstants.asp")
else
	pcStrFileName=Server.Mappath ("../includes/UPSconstants.asp")
end if
select case pcStrPickupType
	case "01"
		pcStrClassificationType="01"
	case "03"
		pcStrClassificationType="03"
	case "11"
		pcStrClassificationType="04"
end select

If pcStrDynamicInsuredValue="" then
	pcStrDynamicInsuredValue = "0"
End If
If pcStrDynamicInsuredValue="0" AND pcCurInsuredValue="" then
	pcCurInsuredValue="100.00"
End If
if pcStrUseNegotiatedRates="1" AND pcStrShipperNumber="" then
	pcStrUseNegotiatedRates="0"
end if
if pcStrUseNegotiatedRates<>"1" then
	pcStrUseNegotiatedRates="0"
end if

Set objFile = objFS.OpenTextFile (pcStrFileName, 2, True, 0)
objFile.WriteLine CHR(60)&CHR(37)&"'// UPS Constants //" & vbCrLf
objFile.WriteLine "private const UPS_PICKUP_TYPE = """&pcStrPickupType&"""" & vbCrLf
objFile.WriteLine "private const UPS_PACKAGE_TYPE = """&pcStrPackageType&"""" & vbCrLf
objFile.WriteLine "private const UPS_CLASSIFICATION_TYPE = """&pcStrClassificationType&"""" & vbCrLf
objFile.WriteLine "private const UPS_HEIGHT = """&pcStrPackageHeight&"""" & vbCrLf
objFile.WriteLine "private const UPS_WIDTH = """&pcStrPackageWidth&"""" & vbCrLf
objFile.WriteLine "private const UPS_LENGTH = """&pcStrPackageLength&"""" & vbCrLf
objFile.WriteLine "private const UPS_DIM_UNIT = """&pcStrPackageDimUnit&"""" & vbCrLf
objFile.WriteLine "private const UPS_COMPANYNAME = """&pcStrShipperCompanyName&"""" & vbCrLf
objFile.WriteLine "private const UPS_ATTENTION = """&pcStrShipperAttentionName&"""" & vbCrLf
objFile.WriteLine "private const UPS_ADDRESS1 = """&pcStrShipperAddress1&"""" & vbCrLf
objFile.WriteLine "private const UPS_ADDRESS2 = """&pcStrShipperAddress2&"""" & vbCrLf
objFile.WriteLine "private const UPS_ADDRESS3 = """&pcStrShipperAddress3&"""" & vbCrLf
objFile.WriteLine "private const UPS_CITY = """&pcStrShipperCity&"""" & vbCrLf
objFile.WriteLine "private const UPS_STATE = """&pcStrShipperState&"""" & vbCrLf
objFile.WriteLine "private const UPS_POSTALCODE = """&pcStrShipperPostalCode&"""" & vbCrLf
objFile.WriteLine "private const UPS_COUNTRY = """&pcStrShipperCountryCode&"""" & vbCrLf
objFile.WriteLine "private const UPS_PHONE = """&pcStrShipperPhone&"""" & vbCrLf
objFile.WriteLine "private const UPS_FAX = """&pcStrShipperFax&"""" & vbCrLf
objFile.WriteLine "private const UPS_INSUREDVALUE = """&pcCurInsuredValue&"""" & vbCrLf
objFile.WriteLine "private const UPS_DYNAMICINSUREDVALUE = """&pcStrDynamicInsuredValue&"""" & vbCrLf
objFile.WriteLine "private const UPS_USENEGOTIATEDRATES = """&pcStrUseNegotiatedRates &"""" & vbCrLf
objFile.WriteLine "private const UPS_SHIPPERNUM = """&pcStrShipperNumber &"""" & vbCrLf
objFile.WriteLine "'// UPS Constants // " &CHR(37)&CHR(62)& vbCrLf
objFile.Close

set objFS=nothing
set objFile=nothing

%>
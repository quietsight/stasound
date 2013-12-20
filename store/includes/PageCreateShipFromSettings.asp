<!--#include file="adminv.asp"-->
<!--#include file="storeconstants.asp"-->
<!--#include file="shipFromSettings.asp"-->
<!--#include file="secureadminfolder.asp"-->
<!--#include file="rc4.asp"--> 

<% 
' form parameters		
pShipFromPersonName=Session("pcAdminpShipFromPersonName")
pShipFromPersonName=replace(pShipFromPersonName,"''","'")
pShipFromName=Session("pcAdminpShipFromName")
pShipFromName=replace(pShipFromName,"''","'")
pShipFromDepartment=Session("pcAdminpShipFromDepartment")
pShipFromDepartment=replace(pShipFromDepartment,"''","'")
pShipFromPhone=Session("pcAdminpShipFromPhone")
pShipFromPage=Session("pcAdminpShipFromPage")
pShipFromFax=Session("pcAdminpShipFromFax")
pShipFromAddress1=Session("pcAdminpShipFromAddress1")
pShipFromAddress2=Session("pcAdminpShipFromAddress2")
pShipFromAddress3=Session("pcAdminpShipFromAddress3")
pShipFromCity=Session("pcAdminpShipFromCity")
pShipFromCity=replace(pShipFromCity,"''","'")
pShipFromPostalCode=Session("pcAdminpShipFromPostalCode")
pShipFromZip4=Session("pcAdminpShipFromZip4")
if Session("pcAdminpShipFromProvince") <> "" then
	pShipFromState=Session("pcAdminpShipFromProvince")
else
	pShipFromState=Session("pcAdminpShipFromState")
end if
pShipFromPostalCountry=Session("pcAdminpShipFromPostalCountry")
pPackageWeightLimit=Session("pcAdminpackageWeightLimit")

if NOT isNumeric(pPackageWeightLimit) then
	pPackageWeightLimit=0
end if

pDefaultProvider=Session("pcAdminDefaultProvider")

pAlwAltShipAddress=Session("pcAdminAlwAltShipAddress")
if pAlwAltShipAddress="" then
	pAlwAltShipAddress="0"
end if

If pAlwAltShipAddress=0 Then
	pHideShipAddress="1"
End If

pComResShipAddress=Session("pcAdminComResShipAddress")
if pComResShipAddress="" then
	pComResShipAddress="0"
end if

pAlwNoShipRates=Session("pcAdminAlwNoShipRates")
if pAlwNoShipRates="" then
	pAlwNoShipRates="0"
end if
pShowProductWeight=Session("pcAdminpShowProductWeight")
if pShowProductWeight="" then
	pShowProductWeight="0"
end if

'Start SDBA
tShipNotifySeparate=Session("pcAdminsds_NotifySeparate")
if tShipNotifySeparate="" then
	tShipNotifySeparate="0"
end if
'End SDBA

pShowCartWeight=Session("pcAdminpShowCartWeight")
if pShowCartWeight="" then
	pShowCartWeight="0"
end if
pShowEstimateLink=Session("pcAdminpShowEstimateLink")
if pShowEstimateLink="" then
	pShowEstimateLink="0"
end if
pHideProductPackage=Session("pcAdminpHideProductPackage")
if pHideProductPackage="" then
	pHideProductPackage="0"
end if

pHideEstimateDeliveryTimes=Session("pcAdminpHideEstimateDeliveryTimes")
if pHideEstimateDeliveryTimes="" then
	pHideEstimateDeliveryTimes="0"
end if

pShipFromWeightUnit=scShipFromWeightUnit
if pShipFromWeightUnit="" then
	pShipFromWeightUnit="LBS"
end if
pSectionShow=Session("pcAdminsectionShow")
select case pSectionShow
	case "NA"
		pRatesOnly="NO"
		pShipDetailTitle=""
		pShipDetails=""
	case "TOP"
		pRatesOnly="NO"
		pShipDetailTitle=Session("pcAdminshipDetailTitle")
		pShipDetailTitle=replace(pShipDetailTitle,"''","'")
		pShipDetailTitle=replace(pShipDetailTitle,"""","&quot;")
		if pShipDetailTitle="" then
			strErr="Ship Details Title is a required field."
		end if
		pShipDetails=Session("pcAdminshipDetails")
		pShipDetails=replace(pShipDetails,vbCrlF,"<BR>")
		pShipDetails=replace(pShipDetails,"""","""""")
		if pShipDetails="" then
			if strErr<>"" then
				strErr=strErr&"<BR>"
			else
				strErr=strErr&"Ship Details is a required field."
			end if
		end if
	case "BTM"
		pRatesOnly=Session("pcAdminratesOnly")
		if pRatesOnly="YES" then
		else
			pRatesOnly="NO"
		end if
		pShipDetailTitle=Session("pcAdminshipDetailTitle")
		pShipDetailTitle=replace(pShipDetailTitle,"''","'")
		pShipDetailTitle=replace(pShipDetailTitle,"""","&quot;")
		if pShipDetailTitle="" then
			strErr="Ship Details Title is a required field."
		end if
		pShipDetails=Session("pcAdminshipDetails")
		pShipDetails=replace(pShipDetails,vbCrlF,"<BR>")
		pShipDetails=replace(pShipDetails,"""","""""")
		if pShipDetails="" then
			if strErr<>"" then
				strErr=strErr&"<BR>"
			else
				strErr=strErr&"Ship Details is a required field."
			end if
		end if
end select

'check permissions on include folder
Dim q, PageName, findit, Body, f, fso
' request values
q=Chr(34)
PageName="shipFromSettings.asp"
findit=Server.MapPath(PageName)

Body=CHR(60)&CHR(37)&"private const scShipFromName="&q&pShipFromName&q&CHR(10)
Body=Body & "private const scOriginPersonName="&q&pShipFromPersonName&q&CHR(10)
Body=Body & "private const scOriginDepartment="&q&pShipFromDepartment&q&CHR(10)
Body=Body & "private const scOriginPhoneNumber="&q&pShipFromPhone&q&CHR(10)
Body=Body & "private const scOriginPagerNumber="&q&pShipFromPage&q&CHR(10)
Body=Body & "private const scOriginFaxNumber="&q&pShipFromFax&q&CHR(10)
Body=Body & "private const scShipFromAddress1="&q&pShipFromAddress1&q&CHR(10)
Body=Body & "private const scShipFromAddress2="&q&pShipFromAddress2&q&CHR(10)
Body=Body & "private const scShipFromAddress3="&q&pShipFromAddress3&q&CHR(10)
Body=Body & "private const scShipFromCity="&q&pShipFromCity&q&CHR(10)
Body=Body & "private const scShipFromState="&q&pShipFromState&q&CHR(10)
Body=Body & "private const scShipFromPostalCode="&q&pShipFromPostalCode&q&CHR(10)
Body=Body & "private const scShipFromZip4="&q&pShipFromZip4&q&CHR(10)
Body=Body & "private const scAlwAltShipAddress="&q&pAlwAltShipAddress&q&CHR(10)
Body=Body & "private const scComResShipAddress="&q&pComResShipAddress&q&CHR(10)
Body=Body & "private const scAlwNoShipRates="&q&pAlwNoShipRates&q&CHR(10)
Body=Body & "private const scShipFromPostalCountry="&q&pShipFromPostalCountry&q&CHR(10)
Body=Body & "private const scShowProductWeight="&q&pShowProductWeight&q&CHR(10)
Body=Body & "private const scPackageWeightLimit="&q&pPackageWeightLimit&q&CHR(10)
Body=Body & "private const scShowCartWeight="&q&pShowCartWeight&q&CHR(10)
Body=Body & "private const scShowEstimateLink="&q&pShowEstimateLink&q&CHR(10)
Body=Body & "private const scHideProductPackage="&q&pHideProductPackage&q&CHR(10)

Body=Body & "private const scHideEstimateDeliveryTimes="&q&pHideEstimateDeliveryTimes&q&CHR(10)

Body=Body & "private const scShipFromWeightUnit="&q&pShipFromWeightUnit&q&CHR(10)
Body=Body & "private const scDefaultProvider="&q&pDefaultProvider&q&CHR(10)
Body=Body & "private const scHideShipAddress="&q&pHideShipAddress&q&CHR(10)

'Start SDBA
Body=Body & "private const scShipNotifySeparate="&q&tShipNotifySeparate&q&CHR(10)
'End SDBA

Body=Body & "private const PC_SECTIONSHOW="&q&pSectionShow&q&CHR(10)
Body=Body & "private const PC_RATESONLY="&q&pRatesOnly&q&CHR(10)
Body=Body & "private const PC_SHIP_DETAIL_TITLE="&q&pShipDetailTitle&q&CHR(10)
Body=Body & "private const PC_SHIP_DETAILS="&q&pShipDetails&q&CHR(10)&CHR(37)&CHR(62)

'on error resume next
Set fso=server.CreateObject("Scripting.FileSystemObject")
Set f=fso.GetFile(findit)
Err.number=0
f.Delete
if Err.number>0 then
	response.redirect "../"&scAdminFolderName&"/techErr.asp?error="&Server.URLEncode("Permissions Not Set to Modify Constants")
end if
Set f=nothing
Set f=fso.OpenTextFile(findit, 2, True)
f.Write Body
f.Close
Set fso=nothing
Set f=nothing

if trim(strErr)<>"" then
	response.redirect "../"&scAdminFolderName&"/modFromShipper.asp?msg="&strErr
	else
	response.redirect "../"&scAdminFolderName&"/modFromShipper.asp?s=1&message="&Server.URLEncode("Shipping settings updated successfully")
end if
%>
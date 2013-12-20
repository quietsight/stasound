<!--#include file="adminv.asp"-->
<!--#include file="settings.asp"-->
<!--#include file="storeconstants.asp"-->
<!--#include file="secureadminfolder.asp"-->
<!--#include file="productcartFolder.asp"-->
<!--#include file="taxsettings.asp"-->
<!--#include file="ppdstatus.inc"-->
<!--#include file="openDb.asp"-->
<%
Dim PageName, Body
Dim FS, f, findit

' request values
q=Chr(34)
if request.queryString("sa")<>"" then
	tTaxonCharges=pTaxonCharges
	if tTaxonCharges="" then
		tTaxonCharges=0
	end if
	tTaxonFees=pTaxonFees
	if tTaxonFees="" then
		tTaxonFees=0
	end if
	ttaxfile=ptaxfile
	if ttaxfile="" then
		ttaxfile=0
	end if
	ttaxshippingaddress=ptaxshippingaddress
	ttaxseparate=ptaxseparate
	ttaxwholesale=ptaxwholesale
	tshowVatID = pshowVatID
	tshowSSN = pshowSSN
	tVatIDReq = pVatIDReq
	tSSNReq = pSSNReq
	ttaxfilename=ptaxfilename
	ttaxCanada=ptaxCanada
	if ttaxfile="1" then
		ttaxCanada="0"
	end if
	PageName="taxsettings.asp"
	refpage="AdminTaxSettings_file.asp"
	strDelete=request.queryString("sa")
	rateArray=split(ptaxRateDefault,", ")
	stateArray=split(ptaxRateState,", ")
	taxSNHArray=split(ptaxSNH,", ")
	ttaxRateDefault=""
	ttaxRateState=""
	ttaxSNH=""
	for i=0 to ubound(stateArray)-1
		if stateArray(i)<>strDelete then
			ttaxRateDefault=ttaxRateDefault&rateArray(i)&", "
			ttaxRateState=ttaxRateState&stateArray(i)&", "
			ttaxSNH=ttaxSNH&taxSNHArray(i)&", "
		end if
	next
else
	if request.Form("RateOnly")="1" OR request("ActivateZone")="1" then
		tTaxonCharges=pTaxonCharges
		if tTaxonCharges="" then
			tTaxonCharges=0
		end if
		tTaxonFees=pTaxonFees
		if tTaxonFees="" then
			tTaxonFees=0
		end if
		ttaxfile=ptaxfile
		if ttaxfile="" then
			ttaxfile=0
		end if
		pcv_DefaultCheck=replace(trim(ptaxRateState),",","")
		if pcv_DefaultCheck="" then
			instAddComma=""
			ttaxRateDefault=""
			ttaxRateState=""
			ttaxSNH=""
		else
			stateArray=split(ptaxRateState,", ")
			if ptaxRateState<>"" and ubound(stateArray)=0 then
				instAddComma=", "
			else
				instAddComma=""
			end if
			ttaxRateDefault=ptaxRateDefault
			ttaxRateState=ptaxRateState
			ttaxSNH=ptaxSNH
		end if
		if request("RateOnly")="1" then
			ttaxRateDefault=ttaxRateDefault&instAddComma&replace(request.form("taxRateDefault"),"%","")&", "
			if request.Form("PopForm")="YES" then
				ttaxSNH=ttaxSNH&instAddComma&request.Form("taxSNH")&", "
			else
				ttaxSNH=ttaxSNH&instAddComma&request.Form("taxSNH"&stateArray(0))&", "
			end if
			ttaxRateState=ttaxRateState&instAddComma&request.form("taxRateState")&", "
			PageName=request.form("page_name")
			refpage=request.form("refpage")
			ttaxCanada=ptaxCanada
			if ttaxCanada="" then
				ttaxCanada="0"
			end if
			if ttaxfile="1" then
				ttaxCanada="0"
			end if
		else
			ttaxRateDefault=ptaxRateDefault
			ttaxSNH=ptaxSNH
			ttaxRateState=ptaxRateState
			ttaxCanada="1"
			PageName="taxsettings.asp"
			refpage="AddTaxPerZone.asp"
		end if
		ttaxshippingaddress=ptaxshippingaddress
		ttaxseparate=ptaxseparate
		ttaxwholesale=ptaxwholesale
		tshowVatID = pshowVatID
		tshowSSN = pshowSSN
		tVatIDReq = pVatIDReq
		tSSNReq = pSSNReq
		ttaxVATrate=ptaxVATrate
		ttaxVATRate_Code=ptaxVATRate_Code
		ttaxVAT=ptaxVAT
		ttaxdisplayVAT=ptaxdisplayVAT
		ttaxfilename=ptaxfilename
	else
		tTaxonCharges=request.form("TaxonCharges")
		if tTaxonCharges="" then
			tTaxonCharges="0"
		end if
		tTaxonFees=request.form("TaxonFees")
		if tTaxonFees="" then
			tTaxonFees="0"
		end if
		ttaxfile=request.form("taxfile")
		if ttaxfile="" then
			ttaxfile="0"
		end if
		ttaxRateState=request.form("taxRateState")&", "
		tempRateStateArray=split(ttaxRateState,", ")
		if request.Form("PopForm")="YES" then
			ttaxSNH=ttaxSNH&request.Form("taxSNH")&", "
		else
			for j=0 to ubound(tempRateStateArray)-1
				ttaxSNH=ttaxSNH&request.Form("taxSNH"&tempRateStateArray(j))&", "
			next
		end if
		ttaxshippingaddress=request.form("taxshippingaddress")
		if ttaxshippingaddress="" then
			ttaxshippingaddress="0"
		end if
		ttaxseparate=request.form("taxseparate")
		if ttaxseparate="" then
			ttaxseparate="0"
		end if
		ttaxwholesale=request.form("taxwholesale")
		if ttaxwholesale="" then
			ttaxwholesale="0"
		end if		
		tshowVatID=request.form("showVatID")
		if tshowVatID="" then
			tshowVatID="0"
		end if		
		tshowSSN=request.form("showSSN")
		if tshowSSN="" then
			tshowSSN="0"
		end if
		tVatIDReq=request.form("VatIDReq")
		if tVatIDReq="" then
			tVatIDReq="0"
		end if
		tSSNReq=request.form("SSNReq")
		if tSSNReq="" then
			tSSNReq="0"
		end if
		ttaxVATrate=replace(request.form("taxVATrate"),"%","")
		if ttaxVATrate="" then
			ttaxVATrate="0"
		end if
		ttaxVATRate_Code=request.form("taxVATRate_Code")
		ttaxVAT=request.form("taxVAT")
		if ttaxVAT="" then
			ttaxVAT="0"
		end if
		ttaxdisplayVAT=request.form("taxdisplayVAT")
		if ttaxdisplayVAT="" then
			ttaxdisplayVAT="0"
		end if
		ttaxRateDefault=replace(request.form("taxRateDefault"),"%","")&", "
		if ttaxRateState="" then
			ttaxRateDefault="0"
		end if
		If NOT isNumeric(ttaxRateDefault) then
			'ttaxRateDefault="0"
		End If
		ttaxfilename=request.form("taxfilename")
		ttaxCanada=ptaxCanada
		if ttaxfile="1" then
			ttaxCanada="0"
		end if
		if ttaxfile="1" then
			dim query, conntemp, rs
			'on error resume next
			call openDb()
			query="DELETE FROM pcTaxZoneRates;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			query="DELETE FROM pcTaxZonesGroups;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			query="DELETE FROM pcTaxGroups;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			query="DELETE FROM pcTaxZoneDescriptions;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			query="DELETE FROM pcTaxZones;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			set rs=nothing
			call closedb()
		end if
		PageName=request.form("page_name")
		refpage=request.form("refpage")
		if ttaxfile="1" and ttaxfilename="" then
			response.redirect "../"&scAdminFolderName&"/AdminTaxSettings_file.asp?nofilename=1"
		end if
	end if
end if

on error resume next
if PPD="1" then
	findit=Server.MapPath("/"&scPcFolder&"/pc/tax/"&ttaxfilename)
else
	findit=Server.MapPath("../pc/tax/"&ttaxfilename)
end if

Set fso=server.CreateObject("Scripting.FileSystemObject")
fileLocate=findit
Set f=fso.GetFile(fileLocate)
if err.number>0 then
	nofile=1
else
	nofile=0
end if
err.number=0
findit=Server.MapPath(PageName)
Body=CHR(60)&CHR(37)&"private const pTaxonCharges="&tTaxonCharges&CHR(10)
Body=Body & "private const pTaxonFees="&tTaxonFees&CHR(10)
Body=Body & "private const ptaxfile="&ttaxfile&CHR(10)
Body=Body & "private const ptaxsetup=1"&CHR(10)
Body=Body & "private const ptaxshippingaddress="&q&ttaxshippingaddress&q&CHR(10)
Body=Body & "private const ptaxseparate="&q&ttaxseparate&q&CHR(10)
Body=Body & "private const ptaxwholesale="&q&ttaxwholesale&q&CHR(10)
Body=Body & "private const pshowVatID="&q&tshowVatID&q&CHR(10)
Body=Body & "private const pVatIdReq="&q&tVatIdReq&q&CHR(10)
Body=Body & "private const pshowSSN="&q&tshowSSN&q&CHR(10)
Body=Body & "private const pSSNReq="&q&tSSNReq&q&CHR(10)
Body=Body & "private const ptaxVATrate="&q&ttaxVATrate&q&CHR(10)
Body=Body & "private const ptaxVATRate_Code="&q&ttaxVATRate_Code&q&CHR(10)
Body=Body & "private const ptaxVAT="&q&ttaxVAT&q&CHR(10)
Body=Body & "private const ptaxdisplayVAT="&q&ttaxdisplayVAT&q&CHR(10)
Body=Body & "private const ptaxRateDefault="&q&ttaxRateDefault&q&CHR(10)
Body=Body & "private const ptaxRateState="&q&ttaxRateState&q&CHR(10)
Body=Body & "private const ptaxSNH="&q&ttaxSNH&q&CHR(10)
Body=Body & "private const ptaxCanada="&q&ttaxCanada&q&CHR(10)
Body=Body & "private const ptaxfilename="&q&ttaxfilename&q&CHR(37)&CHR(62) 

' create the file using the FileSystemObject

	Set fso=server.CreateObject("Scripting.FileSystemObject")
	Set f=fso.GetFile(findit)
	Err.number=0
	f.Delete
	if Err.number>0 then
		response.redirect "../"&scAdminFolderName&"/techErr.asp?error="&Server.URLEncode("Permissions Not Set to Modify Tax")
	end if
	Set f=nothing
	
	Set f=fso.OpenTextFile(findit, 2, True)
	f.Write Body
	f.Close
	
	Set fso=nothing
	Set f=nothing
	
	if request.Form("RateOnly")="1" then
		response.redirect "../"&scAdminFolderName&"/"&refpage&"?ro=1&rstate="&request.form("taxRateState")&"&nofile="&nofile
	else
		response.redirect "../"&scAdminFolderName&"/"&refpage&"?nofile="&nofile
	end if %>
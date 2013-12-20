<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<% response.Buffer=true %>
<% PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include FILE="../includes/pcProductOptionsCode.asp"--> 
<%
Dim f, query, conntemp, rstemp
Dim pcv_strOptionGroupDesc, pcv_intOptionGroupCount, pcv_strOptionGroupCount, pcv_strOptionGroupID, pcv_strOptionRequired
Dim xOptionsCnt, pcv_strNumberValidations, pcv_strRemoveFeature
'--> open database connection
call openDB()
%>
<!--#INCLUDE FILE="../pc/viewPrdCode.asp"-->
<%
pcv_strAdminPrefix="1" '// This value is usually retrieved from the header, but this popup does not use the header.
pcv_strRemoveFeature="1" '// Activate the Remove Option Feature.

pidproduct=Request("idproduct")
poption=Request("option")
pidProductOrdered=Request("idProductOrdered")
idOptionArray=Request("idOptionArray")

query="SELECT ProductsOrdered.quantity, ProductsOrdered.unitPrice, ProductsOrdered.QDiscounts, ProductsOrdered.pcPrdOrd_OptionsPriceArray FROM ProductsOrdered WHERE idProductOrdered="&pidProductOrdered
set rstemp=conntemp.execute(query)
if NOT rstemp.eof then
	pta_Quantity = rstemp("quantity")
	pta_Price = rstemp("unitPrice")
	pta_QDiscounts = rstemp("QDiscounts")	
	pcv_strOptionsPriceArray = rstemp("pcPrdOrd_OptionsPriceArray")
end if
set rstemp = nothing

'// Original Price for ReCalulating Quantity Discounts	
pta_Price = ((pta_Price + pcv_strOptionsPriceTotal) + (pta_QDiscounts/int(pta_Quantity)))	

'// Obtain the Existing Options Total
pOpPrices=0
dim pcv_tmpOptionLoopCounter, pcArray_TmpCounter
If len(pcv_strOptionsPriceArray)>0 then

	pcArray_TmpCounter = split(pcv_strOptionsPriceArray,chr(124))
	For pcv_tmpOptionLoopCounter = 0 to ubound(pcArray_TmpCounter)
		pOpPrices = pOpPrices + pcArray_TmpCounter(pcv_tmpOptionLoopCounter)
	Next
	
end if

if NOT isNumeric(pOpPrices) then
	pOpPrices=0
end if								
					
'// Subtract Old Options Total from Original Price								
pta_Price = pta_Price - pOpPrices

'response.Write(pta_Price)
'response.End()

if request.form("action")="update" then
	
	'--> New Product Options
	pcv_intOptionGroupCount = getUserInput(request.Form("OptionGroupCount"),0)
	if IsNull(pcv_intOptionGroupCount) OR pcv_intOptionGroupCount="" then
		pcv_intOptionGroupCount = 0
	end if
	pcv_intOptionGroupCount = cint(pcv_intOptionGroupCount)
	
	xOptionGroupCount = 0
	pcv_strSelectedOptions = ""
	do until xOptionGroupCount = pcv_intOptionGroupCount	
		xOptionGroupCount = xOptionGroupCount + 1
		pcvstrTmpOptionGroup = request.Form("idOption"&xOptionGroupCount)
		if pcvstrTmpOptionGroup <> "" then			
			pcv_strSelectedOptions = pcv_strSelectedOptions & pcvstrTmpOptionGroup & chr(124)	
		end if	
	loop
	' trim the last pipe if there is one
	xStringLength = len(pcv_strSelectedOptions)
	if xStringLength>0 then
		pcv_strSelectedOptions = left(pcv_strSelectedOptions,(xStringLength-1))
	end if
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Get the Options for the item
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcArray_SelectedOptions = Split(pcv_strSelectedOptions,chr(124))
	
	pcv_strOptionsArray = ""
	pcv_strOptionsPriceArray = ""
	pcv_strOptionsPriceArrayCur = ""
	pcv_strOptionsPriceTotal = 0
	xOptionsArrayCount = 0
	pPriceToAdd = 0
	
	For cCounter = LBound(pcArray_SelectedOptions) TO UBound(pcArray_SelectedOptions)
		
		' SELECT DATA SET
		' TABLES: optionsGroups, options, options_optionsGroups
		query = 		"SELECT optionsGroups.optionGroupDesc, options.optionDescrip, options_optionsGroups.price, options_optionsGroups.Wprice "
		query = query & "FROM optionsGroups, options, options_optionsGroups "
		query = query & "WHERE idoptoptgrp=" & pcArray_SelectedOptions(cCounter) & " "
		query = query & "AND options_optionsGroups.idOption=options.idoption "
		query = query & "AND options_optionsGroups.idOptionGroup=optionsGroups.idoptiongroup "	
		
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		
		
		if rs.eof then 
			set rs=nothing
			call closeDb()
			response.redirect "msg.asp?message=42"   	  
		else
			
			xOptionsArrayCount = xOptionsArrayCount + 1
			
			pOptionDescrip=""
			pOptionGroupDesc=""
			pPriceToAdd=""
			pOptionDescrip=rs("optiondescrip")
			pOptionGroupDesc=rs("optionGroupDesc")
			
			If Session("customerType")=1 Then
				pPriceToAdd=rs("Wprice")
				If rs("Wprice")=0 then
					pPriceToAdd=rs("price")
				End If
			Else
				pPriceToAdd=rs("price")
			End If	
			
			'// Generate Our Strings
			if xOptionsArrayCount > 1 then
				pcv_strOptionsArray = pcv_strOptionsArray & chr(124)
				pcv_strOptionsPriceArray = pcv_strOptionsPriceArray & chr(124)
				pcv_strOptionsPriceArrayCur = pcv_strOptionsPriceArrayCur & chr(124)
			end if
			'// Column 4) This is the Array of Product "option groups: options"
			pcv_strOptionsArray = pcv_strOptionsArray & pOptionGroupDesc & ": " & pOptionDescrip
			'// Column 25) This is the Array of Individual Options Prices
			pcv_strOptionsPriceArray = pcv_strOptionsPriceArray & pPriceToAdd
			'// Column 26) This is the Array of Individual Options Prices, but stored as currency "scCurSign & money(pcv_strOptionsPriceTotal) "
			pcv_strOptionsPriceArrayCur = pcv_strOptionsPriceArrayCur & scCurSign & formatnumber(pPriceToAdd, 2)
			'// Column 5) This is the total of all option prices
			pcv_strOptionsPriceTotal = pcv_strOptionsPriceTotal + pPriceToAdd
			
		end if
		
		set rs=nothing
	Next
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Get the Options for the item
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


	'// Add New Options Total from Original Price									
	pta_Price = pta_Price + pcv_strOptionsPriceTotal

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START: ReCalculate Product Quantity Discounts
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	pOrigPrice = pta_Price	
	if pOrigPrice<>"" and isNumeric(pOrigPrice) then
		'// add on a fraction to help vbscript round function up to the actual price
		pOrigPrice=round((pOrigPrice + .001), 2)
	end if
	t=pOrigPrice
	'response.write pOrigPrice
	'response.end
	
	query="SELECT * FROM discountsPerQuantity WHERE idProduct=" &pidproduct& " AND quantityFrom<=" &pta_Quantity& " AND quantityUntil>=" &pta_Quantity
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=conntemp.execute(query)	
	tempNum=0
	
	pOrigPrice = (pOrigPrice*cdbl(pta_Quantity))


	'// Exclude options from the price
	pOrigPriceNoOptions = pOrigPrice
	if isNULL(pcv_strOptionsPriceTotal)=False and len(pcv_strOptionsPriceTotal)>0 then	'// there is pricing on options	
		pOrigPriceNoOptions=(pOrigPrice-(pcv_strOptionsPriceTotal*cdbl(pta_Quantity)))			
	end if
	
	pta_Price=pOrigPrice
	if not rstemp.eof and err.number<>9 then
		'// There are quantity discounts defined for that quantity 
		pDiscountPerUnit=rstemp("discountPerUnit")
		pDiscountPerWUnit=rstemp("discountPerWUnit")
		pPercentage=rstemp("percentage")
		pbaseproductonly=rstemp("baseproductonly")
		if pbaseproductonly="-1" then
			pOrigPrice=pOrigPriceNoOptions
		end if			
		if session("customerType")<>1 then
			if pPercentage="0" then 
				pta_Price=pta_Price - pDiscountPerUnit
				tempNum=tempNum + (pDiscountPerUnit * pta_Quantity)
			else
				pta_Price=pta_Price - ((pDiscountPerUnit/100) * pOrigPrice)
				tempNum=tempNum + ((pDiscountPerUnit/100) * pOrigPrice)
			end if
		else
			if pPercentage="0" then 
				pta_Price=pta_Price - pDiscountPerWUnit
				tempNum=tempNum + (pDiscountPerWUnit * pta_Quantity)
			else
				pta_Price=pta_Price - ((pDiscountPerWUnit/100) * pOrigPrice)
				tempNum=tempNum + ((pDiscountPerWUnit/100) * pOrigPrice)
			end if
		end if
	end if		
	if TempNum="" OR isNULL(TempNum)=True then
		TempNum=0
	end if
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END: ReCalculate Product Quantity Discounts
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	
	'// Unit Price
	tempVar1 = (pta_Price/pta_Quantity)
	
	'response.write tempVar1 &" - "& pcv_strOptionsPriceTotal &" - "& TempNum
	'response.end
	
	pcv_strOptionsArray = replace(pcv_strOptionsArray,"'","''")
	
	query = 		"Update productsOrdered SET "
	query = query & "unitPrice=" & tempVar1 & ", "
	query = query & "QDiscounts=" & TempNum & ", "
	query = query & "pcPrdOrd_SelectedOptions='" & pcv_strSelectedOptions & "', "
	query = query & "pcPrdOrd_OptionsPriceArray='" & pcv_strOptionsPriceArray & "', "
	query = query & "pcPrdOrd_OptionsArray='" & pcv_strOptionsArray & "' "
	query = query & "WHERE idProductOrdered="&pidProductOrdered&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	
	msg="This order has been updated!"
	btn="2"
	response.redirect("options_popup.asp?idProductOrdered="&pidProductOrdered&"&idOptionArray="&pcv_strSelectedOptions&"&idproduct="&pidproduct&"&msg="&msg&"")
	
end if
%>

<html>
<head>
<title>Modify Order Options</title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body>
<div id="pcCPmain" style="width:450px; background-image: none;">
<form name="form1" method="post" action="options_popup.asp" class="pcForms">
<table class="pcCPcontent">
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
	<tr>
		<th colspan="2">Available Product Options</th>
	</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
	<tr>
	<tr>
		<td>
		<p>This product currently has the following options available. When you click 'Save', the options you select <u>will replace ALL existing options</u> for the product, with the current option prices.</p>
    
    <div class="pcCPmessage">NOTE: the product subtotal will change if option prices have changed since the time the order was placed</div>
		</td>
	</tr>
	<tr>
		<td>		
			<table width="98%" border="0" align="center" cellpadding="4" cellspacing="0">
				<% 
				pidproduct=request("idproduct")
				pidProductOrdered=request("idProductOrdered") 
				%>
				
				<tr class="normal">
					<td>		
					<% 
					pcs_OptionsN
					%>			
					<input type=hidden name="idproduct" value="<%=pidproduct%>">
					<input type=hidden name="idProductOrdered" value="<%=pidProductOrdered%>">
					<input type=hidden name="option" value="<%=request.QueryString("option")%>">
					<input type=hidden name="idOptionArray" value="<%=idOptionArray%>">					
					<input type="hidden" name="action" value="update">
					</td>
				</tr>
				<tr> 
					<td colspan="2" align="center"> 
						<% 
						if Request("msg")<>"" then 
							msg=Request("msg")
						end if 
						%>
						<font color=red><%=msg%></font>
					</td>
				</tr>
			  	<tr>
					<td align="center">
					<input type="submit" name="Submit" value="Save">
					<input type="button" name="Back" value="Close Window" onClick="opener.location.reload(); self.close();">
					</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
			</table>			     
		</td>
	</tr>
</table>
</form>
</div>
</body>
</html>
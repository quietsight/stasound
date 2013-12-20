<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% pageTitle="Modify Custom Shipping Option" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/shipFromSettings.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="AdminHeader.asp"-->

<%
Dim rstemp, connTemp, mySQL,tempfrom, tempfromSub

'<<===================If mode is to Activate======================================>>
if request.queryString("mode")="Act" then
	idFlatShipType=request.querystring("idFlatShipType")
	call openDb()
			
	Set rstemp=Server.CreateObject("ADODB.Recordset")
	
	mySQL="UPDATE shipService SET serviceActive = -1 WHERE serviceCode='C"&idFlatShipType&"';"
	Set rstemp=connTemp.execute(mySQL)
	
	connTemp.Close
	Set rstemp=Nothing
	Set connTemp=Nothing
	
	response.redirect request("refer") & "?panel=4"
end if
'<<==================End Activate Shiptype========================================>>

'<<===================If mode is to Inactivate======================================>>
if request.queryString("mode")="InAct" then
	idFlatShipType=request.querystring("idFlatShipType")
	call openDb()
			
	Set rstemp=Server.CreateObject("ADODB.Recordset")
	
	mySQL="UPDATE shipService SET serviceActive = 0 WHERE serviceCode='C"&idFlatShipType&"';"
	Set rstemp=connTemp.execute(mySQL)
	
	connTemp.Close
	Set rstemp=Nothing
	Set connTemp=Nothing
	
	response.redirect request("refer") & "?panel=4"
end if
'<<==================End Inactivate Shiptype========================================>>

'<<===================If mode is to Delete======================================>>
if request.queryString("mode")="DEL" then
	idFlatShipType=request.querystring("idFlatShipType")
	call openDb()
			
	Set rstemp=Server.CreateObject("ADODB.Recordset")
	
	mySQL="DELETE FROM FlatShipTypeRules WHERE idFlatShipType="& idFlatShipType
	Set rstemp=connTemp.execute(mySQL)
	
	mySQL="DELETE FROM FlatShipTypes WHERE idFlatShipType="& idFlatShipType
	Set rstemp=connTemp.execute(mySQL)
	
	mySQL="DELETE FROM shipService WHERE serviceCode='C"&idFlatShipType&"';"
	Set rstemp=connTemp.execute(mySQL)
	
	connTemp.Close
	Set rstemp=Nothing
	Set connTemp=Nothing
	
	response.redirect request("refer") & "?panel=4"
end if
'<<==================End Delete Shiptype========================================>>

	
'<<===================If for submitted, Begin update============================>>
sMode=Request.Form("Submit")
If sMode="Save" Then
	iCnt=10
	'save all inputs in temporary session state
	Session("adminidFlatShipType")=Request("idFlatShipType")
	Session("adminFlatShipTypeDesc")=Request("FlatShipTypeDesc")
	If Session("adminFlatShipTypeDesc")="" then
		msg="<b>Error:</b> You must specify a Name for this shipping type"
		response.redirect "ModFlatShippingRates.asp?idFlatShipType="&Session("adminidFlatShipType")&"&msg="&Server.Urlencode(msg)
	End If
	Session("adminFlatShipTypeDelivery")=Request("FlatShipTypeDelivery")
	Session("adminFlatShipTypePref")=trim(Request("FlatShipTypePref"))
	If Session("adminFlatShipTypePref")="" then
		msg="<b>Error:</font></b> You must specify if this ship type will be calculated by weight, quantity or the sub-total of the cart."
		response.redirect "ModFlatShippingRates.asp?idFlatShipType="&Session("adminidFlatShipType")&"&msg="&Server.Urlencode(msg)
	End If
	If Session("adminFlatShipTypePref")="I" then 
		Session("adminstartIncrement")=trim(Request("startIncrement"))
		if NOT isNumeric(Session("adminstartIncrement")) then
			msg="<b>Error:</b> Your first unit price must be a numeric value."
			response.redirect "ModFlatShippingRates.asp?idFlatShipType="&Session("adminidFlatShipType")&"&msg="&Server.Urlencode(msg)
		end if
		if Session("adminstartIncrement")="" then
			msg="<b>Error:</b> You must specify a first unit price for this shipping option."
			response.redirect "ModFlatShippingRates.asp?idFlatShipType="&Session("adminidFlatShipType")&"&msg="&Server.Urlencode(msg)
		end if
	End If
	Dim ErrIsNumeric
	ErrIsNumeric=0
	for s=1 to iCnt
		Session("adminShippingPrice" & s)=replacecomma(Request("ShippingPrice" & s))
		If NOT IsNumeric(Session("adminShippingPrice"&s)) AND Session("adminShippingPrice"&s)<>"" then
			ErrIsNumeric=ErrIsNumeric+1
		End If
		if Session("adminFlatShipTypePref")="P" then
			Session("adminquantityfrom" & s)=replacecomma(Request("quantityfrom" & s))
			Session("adminquantityto" & s)=replacecomma(Request("quantityto" & s))
		else
			Session("adminquantityfrom" & s)=Request("quantityfrom" & s)
			Session("adminquantityto" & s)=Request("quantityto" & s)
			if Session("adminFlatShipTypePref")="W" Then
				Session("adminquantityfromSub" & s)=Request("quantityfromSub" & s)
				if Request("quantityfromSub" & s)<>"" OR Request("quantityfrom" & s)<>"" then
					if Request("quantityfrom" & s)="" then
						Session("adminquantityfrom" & s)=0	
					end if
					if Request("quantityfromSub" & s)="" then
						Session("adminquantityfromSub" & s)=0	
					end if
				end if
				Session("adminquantitytoSub" & s)=Request("quantitytoSub" & s)
				if Request("quantitytoSub" & s)<>"" OR Request("quantityto" & s)<>"" then
					if Request("quantityto" & s)="" then
						Session("adminquantityto" & s)=0	
					end if
					if Request("quantitytoSub" & s)="" then
						Session("adminquantitytoSub" & s)=0	
					end if
				end if
			end if
		end if
		If NOT IsNumeric(Session("adminquantityfrom"&s)) AND Session("adminquantityfrom"&s)<>"" then
			ErrIsNumeric=ErrIsNumeric+1
		End If
		If NOT IsNumeric(Session("adminquantityto"&s)) AND Session("adminquantityto"&s)<>"" then
			ErrIsNumeric=ErrIsNumeric+1
		End If
		Session("adminidFlatShipTypeRule" & s)=Request("idFlatShipTypeRule" & s)
	next
	If ErrIsNumeric<>0 then
		msg="<b>Error:</b> &nbsp;""To"", ""From"" and ""Ship Rate"" entries must be numeric values only."
		response.redirect "ModFlatShippingRates.asp?idFlatShipType="&Session("adminidFlatShipType")&"&msg="&Server.Urlencode(msg)
		response.End()
	End If
	'validate for commas
	If instr( Session("adminFlatShipTypeDelivery"),",") then
		msg="<b>Error:</b>  The ""Delivery Time"" must not contain commas. "
		response.redirect "ModFlatShippingRates.asp?idFlatShipType="&Session("adminidFlatShipType")&"&msg="&Server.Urlencode(msg)
		response.End()
	End if
	If instr(Session("adminFlatShipTypeDesc") ,",") then
		msg="<b>Error:</b> The ""Shipping Option Name"" must not contain commas."
		response.redirect "ModFlatShippingRates.asp?idFlatShipType="&Session("adminidFlatShipType")&"&msg="&Server.Urlencode(msg)
		response.End()
	End if
	If instr(Session("adminFlatShipTypeDesc") ,"'") or instr(Session("adminFlatShipTypeDesc") ,"""") then
		msg="<b>Error:</b> The ""Shipping Option Name"" must not contain apostrophes or double quotes."
		response.redirect "ModFlatShippingRates.asp?idFlatShipType="&Session("adminidFlatShipType")&"&msg="&Server.Urlencode(msg)
		response.End()
	End if
	if Session("adminFlatShipTypePref")="I" then
		Session("adminstartIncrement")=replacecomma(Session("adminstartIncrement"))
	else
		Session("adminstartIncrement")="0"
	end if
	for c=1 to iCnt
	'check to make sure there are no overlaps
	if Session("adminFlatShipTypePref")="P" or Session("adminFlatShipTypePref")="O" then
		quantityfrom=replacecomma(Request("quantityfrom" & c))
		quantityto=replacecomma(Request("quantityto" & c))
		tempPrice=replacecomma(Request("ShippingPrice" & c))
	else
		if Session("adminFlatShipTypePref")="W" then
			quantityfrom=Request("quantityfrom" & c)
			quantityfromSub=Request("quantityfromSub" & c)
			if quantityfrom<>"" OR quantityfromSub<>"" then
				if quantityfrom="" then
					quantityfrom=0
				end if
				if quantityfromSub="" then
					quantityfromSub=0
				end if
			end if
			quantityto=Request("quantityto" & c)
			quantitytoSub=Request("quantitytoSub" & c)
			if quantityto<>"" OR quantitytoSub<>"" then
				if quantityto="" then
					quantityto=0
				end if
				if quantitytoSub="" then
					quantitytoSub=0
				end if
			end if
			if quantityfrom<>"" and quantityto<>"" and quantityfromSub<>"" and quantitytoSub<>"" then
				if scShipFromWeightUnit="KGS" then
					quantityfrom=(int(quantityfrom*1000)+int(quantityfromSub))
					quantityto=(int(quantityto*1000)+int(quantitytoSub))
				else
					quantityfrom=(int(quantityfrom*16)+int(quantityfromSub))
					quantityto=(int(quantityto*16)+int(quantitytoSub))
				end if
			end if
			tempPrice=replacecomma(Request("ShippingPrice" & c))
		else
			quantityfrom=Request("quantityfrom" & c)
			quantityto=Request("quantityto" & c)
			tempPrice=replacecomma(Request("ShippingPrice" & c))
		end if
	end if
	if quantityfrom <> "" AND quantityto="" AND tempPrice="" then
		msg="<b>Error:</b> You must specify a shipping rate for each tier."
		response.redirect "ModFlatShippingRates.asp?idFlatShipType="&Session("adminidFlatShipType")&"&msg="&Server.Urlencode(msg)
	end if
	if quantityfrom <> "" AND quantityto <> "" AND tempPrice="" then
		msg="<b>Error:</b> You must specify a shipping rate for each tier."
		response.redirect "ModFlatShippingRates.asp?idFlatShipType="&Session("adminidFlatShipType")&"&msg="&Server.Urlencode(msg)
	end if
	if quantityfrom <> "" AND quantityto="" AND tempPrice<>"" then
		msg="<b>Error:</b> You must specify a From, To and a shipping rate for each tier."
		response.redirect "ModFlatShippingRates.asp?idFlatShipType="&Session("adminidFlatShipType")&"&msg="&Server.Urlencode(msg)
	end if
	if quantityfrom="" AND quantityto <> "" AND tempPrice<>"" then
		msg="<b>Error:</b> You must specify a From, To and a shipping rate for each tier."
		response.redirect "ModFlatShippingRates.asp?idFlatShipType="&Session("adminidFlatShipType")&"&msg="&Server.Urlencode(msg)
	end if
	if quantityfrom="" AND quantityto <> "" AND tempPrice="" then
		msg="<b>Error:</b> You must specify a From, To and a shipping rate for each tier."
		response.redirect "E-ModFlatShippingRates.asp?idFlatShipType="&Session("adminidFlatShipType")&"&msg="&Server.Urlencode(msg)
	end if
	if quantityfrom="" AND quantityto="" AND tempPrice<>"" then
		msg="<b>Error:</b> You must specify a From, To and a shipping rate for each tier."
		response.redirect "ModFlatShippingRates.asp?idFlatShipType="&Session("adminidFlatShipType")&"&msg="&Server.Urlencode(msg)
	end if
	if c<>1 then
		d=c-1
		if quantityfrom<>"" then
			if Session("adminFlatShipTypePref")="P" or Session("adminFlatShipTypePref")="O" then
				tempfrom=replacecomma(Request("quantityfrom" & d))
				tempto=replacecomma(Request("quantityto" & d))
			else
				if Session("adminFlatShipTypePref")="W" then
					tempfrom=Request("quantityfrom" & d)
					tempfromSub=Request("quantityfromSub" & d)
					if tempfrom<>"" OR tempfromSub<>"" then
						if tempfrom="" then
							tempfrom=0
						end if
						if tempfromSub="" then
							tempfromSub=0
						end if
					end if
					tempto=Request("quantityto" & d)
					temptoSub=Request("quantitytoSub" & d)
					if tempto<>"" OR temptoSub<>"" then
						if tempto="" then
							tempto=0
						end if
						if temptoSub="" then
							temptoSub=0
						end if
					end if
					if scShipFromWeightUnit="KGS" then
						tempfrom=(int(tempfrom*1000)+int(tempfromSub))
						tempto=(int(tempto*1000)+int(temptoSub))
					else
						tempfrom=(int(tempfrom*16)+int(tempfromSub))
						tempto=(int(tempto*16)+int(temptoSub))
					end if
				else
					tempfrom=Request("quantityfrom" & d)
					tempto=Request("quantityto" & d)
				end if
			end if
			If Cdbl(quantityfrom) <> "" AND Cdbl(quantityfrom)=> Cdbl(tempfrom) AND Cdbl(quantityfrom) > Cdbl(tempto) AND Cdbl(quantityto) <> "" AND Cdbl(quantityto) => Cdbl(quantityfrom) then
			else
				msg="Error: Your entries are conflicting with each other. It appears that you have created two or more entries that contain at least one value that is the same. You cannot have more then one shipping price assigned to any one quantity/weight per ship type."
				response.redirect "ModFlatShippingRates.asp?idFlatShipType="& Session("adminidFlatShipType") &"&msg="&msg
			end if
		end if
	end if
	next
	for i=1 to iCnt
		if Session("adminFlatShipTypePref")="P" or Session("adminFlatShipTypePref")="O" then
			quantityfrom=replacecomma(Request("quantityfrom" & i))
			quantityto=replacecomma(Request("quantityto" & i))
		else
			if Session("adminFlatShipTypePref")="W" then
				tempfrom=Session("adminquantityfrom" & i)
				tempfromSub=Session("adminquantityfromSub" & i)
				if tempfrom<>"" OR tempfromSub<>"" then
					if tempfrom="" then
						tempfrom=0
					end if
					if tempfromSub="" then
						tempfromSub=0
					end if
				end if
				tempto=Session("adminquantityto" & i)
				temptoSub=Session("adminquantitytoSub" & i)
				if tempto<>"" OR temptoSub<>"" then
					if tempto="" then
						tempto=0
					end if
					if temptoSub="" then
						temptoSub=0
					end if
				end if
				if tempFrom<>"" then
				if scShipFromWeightUnit="KGS" then
					quantityfrom=(int(tempfrom)*1000)+int(tempfromSub)
					quantityto=(int(tempto)*1000)+int(temptoSub)
				else
					quantityfrom=(int(tempfrom)*16)+int(tempfromSub)
					quantityto=(int(tempto)*16)+int(temptoSub)
				end if
				end if
			else
				quantityfrom=Request("quantityfrom" & i)
				quantityto=Request("quantityto" & i)
			end if
		end if
		ShippingPrice=replacecomma(Request("ShippingPrice" & i))

		idFlatShipTypeRule=Request("idFlatShipTypeRule" & i)

		If ShippingPrice <> "" AND quantityfrom <> "" AND quantityto <> "" AND idFlatShipTypeRule <> "" Then
			call openDb()
			
			Set rstemp=Server.CreateObject("ADODB.Recordset")
			
			mySQL="UPDATE FlatShipTypeRules SET quantityFrom="& quantityfrom &", quantityTo="& quantityto &", ShippingPrice="& ShippingPrice &", num="& i &" WHERE idFlatShipTypeRule="&idFlatShipTypeRule 

			rstemp.Open mySQL, connTemp
			connTemp.Close
			Set rstemp=Nothing
			Set connTemp=Nothing
		Else
			if idFlatShipTypeRule<>"" then
				call openDb()
				
				Set rstemp=Server.CreateObject("ADODB.Recordset")
				
				mySQL="DELETE FROM FlatShipTypeRules WHERE idFlatShipTypeRule="&idFlatShipTypeRule 
				
				rstemp.Open mySQL, connTemp
				connTemp.Close
				Set rstemp=Nothing
				Set connTemp=Nothing
			end if
		End If
		'======
		
		'======
		If ShippingPrice <> "" AND quantityfrom <> "" AND quantityto <> "" AND idFlatShipTypeRule="" Then
			call openDb()
			
			Set rstemp=Server.CreateObject("ADODB.Recordset")
			mySQL="INSERT INTO FlatShipTypeRules (idFlatShipType, quantityFrom, quantityTo, ShippingPrice, Num) VALUES ("& Session("adminidFlatShipType") &", "&quantityfrom&", "&quantityto&", "& ShippingPrice &", "&i&")"
			rstemp.Open mySQL, connTemp
			connTemp.Close
			Set rstemp=Nothing
			Set connTemp=Nothing
		End If
	next
	
	call openDb()
			
	Set rstemp=Server.CreateObject("ADODB.Recordset")
	
	mySQL="UPDATE FlatShipTypes SET FlatShipTypeDesc='"& replace(Session("adminFlatShipTypeDesc"),"'","''") &"', WQP='"& Session("adminFlatShipTypePref") &"', FlatShipTypeDelivery='"& replace(Session("adminFlatShipTypeDelivery"),"'","''") &"', startIncrement="&Session("adminstartIncrement")&" WHERE idFlatShipType="& Session("adminidFlatShipType")

	rstemp.Open mySQL, connTemp
	
	mySQL="UPDATE shipService SET serviceDescription='"& replace(Session("adminFlatShipTypeDesc"),"'","''")&"' WHERE serviceCode='C"& Session("adminidFlatShipType")&"';"

	rstemp.Open mySQL, connTemp
	
	connTemp.Close
	Set rstemp=Nothing
	Set connTemp=Nothing
	
	'insert into shipService
	If request.form("free")="YES" then
		serviceFree="-1"
		freeamt=replacecomma(request.form("amt"))
			If Not isNumeric(freeamt) then
				freeamt=0
			End If				
	Else
		serviceFree="0"
		freeamt="0"
	End If	
	If request.form("handling")<>"0" AND request.form("handling")<>"" then
		If isNumeric(request.form("handling"))=true then
			thandling=replacecomma(request.form("handling"))
			shfee=request.form("shfee")
		Else
			thandling="0"
			shfee="0"
		End If
	Else
		thandling="0"
		shfee="0"
	End If
	serviceLimitation=request.Form("serviceLimitation")
	if serviceLimitation="" then
		serviceLimitation=0
	end if
	call openDb()
			
	Set rstemp=Server.CreateObject("ADODB.Recordset")
	
	mySQL="UPDATE shipService SET serviceFree="&serviceFree&", serviceFreeOverAmt="&freeAmt&", serviceHandlingFee="&thandling&", serviceShowHandlingFee="&shfee&", serviceLimitation="&serviceLimitation&" WHERE serviceCode='C"&Session("adminidFlatShipType")&"';"

	rstemp.Open mySQL, connTemp
	connTemp.Close
	Set rstemp=Nothing
	Set connTemp=Nothing
	
	Session("adminidFlatShipType")=""
	Session("adminFlatShipTypePref")=""
	Session("adminFlatShipTypeDesc")=""
	Session("adminstartIncrement")=""
	Session("adminFlatShipTypeDelivery")=""
	i=1
	do until i=11
		Session("adminShippingPrice"&i)=""
		Session("adminquantityfrom"&i)=""
		Session("adminquantityto"&i)=""
		Session("adminquantityfromSub"&i)=""
		Session("adminquantitytoSub"&i)=""
		Session("adminidFlatShipTypeRule"&i)=""
	i=i+1
	loop
	referURL=request("refer") & "?panel=4"
	if referURL="" then
		referURL="ModFlatShippingRates.asp?idFlatShipType="&request("idFlatShipType")
	end if
	response.redirect referURL
End If
'<<===================End update============================>>

'<<===================Display info of the shiptype you want to modify============================>>
msg=request.QueryString("msg")
idFlatShipType=request.QueryString("idFlatShipType")
call openDb()
Set rstemp=Server.CreateObject("ADODB.Recordset")

mySQL="SELECT FlatShipTypeDesc, WQP, FlatShipTypeDelivery FROM FlatShipTypes WHERE idFlatShipType="&idFlatShipType&";"
set rstemp=connTemp.execute(mySQL)
FlatShipTypePref=rstemp("WQP")
FlatShipTypeDesc=rstemp("FlatShipTypeDesc")
FlatShipTypeDelivery=rstemp("FlatShipTypeDelivery")
Dim VarShipType, VarFrom, VarTo, VarRate, VarSign, VarSign2, VarSign3, VarRateSign, VarIncrem, VarBased,VarInput, u1, u2, VarWeight
VarShipType=FlatShipTypePref
VarIncrem=0
VarWeight=0
VarPrice=0
select case VarShipType
case "W"
	VarBased="Weight"
	VarInput="<input type=""hidden"" name=""FlatShipTypePref"" value=""W"">"
	VarFrom="From (weight)"
	VarTo="To (weight)"
	VarRate="&nbsp;&nbsp;Ship Rate"
	VarWeight=1
	if scShipFromWeightUnit="KGS" then
		u1="kg"
		u2="g"
	else
		u1="lb"
		u2="oz"
	end if
	VarSign=""
	VarSign2=""
	VarSign3=scCurSign
	VarRateSign=""
case "Q"
	VarBased="Quantity"
	VarInput="<input type=""hidden"" name=""FlatShipTypePref"" value=""Q"">"
	VarFrom="From (units)"
	VarTo="To (units)"
	VarRate="&nbsp;&nbsp;Ship Rate"
	VarSign=""
	VarSign2=""
	VarSign3=scCurSign
	VarRateSign=""
case "P"
	VarBased="Order Amount"
	VarInput="<input type=""hidden"" name=""FlatShipTypePref"" value=""P"">"
	VarFrom="From (price)"
	VarTo="&nbsp;&nbsp;To (price)"
	VarRate="&nbsp;&nbsp;Ship Rate"
	VarSign=scCurSign
	VarSign2=scCurSign
	VarSign3=scCurSign
	VarRateSign=""
	VarPrice=1
case "O"
	VarBased="Percentage of Order Amount"
	VarInput="<input type=""hidden"" name=""FlatShipTypePref"" value=""O"">"
	VarFrom="From (price)"
	VarTo="&nbsp;&nbsp;To (price)"
	VarRate="Percentage"
	VarSign=scCurSign
	VarSign2=scCurSign
	VarSign3=""
	VarRateSign="%"
	VarPrice=1
case "I"
	VarBased="Incremental Calculation"
	VarInput="<input type=""hidden"" name=""FlatShipTypePref"" value=""I"">"
	VarFrom="From (unit)"
	VarTo="To (unit)"
	VarRate="&nbsp;&nbsp;Ship Rate"
	VarSign=""
	VarSign2=""
	VarSign3=scCurSign
	VarRateSign=""
	VarIncrem=1
end select 

Set rstemp=Server.CreateObject("ADODB.Recordset")

mySQL="SELECT FlatShipTypeDesc, FlatShipTypeDelivery, WQP, startIncrement FROM FlatShipTypes WHERE idFlatShipType="&idFlatShipType
Set rstemp=connTemp.execute(mySQL)
%>


<form method="POST" action="ModFlatShippingRates.asp" class="pcForms"> 
<input name="refer" type="hidden" id="refer" value="<%=request("refer")%>"> 
<table class="pcCPcontent">
    <tr>
        <td colspan="4" class="pcCPspacer">
            <% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
        </td>
    </tr>
	<tr> 
		<td colspan="2">Shipping Option Name: </td>
        <td colspan="2"><input type="text" name="FlatShipTypeDesc" value="<%=rstemp("FlatShipTypeDesc")%>"> <span class="pcSmallText">Note: do <strong>not</strong> enter commas, apostrophes, or double quotes</span>
		</td>
	</tr>                
    <tr> 
    	<td colspan="2">Delivery Time:</td>
        <td colspan="2"><input name="FlatShipTypeDelivery" type="text" value="<%=rstemp("FlatShipTypeDelivery")%>" size="30" maxlength="100"> <span class="pcSmallText">This field is optional - Note: do not enter commas</span></td>
    </tr>
    <tr> 
        <td colspan="4">Based on: <b><%=VarBased%></b> <%=VarInput%> <input type="hidden" name="idFlatShipType" size="40" value="<%=idFlatShipType%>"> 
        </td>
    </tr>
	  <% if VarIncrem=1 then %>
      <tr> 
        <td colspan="4">Amount to be charged on first unit: 
          <input name="startIncrement" type="text" size="10" value="<%=rstemp("startIncrement")%>"> <br />Enter the amount to be charged on the first unit, then use the brackets below to specify charges on additional units.
        </td>
      </tr>
      <% end if %>
    <tr> 
        <td colspan="4" class="pcCPspacer"></td>
    </tr>
    <tr> 
        <th>Range:</th>
        <th nowrap><%=VarFrom%></th>
        <th nowrap><%=VarTo%></th>
        <th nowrap><%=VarRate%></th>
    </tr>
    <tr> 
        <td colspan="4" class="pcCPspacer"></td>
    </tr>
    
          <% mySQL="SELECT idFlatShipTypeRule, idFlatshipType, quantityFrom, quantityTo, shippingPrice, num FROM FlatShipTypeRules WHERE idFlatShipType="&idFlatShipType&" ORDER BY num"
set rstemp=connTemp.execute(mySQL)

		r=rstemp("num")
		do until rstemp.eof
			Session("adminShippingPrice" & r)=rstemp("shippingPrice")
			Session("adminquantityfrom" & r)=rstemp("quantityFrom")
			Session("adminquantityto" & r)=rstemp("quantityTo")
			if VarWeight=1 then
				if scShipFromWeightUnit="KGS" then
					tWeightfrom=rstemp("quantityFrom")
					tempfrom=int(tWeightfrom/1000)
					tempfromSub=int(tWeightfrom)-(int(tempfrom)*1000)
					Session("adminquantityfrom" & r)=tempfrom
					Session("adminquantityfromSub" & r)=tempfromSub
					tWeightto=rstemp("quantityto")
					tempto=int(tWeightto/1000)
					temptoSub=int(tWeightto)-(int(tempto)*1000)
					Session("adminquantityto" & r)=tempto
					Session("adminquantitytoSub" & r)=temptoSub
				Else
					tWeightfrom=rstemp("quantityFrom")
					tempfrom=int(tWeightfrom/16)
					tempfromSub=int(tWeightfrom)-(int(tempfrom)*16)
					Session("adminquantityfrom" & r)=tempfrom
					Session("adminquantityfromSub" & r)=tempfromSub
					tWeightto=rstemp("quantityto")
					tempto=int(tWeightto/16)
					temptoSub=int(tWeightto)-(int(tempto)*16)
					Session("adminquantityto" & r)=tempto
					Session("adminquantitytoSub" & r)=temptoSub
				End if
			end if
			Session("adminidFlatShipTypeRule" & r)=rstemp("idFlatShipTypeRule")
			r=r + 1
			rstemp.movenext
		loop
		dim iRcnt
		iRcnt=1
		do until iRcnt=11 %>
          <tr> 
            <td align="right"><%=VarSign%></td>
            <td nowrap> <% if VarWeight=1 then  %> <input name="quantityFrom<%=iRcnt%>" size="2" value="<%=Session("adminquantityfrom" & iRcnt)%>"> 
              &nbsp;<%=u1%>&nbsp; <input name="quantityFromSub<%=iRcnt%>" size="2" value="<%=Session("adminquantityfromSub" & iRcnt)%>"> 
              &nbsp;<%=u2%> <% else
				if VarPrice=1 then %> <input name="quantityFrom<%=iRcnt%>" size="6" value="<% if cstr(Session("adminquantityfrom" & iRcnt))<>"" then response.write money(Session("adminquantityfrom" & iRcnt))%>"> 
              <% else %> <input name="quantityFrom<%=iRcnt%>" size="6" value="<%=Session("adminquantityfrom" & iRcnt)%>"> 
              <% end if %> <% end if %> <input type="hidden" name="idFlatShipTypeRule<%=iRcnt%>" size="40" value="<%=Session("adminidFlatShipTypeRule" & iRcnt)%>"> 
            </td>
            <td nowrap> <% if VarWeight=1 then  %> <input name="quantityto<%=iRcnt%>" size="2" value="<%=Session("adminquantityto" & iRcnt)%>"> 
              &nbsp;<%=u1%>&nbsp; <input name="quantitytoSub<%=iRcnt%>" size="2" value="<%=Session("adminquantitytoSub" & iRcnt)%>"> 
              &nbsp;<%=u2%> <% else %> <%=VarSign2%>&nbsp; <% if VarPrice=1 then %> <input name="quantityTo<%=iRcnt%>" size="6" value="<%if cstr(Session("adminquantityto" & iRcnt)) <> "" then response.write money(Session("adminquantityto" & iRcnt))%>"> 
              <% else %> <input name="quantityTo<%=iRcnt%>" size="6" value="<%=Session("adminquantityto" & iRcnt)%>"> 
              <% end if %> <% end if %> </td>
            <td nowrap> <%=VarSign3%>&nbsp; <% if VarPrice=1 then %> <input name="ShippingPrice<%=iRcnt%>" size="6" value="<%if cstr(Session("adminShippingPrice" & iRcnt))<>"" then response.write money(Session("adminShippingPrice" & iRcnt))%>"> 
              <% else %> <input name="ShippingPrice<%=iRcnt%>" size="10" value="<%=(Session("adminShippingPrice" & iRcnt))%>"> 
              <% end if %> <%=VarRateSign%></td>
          </tr>
          <% iRcnt=iRcnt+1
			loop %>
          <% mySQL="SELECT serviceFree, serviceFreeOverAmt,serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation FROM shipService WHERE serviceCode='C"&idFlatShipType&"';"
Set rstemp=connTemp.execute(mySQL) %>
            <tr> 
                <td colspan="4" class="pcCPspacer"><hr></td>
            </tr>
            <tr> 
            <td>&nbsp;</td>
            <td colspan="3"> <input name="free" type="checkbox" class="clearBorder" id="free" value="YES" <% if rstemp("serviceFree")="-1" then%>checked<%end if%> >
              Offer free shipping for orders over <%=VarSign3%> <input name="amt" type="text" id="amt" size="10" maxlength="10" value="<%=money(rstemp("serviceFreeOverAmt"))%>"> 
            </td>
          </tr>
            <tr> 
                <td colspan="4" class="pcCPspacer"><hr></td>
            </tr>

          <tr> 
            <td>&nbsp;</td>
            <td colspan="3">Add Handling Fee <%=VarSign3%> <input name="handling" type="text" id="handling" size="10" maxlength="10" value="<%=money(rstemp("serviceHandlingFee"))%>"> 
            </td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td colspan="3">
            	<input type="radio" name="shfee" value="-1" <% if rstemp("serviceShowHandlingFee")="-1" then%>checked<%end if%> class="clearBorder">
              Display as a &quot;Shipping &amp; Handling&quot; charge.
              <br> 
              <input type="radio" name="shfee" value="0" <% if rstemp("serviceShowHandlingFee")="0" then%>checked<%end if%> class="clearBorder"> Integrate into shipping rate.</td>
          </tr>
            <tr> 
                <td colspan="4" class="pcCPspacer"><hr></td>
            </tr>
            <tr> 
            <td>&nbsp;</td>
            <td colspan="3">Is there a limitation to which customers will see this shipping option?</td>
            </tr>
            <tr> 
            <td>&nbsp;</td>
            <td colspan="3"><input name="serviceLimitation" type="radio" value="0" checked class="clearBorder"> Show Rate to All</td>
            </tr>
            <tr> 
            <td>&nbsp;</td>
            <td colspan="3"><input type="radio" name="serviceLimitation" value="1" class="clearBorder" <% if rstemp("serviceLimitation")="1" then%>checked<%end if%>> International Only</td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td colspan="3"><input type="radio" name="serviceLimitation" value="2" class="clearBorder" <% if rstemp("serviceLimitation")="2" then%>checked<%end if%>>
                Domestic Only</td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td colspan="3">For stores shipping out of the United States</td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td colspan="3"><input type="radio" name="serviceLimitation" value="3" class="clearBorder" <% if rstemp("serviceLimitation")="3" then%>checked<%end if%>>
Continental United States Only</td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td colspan="3"><input type="radio" name="serviceLimitation" value="4" class="clearBorder" <% if rstemp("serviceLimitation")="4" then%>checked<%end if%>>
Alaska &amp; Hawaii Only</td>
            </tr>
            <tr> 
                <td colspan="4" class="pcCPspacer"><hr></td>
            </tr>
          <tr> 
            <td>&nbsp; </td>
            <td colspan="3"> 
            <input type="submit" name="Submit" value="Save" class="submit2">&nbsp;
			<input type="button" name="Button" value="Back" onclick="javascript:document.location.href='viewShippingOptions.asp'" class="ibtnGrey">
            </td>
          </tr>
        </table>
</form>
<%
Session("adminFlatShipTypePref")=""
Session("adminidFlatShipType")=""
Session("adminFlatShipTypeDesc")=""
Session("adminFlatShipTypeDelivery")=""
Session("adminstartIncrement")=""
i=1
do until i=11
	Session("adminShippingPrice"&i)=""
	Session("adminquantityfrom"&i)=""
	Session("adminquantityto"&i)=""
	Session("adminquantityfromSub"&i)=""
	Session("adminquantitytoSub"&i)=""
	Session("adminidFlatShipTypeRule"&i)=""
	i=i+1
loop
'<<===================End Displaying info of the shiptype you want to modify============================>> %>
<!--#include file="AdminFooter.asp"-->
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% pageTitle="Add New Custom Shipping Option" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ShipFromSettings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="AdminHeader.asp"-->
<% 

Dim rstemp, connTemp, mySQL, ErrMsgNum
ErrMsgNum=0

'<<===================If form submitted============================>>
sMode=Request.Form("Submit")
If sMode="Save" Then
	iCnt=10
	'save all inputs in temporary session state
	Session("adminFlatShipTypeDesc")=Request("FlatShipTypeDesc")
	If Session("adminFlatShipTypeDesc")="" then
		msg="<b>Error:</b> You must specify a Name for this shipping type"
		response.redirect "AddFlatShippingRates.asp?Type="&trim(Request("FlatShipTypePref"))&"&msg="&Server.Urlencode(msg)
	End If
	Session("adminFlatShipTypeDelivery")=Request("FlatShipTypeDelivery")
	Session("adminFlatShipTypePref")=trim(Request("FlatShipTypePref"))
	If Session("adminFlatShipTypePref")="" then
		msg="<b>Error:</b> You must specify if this ship type will be calculated by weight, quantity or the sub-total of the cart."
		ErrMsgNum=ErrMsgNum+1
	End If
	If Session("adminFlatShipTypePref")="I" then 
		Session("adminstartIncrement")=trim(Request("startIncrement"))
		if NOT isNumeric(Session("adminstartIncrement")) then
			msg="<b>Error:</b> Your first unit price must be a numeric value."
			ErrMsgNum=ErrMsgNum+1
		end if
		if Session("adminstartIncrement")="" then
			msg="<b>Error:</b> You must specify a first unit price for this shipping option."
			ErrMsgNum=ErrMsgNum+1
		end if
	End If
	Dim ErrIsNumeric
	ErrIsNumeric=0
	for s=1 to iCnt
		Session("adminShippingPrice" & s)=replacecomma(Request("ShippingPrice" & s))
		If NOT IsNumeric(Session("adminShippingPrice"&s)) AND Session("adminShippingPrice"&s)<>"" then
			ErrIsNumeric=ErrIsNumeric+1
		End If
		if (Session("adminFlatShipTypePref")="P") OR (Session("adminFlatShipTypePref")="O") then
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
	next
	If ErrIsNumeric<>0 then
		msg="<b>Error:</b> &nbsp;""To"", ""From"" and ""Ship Rate"" entries must be numeric values only."
		response.redirect "AddFlatShippingRates.asp?Type="&trim(Request("FlatShipTypePref"))&"&msg="&Server.Urlencode(msg)
		response.End()
	End If
	'validate for commas
	If instr(Session("adminFlatShipTypeDelivery"),",") then
		msg="<b>Error:</b>  The ""Delivery Time"" must not contain commas. "
		response.redirect "AddFlatShippingRates.asp?Type="&trim(Request("FlatShipTypePref"))&"&msg="&Server.Urlencode(msg)
		response.End()
	End if
	If instr(Session("adminFlatShipTypeDesc") ,",") then
		msg="<b>Error:</b> The ""Shipping Option Name"" must not contain commas."
		response.redirect "AddFlatShippingRates.asp?Type="&trim(Request("FlatShipTypePref"))&"&msg="&Server.Urlencode(msg)
		response.End()
	End if
	If instr(Session("adminFlatShipTypeDesc") ,"'") or instr(Session("adminFlatShipTypeDesc") ,"""") then
		msg="<b>Error:</b> The ""Shipping Option Name"" must not contain apostrophes or double quotes."
		response.redirect "AddFlatShippingRates.asp?Type="&trim(Request("FlatShipTypePref"))&"&msg="&Server.Urlencode(msg)
		response.End()
	End if
	if Session("adminFlatShipTypePref")="I" then
		Session("adminstartIncrement")=replacecomma(Session("adminstartIncrement"))
	else
		Session("adminstartIncrement")="0"
	end if
	for c=1 to iCnt
		'check that first row has values
		if Session("adminquantityfrom1") = "" or Session("adminquantityto1") = "" then
			msg="<b>Error</b>: You must specify both a From and a To for the first tier"
			ErrMsgNum=ErrMsgNum+1
		end if
		'check to make sure there are no overlaps
		if (Session("adminFlatShipTypePref")="P") OR (Session("adminFlatShipTypePref")="O") then
			quantityfrom=replacecomma(Request("quantityfrom" & c))
			quantityto=replacecomma(Request("quantityto" & c))
			reqshippingprice=replacecomma(Request("ShippingPrice" & c))
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
				reqshippingprice=replacecomma(Request("ShippingPrice" & c))
			else
				quantityfrom=Request("quantityfrom" & c)
				quantityto=Request("quantityto" & c)
				reqshippingprice=replacecomma(Request("ShippingPrice" & c))
			end if
		end if
		if quantityfrom <> "" AND quantityto="" AND reqshippingprice="" then
			msg="<b>Error:</b> You must specify a shipping rate for each tier"
			ErrMsgNum=ErrMsgNum+1
		end if
		if quantityfrom <> "" AND quantityto <> "" AND reqshippingprice="" then
			msg="<b>Error:</b> You must specify a shipping rate for each tier"
			ErrMsgNum=ErrMsgNum+1
		end if
		if quantityfrom <> "" AND quantityto="" AND reqshippingprice<>"" then
			msg="<b>Error:</b> You must specify a From, To and a shipping rate for each tier"
			ErrMsgNum=ErrMsgNum+1
		end if
		if quantityfrom="" AND quantityto <> "" AND reqshippingprice<>"" then
			msg="<b>Error:</b> You must specify a From, To and a shipping rate for each tier"
			ErrMsgNum=ErrMsgNum+1
		end if
		if quantityfrom="" AND quantityto <> "" AND reqshippingprice="" then
			msg="<b>Error:</b> You must specify a From, To and a shipping rate for each tier"
			ErrMsgNum=ErrMsgNum+1
		end if
		if quantityfrom="" AND quantityto="" AND reqshippingprice<>"" then
			msg="<b>Error:</b> You must specify a From, To and a shipping rate for each tier"
			ErrMsgNum=ErrMsgNum+1
		end if
		if c<>1 then
			d=c-1
			if quantityfrom<>"" then
				if (Session("adminFlatShipTypePref")="P") OR (Session("adminFlatShipTypePref")="O") then
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
					msg="<b>Conflict:</b> Your entries are conflicting with each other. It appears that you have created two or more entries that contain at least one value that is the same. You cannot have more then one shipping price assigned to any one quantity/weight per ship type. "
					ErrMsgNum=ErrMsgNum+1
				end if
			end if
		end if
	next
	'validate for commas
	If instr( Session("adminFlatShipTypeDelivery"),",") then
		msg="<b>Error:</b> The ""Delivery Time"" must not contain commas. "
		ErrMsgNum=ErrMsgNum+1
	End if
	If instr(Session("adminFlatShipTypeDesc") ,",") then
		msg="<b>Error:</b> The ""Shipping Option Name"" must not contain commas. "
		ErrMsgNum=ErrMsgNum+1
	End if
	'if any errors exists, let customer know
	If ErrMsgNum<>0 then
		response.redirect "AddFlatShippingRates.asp?Type="&trim(Request("FlatShipTypePref"))&"&msg="&Server.Urlencode(msg)
	End if
	'Create new FlatShipType
	call openDb()
	strCreateErrMsgNum=0
	err.clear
	err.number=0
			
	Set rstemp=Server.CreateObject("ADODB.Recordset")
	mySQL="INSERT INTO FlatShipTypes (FlatShipTypeDesc, WQP, FlatShipTypeDelivery,startIncrement) VALUES ('"& replace(Session("adminFlatShipTypeDesc"),"'","''") &"', '"& Session("adminFlatShipTypePref") &"','"& replace(Session("adminFlatShipTypeDelivery"),"'","''") &"',"&Session("adminstartIncrement")&");"
	rstemp.Open mySQL, connTemp
	if err.number<>0 then
		strCreateErrMsgNum=strCreateErrMsgNum+1
	end if
	
	if strCreateErrMsgNum=0 then
		mySQL="SELECT idFlatShipType FROM FlatShipTypes WHERE FlatShipTypeDesc='"& replace(Session("adminFlatShipTypeDesc"),"'","''") &"'"
		rstemp.Open mySQL, connTemp
		do until rstemp.eof
			idFlatShipType=rstemp("idFlatShipType")
			rstemp.movenext
		loop
	
		connTemp.Close
		Set rstemp=Nothing
		Set connTemp=Nothing
		
		for i=1 to iCnt
			if (Session("adminFlatShipTypePref")="P") OR (Session("adminFlatShipTypePref")="O") then
				quantityfrom=replacecomma(Request("quantityfrom" & i))
				quantityto=replacecomma(Request("quantityto" & i))
			else
				if Session("adminFlatShipTypePref")="W" then
					tempfrom=Request("quantityfrom" & i)
					tempfromSub=Request("quantityfromSub" & i)
					if tempfrom<>"" OR tempfromSub<>"" then
						if tempfrom="" then
							tempfrom=0
						end if
						if tempfromSub="" then
							tempfromSub=0
						end if
					end if
					tempto=Request("quantityto" & i)
					temptoSub=Request("quantitytoSub" & i)
					if tempto<>"" OR temptoSub<>"" then
						if tempto="" then
							tempto=0
						end if
						if temptoSub="" then
							temptoSub=0
						end if
					end if
					if tempfrom<>"" then
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
			If ShippingPrice <> "" AND quantityfrom <> "" AND quantityto <> "" Then
		
				call openDb()
			
				Set rstemp=Server.CreateObject("ADODB.Recordset")
			
				mySQL="INSERT INTO FlatShipTypeRules (idFlatShipType, quantityFrom, quantityTo, ShippingPrice, Num) VALUES ("& idFlatShipType &", "& quantityfrom &", "& quantityto &", "& ShippingPrice &", "& i &")"
				rstemp.Open mySQL, connTemp
				connTemp.Close
				Set rstemp=Nothing
				Set connTemp=Nothing
			End If
		next

		'insert into shipService
		If request.form("free")="YES" then
			serviceFree="-1"
			freeamt=request.form("amt")
			If Not isNumeric(freeamt) then
				freeamt=0
			End If	
		Else
			serviceFree="0"
			freeamt="0"
		End If	
		If request.form("handling")<>"0" AND request.form("handling")<>"" then
			If isNumeric(request.form("handling"))=true then
				thandling=request.form("handling")
				shfee=request.form("shfee")
			Else
				thandling="0"
				shfee="0"
			End If
		Else
			thandling="0"
			shfee="0"
		End If
		thandling = replacecomma(thandling)
		serviceLimitation=request.Form("serviceLimitation")
		if serviceLimitation="" then
			serviceLimitation=0
		end if
		call openDb()
				
		Set rstemp=Server.CreateObject("ADODB.Recordset")
		
		mySQL="INSERT INTO shipService  (serviceCode, serviceActive, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee,serviceShowHandlingFee, serviceLimitation) VALUES ('C"&idFlatShipType&"', -1, 0, '"&replace(Session("adminFlatShipTypeDesc"),"'","''")&"', "&serviceFree&", "&freeAmt&", "&thandling&","&shfee&","&serviceLimitation&");"
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
		do until i=10
			Session("adminShippingPrice"&sCnt)=""
			Session("adminquantityfrom"&sCnt)=""
			Session("adminquantityto"&sCnt)=""
			Session("adminquantityfromSub"&sCnt)=""
			Session("adminquantitytoSub"&sCnt)=""
			i=i+1
		loop
		response.redirect "viewShippingOptions.asp?panel=4"
	end if
End If
'<<===================End update============================>>

'<<===================Create Form============================>>
%>
<form method="POST" action="AddFlatShippingRates.asp" class="pcForms"> 
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
        <td colspan="2"><input type="text" name="FlatShipTypeDesc" value="<%=Session("adminFlatShipTypeDesc")%>"> <span class="pcSmallText">Note: do <strong>not</strong> enter commas, apostrophes, or double quotes</span>
		</td>
	</tr>                
    <tr> 
    	<td colspan="2">Delivery Time:</td>
        <td colspan="2"><input name="FlatShipTypeDelivery" type="text" value="<%=Session("adminFlatShipTypeDelivery")%>" size="30" maxlength="100"> <span class="pcSmallText">This field is optional - Note: do not enter commas</span></td>
    </tr>
                
	<tr> 
		<td colspan="4">Will be calculated based on:
        	<b> 
			<% Dim VarShipType, VarFrom, VarTo, VarRate, VarSign, VarSign2, VarSign3, VarRateSign, VarIncrem
                VarShipType=request.querystring("Type")
                VarIncrem=0
                VarWeight=0
                select case VarShipType
                case "W"
                    response.write "Weight"
                    response.write "<input type=""hidden"" name=""FlatShipTypePref"" value=""W"">"
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
                    response.write "Quantity"
                    response.write "<input type=""hidden"" name=""FlatShipTypePref"" value=""Q"">"
                    VarFrom="From (units)"
                    VarTo="To (units)"
                    VarRate="&nbsp;&nbsp;Ship Rate"
                    VarSign=""
                    VarSign2=""
                    VarSign3=scCurSign
                    VarRateSign=""
                case "P"
                    response.write "Order Amount"
                    response.write "<input type=""hidden"" name=""FlatShipTypePref"" value=""P"">"
                    VarFrom="From (price)"
                    VarTo="&nbsp;&nbsp;To (price)"
                    VarRate="&nbsp;&nbsp;Ship Rate"
                    VarSign=scCurSign
                    VarSign2=scCurSign
                    VarSign3=scCurSign
                    VarRateSign=""
                case "O"
                    response.write "Percentage of Order Amount"
                    response.write "<input type=""hidden"" name=""FlatShipTypePref"" value=""O"">"
                    VarFrom="From (price)"
                    VarTo="&nbsp;&nbsp;To (price)"
                    VarRate="Percentage"
                    VarSign=scCurSign
                    VarSign2=scCurSign
                    VarSign3=""
                    VarRateSign="%"
                case "I"
                    response.write "Incremental Calculation"
                    response.write "<input type=""hidden"" name=""FlatShipTypePref"" value=""I"">"
                    VarFrom="From (unit)"
                    VarTo="To (unit)"
                    VarRate="&nbsp;&nbsp;Ship Rate"
                    VarSign=""
                    VarSign2=""
                    VarSign3=scCurSign
                    VarRateSign=""
                    VarIncrem=1
                end select %>
                </b>
			</td>
		</tr>
                
	<% if VarIncrem=1 then %>
                    
        <tr> 
            <td colspan="2" valign="top">Charge on first unit:</td>
            <td colspan="2"><input name="startIncrement" type="text" size="10"> <div style="margin-top: 5px;">Enter the amount to be charged on the first unit, then below specify the charges on additional units.</div></td>
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
	<%
	iRcnt=1
		do until iRcnt=11 %>
                
            <tr> 
            <td align="right" nowrap><%=VarSign%></td>
            <td nowrap>
            <% if VarIncrem=1 AND iRcnt=1 then 
            Session("adminquantityfrom1")="2"
            end if %>
            <% if VarWeight=1 then %>
            <input name="quantityFrom<%=iRcnt%>" size="2" value="<%=Session("adminquantityfrom"&iRcnt)%>">
             <%=u1%> 
            <input name="quantityFromSub<%=iRcnt%>" size="2" value="<%=Session("adminquantityfromSub"&iRcnt)%>">
             <%=u2%>
            <% else %>
            <input name="quantityFrom<%=iRcnt%>" size="6" value="<%=Session("adminquantityfrom"&iRcnt)%>">
            <% end if %>
            </td>
            <td nowrap> 
            <%=VarSign2%> 
            <% if VarWeight=1 then %>
            <input name="quantityto<%=iRcnt%>" size="2" value="<%=Session("adminquantityto"&iRcnt)%>">
             <%=u1%> 
            <input name="quantitytoSub<%=iRcnt%>" size="2" value="<%=Session("adminquantitytoSub"&iRcnt)%>">
             <%=u2%>
            <% else %>
            <input name="quantityto<%=iRcnt%>" size="6" value="<%=Session("adminquantityto"&iRcnt)%>">
            <% end if %>
            </td>
            <td nowrap> 
            <%=VarSign3%>  
            <input name="ShippingPrice<%=iRcnt%>" size="10" value="<%=Session("adminShippingPrice"&iRcnt)%>">
            <%=VarRateSign%> </td>
            </tr>
                
<% 		
		iRcnt=iRcnt+1
		loop 
%>
                
    <tr> 
        <td colspan="4" class="pcCPspacer"><hr></td>
    </tr>
    <tr> 
    <td>&nbsp;</td>
    <td colspan="3"> 
        <input name="free" type="checkbox" id="free" value="YES" class="clearBorder"> Offer free shipping for orders over <%=VarSign3%>  
        <input name="amt" type="text" id="amt" size="10" maxlength="10">
    </td>
    </tr>
    <tr> 
        <td colspan="4" class="pcCPspacer"><hr></td>
    </tr>
    <tr> 
    <td>&nbsp;</td>
    <td colspan="3">Add Handling Fee <%=VarSign3%> <input name="handling" type="text" id="handling" size="10" maxlength="10">
    </td>
    </tr>
    <tr> 
    <td>&nbsp;</td>
    <td colspan="3"> 
    <input type="radio" name="shfee" value="-1" checked class="clearBorder"> Display as a "Shipping & Handling" charge.<br>
    <input type="radio" name="shfee" value="0" class="clearBorder"> Integrate into shipping rate.</td>
    </tr>
    <tr> 
        <td colspan="4" class="pcCPspacer"><hr></td>
    </tr>
    <tr> 
        <td>&nbsp;</td>
        <td colspan="3">Is there a limitation to which customers will see this shipping option? </td>
    </tr>
    <tr> 
        <td>&nbsp;</td><td colspan="3">
        <input name="serviceLimitation" type="radio" value="0" checked class="clearBorder"> Show Rate to All</td>
    </tr>
    <tr> 
        <td>&nbsp;</td>
        <td colspan="3"><input type="radio" name="serviceLimitation" value="1" class="clearBorder"> International Only</td>
    </tr>
    <tr> 
        <td>&nbsp;</td>
        <td colspan="3"><input type="radio" name="serviceLimitation" value="2" class="clearBorder"> Domestic Only</td>
    </tr>
    <tr> 
        <td>&nbsp;</td>
        <td colspan="3">For stores shipping out of the United States</td>
    </tr>
    <tr> 
        <td>&nbsp;</td>
        <td colspan="3"><input type="radio" name="serviceLimitation" value="3" class="clearBorder"> Continental United States Only</td>
    </tr>
    <tr> 
        <td>&nbsp;</td>
        <td colspan="3"><input type="radio" name="serviceLimitation" value="4" class="clearBorder"> Alaska &amp; Hawaii Only</td>
    </tr>
    <tr> 
        <td colspan="4" class="pcCPspacer"><hr></td>
    </tr>
    <tr> 
        <td></td>
        <td colspan="3"> 
        <input type="submit" name="Submit" value="Save" class="submit2">&nbsp;
        <input type="button" name="Button" value="Back" onclick="javascript:document.location.href='viewShippingOptions.asp'" class="ibtnGrey">
        </td>
    </tr>
</table>
</form>
<%
	Session("adminidFlatShipType")=""
	Session("adminFlatShipTypePref")=""
	Session("adminFlatShipTypeDesc")=""
	Session("adminFlatShipTypeDelivery")=""
	Session("adminstartIncrement")=""
	Session("adminShippingPrice1")=""
	Session("adminShippingPrice2")=""
	Session("adminShippingPrice3")=""
	Session("adminShippingPrice4")=""
	Session("adminShippingPrice5")=""
	Session("adminShippingPrice6")=""
	Session("adminShippingPrice7")=""
	Session("adminShippingPrice8")=""
	Session("adminShippingPrice9")=""
	Session("adminShippingPrice10")=""
	Session("adminquantityfrom1")=""
	Session("adminquantityfrom2")=""
	Session("adminquantityfrom3")=""
	Session("adminquantityfrom4")=""
	Session("adminquantityfrom5")=""
	Session("adminquantityfrom6")=""
	Session("adminquantityfrom7")=""
	Session("adminquantityfrom8")=""
	Session("adminquantityfrom9")=""
	Session("adminquantityfrom10")=""
	Session("adminquantityto1")=""
	Session("adminquantityto2")=""
	Session("adminquantityto3")=""
	Session("adminquantityto4")=""
	Session("adminquantityto5")=""
	Session("adminquantityto6")=""
	Session("adminquantityto7")=""
	Session("adminquantityto8")=""
	Session("adminquantityto9")=""
	Session("adminquantityto10")=""

'<<===================End Displaying info of the shiptype you want to modify============================>> 
%>
<!--#include file="AdminFooter.asp"-->
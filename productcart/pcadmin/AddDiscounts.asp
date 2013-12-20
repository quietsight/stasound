<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add New Discount by Code" %>
<% Section="specials" %>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/shipFromSettings.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/SQLFormat.txt"--> 
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="AdminHeader.asp"-->

<% Dim rs, connTemp, query

if (pcv_Query="") and (Session("Admin_DC_Status")="") then
	session("admin_DiscFPrds")=""
	session("admin_DiscFCATs")=""
	session("admin_DiscFCusts")=""
	session("admin_DiscShips")=""
	session("admin_DiscFCustPriceCATs")=""
end if

dim intRequestSubmit
intRequestSubmit=0

if request("submit2")<>"" then
	intRequestSubmit=1
	Count1=request("Count1")
	if Count1="" then
		Count1="0"
	end if
	
	For i=1 to Count1
		if request("Pro" & i)="1" then
			IDPro=request("IDPro" & i)
			session("admin_DiscFPrds")=replace(session("admin_DiscFPrds"),IDPro & ",","")
		end if
	Next
end if

if (request("submit3")<>"") OR (request("submit3A")<>"") then
	intRequestSubmit=1
	Count2=request("Count2")
	if Count2="" then
		Count2="0"
	end if
	
	if (request("submit3")<>"") then
		For i=1 to Count2
			if request("CAT" & i)="1" then
				IDCat=request("IDCat" & i)
				SubCat=mid(session("admin_DiscFCATs"),instr(session("admin_DiscFCATs"),IDCat & "-")+len(IDCat & "-"),1)
				session("admin_DiscFCATs")=replace(session("admin_DiscFCATs"),IDCat & "-" & SubCat & ",","")
			end if
		Next
	else
		For i=1 to Count2
			if request("CAT" & i)="1" then
				IDCat=request("IDCat" & i)
				IncSub=request("IncSub" & i)
				if IncSub="" then
					IncSub="0"
				end if
				SubCat=mid(session("admin_DiscFCATs"),instr(session("admin_DiscFCATs"),IDCat & "-")+len(IDCat & "-"),1)
				session("admin_DiscFCATs")=replace(session("admin_DiscFCATs"),IDCat & "-" & SubCat & ",",IDCat & "-" & IncSub & ",")
			end if
		Next
	end if
end if

if request("submit4")<>"" then
	intRequestSubmit=1
	Count3=request("Count3")
	if Count3="" then
		Count3="0"
	end if
	
	For i=1 to Count3
		if request("Cust" & i)="1" then
			IDCust=request("IDCust" & i)
			session("admin_DiscFCusts")=replace(session("admin_DiscFCusts"),IDCust & ",","")
		end if
	Next
end if

if request("submit8")<>"" then
	intRequestSubmit=1
	Count4=request("Count4")
	if Count4="" then
		Count4="0"
	end if
	'// Get array of selected Pricing categories
	pcInt_pCatArray = split(session("admin_DiscFCustPriceCATs"),",")
	For i=1 to Count4
		if request("CustCat" & i)="1" then
			IDCustCat=request("IDCustCat" & i)
			session("admin_DiscFCustPriceCATs")=replace(session("admin_DiscFCustPriceCATs"),IDCustCat & ",","")
		end if
	Next
end if

if Session("Admin_DC_Status")="ok" then
	Session("Admin_DC_Status")=""
	response.redirect "AddDiscounts.asp" & "?" & Session("Admin_DC_Query")
else
	if Request("GoURL")<>"" then
		Session("Admin_DC_Status")="ok"
		tmpPost=""
		For i = 1 to request.form.count
		    fieldName = request.form.key(i)
		    fieldValue = request.form.item(i)
		    if ucase(fieldName)<>"GOURL" then
			    tmpPost=tmpPost & "&" & fieldName & "=" & Server.URLEncode(fieldValue)
		    end if
		Next
		Session("Admin_DC_Query")=pcv_Query & tmpPost
		response.redirect Request("GoURL")
	end if
end if

msg=""
msg=Request("msg")

dim intRequestSubmit1
intRequestSubmit1=Request("Submit1")

pcDisc_PerToFlatDiscount=0
pcDisc_PerToFlatCartTotal=0

if intRequestSubmit1<>"" OR intRequestSubmit=1 then
	discountType=Request("discountType")
	ShipCount=-1
	
	If discountType="1" then
		pricetodiscount=replacecomma(Request("pricetodiscount"))
		if pricetodiscount="" then
			pricetodiscount=0
		end if
		percentagetodiscount=0
		pcDisc_PerToFlatDiscount=0
		pcDisc_PerToFlatCartTotal=0
	Else
		If discountType="2" then
			percentagetodiscount=Request("percentagetodiscount")
			if percentagetodiscount="" then
				percentagetodiscount=0
			end if
			pricetodiscount=0
			pcDisc_PerToFlatCartTotal=replacecomma(request("pcDisc_PerToFlatCartTotal"))
			if NOT isNumeric(pcDisc_PerToFlatCartTotal) or pcDisc_PerToFlatCartTotal="" then
				pcDisc_PerToFlatCartTotal=0
			end if
			if pcDisc_PerToFlatCartTotal=0 then
				pcDisc_PerToFlatDiscount=0
			else
				pcDisc_PerToFlatDiscount=replacecomma(request("pcDisc_PerToFlatDiscount"))
				if NOT isNumeric(pcDisc_PerToFlatDiscount) or pcDisc_PerToFlatDiscount="" then
					pcDisc_PerToFlatDiscount=0
					pcDisc_PerToFlatCartTotal=0
				end if
			end if
		Else
			SHIPS=split(request("pcv_intShippingDiscount"),", ")
			if intRequestSubmit1<>"" then
				call opendb()
				session("admin_DiscShips")=""
				for i=lbound(SHIPS) to ubound(SHIPS)
					if SHIPS(i)<>"" then
						session("admin_DiscShips")=session("admin_DiscShips") & SHIPS(i) & ","
					end if
				next
				pricetodiscount=0
				percentagetodiscount=0
				pcDisc_PerToFlatDiscount=0
				pcDisc_PerToFlatCartTotal=0
				call closedb()
			end if
		End if
	end if
	
	discountcode=replace(Request("discountcode"),"'","")
	discountcode=replace(Request("discountcode"),",","")
	if discountcode="" then
		if intRequestSubmit1<>"" then
			msg="You must supply a discount code. Discount was not added."
		end if
	End if
else
	if request("pcv_intShippingDiscount")<>"" then
		SHIPS=split(request("pcv_intShippingDiscount"),", ")
	end if
	discountType=Request("discountType")
	pricetodiscount=Request("pricetodiscount")
	percentagetodiscount=Request("percentagetodiscount")
	discountcode=replace(Request("discountcode"),"'","")
end if

discountdesc=Request("discountdesc")
if instr(discountdesc,"'")>0 then
	response.redirect("AddDiscounts.asp?msg=The description cannot include an apostrophe.")
end if
discountdesc=replace(discountdesc,"'","''")
discountdesc=replace(discountdesc,",","")
if intRequestSubmit1<>"" then
	if discountdesc="" then
		discountdesc="No Description"
	end if
end if

active=Request("active")
if active="" then
	active=0
end if

expDate=Request("expDate")
if msg="" AND trim(expDate) <> "" then
	if not isDate(expDate) then
		msg="Please enter a valid expiration date"
	end if
	'// Reverse "International" Date Format for comparison and db entry 
	if scDateFrmt="DD/MM/YY" then
		pcArray_FixDate=split(expDate,"/")
		expDate=pcArray_FixDate(1) & "/" & pcArray_FixDate(0) & "/" & pcArray_FixDate(2)
	end if	
end if
	
startDate=Request("startDate")
if msg="" AND trim(startDate) <> "" then
	if not isDate(startDate) then
		msg="Please enter a valid start date"
	end if
	'// Reverse "International" Date Format for comparison and db entry 
	if scDateFrmt="DD/MM/YY" then
		pcArray_FixDate=split(startDate,"/")
		startDate=pcArray_FixDate(1) & "/" & pcArray_FixDate(0) & "/" & pcArray_FixDate(2)
	end if	
end if

'Check that exp date is after start date, if there is one
if msg="" AND trim(startDate) <> "" AND trim(expDate) <> "" then
	if datediff("d", startDate, expDate)<1 then
		msg="Your Start Date must be at least one day before your Expiration Date."
	end if
end if


onetime=Request("onetime")
if onetime="" then
	onetime=0
end if

quantityfrom=Request("quantityfrom")
if quantityfrom="" then
	quantityfrom=0
end if

quantityUntil=Request("quantityuntil")
if quantityuntil="" then
	quantityuntil=9999
end if

weightfrom=Request("weightfrom")
if weightfrom="" then
	weightfrom=0
end if

weightfromSub=Request("weightfromSub")
if weightfromSub="" then
	weightfromSub=0
end if

weightUntil=Request("weightuntil")
weightUntilSub=Request("weightuntilSub")
if weightuntil="" AND weightUntilSub="" then
	weightuntil=999
	weightUntilSub=0
end if

pricefrom=replacecomma(Request("pricefrom"))
if pricefrom="" then
	pricefrom=0
end if

priceUntil=replacecomma(Request("priceuntil"))
if priceuntil="" then
	priceuntil=9999
end if

pcSeparate=Request("pcSeparate")
if pcSeparate="" then
	pcSeparate=0
end if

idproduct=Request("idproduct")
if idproduct="" then
	idproduct=0
end if

pcAuto=Request("pcAuto")
if pcAuto="" then
	pcAuto=0
end if

pcIncExcPrd=Request("IncExcPrd")
if pcIncExcPrd="" then
	pcIncExcPrd=0
end if

pcIncExcCat=Request("IncExcCat")
if pcIncExcCat="" then
	pcIncExcCat=0
end if

pcIncExcCust=Request("IncExcCust")
if pcIncExcCust="" then
	pcIncExcCust=0
end if

pcIncExcCPrice=Request("IncExcCPrice")
if pcIncExcCPrice="" then
	pcIncExcCPrice=0
end if

pcRetail=Request("Retail")
if pcRetail="" then
	pcRetail="0"
end if

pcWholesale=Request("Wholesale")
if pcWholesale="" then
	pcWholesale="0"
end if

if (intRequestSubmit1="Save") and (msg="") then
	if scShipFromWeightUnit="KGS" then
		weightfrom=(int(weightfrom)*1000)+int(weightfromSub)
		WeightUntil=(WeightUntil*1000)+WeightUntilSub
	else
		weightfrom=(weightfrom*16)+weightfromSub
		WeightUntil=(WeightUntil*16)+WeightUntilSub
	end if
	
	call opendb()
	
	if scDB="SQL" then
		strDtDelim="'"
	else
		strDtDelim="#"
	end if

	query="SELECT discountdesc FROM discounts WHERE discountdesc='"& discountdesc &"' "
	set rsValidateDC=server.CreateObject("ADODB.RecordSet")
	set rsValidateDC=conntemp.execute(query)
	if not rsValidateDC.eof then
		response.redirect("AddDiscounts.asp?msg=A discount already exists with that description.")
	end if
	set rsValidateDC=nothing
	
	query="SELECT discountcode FROM discounts WHERE discountcode='"& discountcode &"' "
	set rsValidateDC=server.CreateObject("ADODB.RecordSet")
	set rsValidateDC=conntemp.execute(query)
	if not rsValidateDC.eof then
		response.redirect("AddDiscounts.asp?msg=A discount already exists with that code.")
	end if
	set rsValidateDC=nothing

	query="INSERT INTO discounts (pricetodiscount, percentagetodiscount, discountcode, discountdesc, active,"
	If expDate<>"" Then
		if SQL_Format="1" then
			expDate=(day(expDate)&"/"&month(expDate)&"/"&year(expDate))
		else
			expDate=(month(expDate)&"/"&day(expDate)&"/"&year(expDate))
		end if
		query=query & "expDate, "
	End If
	If startDate<>"" Then
		if SQL_Format="1" then
			startDate=(day(startDate)&"/"&month(startDate)&"/"&year(startDate))
		else
			startDate=(month(startDate)&"/"&day(startDate)&"/"&year(startDate))
		end if
		query=query & "pcDisc_StartDate, "
	End If
	query=query & "onetime, quantityfrom, quantityuntil, weightfrom, weightuntil, pricefrom, priceuntil, idProduct, pcSeparate, pcDisc_Auto, pcRetailFlag, pcWholesaleFlag, pcDisc_PerToFlatCartTotal, pcDisc_PerToFlatDiscount,pcDisc_IncExcPrd,pcDisc_IncExcCat,pcDisc_IncExcCust,pcDisc_IncExcCPrice) VALUES ("&pricetodiscount&","&percentagetodiscount&",'"&discountcode&"','"&discountdesc&"',"&active&","
	If expDate<>"" Then
		query=query & strDtDelim & expDate & strDtDelim &","
	End If
	If startDate<>"" Then
		query=query & strDtDelim & startDate & strDtDelim &","
	End If
	query=query & ""&onetime&","&quantityfrom&","&quantityuntil&","&weightfrom&","& weightuntil&","&pricefrom&","&priceuntil&","&idProduct&","&pcSeparate&","&pcAuto&","&pcRetail&","&pcWholesale&","&pcDisc_PerToFlatCartTotal&","&pcDisc_PerToFlatDiscount&"," & pcIncExcPrd & "," & pcIncExcCat & "," & pcIncExcCust & "," & pcIncExcCPrice & ")"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	set rs=nothing

	query="SELECT iddiscount FROM discounts ORDER BY iddiscount desc;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	if not rs.eof then
		pIDDiscount=rs("iddiscount")
		
		pcArray=split(session("admin_DiscFPrds"),",")
		
		For i=lbound(pcArray) to ubound(pcArray)
		if trim(pcArray(i))<>"" then
		query="INSERT INTO pcDFProds (pcFPro_IDDiscount,pcFPro_IDProduct) VALUES (" & pIDDiscount & "," & pcArray(i) & ");"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		set rs=nothing
		end if
		next
		session("admin_DiscFPrds")=""
		
		pcArray=split(session("admin_DiscFCATs"),",")
		
		For i=lbound(pcArray) to ubound(pcArray)
			if trim(pcArray(i))<>"" then
				pcArray1=split(pcArray(i),"-")
				query="INSERT INTO pcDFCats (pcFCat_IDDiscount,pcFCat_IDCategory,pcFCat_SubCats) VALUES (" & pIDDiscount & "," & pcArray1(0) & "," & pcArray1(1) & ");"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				set rs=nothing
			end if
		next
		session("admin_DiscFCATs")=""
		
		pcArray=split(session("admin_DiscFCusts"),",")
		
		For i=lbound(pcArray) to ubound(pcArray)
			if trim(pcArray(i))<>"" then
				query="INSERT INTO pcDFCusts (pcFCust_IDDiscount,pcFCust_IDCustomer) VALUES (" & pIDDiscount & "," & pcArray(i) & ");"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				set rs=nothing
			end if
		next
		session("admin_DiscFCusts")=""
		
		pcArray=split(session("admin_DiscShips"),",")
		
		For i=lbound(pcArray) to ubound(pcArray)
			if trim(pcArray(i))<>"" then
				query="INSERT INTO pcDFShip (pcFShip_IDDiscount, pcFShip_IDShipOpt) VALUES (" & pIDDiscount & "," & pcArray(i) & ");"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				set rs=nothing
			end if
		next
		session("admin_DiscShips")=""
		
		pcArray=split(session("admin_DiscFCustPriceCATs"),",")
		For i=lbound(pcArray) to ubound(pcArray)
			if trim(pcArray(i))<>"" then
				query="INSERT INTO pcDFCustPriceCats (pcFCPCat_IDDiscount,pcFCPCat_IDCategory) VALUES (" & pIDDiscount & "," & pcArray(i) & ");"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				set rs=nothing
			end if
		next
		session("admin_DiscFCustPriceCATs")=""
		
	end if
	
	call closedb()
	response.redirect "AdminDiscounts.asp?s=1&message="&Server.Urlencode("The electronic coupon was added successfully. You will find it listed below.") 
end if

pcv_Filter=0

call opendb()

if session("admin_DiscFPrds")<>"" then
	session("admin_DiscFCATs")=""
	pcv_Filter=1
else
	if session("admin_DiscFCATs")<>"" then
		session("admin_DiscFPrds")=""
		pcv_Filter=2
	end if
end if

call closedb()
%>

<script>
function Form1_Validator(theForm)
{
	if (theForm.discountdesc.value.indexOf("'")>=0)
	{
		alert("You cannot use apostrophes in the discount description.");
		theForm.discountdesc.focus();
		return (false);
	}
	if (theForm.discountcode.value.indexOf("'")>=0)
	{
		alert("You cannot use apostrophes in the discount code.");
		theForm.discountcode.focus();
		return (false);
	}
	if (theForm.discountdesc.value.indexOf(",")>=0)
	{
		alert("You cannot use commas in the discount description.");
		theForm.discountdesc.focus();
		return (false);
	}
	if (theForm.discountcode.value.indexOf(",")>=0)
	{
		alert("You cannot use commas in the discount code.");
		theForm.discountcode.focus();
		return (false);
	}
	if (theForm.discountcode.value=="")
	{
		alert("Please enter a discount code.");
		theForm.discountcode.focus();
		return (false);
	}
	if (theForm.discountdesc.value=="")
	{
		alert("Please enter a discount description.");
		theForm.discountdesc.focus();
		return (false);
	}
	if (theForm.clicksav.value=="1")
	{
		if ((theForm.discount1.value=="") || (theForm.discount1.value=="0"))
		{
			alert("Please select a discount type.");
			theForm.pricetodiscount.focus();
			return (false);
		}
		if ((theForm.discount1.value=="3") && ((theForm.pcv_intShippingDiscount.value=="0") || (theForm.pcv_intShippingDiscount.value=="")))
		{
			alert("Please select a shipping service.");
			theForm.pcv_intShippingDiscount.focus();
			return( false);
		}
	}
	return (true);
}
</script>

<form method="post" name="hForm" action="Adddiscounts.asp?act=add" onSubmit="return Form1_Validator(this)" class="pcForms">
	<input type="hidden" value="<%=discountType%>" name="discount1">
	<input type="hidden" name="iddiscount" value="<%=iddiscount%>">        
		<table class="pcCPcontent">
            <tr>
                <td colspan="3" class="pcCPspacer">
                    <% ' START show message, if any %>
                        <!--#include file="pcv4_showMessage.asp"-->
                    <% 	' END show message %>
                </td>
            </tr>      

      <tr>
      <tr>
				<th colspan="3">General Information</th>
			</tr>   
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr>  
			<tr>
				<td nowrap width="20%">Discount Description:</td>
				<td colspan="2" width="80%">
					<input name="discountdesc" id="discountdesc" size="30" value="<%=discountdesc%>">&nbsp;Shown to the customer during checkout and on order receipts.
                </td>
			</tr>
			<tr>
				<td>Discount Code:</td>
				<td colspan="2"><input name="discountcode" id"discountcode" size="30" value="<%=discountcode%>">&nbsp;Used by customers to apply the discount to an order
                </td>
			</tr>
			<tr>
				<td colspan="3"><hr></td>
			</tr>
			<tr>
				<td colspan="3">Type of Discount:</td>
			</tr>
			<tr>
				<td colspan="3"> 
					<table width="100%" border="0" cellspacing="0" cellpadding="2">
						<tr>
							<td width="5%" align="right"><input type="radio" name="discountType" value="1" onClick="hForm.discount1.value='1';" <%if discountType=1 then%>checked<%end if%> class="clearBorder"></td>
							<td width="25%">Price Discount</td>
							<td width="70%"><%=scCurSign%>&nbsp;<input name="pricetodiscount" size="16" value="<%=money(pricetodiscount)%>"></td>
						</tr>

						<tr>
							<td align="right">
							<input type="radio" name="discountType" value="2" onClick="hForm.discount1.value='2';" <%if discountType=2 then%>checked<%end if%> class="clearBorder"></td>
							<td>Percent Discount</td>
                            <td>%&nbsp;<input name="percentagetodiscount" size="16" value="<%=percentagetodiscount%>"></td>
						</tr>
						<tr>
                        	<td>&nbsp;</td>
						  	<td colspan="2">
                          	<div class="pcCPnotes">Use the settings below to set values that will allow you to change the &quot;Percent Discount&quot; to a &quot;Price Discount&quot; fee based on the cart total. Setting the values to zero (0) will ignore this feature and always use the percentage set above.</div>
                            <div style="padding: 5px; margin-bottom: 15px;">
                                If the total ordered is over: <%=scCurSign%>&nbsp;
                                <input type="text" name="pcDisc_PerToFlatCartTotal" value="<%=money(pcDisc_PerToFlatCartTotal)%>" size="16">
                                ... switch to the following flat discount: <%=scCurSign%>&nbsp;
                                <input type="text" name="pcDisc_PerToFlatDiscount" value="<%=money(pcDisc_PerToFlatDiscount)%>" size="16">
                            </div>
                          </td>
					  	</tr>
						<tr>
							<td align="right" valign="top">
								<input type="radio" name="discountType" value="3" onClick="hForm.discount1.value='3';" <%if discountType=3 then%>checked<%end if%> class="clearBorder"></td>
							<td valign="top">Free Shipping <div class="pcCPnotes">To select multiple shipping services use the CTRL key.</div>
                            </td>
							<td>
								<select name="pcv_intShippingDiscount" size="5" multiple>
								<% call opendb()
								query="SELECT idshipservice, servicePriority, serviceDescription FROM shipService WHERE serviceActive=-1 ORDER BY servicePriority;"
								set rs=server.CreateObject("ADODB.RecordSet")
								set rs=connTemp.execute(query)
								do until rs.eof%>
									<option value="<%=rs("idshipservice")%>" <%
									if request("pcv_intShippingDiscount")<>"" then
										SHIPS=split(request("pcv_intShippingDiscount"),", ")
										for i=lbound(SHIPS) to ubound(SHIPS)
											if SHIPS(i)<>"" then
												if clng(SHIPS(i))=clng(rs("idshipservice")) then%>selected<%end if
											end if
										next
									end if%>><%=rs("serviceDescription")%></option>
									<%rs.MoveNext
								loop
								set rs=nothing
								call closedb() %>
								</select>
								</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr>  
			<tr><th colspan="3">Status and Expiration</th>
			</tr>
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td>Active:</td>
				<td colspan="2">Yes <input type="checkbox" name="active" value="-1" checked class="clearBorder"></td>
			</tr>
			<tr>
				<td>Start date:</td>
				<td colspan="2">
					<input type="text" name="startDate" value="<%=startDate%>">
					<span class="pcCPnotes">Format: <%=lcase(scDateFrmt)%></span>
                </td>
			</tr>
			<tr>
				<td>Expiration date:</td>
				<td colspan="2">
					<input type="text" name="expDate" value="<%=expDate%>">
					<span class="pcCPnotes">Format: <%=lcase(scDateFrmt)%></span>
                </td>
			</tr>
			<tr>
				<td>One Time Only:</td>
                <td colspan="2">Yes 
				<input type="checkbox" name="onetime" value="-1" <%if onetime="-1" then%>checked<%end if%> class="clearBorder">&nbsp;Check this option for discounts that a customer can only use once</td>
			</tr>
			<tr>
            	<td valign="top">Automatically Apply?</td>
				<td colspan="2">
				  <input name="pcAuto" type="radio" value="1" class="clearBorder" <%if request("pcAuto")="1" then%>checked<%end if%>>Yes&nbsp; 
                  <input name="pcAuto" type="radio" value="0" class="clearBorder" <%if request("pcAuto")<>"1" then%>checked<%end if%>>No&nbsp;|&nbsp;
				  The discount is automatically applied to the order, when applicable, and shown on the order verification page.
               </td>
			</tr>
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr>
			<tr>
				<th colspan="3">Parameters that Restrict Applicability</th>
			</tr>
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr>  
			<tr>
				<td>Order Quantity:</td>
				<td nowrap>From: 
				<input name="quantityFrom" size="6" value="<%=quantityFrom%>">
                </td>
				<td nowrap>To: 
				<input name="quantityUntil" size="6" value="<%=quantityUntil%>">
                </td>
			</tr>
			<%
				if scShipFromWeightUnit="KGS" then
					u1="kg"
					u2="g"
				else
					u1="lb"
					u2="oz"
				end if
			%>
			<tr> 
				<td>Order Weight:</td>
				<td nowrap>From: 
				<input name="WeightFrom" size="2" value="<%=WeightFrom%>">
				&nbsp;<%=u1%>&nbsp;
			
				<input name="WeightFromSub" size="2" value="<%=WeightFromSub%>">
				&nbsp;<%=u2%>
				</td>
				<td nowrap>To: 
				<input name="WeightUntil" size="2" value="<%=WeightUntil%>">
				&nbsp;<%=u1%>&nbsp;
			
				<input name="WeightUntilSub" size="2" value="<%=WeightUntilSub%>">
				&nbsp;
				<%=u2%>
                </td>
			</tr>
			<tr>
				<td>Order Amount:</td>
				<td nowrap>From: <%=scCurSign%> 
				<input name="priceFrom" size="6" value="<%=money(priceFrom)%>">
                </td>
				<td nowrap>To: <%=scCurSign%> 
				<input name="priceUntil" size="6" value="<%=money(priceUntil)%>">
                </td>
			</tr>
			<tr>
				<td colspan="3"><hr></td>
			</tr>
			<tr>
            	<td>Use with others?</td>
				<td colspan="2">
					<input name="pcSeparate" type="radio" value="1" class="clearBorder" <%if request("pcSeparate")="1" then%>checked<%end if%>>Yes&nbsp; 
					<input name="pcSeparate" type="radio" value="0" class="clearBorder" <%if request("pcSeparate")<>"1" then%>checked<%end if%>>No&nbsp;|&nbsp;If "yes" this discount code can be used with other discount codes.</td>
			</tr>
			<tr>
				<td colspan="3"><hr></td>
			</tr>
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td colspan="3"><h2>Filter by Product(s)</h2>
                If no products are selected, the discount applies to orders that contain any products.
                </td>
			</tr>
            <tr>
				<td colspan="3"><input type="radio" name="IncExcPrd" value="0" class="clearBorder" <%if request("IncExcPrd")<>"1" then%>checked<%end if%>> Include selected products&nbsp;&nbsp;<input type="radio" name="IncExcPrd" value="1" class="clearBorder" <%if request("IncExcPrd")="1" then%>checked<%end if%>> Exclude selected products </td>
			</tr>
			<tr>
				<td colspan="3">
					<table class="pcCPcontent" style="width:auto; border:1px solid #E1E1E1;">
						<% pcArray=split(session("admin_DiscFPrds"),",")
						Count1=0
						call opendb()
						For i=lbound(pcArray) to ubound(pcArray)
							if trim(pcArray(i))<>"" then
								Count1=Count1+1
								pIDProduct=pcArray(i)
								query="SELECT description FROM products WHERE Idproduct=" & pIDProduct
								set rs=connTemp.execute(query)
								pName=rs("description")
								set rs=nothing%>
								<tr>
									<td><%=pName%></td><td>
										<input type="checkbox" name="Pro<%=Count1%>" value="1" class="clearBorder">
										<input type="hidden" name="IDPro<%=Count1%>" value="<%=pIDProduct%>">
									</td>
								</tr>
								<%
							end if
						next
						set rs=nothing
						call closedb() %>
						<tr>
							<td colspan="2">
								<%if Count1>0 then%>
									<a href="javascript:checkAllPrd();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAllPrd();">Uncheck All</a>
									<script language="JavaScript">
									<!--
									function checkAllPrd() {
									for (var j = 1; j <= <%=count1%>; j++) {
									box = eval("document.hForm.Pro" + j); 
									if (box.checked == false) box.checked = true;
									}
									}
									
									function uncheckAllPrd() {
									for (var j = 1; j <= <%=count1%>; j++) {
									box = eval("document.hForm.Pro" + j); 
									if (box.checked == true) box.checked = false;
									}
									}
									
									//-->
									</script>
								<%else%>
									No Items to display.
								<%end if%>
                            </td>
						</tr>
					</table>
                </td>
			</tr>
			
			<tr>
				<td colspan="3">
					<%if Count1>0 then%>
						<input type="hidden" name="Count1" value="<%=Count1%>">
						<input type="submit" name="submit2" value="Remove Selected Product(s)">
						&nbsp;
					<%end if%>
					<input type="submit" name="submit5" value="Add Products" onclick="document.hForm.GoURL.value='addprdsToDc.asp?idcode=0';" <%if pcv_Filter=2 then%>disabled="disabled"<%end if%>>				</td>
			</tr>
			<tr>
				<td colspan="3" class="pcCPspacer"><img src="images/pc_admin.gif" width="85" height="19" alt="Or"></td>
			</tr>
			<tr>
				<td colspan="3"><h2>Filter by Category</h2>
                If no categories and no products are selected, the discount applies to orders that contain any products.
                </td>
			</tr>
            <tr>
				<td colspan="3"><input type="radio" name="IncExcCat" value="0" class="clearBorder" <%if request("IncExcCat")<>"1" then%>checked<%end if%>> Include selected categories&nbsp;&nbsp;<input type="radio" name="IncExcCat" value="1" class="clearBorder" <%if request("IncExcCat")="1" then%>checked<%end if%>> Exclude selected categories</td>
			</tr>
			<tr>
				<td colspan="3">
					<table class="pcCPcontent" style="width:auto; border:1px solid #E1E1E1;">
						<% pcArray=split(session("admin_DiscFCATs"),",")
						Count2=0
						call opendb()
						For i=lbound(pcArray) to ubound(pcArray)
							if trim(pcArray(i))<>"" then
								pcArray1=split(pcArray(i),"-")
								Count2=Count2+1
								if Count2=1 then
								%>
								<tr><td></td><td>Include Subcategories</td><td></td></tr>
								<%
								end if
								pIDCat=pcArray1(0)
								query="SELECT categoryDesc FROM categories WHERE IDCategory=" & pIDCat
								set rs=connTemp.execute(query)
								pName=rs("categoryDesc")
								pSubCats=pcArray1(1)
								if pSubCats<>"" then
								else
									pSubCats="0"
								end if%>
								<tr>
									<td><%=pName%></td>
									<td>
									<input type="checkbox" name="IncSub<%=Count2%>" value="1" <%if pSubCats="1" then%>checked<%end if%> class="clearBorder">
									</td>
                                    </td>
									<td>
										<input type="checkbox" name="CAT<%=Count2%>" value="1" class="clearBorder">
										<input type="hidden" name="IDCat<%=Count2%>" value="<%=pIDCAT%>">
                                    </td>
								</tr>
								<%
							end if
						next
						set rs=nothing 
						call closedb() %>
						<tr>
							<td colspan="2">
								<%if Count2>0 then%>
									<a href="javascript:checkAllCat();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAllCat();">Uncheck All</a>
									<script language="JavaScript">
									<!--
									function checkAllCat() {
									for (var j = 1; j <= <%=count2%>; j++) {
									box = eval("document.hForm.CAT" + j); 
									if (box.checked == false) box.checked = true;
									}
									}
									
									function uncheckAllCat() {
									for (var j = 1; j <= <%=count2%>; j++) {
									box = eval("document.hForm.CAT" + j); 
									if (box.checked == true) box.checked = false;
									}
									}
									
									//-->
									</script>
								<%else%>
								No categories to display.
								<%end if%>
                            </td>
						</tr>
					</table>
                </td>
			</tr>
			<tr>
				<td colspan="3">
				<%if Count2>0 then%>
					<input type="hidden" name="Count2" value="<%=Count2%>">
					<input type="submit" name="submit3A" value="Update Selected Categories">
					&nbsp;
					<input type="submit" name="submit3" value="Remove Selected Categories">
					&nbsp;
				<%end if%>
					<input type="submit" name="submit6" value="Add Categories" onclick="document.hForm.GoURL.value='addcatsToDc.asp?idcode=0';" <%if pcv_Filter=1 then%>disabled="disabled"<%end if%>>
                </td>
			</tr>
			<tr>
				<td colspan="3" class="pcCPspacer"><div class="pcCPnotes" style="margin: 10px 0 10px 0;"><img src="images/pcv3_infoIcon.gif" width="16" height="16" alt="Note about this feature" hspace="5">You cannot use <em><strong>Filter by Product(s)</strong></em> and <em><strong>Filter by Category</strong></em> at the same time. <a href="http://wiki.earlyimpact.com/productcart/marketing-discounts_by_code#limiting_applicability" target="_blank">See the documentation</a> for details.</div></td>
			</tr>
			<tr>
				<td colspan="3"><h2>Filter by Customer(s)</h2>
                If no customers are selected, the discount can be used by anyone.
                </td>
			</tr>
            <tr>
				<td colspan="3"><input type="radio" name="IncExcCust" value="0" class="clearBorder" <%if request("IncExcCust")<>"1" then%>checked<%end if%>> Include selected customers&nbsp;&nbsp;<input type="radio" name="IncExcCust" value="1" class="clearBorder" <%if request("IncExcCust")="1" then%>checked<%end if%>> Exclude selected customers</td>
			</tr>
			<tr>
				<td colspan="3">
					<table class="pcCPcontent" style="width:auto; border:1px solid #E1E1E1;">
						<% pcArray=split(session("admin_DiscFCusts"),",")
						Count3=0
						call opendb()
						For i=lbound(pcArray) to ubound(pcArray)
							if trim(pcArray(i))<>"" then
								Count3=Count3+1
								pIDCust=pcArray(i)
								query="SELECT name,lastname FROM customers WHERE idcustomer=" & pIDCust
								set rs=connTemp.execute(query)
								pName=rs("name") & " " & rs("lastname")%>
								<tr>
									<td><%=pName%></td><td>
										<input type="checkbox" name="Cust<%=Count3%>" value="1" class="clearBorder">
										<input type="hidden" name="IDCust<%=Count3%>" value="<%=pIDCust%>">
									</td>
								</tr>
								<%
							end if
						next
						set rs=nothing 
						call closedb() %>
						<tr>
							<td colspan="2">
							<%if Count3>0 then%>
								<a href="javascript:checkAllCust();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAllCust();">Uncheck All</a>
								<script language="JavaScript">
								<!--
								function checkAllCust() {
								for (var j = 1; j <= <%=count3%>; j++) {
								box = eval("document.hForm.Cust" + j); 
								if (box.checked == false) box.checked = true;
								}
								}
								
								function uncheckAllCust() {
								for (var j = 1; j <= <%=count3%>; j++) {
								box = eval("document.hForm.Cust" + j); 
								if (box.checked == true) box.checked = false;
								}
								}
								
								//-->
								</script>
							<%else%>
								No Customers to display.
							<%end if%>
                           </td>
						</tr>
					</table>
                </td>
			</tr>
			<tr>
				<td colspan="3">
					<%if Count3>0 then%>
						<input type="hidden" name="Count3" value="<%=Count3%>">
						<input type="submit" name="submit4" value="Remove Selected Customer(s)">
						&nbsp;
					<%end if%>
					<input type="submit" name="submit7" value="Add Customers" onclick="document.hForm.GoURL.value='addcustsToDc.asp?idcode=0';"></td>
			</tr>
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td colspan="3"><h2>Filter by Customer Pricing Category</h2>
				If no categories are selected, the discount can be used by anyone.</td>
			</tr>
            <tr>
				<td colspan="3"><input type="radio" name="IncExcCPrice" value="0" class="clearBorder" <%if request("IncExcCPrice")<>"1" then%>checked<%end if%>> Include selected customer categories&nbsp;&nbsp;<input type="radio" name="IncExcCPrice" value="1" class="clearBorder" <%if request("IncExcCPrice")="1" then%>checked<%end if%>> Exclude selected customer categories</td>
			</tr>
			<tr>
				<td colspan="3">
					<table class="pcCPcontent" style="width:auto; border:1px solid #E1E1E1;">
						<% pcArray=split(session("admin_DiscFCustPriceCATs"),",")
						Count4=0
						call opendb()
						For i=lbound(pcArray) to ubound(pcArray)
							if trim(pcArray(i))<>"" then
								Count4=Count4+1					
								pIDCustCat=pcArray(i)								
								query="SELECT pcCC_Name FROM pcCustomerCategories WHERE idcustomerCategory=" & pIDCustCat
								SET rs=Server.CreateObject("ADODB.RecordSet")
								SET rs=conntemp.execute(query)
								if NOT rs.eof then 								
								strpcCC_Name=rs("pcCC_Name")
						%>								
								<tr>
									<td><%=strpcCC_Name%></td><td>
										<input type="checkbox" name="CustCat<%=Count4%>" value="1" class="clearBorder">
										<input type="hidden" name="IDCustCat<%=Count4%>" value="<%=pIDCustCat%>">
									</td>
								</tr>
								<%
							end if
						  end if
						next
						set rs = nothing							
						call closedb() %>
						<tr>
							<td colspan="2">
							<%if Count4>0 then%>
								<a href="javascript:checkAllCustCat();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAllCustCat();">Uncheck All</a>
								<script language="JavaScript">
								<!--
								function checkAllCustCat() {
								for (var j = 1; j <= <%=count4%>; j++) {
								box = eval("document.hForm.CustCat" + j); 
								if (box.checked == false) box.checked = true;
								}
								}
								
								function uncheckAllCustCat() {
								for (var j = 1; j <= <%=count4%>; j++) {
								box = eval("document.hForm.CustCat" + j); 
								if (box.checked == true) box.checked = false;
								}
								}
								
								//-->
								</script>
							<%else%>
								No Customer Pricing Categories to display.
							<%end if%>
                           </td>
						</tr>
					</table>
                </td>
			</tr>
			<tr>
				<td colspan="3">
					<%if Count4>0 then%>
						<input type="hidden" name="Count4" value="<%=Count4%>">
						<input type="submit" name="submit8" value="Remove Selected Pricing Categories">
						&nbsp;
					<%end if%>
					<input type="submit" name="submit9" value="Add Pricing Catgeories" onclick="document.hForm.GoURL.value='addCustPriceCatsToDc.asp?idcode=0';">
                </td>
			</tr>
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr>
			<tr>
			  <td colspan="3"><h2>Filter by Customer Type</h2>
If no Types are selected, the discount can be used by anyone.(check all that apply)</td>
		    </tr>
			<tr>
			  <td >Retail Customers:</td>
		      <td ><input type="checkbox" name="Retail" value="1" class="clearBorder" <%if pcRetail ="1" then %> checked <% end if %> >&nbsp;</td>
		      <td >&nbsp;</td>
			</tr>
			<tr>
				<td>Wholesale Customers: </td>
			    <td><input type="checkbox" name="Wholesale" value="1" class="clearBorder"  <%if pcWholeSale ="1" then %> checked <% end if %>>&nbsp;</td>
			    <td>&nbsp;</td>
			</tr>
			<tr>
				<td colspan="3"><hr></td>
			</tr>  
			<tr> 
				<td colspan="3" align="center">
					<input type="submit" name="submit1" value="Save" onclick="hForm.clicksav.value='1';" class="submit2">
					<input type="hidden" name="clicksav" value="">
					&nbsp;
					<input type="button" name="back" value="Back" onClick="javascript:history.back()">
                </td>
			</tr>
		</table>
		<input type="hidden" name="GoURL" value="">
	</form>
<!--#include file="AdminFooter.asp"-->
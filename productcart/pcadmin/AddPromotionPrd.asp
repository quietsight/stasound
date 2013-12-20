<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Section="specials" %>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/shipFromSettings.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/adovbs.inc"-->

<%
Dim rs, connTemp, query, pcvStrProductName, pidProduct

pidProduct=request("idproduct")
if pidProduct="" OR pidProduct="0" then
	response.redirect "PromotionPrdSrc.asp"
end if

call opendb()
query="SELECT description FROM Products WHERE idproduct=" & pidProduct & ";"
set rs=Server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query)
if not rs.eof then
	pcvStrProductName=rs("description")
end if
set rs=nothing
pageIcon="pcv4_icon_promo.png"
pageTitle="Create New Promotion for: <strong>" & pcvStrProductName & "</strong>"
%>

<!--#include file="AdminHeader.asp"-->
<%
session("srcprd_DiscArea")=""

CanNotRun=0

query="SELECT DISTINCT idproduct FROM discountsPerQuantity WHERE idproduct=" & pidProduct & ";"
set rs=connTemp.execute(query)

if not rs.eof then
CanNotRun=1%>
<table class="pcCPcontent">
       <tr>
        	<td colspan="3">
				<div class="pcCPmessage">You cannot add a promotion to this product because it already has quantity discounts assigned to it. <a href="ModDctQtyPrd.asp?idproduct=<%=pIdProduct%>">Review</a></div>
			</td>
		</tr>
		<tr>
			<td>
				<input type="button" name="back" value=" Product Promotions " onclick="location='PromotionPrdSrc.asp';" class="ibtnGrey">
				&nbsp;&nbsp;<input type="button" name="back2" value=" Back to Main menu " onclick="location='menu.asp';" class="ibtnGrey">
			</td>
		</tr>		
</table>
<%set rs=nothing
end if
set rs=nothing

IF CanNotRun=0 THEN
query="SELECT DISTINCT pcCD_idcategory FROM pcCatDiscounts WHERE pcCD_idcategory IN (SELECT DISTINCT idcategory FROM categories_products WHERE idproduct=" & pidProduct & ");"
set rs=connTemp.execute(query)

if not rs.eof then
CanNotRun=1
pIDcategory=rs("pcCD_idcategory")%>
<table class="pcCPcontent">
       <tr>
        	<td colspan="3">
				<div class="pcCPmessage">You cannot add a promotion to this product because it is assigned to one or more categories for which quantity discounts have been entered. <a href="ModDctQtyCat.asp?idcategory=<%=pIDcategory%>">Review</a></div>
			</td>
		</tr>
		<tr>
			<td>
				<input type="button" name="back" value=" Product Promotions " onclick="location='PromotionPrdSrc.asp';" class="ibtnGrey">
				&nbsp;&nbsp;<input type="button" name="back2" value=" Back to Main menu " onclick="location='menu.asp';" class="ibtnGrey">
			</td>
		</tr>		
</table>
<%set rs=nothing
end if
set rs=nothing
END IF

IF CanNotRun=0 THEN

query="SELECT idproduct FROM pcPrdPromotions WHERE idproduct=" & pidProduct & ";"
set rs=connTemp.execute(query)
if not rs.eof then
	set rs=nothing
	call closedb()
	response.redirect "ModPromotionPrd.asp?idproduct=" & pidProduct & "&iMode=start"
end if
set rs=nothing

if (request("second")<>"1") and (Session("Admin_DC_Status")="") then
	session("admin_PromoFPrds")=pidProduct & ","
	session("admin_PromoFCATs")=""
	session("admin_PRFCusts")=""
	session("admin_PRFCustPriceCATs")=""
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
			session("admin_PromoFPrds")=replace(session("admin_PromoFPrds"),IDPro & ",","")
		end if
	Next
end if

if request("submit3")<>"" then
	intRequestSubmit=1
	Count2=request("Count2")
	if Count2="" then
		Count2="0"
	end if
	
	For i=1 to Count2
		if request("CAT" & i)="1" then
			IDCat=request("IDCat" & i)
			SubCat=mid(session("admin_PromoFCATs"),instr(session("admin_PromoFCATs"),IDCat & "-")+len(IDCat & "-"),1)
			session("admin_PromoFCATs")=replace(session("admin_PromoFCATs"),IDCat & "-" & SubCat & ",","")
		end if
	Next
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
			session("admin_PRFCusts")=replace(session("admin_PRFCusts"),IDCust & ",","")
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
	pcInt_pCatArray = split(session("admin_PRFCustPriceCATs"),",")
	For i=1 to Count4
		if request("CustCat" & i)="1" then
			IDCustCat=request("IDCustCat" & i)
			session("admin_PRFCustPriceCATs")=replace(session("admin_PRFCustPriceCATs"),IDCustCat & ",","")
		end if
	Next
end if

if Session("Admin_DC_Status")="ok" then
	Session("Admin_DC_Status")=""
	response.redirect "AddPromotionPrd.asp" & "?" & Session("Admin_DC_Query")
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
		if Left(Session("Admin_DC_Query"),1)="&" then
			Session("Admin_DC_Query")=Mid(Session("Admin_DC_Query"),2,len(Session("Admin_DC_Query")))
		end if
		response.redirect Request("GoURL")
	end if
end if

msg=""
msg=Request("msg")

pcv_ShowMain=1

dim intRequestSubmit1
intRequestSubmit1=Request("Submit1")

if intRequestSubmit1<>"" OR intRequestSubmit=1 then
	discountType=Request("discountType")
	
	If discountType="1" then
		pricetodiscount=replacecomma(Request("pricetodiscount"))
		if not isNumeric(pricetodiscount) then 
			msg="Invalid discount amount."
			pricetodiscount=0
		end if
		if pricetodiscount="" then
			pricetodiscount=0
		end if
		discountvalue=pricetodiscount
		percentagetodiscount=0
	Else
		If discountType="2" then
			percentagetodiscount=Request("percentagetodiscount")
			percentagetodiscount=replace(percentagetodiscount,"%","")
			if not isNumeric(percentagetodiscount) then 
				msg="Invalid percentage discount."
				percentagetodiscount=0
			end if
			if percentagetodiscount="" then
				percentagetodiscount=0
			end if
			pricetodiscount=0
			discountvalue=percentagetodiscount
		End if
	end if
else
	discountType=Request("discountType")
	pricetodiscount=Request("pricetodiscount")
	percentagetodiscount=Request("percentagetodiscount")
end if

qtytrigger=request("qtytrigger")
if qtytrigger="" OR qtytrigger<="0" OR (NOT (IsNumeric(qtytrigger))) then
	qtytrigger="1"
end if

applyunits=request("applyunits")
if applyunits="" OR applyunits<"0" OR (NOT (IsNumeric(applyunits))) then
	applyunits="1"
end if

promomsg=request("promomsg")
confirmmsg=request("confirmmsg")
descmsg=request("descmsg")
prinactive=request("prinactive")

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
	if promomsg<>"" then
		promomsg=pcf_ReplaceCharacters(promomsg)
		promomsg=pcf_ReplaceQuotes(promomsg)
	end if
	if confirmmsg<>"" then
		confirmmsg=pcf_ReplaceCharacters(confirmmsg)
		confirmmsg=pcf_ReplaceQuotes(confirmmsg)
	end if
	if descmsg<>"" then
		descmsg=pcf_ReplaceCharacters(descmsg)
		descmsg=pcf_ReplaceQuotes(descmsg)
	end if
	if prinactive="" OR IsNull(prinactive) then
		prinactive=0
	end if
	
	query="INSERT INTO pcPrdPromotions (idproduct,pcPrdPro_QtyTrigger,pcPrdPro_DiscountType,pcPrdPro_DiscountValue,pcPrdPro_ApplyUnits,pcPrdPro_PromoMsg,pcPrdPro_ConfirmMsg,pcPrdPro_SDesc,pcPrdPro_Inactive,pcPrdPro_IncExcCust,pcPrdPro_IncExcCPrice,pcPrdPro_RetailFlag,pcPrdPro_WholesaleFlag)"
	query=query & " VALUES (" & pIdProduct & "," & qtytrigger & "," & discountType & "," & discountvalue & "," & applyunits & ",'" & promomsg & "','" & confirmmsg & "','" & descmsg & "'," & prinactive & "," & pcIncExcCust & "," & pcIncExcCPrice & "," & pcRetail & "," & pcWholesale & ");"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	set rs=nothing

	query="SELECT pcPrdPro_id FROM pcPrdPromotions ORDER BY pcPrdPro_id desc;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	if not rs.eof then
		pIDPromo=rs("pcPrdPro_id")
		
		pcArray=split(session("admin_PromoFPrds"),",")
		
		For i=lbound(pcArray) to ubound(pcArray)
		if trim(pcArray(i))<>"" then
		query="INSERT INTO pcPPFProducts (pcPrdPro_id,idproduct) VALUES (" & pIDPromo & "," & pcArray(i) & ");"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		set rs=nothing
		end if
		next
		session("admin_PromoFPrds")=""
		
		pcArray=split(session("admin_PromoFCATs"),",")
		
		For i=lbound(pcArray) to ubound(pcArray)
			if trim(pcArray(i))<>"" then
				pcArray1=split(pcArray(i),"-")
				query="INSERT INTO pcPPFCategories (pcPrdPro_id,idcategory,pcPPFCats_IncSubCats) VALUES (" & pIDPromo & "," & pcArray1(0) & "," & pcArray1(1) & ");"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				set rs=nothing
			end if
		next
		session("admin_PromoFCATs")=""
		
		pcArray=split(session("admin_PRFCusts"),",")
		
		For i=lbound(pcArray) to ubound(pcArray)
			if trim(pcArray(i))<>"" then
				query="INSERT INTO pcPPFCusts (pcPrdPro_id,IDCustomer) VALUES (" & pIDPromo & "," & pcArray(i) & ");"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				set rs=nothing
			end if
		next
		session("admin_PRFCusts")=""
		
		pcArray=split(session("admin_PRFCustPriceCATs"),",")
		For i=lbound(pcArray) to ubound(pcArray)
			if trim(pcArray(i))<>"" then
				query="INSERT INTO pcPPFCustPriceCats (pcPrdPro_id,idCustomerCategory) VALUES (" & pIDPromo & "," & pcArray(i) & ");"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				set rs=nothing
			end if
		next
		session("admin_PRFCustPriceCATs")=""
		
	end if
	
	call closedb()
	%>
	<table class="pcCPcontent">
       <tr>
        	<td colspan="3">
				<div class="pcCPmessageSuccess">Promotion added successfully! <a href="../pc/viewPrd.asp?idproduct=<%=pIdProduct%>" target="_blank">Preview it</a> in the storefront &gt;&gt;</div>
			</td>
		</tr>
		<tr>
			<td>
				<input type="button" name="back" value=" Product Promotions " onclick="location='PromotionPrdSrc.asp';" class="ibtnGrey">
				&nbsp;&nbsp;<input type="button" name="back1" value=" View/Edit the Promotion " onclick="location='ModPromotionPrd.asp?idproduct=<%=pidproduct%>&iMode=start';" class="ibtnGrey">
				&nbsp;&nbsp;<input type="button" name="back2" value=" Back to Main menu " onclick="location='menu.asp';" class="ibtnGrey">
			</td>
		</tr>		
	</table>	
	<%
	pcv_ShowMain=0
end if

IF pcv_ShowMain=1 THEN

pcv_Filter=0

if session("admin_PromoFPrds")<>"" then
	session("admin_PromoFCATs")=""
	pcv_Filter=1
else
	if session("admin_PromoFCATs")<>"" then
		session("admin_PromoFPrds")=""
		pcv_Filter=2
	end if
end if
%>

<script>
function Form1_Validator(theForm)
{
	if (theForm.clicksav.value=="1")
	{
	if ((theForm.discount1.value=="") || (theForm.discount1.value=="0"))
	{
		alert("Please select a discount type.");
		return(false);
	}
	}
	return(true);
}
</script>

<form method="post" name="hForm" action="AddPromotionPrd.asp?act=add" onSubmit="return Form1_Validator(this)" class="pcForms">
	<input type=hidden value="1" name="second">
	<input type=hidden value="<%=discountType%>" name="discount1">
	<input type="hidden" name="idproduct" value="<%=pIdProduct%>">
	<input type="hidden" name="idcode" value="0">
		<table class="pcCPcontent">
            <tr>
                <td colspan="3" class="pcCPspacer">
                    <% ' START show message, if any %>
                        <!--#include file="pcv4_showMessage.asp"-->
                    <% 	' END show message %>
                </td>
            </tr>
      		<tr>
				<th colspan="3">Promotion Settings</th>
			</tr>   
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr>
			<%
			promomsg=pcf_ReplaceQuotes(promomsg)
			confirmmsg=pcf_ReplaceQuotes(confirmmsg)
			descmsg=pcf_ReplaceQuotes(descmsg)
			%>
			<tr>
				<td nowrap width="20%">Promotion Message:</td>
				<td colspan="2" width="80%">
					<input name="promomsg" size="60" value="<%=promomsg%>">&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=214')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
				</td>
			</tr>
            <tr>
				<td nowrap width="20%">Confirmation Message:</td>
				<td colspan="2" width="80%">
					<input name="confirmmsg" size="60" value="<%=confirmmsg%>">&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=215')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
				</td>
			</tr>
            <tr>
				<td nowrap width="20%">Short Description:</td>
				<td colspan="2" width="80%">
					<input name="descmsg" size="60" value="<%=descmsg%>">&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=216')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
				</td>
			</tr>
            <tr>
				<td nowrap width="20%">Quantity Trigger:</td>
				<td colspan="2" width="80%">
					<input name="qtytrigger" size="4" value="<%=qtytrigger%>">&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=218')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
				</td>
			</tr>
			<tr>
				<td nowrap width="20%">Apply to next N Units:</td>
				<td colspan="2" width="80%">
					<input name="applyunits" size="4" value="<%=applyunits%>"> unit(s)&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=217')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
				</td>
			</tr>
			<tr>
				<td nowrap width="20%" align="right"><input type="checkbox" name="prinactive" value="1" <%if prinactive="1" then%>checked<%end if%> class="clearBorder"></td>
				<td colspan="2" width="80%">
					Inactive
				</td>
			</tr>
			<tr>
				<td colspan="3">Type of Discount:</td>
			</tr>
			<tr>
				<td colspan="3"> 
					<table width="100%" border="0" cellspacing="0" cellpadding="2">
						<tr>
							<td width="5%" align="right"><input type="radio" name="discountType" value="1" onClick="hForm.discount1.value='1';" <%if discountType=1 then%>checked<%end if%> class="clearBorder"></td>
							<td width="20%">Price Discount</td>
							<td width="75%"><%=scCurSign%> <input name="pricetodiscount" size="16" value="<%=money(pricetodiscount)%>">
                            &nbsp;<span class="pcSmallText"><a href="FindProductType.asp?idproduct=<%=pIdProduct%>" target="_blank">Look up current product prices</a></span>.</td>
						</tr>

						<tr>
							<td align="right"><input type="radio" name="discountType" value="2" onClick="hForm.discount1.value='2';" <%if discountType=2 then%>checked<%end if%> class="clearBorder"></td>
							<td>Percent Discount</td>
                            <td width="70%">% <input name="percentagetodiscount" size="16" value="<%=percentagetodiscount%>"></td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr>
			<tr>
				<th colspan="3">Parameters that Restrict Applicability</th>
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
						<% pcArray=split(session("admin_PRFCusts"),",")
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
					<input type="submit" name="submit7" value="Add Customers" onclick="document.hForm.GoURL.value='addcustsToPR.asp?idcode=0&idproduct=<%=pIDProduct%>';"></td>
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
						<% pcArray=split(session("admin_PRFCustPriceCATs"),",")
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
					<input type="submit" name="submit9" value="Add Pricing Categories" onclick="document.hForm.GoURL.value='addCustPriceCatsToPR.asp?idcode=0&idproduct=<%=pIDProduct%>';">
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
            <!-- COMMENT OUT FOR v4
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr>  
			<tr>
				<th colspan="3">Parameters that Restrict Applicability&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=219')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></th>
			</tr>
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr> 
 			<tr>
				<td colspan="3">You can apply a promotion to either one or more products OR to one or more categories. So if a product is listed below, the button to add a category is hidden, and vice versa.</td>
			</tr>            
			<tr>
				<td colspan="3"><h2>Product(s)</h2></td>
			</tr>
			<tr>
				<td colspan="3">
					<table class="pcCPcontent" style="width: 400px; border: 1px dashed #e1e1e1;">
						<% pcArray=split(session("admin_PromoFPrds"),",")
						Count1=0
						call opendb()
						For i=lbound(pcArray) to ubound(pcArray)
							if trim(pcArray(i))<>"" then
								Count1=Count1+1
								pIDPro=pcArray(i)
								query="SELECT description FROM products WHERE Idproduct=" & pIDPro
								set rs=connTemp.execute(query)
								pName=rs("description")
								set rs=nothing%>
								<tr>
									<td><a href="FindProductType.asp?id=<%=pIDPro%>" target="_blank"><%=pName%></a></td>
                                    <td align="right">
                                -->
                                <!-- COMMENT OUT FOR v4: Turn this next input field from a checkbox into a hidden field -->
										<input type="hidden" name="Pro<%=Count1%>" value="1" class="clearBorder">
										<input type="hidden" name="IDPro<%=Count1%>" value="<%=pIDPro%>">
                                <!-- COMMENT OUT FOR v4
									</td>
								</tr>
								<%
							end if
						next
						set rs=nothing
						call closedb() %>
						<tr>
							<td colspan="2"<%if Count1>0 then%>align="right"<%end if%>>
								<%if Count1>0 then%>
									<a href="javascript:checkAllPrd();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAllPrd();">Uncheck All</a>
									<script language="JavaScript">
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
                    <%if pcv_Filter=2 then%>
                    	To apply the promotion to one or more products, please remove the categories to which it currently applies.
                    <%else%>
					<input type="submit" name="submit5" value="Add Products" onclick="document.hForm.GoURL.value='addprdsToPR.asp?idcode=0&idproduct=<%=pIDProduct%>';">
                    <%end if%>
                </td>
			</tr>
			<tr>
				<td colspan="3"><img src="images/pc_admin.gif" width="85" height="19" alt="Alternative selections" vspace="15"></td>
			</tr> 
			<tr>
				<td colspan="3"><h2>Categories</h2></td>
			</tr>
			<tr>
				<td colspan="3">
					<table class="pcCPcontent" style="width: 400px; border: 1px dashed #e1e1e1;">
						<% pcArray=split(session("admin_PromoFCATs"),",")
						Count2=0
						call opendb()
						For i=lbound(pcArray) to ubound(pcArray)
							if trim(pcArray(i))<>"" then
								pcArray1=split(pcArray(i),"-")
								Count2=Count2+1
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
									<td>
									<a href="modcata.asp?idcategory=<%=pIDCat%>" target="_blank"><%=pName%></a>&nbsp;
									<%if pSubCats="1" then%>
										(including its subcategories)
									<%end if%>
                                    </td>
									<td align="right">
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
							<td colspan="2"<%if Count2>0 then%>align="right"<%end if%>>
								<%if Count2>0 then%>
									<a href="javascript:checkAllCat();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAllCat();">Uncheck All</a>
									<script language="JavaScript">
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
				<%if Count2>0 then%>
					<input type="hidden" name="Count2" value="<%=Count2%>">
					<input type="submit" name="submit3" value="Remove Selected Categories">
					&nbsp;
				<%end if%>
				<%if pcv_Filter=1 then%>
                    To apply the promotion to one or more categories, please remove the products to which it currently applies.
                <%else%>
					<input type="submit" name="submit6" value="Add Categories" onclick="document.hForm.GoURL.value='addcatsToPR.asp?idcode=0&idproduct=<%=pIDProduct%>';">
                <%end if%>
                </td>
			</tr>
            -->
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
<%END IF 'pcv_ShowMain=1%>
<%END IF 'CanNotRun%>
<!--#include file="AdminFooter.asp"-->
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
Dim rs, connTemp, query, pcvStrCatName, categoryName

session("srcCat_DiscArea")=""

pidCategory=request("idCategory")
if pidCategory="" OR pidCategory="0" then
	response.redirect "PromotionCatSrc.asp"
end if

categoryName=""

call opendb()
query="SELECT categoryDesc FROM Categories WHERE idCategory=" & pidCategory & ";"
set rs=Server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query)
if not rs.eof then
	pcvStrCatName=rs("categoryDesc")	
	categoryName="Category Name: <a href='modCata.asp?idcategory=" & pidCategory & "'><b>" & rs("categoryDesc") & "</b></a>"
end if
set rs=nothing

pageTitle="Create New Promotion for: " & pcvStrCatName

query="SELECT idCategory FROM pcCatPromotions WHERE idCategory=" & pidCategory & ";"
set rs=connTemp.execute(query)
if not rs.eof then
	set rs=nothing
	call closedb()
	response.redirect "ModPromotionCat.asp?idCategory=" & pidCategory & "&iMode=start"
end if
set rs=nothing

%>
<!--#include file="AdminHeader.asp"-->
<% 

if (request("second")<>"1") and (Session("Admin_DC_Status")="") then
	session("admin_PromoFPrds")=""
	session("admin_PromoFCATs")=pidCategory & "-1,"
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

if Session("Admin_DC_Status")="ok" then
	Session("Admin_DC_Status")=""
	response.redirect "AddPromotionCat.asp" & "?" & Session("Admin_DC_Query")
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
	query="INSERT INTO pcCatPromotions (idCategory,pcCatPro_QtyTrigger,pcCatPro_DiscountType,pcCatPro_DiscountValue,pcCatPro_ApplyUnits,pcCatPro_PromoMsg,pcCatPro_ConfirmMsg,pcCatPro_SDesc)"
	query=query & " VALUES (" & pidCategory & "," & qtytrigger & "," & discountType & "," & discountvalue & "," & applyunits & ",'" & promomsg & "','" & confirmmsg & "','" & descmsg & "');"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	set rs=nothing

	query="SELECT pcCatPro_id FROM pcCatPromotions ORDER BY pcCatPro_id desc;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	if not rs.eof then
		pIDPromo=rs("pcCatPro_id")
		
		pcArray=split(session("admin_PromoFPrds"),",")
		
		For i=lbound(pcArray) to ubound(pcArray)
		if trim(pcArray(i))<>"" then
		query="INSERT INTO pcCPFProducts (pcCatPro_id,idProduct) VALUES (" & pIDPromo & "," & pcArray(i) & ");"
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
				query="INSERT INTO pcCPFCategories (pcCatPro_id,idcategory,pcCPFCats_IncSubCats) VALUES (" & pIDPromo & "," & pcArray1(0) & "," & pcArray1(1) & ");"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				set rs=nothing
			end if
		next
		session("admin_PromoFCATs")=""
		
	end if
	
	call closedb()
	%>
	<table class="pcCPcontent">
		<tr>
        	<td colspan="3">
				<div class="pcCPmessageSuccess">Promotion added successfully! <a href="../pc/viewcategories.asp?idcategory=<%=pidCategory%>" target="_blank">Preview it</a> in the storefront &gt;&gt;</div>
			</td>
		</tr>
		<tr>
			<td>
				<input type="button" name="back" value=" Category Promotions " onclick="location='PromotionCatSrc.asp';" class="ibtnGrey">
				&nbsp;&nbsp;<input type="button" name="back1" value=" View/Edit the Promotion " onclick="location='ModPromotionCat.asp?idcategory=<%=pidcategory%>&iMode=start';" class="ibtnGrey">
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

<form method="post" name="hForm" action="AddPromotionCat.asp?act=add" onSubmit="return Form1_Validator(this)" class="pcForms">
	<input type=hidden value="1" name="second">
	<input type=hidden value="<%=discountType%>" name="discount1">
	<input type="hidden" name="idCategory" value="<%=pidCategory%>">
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
				<td colspan="3">Type of Discount:</td>
			</tr>
			<tr>
				<td colspan="3"> 
					<table width="100%" border="0" cellspacing="0" cellpadding="2">
						<tr>
							<td width="5%" align="right"><input type="radio" name="discountType" value="1" onClick="hForm.discount1.value='1';" <%if discountType=1 then%>checked<%end if%> class="clearBorder"></td>
							<td width="20%">Price Discount</td>
							<td width="75%"><%=scCurSign%> <input name="pricetodiscount" size="16" value="<%=money(pricetodiscount)%>">
                            &nbsp;<span class="pcSmallText"><a href="editCategories.asp?lid=<%=pidCategory%>">Look up products in the category</a></span>.</td>
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
										<input type="checkbox" name="Pro<%=Count1%>" value="1" class="clearBorder">
										<input type="hidden" name="IDPro<%=Count1%>" value="<%=pIDPro%>">
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
                    <%if pcv_Filter=2 then%>
                    	To apply the promotion to one or more products, please remove the categories to which it currently applies.
                    <%else%>
					<input type="submit" name="submit5" value="Add Products" onclick="document.hForm.GoURL.value='addprdsToPR_cat.asp?idcode=0&idCategory=<%=pidCategory%>';">
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
					<input type="submit" name="submit6" value="Add Categories" onclick="document.hForm.GoURL.value='addcatsToPR_cat.asp?idcode=0&idCategory=<%=pidCategory%>';">
                <%end if%>
                </td>
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
<%END IF 'pcv_ShowMain=1%>
<!--#include file="AdminFooter.asp"-->
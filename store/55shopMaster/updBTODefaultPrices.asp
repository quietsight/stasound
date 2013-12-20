<%@LANGUAGE="VBSCRIPT"%>
<%On Error Resume Next%>
<% pageTitle = "Quickly update all BTO Product default prices" %>
<% Section = "" %>
<%PmAdmin=2%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp" -->
<!--#include file="pcCalculateBTODefaultPrices.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="inc_UpdateDates.asp" -->
<!--#include file="AdminHeader.asp"-->
<%dim conntemp, rs, pcArr, intCount,i,j,query,pcArr1, intCount1,pcArr2, intCount2

IF request("action")="upd" THEN
	
	call opendb()
	
	query="SELECT idcustomerCategory, pcCC_WholesalePriv, pcCC_CategoryType, pcCC_ATB_Percentage, pcCC_ATB_Off FROM pcCustomerCategories;"
	Set rs=Server.CreateObject("ADODB.RecordSet")
	Set rs=conntemp.execute(query)
	intCount2=-1
	if not rs.eof then
		pcArr2=rs.getRows()		
		intCount2=ubound(pcArr2,2)
	end if
	set rs=nothing
	
	query="SELECT idproduct,price,bToBPrice,NoPrices FROM products WHERE serviceSpec<>0 AND removed=0;"
	Set rs=Server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	if not rs.eof then
		pcArr=rs.getRows()
		set rs=nothing
		intCount=ubound(pcArr,2)
		For i=0 to intCount
			pidProduct=pcArr(0,i)
			pPrice=pcArr(1,i)
			pWPrice=pcArr(2,i)
			if pWPrice=0 then
				pWPrice=pPrice
			end if
			pBtoBPrice=pWPrice
			pnoprices=pcArr(3,i)
			
			save_pPrice=pPrice
			save_pBtoBPrice=pBtoBPrice

			query="SELECT configSpec_products.price,configSpec_products.Wprice,products.pcprod_minimumqty FROM (configSpec_products INNER JOIN products ON configSpec_products.configProduct = products.idProduct) INNER JOIN categories ON configSpec_products.configProductCategory = categories.idCategory WHERE (((configSpec_products.specProduct)=" & pidProduct & ") AND ((configSpec_products.cdefault)<>0)) ORDER BY configSpec_products.catSort, categories.idCategory, configSpec_products.prdSort,products.description;"
			Set rs=Server.CreateObject("ADODB.RecordSet")
			Set rs=conntemp.execute(query)

			iAddDefaultPrice=Cdbl(0)
			iAddDefaultWPrice=Cdbl(0)
			
			if not rs.eof then
				If pnoprices<2 Then
					pcArr1=rs.getRows()					
					intCount1=ubound(pcArr1,2)
					For j=0 to intCount1
						dblprice=pcArr1(0,j)
						dblWprice=pcArr1(1,j)
						pcv_qty=pcArr1(2,j)
						if IsNull(pcv_qty) or pcv_qty="" then
							pcv_qty=0
						end if
						if clng(pcv_qty)=0 then
							pcv_qty=1
						end if
						if dblWprice=0 then
							dblWprice=dblprice
						end if
						iAddDefaultPrice=Cdbl(iAddDefaultPrice+dblprice*pcv_qty)
						iAddDefaultWPrice=Cdbl(iAddDefaultWPrice+dblWprice*pcv_qty)
					Next
					pPrice=Cdbl(pPrice+iAddDefaultPrice)
					pWPrice=Cdbl(pWPrice+iAddDefaultWPrice)
				else
					pPrice=0
					pWPrice=0
				end if
			end if
			set rs=nothing

			query="UPDATE Products SET pcProd_BTODefaultPrice=" & pPrice & ",pcProd_BTODefaultWPrice=" & pWPrice & " WHERE idProduct=" & pidProduct & ";"
			Set rs2=Server.CreateObject("ADODB.RecordSet")
			Set rs2=conntemp.execute(query)
			set rs2=nothing
			
			call updPrdEditedDate(pidProduct)
			
			For k=0 to intCount2
			session("admin_tmp_customerCategory")=pcArr2(0,k)
			session("admin_tmp_customertype")=pcArr2(1,k)
			session("admin_tmp_customerCategoryType")=pcArr2(2,k)
			if session("admin_tmp_customerCategoryType")="ATB" then
				session("admin_tmp_ATBCustomer")=1
				session("admin_tmp_ATBPercentage")=pcArr2(3,k)
				intpcCC_ATB_Off=pcArr2(4,k)
				if intpcCC_ATB_Off="Retail" then
					session("admin_tmp_ATBPercentOff")=0
				else
					session("admin_tmp_ATBPercentOff")=1
				end if
			else
				session("admin_tmp_ATBCustomer")=0
				session("admin_tmp_ATBPercentage")=0
				session("admin_tmp_ATBPercentOff")=0
			end if
			
			pPrice=save_pPrice
			pBtoBPrice=save_pBtoBPrice
			
			Call CalBTODefaultPriceCat()
		
			query="DELETE FROM pcBTODefaultPriceCats WHERE idproduct=" & pidProduct & " AND idcustomerCategory=" & session("admin_tmp_customerCategory") & ";"
			Set rs=Server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
			set rs=nothing
			query="INSERT INTO pcBTODefaultPriceCats (idproduct,idcustomerCategory,pcBDPC_Price) VALUES (" & pidProduct  & "," & session("admin_tmp_customerCategory") & "," & dblpcCC_Price & ");"
			Set rs=Server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
			set rs=nothing
			Next
		Next
	end if
	call closedb()

	%>	
<table class="pcCPcontent">
<tr>
	<td>
		<div class="pcCPmessageSuccess">
			All default prices were updated successfully!
		</div>
		<br>
		<br>
        <input type="button" name="updateDefault" value="Update Base Prices" onClick="document.location.href='updBTOPrdPrices.asp'" class="ibtnGrey">&nbsp;
        <input type="button" name="updateDefault" value="Update Configuration Prices" onClick="document.location.href='updateBTOprices.asp'" class="ibtnGrey">&nbsp;
	</td>
</tr>
</table>
<%ELSE%>
<form action="updBTODefaultPrices.asp?action=upd" method="post" class="pcForms">
<table class="pcCPcontent">
<tr>
	<td>
		Select <strong>Update BTO Product default prices</strong> to update all BTO Default prices. See <a href="http://wiki.earlyimpact.com/bto/btoaddnew#understanding_base_and_default_price" target="_blank">Understanding Base and Default Prices</a> for details on what a Default Price is.
	</td>
</tr>
<tr>
	<td><hr></td>
</tr>
<tr>
	<td align="center">
		<input type="submit" name="submit" value="Update BTO Product default prices" class="submit2">&nbsp;
        <input type="button" name="updateDefault" value="Update Base Prices" onClick="document.location.href='updBTOPrdPrices.asp'">&nbsp;
        <input type="button" name="updateDefault" value="Update Configuration Prices" onClick="document.location.href='updateBTOprices.asp'">&nbsp;
	</td>
</tr>
</table>
</form>
<%END IF%><!--#include file="AdminFooter.asp"-->
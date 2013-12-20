<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3
section="specials"
pcStrPageName="sm_addedit_S4.asp"
%>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/emailsettings.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/currencyformatinc.asp" -->
<% 
Dim query, conntemp, rstemp, rstemp1

call openDB()

IF request("a")="post" THEN
	UP=request("UP1")
	session("sm_UP1")=UP
	TempStr2=""
	tmpChangeName=""
	
	if UP="1" then
		session("sm_UP2")=request("priceSelect")
		session("sm_UP3")=request("cprice")
		session("sm_UP4")=request("cpriceType")
		session("sm_UP5")=request("cpriceRound")
		priceSelect=request("priceSelect")
		Select Case priceSelect
		Case "0": TempStr2="The Online Price"
		tmpChangeName="Online Price"
		Case "-1": TempStr2="The Wholesale Price"
		tmpChangeName="Wholesale Price"
		Case Else:
			tmpquery="Select pcCC_Name FROM pcCustomerCategories WHERE idcustomerCategory=" & priceSelect
			set rstemp4=connTemp.execute(tmpquery)
			TempStr2="The Price in Pricing Category: '" & rstemp4("pcCC_Name") & "'"
			tmpChangeName=rstemp4("pcCC_Name")
			set rstemp4=nothing
		End Select
		
		Select Case request("cpriceType")
		Case "0": TempStr2=TempStr2 & " will be reduced by: " & request("cprice") & "%"
		Case "1": TempStr2=TempStr2 & " will be reduced by: " & scCurSign & request("cprice")
		End Select
		if request("cpriceRound")="1" then
			TempStr2=TempStr2 & "<br>The updated price will be rounded to the nearest integer."
		else
			if request("cpriceRound")="0" then
				TempStr2=TempStr2 & "<br>The updated price will be rounded to the nearest hundredth."
			end if
		end if
	end if
	
	if UP="2" then
		session("sm_UP2")=request("priceSelect1")
		session("sm_UP3")=request("wprice")
		session("sm_UP4")=request("priceSelect2")
		session("sm_UP5")=request("cpriceRound1")
		priceSelect1=request("priceSelect1")
		priceSelect2=request("priceSelect2")
		Select Case priceSelect1
		Case "0": TempStr2="Make the online price " & request("wprice") & "% "
		tmpChangeName="Online Price"
		Case "-1": TempStr2="Make the wholesale price " & request("wprice") & "% "
		tmpChangeName="Wholesale Price"
		Case Else:
			tmpquery="Select pcCC_Name FROM pcCustomerCategories WHERE idcustomerCategory=" & priceSelect1
			set rstemp4=connTemp.execute(tmpquery)
			TempStr2="Change Price in Pricing Category: '" & rstemp4("pcCC_Name") & "' " & request("wprice") & "% "
			tmpChangeName=rstemp4("pcCC_Name")
			set rstemp4=nothing
		End Select
		Select Case priceSelect2
		Case "0": TempStr2=TempStr2 & "of the online price."
		Case "-2": TempStr2=TempStr2 & "of the list price."
		Case "-1": TempStr2=TempStr2 & "of the wholesale price."
		'Start SDBA
		Case "-3": TempStr2=TempStr2 & "of the product cost."
		'End SDBA
		Case Else:
			tmpquery="Select pcCC_Name FROM pcCustomerCategories WHERE idcustomerCategory=" & priceSelect2
			set rstemp4=connTemp.execute(tmpquery)
			TempStr2=TempStr2 & "of the price in pricing category: '" & rstemp4("pcCC_Name") & "'"
			set rstemp4=nothing
		End Select
		if request("cpriceRound1")="1" then
			TempStr2=TempStr2 & "<br>The updated price will be rounded to the nearest integer."
		else
			if request("cpriceRound1")="0" then
				TempStr2=TempStr2 & "<br>The updated price will be rounded to the nearest hundredth."
			end if
		end if
	end if
		
	session("sm_ChangeTxt")=TempStr2
	session("sm_ChangeName")=tmpChangeName
	session("sm_TechDetails")=""
	session("sm_TechDetails1")=""
	session("sm_TechDetails")=session("sm_TechDetails") & "Price Changes:<div style=""padding: 5px 0 12px 0; font-weight:bold;"">" & session("sm_ChangeTxt") & "</div>"
	session("sm_TechDetails1")="Products:<div style=""padding-top: 5px; font-weight:bold;"">The sale will affect " & session("sm_PrdCount") & " product(s) in your store.</div><div style=""padding-top: 5px;"" class=""pcSmallText"">If you just updated the product included in the sale, this number may be incorrect. it will be updated when you save the sale.</div>"
	
	

END IF

UP=session("sm_UP1")

if (UP="") OR (UP="0") then
	call closedb()
	response.redirect "sm_addedit_S1.asp"
	response.end
end if

pcSaleID=session("sm_pcSaleID")

pageIcon="pcv4_icon_salesManager.png"
if pcSaleID="" then
	pageTitle="Sales Manager - Create New Sale - Step 3: Preview Sale"
else
	pageTitle="Sales Manager - Edit Sale - Step 3: Preview Sale"
end if
	
%>
<!--#include file="AdminHeader.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<table class="pcCPcontent">
<tr> 
	<td colspan="4">
    	<div class="pcCPmessageInfo">
        	<strong>SUMMARY</strong>
            <div style="padding-top: 5px;">The sale will affect <b><%=session("sm_PrdCount")%></b> product(s) in your store. You can add/remove products by editing the sale.</div>
            <div><%=session("sm_ChangeTxt")%></div>
        </div>
	</td>
</tr>
<tr> 
	<td colspan="4">
		The following table shows how the Sale will be applied to the <strong><%if int(session("sm_PrdCount"))>5 then%>first 5 of the <%=session("sm_PrdCount")%><%else%><%=session("sm_PrdCount")%><%end if%> products</strong> that you have selected<% if pcSaleID<>"" then %> (<a href="sm_addedit_S1.asp?a=rev">edit</a>)<% end if %>.
	</td>
</tr>
<tr>
	<td colspan="4" class="pcCPspacer"></td>
</tr>
<tr>
	<th width="10%">SKU</th>
	<th width="45%">Product Name</th>
	<th width="15%"><div align="right"><%=session("sm_ChangeName")%></div></th>
	<th width="15%"><div align="right">Sale Price</div></th>
</tr>
<tr>
	<td colspan="4" class="pcCPspacer"></td>
</tr>
<%query="SELECT TOP 5 Products.IDProduct,products.price,products.listprice,products.btoBprice,products.cost,products.sku,products.description FROM " & session("sm_Param1") & " WHERE " & session("sm_Param2") & ";"
set rs=connTemp.execute(query)

if not rs.eof then
	pcArr=rs.getRows()
	set rs=nothing
	intCount=ubound(pcArr,2)
	For m=0 to intCount
	
	pidproduct=pcArr(0,m)
	pcv_Price=pcArr(1,m)
	pcv_ListPrice=pcArr(2,m)
	pcv_BtoBPrice=pcArr(3,m)
	pcv_Cost=pcArr(4,m)
	pcv_sku=pcArr(5,m)
	pcv_name=pcArr(6,m)
	
	if Cdbl(pcv_BtoBPrice)=0 then
		pcv_BtoBPrice=pcv_Price
	end if
	
	tmpOrgPrice=0
	tmpUpdPrice=0
	if UP="1" then
		priceSelect=session("sm_UP2")
		if priceSelect="0" then
			tempPrice=cdbl(pcv_Price)
			tmpOrgPrice=tempPrice
			if session("sm_UP4")="0" then
				tempPrice=tempPrice-cdbl(tempPrice*cdbl(replacecomma(session("sm_UP3")))*0.01)
			else
				if session("sm_UP4")="1" then
					tempPrice=tempPrice-cdbl(replacecomma(session("sm_UP3")))
				end if
			end if
			if session("sm_UP5")="1" then
				tempPrice=round(tempPrice)
			else
				if session("sm_UP5")="0" then
					tempPrice=round(tempPrice,2)
				end if
			end if
			tmpUpdPrice=tempPrice
		end if
		if priceSelect="-1" then
			tempPrice=cdbl(pcv_BtoBPrice)
			tmpOrgPrice=tempPrice
			if session("sm_UP4")="0" then
				tempPrice=tempPrice-cdbl(tempPrice*cdbl(replacecomma(session("sm_UP3")))*0.01)
			else
				if session("sm_UP4")="1" then
					tempPrice=tempPrice-cdbl(replacecomma(session("sm_UP3")))
				end if
			end if
			if session("sm_UP5")="1" then
				tempPrice=round(tempPrice)
			else
				if session("sm_UP5")="0" then
					tempPrice=round(tempPrice,2)
				end if
			end if
			tmpUpdPrice=tempPrice
		end if
		if Cint(priceSelect)>0 then
			tempPrice=0
			tmp_BTOTable=0
			query="SELECT pcCC_Price FROM pcCC_Pricing WHERE idproduct=" & pidproduct & " AND idcustomerCategory=" & priceSelect
			set rstemp1=connTemp.execute(query)
			if not rstemp1.eof then
				tempPrice=rstemp1("pcCC_Price")
				tempPrice=pcf_Round(tempPrice, 2)
				if IsNull(tempPrice) or tempPrice="" then
					tempPrice=0
				end if
			else
				query="SELECT pcCC_BTO_Price FROM pcCC_BTO_Pricing WHERE idBTOItem=" & pidproduct & " AND idcustomerCategory=" & priceSelect
				set rstemp1=connTemp.execute(query)
				if not rstemp1.eof then
					tempPrice=rstemp1("pcCC_BTO_Price")
					if IsNull(tempPrice) or tempPrice="" then
						tempPrice=0
					end if
					tmp_BTOTable=1
				end if
			end if
			set rstemp1=nothing
			
			if tempPrice<>"0" then
			else
					query="SELECT idcustomerCategory, pcCC_Name, pcCC_CategoryType, pcCC_ATB_Percentage, pcCC_ATB_Off FROM pcCustomerCategories WHERE idcustomerCategory=" & priceSelect
					SET rstemp1=Server.CreateObject("ADODB.RecordSet")
					SET rstemp1=conntemp.execute(query)
					if NOT rstemp1.eof then 
						intIdcustomerCategory=rstemp1("idcustomerCategory")
						strpcCC_Name=rstemp1("pcCC_Name")
						strpcCC_CategoryType=rstemp1("pcCC_CategoryType")
						intpcCC_ATBPercentage=rstemp1("pcCC_ATB_Percentage")
						intpcCC_ATB_Off=rstemp1("pcCC_ATB_Off")
						if intpcCC_ATB_Off="Retail" then
							intpcCC_ATBPercentOff=0
						else
							intpcCC_ATBPercentOff=1
						end if
						
						SP_price=pcv_Price
						SP_wprice=pcv_BtoBPrice
		
						if (SP_wprice>"0") then
							SPtempPrice=SP_wprice
						else
							SPtempPrice=SP_price
						end if
						' Calculate the "across the board" price
						if strpcCC_CategoryType="ATB" then
							if intpcCC_ATBPercentOff=0 then
								tempPrice=SP_price-(pcf_Round(SP_price*(cdbl(intpcCC_ATBPercentage)/100),2))
							else
								tempPrice=SPtempPrice-(pcf_Round(SPtempPrice*(cdbl(intpcCC_ATBPercentage)/100),2))
							end if
						end if
					end if
			end if
			
			tmpOrgPrice=tempPrice
			
			if session("sm_UP4")="0" then
				tempPrice=tempPrice-cdbl(tempPrice*cdbl(replacecomma(session("sm_UP3")))*0.01)
			else
				if session("sm_UP4")="1" then
					tempPrice=tempPrice-cdbl(replacecomma(session("sm_UP3")))
				end if
			end if
			if session("sm_UP5")="1" then
				tempPrice=round(tempPrice)
			else
				if session("sm_UP5")="0" then
					tempPrice=round(tempPrice,2)
				end if
			end if
				
			tmpUpdPrice=tempPrice

		end if				
	end if
	if UP="2" then
		priceSelect1=session("sm_UP2")
		priceSelect2=session("sm_UP4")
		tempPrice=0
		tmp_BTOTable=0
		
		if Cint(priceSelect1)>0 then
			query="SELECT pcCC_Price FROM pcCC_Pricing WHERE idproduct=" & pidproduct & " AND idcustomerCategory=" & priceSelect1
			set rstemp1=connTemp.execute(query)
			if not rstemp1.eof then
				tempPrice=rstemp1("pcCC_Price")
				tempPrice=pcf_Round(tempPrice, 2)
				if IsNull(tempPrice) or tempPrice="" then
					tempPrice=0
				end if
			else
				query="SELECT pcCC_BTO_Price FROM pcCC_BTO_Pricing WHERE idBTOItem=" & pidproduct & " AND idcustomerCategory=" & priceSelect1
				set rstemp1=connTemp.execute(query)
				if not rstemp1.eof then
					tempPrice=rstemp1("pcCC_BTO_Price")
					if IsNull(tempPrice) or tempPrice="" then
						tempPrice=0
					end if
					tmp_BTOTable=1
				end if
			end if
			set rstemp1=nothing
			
			if tempPrice<>"0" then
			else
					query="SELECT idcustomerCategory, pcCC_Name, pcCC_CategoryType, pcCC_ATB_Percentage, pcCC_ATB_Off FROM pcCustomerCategories WHERE idcustomerCategory=" & priceSelect1
					SET rstemp1=Server.CreateObject("ADODB.RecordSet")
					SET rstemp1=conntemp.execute(query)
					if NOT rstemp1.eof then 
						intIdcustomerCategory=rstemp1("idcustomerCategory")
						strpcCC_Name=rstemp1("pcCC_Name")
						strpcCC_CategoryType=rstemp1("pcCC_CategoryType")
						intpcCC_ATBPercentage=rstemp1("pcCC_ATB_Percentage")
						intpcCC_ATB_Off=rstemp1("pcCC_ATB_Off")
						if intpcCC_ATB_Off="Retail" then
							intpcCC_ATBPercentOff=0
						else
							intpcCC_ATBPercentOff=1
						end if
						
						SP_price=pcv_Price
						SP_wprice=pcv_BtoBPrice
		
						if (SP_wprice>"0") then
							SPtempPrice=SP_wprice
						else
							SPtempPrice=SP_price
						end if
						' Calculate the "across the board" price
						if strpcCC_CategoryType="ATB" then
							if intpcCC_ATBPercentOff=0 then
								tempPrice=SP_price-(pcf_Round(SP_price*(cdbl(intpcCC_ATBPercentage)/100),2))
							else
								tempPrice=SPtempPrice-(pcf_Round(SPtempPrice*(cdbl(intpcCC_ATBPercentage)/100),2))
							end if
						end if
					end if
			end if
			tmpOrgPrice=tempPrice
		end if
		
		if Cint(priceSelect2)>0 then
			query="SELECT pcCC_Price FROM pcCC_Pricing WHERE idproduct=" & pidproduct & " AND idcustomerCategory=" & priceSelect2
			set rstemp1=connTemp.execute(query)
			if not rstemp1.eof then
				tempPrice=rstemp1("pcCC_Price")
				tempPrice=pcf_Round(tempPrice, 2)
				if IsNull(tempPrice) or tempPrice="" then
					tempPrice=0
				end if
			else
				query="SELECT pcCC_BTO_Price FROM pcCC_BTO_Pricing WHERE idBTOItem=" & pidproduct & " AND idcustomerCategory=" & priceSelect2
				set rstemp1=connTemp.execute(query)
				if not rstemp1.eof then
					tempPrice=rstemp1("pcCC_BTO_Price")
					if IsNull(tempPrice) or tempPrice="" then
						tempPrice=0
					end if
					tmp_BTOTable=1
				end if
			end if
			set rstemp1=nothing
			
			if tempPrice<>"0" then
			else
					query="SELECT idcustomerCategory, pcCC_Name, pcCC_CategoryType, pcCC_ATB_Percentage, pcCC_ATB_Off FROM pcCustomerCategories WHERE idcustomerCategory=" & priceSelect2
					SET rstemp1=Server.CreateObject("ADODB.RecordSet")
					SET rstemp1=conntemp.execute(query)
					if NOT rstemp1.eof then 
						intIdcustomerCategory=rstemp1("idcustomerCategory")
						strpcCC_Name=rstemp1("pcCC_Name")
						strpcCC_CategoryType=rstemp1("pcCC_CategoryType")
						intpcCC_ATBPercentage=rstemp1("pcCC_ATB_Percentage")
						intpcCC_ATB_Off=rstemp1("pcCC_ATB_Off")
						if intpcCC_ATB_Off="Retail" then
							intpcCC_ATBPercentOff=0
						else
							intpcCC_ATBPercentOff=1
						end if
						
						SP_price=pcv_Price
						SP_wprice=pcv_BtoBPrice
		
						if (SP_wprice>"0") then
							SPtempPrice=SP_wprice
						else
							SPtempPrice=SP_price
						end if
						' Calculate the "across the board" price
						if strpcCC_CategoryType="ATB" then
							if intpcCC_ATBPercentOff=0 then
								tempPrice=SP_price-(pcf_Round(SP_price*(cdbl(intpcCC_ATBPercentage)/100),2))
							else
								tempPrice=SPtempPrice-(pcf_Round(SPtempPrice*(cdbl(intpcCC_ATBPercentage)/100),2))
							end if
						end if
					end if
			end if
		end if	
		if priceSelect1="0" then
		tmpOrgPrice=cdbl(pcv_Price)
		Select Case priceSelect2
		Case "0": tempPrice=cdbl(pcv_Price)
		Case "-2": tempPrice=cdbl(pcv_ListPrice)
		Case "-1": tempPrice=cdbl(pcv_BtoBPrice)
		Case "-3": tempPrice=cdbl(pcv_Cost)
		End Select
			if ((priceSelect2="-3") and (cdbl(tempPrice)<>0)) or (priceSelect2<>"-3") then
				tempPrice=tempPrice*cdbl(replacecomma(session("sm_UP3")))*0.01
				if session("sm_UP5")="1" then
					tempPrice=round(tempPrice)
				else
					if session("sm_UP5")="0" then
						tempPrice=round(tempPrice,2)
					end if
				end if
			end if
			tmpUpdPrice=tempPrice
		end if
		if priceSelect1="-1" then
		tmpOrgPrice=cdbl(pcv_BtoBPrice)
		Select Case priceSelect2
		Case "0": tempPrice=cdbl(pcv_Price)
		Case "-2": tempPrice=cdbl(pcv_ListPrice)
		Case "-1": tempPrice=cdbl(pcv_BtoBPrice)
		Case "-3": tempPrice=cdbl(pcv_Cost)
		End Select
			if ((priceSelect2="-3") and (cdbl(tempPrice)<>0)) or (priceSelect2<>"-3") then
				tempPrice=tempPrice*cdbl(replacecomma(session("sm_UP3")))*0.01
				if session("sm_UP5")="1" then
					tempPrice=round(tempPrice)
				else
					if session("sm_UP5")="0" then
						tempPrice=round(tempPrice,2)
					end if
				end if
			end if
			tmpUpdPrice=tempPrice
		end if
		
		if Cint(priceSelect1)>0 then
			Select Case priceSelect2
				Case "0": tempPrice=cdbl(pcv_Price)
				Case "-2": tempPrice=cdbl(pcv_ListPrice)
				Case "-1": tempPrice=cdbl(pcv_BtoBPrice)
				Case "-3": tempPrice=cdbl(pcv_Cost)
			End Select

			tempPrice=tempPrice*cdbl(replacecomma(session("sm_UP3")))*0.01
			if session("sm_UP5")="1" then
				tempPrice=round(tempPrice)
			else
				if session("sm_UP5")="0" then
					tempPrice=round(tempPrice,2)
				end if
			end if
			
			tmpUpdPrice=tempPrice
				
		end if
		
	end if%>
	<tr>
		<td><%=pcv_sku%></td>
		<td><%=pcv_name%></td>
		<td align="right"><%=scCurSign & money(tmpOrgPrice)%></td>
		<td align="right"><%=scCurSign & money(tmpUpdPrice)%></td>
	</tr>
	<%Next
end if
set rs=nothing%>
<tr>
	<td colspan="4" class="pcCPspacer"><hr></td>
</tr>
<tr align="center">
	<td colspan="4">
		<%if Clng(session("sm_PrdCount"))>0 then%>
		<input type="button" name="Go" value="Continue" onClick="location='sm_addedit_S5.asp';" class="submit2">&nbsp;
		<%end if%>
		<input type="button" name="Back" value="Change Sale Settings" onClick="location='sm_addedit_S3.asp?a=back';" class="ibtnGrey">
	</td>
</tr>
</table>
<% 
call closeDb()
%>
<!--#include file="AdminFooter.asp"-->
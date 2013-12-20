<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="sds_LIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/statusAPP.inc"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="header.asp"-->
<% Dim connTemp,rs,query
call opendb()

pcv_IdOrder=request("idorder")
if pcv_IdOrder="" then
	pcv_IdOrder=0
end if

if pcv_IdOrder=0 then
	response.redirect "menu.asp"
end if

	query="SELECT ord_OrderName FROM Orders WHERE idorder=" & pcv_IdOrder & ";"
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		pcv_OrderName=rs("ord_OrderName")
	end if
	set rs=nothing%>
<div id="pcMain">
	<table class="pcMainTable">
	<tr>
		<td valign="top">
		<table class="pcShowContent">
		<tr>
			<td colspan="6"><h1><%response.write dictLanguage.Item(Session("language")&"_sds_viewpast_1c")%> - <%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_1")%> <%=(scpre+int(pcv_IdOrder))%></h1></td>
		</tr>
		<tr>
			<td colspan="6" class="pcSpacer"></td>
		</tr>
		<tr>
			<td width="5%" align="center"><img border="0" src="images/step1a.gif"></td>
			<td width="28%" nowrap><b><%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_3")%></b></td>
			<td width="5%" align="center"><img border="0" src="images/step2.gif"></td>
			<td width="28%" nowrap><font color="#A8A8A8"><%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_4")%></font></td>
			<td width="5%" align="center"><img border="0" src="images/step3.gif"></td>
			<td width="29%" nowrap><font color="#A8A8A8"><%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_5")%></font></td>
		</tr>
		<tr>
			<td colspan="6" class="pcSpacer"></td>
		</tr>
		</table>
		</td>
	</tr>
	</table>
	<%
	If statusAPP="1" Then
		query = "SELECT idProduct FROM ProductsOrdered where idOrder = " & pcv_IdOrder & " AND ProductsOrdered.pcDropShipper_ID=" & session("pc_idsds") & ";"
		set rs=connTemp.execute(query)
		do until rs.eof
			dsdProductID = rs("idProduct")
			query = "SELECT * FROM pcDropShippersSuppliers WHERE idProduct = "&dsdProductID&" AND pcDS_IsDropShipper=" & session("pc_sdsIsDropShipper") & ";"
			set rsGObj = connTemp.execute(query)
			if rsGObj.eof then
				query = "INSERT INTO pcDropShippersSuppliers (idProduct,pcDS_IsDropShipper ) VALUES ("&dsdProductID&", " & session("pc_sdsIsDropShipper") & ");"
				set rsGOb2j = connTemp.execute(query)
				set rsGOb2j = nothing
			end if
			rs.moveNext
		loop
		set rsGObj = nothing
	End If
	query="SELECT Products.idproduct,Products.Description,Products.Stock,Products.sku,Products.pcProd_IsDropShipped,Products.pcDropShipper_ID,ProductsOrdered.quantity,ProductsOrdered.pcPrdOrd_BackOrder,ProductsOrdered.pcPrdOrd_Shipped FROM pcDropShippersSuppliers INNER JOIN (Products INNER JOIN ProductsOrdered ON Products.idproduct=ProductsOrdered.idproduct) ON (pcDropShippersSuppliers.idproduct=products.idproduct AND pcDropShippersSuppliers.pcDS_IsDropShipper=" & session("pc_sdsIsDropShipper") & ")  WHERE ProductsOrdered.idorder=" & pcv_IdOrder & " AND ProductsOrdered.pcDropShipper_ID=" & session("pc_idsds") & ";"
	set rs=connTemp.execute(query)
	
	IF rs.eof THEN
		set rs=nothing%>
		<div class="pcErrorMessage"><%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_6")%></div>
		<br>
		<br>
		<input type=button name="Back" value="javascript:history.back();" class="ibtnGrey">
<%
	ELSE
%>

		<Form name="form1" method="post" action="sds_ShipOrderWizard2.asp" class="pcForms">
		<table class="pcMainTable">
		<tr>
			<th>&nbsp;</th>
			<th width="35%"><%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_7")%></th>
			<th nowrap><%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_8")%></th>
			<th width="35%"><%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_9")%></th>
		</tr>
		<tr>
			<td colspan="4" class="pcSpacer"></td>
		</tr>
		<%
		pcv_count=0
		pcv_available=0
		Do while not rs.eof
			pcv_cancheck=0
			pcv_count=pcv_count+1
			pcv_IDProduct=rs("idproduct")
			pcv_Description=rs("description")
			pcv_Stock=rs("stock")
			pcv_Sku=rs("sku")			
			pcv_IsDropShipped=rs("pcProd_IsDropShipped")
			if IsNull(pcv_IsDropShipped) or pcv_IsDropShipped="" then
				pcv_IsDropShipped=0
			end if
			pcv_IDDropShipper=rs("pcDropShipper_ID")
			if IsNull(pcv_IDDropShipper) or pcv_IDDropShipper="" then
				pcv_IDDropShipper=0
			end if
			pcv_Qty=rs("quantity")
			if IsNull(pcv_Qty) or pcv_Qty="" then
				pcv_Qty=0
			end if
			pcv_BackOrder=rs("pcPrdOrd_BackOrder")
			if IsNull(pcv_BackOrder) or pcv_BackOrder="" then
				pcv_BackOrder=0
			end if
			pcv_Shipped=rs("pcPrdOrd_Shipped")
			if IsNull(pcv_Shipped) or pcv_Shipped="" then
				pcv_Shipped=0
			end if
			
			if (pcv_Shipped=0) or (pcv_BackOrder=1) then
				pcv_cancheck=1
				pcv_available=pcv_available+1
			end if
			%>
			<tr>
				<td><input type="checkbox" name="C<%=pcv_count%>" value="1" <%if pcv_cancheck=1 then%><%if (clng(pcv_Stock)>=clng(pcv_Qty)) and (pcv_BackOrder=0) then%>checked<%end if%><%else%>disabled<%end if%> class="clearBorder">
				<input type="hidden" name="IDPrd<%=pcv_count%>" value="<%=pcv_IDProduct%>"></td>
				<td><%if pcv_cancheck=0 then%><font color="#666666"><i><%end if%><%=pcv_Description%> (<%=pcv_sku%>)<%if pcv_cancheck=0 then%></i></font><%end if%></td>
				<td><%if pcv_cancheck=0 then%><font color="#666666"><i><%end if%><%=pcv_Qty%><%if pcv_cancheck=0 then%></i></font><%end if%></td>
				<td nowrap><%if pcv_cancheck=0 then%><font color="#666666"><i><%end if%>
				<%IF (pcv_Shipped=1) THEN%>
				<%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_10")%>
				<%ELSE%>
					<%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_11")%>
					<%if (pcv_BackOrder=1) then%>
					<%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_12")%>
					<%end if%>
				<%END IF%>
				<%if pcv_cancheck=0 then%></i></font><%end if%></td>
			</tr>
		<%	rs.MoveNext
		loop
		set rs=nothing%>
		<tr>
			<td colspan="4" class="pcSpacer"></td>
		</tr>
		<tr>
			<td colspan="4"><hr></td>
		</tr>
		<tr>
			<td colspan="4"><%if pcv_available>0 then%><input type="image" src="<%=rslayout("pcLO_processShip")%>" name="submit1" value="<%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_13")%>" border="0" id="submit"><%end if%>&nbsp;<a href="javascript:history.back();"><img src="<%=rslayout("back")%>" border="0"></a>
			<input type=hidden name="count" value="<%=pcv_count%>">
			<input type=hidden name="idorder" value="<%=pcv_IdOrder%>">
			</td>
		</tr>
		</table>
	</Form>
	<%END IF%>
<table class="pcMainTable">
	<tr>
    	<td>&nbsp;</td>
    </tr>
    <tr>
        <td><div align="center"><a href="sds_MainMenu.asp"><%response.write(dictLanguage.Item(Session("language")&"_CustPref_1"))%></a> - <a href="sds_ViewPast.asp"><%response.write(dictLanguage.Item(Session("language")&"_sdsMain_3"))%></a></div></td>
    </tr>
</table>
</div>
<%call closedb()%>
<!--#include file="footer.asp"-->
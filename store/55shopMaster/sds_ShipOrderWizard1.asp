<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%pageTitle="Shipping Wizard" %>
<% response.Buffer=true %>
<% section="orders" %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/FedExconstants.asp"-->
<!--#include file="../includes/pcFedExClass.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="AdminHeader.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<%
Dim connTemp,rs,query

call openDb()

'// Page Action
pcv_strPageAction = getUserInput(Request("PageAction"),7)
pcv_strPageMode = getUserInput(Request("m"),3)
if lcase(pcv_strPageMode)="new" then
	'Clear Sessions
	Session("pcAdminPackageCount")=""
end if

select case pcv_strPageAction
case "FedEx"
	pcv_strPageAction = "FedEx"
	pcv_strPageFormAction = "FedEx_ManageShipmentsRequest.asp"
case "FedExWs"
	pcv_strPageAction = "FedExWs"
	pcv_strPageFormAction = "FedExWS_ManageShipmentsRequest.asp"
case "UPS"
	pcv_strPageAction = "UPS"
	pcv_strPageFormAction = "UPS_ManageShipmentsRequest.asp"
case "USPS"
	pcv_strPageAction = "USPS"
	pcv_strPageFormAction = "USPS_ManageShipmentsRequest.asp"
case else
	pcv_strPageAction = "SDBA"
	pcv_strPageFormAction = "sds_ShipOrderWizard2.asp"
end select

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Shipping Center Config
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// Manage package count for Shipping Center
pcPackageCount = Request("PackageCount")
if len(pcPackageCount)>0 then
	pcArraySize = pcPackageCount
	pcPackageCount = pcPackageCount + 1
end if
'// Mangage Package Array for Shipping Center
pcv_strItemsList = Request("ItemsList")
if pcv_strItemsList<>"" then
	Dim pcLocalArray()
	ReDim pcLocalArray(pcArraySize)
	pcArray_TmpGlobalReturn = split(pcv_strItemsList, chr(124))
	For xPackageCount = LBound(pcArray_TmpGlobalReturn) TO UBound(pcArray_TmpGlobalReturn)
		pcLocalArray(xPackageCount) = pcArray_TmpGlobalReturn(xPackageCount)
	Next
end if
'// Manage unavailable items
Public Function pcf_ItemAvailable(productid)
	pcf_ItemAvailable = true
	if pcv_strItemsList<>"" then
		pcArray_TmpItemList = join(pcArray_TmpGlobalReturn,",")
		pcArray_TmpItemList = split(pcArray_TmpItemList,",")
		For xPackageCount = LBound(pcArray_TmpItemList) TO UBound(pcArray_TmpItemList)
			if cstr(productid) = pcArray_TmpItemList(xPackageCount) then
				pcf_ItemAvailable=false
			end if
		Next
	end if
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' End: Shipping Center Config
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'// Button
Dim pcv_strButtonText
if pcv_strPageAction = "FedExWs" OR pcv_strPageAction = "FedEx" OR pcv_strPageAction = "UPS" OR pcv_strPageAction = "USPS" then
	if pcPackageCount>1 then
		pcv_strButtonText = "Add this Package"
	else
		pcv_strButtonText = "Start Shipment"
		if pcv_strPageAction = "USPS" then
			pcv_strButtonText = "Continue"
		end if
	end if
else
	pcv_strButtonText = "Enter Shipment Details"
	if pcv_strPageAction = "USPS" then
		pcv_strButtonText = "Continue"
	end if
end if

' Validate order ID
pcv_IdOrder=getUserInput(request("idorder"),10)
if not validNum(pcv_IdOrder) then pcv_IdOrder=0
if pcv_IdOrder=0 then response.redirect "menu.asp"

query="SELECT ord_OrderName FROM Orders WHERE idorder=" & pcv_IdOrder & ";"
set rs=Server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query)

if not rs.eof then
	pcv_OrderName=rs("ord_OrderName")
end if
set rs=nothing
%>
<table class="pcCPcontent">
	<tr>
		<td colspan="2">Order ID#: <b><%=(scpre+int(pcv_IdOrder))%></b></td>
	</tr>
	<%
	'// Heading for FedEx
	if pcv_strPageAction="FedEx" OR pcv_strPageAction="UPS" or pcv_strPageAction="USPS" then
		select case pcv_strPageAction
			case "UPS"
				pcv_DescriptionString="UPS OnLine&reg; Tools Shipping"
			case "FedEx"
				pcv_DescriptionString="FedEx&reg; Shipment Request"
			case "FedExWs"
				pcv_DescriptionString="FedEx&reg; Shipment Request"
			case "USPS"
				pcv_DescriptionString="U.S.P.S. Shipment Request"
		end select %>
		<tr>
			<th colspan="2"><%=pcv_DescriptionString%></th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<% IF pcv_strPageAction="UPS" then %>
			<tr>
				<td colspan="2">
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td width="4%"><img src="../UPSLicense/LOGO_S2.jpg" width="45" height="50"></td>
							<td width="96%">UPS, THE UPS SHIELD TRADEMARK, THE UPS READY MARK, <br>
							THE UPS ONLINE TOOLS MARK AND THE COLOR BROWN ARE TRADEMARKS OF<br>
							UNITED PARCEL SERVICE OF AMERICA, INC. ALL RIGHTS RESERVED.</td>
						</tr>
					</table>
					<br />
					<br />Select the products to be included with this package, then click the "Process Shipment" button.<br>
<br>
				If the quantity of a product ordered is greater than one and you will be shipping it in a separate box, you may split the quantity of that product. For example, the customer purchases 10 widgets and you can only ship 5 in one package, you can use the <em><strong>&quot;Split Package Quantity&quot;</strong></em> link to divide the products for packaging purposes. When you split the quantity, the order will than contain two packages, each with the new quantity. In this case there would be two line items of 5 widgets each.</td>
			</tr>
		<% end if
		IF pcv_strPageAction="FedEx" or pcv_strPageAction="FedExWs" then  %>
			<tr>
				<td colspan="2"><img src="images/Clct_Prf_2c_Pos_Plt_150.png"><br /><br />Select the products to be included with this package, then click the "Process Shipment" button.
<br>
				If the quantity of a product ordered is greater than one and you will be shipping it in a separate box, you may split the quantity of that product. For example, the customer purchases 10 widgets and you can only ship 5 in one package, you can use the <em><strong>&quot;Split Package Quantity&quot;</strong></em> link to divide the products for packaging purposes. When you split the quantity, the order will than contain two packages, each with the new quantity. In this case there would be two line items of 5 widgets each.</td>
			</tr>
		<% end if
		IF pcv_strPageAction="USPS" then  %>
			<tr>
				<td colspan="2"><img src="images/hdr_uspsLogo.jpg"><br /><br />Select the products to be included with this package, then click the "Process Shipment" button.
<br>
				If the quantity of a product ordered is greater than one and you will be shipping it in a separate box, you may split the quantity of that product. For example, the customer purchases 10 widgets and you can only ship 5 in one package, you can use the <em><strong>&quot;Split Package Quantity&quot;</strong></em> link to divide the products for packaging purposes. When you split the quantity, the order will than contain two packages, each with the new quantity. In this case there would be two line items of 5 widgets each.</td>
			</tr>
			<%call opendb()
			query="SELECT pcES_UserID,pcES_PassP,pcES_AutoRefill,pcES_TriggerAmount,pcES_LogTrans,pcES_Reg,pcES_TestMode,pcES_AutoRmvLogs FROM pcEDCSettings WHERE pcES_Reg=1;"
			set rsQ=connTemp.execute(query)

			tmpEDCUserID=0
			if not rsQ.eof then
				EndiciaReg=1
				tmpEDCUserID=rsQ("pcES_UserID")
				EDCAutoRmv=rsQ("pcES_AutoRmvLogs")
				if tmpEDCUserID>"0" then
					pcv_strPageFormAction="EDCUSPS_ManageShipmentsRequest.asp"
					If EDCAutoRmv>"0" then
						dim dtTodaysDate
						dtTodaysDate=Date()-Clng(EDCAutoRmv)
						if SQL_Format="1" then
							dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
						else
							dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
						end if
						if scDB="SQL" then
							query="DELETE FROM pcEDCLogs WHERE pcET_ID IN (SELECT DISTINCT pcET_ID FROM pcEDCTrans WHERE pcET_TransDate<='" & dtTodaysDate & "');"
						else
							query="DELETE FROM pcEDCLogs WHERE pcET_ID IN (SELECT DISTINCT pcET_ID FROM pcEDCTrans WHERE pcET_TransDate<=#" & dtTodaysDate & "#);"
						end if
						set rsQ=connTemp.execute(query)
						set rsQ=nothing
						if scDB="SQL" then
							query="DELETE FROM pcEDCTrans WHERE pcET_TransDate<='" & dtTodaysDate & "';"
						else
							query="DELETE FROM pcEDCTrans WHERE pcET_TransDate<=#" & dtTodaysDate & "#;"
						end if
						set rsQ=connTemp.execute(query)
						set rsQ=nothing
					end if
				end if
			else
				EndiciaReg=0
			end if
			set rsQ=nothing%>
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
				<th colspan="2">Endicia's Postage Label Services for USPS</th>
			</tr>
			<tr valign="top">
				<td>
					<%if EndiciaReg=0 then%>
						You can choose Endicia's Postage Label Services to print USPS postage.<br>
						<a href="EDC_signup.asp">Click here</a> to sign-up for an Endicia account.
					<%else%>
						<%if tmpEDCUserID="0" OR tmpEDCUserID="" then%>
						You signed up to use Endicia's service to print USPS postage.<br>
						Please <a href="EDC_manage.asp">click here</a> to complete the sign up process and activate your account
						<%else%>
						You are using Endicia's Postage Label Services.<br>
						<a href="EDC_manage.asp">Click here</a> to manage your Endicia account.
						<%end if%>
					<%end if%>
				</td>
				<td align="right">
					<img src="images/PoweredByEndicia_small.jpg" border="0">
				</td>
			</tr>
		<% end if
	else
		'// Heading for Drop Shipping
		%>
		<tr>
			<td><b>Steps</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td width="5%" align="center"><img border="0" src="images/step1a.gif"></td>
			<td width="95%"><b>Select products</b></td>
		</tr>
		<tr>
			<td align="center"><img border="0" src="images/step2.gif"></td>
			<td><font color="#A8A8A8">Specify Shipment Details</font></td>
		</tr>
		<tr>
			<td align="center"><img border="0" src="images/step3.gif"></td>
			<td><font color="#A8A8A8">Finalize Shipment</font></td>
		</tr>
	<% end if %>
</table>

	<%
	query="SELECT Products.idproduct, Products.Description, Products.sku, Products.Stock, Products.noStock, Products.pcProd_IsDropShipped, Products.pcDropShipper_ID,  ProductsOrdered.idProductOrdered, ProductsOrdered.quantity, ProductsOrdered.pcPrdOrd_BackOrder, ProductsOrdered.pcPrdOrd_Shipped FROM Products INNER JOIN ProductsOrdered ON Products.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.idorder=" & pcv_IdOrder

	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	IF rs.eof THEN
		set rs=nothing%>
		<div class="pcCPmessage">
			No products found.<br /><br />
			<a href="#" onClick="javascript:history.back();">Back</a>
		</div>
	<% ELSE %>
		<form name="form1" method="post" action="<%=pcv_strPageFormAction%>" class="pcForms">
			<input name="PackageCount" type="hidden" value="<%=pcPackageCount%>">
			<%
			'// if there is a list of packages, then we must maintain the list with hidden fields.
			If pcv_strItemsList<>"" Then
				For xArrayCount = LBound(pcLocalArray) TO UBound(pcLocalArray)
				%>
					<input type="hidden" name="<%="pcAdminPrdList"&(xArrayCount+1)%>" value="<%=pcLocalArray(xArrayCount)%>">
				<%
				Next
			End If
			%>
			<table class="pcCPcontent">
				<tr>
					<td colspan="4" class="pcCPspacer"></td>
				</tr>
				<tr>
					<th width="80%" colspan="2">Product Name</th>
					<th>Quantity</th>
					<th>Status</th>
				</tr>
				<tr>
					<td colspan="4" class="pcCPspacer"></td>
				</tr>
				<%
				pcv_count=0
				pcv_available=0
				Do while not rs.eof
					pcv_cancheck=0
					pcv_count=pcv_count+1
					pcv_IDProduct=rs("idproduct")
					pcv_IDProductOrdered=rs("idProductOrdered")
					pcv_Description=rs("description")
					pcv_Sku=rs("sku")
					pcv_Stock=rs("stock")
					pcv_DisregardStock=rs("noStock")
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

					'// The backorder and stock checks below are only relevant if scOutOfStockPurchase = -1. Otherwise we only need check if its been shipped.
					if (pcv_Shipped=0) AND (((pcv_BackOrder="1") and (clng(pcv_Stock)>=clng(pcv_Qty))) or ((pcv_BackOrder="0") and (clng(pcv_Stock)>=0)) or ((pcv_DisregardStock<>0) and (pcv_IDDropShipper=0)) OR (scOutOfStockPurchase=0)) then
						pcv_cancheck=1
						pcv_available=pcv_available+1
					end if

					if (NOT pcf_ItemAvailable(pcv_IDProductOrdered)) AND (pcv_strPageAction = "FedExWs" OR pcv_strPageAction = "FedEx" OR pcv_strPageAction = "UPS" OR pcv_strPageAction = "USPS") then
						pcv_cancheck=0
					end if
					%>
					<tr>
						<td align="center" width="5%">
							<% dim pcv_showLink
							pcv_showLink = 0
							if (clng(pcv_Stock)>=clng(pcv_Qty)) then
								pcv_showLink = 1
							end if %>
							<input type="checkbox" name="C<%=pcv_count%>" value="1" <%if pcv_cancheck=1 then%><%if (clng(pcv_Stock)>=clng(pcv_Qty)) or (pcv_DisregardStock<>0 and pcv_IDDropShipper=0) then%>checked<%end if%><%else%>disabled<%end if%> class="clearBorder">
											<%if pcv_cancheck=1 then
											else
											response.write "<script language=""javascript"">"&vbcrlf
											response.write "function enableFieldC"&pcv_count&"()"&vbcrlf
											response.write "{"&vbcrlf
											response.write "document.form1.C"&pcv_count&".disabled=false;"&vbcrlf
											response.write "document.getElementById('showSubmit').style.display='';"&vbcrlf
											response.write "}"&vbcrlf
											response.write "</script>"&vbcrlf
											end if %>

							<input type="hidden" name="IDPrd<%=pcv_count%>" value="<%=pcv_IDProductOrdered%>">                        </td>
						<td align="left" width="70%"><%if pcv_cancheck=0 then%><font color="#666666"><i><%end if%><a href="FindProductType.asp?id=<%=pcv_IDProduct%>" target="_blank"><%=pcv_Description%></a> (<%=pcv_Sku%>)<%if pcv_cancheck=0 then%></i></font><%end if%></td>
						<td width="5%"><%if pcv_cancheck=0 then%><font color="#666666"><i><%end if%><%=pcv_Qty%><%if pcv_cancheck=0 then%></i></font><%end if%></td>
						<td nowrap width="20%"><%if pcv_cancheck=0 then%><font color="#666666"><i><%end if%>
						<%IF (pcv_Shipped=1) THEN%>
							Shipped
						<%ELSE%>
							<%if (pcv_DisregardStock<>0 OR scOutOfStockPurchase=0) AND (pcv_IDDropShipper=0) then%>
								Available
							<%else%>
								<%if (clng(pcv_Stock)>=0) and (pcv_IDDropShipper=0) then%>
									<% if pcf_ItemAvailable(pcv_IDProductOrdered) then %>
										Available
										<% if (pcv_showLink=0) then %>
											<br>
											<a href="javascript:enableFieldC<%=pcv_count%>()">Click here to override availability<a/>
										<% end if %>
									<% else %>
										Unavailable (Already in package)
									<% end if %>
								<%elseif (pcv_showLink=0) AND (pcv_BackOrder=1) AND (pcv_IDDropShipper=0) then%>
									Not available (Back-ordered)
									<br>
									<a href="javascript:enableFieldC<%=pcv_count%>()">Click here to override availability<a/>
								<%elseif (pcv_showLink=0) AND (pcv_BackOrder=0) AND (pcv_IDDropShipper=0) then%>
									Not available (Stock Level)
									<br>
									<a href="javascript:enableFieldC<%=pcv_count%>()">Click here to override availability<a/>
								<%end if%>
								<%if pcv_IsDropShipped=1 and pcv_IDDropShipper=0 then%>
									Already drop-shipped
								<%end if%>
								<%if pcv_IDDropShipper>0 then%>
									Will be drop-shipped
								<%end if%>
							<%end if%>
						<%END IF%>
						<%if pcv_cancheck=0 then%></i></font><%end if%>
						</td>
					</tr>
<script language="JavaScript">
<!--
	function newWindow2(file,window)
	{
		catWindow=open(file,window,'resizable=no,width=480,height=360,scrollbars=1');
		if (catWindow.opener == null) catWindow.opener = self;
		checkwin();
	}

	function checkwin()
	{
		if (catWindow.closed)
		{
			location="sds_ShipOrderWizard1.asp?idorder=<%=pcv_IdOrder%>&PageAction=<%=pcv_strPageAction%>"
		}
		else
		{
			setTimeout('checkwin()',500);
		}
	}

//-->
</script>
				<% IF pcv_Qty>1 then %>
						<TR>
							<TD colspan="2">&nbsp;</TD>
							<TD colspan="2"><a href="#" onClick="newWindow2('pcSplitPOQty_popup.asp?Qty=<%=pcv_Qty%>&idProductOrdered=<%=pcv_IDProductOrdered%>','window2')">Split Package Quantity</a></TD>
						</TR>
					<% end if
				rs.MoveNext
				loop
				set rs=nothing%>
				<script language="javascript">

				function enableFieldC1()
				{
				document.form1.C1.disabled=false;
				document.getElementById('showSubmit').style.display='';
				}

				</script>

				<tr>
					<td colspan="4" class="pcCPspacer"></td>
				</tr>
		<tr>
			<td colspan="4" align="left">
				<%if pcv_available>0 then%>
					<a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a>
					<br />
					<br />
					<script language="JavaScript">
					<!--
					function checkAll() {
					for (var j = 1; j <= <%=pcv_count%>; j++) {
					box = eval("document.form1.C" + j);
					if (box.checked == false) box.checked = true;
						 }
					}

					function uncheckAll() {
					for (var j = 1; j <= <%=pcv_count%>; j++) {
					box = eval("document.form1.C" + j);
					if (box.checked == true) box.checked = false;
						 }
					}

					//-->
					</script>
				<%end if%>
			<%if pcv_available>0 then%>
				<input type="submit" name="submit1" value="<%=pcv_strButtonText%>" class="submit2" <%if EndiciaReg=1 then%>onclick="javascript:pcf_Open_EndiciaPop();"<%end if%>>
			<%else%>
			<span id="showSubmit" style="display:none"><input type="submit" name="submit1" value="<%=pcv_strButtonText%>" class="submit2" <%if EndiciaReg=1 then%>onclick="javascript:pcf_Open_EndiciaPop();"<%end if%>></span>
			<%end if%>
			&nbsp;<input type="button" name="Back" value=" Back " onclick="javascript:history.back();">
			<input type="hidden" name="count" value="<%=pcv_count%>">
			<input type="hidden" name="idorder" value="<%=pcv_IdOrder%>">
			</td>
		</tr>
		<tr>
			<td colspan="4"><hr></td>
		</tr>
		<% if pcv_strPageAction="FedEx" OR pcv_strPageAction="FedExWs" then %>
			<tr>
			 <td colspan="4" align="center"><% = pcf_FedExWriteLegalDisclaimers %></td>
		  </tr>
		<% end if %>
	</table>
	</Form>
<%END IF%>
<%call closedb()%>
<%Response.write(pcf_ModalWindow("Connecting to Endicia Label Server... ","EndiciaPop", 300))%>
<!--#include file="AdminFooter.asp"-->
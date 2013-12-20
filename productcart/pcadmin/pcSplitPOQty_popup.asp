<% response.Buffer=true %>
<% PmAdmin=7%>
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
dim query, conntemp, rs

'--> open database connection
call openDB()
pcv_IDProductOrdered=request("idProductOrdered")
pcv_ProductQty=request("Qty")

if request("action")="upd" then
	pcv_Qty1=getUserInput(request("qty1"),3)
	pcv_Qty2=getUserInput(request("qty2"),3)
	
	if trim(pcv_Qty1)="" then
		pcv_Qty1=0
	end if
	
	if trim(pcv_Qty2)="" then
		pcv_Qty2=0
	end if
	
	if not validNum(pcv_Qty1) or not validNum(pcv_Qty2) then
		msg="You must enter valid quantities in each quantity field."
		response.redirect "pcSplitPOQty_popup.asp?idProductOrdered="&pcv_IDProductOrdered&"&Qty="&pcv_ProductQty&"&msg="&msg
		response.End()
	end if
	
	'Ensure that two Quanties equal the whole
	if (int(pcv_Qty1)+int(pcv_Qty2))<>int(pcv_ProductQty) then
		msg="The total of the new quantities must equal<br /> the original quantity of the order.<br />The original quantity is "&pcv_ProductQty&".<br /><br /><a href='pcSplitPOQty_popup.asp?idProductOrdered="&pcv_IDProductOrdered&"&Qty="&pcv_ProductQty&"'>Back</a>"
	else
		'Change quantity of the original ProductOrdered
		query="UPDATE ProductsOrdered SET quantity="&pcv_Qty1&" WHERE (((ProductsOrdered.idProductOrdered)="&pcv_IDProductOrdered&"));" 
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		'Get all the attributes from the original
		query="SELECT ProductsOrdered.idOrder, ProductsOrdered.idProduct, ProductsOrdered.service, ProductsOrdered.quantity, ProductsOrdered.idOptionA, ProductsOrdered.idOptionB, ProductsOrdered.unitPrice, ProductsOrdered.unitCost, ProductsOrdered.xfdetails, ProductsOrdered.idconfigSession, ProductsOrdered.rmaSubmitted, ProductsOrdered.QDiscounts, ProductsOrdered.ItemsDiscounts, ProductsOrdered.pcPackageInfo_ID, ProductsOrdered.pcDropShipper_ID, ProductsOrdered.pcPrdOrd_Shipped, ProductsOrdered.pcPrdOrd_BackOrder, ProductsOrdered.pcPrdOrd_SentNotice, ProductsOrdered.pcPrdOrd_SelectedOptions, ProductsOrdered.pcPrdOrd_OptionsPriceArray, ProductsOrdered.pcPrdOrd_OptionsArray, ProductsOrdered.pcPO_EPID, ProductsOrdered.pcPO_GWOpt, ProductsOrdered.pcPO_GWNote, ProductsOrdered.pcPO_GWPrice FROM ProductsOrdered WHERE (((ProductsOrdered.idProductOrdered)="&pcv_IDProductOrdered&"));"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		pcvt_idOrder=rs("idOrder")
		pcvt_idProduct=rs("idProduct")
		pcvt_service=rs("service")
		pcvt_service=rs("service")
		if pcvt_service=0 then
			pcvt_service="0"
		end if
		pcvt_quantity=rs("quantity")
		pcvt_idOptionA=rs("idOptionA")
		if isNULL(pcvt_idOptionA) or pcvt_idOptionA="" then
			pcvt_idOptionA=0
		end if
		pcvt_idOptionB=rs("idOptionB")
		if isNULL(pcvt_idOptionB) or pcvt_idOptionB="" then
			pcvt_idOptionB=0
		end if
		pcvt_unitPrice=rs("unitPrice")
		pcvt_unitCost=rs("unitCost")
		pcvt_xfdetails=rs("xfdetails")
		pcvt_idconfigSession=rs("idconfigSession")
		pcvt_rmaSubmitted=rs("rmaSubmitted")
		pcvt_QDiscounts=rs("QDiscounts")
		pcvt_ItemsDiscounts=rs("ItemsDiscounts")
		pcvt_pcPackageInfo_ID=rs("pcPackageInfo_ID")
		pcvt_pcDropShipper_ID=rs("pcDropShipper_ID")
		pcvt_pcPrdOrd_Shipped=rs("pcPrdOrd_Shipped")
		pcvt_pcPrdOrd_BackOrder=rs("pcPrdOrd_BackOrder")
		pcvt_pcPrdOrd_SentNotice=rs("pcPrdOrd_SentNotice")
		pcvt_pcPrdOrd_SelectedOptions=rs("pcPrdOrd_SelectedOptions")
		pcvt_pcPrdOrd_OptionsPriceArray=rs("pcPrdOrd_OptionsPriceArray")
		pcvt_pcPrdOrd_OptionsArray=rs("pcPrdOrd_OptionsArray")
		pcvt_pcPO_EPID=rs("pcPO_EPID")
		pcvt_pcPO_GWOpt=rs("pcPO_GWOpt")
		pcvt_pcPO_GWNote=rs("pcPO_GWNote")
		pcvt_pcPO_GWPrice=rs("pcPO_GWPrice")
		
		'Add new record with the same attributes as the original with the new Quantity
		query="INSERT INTO ProductsOrdered (idOrder, idProduct, service, quantity, idOptionA, idOptionB, unitPrice, unitCost, xfdetails, idconfigSession, rmaSubmitted, QDiscounts, ItemsDiscounts, pcPackageInfo_ID, pcDropShipper_ID, pcPrdOrd_Shipped, pcPrdOrd_BackOrder, pcPrdOrd_SentNotice, pcPrdOrd_SelectedOptions, pcPrdOrd_OptionsPriceArray, pcPrdOrd_OptionsArray, pcPO_EPID, pcPO_GWOpt, pcPO_GWNote, pcPO_GWPrice) VALUES ("&pcvt_idOrder&", "& pcvt_idProduct&", "& pcvt_service&", "& pcv_Qty2&", "& pcvt_idOptionA&", "& pcvt_idOptionB&", "& pcvt_unitPrice&", "& pcvt_unitCost&", '"& pcvt_xfdetails&"', "& pcvt_idconfigSession&", "& pcvt_rmaSubmitted&", 0, "& pcvt_ItemsDiscounts&", "& pcvt_pcPackageInfo_ID&", "& pcvt_pcDropShipper_ID&", "& pcvt_pcPrdOrd_Shipped&", "& pcvt_pcPrdOrd_BackOrder&", "& pcvt_pcPrdOrd_SentNotice&", '"& pcvt_pcPrdOrd_SelectedOptions&"', '"& pcvt_pcPrdOrd_OptionsPriceArray&"', '"& pcvt_pcPrdOrd_OptionsArray&"', "& pcvt_pcPO_EPID&", "& pcvt_pcPO_GWOpt&", '"& pcvt_pcPO_GWNote&"', "& pcvt_pcPO_GWPrice &");"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		set rs=nothing
		
		msg="You have successfully split the selected package!<br /><br />Please use the &quot;Close Window&quot; below to close this window and update the order details."
		btn="2"
		msgType=1
	end if
end if
%>

<html>
<head>
<title>Split Product Quantity</title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body>
<div id="pcCPmain" style="width:450px; background-image: none;">
<%if msg<>"" then%>
	<table class="pcCPcontent">
	<tr> 
		<td align="center" class="pcCPspacer"> 
			<% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
		</td>
	</tr>
	<tr>
		<td align="center">
        	<% if msgType=1 then %>
			<input type="button" name="Back" value="Close Window" onClick="opener.location.reload(); self.close();" class="ibtnGrey">
            <% end if %>
		</td>
	</tr>
</table>
<%else%>
<form name="form1" method="post" action="pcSplitPOQty_popup.asp?action=upd" class="pcForms">
<input type="hidden" name="idProductOrdered" value="<%=pcv_IDProductOrdered%>">
<input type="hidden" name="Qty" value="<%=pcv_ProductQty%>">

        <table class="pcCPcontent">
            <tr> 
                <td colspan="2" align="center" class="pcCPspacer">
                    <% 
                    msg=request("msg")
                    if trim(msg)<>"" then %>
                        <div class="pcCPmessage">
                            <%=msg%>
                        </div>
                    <% end if %>
                </td>
            </tr>
            <tr>
                <th colspan="2">Split Package Quantity</th>
            </tr>
            <tr>
                <td colspan="2" class="pcCPspacer"></td>
            </tr>	
				<tr>
				<td colspan="2">The split quantity must equal the total quantity of the original product ordered. If you wish to split the quantity into 3 or more packages, you may simply repeat the process again by clicking on the &quot;Split Package Quantity&quot; of the product again.</td>
                </tr>
				<tr>
					<td colspan="2">&nbsp;</td>
                </tr>
                <tr>
                    <td colspan="2">
                    The customer has ordered <strong><em><%=pcv_ProductQty%></em></strong> of this product.
                    </td>
                </tr>
				<tr>
					<td width="10%" align="right">
					  <input type="text" name="qty1" id="qty1" maxlength="4" size="4">
					</td>
					<td width="90%">Quantity of Package 1 </td>
				</tr>
				<tr>
				  <td width="10%" align="right">
						<input type="text" name="qty2" id="qty2" maxlength="4" size="4">
					</td>
					<td width="90%">Quantity of Package 2 </td>
				</tr>
				<tr>
					<td colspan="2"><hr></td>
				</tr>
			  <tr>
					<td colspan="2" align="center">
					<input type="submit" name="Submit" value="Save" class="submit2">
					<input type="button" name="Back" value="Close Window" onClick="opener.location.reload(); self.close();">
					</td>
				</tr>
				<tr>
					<td colspan="2">&nbsp;</td>
				</tr>
			</table>
        </form>
<%end if%>
</div>
</body>
</html>
<%call closedb()%>
<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<% 
Dim rs, conntemp, strSql
Dim il_intListID, il_categoryDesc, il_strMode, pcStrPageName

pcStrPageName="editCategories.asp"

if request("iPageCurrent")="" then
	iPageCurrent=1 
else
	iPageCurrent=Request("iPageCurrent")
end If
if request("nav")="1" then
	section="services"
else
	section="products"
end if

il_strMode=""

Dim reqstr, reqidproduct, reqidcategory
reqstr=Request.QueryString("reqstr")
reqidproduct=Request.QueryString("reqidproduct")
reqidcategory=Request.QueryString("reqidcategory")
If reqstr="" then
	reqstr=Request.Form("reqstr")
	reqidproduct=Request.Form("reqidproduct")
	reqidcategory=Request.Form("reqidcategory")
End If
	
If isNumeric(Request.QueryString("lid")) AND Request.QueryString("lid") <> "" Then
	il_intListID=Request.QueryString("lid")
ElseIf isNumeric(Request.QueryString("lid")) AND Request.Form("lid") <> "" Then
	il_intListID=Request.Form("lid")
Else
	Response.Redirect "../manageCategories.asp"
End If

If Request.QueryString("mode") <> "" Then
	il_strMode=Request.QueryString("mode")
ElseIf Request.Form("mode") <> "" Then
	il_strMode=Request.Form("mode")
Else
	il_strMode="view"
End If

Set conntemp=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.Recordset")
call Opendb()

strSQL="SELECT categoryDesc, idCategory FROM categories WHERE idCategory=" & il_intListID & " ORDER BY categoryDesc"
rs.Open strSQL, conntemp, adOpenStatic, adLockReadOnly

If Err Then
	TrapError Err.Description
Else
	il_categoryDesc=rs("categoryDesc")
End If
	
rs.Close

SOrder=request("SOrder")
if not SOrder<>"" then
	SOrder="sku"
end if
Sort=request("Sort")
if not Sort<>"" then
	Sort="asc"
end if

pageTitle="Products currently assigned to: " & il_categoryDesc
%>
<!--#include file="AdminHeader.asp"-->
<%src_checkPrdType=request("cptype")
if src_checkPrdType="" then
	src_checkPrdType="0"
end if%>

	<div style="margin: 5px 0 15px 15px;"><a href="JavaScript:;" onClick="document.getElementById('FindProducts').style.display=''; document.getElementById('pcHideInactive').style.display='none';">Add new products</a> to &quot;<%=il_categoryDesc%>&quot;&nbsp;|&nbsp;<a href="updPrdPrices.asp?idcategory=<%=il_intListID%>">Update Product Prices</a>&nbsp;|&nbsp;<a href="../pc/viewcategories.asp?idcategory=<%=il_intListID%>" target="_blank">View in the storefront</a>
	</div>
    
    <table id="FindProducts" class="pcCPcontent" style="display:none;">
        <tr>
            <td>
            <div id="FindProductsClose" style="float: right; padding: 18px 80px 0 0;"><a href="JavaScript:;" onClick="document.getElementById('FindProducts').style.display='none'; document.getElementById('pcHideInactive').style.display='';"><img src="images/pcIconDelete.jpg" width="12" height="12" alt="Close this panel"></a></div>
            <%
                src_ShowPrdTypeBtns=1
                src_FormTitle1="Find Products"
                src_FormTitle2="Add Products to " & il_categoryDesc
                src_FormTips1="Use the following filters to look for products in your store."
                src_FormTips2="Select the products that you would like to add to the '" & il_categoryDesc & "' category."
                src_IncNormal=0
                src_IncBTO=0
                src_IncItem=0
                src_DisplayType=1
                src_ShowLinks=0
                src_FromPage="editCategories.asp?iPageCurrent=1&Sorder=" & request("SOrder") & "&Sort=" & request("Sort") & "&nav=" & request("nav") & "&lid=" & il_intListID & "&mode=view"
                src_ToPage="actionCategories.asp?Sorder=" & request("SOrder") & "&Sort=" & request("Sort") & "&nav=" & request("nav") & "&lid=" & il_intListID & "&mode=edit&reqstr=" & reqstr & "&reqidproduct=" & reqidproduct & "&reqidcategory=" & reqidcategory
                src_Button1=" Search "
                src_Button2=" Add products to " & il_categoryDesc
                src_Button3=" Back "
                src_PageSize=15
                UseSpecial=1
                session("srcprd_from")=""
                session("srcprd_where")=" AND (products.idProduct NOT IN (SELECT idProduct FROM categories_products WHERE idCategory=" & il_intListID & ")) "
            %>
                <!--#include file="inc_srcPrds.asp"-->
            </td>
        </tr>
    </table>
	
	<% ' START show message, if any %>
        <!--#include file="pcv4_showMessage.asp"-->
    <% 	' END show message %>
	 
	<form method="POST" action="actionCategories.asp" name="myForm" class="pcForms">
		<input type="hidden" name="mode" value="<%=il_strMode%>"><input type="hidden" name="lid" value="<%= il_intListID %>">
		<input type="hidden" name="reqstr" value="<%=reqstr%>">
		<input type="hidden" name="reqidproduct" value="<%=reqidproduct%>">
		<input type="hidden" name="reqidcategory" value="<%=reqidcategory%>">
		<input type="hidden" name="nav" value="<%=request("nav")%>"> 
		<input type="hidden" name="SOrder" value="<%=SOrder%>">
		<input type="hidden" name="Sort" value="<%=Sort%>">
		<input type="hidden" name="frmHideInactiveProducts" value="<%=Request("frmHideInactiveProducts")%>" >
        
    	<div id="pcHideInactive" style="position: absolute; left: 635px; top: 144px;">
        <input class="clearBorder" onclick="javascript:reloadFormToShowHideInactiveProducts();" type="checkbox" <%if Request("frmHideInactiveProducts") = "true" then%>checked<%end if %> name="chkHideInActiveProducts" />&nbsp;Hide inactive products
        </div>
		
		<table class="pcCPcontent">
			<tr>
				<th width="5%">&nbsp;</th>
				<th width="20%"><a href="editCategories.asp?Sorder=sku&Sort=ASC&nav=<%=request("nav")%>&lid=<%= il_intListID %>&mode=<%=request("mode")%>&reqstr=<%=reqstr%>&reqidproduct=<%=reqidproduct%>&reqidcategory=<%=reqidcategory%>&frmHideInactiveProducts=<%=Request("frmHideInactiveProducts")%>"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="editCategories.asp?Sorder=sku&Sort=Desc&nav=<%=request("nav")%>&lid=<%= il_intListID %>&mode=<%=request("mode")%>&reqstr=<%=reqstr%>&reqidproduct=<%=reqidproduct%>&reqidcategory=<%=reqidcategory%>&frmHideInactiveProducts=<%=Request("frmHideInactiveProducts")%>"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a> SKU
				</th>
				<th width="65%"><a href="editCategories.asp?Sorder=Description&Sort=ASC&nav=<%=request("nav")%>&lid=<%= il_intListID %>&mode=<%=request("mode")%>&reqstr=<%=reqstr%>&reqidproduct=<%=reqidproduct%>&reqidcategory=<%=reqidcategory%>&frmHideInactiveProducts=<%=Request("frmHideInactiveProducts")%>"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="editCategories.asp?Sorder=Description&Sort=Desc&nav=<%=request("nav")%>&lid=<%= il_intListID %>&mode=<%=request("mode")%>&reqstr=<%=reqstr%>&reqidproduct=<%=reqidproduct%>&reqidcategory=<%=reqidcategory%>&frmHideInactiveProducts=<%=Request("frmHideInactiveProducts")%>"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a> Product Name
				</th>
				<th width="10%" nowrap><a href="editCategories.asp?Sorder=AL.POrder&Sort=ASC&nav=<%=request("nav")%>&lid=<%= il_intListID %>&mode=<%=request("mode")%>&reqstr=<%=reqstr%>&reqidproduct=<%=reqidproduct%>&reqidcategory=<%=reqidcategory%>&frmHideInactiveProducts=<%=Request("frmHideInactiveProducts")%>"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="editCategories.asp?Sorder=AL.POrder&Sort=Desc&nav=<%=request("nav")%>&lid=<%= il_intListID %>&mode=<%=request("mode")%>&reqstr=<%=reqstr%>&reqidproduct=<%=reqidproduct%>&reqidcategory=<%=reqidcategory%>&frmHideInactiveProducts=<%=Request("frmHideInactiveProducts")%>"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a> Order
				</th>
			</tr>
			<%
			call opendb()
			strSQL="SELECT A.idProduct AS idProduct,description,active,A.configOnly,A.serviceSpec,A.sku,AL.POrder FROM products A, categories_products AL, categories L WHERE A.removed=0"
			
			if ( Request("frmHideInactiveProducts") = "true" ) then
			    strSQL=strSQL&" AND A.Active = -1 AND A.idProduct=AL.idProduct AND AL.idCategory=L.idCategory AND L.idCategory=" & il_intListID & " ORDER BY " & Sorder & " " & Sort
			else
			    strSQL=strSQL&" AND A.idProduct=AL.idProduct AND AL.idCategory=L.idCategory AND L.idCategory=" & il_intListID & " ORDER BY " & Sorder & " " & Sort
			end if
			    
			set rs=Server.CreateObject("ADODB.Recordset")
			rs.CacheSize=50
			rs.PageSize=50
			rs.Open strSQL, conntemp, adOpenStatic, adLockReadOnly
			If Err Then
				rs.Close
				conntemp.Close
			ElseIf rs.EOF OR rs.BOF Then
			%>
			<tr> 
				<td colspan="4">No Products Found</td>
			</tr>
			<% Else
				rs.MoveFirst
	
				' get the max number of pages
				Dim iPageCount
				iPageCount=rs.PageCount
				If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
				If iPageCurrent < 1 Then iPageCurrent=1
						
				' set the absolute page
				rs.AbsolutePage=iPageCurrent
				'end for nav

				Dim Count
				Count=0
				Do While NOT rs.EOF And Count < rs.PageSize
					%>
					<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
						<td width="5%" align="center" bgcolor="<%= strCol %>"><input type="checkbox" name="Address" value="<%= rs("idProduct") %>" class="clearBorder"></td>
						<td width="10%" nowrap><%=rs("sku")%></td>
						<td width="75%">
							<%if (rs("configOnly")=0) and (rs("serviceSpec")<>0) then%>
								<a href="FindProductType.asp?id=<%= rs("idProduct") %>"><%= rs("description") %></a>&nbsp;
								<font color="#FF0000">(BTO product)</font>
							<%elseif (rs("configOnly")<>0) and (rs("serviceSpec")=0) then%>
								<a href="FindProductType.asp?id=<%= rs("idProduct") %>"><%= rs("description") %></a>&nbsp;
								<font color="#FF0000">(BTO item)</font>
							<%else%>
								<a href="FindProductType.asp?id=<%= rs("idProduct") %>"><%= rs("description") %></a>
							<%end if%>
							<% if rs("active")<>-1 then %>
								&nbsp;<font color="#FF0000">(Inactive)</font>
							<% end if %>
						</td>
						<td width="10%" align="center">
							<%if il_strMode="view" then
								POrder=rs("POrder")
							if POrder<>"" then
							else
								POrder="0"
							end if
								%>
							<input type="text" name="POrder" value="<%=POrder%>" size="3">
							<input type="hidden" name="listidproduct" value="<%= rs("idProduct") %>">
							<%else%>
								&nbsp;
							<%end if%>
						</td>
						</tr>
						<% count=count + 1
						rs.MoveNext
					Loop
					rs.Close
					%>
					<tr> 
						<td align="center"> 
							<input type="checkbox" value="ON" onClick="javascript:checkTheBoxes();" class="clearBorder">
						</td>
						<td colspan="3">Select All</td>
					</tr>
					<tr>
						<td colspan="4"><hr></td>
					</tr>
					<tr>
						<td colspan="4">
							<input type="submit" value="Remove Checked" class="submit2" onclick="javascript:if (!(checkedBoxes())) {alert('Please select at least one product to remove from this category.'); return(false)}">
							&nbsp;<input type="submit" name="UpdateOrder" value="Update Product Order" class="submit2">
                            &nbsp;<input type="submit" name="ResetOrder" value="Reset Product Order" class="submit2" onClick="JavaScript: if (confirm('PLEASE NOTE: you are about to reset to &quot;0&quot; the Order value for all products in this category. If you do so, products will be ordered based on the general product sorting criteria set under &quot;Store Settings > Display Settings&quot;. Would you like to continue?'));">
                            &nbsp;<input type="submit" name="CopyTo" value="Copy Selected to..." onclick="javascript:if (!(checkedBoxes())) {alert('Please select at least one product to copy to another category.'); return(false)}">
                            &nbsp;<input type="submit" name="MoveTo" value="Move Selected to...." onclick="javascript:if (!(checkedBoxes())) {alert('Please select at least one product to move to another category.'); return(false)}">
						</td>
					</tr>
		</table>
					
		<!-- page navigation-->
		<% If iPageCount>1 Then %>
		<p class="pcPageNav">
			<%response.Write("Currently viewing page "& iPageCurrent & " of "& iPageCount & " <P>")%>
		</p>
		<table class="pcPageNav">
			<tr>
				<td width="100%">
					<%' display Next / Prev buttons
					if iPageCurrent > 1 then %>
						<a href="editCategories.asp?iPageCurrent=<%=iPageCurrent-1%>&Sorder=<%=request("SOrder")%>&Sort=<%=request("Sort")%>&nav=<%=request("nav")%>&lid=<%= il_intListID %>&mode=<%=request("mode")%>&frmHideInactiveProducts=<%=Request("frmHideInactiveProducts")%>"><img src="../pc/images/prev.gif" width="10" height="10" border="0"></a> 
					<% end If
					For I=1 To iPageCount
						If Cint(I)=Cint(iPageCurrent) Then %>
							<%=I%>
						<% Else %>
							<a href="editCategories.asp?iPageCurrent=<%=I%>&Sorder=<%=request("SOrder")%>&Sort=<%=request("Sort")%>&nav=<%=request("nav")%>&lid=<%= il_intListID %>&mode=<%=request("mode")%>&frmHideInactiveProducts=<%=Request("frmHideInactiveProducts")%>"><%=I%></a> 
						<% End If %>
					<% Next %>
					<% if CInt(iPageCurrent) < CInt(iPageCount) then %>
						<a href="editCategories.asp?iPageCurrent=<%=iPageCurrent+1%>&Sorder=<%=request("SOrder")%>&Sort=<%=request("Sort")%>&nav=<%=request("nav")%>&lid=<%= il_intListID %>&mode=<%=request("mode")%>&frmHideInactiveProducts=<%=Request("frmHideInactiveProducts")%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
					<% end If %>
				</td>
			</tr>
		</table>
		<% End If %>
		<!-- end page navigation -->
	<table class="pcCPcontent">
	<% End If ' end of rs.eof%>
	<% if reqstr<>"" then %>
	<tr>
		<td colspan="4">
			<% if il_intListID <> 1 then %>
			<input type="button" value="Edit Category" onClick="document.location.href='modCata.asp?idcategory=<%=il_intListID%>'" name="button">&nbsp;
			<% end if %>
			<input type="button" value="Manage Categories" onClick="document.location.href='<%=reqstr%>&nav=<%=request("nav")%>&idproduct=<%=reqidproduct%>&idcategory=<%=reqidcategory%>';" name="button">
		</td>
	</tr>
	<% else %>
	<tr>
		<td colspan="4">
			<% if il_intListID <> 1 then %>
			<input type="button" value="Edit Category" onClick="document.location.href='modCata.asp?idcategory=<%=il_intListID%>'">&nbsp;
			<% end if %>
			<input type="button" value="Manage Categories" onClick="document.location.href='manageCategories.asp?nav=<%=request("nav")%>';">
		</td>
	</tr>
	<% end if %>
	</table>
</form>
<script>
function checkedBoxes()
{
	with (document.myForm) {
		for(i=0; i < elements.length -1; i++) {
			if ( elements[i].name== "Address" ) {
				if (elements[i].checked==true) return(true);
			}
		}
	}
	return(false);
}
</script>
<%call closedb()%><!--#include file="AdminFooter.asp"-->
<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Section="products" %>
<%PmAdmin="2*3*"%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% 
dim mySQL, conntemp, rstemp
call openDb()
IF request("action")="update" then
	IF request("submit1")<>"" then
		Count=request("count")
		if (Count>"0") and (IsNumeric(Count)) then
			For k=1 to Count
				if request("C" & k)="1" then
					pcv_ID=request("ID" & k)
					if pcv_ID<>"" then
						query="DELETE FROM cs_relationships WHERE idproduct=" & pcv_ID
						set rs=connTemp.execute(query)
					end if
				end if
			Next
		set rs=nothing
		call closedb()
		response.Redirect "crossSellView.asp?s=1&msg=" & server.URLEncode("The selected cross-selling relationships were successfully removed.")
		end if
	End if
END IF


	iPageCurrent=request("iPageCurrent")
	if iPageCurrent="" or iPageCurrent="0" then
		iPageCurrent=1
	end if
	iPageCount=0
	
	iPageSize=request("iPageSize")
	if iPageSize="" or iPageSize="0" then
		iPageSize=10
	end if
	
	sOrder=request("sOrder")
	if sOrder="" then
		sOrder="ASC"
	end if
	
	'// Filter by category
	Dim pcIntCategoryID, queryCat
	pcIntCategoryID=request("idcategory")
		if not validNum(pcIntCategoryID) then
			pcIntCategoryID=request("idcat")
		end if
	if validNum(pcIntCategoryID) then
		queryCat="WHERE products.idproduct IN (SELECT DISTINCT categories_products.idproduct FROM categories_products WHERE categories_products.idcategory=" & pcIntCategoryID & ") "
		' Get Category Name:
		query="SELECT categoryDesc FROM categories WHERE idCategory="&pcIntCategoryID
		set rstemp=Server.CreateObject("ADODB.Recordset")
		set rstemp=conntemp.execute(query)
		pcStrCategoryName=rstemp("categoryDesc")
		set rstemp=nothing
	end if
	
	query="SELECT DISTINCT cs_relationships.idproduct, products.description, products.sku FROM cs_relationships INNER JOIN products ON cs_relationships.idproduct=products.idProduct " & queryCat & "ORDER BY products.description " & sOrder & ";"
	set rstemp=Server.CreateObject("ADODB.Recordset") 
	rstemp.CacheSize=iPageSize
	rstemp.PageSize=iPageSize
	rstemp.Open query, connTemp, adOpenStatic, adLockReadOnly
	
	if not rstemp.eof then
		rstemp.AbsolutePage=iPageCurrent
		iPageCount=rstemp.PageCount
	end if

	pcv_intCount=-1
	if not rstemp.eof then
	pcArray=rstemp.getRows(iPageSize)
	pcv_intCount=ubound(pcArray,2)
	end if
	set rstemp=nothing
	
if validNum(pcIntCategoryID) then
	pageTitle="Cross Selling Relationships in <strong>" & pcStrCategoryName & "</strong>"
	else
	pageTitle="Cross Selling Relationships"
end if	
	
%>
<!--#include file="AdminHeader.asp"-->
<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<form method="POST" action="crossSellView.asp?action=update" name="checkboxform" class="pcForms">
	<table class="pcCPcontent">
		<tr>
			<td colspan="4" align="right">
			<%if pcv_intCount>-1 then%>
			Results per page:
			<select name="iPageSize" onchange="javascript:document.checkboxform.submit();">
				<option value="5" <%if iPageSize="5" then%>selected<%end if%>>5</option>
				<option value="10" <%if iPageSize="10" then%>selected<%end if%>>10</option>
				<option value="15" <%if iPageSize="15" then%>selected<%end if%>>15</option>
				<option value="20" <%if iPageSize="20" then%>selected<%end if%>>20</option>
				<option value="25" <%if iPageSize="25" then%>selected<%end if%>>25</option>
			</select>
            
            &nbsp;&nbsp;
            Only show products from:
            <%
            
            cat_DropDownName="idcat"
            cat_Type="1"
            cat_DropDownSize="1"
            cat_MultiSelect="0"
            cat_ExcBTOHide="0"
            cat_StoreFront="0"
            cat_ShowParent="1"
            cat_DefaultItem=""
            cat_SelectedItems="0,"
            cat_ExcItems=""
            cat_ExcSubs="0"
            cat_EventAction="onchange=""location='crossSellView.asp?idcat=' + document.checkboxform.idcat.value + ''"""
            %>
            <!--#include file="../includes/pcCategoriesList.asp"-->
            <%call pcs_CatList()%>
            
			<%end if%>
			</td>
		</tr>
		<tr>
			<th colspan="2" width="40%" nowrap>
				<a href="crossSellView.asp?idcat=<%=pcIntCategoryID%>&iPageCurrent=<%=iPageCurrent%>&iPageSize=<%=iPageSize%>&sOrder=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a>
				<a href="crossSellView.asp?idcat=<%=pcIntCategoryID%>&iPageCurrent=<%=iPageCurrent%>&iPageSize=<%=iPageSize%>&sOrder=DESC"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>
			Primary Product</th>
			<th width="60%" colspan="2" nowrap>Related Products</th>
		</tr>
		<tr>
			<td colspan="4" class="pcCPspacer"></td>
		</tr>
	<%
	Count=0
	For i=0 to pcv_intCount
	Count=Count+1
	idmain=pcArray(0,i)
	mainName=pcArray(1,i)
	mainNameSku=pcArray(2,i)
	query="SELECT DISTINCT cs_relationships.idproduct, cs_relationships.idrelation, products.description FROM cs_relationships INNER JOIN products ON (cs_relationships.idrelation=products.idProduct) WHERE (((cs_relationships.idproduct)="&idmain&"));"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=conntemp.execute(query)
	%>
		<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
			<td valign="top" width="1%">
				<input type="checkbox" size="3" name="C<%=count%>" value="1" class="clearBorder">
				<input type="hidden" name="ID<%=count%>" value="<%=idmain%>"> 
            </td>
            <td valign="top" width="39%">
				&nbsp;<a href="FindProductType.asp?id=<%=idmain%>" target="_blank"><%=mainName%></a><br /><span class="pcSmallText">(<%=mainNameSku%>)</span>
			</td>
			<td class="pcSmallText" valign="top">
			<%
				cnt=0
				pcv_intCount1=-1
				if not rs.eof then
					pcArray1=rs.getRows()
					pcv_intCount1=ubound(pcArray1,2)
				end if
				set rs=nothing
				For j=0 to pcv_intCount1
					if cnt<>0 then
						response.Write ", "
					end if 
					response.Write pcArray1(2,j)
					cnt=cnt+1
				Next
				%>
			</td>
			<td nowrap="nowrap" class="cpLinksList" valign="top">
				<a href="crossSellEdit.asp?idmain=<%=idmain%>">View/Edit</a>&nbsp;
				<a href="crossSellMultiAdd.asp?idmain=<%=idmain%>">Copy</a>
			</td>
		</tr>
	<%
	Next
	%>
	<%if count>"0" then%>
	<script language="JavaScript">
	<!--
	function checkAll() {
	for (var j = 1; j <= <%=count%>; j++) {
	box = eval("document.checkboxform.C" + j); 
	if (box.checked == false) box.checked = true;
	}
	}
	
	function uncheckAll() {
	for (var j = 1; j <= <%=count%>; j++) {
	box = eval("document.checkboxform.C" + j); 
	if (box.checked == true) box.checked = false;
	}
	}
	//-->
	</script>
    <tr>
    	<td colspan="4" class="pcCPspacer"></td>
    </tr>
	<tr>
		<td colspan="2" class="cpLinksList">
		<input type="hidden" name="count" value=<%=count%>>
		<a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a>
		</td>
        <td colspan="2" align="right" class="cpLinksList">
        <%
		If iPageCount>1 Then
		' display Next / Prev buttons
		if iPageCurrent > 1 then %>
		<a href="crossSellView.asp?idcat=<%=pcIntCategoryID%>&iPageCurrent=<%=iPageCurrent-1%>&iPageSize=<%=iPageSize%>&sOrder=<%=sOrder%>"><img src="../pc/images/prev.gif" width="10" height="10" border="0"></a> 
		<%
		end If
		For I=1 To iPageCount
		If Cint(I)=Cint(iPageCurrent) Then %>
			<b><%=I%></b> 
		<%
		Else
		%>
			<a href="crossSellView.asp?idcat=<%=pcIntCategoryID%>&iPageCurrent=<%=I%>&iPageSize=<%=iPageSize%>&sOrder=<%=sOrder%>"> 
			<%=I%></a> 
		<%
		End If
		Next
			if CInt(iPageCurrent) < CInt(iPageCount) then %>
				<a href="crossSellView.asp?idcat=<%=pcIntCategoryID%>&iPageCurrent=<%=iPageCurrent+1%>&iPageSize=<%=iPageSize%>&sOrder=<%=sOrder%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
		<%
			end If
		End if%>
		</td>
	</tr>
	<%
end if
call closeDb()
%>								

	<tr>
		<td colspan="4" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<td align="center" colspan="4">
		<input type="submit" value="Delete selected" name="submit1" class="submit2" onclick="return(confirm('You are about to completely remove the selected cross selling relationships. If you only want to remove one or more products from a specific relationship, click CANCEL and then select EDIT for that cross selling relationship. Click OK to confirm the removal.'));">&nbsp;
		<input type="button" value="Add new" onClick="location.href='crossSellAdd.asp'">&nbsp;
		<input type="button" value="Edit Settings" onClick="location.href='crossSellSettings.asp'">&nbsp;
		</td>
	</tr>

</table>
</form>
<!--#include file="AdminFooter.asp"-->
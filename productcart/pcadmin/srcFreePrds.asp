<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% 
Server.ScriptTimeout=5400
%>
<%PmAdmin=2%><!--#include file="adminv.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"-->
<% pageTitle = "Manage Categories - Orphaned Products" %>
<% section = "products" %>
<!--#include file="Adminheader.asp"--> 
<%
dim iPageSize
dim query, connTemp, rs, rsTemp, pIdCategory, pcv_ID
dim iPageCurrent

'// MOVE PRODUCTS - START
IF request("action")="update" then
	IF request("Submit1")<>"" then
		call openDb()
		Count=request("count")
		pIdCategory=request("idcategory")
		if validNum(Count) and validNum(pIdCategory) then	
			totalMoved=0
			For k=0 to Count
				if request("C" & k)="1" then
					pcv_ID=request("ID" & k)
					if validNum(pcv_ID) then
						query="INSERT INTO categories_products (idProduct, idCategory) VALUES (" &pcv_ID& "," &pIdCategory& ")"
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=conntemp.execute(query)
						if err.number <> 0 then
							set rs=nothing
							call closeDb()
						  response.redirect "techErr.asp?error="& Server.Urlencode("Error moving a product on srcFreePrds.asp") 
						end If
						set rs=nothing
						totalMoved=totalMoved+1
					end if
				end if
			Next
		end if
		pIdCategory=""
		pcv_ID=""
		call closeDb()
		response.Redirect "srcFreePrds.asp?message=1&total="&totalMoved
	End if
END IF
'// MOVE PRODUCTS - END

'// DELETE PRODUCTS - START

IF request("action")="update" then
	IF request("Submit2")<>"" then
		call openDb()
		Count=request("count")
		if (Count>"0") and (IsNumeric(Count)) then
			totalDeleted=0
			For k=0 to Count
				if request("C" & k)="1" then
					pcv_ID=request("ID" & k)
					if validNum(pcv_ID) then
													
							' delete from taxPrd - QueryA
							query="DELETE FROM taxPrd WHERE idProduct=" &pcv_ID
							set rs=Server.CreateObject("ADODB.Recordset")
							set rs=conntemp.execute(query)
							
							if err.number <> 0 then
								set rs=nothing
								call closeDb()
							  response.redirect "techErr.asp?error="& Server.Urlencode("Error deleting a product in QueryA") 
							end If
							
							' delete product from configSpec_products - QueryB
							query="DELETE FROM configSpec_products WHERE configProduct=" &pcv_ID
							set rs=conntemp.execute(query)
							
							if err.number <> 0 then
								set rs=nothing
								call closeDb()
							  response.redirect "techErr.asp?error="& Server.Urlencode("Error deleting a product in QueryB") 
							end If
							
							' delete product from cs_relationships - QueryC
							query="DELETE FROM cs_relationships WHERE idProduct=" &pcv_ID
							set rs=conntemp.execute(query)
							
							if err.number <> 0 then
								set rs=nothing
								call closeDb()
							  response.redirect "techErr.asp?error="& Server.Urlencode("Error deleting a product in QueryC") 
							end If
							
							' delete product from categories_products - QueryD
							query="DELETE FROM categories_products WHERE idProduct=" &pcv_ID
							set rs=conntemp.execute(query)
							
							if err.number <> 0 then
								set rs=nothing
								call closeDb()
							  response.redirect "techErr.asp?error="& Server.Urlencode("Error deleting a product in QueryD") 
							end If
							
							' delete product from products table - QueryE
							query="UPDATE products SET active=0, removed=-1 WHERE idproduct=" &pcv_ID
							set rs=conntemp.execute(query)
							
							if err.number <> 0 then
								set rs=nothing
								call closeDb()
							  response.redirect "techErr.asp?error="& Server.Urlencode("Error deleting a product in QueryE") 
							end If
							
							set rs=nothing
							totalDeleted=totalDeleted+1
					end if
				end if
			Next
		end if
		pcv_ID=""
		call closeDb()
		response.Redirect "srcFreePrds.asp?message=2&total="&totalDeleted
	End if
END IF

'// DELETE PRODUCTS - END


	iPageSize=25
	
	if request.queryString("iPageCurrent")="" then
		iPageCurrent=1 
	else
		iPageCurrent=server.HTMLEncode(request.querystring("iPageCurrent"))
	end if

	pIdCategory=getUserInput(request.querystring("idCategory"),0)

call opendb()

query="SELECT DISTINCT products.idProduct, products.description, products.sku, products.active FROM products WHERE ((products.idProduct) Not in (SELECT DISTINCT categories_products.idProduct FROM categories_products)) AND products.removed=0 ORDER BY products.description"

Set rsTemp=Server.CreateObject("ADODB.Recordset")     

rstemp.CacheSize=iPageSize
rstemp.PageSize=iPageSize

'response.end
rsTemp.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

if err.number <> 0 then
	set rstemp=nothing
	call closedb()
  	response.redirect "techErr.asp?error="&Server.UrlEncode("Error in page SrcFreePrds. Error: "&err.description)
end If

if rsTemp.eof then 
              
else

	dim iPageCount
	iPageCount=rstemp.PageCount
	
	If Cint(iPageCurrent) > Cint(iPageCount) Then Cint(iPageCurrent)=Cint(iPageCount)
	If Cint(iPageCurrent) < 1 Then iPageCurrent=Cint(1)
	'rstemp.MoveFirst
	rstemp.AbsolutePage=iPageCurrent
end if	

' counting variable for our recordset
   
%>
<form action="srcFreePrds.asp" method="post" class="pcForms">
<table class="pcCPcontent">
	<tr>
    	<td colspan="4" class="pcCPspacer">
        	<% 
			message=request.QueryString("message")
			total=request.QueryString("total")
			if message=1 then
			%>
            <div class="pcCPmessageSuccess"><%=total%> products were moved to the selected category</div>
            <%
			elseif message=2 then
			%>
            <div class="pcCPmessageSuccess"><%=total%> products were removed from the Control Panel</div>
            <%
			end if
			%>
        </td>
	<tr>
		<td colspan="4">The following products are currently not assigned to any category. To assign a product to a category, click on the product to be taken to the <em>Modify Product</em> page where category assignments are defined. To reassign multiple products, use the drop-down at the bottom.</td>
	</tr>
	<tr>
		<td colspan="4" class="pcCPspacer"></td>
	</tr>       
	<tr> 
		<th width="20%" nowrap>SKU</th>
		<th width="80%" colspan="3">Product</th>
	</tr>
	<tr>
		<td colspan="4" class="pcCPspacer"></td>
	</tr>
                      
<% 
Dim pcNoProducts
If rstemp.eof Then
	pcNoProducts=1
	set rstemp=nothing%>
	<tr> 
		<td colspan="4">No Products Found</td>
	</tr>
	<tr>
		<td colspan="4" class="pcCPspacer"></td>
	</tr>
<% ELSE 
	pcNoProducts=0
		Dim Count
		Count = 0

		' set count equal to zero
		count=0
		pcArr=rsTemp.GetRows(iPageSize)
		count=ubound(pcArr,2)
		set rsTemp=nothing
		
		For i=0 to count
								
		pidProduct=pcArr(0,i)
		tempidProduct=pidProduct
		pDescription=pcArr(1,i)
		psku=pcArr(2,i)
		pActive=pcArr(3,i)
		%>
				 
		<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
			<td nowrap="nowrap">
            	<input type="checkbox" name="C<%=i%>" value="1" class="clearBorder">
                <input type="hidden" name="ID<%=i%>" value="<%=tempidProduct%>"> 
				&nbsp;<%=psku%>
            </td>
			<td colspan="2"><a href="FindProductType.asp?id=<%=pidproduct %>" target="_blank"><%=pdescription%></a></td>
			<td>
				<% if pactive=0 then%>
					<img src="images/notactive.gif" width="32" height="16">
				<% else %>
					&nbsp;
				<% end if %>
			</td>
		</tr>
		<%
		Next
END IF
%>

<% If iPageCount>1 Then %>
	<tr>
		<td colspan="4">
            <br />
			<%response.Write("Currently viewing page "& iPageCurrent & " of "& iPageCount & " <br />")%>
				<%' display Next / Prev buttons
                if iPageCurrent > 1 then %>
                <a href="srcFreePrds.asp?iPageSize=<%=iPageSize%>&iPageCurrent=<%=iPageCurrent - 1%>&idCategory=<%=pIdCategory%>"><img src="../pc/images/prev.gif" width="10" height="10" border="0"></a> 
                <% end If
                For I = 1 To iPageCount
                If Cint(I) = Cint(iPageCurrent) Then %>
                <b><%=I%></b>
                <% Else %>
                <a href="srcFreePrds.asp?iPageSize=<%=iPageSize%>&iPageCurrent=<%=I%>&idCategory=<%=pIdCategory%>"><%=I%></a> 
                <% End If %>
                <% Next %>
                <% if CInt(iPageCurrent) < CInt(iPageCount) then %>
                <a href="srcFreePrds.asp?iPageSize=<%=iPageSize%>&iPageCurrent=<%= iPageCurrent + 1%>&idCategory=<%=pIdCategory%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
                <% end If %>
                <hr>
		</td>
	</tr>
<% End If

set rstemp=nothing
call closeDb()

if pcNoProducts=0 then
%>
	<tr>
		<td colspan="4">Move selected products to:
		<%
		cat_DropDownName="idcategory"
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
		cat_EventAction=""
		%>
		<!--#include file="../includes/pcCategoriesList.asp"-->
		<%call pcs_CatList()%>
        <input type="submit" name="Submit1" value="Submit" class="submit2">
        </td>
	</tr>
<%
end if
%>
	<tr>
		<td colspan="4">
        <hr>
        <%
		if pcNoProducts=0 then
		%>
        <input type="submit" name="Submit2" value="Delete Selected" onclick="return(confirm('You are about to remove the selected products from your Control Panel. Are you sure to continue?'));">&nbsp;
        <%
		end if
		%>
        <input type="button" value="Manage Categories" onClick="document.location.href='manageCategories.asp'">
        </td>
	</tr>
</table>
<input type="hidden" name="count" value=<%=count%>>
<input type="hidden" name="action" value="update">
</form>
<!--#include file="Adminfooter.asp"-->
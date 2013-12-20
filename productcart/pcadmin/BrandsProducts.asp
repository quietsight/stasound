<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% 
Dim connTemp, query, rs, rstemp, pcIntBrandID

pcIntBrandID = request("idbrand")
	if not validNum(pcIntBrandID) then
		response.Redirect("BrandsManage.asp")
	end if

call opendb()

if request("action")="update" then
	count=request("count")
		if count<>"" then
			Count1=clng(Count)
			For i=0 to Count1
				if request("C" & i)<>1 then
					IDproduct=request("ID" & i)
					if validNum(IDproduct) and IDproduct<>0 then
						query="UPDATE Products SET IDBrand=0 WHERE IDproduct=" & IDproduct
						set rstemp=Server.CreateObject("ADODB.Recordset")
						set rstemp=connTemp.execute(query)
						set rstemp=nothing
					end if
				end if
			Next
			msg="Product assignments updated successfully."
			msgType=1
		end if
end if	
 
query="SELECT BrandName FROM Brands WHERE IDBrand="&pcIntBrandID
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=connTemp.execute(query)
BrandName=rstemp("BrandName")
query="SELECT idproduct,sku,description,serviceSpec FROM products WHERE active=-1 AND configOnly=0 AND removed=0 AND IDBrand=" & pcIntBrandID
set rstemp=conntemp.execute(query) 
Dim pcArrayBrands, pcIntBrandsCount, iCount
if rstemp.EOF then
	pcIntBrandsCount=0
else	
	pcArrayBrands=rstemp.getRows()
	pcIntBrandsCount=ubound(pcArrayBrands,2)+1
end if
set rstemp=nothing
call closeDb()

pageTitle="Manage Brands - Product Assigned to: " & BrandName

%>
<!--#include file="AdminHeader.asp"-->


	<form name="form1" method="post" action="BrandsProducts.asp" class="pcForms">
    	<input type="hidden" name="IDBrand" value="<%=pcIntBrandID%>">
        <input type="hidden" name="action" value="update">
        <table class="pcCPcontent">
            <tr>
                <td colspan="3" class="pcCPspacer">
                    <% ' START show message, if any %>
                        <!--#include file="pcv4_showMessage.asp"-->
                    <% 	' END show message %>
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    These are the products assigned to under <strong><a href="BrandsEdit.asp?idbrand=<%=pcIntBrandID%>"><%=BrandName%></a></strong>. <a href="BrandsAssign.asp?idbrand=<%=pcIntBrandID%>">Assign other products</a>.
                </td>
            </tr>
            <tr>
                <td colspan="3" class="pcCPspacer"></td>
            </tr>
			<tr>
            	<th width="5%"></th>		
				<th width="10%">SKU</th>
                <th width="85%">Name</th>
			</tr>
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr>
            <%
			if pcIntBrandsCount=0 then
			%>
			<tr>
				<td colspan="3"><div class="pcCPmessage">No products have yet been assigned to this brand.</div></td>
			</tr>
            <%
			else
				for iCount=1 to pcIntBrandsCount
					IDProduct=pcArrayBrands(0,iCount-1)
					pcStrSku=pcArrayBrands(1,iCount-1)
					pcStrDesc=pcArrayBrands(2,iCount-1)
					ServiceSpec=pcArrayBrands(3,iCount-1)
			%>
                <tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
                    <td>
                        <input type="checkbox" name="C<%=iCount%>" value="1" checked>
                        <input type="hidden" name="ID<%=iCount%>" value="<%=IDProduct%>">
                    </td>
                    <td><%=pcStrSku%></td>
                    <td><a href="FindProductType.asp?idproduct=<%=IDProduct%>" target="_blank"><%=pcStrDesc%></a> <%if serviceSpec<>0 then%>(BTO)<%end if%></td>
                </tr>
            <%
				next
			end if
			%>
            <tr>
                <td colspan="3" class="pcCPspacer"></td>
            </tr>
            <tr>
            	<td colspan="3">
					<% If pcIntBrandsCount<>0 then %>
                    <input type="hidden" name="Count" value="<%=pcIntBrandsCount%>">
                    <input type="submit" name="submit" value="Remove Unchecked" class="submit2">&nbsp;
                    <% end if %>
                    <input type="button" value="Assign Other Products" onClick="location='BrandsAssign.asp?idbrand=<%=pcIntBrandID%>';">&nbsp;
                    <input type="button" value="Manage Brands" onClick="location='BrandsManage.asp';">
                </td>
             </tr>            
		</table>
	</form>
<!--#include file="AdminFooter.asp"-->
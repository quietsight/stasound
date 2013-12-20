<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'// Get more category details
'// This page shows category thumbnails only
	query = "SELECT categoryDesc,image FROM Categories WHERE idCategory = " & intIdCategory
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	pcStrCategoryDesc=rs("categoryDesc")
	pcStrCategoryImg=rs("image")
	SET rs=nothing
	
	'// Call SEO Routine
	pcGenerateSeoLinks
	'//
%>
<table class="pcShowCategory">
	<tr>
		<td class="pcShowCategoryImage">
		<%if pcStrCategoryImg<>"" then%>
 			<p><a href="<%=pcStrCatLink%>" <%if scStoreUseToolTip="1" or scStoreUseToolTip="3" then%>onmouseover="javascript:document.getCatPre.idcategory.value='<%=intIdCategory%>'; sav_CatPrecallxml='1'; return runPreCatXML('cat_<%=intIdCategory%>');" onmouseout="javascript: sav_CatPrecallxml=''; hidetip();"<%else%>title="<%=pcStrCategoryDesc%>"<%end if%>><img src="catalog/<%response.write pcStrCategoryImg%>" alt="<%=pcStrCategoryDesc%>"></a></p>
		<% else %>
			<p><a href="<%=pcStrCatLink%>" <%if scStoreUseToolTip="1" or scStoreUseToolTip="3" then%>onmouseover="javascript:document.getCatPre.idcategory.value='<%=intIdCategory%>'; sav_CatPrecallxml='1'; return runPreCatXML('cat_<%=intIdCategory%>');" onmouseout="javascript: sav_CatPrecallxml=''; hidetip();"<%else%>title="<%=pcStrCategoryDesc%>"<%end if%>><img src="catalog/no_image.gif" width="50" height="50" alt="<%=pcStrCategoryDesc%>"></a></p>
		<%end if %>
		</td>
	</tr>
</table>
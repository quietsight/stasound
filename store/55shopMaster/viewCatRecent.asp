<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% if request("nav")="1" then
	section="services"
else
	section="products"
end if
%>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
dim query, conntemp, rsCat, idCategory
	
'// Set Page Title
	pageTitle="Recently Edited Categories"
	pcStrPageName="viewCatRecent.asp"

'// Load Subcategories
	call openDb()
	query="SELECT TOP 10 idCategory, categoryDesc FROM categories WHERE idcategory > 1 ORDER BY pcCats_EditedDate DESC"
	set rsCat=server.CreateObject("ADODB.RecordSet")
	set rsCat=conntemp.execute(query)

	if err.number <> 0 then
		set rsCat=nothing
		call closedb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error getting category list") 
	end If	

%>
<!--#include file="AdminHeader.asp"-->
	<table class="pcCPcontent">
			<% 
			if rsCat.eof then
				response.write "No categories found."
			else
				count=0
				while not rsCat.eof and count<10
					tcategoryDesc=rsCat("categoryDesc")
					tidCategory=rsCat("idCategory")
				%>
					<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
						<td> 
							<a href="modcata.asp?idcategory=<%=tidcategory%>" target="_blank"><%=tcategoryDesc%></a>
							<input type="hidden" name="idcategory<%=iCnt%>" value="<%=tidcategory%>">
						</td>
					</tr>
				<%
				rsCat.MoveNext
				count=count + 1
				Wend
			end if
			Set rsCat=Nothing 
			call closeDb()
			count=0
	        %>
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
            <tr> 
                <td colspan="2"> 
                <input name="back" type="button" onClick="document.location.href='manageCategories.asp?'" value="Manage Categories" class="ibtnGrey">
                </td>
            </tr>
        </table>
<!--#include file="AdminFooter.asp"-->
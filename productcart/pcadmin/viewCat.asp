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
<% 'on error resume next
dim parent, pcIntHidden
parent=request("parent")
nav=request("nav")
top=request("top")
pcIntHidden=request("hidden")
	if not validNum(pcIntHidden) then
		pcIntHidden=0
	end if

dim query, conntemp, rs, rstemp, pcStrParentName
dim i, idCategory, priority
sMode=request.form("submitForm")
if sMode<>"" then
	call openDb()
	parent=request.form("parent")
	nav=request.form("nav")
	top=request.form("top")
	iCnt=request.form("iCnt")
	set rs=server.CreateObject("ADODB.RecordSet")
	for i=1 to iCnt
		idCategory=request.form("idCategory"&i)
		priority=request.form("priority"&i)
		query="UPDATE categories SET "
		query=query & "priority=" &priority
		query=query & " WHERE idCategory="&idCategory&";"
		on error resume next
		rs.open query,conntemp
	next
	set rs=nothing
	call closeDb()	
	response.redirect "viewCat.asp?top="&top&"&nav="&nav&"&parent="&parent
	response.end
end if

'// Get Parent Category Name
	call openDb()
	query="SELECT categoryDesc FROM categories WHERE idCategory="&parent
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)
	pcStrParentName=rstemp("categoryDesc")
	set rstemp=nothing
	
'// Set Page Title
	pageTitle="Order Subcategories of &quot;" &  pcStrParentName & "&quot;"
	pcStrPageName="viewCat.asp"

'// Load Subcategories
	query="SELECT idCategory, categoryDesc, priority FROM categories WHERE idcategory > 1 AND idparentCategory="&parent&" ORDER BY priority, categoryDesc ASC"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	if err.number <> 0 then
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error getting category list") 
	end If	

	if rs.eof then
		set rs=nothing
		call closedb()
		response.redirect "modcata.asp?idcategory=" & parent &"&msg="& Server.Urlencode("No subcategories found.")
	end If
%>
<!--#include file="AdminHeader.asp"-->
<form name="form1" method="post" action="viewCat.asp" class="pcForms">
	<input type="hidden" name="parent" value="<%=parent%>">
	<input type="hidden" name="nav" value="<%=nav%>">
	<input type="hidden" name="top" value="<%=top%>">
	<table class="pcCPcontent">
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr> 					
				<th width="5%" nowrap>Order</th>
				<th width="95%" nowrap>Category Name</th>
			</tr>
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<% 
			Dim iCnt
			iCnt=0
			do until rs.eof
				tcategoryDesc=rs("categoryDesc")
				tidCategory=rs("idCategory")
				tpriority=rs("priority")
				iCnt=iCnt+1 %>
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
					<td align="center" nowrap> 
						<input type="text" name="priority<%=iCnt%>" size="2" maxlength="4" value="<%=tpriority%>">
					</td>
					<td> 
						<a href="modcata.asp?idcategory=<%=tidcategory%>" target="_blank"><%=tcategoryDesc%></a>
						<input type="hidden" name="idcategory<%=iCnt%>" value="<%=tidcategory%>">
					</td>
				</tr>
				<%
				rs.movenext
			loop
			set rs=nothing
			call closedb()
			%>
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
            <tr> 
                <td colspan="2"> 
                <input type="hidden" name="iCnt" value="<%=iCnt%>">
                <input type="hidden" name="CntParent" value="<%=Cntparent%>">
                <input name="submitForm" type="submit" class="submit2" value="Update Category Order">&nbsp;
                <% if pcIntHidden=0 then %>
                <input name="preview" type="button" value="Preview" onClick="window.open('../pc/viewcategories.asp?idcategory=<%=parent%>')">&nbsp;
                <% end if %>
                <% if parent>1 then %>
                <input name="edit" type="button" value="Edit Parent" onClick="document.location.href='modcata.asp?idcategory=<%=parent%>'">&nbsp;
                <% end if %>
                <input name="back" type="button" onClick="document.location.href='manageCategories.asp?top=<%=top%>&nav=<%=nav%>&parent=<%=parent%>'" value="Back">
                <% if pcIntHidden=1 then %>
                <div class="pcSmallText" style="margin-top: 10px;">This category is hidden in the storefront.</div>
                <% end if %>
                </td>
            </tr>
        </table>
	</form>
<!--#include file="AdminFooter.asp"-->
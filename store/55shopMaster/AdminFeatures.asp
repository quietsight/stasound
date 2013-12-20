<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Featured Products" %>
<% section="specials" %>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
Dim rstemp, connTemp, strSQL, pid

'****************************
'* Action Items
'****************************

	' UPDATE list of featured products
		sMode=Request("Submit")
		If sMode <> "" Then
		
			call openDb()
			if (request("prdlist")<>"") and (request("prdlist")<>",") then
			prdlist=split(request("prdlist"),",")
			Set rstemp=Server.CreateObject("ADODB.Recordset")   
			For i=lbound(prdlist) to ubound(prdlist)
				id=prdlist(i)
				If (id<>"0") and (id<>"") then
					query="UPDATE products SET showInHome=-1 WHERE idproduct="&id
					Set rstemp=connTemp.execute(query)
				end if
			Next
			Set rstemp=nothing
			call closeDb()
			response.redirect "AdminFeatures.asp?s=1&msg="&msg
			end if
			
	
			' UPDATE featured products order
			PCount=request("PCount")
			if (pCount<>"") and (PCount<>"0") then
				set rstemp=Server.CreateObject("ADODB.Recordset")
				For i=1 to Cint(PCount)
					idproduct=request("IDFP" & i)
					OrdInHome=request("FPOrd" & i)
					query="UPDATE products SET pcprod_OrdInHome=" & OrdInHome & " WHERE idproduct="&idproduct
					set rstemp=connTemp.execute(query)
				Next
				set rstemp=nothing
				call closeDb()
				response.redirect "AdminFeatures.asp?s=1&msg="&Server.URLEncode("Product Orders were updated successfully!")
			end if
		End If

'****************************
'* END Action Items
'****************************

	' Paging and sorting
	
	if request("iPageCurrent")="" then
    iPageCurrent=1 
	else
		iPageCurrent=Request("iPageCurrent")
	end If
	
	Dim strORD
	
	strORD=request("order")
	if strORD="" then
		strORD="pcprod_OrdInHome"
	End If
	
	strSort=request("sort")
	if strSort="" Then
		strSort="ASC"
	End If 
	
	call openDb()

	' gets group assignments
	query="SELECT pcprod_OrdInHome,idproduct,sku,description FROM products WHERE active=-1 AND showInHome=-1 ORDER BY "& strORD &" "& strSort

	Set rstemp=Server.CreateObject("ADODB.Recordset")   
	rstemp.CacheSize=100
	rstemp.PageSize=100

	rstemp.Open query, connTemp, adOpenStatic, adLockReadOnly
	dontshow="0"
	If rstemp.eof Then 
		dontshow="1"
	end if
	
	' Find out if all products have been set as Featured Product
	query="SELECT idProduct FROM products WHERE active=-1 AND configOnly=0 AND removed=0 AND showInHome=0"
	Set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if rs.EOF then
		pcAllFeatured = 1
	end if
	set rs = nothing
%>
	<table class="pcCPcontent">
		<tr>
			<td>
			<p>Featured products are shown on the &quot;<a href="http://wiki.earlyimpact.com/productcart/marketing-featured_products" target="_blank">home page</a>&quot; and on the &quot;featured product&quot; page, in the order specified below.&nbsp;<a href="http://wiki.earlyimpact.com/productcart/marketing-featured_products" target="_blank"><img src="images/pcv3_infoIcon.gif" alt="More information about this feature." width="16" height="16" border="0"></a></p>
            <div class="cpOtherLinks"><% if pcAllFeatured <> 1 then %><a href="JavaScript:;" onClick="javascript:document.getElementById('FindProducts').style.display=''">Add New</a>&nbsp;|&nbsp;<% end if %>View the store's <a href="../pc/home.asp" target="_blank">home page</a>&nbsp;|&nbsp;View the store's <a href="../pc/showfeatured.asp" target="_blank">featured products page</a>&nbsp;|&nbsp;<a href="manageHomePage.asp">Manage the home page</a></div>
            <table id="FindProducts" class="pcCPcontent" style="display:none;">
                <tr>
                    <td>
                    <%
                        src_FormTitle1="Find Products"
                        src_FormTitle2="Add New Featured Products"
                        src_FormTips1="Use the following filters to look for products in your store."
                        src_FormTips2="Select the products that you would like to add to your list of featured products."
                        src_IncNormal=1
                        src_IncBTO=1
                        src_IncItem=0
                        src_Featured=2
                        src_DisplayType=1
                        src_ShowLinks=0
                        src_FromPage="AdminFeatures.asp"
                        src_ToPage="AdminFeatures.asp?submit=yes"
                        src_Button1=" Search "
                        src_Button2=" Add as Featured Product "
                        src_Button3=" Back "
                        src_PageSize=15
                    %>
                        <!--#include file="inc_srcPrds.asp"-->
                    </td>
                </tr>
            </table>

			<% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
        </td>
    </tr>
	 
	 <tr>
	 	<td>
	
			<form action="AdminFeatures.asp" name="form1" class="pcForms">
				<table class="pcCPcontent">
					<tr>
						<th width="5%" nowrap> 
						<a href="AdminFeatures.asp?iPageCurrent=<%=iPageCurrent%>&order=pcprod_OrdInHome&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="AdminFeatures.asp?iPageCurrent=<%=iPageCurrent%>&order=pcprod_OrdInHome&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Order</th>
						<th width="20%" nowrap> 
							<a href="AdminFeatures.asp?iPageCurrent=<%=iPageCurrent%>&order=sku&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="AdminFeatures.asp?iPageCurrent=<%=iPageCurrent%>&order=sku&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;SKU</th>
						<th width="75%" nowrap colspan="2">
							<a href="AdminFeatures.asp?iPageCurrent=<%=iPageCurrent%>&order=description&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="AdminFeatures.asp?iPageCurrent=<%=iPageCurrent%>&order=description&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Product&nbsp;&nbsp;<% if pcAllFeatured <> 1 then %><a href="JavaScript:;" onClick="javascript:document.getElementById('FindProducts').style.display=''" class="pcSmallText">Add New</a><% end if %>
						</th>
					</tr>
                    <tr>
                    	<td colspan="4" class="pcCPspacer"></td>
                    </tr>
										 
					<%If rstemp.eof Then%>
					<tr> 
						<td colspan="4">
							<div class="pcCPmessage">No Featured Items Found</div>
						</td>
					</tr>
					<% Else 
						rstemp.MoveFirst
						' get the max number of pages
						Dim iPageCount
						iPageCount=rstemp.PageCount
						If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
						If iPageCurrent < 1 Then iPageCurrent=1
														
						' set the absolute page
						rstemp.AbsolutePage=iPageCurrent
						Dim Count
						hCnt=0
						Count=0
						Do While NOT rstemp.EOF
							Count=Count+1
					%>
													
					<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
					<%FPOrd=rstemp("pcprod_OrdInHome") 
					if FPOrd<>"" then
					else
						FPOrd="0"
					end if%>
						<td valign="top">
							<input type=text name="FPOrd<%=Count%>" value="<%=FPOrd%>" size=2>
							<input type=hidden name="IDFP<%=Count%>" value="<%=rstemp("idproduct")%>">
						</td>
						<td valign="top"> 
							<a href="FindProductType.asp?id=<%=rstemp("idproduct")%>" target="_blank"><%=rstemp("sku")%></a>
						</td>
						<td valign="top"> 
							<a href="FindProductType.asp?id=<%=rstemp("idproduct")%>" target="_blank"><%=rstemp("description") %></a>
						</td>
						<td align="right" width="1%" valign="top"> 
							<a href="javascript:if (confirm('You are about to remove this item as a featured item. Are you sure you want to complete this action?')) location='delFeaturesb.asp?idproduct=<%= rstemp("idproduct") %>'"><img src="images/pcIconDelete.jpg"></a>
						</td>
					</tr>
													
					<%
					rstemp.MoveNext
					Loop
					%>
					<input type=hidden name="PCount" value="<%=Count%>">										
					<%End If %>
					
					<%if dontshow="0" then%>
					<tr>
						<td colspan="4" class="pcCPspacer"></td>
					</tr>
					<tr>
						<td colspan="4" align="center">
							<input type="Submit" name="Submit" value="Update Order" class="submit2">
							&nbsp;<input type="button" onClick="location.href='manageHomePage.asp'" value="Manage Home Page">
						</td>
					</tr>
					<%
					end if
					set rstemp = nothing
					call closeDb()

					%>
				</table>
			</form>
		</td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Affiliates" %>
<% section="misc" %>
<%PmAdmin=8%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../pc/pcSeoFunctions.asp"-->
<!--#include file="AdminHeader.asp"-->
<script language="JavaScript" type="text/javascript" src="../includes/spry/xpath.js"></script>
<script language="JavaScript" type="text/javascript" src="../includes/spry/SpryData.js"></script>
<script type="text/javascript">
	var dscategories = new Spry.Data.XMLDataSet("pcSpryCategoriesXML.asp?CP=1&idRootCategory=<%=pcv_IdRootCategory%>", "categories/category",{sortOnLoad:"categoryDesc",sortOrderOnLoad:"ascending",useCache:false});
	dscategories.setColumnType("idCategory", "number");
	dscategories.setColumnType("idParentCategory", "number");
	var dsProducts = new Spry.Data.XMLDataSet("categories_productsxml.asp?mode=4&x_idCategory={dscategories::idCategory}", "/root/row");
</script>

<% sMode=request("action")

if sMode <> "" then
	sMode="1"
	idRequestProduct=request.Form("product")
	idRequestAffiliate=request.Form("affiliate")
end If

Dim rs, connTemp, strSQL, pid

if request("iPageCurrent")="" then
	iPageCurrent=1 
else
	iPageCurrent=Request("iPageCurrent")
end If

'sorting order
Dim strORD

strORD=request("order")
if strORD="" then
	strORD="pcAff_Active"
End If

strSort=request("sort")
if strSort="" Then
	strSort="ASC"
End If 

dim strDateFormat
strDateFormat="mm/dd/yyyy"
if scDateFrmt="DD/MM/YY" then
	strDateFormat="dd/mm/yyy"
end if

call openDb()

' gets group assignments
query="SELECT idaffiliate, affiliateName, affiliateCompany, commission, affiliateEmail, pcAff_Active, pcAff_website FROM affiliates WHERE idaffiliate > 1 ORDER BY "& strORD &" "& strSort

Set rs=Server.CreateObject("ADODB.Recordset")

rs.CacheSize=200
rs.PageSize=200

rs.Open query, connTemp, adOpenStatic, adLockReadOnly

If not rs.eof Then
	rs.MoveFirst

	' get the max number of pages
	Dim iPageCount
	iPageCount=rs.PageCount
	If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
	If iPageCurrent < 1 Then iPageCurrent=1
		
	' set the absolute page
	rs.AbsolutePage=iPageCurrent

End if
	
%>
<form method="POST" action="Affiliate_action.asp" class="pcForms">
	<div style="width: 790px; overflow:auto;">
	<table class="pcCPcontent">
        <tr>
            <td colspan="5" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>

		<% If rs.eof Then 
				affrpt="0"%>
				<tr>
					<td colspan="7">
						<div class="pcCPmessage">No Affiliates Found!</div>
					</td>
				</tr>
		<% Else%>
			<tr> 
				<th align="center" nowrap><a href="AdminAffiliates.asp?iPageCurrent=<%=iPageCurrent%>&order=idaffiliate&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="AdminAffiliates.asp?iPageCurrent=<%=iPageCurrent%>&order=idaffiliate&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;ID</th>
				<th align="left" nowrap><a href="AdminAffiliates.asp?iPageCurrent=<%=iPageCurrent%>&order=affiliateName&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="AdminAffiliates.asp?iPageCurrent=<%=iPageCurrent%>&order=affiliateName&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Name</th>
				<th align="left" nowrap>Company</th>
				<th align="left" nowrap>Web Site</th>
				<th align="center" nowrap>Commission (%)</th>
				<th align="center" nowrap>Active</th>        
				<th align="center" nowrap></th>
			</tr>
			<tr>
				<td colspan="7" class="pcCPspacer"></td>
			</tr>
			<%
			Do While NOT rs.EOF
				AffiliateId=rs("idaffiliate")
				AffiliateName=rs("affiliateName")
				AffiliateCompany=rs("affiliateCompany")
				AffiliateCommission=rs("commission")
				AffiliateEmail=rs("affiliateEmail")
				AffiliateActive=rs("pcAff_Active")
				AffiliateWeb=rs("pcAff_website")
					if trim(AffiliateWeb)="" or isNull(AffiliateWeb) then
						AffiliateWeb=""
					end if
					if instr(AffiliateWeb,"http://")=0 and instr(AffiliateWeb,"https://")=0 then
						AffiliateWeb="http://" & AffiliateWeb
					end if
					tempURL=replace(AffiliateWeb,"//","/")
					tempURL=replace(tempURL,"https:/","https://")
					tempURL=replace(tempURL,"http:/","http://")
					AffiliateWeb=tempURL
		
				pcv_ShowNotActive=0
				if AffiliateActive="1" then
				else
					pcv_ShowNotActive=1
				end if %>
													
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
				<td align="center"><%=AffiliateId%></td>
				<td><a href="modAffa.asp?idAffiliate=<%=AffiliateId%>"><%=AffiliateName%></a></td>
				<td><a href="modAffa.asp?idAffiliate=<%=AffiliateId%>"><%=AffiliateCompany%></a></td>
				<td><% if trim(AffiliateWeb)<>"" then %><a href="<%=AffiliateWeb%>" target="_blank"><%=AffiliateWeb%></a><%end if%></td>        
				<td align="center"><%=AffiliateCommission%></td>
        <td align="center"><%if pcv_ShowNotActive=1 then%>No<%else%>Yes<%end if%>
				<td align="center" class="cpLinksList"><a href="pcAffiliateSendEmail.asp?idaffiliate=<%=AffiliateId%>">Notify</a>&nbsp;|&nbsp;<a href="#links">Links</a>&nbsp;|&nbsp;<a href="modAffa.asp?idAffiliate=<%=AffiliateId%>">Edit</a>&nbsp;|&nbsp;<a href="javascript:if (confirm('You are about to permanantly delete this affiliate from the database. Are you sure you want to complete this action?')) location='delAffb.asp?idAffiliate=<%=AffiliateId%>'">Remove</a></td>
				</tr>
				<% 
				rs.MoveNext
			Loop
			set rs=nothing
		End If %>
	
		<tr>
			<td colspan="7" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td align="center" colspan="7">
			<input type="button" value="Add New Affiliate" onClick="location.href='instAffa.asp'">&nbsp;
			<input type="button" value="Back" onClick="javascript:history.back()">
			</td>
		</tr>          
	</table>
    </div>
</form>

<% if not affrpt="0" then %>
	<table class="pcCPcontent">
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>         
		<tr> 
			<th colspan="2">Affiliate Sale Reports</th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2">
				<form action="salesReportAffiliate.asp" name="aff_form" target="_blank" class="pcForms">
					<p>Specify a date range to view all sales by affiliate in that period. <b>Note</b>: You must enter both dates in the format <%=strDateFormat%>. Then enter the affiliate ID, or choose from the drop-down list.</p>
					<p style="padding-top:10px;">From:
					<input class="textbox" type="text" size="10" name="FromDate" value="<%=dtInputStr%>">
					To:
					<input class="textbox" type="text" size="10" name="ToDate" value="<%=dtInputStr%>">
					</p>
					<p style="padding-top:10px;">ID: <input type="text" size="5" maxlength="100" name="idaffiliate1">
					<b>OR </b> 
					Name: 
					<% query="SELECT idAffiliate, affiliateName FROM affiliates WHERE idaffiliate>1 ORDER BY affiliateName"
					set rsAffObj=Server.CreateObject("ADODB.RecordSet")
					Set rsAffObj=connTemp.execute(query)
					%>
					<select name="idaffiliate2">
						<option value="0">Select Affiliate</option>
						<option value="ALL">Show All</option>
						<% if NOT rsAffObj.eof then
							do until rsAffObj.eof %>
								<option value="<%=rsAffObj("idAffiliate")%>"><%=rsAffObj("affiliateName")%></option>
								<% rsAffObj.moveNext
							loop 
							set rsAffObj=nothing
						End If %>
					</select>
					</p>
					<p style="padding-top:10px;">
					<input type="submit" value="Search" name="submit" class="submit2">&nbsp;
					<input type="reset" name="Submit2" value="Clear">
					</p>
				</form>
			</td>
		</tr>
	</table>
	<br>
	<table class="pcCPcontent">    
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>      
		<tr> 
			<th colspan="2">Affiliate Payment History</th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td valign="top" nowrap>Affiliate Name:</td>
			<td width="95%" valign="top">
			<form name="form1" action="ReportAffiliateHistory.asp" method="post" target="_blank" class="pcForms">
				<% query="SELECT idAffiliate, affiliateName FROM affiliates WHERE idaffiliate>1 ORDER BY affiliateName"
				set rsAffObj=server.CreateObject("ADODB.RecordSet")
				Set rsAffObj=connTemp.execute(query) %>
				<select name="idaffiliate">
				<% if not rsAffObj.eof then
					do until rsAffObj.eof %>
						<option value="<%=rsAffObj("idAffiliate")%>"><%=rsAffObj("affiliateName")%></option>
						<% rsAffObj.moveNext
					loop 
					set rsAffObj=nothing
				End If %>
				</select>&nbsp;<input name="submit" type="submit" value="View history" class="submit2">
			</form>
			</td>
		</tr>
	</table>
	<br>

	<form method="post" name="links" action="AdminAffiliates.asp?action=1" class="pcForms">
		<table class="pcCPcontent">
			<tr>
				<td colspan="2" class="pcCPspacer"><a name="links"></a></td>
			</tr>
			<tr class="normal">
				<th colspan="2">Generate Affiliate Product Links</th>
			</tr>
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr> 
				<td width="15%" nowrap>Select Affiliate:</td>
				<td width="85%"> 
					<% call openDb()
					query="SELECT idaffiliate, affiliateName FROM affiliates ORDER BY affiliateName ASC"
					set rsAff=Server.CreateObject("adodb.recordset")
					set rsAff=conntemp.execute(query)

					if err.number <> 0 then
						response.redirect "techErr.asp?error="& Server.Urlencode("Error in retreiving products from database: "&Err.Description) 
					end If %>
					<select name="affiliate">
					<% do until rsAff.eof
						pIntIdAffiliate=rsAff("idaffiliate")
						pStrAffiliateName=rsAff("affiliateName")
						if sMode="1" And Cint(idRequestAffiliate)= Cint(pIntIdAffiliate) then
							pRequestAffiliateName=pStrAffiliateName
							%>
							<option value="<%=pIntIdAffiliate%>" selected> 
						<% else %>
							<option value="<%=pIntIdAffiliate%>"> 
						<% end if %>
						<%=pStrAffiliateName%>
						</option>
						<% rsAff.movenext
					loop
					set rsAff=nothing %>
					</select>
				</td>
			</tr>
			<tr>
                <td valign="top">Select a product:</td>
                <td>
			    <div spry:region="dscategories" id="categorySelector">
			      <div spry:state="loading"><img src="images/pc_AjaxLoader.gif"/>&nbsp;Loading Categories...</div>
			      <div spry:state="ready">
			        <select spry:repeatchildren="dscategories" name="categorySelect" onchange="document.links.product.disabled = true; dscategories.setCurrentRowNumber(this.selectedIndex);">
			          <option spry:if="{ds_RowNumber} == {ds_CurrentRowNumber}" value="{idCategory}" selected="selected">{categoryDesc} {pcCats_BreadCrumbs}</option>
			          <option spry:if="{ds_RowNumber} != {ds_CurrentRowNumber}" value="{idCategory}">{categoryDesc} {pcCats_BreadCrumbs}</option>
		            </select>
			        </div>
		        </div>
			    <br>
			    <div spry:region="dsProducts dscategories" id="productSelector">
			      <div spry:state="loading"><img src="images/pc_AjaxLoader.gif"/>&nbsp;Loading Products...</div>
			      <div spry:state="ready">
			        <select spry:repeatchildren="dsProducts" id="productSelect" name="product">
			          <option spry:if="{dscategories::ds_RowNumber} == {dscategories::ds_CurrentRowNumber}" value="{idProduct}" selected="selected">{description}</option>
			          <option spry:if="{dscategories::ds_RowNumber} != {dscategories::ds_CurrentRowNumber}" value="{idProduct}">{description}</option>
		            </select>
			        </div>
	          	</div>
              	</td>
		  	</tr>
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr class="normal"> 
				<td colspan="2" align="center">
				<input type="submit" name="submit1" value="Generate Links" class="submit2">
				</td>
			</tr>
			<% If sMode="1" then
				call openDb()
				query="SELECT idproduct, description FROM products WHERE idproduct="&idRequestProduct
				set rsPrd=Server.CreateObject("adodb.recordset")
				set rsPrd=conntemp.execute(query)
				pProductDesc=rsPrd("description")
				set rsPrd=nothing
				call closeDb()
				%>
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
				<tr>
					<td colspan="2">
						<table class="pcCPcontent">
							<tr>
								<td colspan="2">Link to <strong><%=pProductDesc%></strong> for <strong><%=pRequestAffiliateName%></strong>:</td>
							</tr>
							<tr> 
								<td width="6%">
									<a class="highlighttext" href="javascript:HighlightAll('links.link1')"><img src="images/edit2.gif" width="25" height="23" border="0"></a>
								</td>
								<td width="94%">
									<%
									'// SEO Links
									'// Build Navigation Product Link
									'// Get the first category that the product has been assigned to, filtering out hidden categories
										query="SELECT categories_products.idCategory FROM categories_products INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE categories_products.idProduct="& idRequestProduct &" AND categories.iBTOhide<>1 AND categories.pccats_RetailHide<>1"
										call openDb()
										set rs=server.CreateObject("ADODB.RecordSet")
										set rs=conntemp.execute(query)
										if not rs.EOF then
											pIdCategory=rs("idCategory")
										else
											pIdCategory=1
										end if
										set rs=nothing
										call closeDb()

										if scSeoURLs=1 then
											pcStrPrdLink=pProductDesc & "-" & pIdCategory & "p" & idRequestProduct & ".htm"
											pcStrPrdLink=removeChars(pcStrPrdLink)
											pcStrPrdLink=pcStrPrdLink & "?"
										else
											pcStrPrdLink="viewPrd.asp?idproduct=" & idRequestProduct &"&"
										end if
									'//
									
									tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/"&pcStrPrdLink&"idaffiliate="&idRequestAffiliate),"//","/")
									tempURL=replace(tempURL,"http:/","http://") %>
									<input type="text" name="link1" size="100" value="<%=tempURL%>">
								</td>
							</tr>
							<tr> 
								<td colspan="2">Link to the <strong>storefront</strong> for <%=pRequestAffiliateName%>:</td>
							</tr>
							<tr> 
								<td width="6%">
									<a class="highlighttext" href="javascript:HighlightAll('links.link2')"><img src="images/edit2.gif" width="25" height="23" border="0"></a>
								</td>
								<td width="94%">
									<% tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/home.asp?idaffiliate="&idRequestAffiliate),"//","/")
										tempURL=replace(tempURL,"http:/","http://") %>
									<input type="text" name="link2" size="100" value="<%=tempURL%>">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			<% end if %>                
		</table>
	</form>
<% end if %>  
<!--#include file="AdminFooter.asp"-->
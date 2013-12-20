<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin="1*6*"%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/taxsettings.asp"-->

<%
Dim query, rsVAT, rsVATTemp, connTemp, pcDontshow, pIdProduct, pSku, pDescription, pcMessage, pAction, sMode, strORD, iPageCount, strCol, Count

pcv_intVATID=Request("VATID")
	
'****************************
'* Action Items
'****************************

	'// REMOVE selected product from VAT Category
	pAction=request.QueryString("action")
	if pAction = "remove" then
		pIdproduct=request.Querystring("idproduct")
			if trim(pIdproduct)="" then
				 response.redirect "msg.asp?message=2"
			end if
		call openDb()
		query="DELETE FROM pcProductsVATRates WHERE idproduct="&pIdproduct&" AND pcVATRate_ID="&pcv_intVATID&";"
		set rsVAT=Server.CreateObject("ADODB.Recordset")
		set rsVAT=conntemp.execute(query)		
			if err.number <> 0 then
				set rsVAT=nothing
				call closedb()
				response.redirect "techErr.asp?error="& Server.Urlencode("Error in delSpc: "&Err.Description) 
			end If				
		set rsVAT=nothing
		call closeDb()
		response.Redirect("manageVATCategories.asp?s=1&msg=" & Server.URLEncode("The product was successfully removed from the VAT Category") & "&VATID=" & pcv_intVATID)
	end if
	
	'// REMOVE - Bulk remove products
	if request("action")="update" then
		call openDb()	
		Count=request("Count")
		For k=1 to Count
			if not request("C" & k)="1" then
				query="DELETE FROM pcProductsVATRates WHERE pcProductsVATRates.idProduct=" & request("idProduct" & k) & ";"
				set rsVAT=Server.CreateObject("ADODB.Recordset")
				set rsVAT=connTemp.execute(query)
				set rsVAT=nothing
			end if
		Next
		call closeDb()
		response.redirect "manageVATCategories.asp?msg="&msg&"&VATID=" & pcv_intVATID
	end if	
	
	'// Insert into VAT Category
	sMode=Request("Submit")
	If sMode <> "" Then
		prdlist=split(request("prdlist"),",")
		call openDb()
		For i=lbound(prdlist) to ubound(prdlist)
		id=prdlist(i)
		If (id<>"0") and (id<>"") then				
			query="INSERT INTO pcProductsVATRates (idProduct, pcVATRate_ID) VALUES ("&id&","&pcv_intVATID&");"
			Set rsVAT=Server.CreateObject("ADODB.Recordset")
			set rsVAT=connTemp.execute(query)
			set rsVAT=nothing
		end if
		Next
		call closedb()	
		response.redirect "manageVATCategories.asp?msg="&msg&"&VATID=" & pcv_intVATID	
	End If
	
'****************************
'* END Action Items
'****************************
	
	'// Paging and sorting
	
	if request("iPageCurrent")="" then
    iPageCurrent=1 
	else
		iPageCurrent=Request("iPageCurrent")
	end If
	
	strORD=request("order")
	if strORD="" then
		strORD="description"
	End If
	
	strSort=request("sort")
	if strSort="" Then
		strSort="ASC"
	End If 
	
	call openDb()

	' Retrieve VAT Category Products
	query="SELECT pcProductsVATRates.idProduct, products.description, products.sku, products.idProduct, pcVATRates.pcVATRate_Category, pcVATRates.pcVATRate_Rate, pcVATCountries.pcVATCountry_State, pcVATCountries.pcVATCountry_Code "
	query=query&"FROM pcProductsVATRates, products, pcVATRates, pcVATCountries "
	query=query&"WHERE pcProductsVATRates.idProduct=products.idProduct AND pcProductsVATRates.pcVATRate_ID=pcVATRates.pcVATRate_ID AND pcVATCountries.pcVATCountry_Code=pcVATRates.pcVATCountry_Code AND pcVATRates.pcVATRate_ID="&pcv_intVATID&" " 		
	query=query&"ORDER BY "& strORD &" "& strSort	
	Set rsVAT=Server.CreateObject("ADODB.Recordset")   
	rsVAT.CacheSize=100
	rsVAT.PageSize=100	
	rsVAT.Open query, connTemp, adOpenStatic, adLockReadOnly

		if err.number <> 0 then
			SET rsVAT=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error in retreiving VAT Category Products from database on manageVATCategories: "&Err.Description) 
		end If
	
	' Set a variable if no specials are found
	if rsVAT.EOF then
		pcDontshow = 1
	end if
	
	query="SELECT pcVATRates.pcVATRate_Category, pcVATRates.pcVATRate_Rate, pcVATCountries.pcVATCountry_State, pcVATCountries.pcVATCountry_Code "
	query=query&"FROM pcVATRates, pcVATCountries "
	query=query&"WHERE pcVATCountries.pcVATCountry_Code=pcVATRates.pcVATCountry_Code  AND pcVATRates.pcVATRate_ID="&pcv_intVATID&";" 
	Set rsVATTemp=Server.CreateObject("ADODB.Recordset")
	rsVATTemp.Open query, connTemp, adOpenStatic, adLockReadOnly
	if rsVATTemp.EOF then
		pcNoCategory = 1
	else
		pcv_strVATState = rsVATTemp("pcVATCountry_State")
		pcv_strVATCode = rsVATTemp("pcVATCountry_Code")
		pcv_strVATCategory = rsVATTemp("pcVATRate_Category")
	end if
	set rsVATTemp = nothing
%>
<% pageTitle="Products in VAT Category &quot;" & pcv_strVATCategory & "&quot; - " & pcv_strVATState & " (" & pcv_strVATCode & ")" %>
<% section="taxmenu" %>
<!--#include file="AdminHeader.asp"-->

	<table class="pcCPcontent">
		<% If pcNoCategory = 1 then '// Category Not Found %>		
		<tr>
			<td colspan="6">We could not locate the specified Category.</td>
		</tr>
		<% Else %>
		<tr>
			<td colspan="6">
				<p>Add/remove products to/from this VAT Category. All products not assigned to a VAT category use the default VAT rate.</p>
				<ul>
					<li><% if ptaxVATRate_Code<>"" AND pcNoCategory<>1 then %><a href="JavaScript:;" onClick="javascript:document.getElementById('FindProducts').style.display=''">Add Products to this Category</a><% end if %>
					  <% if ptaxVATRate_Code<>"" AND pcNoCategory<>1 then %>
					</li>
					<li><a href="EditVATCategory.asp?VATID=<%=pcv_intVATID%>">Modify this Category</a></li>
					<li><a href="AdminTaxSettings_VAT.asp">Manage VAT Settings</a></li>
				</ul>
                
				<% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>

            </td>
		</tr>	
		<% End If %>
		<tr>
			<td colspan="6">				
				<table id="FindProducts" class="pcCPcontent" style="display:none;">
					<tr>
						<td>
						<%
						src_FormTitle1="Find Products"
						src_FormTitle2="Manage VAT - Add Products"
						src_FormTips1="Use the following filters to look for products in your store."
						src_FormTips2="Select the products that you would like to add to a VAT Category."							
						src_IncNormal=1
						src_IncBTO=1
						src_IncItem=0
						src_DisplayType=1
						src_ShowLinks=0							
						src_FromPage="manageVATCategories.asp?VATID=" & pcv_intVATID
						src_ToPage="manageVATCategories.asp?submit=yes&VATID=" & pcv_intVATID
						src_Button1=" Search "
						src_Button2=" Add Selected Products "
						src_Button3=" Back "
						src_PageSize=20
						UseSpecial=1
						session("srcprd_from")=""
						session("srcprd_where")=" AND products.idproduct NOT IN (SELECT idProduct FROM pcProductsVATRates) "
						%>
						<!--#include file="inc_srcPrds.asp"-->						
						</td>
					</tr>
				</table>
            </td>
		</tr>
		<tr>
			<td colspan="6" class="pcCPspacer"></td>
		</tr>
	</table>
	<form name="form1" method="post" action="manageVATCategories.asp?action=update" class="pcForms">
	<table class="pcCPcontent">
		<tr>
			<th width="7%" nowrap>
				<a href="manageVATCategories.asp?iPageCurrent=<%=iPageCurrent%>&VATID=<%=pcv_intVATID%>&order=sku&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" alt="Sort Ascending"></a><a href="manageVATCategories.asp?iPageCurrent=<%=iPageCurrent%>&VATID=<%=pcv_intVATID%>&order=sku&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" alt="Sort Descending"></a>			</th>
			<th width="8%" nowrap>SKU</th>
			<th width="7%" nowrap> 
				<a href="manageVATCategories.asp?iPageCurrent=<%=iPageCurrent%>&VATID=<%=pcv_intVATID%>&order=description&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" alt="Sort Ascending"></a><a href="manageVATCategories.asp?iPageCurrent=<%=iPageCurrent%>&VATID=<%=pcv_intVATID%>&order=description&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" alt="Sort Descending"></a>			</th>
			<th width="39%" nowrap>Product&nbsp;&nbsp;<a href="JavaScript:;" onClick="javascript:document.getElementById('FindProducts').style.display=''">Add New</a>
		    <% end if %></th>
			<th width="30%" colspan="2" nowrap>VAT Category</th>
		</tr>
		<% If pcDontshow=1 then '// No Products found %>
		<tr>
			<td colspan="6">
				<div class="pcCPmessage">No products have been added.</div>
           	</td>
		</tr>                      
		<% 
		Else 
		'// Products found		
			
			  rsVAT.MoveFirst
			  
			  ' Get the max number of pages
			  iPageCount=rsVAT.PageCount
			  If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
			  If iPageCurrent < 1 Then iPageCurrent=1
										  
			  ' Set the absolute page
			  rsVAT.AbsolutePage=iPageCurrent
			  hCnt=0
			  Count=0
			  Do While NOT rsVAT.EOF
				  Count=Count + 1
				  ' Assign values							
				  pIdProduct=rsVAT("idproduct")
				  pSku=rsVAT("sku")
				  pDescription=rsVAT("description")
				  pVATCategory=rsVAT("pcVATRate_Category")
				  pVATRate=rsVAT("pcVATRate_Rate")
				  %>
					<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
					  <td>
						  <p><input type="checkbox" name="C<%=Count%>" value="1" checked class="clearBorder"></p>
						  <input type="hidden" name="idProduct<%=Count%>" value="<%=pIdProduct%>">
					  </td>
					  <td><p><%=pSku%></p></td>
					  <td colspan="2">
						  <p><a href="FindProductType.asp?id=<%=pIdProduct%>" target="_blank"><%=pDescription%></a></p>
					  </td>
					  <td><p><%=pVATCategory%>&nbsp;&nbsp;<%=pVATRate%>%</p></td>
					  <td align="center">
						  <a href="javascript:if (confirm('You are about to remove this item from the VAT Category. Are you sure you want to complete this action?')) location='manageVATCategories.asp?action=remove&idproduct=<%=pIdProduct%>&VATID=<%=pcv_intVATID%>'"><img src="images/pcIconDelete.jpg"></a>							
					  
					  </td>
				  </tr>
			  <% 
			  rsVAT.MoveNext
			  Loop
			  End If
			  %>
		<tr>
			<td colspan="6" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="6" class="cpLinksList">
				<a href="javascript:checkAll();">Check All</a>
				&nbsp;|&nbsp;
				<a href="javascript:uncheckAll();">Uncheck All</a>
				<input type="hidden" name="Count" value="<%=Count%>">
				<input type="hidden" name="VATID" value="<%=pcv_intVATID%>">
			</td>
		</tr>
		<tr>
			<td colspan="6"><hr></td>
		</tr>
		<tr>
			<td colspan="6">
				<input name="submit" type=submit class="submit2" value="Remove unchecked">		
				<% if ptaxVATRate_Code<>"" then %>
				<input type="button" name="ManageVATCategories" value="Manage VAT Categories" onclick="location='viewVAT.asp';" class="ibtnGrey">&nbsp;
				<% end if %>
				<input type="button" name="Button" value="Manage VAT Settings" onClick="location='AdminTaxSettings_VAT.asp';" class="ibtnGrey">&nbsp;
				<input type="button" name="Button" value="Back" onClick="JavaScript:history.back()" class="ibtnGrey">
			</td>
		</tr> 
	</table>
</form>
<script language="JavaScript">
<!--
function checkAll() {
for (var j = 1; j <= <%=count%>; j++) {
box = eval("document.form1.C" + j); 
if (box.checked == false) box.checked = true;
   }
}

function uncheckAll() {
for (var j = 1; j <= <%=count%>; j++) {
box = eval("document.form1.C" + j); 
if (box.checked == true) box.checked = false;
   }
}
//-->
</script> 
<% call closeDb() %> 
<!--#include file="AdminFooter.asp"-->
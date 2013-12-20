<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'// If category ID doesn't exist, get the first category that the product has been assigned to, filtering out hidden categories
%>
<!--#include file="pcSeoFirstCat.asp"-->
<%
	'// Call SEO Routine
	pcGenerateSeoLinks
	'//

	'// If product is "Not for Sale", should prices be hidden or shown?
	'// Set pcHidePricesIfNFS = 1 to hide, 0 to show.
	'// Here we leverage the "pcv_intHideBTOPrice" variable to change the behavior (a Control Panel setting could be added in the future)
	pcHidePricesIfNFS = 0
	if (pFormQuantity="-1" and NotForSaleOverride(session("customerCategory"))=0) and pcHidePricesIfNFS=1 then
		pcv_intHideBTOPrice=1
	end if
%>
	<tr class="pcShowProductsL" onmouseover="this.className='pcShowProductsLhover'" onmouseout="this.className='pcShowProductsL'">
		<%if pShowSmallImg <> 0 then%>
		<td width="5%" class="pcShowProductsLCell">
			<%if pSmallImageUrl<>"" then%>
				<a href="<%=pcStrPrdLink%>" <%if scStoreUseToolTip="1" or scStoreUseToolTip="2" then%>onmouseover="javascript:document.getPrd.idproduct.value='<%=pIdProduct%>'; sav_callxml='1'; return runXML1('prd_<%=pIdProduct%>');" onmouseout="javascript: sav_callxml=''; hidetip();"<%end if%>><img src="catalog/<%response.write pSmallImageUrl%>" alt="<%=pDescription%>" class="pcShowProductImageL"></a>
			<% else %>
			&nbsp;
			<%end if %>
		</td>
		<% end if %>
		<td width="65%">
			<div class="pcShowProductName">
				<a href="<%=pcStrPrdLink%>" <%if scStoreUseToolTip="1" or scStoreUseToolTip="2" then%>onmouseover="javascript:document.getPrd.idproduct.value='<%=pIdProduct%>'; sav_callxml='1'; return runXML1('prd_<%=pIdProduct%>');" onmouseout="javascript: sav_callxml=''; hidetip();"<%end if%>><%=pDescription%></a>
				<!--#include file="inc_addPinterest.asp"-->
			</div>
			<% if not psDesc="" then%>
			<div class="pcShowProductSDesc">
				<%=psDesc%>
		<%
		' PRV41 - Product reviews - Start
		%>
		<!-- #include file="pcShowProductReview.asp" -->
		<%
		' PRV41 - Product reviews - End
		%>

			</div>
			<%end if%>
		 </td>
		<%if pShowSKU <> 0 then%>
		<td width="10%" nowrap>
			<%=pSku%>
		</td>
		<% end if %>
		<td width="15%" nowrap> 
				<%ShowSaleIcon=0
				
				if UCase(scDB)="SQL" then	
				if pnoprices=0 then
				query="SELECT pcSales_Completed.pcSC_ID,pcSales_Completed.pcSC_SaveName,pcSales_Completed.pcSC_SaveIcon,pcSales_BackUp.pcSales_TargetPrice FROM (pcSales_Completed INNER JOIN Products ON pcSales_Completed.pcSC_ID=Products.pcSC_ID) INNER JOIN pcSales_BackUp ON pcSales_BackUp.pcSC_ID=pcSales_Completed.pcSC_ID WHERE Products.idproduct=" & pidproduct & " AND Products.pcSC_ID>0;"
				set rsS=Server.CreateObject("ADODB.Recordset")
				set rsS=conntemp.execute(query)
				
				if not rsS.eof then
					ShowSaleIcon=1
					pcSCID=rsS("pcSC_ID")
					pcSCName=rsS("pcSC_SaveName")
					pcSCIcon=rsS("pcSC_SaveIcon")
					pcTargetPrice=rsS("pcSales_TargetPrice") %>
					<script language="JavaScript">
						<!--
									function winSale(fileName)
									{
									myFloater=window.open('','myWindow','scrollbars=auto,status=no,width=450,height=300')
									myFloater.location.href=fileName;
									}
						//-->
					</script>
				<%end if
				set rsS=nothing
				end if
				end if%>
				<% if (pPrice>0) and (pcv_intHideBTOPrice<>"1") then %>
					<div class="pcShowProductPrice">
					<%response.write scCursign & money(pPrice)%>
					<%if (ShowSaleIcon=1) AND (session("customerCategory")=0) AND (pcTargetPrice="0") then%>
					<span class="pcSaleIcon"><a href="javascript:winSale('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="catalog/<%=pcSCIcon%>" title="<%=pcSCName%>" alt="<%=pcSCName%>"></a></span>
					<%end if%>
					<!-- Load quantity discount icon -->
					<!--#Include File="pcShowQtyDiscIcon.asp" -->
					</div>
					<%if (pListPrice-pPrice)>0 AND plistHidden<0 AND session("customerType")<>1 then %>
						<div class="pcShowProductSavings">
						<% response.write dictLanguage.Item(Session("language")&"_prdD1_2") & scCursign & money(pListPrice-pPrice) & " (" & round(((pListPrice-pPrice)/pListPrice)*100) & "%)"%>
						</div>
				<% end if
				end if %>
				<% 'if customer category type logged in - show pricing
				if session("customerCategory")<>0 and (dblpcCC_Price>"0") and (pcv_intHideBTOPrice<>"1") then %>
					<div class="pcShowProductPriceW">
						<% response.write session("customerCategoryDesc")& " " & scCursign & money(dblpcCC_Price)%>
						<%if (ShowSaleIcon=1) AND (clng(session("customerCategory"))=clng(pcTargetPrice)) then%>
						<span class="pcSaleIcon"><a href="javascript:winSale('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="catalog/<%=pcSCIcon%>" title="<%=pcSCName%>" alt="<%=pcSCName%>"></a></span>
						<%end if%>
					</div>
				<% else %>
					<% if (dblpcCC_Price>"0") and (session("customerType")="1") and (pcv_intHideBTOPrice<>"1") then %>
						<div class="pcShowProductPriceW">
							<% response.write dictLanguage.Item(Session("language")&"_prdD1_4")& " " & scCursign & money(dblpcCC_Price)%>
							<%if (ShowSaleIcon=1) AND (clng(pcTargetPrice)=-1) then%>
							<span class="pcSaleIcon"><a href="javascript:winSale('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="catalog/<%=pcSCIcon%>" title="<%=pcSCName%>" alt="<%=pcSCName%>"></a></span>
							<%end if%>
						</div>
					<% end if 
				end if %>
		</td>
	</tr><%

	' Detailed Product Reviews - START
	pcv_IDProduct = pIDProduct
	if pcStrPageNameOR="showRecentlyReviewed.asp" then
	%>
    	<tr>
        	<td colspan="4" align="left">
			<!--#include file="prv_increviews.asp"-->
            </td>
        </tr>
    <%
	End if
	' Detailed Product Reviews - END
	%>
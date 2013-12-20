<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'// START - Show Recently Viewed Products

dim MaxProducts, ViewedPrdList, tmpIndex, tmpIndex1, tmpIndex2, tmpViewedList, tmpVPrdArr, tmpVPrdCount,connTemp2, pcvStrSPname

'// Set maximum products to show
MaxProducts=10

ViewedPrdList=getUserInput2(Request.Cookies("pcfront_visitedPrds"),0)

IF ViewedPrdList<>"" AND ViewedPrdList<>"*" THEN
	
	tmpViewedList=split(ViewedPrdList,"*")
	ViewedPrdList=""
	tmpIndex=0
	tmpIndex1=0
	pcv_ValidateList=0
	pcv_ValidateFailAll=0
	Do While (tmpIndex<ubound(tmpViewedList)) AND (tmpIndex1+1<=MaxProducts)		
		pcv_EvalViewedPrd = tmpViewedList(tmpIndex)		
		if pcv_EvalViewedPrd="" OR validNum(pcv_EvalViewedPrd) then
			pcv_ValidateList=1
		else
			pcv_ValidateFailAll=1
		end if
		if tmpViewedList(tmpIndex)<>"" then
			if ViewedPrdList<>"" then
				ViewedPrdList=ViewedPrdList & ","
			end if
			ViewedPrdList=ViewedPrdList & tmpViewedList(tmpIndex)
			tmpIndex1=tmpIndex1+1
		end if
		tmpIndex=tmpIndex+1
	Loop
	
	tmpViewedList=split(ViewedPrdList,",")
	
	IF pcv_ValidateList=1 AND pcv_ValidateFailAll=0 AND len(ViewedPrdList)>0 THEN '// The cookie was NOT modified or corrupted
	
		Set connTemp2=Server.CreateObject("ADODB.Connection")
		connTemp2.Open scDSN
		query="SELECT products.idproduct,products.description,products.sku,products.smallImageUrl FROM Products WHERE idproduct IN (" & ViewedPrdList & ");"
		set rs=connTemp2.execute(query)
		IF err.number<>0 THEN
			set rs = nothing
			set connTemp2=nothing
		ELSE
		
			IF NOT rs.eof THEN
			%>
				<div id="recentprds">
					<h3><%=dictLanguage.Item(Session("language")&"_viewProducts_1")%></h3>
					<%
					tmpVPrdArr=rs.getRows()
					set rs=nothing
					tmpVPrdCount=ubound(tmpVPrdArr,2)
					For tmpIndex2=0 to tmpIndex1-1
						For tmpIndex=0 to tmpVPrdCount
							if CLng(tmpVPrdArr(0,tmpIndex))=CLng(tmpViewedList(tmpIndex2)) then
							
								' Get product image and sku
								pcvStrSku = tmpVPrdArr(2,tmpIndex)
								pcvStrSmallImage = tmpVPrdArr(3,tmpIndex)
								if pcvStrSmallImage = "" or pcvStrSmallImage = "no_image.gif" then
									pcvStrSmallImage = "hide"
								end if
								
								'Image size
								pcIntSmImgWidth = 20
								pcIntSmImgHeight = 20
								' End get product image
								
								'Show SKU?
								pcIntShowSKU = 0
								
								'Clean up product name
								pcvStrSPname=ClearHTMLTags2(tmpVPrdArr(1,tmpIndex),0)
								if len(pcvStrSPname)>22 then
									pcvStrSPname=left(pcvStrSPname,19) & "..."
								end if
								
							%>
							<div style="clear:both; padding: 0 5px 5px 5px;">
								<% if pcvStrSmallImage <> "hide" then %><a href="viewPrd.asp?idproduct=<%=tmpVPrdArr(0,tmpIndex)%>"><img src="catalog/<%=pcvStrSmallImage%>" width="<%=pcIntSmImgWidth%>" height="<%=pcIntSmImgHeight%>" align="middle" style="border:none; padding-top: 2px; padding-bottom: 5px; padding-right: 4px;"></a><% end if %><a href="viewPrd.asp?idproduct=<%=tmpVPrdArr(0,tmpIndex)%>"><%=pcvStrSPname%></a><% if pcIntShowSKU=1 then%><div class="pcSmallText"><%=pcvStrSku%></div><%end if%>
							</div>
								<%exit for
							end if
						Next
					Next%>
					<div style="clear:both; text-align: right; margin: 5px;">
					<a href="javascript:ClearViewedPrdList();"><%=dictLanguage.Item(Session("language")&"_viewProducts_2")%></a>
					</div>
					<iframe id="clearViewedPrdListCookie" src="clearViewedPrdsList.asp" frameborder="0" width="0" height="0"></iframe>
					<script>
						function ClearViewedPrdList()
						{
							document.getElementById('recentprds').style.display='none';
							document.getElementById('clearViewedPrdListCookie').src="clearViewedPrdsList.asp?action=clear";
						}
					</script>
			</div>
			<%
			END IF ' Empty recordset
		END IF ' Any errors
		set rs=nothing
		set connTemp2=nothing
	
	END IF ' Valid cookie
	
	ViewedPrdList="*" & replace(ViewedPrdList,",","*") & "*"
	
	Response.Cookies("pcfront_visitedPrds")=ViewedPrdList
	Response.Cookies("pcfront_visitedPrds").Expires=Date() + 365

END IF ' Product list exists

'// END - Show Recently Viewed Products
%>
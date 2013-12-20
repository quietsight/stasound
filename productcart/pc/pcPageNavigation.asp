<% '// To hide the "View All" link, set the variable "pcv_ShowViewAllLink" below to "0"
pcv_ShowViewAllLink=1

iRecSize=10

'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'*******************************
' START Page Navigation
'*******************************

If iPageCount>1 then %>
	<div class="pcPageNav">
	<%response.write(dictLanguage.Item(Session("language")&"_advSrcb_4") & iPageCurrent & dictLanguage.Item(Session("language")&"_advSrcb_5") & iPageCount)%>
	&nbsp;-&nbsp;
    <% if iPageCount>iRecSize then %>
		<% if cint(iPageCurrent)>iRecSize then %>
            <a href="<%=pcStrPageName%>?incSale=<%=incSale%>&IDSale=<%=tmpIDSale%>&VA=0&ProdSort=<%=ProdSort%>&iPageCurrent=1&iPageSize=<%=iPageSize%>&PageStyle=<%=pcPageStyle%>&customfield=<%=pcustomfield%>&SearchValues=<%=pCValues%>&exact=<%=intExact%>&keyword=<%=replace(tKeywords,"""","%22")%>&priceFrom=<%=pPriceFrom%>&priceUntil=<%=pPriceUntil%>&idCategory=<%=pIdCategory%>&IdSupplier=<%=pIdSupplier%>&withStock=<%=pWithStock%>&IDBrand=<%=IDBrand%>&order=<%=strORD%>&SKU=<%=pSearchSKU%><%=pcv_strCSFieldQuery%>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_1")%></a>&nbsp;
        <% end if %>
		<% if cint(iPageCurrent)>1 then
            if cint(iPageCurrent)<iRecSize AND cint(iPageCurrent)<iRecSize then
                iPagePrev=cint(iPageCurrent)-1
            else
                iPagePrev=iRecSize
            end if %>
            <a href="<%=pcStrPageName%>?incSale=<%=incSale%>&IDSale=<%=tmpIDSale%>&VA=0&ProdSort=<%=ProdSort%>&iPageCurrent=<%=cint(iPageCurrent)-iPagePrev%>&iPageSize=<%=iPageSize%>&PageStyle=<%=pcPageStyle%>&customfield=<%=pcustomfield%>&SearchValues=<%=pCValues%>&exact=<%=intExact%>&keyword=<%=replace(tKeywords,"""","%22")%>&priceFrom=<%=pPriceFrom%>&priceUntil=<%=pPriceUntil%>&idCategory=<%=pIdCategory%>&IdSupplier=<%=pIdSupplier%>&withStock=<%=pWithStock%>&IDBrand=<%=IDBrand%>&order=<%=strORD%>&SKU=<%=pSearchSKU%><%=pcv_strCSFieldQuery%>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_2")%>&nbsp;<%=iPagePrev%>&nbsp;<%=dictLanguage.Item(Session("language")&"_PageNavigaion_3")%></a>
		<% end if
		if cint(iPageCurrent)+1>1 then
			intPageNumber=cint(iPageCurrent)
		else
			intPageNumber=1
		end if
	else
		intPageNumber=1
	end if
	
	if (cint(iPageCount)-cint(iPageCurrent))<iRecSize then
		iPageNext=cint(iPageCount)-cint(iPageCurrent)
	else
		iPageNext=iRecSize
	end if

	For pageNumber=intPageNumber To (cint(iPageCurrent) + (iPageNext))
		If Cint(pageNumber)=Cint(iPageCurrent) Then %>
			<strong><%=pageNumber%></strong> 
		<% Else %>
            <a href="<%=pcStrPageName%>?incSale=<%=incSale%>&IDSale=<%=tmpIDSale%>&VA=0&ProdSort=<%=ProdSort%>&iPageCurrent=<%=pageNumber%>&iPageSize=<%=iPageSize%>&PageStyle=<%=pcPageStyle%>&customfield=<%=pcustomfield%>&SearchValues=<%=pCValues%>&exact=<%=intExact%>&keyword=<%=replace(tKeywords,"""","%22")%>&priceFrom=<%=pPriceFrom%>&priceUntil=<%=pPriceUntil%>&idCategory=<%=pIdCategory%>&IdSupplier=<%=pIdSupplier%>&withStock=<%=pWithStock%>&IDBrand=<%=IDBrand%>&order=<%=strORD%>&SKU=<%=pSearchSKU%><%=pcv_strCSFieldQuery%>"><%=pageNumber%></a>
		<% End If 
	Next
	
	if (cint(iPageNext)+cint(iPageCurrent))=iPageCount then
	else
		if iPageCount>(cint(iPageCurrent) + (iRecSize-1)) then %>
			<a href="<%=pcStrPageName%>?incSale=<%=incSale%>&IDSale=<%=tmpIDSale%>&VA=0&ProdSort=<%=ProdSort%>&iPageCurrent=<%=cint(intPageNumber)+iPageNext%>&iPageSize=<%=iPageSize%>&PageStyle=<%=pcPageStyle%>&customfield=<%=pcustomfield%>&SearchValues=<%=pCValues%>&exact=<%=intExact%>&keyword=<%=replace(tKeywords,"""","%22")%>&priceFrom=<%=pPriceFrom%>&priceUntil=<%=pPriceUntil%>&idCategory=<%=pIdCategory%>&IdSupplier=<%=pIdSupplier%>&withStock=<%=pWithStock%>&IDBrand=<%=IDBrand%>&order=<%=strORD%>&SKU=<%=pSearchSKU%><%=pcv_strCSFieldQuery%>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_4")%>&nbsp;<%=iPageNext%>&nbsp;<%=dictLanguage.Item(Session("language")&"_PageNavigaion_3")%></a>
		<% end if
    
		if cint(iPageCount)>iRecSize AND (cint(iPageCurrent)<>cint(iPageCount)) then %>
    		&nbsp;<a href="<%=pcStrPageName%>?incSale=<%=incSale%>&IDSale=<%=tmpIDSale%>&VA=0&ProdSort=<%=ProdSort%>&iPageCurrent=<%=cint(iPageCount)%>&iPageSize=<%=iPageSize%>&PageStyle=<%=pcPageStyle%>&customfield=<%=pcustomfield%>&SearchValues=<%=pCValues%>&exact=<%=intExact%>&keyword=<%=replace(tKeywords,"""","%22")%>&priceFrom=<%=pPriceFrom%>&priceUntil=<%=pPriceUntil%>&idCategory=<%=pIdCategory%>&IdSupplier=<%=pIdSupplier%>&withStock=<%=pWithStock%>&IDBrand=<%=IDBrand%>&<%=pcv_strCSFieldQuery%>&order=<%=strORD%>&SKU=<%=pSearchSKU%>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_5")%></a>
    	<% end if 
	end if 
	
	if pcv_ShowViewAllLink=1 then %>
    &nbsp;&nbsp;<a href="<%=pcStrPageName%>?incSale=<%=incSale%>&IDSale=<%=tmpIDSale%>&VA=1&ProdSort=<%=ProdSort%>&PageStyle=<%=pcPageStyle%>&customfield=<%=pcustomfield%>&SearchValues=<%=pCValues%>&exact=<%=intExact%>&keyword=<%=replace(tKeywords,"""","%22")%>&priceFrom=<%=pPriceFrom%>&priceUntil=<%=pPriceUntil%>&idCategory=<%=pIdCategory%>&IdSupplier=<%=pIdSupplier%>&withStock=<%=pWithStock%>&IDBrand=<%=IDBrand%>&order=<%=strORD%>&SKU=<%=pSearchSKU%><%=pcv_strCSFieldQuery%>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_6")%></a>
    <% end if %></div>
<% end if

'*******************************
' END Page Navigation
'*******************************
%>
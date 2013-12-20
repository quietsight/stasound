<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'// START - Show Recently Viewed Products

dim MaxProducts, ViewedPrdList, tmpIndex, tmpIndex1, tmpIndex2, tmpViewedList, tmpVPrdArr, tmpVPrdCount,connTemp2, pcv_ShowSmallRecentProducts, pcvStrSPname

pcv_ShowSmallRecentProducts=0
'//Set on error resume next so that no SQL errors are shown in case of problems in the query
'on error resume next

'// Set maximum products to show
MaxProducts=10

ViewedPrdList=getUserInput2(Request.Cookies("pcfront_visitedPrdsCP"),0)

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
	
	ViewedPrdList="*" & replace(ViewedPrdList,",","*") & "*"
	
	Response.Cookies("pcfront_visitedPrdsCP")=ViewedPrdList
	Response.Cookies("pcfront_visitedPrdsCP").Expires=Date() + 365

END IF ' Product list exists

'// END - Show Recently Viewed Products
%>
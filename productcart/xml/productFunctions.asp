<%

function getUserInputNew(input,stringLength)
 dim tempStr,i,known_bad

 known_bad= array("*","--")
 if stringLength>0 then
  tempStr	= left(trim(input),stringLength) 
 else
  tempStr	= trim(input)
 end if
 for i=lbound(known_bad) to ubound(known_bad)
 	if (instr(1,tempStr,known_bad(i),vbTextCompare)<>0) then
 		tempStr	= replace(tempStr,known_bad(i),"")
 	end if
 next
 tempStr	= replace(tempStr,"'","''")
 '// tempStr	= replace(tempStr,"<","&lt;")
 '// tempStr	= replace(tempStr,">","&gt;")
 tempStr	= replace(tempStr,"%0d","")
 tempStr	= replace(tempStr,"%0D","")
 tempStr	= replace(tempStr,"%0a","")
 tempStr	= replace(tempStr,"%0A","")
 tempStr	= replace(tempStr,"\r\n","")
 tempStr	= replace(tempStr,"\r","")
 tempStr	= replace(tempStr,"\n","")
 tempStr	= replace(tempStr,"\R\N","")
 tempStr	= replace(tempStr,"\R","")
 tempStr	= replace(tempStr,"\N","")
 tempStr	= replace(tempStr,"EXEC(","",1,-1,1) 
	
	if tempStr<>"" then
	 	if IsNumeric(tempStr) then
	 		if InStr(Cstr(10/3),",")>0 then
				if Instr(tempStr,".")>0 then
					tempStr=FormatNumber(tempStr,,,,0)
	 				tempStr=replace(tempStr,".",",")
				end if
	 		end if
	 	end if
	end if
 
 getUserInputNew	= tempStr 
end function

Sub CheckSrcProductsTags()
Dim ChildNodes,strNode,tmpNodeName,tmpNodeValue,tmpValue1
	Set fNode=iRoot.selectSingleNode(cm_filters_name)
	if fNode is Nothing then
		exit Sub
	end if
	if fNode.Text="" then
		exit Sub
	end if
	Set ChildNodes = fNode.childNodes
	
	For Each strNode In ChildNodes
		tmpNodeName=strNode.nodeName
		tmpNodeValue=trim(strNode.Text)
		
		Select Case tmpNodeName
			Case srcCategoryID_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcCategoryID_ex=1
					srcCategoryID_value=tmpNodeValue
				end if
			Case srcCFieldID_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcCFieldID_ex=1
					srcCFieldID_value=tmpNodeValue
				end if
			Case srcCFieldValue_name:
				call CheckValidXMLTag(strNode,0,5,"")
				if tmpNodeValue<>"" then
					srcCFieldValue_ex=1
					srcCFieldValue_value=getUserInputNew(tmpNodeValue,0)
				end if
			Case srcPriceFrom_name:
				call CheckValidXMLTag(strNode,0,3,"")
				if tmpNodeValue<>"" then
					srcPriceFrom_ex=1
					srcPriceFrom_value=tmpNodeValue
				end if
			Case srcPriceTo_name:
				tmpValue1=0
				if CheckExistTag(cm_filters_name & "/" & srcPriceFrom_name) then
					tmpValue1=iRoot.selectSingleNode(cm_filters_name & "/" & srcPriceFrom_name).Text
				end if
				call CheckValidXMLTag(strNode,0,3,tmpValue1)
				if tmpNodeValue<>"" then
					srcPriceTo_ex=1
					srcPriceTo_value=tmpNodeValue
				end if
			Case srcInStock_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcInStock_ex=1
					srcInStock_value=tmpNodeValue
					if srcInStock_value>1 then
						srcInStock_value=1
					end if
				end if
			Case srcSKU_name:
				call CheckValidXMLTag(strNode,0,5,"")
				if tmpNodeValue<>"" then
					srcSKU_ex=1
					srcSKU_value=getUserInputNew(tmpNodeValue,0)
				end if
			Case srcBrandID_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcBrandID_ex=1
					srcBrandID_value=tmpNodeValue
				end if
			Case srcKeyword_name:
				call CheckValidXMLTag(strNode,0,5,"")
				if tmpNodeValue<>"" then
					srcKeyword_ex=1
					srcKeyword_value=getUserInputNew(tmpNodeValue,0)
				end if
			Case srcExactPhrase_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcExactPhrase_ex=1
					srcExactPhrase_value=tmpNodeValue
					if srcExactPhrase_value>1 then
						srcExactPhrase_value=1
					end if
				end if
			Case srcIncInactive_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcIncInactive_ex=1
					srcIncInactive_value=tmpNodeValue
					if srcIncInactive_value>1 then
						srcIncInactive_value=1
					end if
				end if
			Case srcIncDeleted_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcIncDeleted_ex=1
					srcIncDeleted_value=tmpNodeValue
					if srcIncDeleted_value>1 then
						srcIncDeleted_value=1
					end if
				end if
			Case srcIncNormal_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcIncNormal_ex=1
					srcIncNormal_value=tmpNodeValue
					if srcIncNormal_value>1 then
						srcIncNormal_value=1
					end if
				end if
			Case srcIncBTO_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcIncBTO_ex=1
					srcIncBTO_value=tmpNodeValue
					if srcIncBTO_value>1 then
						srcIncBTO_value=1
					end if
				end if
			Case srcIncBTOItems_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcIncBTOItems_ex=1
					srcIncBTOItems_value=tmpNodeValue
					if srcIncBTOItems_value>1 then
						srcIncBTOItems_value=1
					end if
				end if
			Case srcSpecial_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcSpecial_ex=1
					srcSpecial_value=tmpNodeValue
					if srcSpecial_value>1 then
						srcSpecial_value=1
					end if
				end if
			Case srcFeatured_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcFeatured_ex=1
					srcFeatured_value=tmpNodeValue
					if srcFeatured_value>1 then
						srcFeatured_value=1
					end if
				end if
			Case srcSort_name:
				call CheckValidXMLTag(strNode,1,1,"")
				if tmpNodeValue<>"" then
					srcSort_ex=1
					srcSort_value=tmpNodeValue
				end if
			Case srcFromDate_name:
				call CheckValidXMLTag(strNode,0,4,"")
				if tmpNodeValue<>"" then
					srcFromDate_ex=1
					srcFromDate_value=ConvertFromXMLDate(tmpNodeValue)
				end if
			Case srcToDate_name:
				tmpValue1=0
				if CheckExistTag(cm_filters_name & "/" & srcFromDate_name) then
					tmpValue1=iRoot.selectSingleNode(cm_filters_name & "/" & srcFromDate_name).Text
				end if
				call CheckValidXMLTag(strNode,0,4,tmpValue1)
				if tmpNodeValue<>"" then
					srcToDate_ex=1
					srcToDate_value=ConvertFromXMLDate(tmpNodeValue)
				end if
			Case srcHideExported_name:
				if cm_ExportAdmin="1" then
					call CheckValidXMLTag(strNode,1,1,"")
					if tmpNodeValue<>"" then
						srcHideExported_ex=1
						srcHideExported_value=tmpNodeValue
						if srcHideExported_value>1 then
							srcHideExported_value=1
						end if
					end if
				end if
			Case Else:
				call XMLcreateError(105,cm_errorStr_105 & tmpNodeName)
				call returnXML()
		End Select
	Next
End Sub

Sub CheckNewProductsTags()
Dim ChildNodes,strNode,tmpNodeName,tmpNodeValue,tmpValue1
	Set fNode=iRoot.selectSingleNode(cm_filters_name)
	if fNode is Nothing then
		exit Sub
	end if
	if fNode.Text="" then
		exit Sub
	end if
	Set ChildNodes = fNode.childNodes
	
	For Each strNode In ChildNodes
		tmpNodeName=strNode.nodeName
		tmpNodeValue=trim(strNode.Text)
		
		Select Case tmpNodeName
			Case srcFromDate_name:
				call CheckValidXMLTag(strNode,0,4,"")
				if tmpNodeValue<>"" then
					srcFromDate_ex=1
					srcFromDate_value=ConvertFromXMLDate(tmpNodeValue)
				end if
			Case Else:
				call XMLcreateError(105,cm_errorStr_105 & tmpNodeName)
				call returnXML()
		End Select
	Next
End Sub

Sub CheckGetProductDetailsTags()
	
	Dim ChildNodes,strNode,tmpNodeName,tmpNodeValue,tmpValue1
	
	Call CheckRequiredXMLTag(prdID_name)
	Set strNode=iRoot.selectSingleNode(prdID_name)
	call CheckValidXMLTag(strNode,1,1,"")
	prdID_ex=1
	prdID_value=tmpNode.Text

	Set rNode=iRoot.selectSingleNode(cm_requests_name)
	if rNode is Nothing then
		Call SetDefaultProductDetailsTags()
		exit Sub
	else
		if rNode.Text="" then
			Call SetDefaultProductDetailsTags()
			exit Sub
		end if
	end if
	Set ChildNodes = rNode.childNodes
	
	For Each strNode In ChildNodes
		tmpNodeName=strNode.nodeName
		tmpNodeValue=trim(strNode.Text)
		if	tmpNodeName=cm_request_name then
			Select Case tmpNodeValue
				Case cm_requestDefault_name:
					Call SetDefaultProductDetailsTags()
				Case cm_requestAll_name:
					Call SetAllProductDetailsTags()
				Case prdSKU_name:
					prdSKU_ex=1
				Case prdName_name:
					prdName_ex=1
				Case prdDesc_name:
					prdDesc_ex=1
				Case prdSDesc_name:
					prdSDesc_ex=1
				Case prdType_name:
					prdType_ex=1
				Case prdPrice_name:
					prdPrice_ex=1
				Case prdListPrice_name:
					prdListPrice_ex=1
				Case prdWPrice_name:
					prdWPrice_ex=1
				Case prdWeight_name:
					prdWeight_ex=1
				Case prdStock_name:
					prdStock_ex=1
				Case prdCategory_name:
					prdCategory_ex=1
				Case prdBrand_name:
					prdBrand_ex=1
				Case prdSmallImg_name:
					prdSmallImg_ex=1
				Case prdImg_name:
					prdImg_ex=1
				Case prdLargeImg_name:
					prdLargeImg_ex=1
				Case prdActive_name:
					prdActive_ex=1
				Case prdShowSavings_name:
					prdShowSavings_ex=1
				Case prdSpecial_name:
					prdSpecial_ex=1
				Case prdFeatured_name:
					prdFeatured_ex=1
				Case prdOptionGroup_name:
					prdOptionGroup_ex=1
				Case prdRewardPoints_name:
					prdRewardPoints_ex=1
				Case prdNoTax_name:
					prdNoTax_ex=1
				Case prdNoShippingCharge_name:
					prdNoShippingCharge_ex=1
				Case prdNotForSale_name:
					prdNotForSale_ex=1
				Case prdNotForSaleMsg_name:
					prdNotForSaleMsg_ex=1
				Case prdDisregardStock_name:
					prdDisregardStock_ex=1
				Case prdDisplayNoShipText_name:
					prdDisplayNoShipText_ex=1
				Case prdMinimumQty_name:
					prdMinimumQty_ex=1
				Case prdPurchaseMulti_name:
					prdPurchaseMulti_ex=1
				Case prdOversize_name:
					prdOversize_ex=1
				Case prdCost_name:
					prdCost_ex=1
				Case prdBackOrder_name:
					prdBackOrder_ex=1
				Case prdShipNDays_name:
					prdShipNDays_ex=1
				Case prdLowStockNotice_name:
					prdLowStockNotice_ex=1
				Case prdReorderLevel_name:
					prdReorderLevel_ex=1
				Case prdIsDropShipped_name:
					prdIsDropShipped_ex=1
				Case prdSupplierID_name:
					prdSupplierID_ex=1
				Case prdDropShipperID_name:
					prdDropShipperID_ex=1
				Case prdMetaTags_name:
					prdMetaTags_ex=1
				Case prdDownloadable_name:
					prdDownloadable_ex=1
				Case prdDownloadInfo_name:
					prdDownloadInfo_ex=1
				Case prdGiftCertificate_name:
					prdGiftCertificate_ex=1
				Case prdGCInfo_name:
					prdGCInfo_ex=1
				Case prdHideBTOPrices_name:
					prdHideBTOPrices_ex=1
				Case prdHideDefaultConfig_name:
					prdHideDefaultConfig_ex=1
				Case prdDisallowPurchasing_name:
					prdDisallowPurchasing_ex=1
				Case prdSkipPrdPage_name:
					prdSkipPrdPage_ex=1
				Case prdCustomField_name:
					prdCustomField_ex=1
				Case prdCreatedDate_name:
					prdCreatedDate_ex=1
				Case Else:
					call XMLcreateError(106,cm_errorStr_106 & tmpNodeValue)
					call returnXML()
			End Select
		else
			call XMLcreateError(106,cm_errorStr_106 & tmpNodeName)
			call returnXML()
		end if
	Next
End Sub

Sub SetDefaultProductDetailsTags()
	prdSKU_ex=1
	prdName_ex=1
	prdDesc_ex=1
	prdSDesc_ex=1
	prdType_ex=1
	prdPrice_ex=1
	prdListPrice_ex=1
	prdWPrice_ex=1
	prdWeight_ex=1
	prdStock_ex=1
	prdCategory_ex=0
	prdBrand_ex=0
	prdSmallImg_ex=1
	prdImg_ex=1
	prdLargeImg_ex=1
	prdActive_ex=1
	prdShowSavings_ex=1
	prdSpecial_ex=1
	prdFeatured_ex=1
	prdOptionGroup_ex=0
	prdRewardPoints_ex=0
	prdNoTax_ex=1
	prdNoShippingCharge_ex=1
	prdNotForSale_ex=1
	prdNotForSaleMsg_ex=0
	prdDisregardStock_ex=1
	prdDisplayNoShipText_ex=1
	prdMinimumQty_ex=1
	prdPurchaseMulti_ex=1
	prdOversize_ex=0
	prdCost_ex=0
	prdBackOrder_ex=1
	prdShipNDays_ex=1
	prdLowStockNotice_ex=1
	prdReorderLevel_ex=1
	prdIsDropShipped_ex=1
	prdSupplierID_ex=1
	prdDropShipperID_ex=1
	prdMetaTags_ex=0
	prdDownloadable_ex=0
	prdDownloadInfo_ex=0
	prdGiftCertificate_ex=0
	prdGCInfo_ex=0
	prdHideBTOPrices_ex=0
	prdHideDefaultConfig_ex=0
	prdDisallowPurchasing_ex=0
	prdSkipPrdPage_ex=0
	prdCustomField_ex=0
	prdCreatedDate_ex=0
End Sub

Sub SetAllProductDetailsTags()
	prdSKU_ex=1
	prdName_ex=1
	prdDesc_ex=1
	prdSDesc_ex=1
	prdType_ex=1
	prdPrice_ex=1
	prdListPrice_ex=1
	prdWPrice_ex=1
	prdWeight_ex=1
	prdStock_ex=1
	prdCategory_ex=1
	prdBrand_ex=1
	prdSmallImg_ex=1
	prdImg_ex=1
	prdLargeImg_ex=1
	prdActive_ex=1
	prdShowSavings_ex=1
	prdSpecial_ex=1
	prdFeatured_ex=1
	prdOptionGroup_ex=1
	prdRewardPoints_ex=1
	prdNoTax_ex=1
	prdNoShippingCharge_ex=1
	prdNotForSale_ex=1
	prdNotForSaleMsg_ex=1
	prdDisregardStock_ex=1
	prdDisplayNoShipText_ex=1
	prdMinimumQty_ex=1
	prdPurchaseMulti_ex=1
	prdOversize_ex=1
	prdCost_ex=1
	prdBackOrder_ex=1
	prdShipNDays_ex=1
	prdLowStockNotice_ex=1
	prdReorderLevel_ex=1
	prdIsDropShipped_ex=1
	prdSupplierID_ex=1
	prdDropShipperID_ex=1
	prdMetaTags_ex=1
	prdDownloadable_ex=1
	prdDownloadInfo_ex=1
	prdGiftCertificate_ex=1
	prdGCInfo_ex=1
	prdHideBTOPrices_ex=1
	prdHideDefaultConfig_ex=1
	prdDisallowPurchasing_ex=1
	prdSkipPrdPage_ex=1
	prdCustomField_ex=1
	prdCreatedDate_ex=1
End Sub

Function GenSrcProductsQuery()
Dim strSQL, tmpSQL, tmpSQL2, query, PrdTypeStr

	pSKU=srcSKU_value
	pKeywords=srcKeyword_value
	pCValues=srcCFieldValue_value
	tIncludeSKU="false"
	pPriceFrom=srcPriceFrom_value
	if trim(pPriceFrom)="" then
		pPriceFrom=0
	end if
	if NOT isNumeric(pPriceFrom) then
		pPriceFrom=0
	end if
	pPriceUntil=srcPriceTo_value
	if trim(pPriceUntil)="" then
		pPriceUntil=9999999
	end if
	if NOT isNumeric(pPriceUntil) then
		pPriceUntil=9999999
	end if
	pIdCategory=srcCategoryID_value
	if NOT isNumeric(pIdCategory) then
		pIdCategory=0
	end if
	pWithStock=srcInStock_value
	
	pcustomfield=srcCFieldID_value
	if pcustomfield = "" then
		pcustomfield = "0"
		pCValues=""
	end if
	if Not IsNumeric(pcustomfield) then
		pcustomfield = 0
		pCValues=""
	end if
	
	if pcustomfield=0 then
		pCValues=""
	end if
	
	IDBrand=srcBrandID_value
	if NOT isNumeric(IDBrand) then
		IDBrand=0
	end if
	strORD=srcSort_value
	if strORD="" then
		strORD=4
	end if
	if NOT isNumeric(strORD) then
		strORD=4
	end if
	pInactive=srcIncInactive_value
	pIncDeleted=srcIncDeleted_value
	
	form_exact=srcExactPhrase_value
	
	src_IncNormal=srcIncNormal_value
	src_IncBTO=srcIncBTO_value
	src_IncItem=srcIncBTOItems_value
	src_Special=srcSpecial_value
	src_Featured=srcFeatured_value
	
	if src_IncNormal="" then
		src_IncNormal="1"
	end if
	
	if src_IncBTO="" then
		src_IncBTO="1"
	end if
	
	if src_IncItem="" then
		src_IncItem="0"
	end if
	
	if src_Special="" then
		src_Special="0"
	end if

	if src_Featured="" then
		src_Featured="0"
	end if
	
	if (src_IncBTO="0") and (src_IncItem="0") then
		src_IncNormal="1"
	end if

	' create sql statement
	
	if strORD<>"" then
		Select Case StrORD
			Case "0": strORD1="products.description ASC"
			Case "1": strORD1="products.description DESC"
			Case "2": strORD1="products.sku ASC"
			Case "3": strORD1="products.sku DESC"
			Case "4": strORD1="products.idproduct ASC"
			Case "5": strORD1="products.idproduct ASC"
			
		End Select
	Else
		strORD="4"
		strORD1="products.idproduct ASC"
	End If
		
	PrdTypeStr=""
	
	if (src_IncBTO="1") and (src_IncItem="0") and (src_IncNormal="0") then
		PrdTypeStr=" AND serviceSpec<>0 "
	end if
	
	if (src_IncBTO="0") and (src_IncItem="1") and (src_IncNormal="0") then
		PrdTypeStr=" AND configOnly<>0 "
	end if
	
	if (src_IncBTO="1") and (src_IncItem="1") and (src_IncNormal="0") then
		PrdTypeStr=" AND ((serviceSpec<>0) OR (configOnly<>0)) "
	end if
	
	if (src_IncBTO="0") and (src_IncItem="1") and (src_IncNormal="1") then
		PrdTypeStr=" AND serviceSpec=0 "
	end if
	
	if (src_IncBTO="1") and (src_IncItem="0") and (src_IncNormal="1") then
		PrdTypeStr=" AND configOnly=0 "
	end if
	
	if (src_IncBTO="0") and (src_IncItem="0") and (src_IncNormal="1") then
		PrdTypeStr=" AND ((serviceSpec=0) AND (configOnly=0)) "
	end if	
	
	if pIdCategory<>"0" then
		tmpSQL1=",categories_products "
	else
		tmpSQL1=""
	end if
	
	strSQL="SELECT DISTINCT products.idProduct FROM products " & tmpSQL1 & session("srcprd_from") & tmp_addFrom & " WHERE products.price>="&pPriceFrom&" And products.price<=" &pPriceUntil
	
	if len(pSKU)>0 then
		strSQL=strSQL & " AND products.sku like '%"&pSKU&"%'"
	end if
	
	if pIdCategory<>"0" then
		strSQL=strSQL & " AND products.idProduct=categories_products.idProduct AND categories_products.idCategory=" &pIdCategory   
	end if
	
	if pWithStock="1" then
		strSQL=strSQL & " AND (stock>0 OR noStock<>0)" 
	end if
	
	if (IDBrand&""<>"") and (IDBrand&""<>"0") then
		strSQL=strSQL & " AND IDBrand=" & IDBrand
	end if
	
	If srcFromDate_ex=1 then
		tmpFromDate=srcFromDate_Value
		if SQL_Format="1" then
			tmpFromDate=Day(tmpFromDate)&"/"&Month(tmpFromDate)&"/"&Year(tmpFromDate)
		else
			tmpFromDate=Month(tmpFromDate)&"/"&Day(tmpFromDate)&"/"&Year(tmpFromDate)
		end if
		if scDB="Access" then
			strSQL=strSQL & " AND products.pcprod_EnteredOn>=#" & tmpFromDate & "# "
		else
			strSQL=strSQL & " AND products.pcprod_EnteredOn>='" & tmpFromDate & "' "
		end if
	End if
	
	If srcToDate_ex=1 then
		tmpToDate=CDate(srcToDate_Value)
		if SQL_Format="1" then
			tmpToDate=Day(tmpToDate)&"/"&Month(tmpToDate)&"/"&Year(tmpToDate)
		else
			tmpToDate=Month(tmpToDate)&"/"&Day(tmpToDate)&"/"&Year(tmpToDate)
		end if
		if scDB="Access" then
			strSQL=strSQL & " AND products.pcprod_EnteredOn<=#" & tmpToDate & "# "
		else
			strSQL=strSQL & " AND products.pcprod_EnteredOn<='" & tmpToDate & "' "
		end if
	End if
	
	if pInactive="1" then
	else
		strSQL=strSQL & " AND products.active<>0" 
	end if
	
	if pIncDeleted="1" then
	else
		strSQL=strSQL & " AND products.removed=0"
	end if
	
	if src_Special="1" then
	   strSQL=strSQL & " AND products.hotdeal<>0" 
	end if
	
	if src_Special="2" then
	   strSQL=strSQL & " AND products.hotdeal=0" 
	end if
	
	if src_Featured="1" then
	   strSQL=strSQL & " AND products.showInHome<>0" 
	end if
	
	if src_Featured="2" then
	   strSQL=strSQL & " AND products.showInHome=0" 
	end if
	
	if cm_ExportAdmin="1" AND srcHideExported_value="1" then
		strSQL=strSQL & " AND (Products.idproduct NOT IN (SELECT DISTINCT pcXEL_ExportedID FROM pcXMLExportLogs WHERE pcXP_ID=" & pcv_PartnerID & " AND pcXEL_IDType=0)) "
	end if
	
	TestWord=""
	if form_exact<>"1" then
		if Instr(pKeywords," AND ")>0 then
			keywordArray=split(pKeywords," AND ")
			TestWord=" AND "
		else
			if Instr(pKeywords," and ")>0 then
				keywordArray=split(pKeywords," and ")
				TestWord=" AND "
			else
				if Instr(pKeywords,",")>0 then
					keywordArray=split(pKeywords,",")
					TestWord=" OR "
				else
					if (Instr(pKeywords," OR ")>0) then
						keywordArray=split(pKeywords," OR ")
						TestWord=" OR "
					else
						if (Instr(pKeywords," or ")>0) then
							keywordArray=split(pKeywords," or ")
							TestWord=" OR "
						else
							if (Instr(pKeywords," ")>0) then
								keywordArray=split(pKeywords," ")
								TestWord=" AND "
							else
								keywordArray=split(pKeywords,"***")	
								TestWord=" OR "
							end if
						end if
					end if
				end if
			end if
		end if
	else
		keywordArray=split(pKeywords,"***")	
		TestWord=" OR "
	end if
	
	if pcustomfield<>"0" AND pCValues<>"" then
			strCnt=0
			strSQL=strSQL & " AND (("
			strSQL=strSQL & "(custom1="&pcustomfield&" AND content1 LIKE '%"&trim(pCValues)&"%') OR (custom2="&pcustomfield&" AND content2 LIKE '%"&trim(pCValues)&"%') OR (custom3="&pcustomfield&" AND content3 LIKE '%"&trim(pCValues)&"%')"
			strSQL=strSQL & ")"
			if pKeywords="" then
			strSQL=strSQL & ")"
			end if
	end if
	
	if pcustomfield<>"0" AND pKeywords<>"" then
		if pCValues="" then
			strCnt=0
			strSQL=strSQL & " AND (("
			for L=0 to UBound(keywordArray)
				strCnt=strCnt+1
				strSQL=strSQL & "((custom1="&pcustomfield&" AND content1 LIKE '%"&trim(keywordArray(L))&"%') OR (custom2="&pcustomfield&" AND content2 LIKE '%"&trim(keywordArray(L))&"%') OR (custom3="&pcustomfield&" AND content3 LIKE '%"&trim(keywordArray(L))&"%'))"
				if strCnt<>(UBound(keywordArray)+1) then
					strSQL=strSQL&TestWord
				end if
			next
			strSQL=strSQL & ")"
		end if
	else
		if pKeywords<>"" AND pCValues="" then
			strCnt=0
			strSQL=strSQL & " AND (("
			for L=0 to UBound(keywordArray)
				strCnt=strCnt+1
				strSQL=strSQL & "((content1 LIKE '%"&trim(keywordArray(L))&"%') OR (content2 LIKE '%"&trim(keywordArray(L))&"%') OR (content3 LIKE '%"&trim(keywordArray(L))&"%'))"
				if strCnt<>(UBound(keywordArray)+1) then
					strSQL=strSQL&TestWord
				end if
			next
	    strSQL=strSQL&")"	
		end if
		if pcustomfield<>"0" AND pKeywords="" AND pCValues="" then
			strCnt=0
			strSQL=strSQL & " AND ("
			strSQL=strSQL & "(custom1="&pcustomfield&") OR (custom2="&pcustomfield&") OR (custom3="&pcustomfield&")"
			strSQL=strSQL & ")"
		end if
	end if
	
	if pKeywords<>"" then
		if pcustomfield<>"0" AND pCValues<>"" then
			strSQl=strSql & " AND ("
		else
			strSQl=strSql & " OR ("
		end if
		
		tmpSQL="(details LIKE "
		tmpSQL2="(description LIKE "
		tmpSQL3="(sDesc LIKE "
		if tIncludeSKU="true" then
			tmpSQL4="(SKU LIKE "
		end if
		Dim Pos
		Pos=0
		For L=LBound(keywordArray) to UBound(keywordArray)
			if trim(keywordArray(L))<>"" then
			Pos=Pos+1
			if Pos>1 Then
				tmpSQL=tmpSQL  & TestWord & " details LIKE "
				tmpSQL2=tmpSQL2 & TestWord & " description LIKE "
				tmpSQL3=tmpSQL3 & TestWord & " sDesc LIKE "
				if tIncludeSKU="true" then
					tmpSQL4=tmpSQL4 & TestWord & " SKU LIKE "
				end if
			end if
				tmpSQL=tmpSQL  & "'%" & trim(keywordArray(L)) & "%'"
				tmpSQL2=tmpSQL2 & "'%" & trim(keywordArray(L)) & "%'"
				tmpSQL3=tmpSQL3 & "'%" & trim(keywordArray(L)) & "%'"
				if tIncludeSKU="true" then
					tmpSQL4=tmpSQL4 & "'%" & trim(keywordArray(L)) & "%'"
				end if
			end if
		Next
		tmpSQL=tmpSQL & ")"
		tmpSQL2=tmpSQL2 & ")"
		tmpSQL3=tmpSQL3 & ")"
		if tIncludeSKU="true" then
			tmpSQL4=tmpSQL4 & ")"
		end if
		
		strSQL=strSQL & tmpSQL
		strSQL=strSQL & " OR " & tmpSQL2
		if tIncludeSKU="true" then
			strSQL=strSQL & " OR " & tmpSQL3
			strSQL=strSQL & " OR " & tmpSQL4 & "))"
		else	
			strSQL=strSQL & " OR " & tmpSQL3 & "))"
		end if
		query=strSQL & PrdTypeStr & " ORDER BY " & strORD1
	else
		query=strSQL & PrdTypeStr & " ORDER BY " & strORD1
	end if
	
	GenSrcProductsQuery=query
	
End Function

Sub RunSrcProducts()
Dim query,rs1,resultCount,pcArr
Dim requestKey,i,strNode
on error resume next
	query=GenSrcProductsQuery()
	call opendb()
	set rs1=connTemp.execute(query)
	resultCount=0
	if Err.number<>0 then
		set rs1=nothing
		call closedb()
		call XMLcreateError(115,cm_errorStr_115)
		call returnXML()
	end if
	if not rs1.eof then
		pcArr=rs1.getRows()
		resultCount=ubound(pcArr,2)+1
	end if
	set rs1=nothing
	call closedb()
	
	IF cm_LogTurnOn=1 THEN
		requestKey=CreateRequestRecord(pcv_PartnerID,0,0,0,0,resultCount,0,0)
		cm_requestKey_value=requestKey
		Set tmpNode=oXML.createNode(1,cm_requestKey_name,"")
		tmpNode.Text=requestKey
		oRoot.appendChild(tmpNode)
	END IF
	
	Set tmpNode=oXML.createNode(1,cm_requestStatus_name,"")
	tmpNode.Text=cm_SuccessCode
	oRoot.appendChild(tmpNode)
	
	Set tmpNode=oXML.createNode(1,cm_resultCount_name,"")
	tmpNode.Text=resultCount
	oRoot.appendChild(tmpNode)
	
	if resultCount>0 then
	
		Set tmpNode=oXML.createNode(1,cm_products,"")
		oRoot.appendChild(tmpNode)
	
		For i=0 to resultCount-1
			Set strNode=oXML.createNode(1,prdID_name,"")
			strNode.Text=pcArr(0,i)
			tmpNode.appendChild(strNode)
		Next
	
	end if
	
End Sub

Function GenNewProductsQuery()
Dim strSQL, query, tmpFromDate
Dim rs1,tmpLastID
on error resume next

	tmpLastID=0
	
	call opendb()
	
	query="SELECT pcXL_LastID FROM pcXMLLogs WHERE pcXP_id=" & pcv_PartnerID & " AND pcXL_RequestType=6 ORDER BY pcXL_LastID DESC;"
	set rs1=connTemp.execute(query)

	if Err.number<>0 then
		set rs1=nothing
		call closedb()
		call XMLcreateError(115,cm_errorStr_115)
		call returnXML()
	end if
	if not rs1.eof then
		tmpLastID=rs1("pcXL_LastID")
	end if
	set rs1=nothing
	
	call closedb()
	
	strSQL=""
	
	IF Clng(tmpLastID)>0 THEN
		strSQL=strSQL & " products.idproduct>" & tmpLastID
	ELSE
		If srcFromDate_ex=0 then
			srcFromDate_ex=1
			srcFromDate_Value=Date()-7
		End if

		tmpFromDate=srcFromDate_Value
		if SQL_Format="1" then
			tmpFromDate=Day(tmpFromDate)&"/"&Month(tmpFromDate)&"/"&Year(tmpFromDate)
		else
			tmpFromDate=Month(tmpFromDate)&"/"&Day(tmpFromDate)&"/"&Year(tmpFromDate)
		end if
		if scDB="Access" then
			strSQL=strSQL & " products.pcprod_EnteredOn>=#" & tmpFromDate & "# "
		else
			strSQL=strSQL & " products.pcprod_EnteredOn>='" & tmpFromDate & "' "
		end if
	END IF
	
	query="SELECT products.idproduct FROM Products WHERE " & strSQL & " AND configOnly=0 AND active<>0 AND removed=0 ORDER BY products.idproduct ASC;"
	
	GenNewProductsQuery=query
	
End Function

Sub RunNewProducts()
Dim query,rs1,resultCount,pcArr
Dim requestKey,i,strNode,tmpLastID
on error resume next
	query=GenNewProductsQuery()
	call opendb()
	set rs1=connTemp.execute(query)
	resultCount=0
	if Err.number<>0 then
		set rs1=nothing
		call closedb()
		call XMLcreateError(115,cm_errorStr_115)
		call returnXML()
	end if
	if not rs1.eof then
		pcArr=rs1.getRows()
		resultCount=ubound(pcArr,2)+1
	end if
	set rs1=nothing
	
	tmpLastID=0
	query="SELECT products.idProduct FROM Products ORDER BY products.idProduct DESC;"
	set rs1=connTemp.execute(query)
	if Err.number<>0 then
		set rs1=nothing
		call closedb()
		call XMLcreateError(115,cm_errorStr_115)
		call returnXML()
	end if
	if not rs1.eof then
		tmpLastID=rs1("idProduct")
	end if
	set rs1=nothing
	
	call closedb()
	
	IF cm_LogTurnOn=1 THEN
		requestKey=CreateRequestRecord(pcv_PartnerID,6,0,0,0,resultCount,tmpLastID,0)
		cm_requestKey_value=requestKey	
		Set tmpNode=oXML.createNode(1,cm_requestKey_name,"")
		tmpNode.Text=requestKey
		oRoot.appendChild(tmpNode)
	END IF
	
	Set tmpNode=oXML.createNode(1,cm_requestStatus_name,"")
	tmpNode.Text=cm_SuccessCode
	oRoot.appendChild(tmpNode)
	
	Set tmpNode=oXML.createNode(1,cm_resultCount_name,"")
	tmpNode.Text=resultCount
	oRoot.appendChild(tmpNode)
	
	if resultCount>0 then
	
		Set tmpNode=oXML.createNode(1,cm_products,"")
		oRoot.appendChild(tmpNode)
	
		For i=0 to resultCount-1
			Set strNode=oXML.createNode(1,prdID_name,"")
			strNode.Text=pcArr(0,i)
			tmpNode.appendChild(strNode)
		Next
	
	end if
	
End Sub

Sub XMLgetPrdCustomField(parentNode,pcv_IDProduct)
Dim query,rsQ,attNode,k,intCount1,tmpValue
	
	call opendb()
	
	query="SELECT pcSearchFields.idSearchField,pcSearchFields.pcSearchFieldName,pcSearchData.idSearchData,pcSearchData.pcSearchDataName,pcSearchData.pcSearchDataOrder FROM pcSearchFields INNER JOIN (pcSearchData INNER JOIN pcSearchFields_Products ON pcSearchData.idSearchData=pcSearchFields_Products.idSearchData) ON pcSearchFields.idSearchField=pcSearchData.idSearchField WHERE pcSearchFields_Products.idproduct=" & pcv_IDProduct & ";"
	set rsQ=connTemp.execute(query)
	if not rsQ.eof then
		tmpValue=rsQ.getRows()
		set rsQ=nothing
		intCount1=ubound(tmpValue,2)
		For k=0 to intCount1
			Set attNode=oXML.createNode(1,prdCustomField_name,"")
			parentNode.appendChild(attNode)
		
			Call XMLCreateNode(attNode,cfID_name,tmpValue(0,k))
			Call XMLCreateNode(attNode,cfName_name,New_HTMLEncode(tmpValue(1,k)))
			Call XMLCreateNode(attNode,cfValue_name,New_HTMLEncode(tmpValue(3,k)))
		Next
	end if
	set rsQ=nothing
	
	call closedb()

End Sub

Sub XMLgetAttrList(grpNode,tmp_IDProduct,tmp_IDOption)
Dim query,rs2,attNode,subNode,pcArr2,intCount2,j
Dim tmpOptName,tmpOptPrice,tmpOptWPrice,tmpOptOrder,tmpOptInactive

	query =	"SELECT options.optionDescrip, options_optionsGroups.price, options_optionsGroups.Wprice, options_optionsGroups.sortOrder, options_optionsGroups.InActive "
	query = query & "FROM options_optionsGroups "
	query = query & "INNER JOIN options "
	query = query & "ON options_optionsGroups.idOption = options.idOption "
	query = query & "WHERE options_optionsGroups.idOptionGroup=" & tmp_IDOption &" "
	query = query & "AND options_optionsGroups.idProduct=" & tmp_IDProduct &" "
	query = query & "ORDER BY options_optionsGroups.sortOrder;"	
	set rs2=conntemp.execute(query)
	
	if not rs2.eof then
		pcArr2=rs2.getRows()
		set rs2=nothing
		intCount2=ubound(pcArr2,2)
		
		For j=0 to intCount2
			Set attNode=oXML.createNode(1,option_name,"")
			grpNode.appendChild(attNode)
			
			tmpOptName=trim(pcArr2(0,j))
			tmpOptPrice=trim(pcArr2(1,j))
			tmpOptWPrice=trim(pcArr2(2,j))
			tmpOptOrder=trim(pcArr2(3,j))
			tmpOptInactive=trim(pcArr2(4,j))
			
			Call XMLCreateNode(attNode,optName_name,New_HTMLEncode(tmpOptName))
			Call XMLCreateNode(attNode,optPrice_name,tmpOptPrice)
			Call XMLCreateNode(attNode,optWPrice_name,tmpOptWPrice)
			Call XMLCreateNode(attNode,optOrder_name,tmpOptOrder)
			Call XMLCreateNode(attNode,optInactive_name,tmpOptInactive)
		Next
	end if
	
	set rs2=nothing
	
End Sub

Sub XMLgetPrdOptions(prdNode,pcv_IDProduct)
Dim query,rs1,pcArr1,intCount1,i,attNode,subNode
Dim tmpGrpID,tmpGrpName,tmpGrpReq,tmpGrpOrder

	call opendb()

	query="SELECT pcProductsOptions.idOptionGroup,optionsGroups.OptionGroupDesc,pcProductsOptions.pcProdOpt_Required,pcProductsOptions.pcProdOpt_order FROM optionsGroups INNER JOIN pcProductsOptions ON optionsGroups.idOptionGroup=pcProductsOptions.idOptionGroup WHERE pcProductsOptions.idProduct=" & pcv_IDProduct & ";"
	set rs1=connTemp.execute(query)
	if not rs1.eof then
		pcArr1=rs1.getRows()
		set rs1=nothing
		intCount1=ubound(pcArr1,2)
		For i=0 to intCount1
			Set attNode=oXML.createNode(1,prdOptionGroup_name,"")
			prdNode.appendChild(attNode)
			
			tmpGrpID=trim(pcArr1(0,i))
			tmpGrpName=trim(pcArr1(1,i))
			tmpGrpReq=trim(pcArr1(2,i))
			tmpGrpOrder=trim(pcArr1(3,i))
			
			Call XMLCreateNode(attNode,groupName_name,New_HTMLEncode(tmpGrpName))
			Call XMLCreateNode(attNode,groupRequired_name,tmpGrpReq)
			Call XMLCreateNode(attNode,groupOrder_name,tmpGrpOrder)
			
			call XMLgetAttrList(attNode,pcv_IDProduct,tmpGrpID)
		Next
	end if
	set rs1=nothing
	
	call closedb()

End Sub

Sub XMLgetCATInfor(prdNode,pcv_IDProduct)
Dim query,rs1,attNode,subNode,pcArr1,intCount1,i,tmpParentName

	call opendb()

	query="SELECT categories.idCategory,categories.[image],categories.largeimage,categories.idParentCategory,categories.categoryDesc,categories.LDesc,categories.SDesc FROM categories INNER JOIN categories_products ON categories.idCategory=categories_products.idCategory WHERE categories_products.idProduct=" & pcv_IDProduct & ";"
	set rs1=connTemp.execute(query)
	if not rs1.eof then
		pcArr1=rs1.getRows()
		set rs1=nothing
		intCount1=ubound(pcArr1,2)
		For i=0 to intCount1
			Set attNode=oXML.createNode(1,prdCategory_name,"")
			prdNode.appendChild(attNode)
			
			Call XMLCreateNode(attNode,catID_name,trim(pcArr1(0,i)))
			Call XMLCreateNode(attNode,catName_name,New_HTMLEncode(trim(pcArr1(4,i))))
			Call XMLCreateNode(attNode,catLDesc_name,New_HTMLEncode(trim(pcArr1(5,i))))
			Call XMLCreateNode(attNode,catSDesc_name,New_HTMLEncode(trim(pcArr1(6,i))))
			Call XMLCreateNode(attNode,catImg_name,New_HTMLEncode(trim(pcArr1(1,i))))
			Call XMLCreateNode(attNode,catLargeImg_name,New_HTMLEncode(trim(pcArr1(2,i))))
			Call XMLCreateNode(attNode,catParentID_name,trim(pcArr1(3,i)))
			
			query="SELECT categoryDesc FROM categories WHERE idcategory=" & trim(pcArr1(3,i)) & ";"
			set rs1=connTemp.execute(query)
			
			if not rs1.eof then
				tmpParentName=trim(rs1("categoryDesc"))
				Call XMLCreateNode(attNode,catParentName_name,New_HTMLEncode(tmpParentName))
			end if
			
			set rs1=nothing
		Next
	end if
	set rs1=nothing
	
	call closedb()

End Sub

Sub XMLgetGCInfor(parentNode,pcv_IDProduct,pcv_GiftCert)
Dim query,rs1,pcArr1

	call opendb()
	
	IF pcv_GiftCert<>"0" THEN
		query="SELECT pcGC_Exp,pcGC_EOnly,pcGC_CodeGen,pcGC_ExpDate,pcGC_ExpDays,pcGC_GenFile FROM pcGC WHERE pcGC_IDProduct=" & pcv_IDProduct & ";"
		set rs1=connTemp.execute(query)
		if not rs1.eof then
			pcArr1=rs1.getRows()
			set rs1=nothing
			
			Call XMLCreateNode(parentNode,giExpire_name,trim(pcArr1(0,0)))
			Call XMLCreateNode(parentNode,giEOnly_name,trim(pcArr1(1,0)))
			Call XMLCreateNode(parentNode,giUseGen_name,trim(pcArr1(2,0)))
			
			tmpDate=trim(pcArr1(3,0))
			if tmpDate<>"" then
				tmpDate=ConvertToXMLDate(tmpDate)
			end if
			
			Call XMLCreateNode(parentNode,giExpDate_name,tmpDate)
			Call XMLCreateNode(parentNode,giExpNDays_name,trim(pcArr1(4,0)))
			Call XMLCreateNode(parentNode,giCustomGen_name,New_HTMLEncode(trim(pcArr1(5,0))))
			
		end if
		set rs1=nothing
	END IF
	
	call closedb()

End Sub

Sub XMLgetDownloadInfor(parentNode,pcv_IDProduct,pcv_Downloadable)
Dim query,rs1,pcArr1

	call opendb()
	
	IF pcv_Downloadable<>"0" THEN
		query="SELECT ProductURL,URLExpire,ExpireDays,License,LocalLG,RemoteLG,LicenseLabel1,LicenseLabel2,LicenseLabel3,LicenseLabel4,LicenseLabel5,AddToMail FROM DProducts WHERE IDProduct=" & pcv_IDProduct & ";"
		set rs1=connTemp.execute(query)
		if not rs1.eof then
			pcArr1=rs1.getRows()
			set rs1=nothing
			
			Call XMLCreateNode(parentNode,diFileLocation_name,New_HTMLEncode(pcArr1(0,0)))
			Call XMLCreateNode(parentNode,diURLExpire_name,pcArr1(1,0))
			Call XMLCreateNode(parentNode,diURLExpDays_name,pcArr1(2,0))
			Call XMLCreateNode(parentNode,diUseLicenseGen_name,pcArr1(3,0))
			Call XMLCreateNode(parentNode,diLocalGen_name,New_HTMLEncode(pcArr1(4,0)))
			Call XMLCreateNode(parentNode,diRemoteGen_name,New_HTMLEncode(pcArr1(5,0)))
			Call XMLCreateNode(parentNode,diLFLabel1_name,New_HTMLEncode(pcArr1(6,0)))
			Call XMLCreateNode(parentNode,diLFLabel2_name,New_HTMLEncode(pcArr1(7,0)))
			Call XMLCreateNode(parentNode,diLFLabel3_name,New_HTMLEncode(pcArr1(8,0)))
			Call XMLCreateNode(parentNode,diLFLabel4_name,New_HTMLEncode(pcArr1(9,0)))
			Call XMLCreateNode(parentNode,diLFLabel5_name,New_HTMLEncode(pcArr1(10,0)))
			Call XMLCreateNode(parentNode,diAddMsg_name,New_HTMLEncode(pcArr1(11,0)))
		end if
		set rs1=nothing
	END IF
	
	call closedb()
	
End Sub

Sub XMLgetBrandInfor(prdNode,pcv_IDBrand)
Dim query,rs1,attNode,subNode,tmpBrandName,tmpBrandLogo
	
	Set attNode=oXML.createNode(1,prdBrand_name,"")
	prdNode.appendChild(attNode)
	
	Call XMLCreateNode(attNode,brandID_name,pcv_IDBrand)
	
	call opendb()
	
	query="SELECT BrandName,BrandLogo FROM Brands WHERE IdBrand=" & pcv_IDBrand & ";"
	set rs1=connTemp.execute(query)

	if not rs1.eof then
		tmpBrandName=trim(rs1("BrandName"))
		tmpBrandLogo=trim(rs1("BrandLogo"))
		
		Call XMLCreateNode(attNode,brandName_name,New_HTMLEncode(tmpBrandName))
		Call XMLCreateNode(attNode,brandLogo_name,New_HTMLEncode(tmpBrandLogo))
	end if

	set rs1=nothing
	
	call closedb()

End Sub

Sub RunGetProductDetails()
Dim query,rs,prdNode,i,pcArr,pcv_HaveRecords,attNode,subNode,queryQ,rsQ,tmpExportedFlag
	
	call opendb()
	
	query="SELECT sku, description, serviceSpec, configOnly, price, listPrice, bToBPrice, weight, stock, IDBrand, smallImageUrl, imageUrl, largeImageURL, active, listHidden, hotDeal, iRewardPoints, notax, noshipping, formQuantity, emailText, noStock, noshippingtext, pcprod_minimumqty, pcprod_qtyvalidate, OverSizeSpec, cost, pcProd_BackOrder, pcProd_ShipNDays, pcProd_NotifyStock, pcProd_ReorderLevel, pcProd_IsDropShipped, pcSupplier_ID, pcDropShipper_ID, Downloadable, pcprod_GC, pcprod_hidebtoprice, pcprod_HideDefConfig, NoPrices, pcProd_SkipDetailsPage, idproduct, ShowInHome, pcprod_EnteredOn, custom1, content1, custom2, content2, custom3, content3, pcprod_QtyToPound, pcProd_MetaTitle, pcProd_MetaDesc, pcProd_MetaKeywords, details, sDesc FROM Products WHERE removed=0 AND idproduct=" & prdID_value & ";"
	'Last: 54
	
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		pcArr=rs.GetRows()
		pcv_HaveRecords=1
	end if
	set rs=nothing
	
	call closedb()
	
	IF pcv_HaveRecords=1 THEN
		i=0
		
		IF cm_LogTurnOn=1 THEN
			requestKey=CreateRequestRecord(pcv_PartnerID,3,prdID_value,0,0,0,0,0)
			cm_requestKey_value=requestKey
			Set tmpNode=oXML.createNode(1,cm_requestKey_name,"")
			tmpNode.Text=requestKey
			oRoot.appendChild(tmpNode)
		END IF
		
		Set tmpNode=oXML.createNode(1,cm_requestStatus_name,"")
		tmpNode.Text=cm_SuccessCode
		oRoot.appendChild(tmpNode)
		
		if cm_ExportAdmin="1" then
			tmpExportedFlag=0
			call opendb()
			queryQ="SELECT pcXEL_ExportedID FROM pcXMLExportLogs WHERE pcXP_ID=" & pcv_PartnerID & " AND pcXEL_IDType=0 AND pcXEL_ExportedID=" & prdID_value & ";"
			set rsQ=connTemp.execute(queryQ)
			if not rsQ.eof then
				tmpExportedFlag=1
			else
				queryQ="INSERT INTO pcXMLExportLogs (pcXP_ID,pcXEL_ExportedID,pcXEL_IDType) VALUES (" & pcv_PartnerID & "," & prdID_value & ",0);"
				set rsQ=connTemp.execute(queryQ)
			end if
			set rsQ=nothing
			call closedb()
			Set tmpNode=oXML.createNode(1,cm_ExportedFlag_name,"")
			tmpNode.Text=New_HTMLEncode(tmpExportedFlag)
			oRoot.appendChild(tmpNode)
		end if
		
		Set prdNode=oXML.createNode(1,cm_product,"")
		oRoot.appendChild(prdNode)
	
		Set attNode=oXML.createNode(1,prdID_name,"")
		attNode.Text=New_HTMLEncode(prdID_value)
		prdNode.appendChild(attNode)
	
		if prdSKU_ex=1 then
			Call XMLCreateNode(prdNode,prdSKU_name,New_HTMLEncode(trim(pcArr(0,i))))
		end if
		
		if prdName_ex=1 then
			Call XMLCreateNode(prdNode,prdName_name,New_HTMLEncode(trim(pcArr(1,i))))
		end if
	
		if prdDesc_ex=1 then
			Call XMLCreateNode(prdNode,prdDesc_name,New_HTMLEncode(trim(pcArr(53,i))))
		end if
		
		if prdSDesc_ex=1 then
			Call XMLCreateNode(prdNode,prdSDesc_name,New_HTMLEncode(trim(pcArr(54,i))))
		end if
		
		if prdType_ex=1 then
			if pcArr(2,i)<>0 then
				tmpPrdType="1"
			else
				if pcArr(3,i)<>0 then
					tmpPrdType="2"
				else
					tmpPrdType="0"
				end if
			end if
			Call XMLCreateNode(prdNode,prdType_name,tmpPrdType)
		end if
		
		if prdPrice_ex=1 then
			Call XMLCreateNode(prdNode,prdPrice_name,trim(pcArr(4,i)))
		end if
		
		if prdListPrice_ex=1 then
			Call XMLCreateNode(prdNode,prdListPrice_name,trim(pcArr(5,i)))
		end if
		
		if prdWPrice_ex=1 then
			Call XMLCreateNode(prdNode,prdWPrice_name,trim(pcArr(6,i)))
		end if
	
		if prdWeight_ex=1 then
			Set attNode=oXML.createNode(1,prdWeight_name,"")
			prdNode.appendChild(attNode)
			tmpPrdWeight=trim(pcArr(7,i))
			
			if scShipFromWeightUnit="KGS" then
				tmp_weight=Int(tmpPrdWeight/1000)
				tmp_weight1=tmpPrdWeight-(tmp_weight*1000)
				
				Call XMLCreateNode(attNode,Kgs_name,tmp_weight)
				Call XMLCreateNode(attNode,Grams_name,tmp_weight1)
			else
				tmp_weight=Int(tmpPrdWeight/16)
				tmp_weight1=tmpPrdWeight-(tmp_weight*16)
				
				Call XMLCreateNode(attNode,Pounds_name,tmp_weight)
				Call XMLCreateNode(attNode,Ounces_name,tmp_weight1)
			end if
			tmpUnitsToPounds=trim(pcArr(49,i))
			if IsNull(tmpUnitsToPounds) or tmpUnitsToPounds="" then
				tmpUnitsToPounds=0
			end if
			Call XMLCreateNode(attNode,UnitsToPound_name,tmpUnitsToPounds)
		end if
		
		if prdStock_ex=1 then
			Call XMLCreateNode(prdNode,prdStock_name,trim(pcArr(8,i)))
		end if
		
		if prdCategory_ex=1 then
			call XMLgetCATInfor(prdNode,prdID_value)
		end if
		
		if prdBrand_ex=1 then
			call XMLgetBrandInfor(prdNode,trim(pcArr(9,i)))
		end if
		
		if prdSmallImg_ex=1 then
			Call XMLCreateNode(prdNode,prdSmallImg_name,New_HTMLEncode(trim(pcArr(10,i))))
		end if
		
		if prdImg_ex=1 then
			Call XMLCreateNode(prdNode,prdImg_name,New_HTMLEncode(trim(pcArr(11,i))))
		end if
		
		if prdLargeImg_ex=1 then
			Call XMLCreateNode(prdNode,prdLargeImg_name,New_HTMLEncode(trim(pcArr(12,i))))
		end if
		
		if prdActive_ex=1 then
			Call XMLCreateNode(prdNode,prdActive_name,trim(pcArr(13,i)))
		end if
		
		if prdShowSavings_ex=1 then
			Call XMLCreateNode(prdNode,prdShowSavings_name,trim(pcArr(14,i)))
		end if
		
		if prdSpecial_ex=1 then
			Call XMLCreateNode(prdNode,prdSpecial_name,trim(pcArr(15,i)))
		end if
		
		if prdFeatured_ex=1 then
			Call XMLCreateNode(prdNode,prdFeatured_name,trim(pcArr(41,i)))
		end if
		
		if prdOptionGroup_ex=1 then
			call XMLgetPrdOptions(prdNode,prdID_value)
		end if
		
		if prdRewardPoints_ex=1 then
			Call XMLCreateNode(prdNode,prdRewardPoints_name,trim(pcArr(16,i)))
		end if
		
		if prdNoTax_ex=1 then
			Call XMLCreateNode(prdNode,prdNoTax_name,trim(pcArr(17,i)))
		end if
		
		if prdNoShippingCharge_ex=1 then
			Call XMLCreateNode(prdNode,prdNoShippingCharge_name,trim(pcArr(18,i)))
		end if
		
		if prdNotForSale_ex=1 then
			Call XMLCreateNode(prdNode,prdNotForSale_name,trim(pcArr(19,i)))
		end if
		
		if prdNotForSaleMsg_ex=1 then
			Call XMLCreateNode(prdNode,prdNotForSaleMsg_name,New_HTMLEncode(trim(pcArr(20,i))))
		end if
		
		if prdDisregardStock_ex=1 then
			Call XMLCreateNode(prdNode,prdDisregardStock_name,trim(pcArr(21,i)))
		end if
		
		if prdDisplayNoShipText_ex=1 then
			Call XMLCreateNode(prdNode,prdDisplayNoShipText_name,trim(pcArr(22,i)))
		end if
		
		if prdMinimumQty_ex=1 then
			Call XMLCreateNode(prdNode,prdMinimumQty_name,trim(pcArr(23,i)))
		end if
		
		if prdPurchaseMulti_ex=1 then
			Call XMLCreateNode(prdNode,prdPurchaseMulti_name,trim(pcArr(24,i)))
		end if
		
		if prdOversize_ex=1 then
			Set attNode=oXML.createNode(1,prdOversize_name,"")
			prdNode.appendChild(attNode)
			
			tmp_OverSize=trim(pcArr(25,i))
			
			if (tmp_OverSize<>"") AND (tmp_OverSize<>"NO") then
				tmpOverSize=split(tmp_OverSize,"||")
				Call XMLCreateNode(attNode,osWidth_name,tmpOverSize(0))
				Call XMLCreateNode(attNode,osHeight_name,tmpOverSize(1))
				Call XMLCreateNode(attNode,osLength_name,tmpOverSize(2))
			end if
		end if
	
		if prdCost_ex=1 then
			Call XMLCreateNode(prdNode,prdCost_name,trim(pcArr(26,i)))
		end if
		
		if prdBackOrder_ex=1 then
			Call XMLCreateNode(prdNode,prdBackOrder_name,trim(pcArr(27,i)))
		end if
		
		if prdShipNDays_ex=1 then
			Call XMLCreateNode(prdNode,prdShipNDays_name,trim(pcArr(28,i)))
		end if
		
		if prdLowStockNotice_ex=1 then
			Call XMLCreateNode(prdNode,prdLowStockNotice_name,trim(pcArr(29,i)))
		end if
		
		if prdReorderLevel_ex=1 then
			Call XMLCreateNode(prdNode,prdReorderLevel_name,trim(pcArr(30,i)))
		end if
		
		if prdIsDropShipped_ex=1 then
			Call XMLCreateNode(prdNode,prdIsDropShipped_name,trim(pcArr(31,i)))
		end if
		
		if prdSupplierID_ex=1 then
			Call XMLCreateNode(prdNode,prdSupplierID_name,trim(pcArr(32,i)))
		end if
		
		if prdDropShipperID_ex=1 then
			Call XMLCreateNode(prdNode,prdDropShipperID_name,trim(pcArr(33,i)))
		end if
		
		if prdMetaTags_ex=1 then
			tmp_MtTitle=trim(pcArr(50,i))
			tmp_MtDesc=trim(pcArr(51,i))
			tmp_MtKey=trim(pcArr(52,i))
			
			Set attNode=oXML.createNode(1,prdMetaTags_name,"")
			prdNode.appendChild(attNode)
			
			Call XMLCreateNode(attNode,mtTitle_name,New_HTMLEncode(tmp_MtTitle))
			Call XMLCreateNode(attNode,mtDesc_name,New_HTMLEncode(tmp_MtDesc))
			Call XMLCreateNode(attNode,mtKeywords_name,New_HTMLEncode(tmp_MtKey))
		end if
	
		if prdDownloadable_ex=1 then
			Call XMLCreateNode(prdNode,prdDownloadable_name,trim(pcArr(34,i)))
		end if
	
		if prdDownloadInfo_ex=1 then
			Set attNode=oXML.createNode(1,prdDownloadInfo_name,"")
			prdNode.appendChild(attNode)
			Call XMLgetDownloadInfor(attNode,prdID_value,trim(pcArr(34,i)))
		end if
		
		if prdGiftCertificate_ex=1 then
			Call XMLCreateNode(prdNode,prdGiftCertificate_name,trim(pcArr(35,i)))
		end if
		
		if prdGCInfo_ex=1 then
			Set attNode=oXML.createNode(1,prdGCInfo_name,"")
			prdNode.appendChild(attNode)
			Call XMLgetGCInfor(attNode,prdID_value,trim(pcArr(35,i)))
		end if
		
		if prdHideBTOPrices_ex=1 then
			Call XMLCreateNode(prdNode,prdHideBTOPrices_name,trim(pcArr(36,i)))
		end if
		
		if prdHideDefaultConfig_ex=1 then
			Call XMLCreateNode(prdNode,prdHideDefaultConfig_name,trim(pcArr(37,i)))
		end if
		
		if prdDisallowPurchasing_ex=1 then
			Call XMLCreateNode(prdNode,prdDisallowPurchasing_name,trim(pcArr(38,i)))
		end if
		
		if prdSkipPrdPage_ex=1 then
			Call XMLCreateNode(prdNode,prdSkipPrdPage_name,trim(pcArr(39,i)))
		end if
		
		if prdCustomField_ex=1 then
			Call XMLgetPrdCustomField(prdNode,prdID_value)
		end if
		
		if prdCreatedDate_ex=1 then
			tmpCreatedDate=trim(pcArr(42,i))
			if tmpCreatedDate<>"" then
				tmpCreatedDate=ConvertToXMLDate(tmpCreatedDate)
			end if
			Call XMLCreateNode(prdNode,prdCreatedDate_name,tmpCreatedDate)
		end if
		
		Set pXML1=Server.CreateObject("MSXML2.DOMDocument"&scXML)
		pXML1.async=false
		pXML1.load(oXML)
		If (pXML1.parseError.errorCode <> 0) Then	
			Set oXML=nothing
			call InitResponseDocument(cm_GetProductDetailsResponse_name)
			call XMLcreateError(pXML1.parseError.errorCode, pXML1.parseError.reason)
			call returnXML()
		End If
		set pXML1 = nothing
		
	ELSE
		call XMLcreateError(116,cm_errorStr_116)
		call returnXML()
	END IF 'Have product record
	
End Sub

Sub PresetPrdValues()
	prdID_value=0
	prdType_value=0
	prdPrice_value=0
	prdListPrice_value=0
	prdWPrice_value=0
	prdWeight_value=0
	Pounds_value=0
	Ounces_value=0
	Kgs_value=0
	Grams_value=0
	UnitsToPound_value=0
	prdStock_value=0
	brandID_value=0
	prdActive_value=-1
	prdShowSavings_value=0
	prdSpecial_value=0
	prdFeatured_value=0
	prdRewardPoints_value=0
	prdNoTax_value=0
	prdNoShippingCharge_value=0
	prdNotForSale_value=0
	prdDisregardStock_value=0
	prdDisplayNoShipText_value=0
	prdMinimumQty_value=0
	prdPurchaseMulti_value=0
	prdOversize_value="NO"
	osWidth_value=0
	osHeight_value=0
	osLength_value=0
	osX_value=0
	osWeight_value=0
	prdCost_value=0
	prdBackOrder_value=0
	prdShipNDays_value=0
	prdLowStockNotice_value=0
	prdReorderLevel_value=0
	prdIsDropShipped_value=0
	prdSupplierID_value=0
	prdDropShipperID_value=0
	prdDownloadable_value=0
	diURLExpire_value=0
	diURLExpDays_value=0
	diUseLicenseGen_value=0
	prdGiftCertificate_value=0
	giExpire_value=0
	giEOnly_value=0
	giUseGen_value=0
	giExpNDays_value=0
	prdHideBTOPrices_value=0
	prdHideDefaultConfig_value=0
	prdDisallowPurchasing_value=0
	prdSkipPrdPage_value=0
End Sub

Sub CheckAddUpdProduct(requestType)

	Dim ChildNodes,strNode,tmpNodeName,tmpNodeValue,tmpValue1,subNode
	
	Call CheckRequiredXMLTag(cm_product)
	Call PresetPrdValues()
	
	if requestType=0 then
		Call CheckRequiredXMLTag(cm_product & "/" & prdSKU_name)
		Call CheckRequiredXMLTag(cm_product & "/" & prdName_name)
		Call CheckRequiredXMLTag(cm_product & "/" & prdDesc_name)
		Call CheckRequiredXMLTag(cm_product & "/" & prdPrice_name)
	else
		Call CheckRequiredXMLTag(cm_product & "/" & ImportField_name)
	end if
	
	Set rNode=iRoot.selectSingleNode(cm_product)
	Set ChildNodes = rNode.childNodes
	
	For Each strNode In ChildNodes
		tmpNodeName=strNode.nodeName
		tmpNodeValue=trim(strNode.Text)
			Select Case tmpNodeName				
				Case ImportField_name:
					if requestType=1 then
						call CheckValidXMLTag(strNode,1,1,"")
						ImportField_ex=1
						ImportField_value=tmpNodeValue						
						Select Case ImportField_value
							Case 1:
								Call CheckRequiredXMLTag(cm_product & "/" & prdSKU_name)
							Case 2:
								Call CheckRequiredXMLTag(cm_product & "/" & prdID_name)
							Case Else
								call XMLcreateError(128, cm_errorStr_128)
								call returnXML()
						End Select						
					end if	
				Case prdID_name:
					call CheckValidXMLTag(strNode,1,1,"")
					prdID_ex=1
					prdID_value=getUserInputNew(tmpNodeValue,0)			
				Case prdSKU_name:
					call CheckValidXMLTag(strNode,1,5,"")
					prdSKU_ex=1
					prdSKU_value=getUserInputNew(tmpNodeValue,0)					
				Case prdName_name:
					call CheckValidXMLTag(strNode,1,5,"")
					prdName_ex=1
					prdName_value=getUserInputNew(tmpNodeValue,0)
				Case prdDesc_name:
					call CheckValidXMLTag(strNode,1,5,"")
					prdDesc_ex=1
					prdDesc_value=getUserInputNew(tmpNodeValue,0)
				Case prdSDesc_name:
					call CheckValidXMLTag(strNode,0,5,"")
					prdSDesc_ex=1
					prdSDesc_value=getUserInputNew(tmpNodeValue,0)
				Case prdType_name:
					call CheckValidXMLTag(strNode,1,1,"")
					prdType_ex=1
					prdType_value=tmpNodeValue
				Case prdPrice_name:
					call CheckValidXMLTag(strNode,1,3,"")
					prdPrice_ex=1
					prdPrice_value=tmpNodeValue
				Case prdListPrice_name:
					call CheckValidXMLTag(strNode,1,3,"")
					prdListPrice_ex=1
					prdListPrice_value=tmpNodeValue
				Case prdWPrice_name:
					call CheckValidXMLTag(strNode,1,3,"")
					prdWPrice_ex=1
					prdWPrice_value=tmpNodeValue
				Case prdWeight_name:
					if scShipFromWeightUnit="KGS" then
						if CheckExistTag(cm_product & "/" & prdWeight_name & "/" & Kgs_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdWeight_name & "/" & Kgs_name)
							call CheckValidXMLTag(subNode,1,3,"")
							Kgs_ex=1
							Kgs_value=subNode.Text
						end if
						if CheckExistTag(cm_product & "/" & prdWeight_name & "/" & Grams_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdWeight_name & "/" & Grams_name)
							call CheckValidXMLTag(subNode,1,3,"")
							Grams_ex=1
							Grams_value=subNode.Text
						end if
						if (Kgs_ex=1) OR (Grams_ex=1) then
							prdWeight_ex=1
						end if
					else
						if CheckExistTag(cm_product & "/" & prdWeight_name & "/" & Pounds_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdWeight_name & "/" & Pounds_name)
							call CheckValidXMLTag(subNode,1,3,"")
							Pounds_ex=1
							Pounds_value=subNode.Text
						end if
						if CheckExistTag(cm_product & "/" & prdWeight_name & "/" & Ounces_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdWeight_name & "/" & Ounces_name)
							call CheckValidXMLTag(subNode,1,3,"")
							Ounces_ex=1
							Ounces_value=subNode.Text
						end if
						if (Pounds_ex=1) OR (Ounces_ex=1) then
							prdWeight_ex=1
						end if
					end if
					if CheckExistTag(cm_product & "/" & prdWeight_name & "/" & UnitsToPound_name) then
						Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdWeight_name & "/" & UnitsToPound_name)
						call CheckValidXMLTag(subNode,1,3,"")
						UnitsToPound_ex=1
						UnitsToPound_value=subNode.Text
					end if
					if (UnitsToPound_ex=1) then
						prdWeight_ex=1
					end if
				Case prdStock_name:
					call CheckValidXMLTag(strNode,1,0,"")
					prdStock_ex=1
					prdStock_value=tmpNodeValue
				Case prdCategory_name:
					if CheckExistTagEx(strNode,catName_name) then
						Call CheckRequiredXMLTagEx(strNode,catName_name)
					else
						Call CheckRequiredXMLTagEx(strNode,catID_name)
					end if
					prdCategory_ex=1
				Case prdBrand_name:
					if CheckExistTag(cm_product & "/" & prdBrand_name & "/" & brandName_name) then
						Call CheckRequiredXMLTag(cm_product & "/" & prdBrand_name & "/" & brandName_name)
					else
						Call CheckRequiredXMLTag(cm_product & "/" & prdBrand_name & "/" & brandID_name)
					end if
					if CheckExistTag(cm_product & "/" & prdBrand_name & "/" & brandID_name) then
						Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdBrand_name & "/" & brandID_name)
						call CheckValidXMLTag(subNode,1,1,"")
						brandID_ex=1
						brandID_value=subNode.Text
					end if
					if CheckExistTag(cm_product & "/" & prdBrand_name & "/" & brandName_name) then
						Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdBrand_name & "/" & brandName_name)
						call CheckValidXMLTag(subNode,0,5,"")
						brandName_ex=1
						brandName_value=getUserInputNew(subNode.Text,0)
					end if
					if CheckExistTag(cm_product & "/" & prdBrand_name & "/" & brandLogo_name) then
						Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdBrand_name & "/" & brandLogo_name)
						call CheckValidXMLTag(subNode,0,5,"")
						brandLogo_ex=1
						brandLogo_value=getUserInputNew(subNode.Text,0)
					end if
					prdBrand_ex=1
				Case prdSmallImg_name:
					call CheckValidXMLTag(strNode,0,5,"")
					prdSmallImg_ex=1
					prdSmallImg_value=getUserInputNew(tmpNodeValue,0)
				Case prdImg_name:
					call CheckValidXMLTag(strNode,0,5,"")
					prdImg_ex=1
					prdImg_value=getUserInputNew(tmpNodeValue,0)
				Case prdLargeImg_name:
					call CheckValidXMLTag(strNode,0,5,"")
					prdLargeImg_ex=1
					prdLargeImg_value=getUserInputNew(tmpNodeValue,0)
				Case prdActive_name:
					call CheckValidXMLTag(strNode,1,0,"")
					prdActive_ex=1
					prdActive_value=tmpNodeValue
					if cInt(prdActive_value)<>0 then
						prdActive_value=-1
					end if
				Case prdShowSavings_name:
					call CheckValidXMLTag(strNode,1,0,"")
					prdShowSavings_ex=1
					prdShowSavings_value=tmpNodeValue
					if cInt(prdShowSavings_value)<>0 then
						prdShowSavings_value=-1
					end if
				Case prdSpecial_name:
					call CheckValidXMLTag(strNode,1,0,"")
					prdSpecial_ex=1
					prdSpecial_value=tmpNodeValue
					if cInt(prdSpecial_value)<>0 then
						prdSpecial_value=-1
					end if
				Case prdFeatured_name:
					call CheckValidXMLTag(strNode,1,0,"")
					prdFeatured_ex=1
					prdFeatured_value=tmpNodeValue
					if cInt(prdFeatured_value)<>0 then
						prdFeatured_value=-1
					end if
				Case prdOptionGroup_name:
					if tmpNodeValue<>"" then
						Call CheckRequiredXMLTagEx(strNode,groupName_name)
						Call CheckRequiredXMLTagEx(strNode,option_name)
						prdOptionGroup_ex=1
					end if
				Case prdRewardPoints_name:
					call CheckValidXMLTag(strNode,1,1,"")
					prdRewardPoints_ex=1
					prdRewardPoints_value=tmpNodeValue
				Case prdNoTax_name:
					call CheckValidXMLTag(strNode,1,0,"")
					prdNoTax_ex=1
					prdNoTax_value=tmpNodeValue
					if cInt(prdNoTax_value)<>0 then
						prdNoTax_value=-1
					end if
				Case prdNoShippingCharge_name:
					call CheckValidXMLTag(strNode,1,0,"")
					prdNoShippingCharge_ex=1
					prdNoShippingCharge_value=tmpNodeValue
					if cInt(prdNoShippingCharge_value)<>0 then
						prdNoShippingCharge_value=-1
					end if
				Case prdNotForSale_name:
					call CheckValidXMLTag(strNode,1,0,"")
					prdNotForSale_ex=1
					prdNotForSale_value=tmpNodeValue
					if cInt(prdNotForSale_value)<>0 then
						prdNotForSale_value=-1
					end if
				Case prdNotForSaleMsg_name:
					call CheckValidXMLTag(strNode,0,5,"")
					prdNotForSaleMsg_ex=1
					prdNotForSaleMsg_value=getUserInputNew(tmpNodeValue,0)
				Case prdDisregardStock_name:
					call CheckValidXMLTag(strNode,1,0,"")
					prdDisregardStock_ex=1
					prdDisregardStock_value=tmpNodeValue
					if cInt(prdDisregardStock_value)<>0 then
						prdDisregardStock_value=-1
					end if
				Case prdDisplayNoShipText_name:
					call CheckValidXMLTag(strNode,1,0,"")
					prdDisplayNoShipText_ex=1
					prdDisplayNoShipText_value=tmpNodeValue
					if cInt(prdDisplayNoShipText_value)<>0 then
						prdDisplayNoShipText_value=-1
					end if
				Case prdMinimumQty_name:
					call CheckValidXMLTag(strNode,1,1,"")
					prdMinimumQty_ex=1
					prdMinimumQty_value=tmpNodeValue
				Case prdPurchaseMulti_name:
					call CheckValidXMLTag(strNode,1,1,"")
					prdPurchaseMulti_ex=1
					prdPurchaseMulti_value=tmpNodeValue
				Case prdOversize_name:
					if tmpNodeValue<>"" then
						if CheckExistTag(cm_product & "/" & prdOversize_name & "/" & osWidth_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdOversize_name & "/" & osWidth_name)
							call CheckValidXMLTag(subNode,1,3,"")
							osWidth_ex=1
							osWidth_value=subNode.Text
						end if
						if CheckExistTag(cm_product & "/" & prdOversize_name & "/" & osHeight_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdOversize_name & "/" & osHeight_name)
							call CheckValidXMLTag(subNode,1,3,"")
							osHeight_ex=1
							osHeight_value=subNode.Text
						end if
						if CheckExistTag(cm_product & "/" & prdOversize_name & "/" & osLength_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdOversize_name & "/" & osLength_name)
							call CheckValidXMLTag(subNode,1,3,"")
							osLength_ex=1
							osLength_value=subNode.Text
						end if
						prdOversize_ex=1
					end if
				Case prdCost_name:
					call CheckValidXMLTag(strNode,1,3,"")
					prdCost_ex=1
					prdCost_value=tmpNodeValue
				Case prdBackOrder_name:
					call CheckValidXMLTag(strNode,1,1,"")
					prdBackOrder_ex=1
					prdBackOrder_value=tmpNodeValue
				Case prdShipNDays_name:
					call CheckValidXMLTag(strNode,1,1,"")
					prdShipNDays_ex=1
					prdShipNDays_value=tmpNodeValue
				Case prdLowStockNotice_name:
					call CheckValidXMLTag(strNode,1,1,"")
					prdLowStockNotice_ex=1
					prdLowStockNotice_value=tmpNodeValue
				Case prdReorderLevel_name:
					call CheckValidXMLTag(strNode,1,1,"")
					prdReorderLevel_ex=1
					prdReorderLevel_value=tmpNodeValue
				Case prdIsDropShipped_name:
					call CheckValidXMLTag(strNode,1,1,"")
					prdIsDropShipped_ex=1
					prdIsDropShipped_value=tmpNodeValue
				Case prdSupplierID_name:
					call CheckValidXMLTag(strNode,1,1,"")
					prdSupplierID_ex=1
					prdSupplierID_value=tmpNodeValue
				Case prdDropShipperID_name:
					call CheckValidXMLTag(strNode,1,1,"")
					prdDropShipperID_ex=1
					prdDropShipperID_value=tmpNodeValue
				Case prdMetaTags_name:
					if tmpNodeValue<>"" then
						if CheckExistTag(cm_product & "/" & prdMetaTags_name & "/" & mtTitle_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdMetaTags_name & "/" & mtTitle_name)
							call CheckValidXMLTag(subNode,0,5,"")
							mtTitle_ex=1
							mtTitle_value=getUserInputNew(subNode.Text,0)
						end if
						if CheckExistTag(cm_product & "/" & prdMetaTags_name & "/" & mtDesc_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdMetaTags_name & "/" & mtDesc_name)
							call CheckValidXMLTag(subNode,0,5,"")
							mtDesc_ex=1
							mtDesc_value=getUserInputNew(subNode.Text,0)
						end if
						if CheckExistTag(cm_product & "/" & prdMetaTags_name & "/" & mtKeywords_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdMetaTags_name & "/" & mtKeywords_name)
							call CheckValidXMLTag(subNode,0,5,"")
							mtKeywords_ex=1
							mtKeywords_value=getUserInputNew(subNode.Text,0)
						end if
						prdMetaTags_ex=1
					end if
				Case prdDownloadable_name:
					call CheckValidXMLTag(strNode,1,1,"")
					prdDownloadable_ex=1
					prdDownloadable_value=tmpNodeValue
				Case prdDownloadInfo_name:
					if tmpNodeValue<>"" then
						if CheckExistTag(cm_product & "/" & prdDownloadInfo_name & "/" & diFileLocation_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdDownloadInfo_name & "/" & diFileLocation_name)
							call CheckValidXMLTag(subNode,0,5,"")
							diFileLocation_ex=1
							diFileLocation_value=getUserInputNew(subNode.Text,0)
						end if
						if CheckExistTag(cm_product & "/" & prdDownloadInfo_name & "/" & diURLExpire_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdDownloadInfo_name & "/" & diURLExpire_name)
							call CheckValidXMLTag(subNode,1,1,"")
							diURLExpire_ex=1
							diURLExpire_value=subNode.Text
						end if
						if CheckExistTag(cm_product & "/" & prdDownloadInfo_name & "/" & diURLExpDays_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdDownloadInfo_name & "/" & diURLExpDays_name)
							call CheckValidXMLTag(subNode,1,1,"")
							diURLExpDays_ex=1
							diURLExpDays_value=subNode.Text
						end if
						if CheckExistTag(cm_product & "/" & prdDownloadInfo_name & "/" & diUseLicenseGen_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdDownloadInfo_name & "/" & diUseLicenseGen_name)
							call CheckValidXMLTag(subNode,1,1,"")
							diUseLicenseGen_ex=1
							diUseLicenseGen_value=subNode.Text
						end if
						if CheckExistTag(cm_product & "/" & prdDownloadInfo_name & "/" & diLocalGen_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdDownloadInfo_name & "/" & diLocalGen_name)
							call CheckValidXMLTag(subNode,0,5,"")
							diLocalGen_ex=1
							diLocalGen_value=getUserInputNew(subNode.Text,0)
						end if
						if CheckExistTag(cm_product & "/" & prdDownloadInfo_name & "/" & diRemoteGen_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdDownloadInfo_name & "/" & diRemoteGen_name)
							call CheckValidXMLTag(subNode,0,5,"")
							diRemoteGen_ex=1
							diRemoteGen_value=getUserInputNew(subNode.Text,0)
						end if
						if CheckExistTag(cm_product & "/" & prdDownloadInfo_name & "/" & diLFLabel1_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdDownloadInfo_name & "/" & diLFLabel1_name)
							call CheckValidXMLTag(subNode,0,5,"")
							diLFLabel1_ex=1
							diLFLabel1_value=getUserInputNew(subNode.Text,0)
						end if
						if CheckExistTag(cm_product & "/" & prdDownloadInfo_name & "/" & diLFLabel2_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdDownloadInfo_name & "/" & diLFLabel2_name)
							call CheckValidXMLTag(subNode,0,5,"")
							diLFLabel2_ex=1
							diLFLabel2_value=getUserInputNew(subNode.Text,0)
						end if
						if CheckExistTag(cm_product & "/" & prdDownloadInfo_name & "/" & diLFLabel3_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdDownloadInfo_name & "/" & diLFLabel3_name)
							call CheckValidXMLTag(subNode,0,5,"")
							diLFLabel3_ex=1
							diLFLabel3_value=getUserInputNew(subNode.Text,0)
						end if
						if CheckExistTag(cm_product & "/" & prdDownloadInfo_name & "/" & diLFLabel4_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdDownloadInfo_name & "/" & diLFLabel4_name)
							call CheckValidXMLTag(subNode,0,5,"")
							diLFLabel4_ex=1
							diLFLabel4_value=getUserInputNew(subNode.Text,0)
						end if
						if CheckExistTag(cm_product & "/" & prdDownloadInfo_name & "/" & diLFLabel5_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdDownloadInfo_name & "/" & diLFLabel5_name)
							call CheckValidXMLTag(subNode,0,5,"")
							diLFLabel5_ex=1
							diLFLabel5_value=getUserInputNew(subNode.Text,0)
						end if
						if CheckExistTag(cm_product & "/" & prdDownloadInfo_name & "/" & diAddMsg_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdDownloadInfo_name & "/" & diAddMsg_name)
							call CheckValidXMLTag(subNode,0,5,"")
							diAddMsg_ex=1
							diAddMsg_value=getUserInputNew(subNode.Text,0)
						end if
						prdDownloadInfo_ex=1
					end if
				Case prdGiftCertificate_name:
					call CheckValidXMLTag(strNode,1,1,"")
					prdGiftCertificate_ex=1
					prdGiftCertificate_value=tmpNodeValue
				Case prdGCInfo_name:
					if tmpNodeValue<>"" then
						if CheckExistTag(cm_product & "/" & prdGCInfo_name & "/" & giExpire_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdGCInfo_name & "/" & giExpire_name)
							call CheckValidXMLTag(subNode,1,1,"")
							giExpire_ex=1
							giExpire_value=subNode.Text
						end if
						if CheckExistTag(cm_product & "/" & prdGCInfo_name & "/" & giEOnly_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdGCInfo_name & "/" & giEOnly_name)
							call CheckValidXMLTag(subNode,1,1,"")
							giEOnly_ex=1
							giEOnly_value=subNode.Text
						end if
						if CheckExistTag(cm_product & "/" & prdGCInfo_name & "/" & giUseGen_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdGCInfo_name & "/" & giUseGen_name)
							call CheckValidXMLTag(subNode,1,1,"")
							giUseGen_ex=1
							giUseGen_value=subNode.Text
						end if
						if CheckExistTag(cm_product & "/" & prdGCInfo_name & "/" & giExpDate_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdGCInfo_name & "/" & giExpDate_name)
							call CheckValidXMLTag(subNode,0,4,"")
							if subNode.Text<>"" then
								giExpDate_ex=1
								giExpDate_value=ConvertFromXMLDate(subNode.Text)
							end if
						end if
						if CheckExistTag(cm_product & "/" & prdGCInfo_name & "/" & giExpNDays_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdGCInfo_name & "/" & giExpNDays_name)
							call CheckValidXMLTag(subNode,1,1,"")
							giExpNDays_ex=1
							giExpNDays_value=subNode.Text
						end if
						if CheckExistTag(cm_product & "/" & prdGCInfo_name & "/" & giCustomGen_name) then
							Set subNode=iRoot.selectSingleNode(cm_product & "/" & prdGCInfo_name & "/" &giCustomGen_name)
							call CheckValidXMLTag(subNode,0,5,"")
							giCustomGen_ex=1
							giCustomGen_value=getUserInputNew(subNode.Text,0)
						end if
						prdGCInfo_ex=1
					end if
				Case prdHideBTOPrices_name:
					call CheckValidXMLTag(strNode,1,1,"")
					prdHideBTOPrices_ex=1
					prdHideBTOPrices_value=tmpNodeValue
				Case prdHideDefaultConfig_name:
					call CheckValidXMLTag(strNode,1,1,"")
					prdHideDefaultConfig_ex=1
					prdHideDefaultConfig_value=tmpNodeValue
				Case prdDisallowPurchasing_name:
					call CheckValidXMLTag(strNode,1,1,"")
					prdDisallowPurchasing_ex=1
					prdDisallowPurchasing_value=tmpNodeValue
				Case prdSkipPrdPage_name:
					call CheckValidXMLTag(strNode,1,1,"")
					prdSkipPrdPage_ex=1
					prdSkipPrdPage_value=tmpNodeValue
				Case prdCustomField_name:
					if tmpNodeValue<>"" then
						if CheckExistTagEx(strNode,cfName_name) then
							Call CheckRequiredXMLTagEx(strNode,cfName_name)
						else
							Call CheckRequiredXMLTagEx(strNode,cfID_name)
						end if
						Call CheckRequiredXMLTagEx(strNode,cfValue_name)
						prdCustomField_ex=1
					end if
				Case prdCreatedDate_name:
					if tmpNodeValue<>"" then
						if requestType=0 then
							call XMLcreateError(117,cm_errorStr_117 & tmpNodeName)
							call returnXML()
						else
							call CheckValidXMLTag(strNode,0,4,"")
							if tmpNodeValue<>"" then
								prdCreatedDate_ex=1
								prdCreatedDate_value=ConvertFromXMLDate(tmpNodeValue)
							end if
						end if
					end if
				Case Else:
					call XMLcreateError(117,cm_errorStr_117 & tmpNodeName)
					call returnXML()
			End Select
	Next
	if scShipFromWeightUnit="KGS" then
		prdWeight_value=Kgs_value*1000+Grams_value
	else
		prdWeight_value=Pounds_value*16+Ounces_value
	end if

	if osWidth_value+osHeight_value+osLength_value=0 then
		prdOversize_value="NO"
	else
		if prdWeight_ex=1 then
			prdOversize_value=osWidth_value & "||" & osHeight_value & "||" & osLength_value & "||1||" & prdWeight_value
		else
			prdOversize_value=osWidth_value & "||" & osHeight_value & "||" & osLength_value & "||1||"
		end if
	end if
End Sub

Sub XMLBackUpPrdData(BackUpType,tmpID)
	Dim query,rs,tmpTable,tmpField,tmp1,tmp2
	Dim BackupStr1,BackupStr2
	
	BackupStr1=""
	BackupStr2=""

	call opendb()
	
	if prdID_value=0 then
		query="SELECT idproduct FROM Products WHERE sku like '" & prdSKU_value & "';"
		set rstemp=conntemp.execute(query)
		if not rstemp.eof then
			prdID_value=rstemp("idproduct")
		end if
		set rstemp=nothing
	end if
	
	tmp2=""
	
	Select Case BackUpType
		Case 0:
			tmpTable="Products"
			tmpField="idproduct"
			BackupStr1=BackupStr1 & "UPDPRD" & chr(9) & prdID_value
			tmp1=1
		Case 1:
			tmpTable="DProducts"
			tmpField="idproduct"
			BackupStr1=BackupStr1 & "ADDDP" & chr(9) & prdID_value
			tmp1=1
		Case 2:
			tmpTable="pcGC"
			tmpField="pcGC_IDProduct"
			BackupStr1=BackupStr1 & "ADDGC" & chr(9) & prdID_value
			tmp1=1
		Case 3:
			tmpTable="options_optionsGroups"
			tmpField="idProduct"
			BackupStr1=BackupStr1 & "UPDPRDOPT"
			tmp1=0
			tmp2=" AND idoptoptgrp=" & tmpID
		Case 4:
			tmpTable="pcProductsOptions"
			tmpField="idProduct"
			BackupStr1=BackupStr1 & "UPDPRDGRP"
			tmp1=0
			tmp2=" AND idOptionGroup=" & tmpID
	End Select
	
	query="SELECT * FROM " & tmpTable & " WHERE " & tmpField & "=" & prdID_value & tmp2 & ";"
	set rstemp=conntemp.execute(query)
	
	IF not rstemp.eof THEN
		iCols = rstemp.Fields.Count
		For dd=tmp1 to iCols-1
			FType="" & Rstemp.Fields.Item(dd).Type
			if (Ftype="202") or (Ftype="203") or (Ftype="135") then
				PTemp=Rstemp.Fields.Item(dd).Value
				if PTemp<>"" then
					PTemp=replace(PTemp,"'","''")
					PTemp=replace(PTemp,vbcrlf,"DuLTVDu")
				end if
	
				if (scDB="Access") and (Ftype="135") then
					BackupStr2=BackupStr2 & chr(9) & "#" & PTemp & "#"
				else
					BackupStr2=BackupStr2 & chr(9) & "'" & PTemp & "'"
				end if
			else
				PTemp="" & Rstemp.Fields.Item(dd).Value
				if PTemp<>"" then
				else
					PTemp="0"
				end if
				BackupStr2=BackupStr2 & chr(9) & PTemp
			end if
		Next
	END IF
	set rstemp=nothing
	
	if BackupStr2<>"" then
		BackupStr=BackupStr & BackupStr1 & BackupStr2  & vbcrlf
	end if
	
	call closedb()
	
End Sub

Sub XMLUpdCatTier(tmpID)
Dim rs,query,tmptier

if tmpID<>"1" then
call opendb()

query="SELECT tier FROM categories WHERE idCategory IN (SELECT idParentCategory FROM categories WHERE idCategory=" & tmpID & ");"
set rs=connTemp.execute(query)

if not rs.eof then
	tmptier=rs("tier")
	set rs=nothing
	if IsNull(tmptier) or tmptier="" then
		tmptier=0
	end if
	tmptier=clng(tmptier)+1
	
	query="UPDATE categories SET tier=" & tmptier & " WHERE idCategory=" & tmpID & ";"
	set rs=connTemp.execute(query)
	set rs=nothing
end if
set rs=nothing

call closedb()
end if

End Sub

Function XMLCheckParent(pid,pname)
Dim mypname,mypname1
Dim rstemp,query,rstemp1
Dim tmpParentID

	mypname=pname
	mypname1=""
	if mypname<>"" then
		mypname=replace(mypname,"&amp;","&")
		mypname=replace(mypname,"&","&amp;")
		mypname1=replace(pname,"&amp;","&")
	end if
	
	call opendb()
	
	if pid>0 then
		query="Select idCategory from categories where idcategory=" & pid & ";"
	else
		query="Select idCategory from categories where (categorydesc like '" & mypname & "' or categorydesc like '" & mypname1 & "')"
	end if
	set rstemp=conntemp.execute(query)
	
	if rstemp.eof then
		imagename="no_image.gif"
		query="insert into categories (categorydesc,idParentCategory,image,largeimage) values ('" & mypname & "',1,'" & imagename & "','" & imagename & "')"
		set rstemp1=conntemp.execute(query)
		query="Select idCategory from categories where categorydesc='" & mypname & "'"
		set rstemp1=conntemp.execute(query)
		tmpParentID=rstemp1("idCategory")
		set rstemp1=nothing
	else
		tmpParentID=rstemp("idCategory")
		set rstemp=nothing
		if mypname<>"" and pid>0 then
			query="UPDATE categories SET categorydesc='" & mypname & "' WHERE idcategory=" & tmpParentID & ";"
			set rstemp1=conntemp.execute(query)
			set rstemp1=nothing
		end if
	end if
	
	set rstemp=nothing
	call closedb()
	
	call XMLUpdCatTier(tmpParentID)
	
	XMLCheckParent=tmpParentID
	
End Function

Function XMLCheckCategory(pid,cname,pcid,simage,limage,SDesc1,LDesc1)
Dim mycname,mycname1,tcheckcategory
Dim rstemp,query,rstemp1
Dim tmp1
	
	mycname=cname
	mycname=replace(mycname,"&amp;","&")
	mycname=replace(mycname,"&","&amp;")
	
	mycname1=replace(cname,"&amp;","&")
	
	call opendb()
	
	if pid>0 then
		query="Select idCategory from categories where idcategory=" & pid & ";"
	else
		query="Select idCategory from categories where (categorydesc like '" & mycname & "' or categorydesc like '" & mycname1 & "') and idParentCategory=" & pcid
	end if
	set rstemp=conntemp.execute(query)

	if rstemp.eof then
		query1="categoryDesc,idParentCategory,image,largeimage,SDesc,LDesc"
		if simage<>"" then
		 smallimg=simage
		else
		 smallimg="no_image.gif"
		end if
		if limage<>"" then
		 largeimg=limage
		else
		 largeimg="no_image.gif"
		end if
		query2="'" & mycname & "'," & pcid & ",'" & smallimg & "','" & largeimg & "','" & SDesc1 & "','" & LDesc1 & "'"
		query="insert into categories (" & query1 & ") values (" & query2 & ")"
		set rstemp1=conntemp.execute(query)
		query="Select idCategory from categories where categorydesc like '" & mycname & "' and idParentCategory=" & pcid
		set rstemp1=conntemp.execute(query)
		tcheckcategory=rstemp1("idCategory")
	else
		tcheckcategory=rstemp("idCategory")
		set rstemp=nothing
		
		if pid>0 then
			tmp1=""
			if cname<>"" then
				tmp1=tmp1 & ",categoryDesc='" & cname & "'"
			end if
			if simage<>"" then
				tmp1=tmp1 & ",image='" & simage & "'"
			end if
			if limage<>"" then
				tmp1=tmp1 & ",largeimage='" & limage & "'"
			end if
			if SDesc1<>"" then
				tmp1=tmp1 & ",SDesc='" & SDesc1 & "'"
			end if
			if LDesc1<>"" then
				tmp1=tmp1 & ",LDesc='" & LDesc1 & "'"
			end if
			if pcid>0 then
				tmp1=tmp1 & ",idParentCategory='" & pcid & "'"
			end if
			if tmp1<>"" then
				tmp1=mid(tmp1,2,len(tmp1))
			end if
			if tmp1<>"" then
				query="UPDATE categories SET " & tmp1 & " WHERE idcategory=" & tcheckcategory & ";"
				set rstemp1=connTemp.execute(query)
				set rstemp1=nothing
			end if
		end if
	end if
	set rstemp=nothing
	set rstemp1=nothing
	
	call closedb()
	
	call XMLUpdCatTier(tcheckcategory)

	XMLCheckCategory=tcheckcategory

End Function
	
Function XMLCheckTempCategory()
Dim query,rstemp,rstemp1,tmpIDTempCAT

	TempCategory="ImportedProducts"
	
	call opendb()
	
	query="Select idCategory from categories where categorydesc like '" & TempCategory & "' and idParentCategory=1"
	set rstemp=conntemp.execute(query)
	
	if rstemp.eof then
		imagename="no_image.gif"
		query="INSERT INTO categories (categorydesc,idParentCategory,image,largeimage) VALUES ('" & TempCategory & "',1,'" & imagename & "','" & imagename & "')"
		set rstemp1=conntemp.execute(query)
		query="SELECT idCategory FROM categories WHERE categorydesc like '" & TempCategory & "' AND idParentCategory=1"
		set rstemp1=conntemp.execute(query)
		tmpIDTempCAT=rstemp1("idCategory")
		
	else
		tmpIDTempCAT=rstemp("idCategory")
	end if
	set rstemp=nothing
	set rstemp1=nothing
	
	call closedb()
	
	call XMLUpdCatTier(tmpIDTempCAT)
	
	XMLCheckTempCategory=tmpIDTempCAT
	
End Function

Sub XMLAddUpdPrdCat(productID,catID)
Dim query,rs

	call opendb()
	
	query="SELECT idproduct FROM categories_products WHERE idproduct=" & productID & " AND idCategory=" & catID & ";"
	set rs=conntemp.execute(query)
	if rs.eof then 
		query="INSERT INTO categories_products (idProduct, idCategory) VALUES (" &productID& "," & catID & ");"
		set rs=conntemp.execute(query)
		BackupStr=BackupStr & "DELCATPRD" & chr(9) & prdID_value & chr(9) & pIdCategory & vbcrlf
	end if
	set rs=nothing
	
	call closedb()

End Sub

Sub XMLAddUpdCategories(requestType)
Dim attNode,subNode,ChildNodes,rNode
Dim tmpNodeName,tmpNodeValue
Dim pIdCategory,pidParentCategory

	if prdCategory_ex=1 then
		
		Set rNode=iRoot.selectNodes(cm_product & "/" & prdCategory_name)
		For Each attNode In rNode
			If attNode.Text<>"" then
				Set ChildNodes = attNode.childNodes
				catID_value=0
				catName_value=""
				catLDesc_value=""
				catSDesc_value=""
				catImg_value=""
				catLargeImg_value=""
				catParentID_value=0
				catParentName_value=""
			
				For Each subNode In ChildNodes
					tmpNodeName=subNode.NodeName
					tmpNodeValue=subNode.Text
					Select Case tmpNodeName
						Case catID_name:
							if tmpNodeValue<>"" then
								if IsNumeric(tmpNodeValue) then
									if clng(tmpNodeValue)>=0 then
										catID_ex=1
										catID_value=tmpNodeValue
									end if
								end if
							end if
						Case catName_name:
							if tmpNodeValue<>"" then
								catName_ex=1
								catName_value=getUserInputNew(tmpNodeValue,0)
							end if
						Case catLDesc_name:
							if tmpNodeValue<>"" then
								catLDesc_ex=1
								catLDesc_value=getUserInputNew(tmpNodeValue,0)
							end if
						Case catSDesc_name:
							if tmpNodeValue<>"" then
								catSDesc_ex=1
								catSDesc_value=getUserInputNew(tmpNodeValue,0)
							end if
						Case catImg_name:
							if tmpNodeValue<>"" then
								catImg_ex=1
								catImg_value=getUserInputNew(tmpNodeValue,0)
							end if
						Case catLargeImg_name:
							if tmpNodeValue<>"" then
								catLargeImg_ex=1
								catLargeImg_value=getUserInputNew(tmpNodeValue,0)
							end if
						Case catParentID_name:
							if tmpNodeValue<>"" then
								if IsNumeric(tmpNodeValue) then
									if clng(tmpNodeValue)>=0 then
										catParentID_ex=1
										catParentID_value=tmpNodeValue
									end if
								end if
							end if
						Case catParentName_name:
							if tmpNodeValue<>"" then
								catParentName_ex=1
								catParentName_value=getUserInputNew(tmpNodeValue,0)
							end if
					End Select
				Next
				
				if catParentID_value>0 OR catParentName_value<>"" then
					pidParentCategory=XMLCheckParent(catParentID_value,catParentName_value)
				else
					pidParentCategory=1
				end if
				
				pIdCategory=0
				
				if catID_value>0 OR catName_value<>"" then
					pIdCategory=XMLCheckCategory(catID_value,catName_value,pidParentCategory,catImg_value,catLargeImg_value,catSDesc_value,catLDesc_value)
				else
					if requestType=0 then
						pIdCategory=XMLCheckTempCategory()
						call XMLcreateError(119,cm_errorStr_119)
					end if
				end if
				
				if pIdCategory>0 then
					Call XMLAddUpdPrdCat(prdID_value,pIdCategory)
				end if
				
			End If
		Next
	else
		if requestType=0 then
			pIdCategory=XMLCheckTempCategory()
			call XMLcreateError(119,cm_errorStr_119)
			Call XMLAddUpdPrdCat(prdID_value,pIdCategory)
		end if
	end if
	
End Sub

Function XMLcheckOptGrp(GrpName)
Dim rstemp,query

	call opendb()

	query="SELECT idOptionGroup FROM optionsGroups WHERE OptionGroupDesc like '" & GrpName & "'"
	set rstemp=conntemp.execute(query)

	if rstemp.eof then
		query="insert into optionsGroups (OptionGroupDesc) values ('" & GrpName & "')"
		set rstemp=conntemp.execute(query)
		query="SELECT idOptionGroup FROM optionsGroups WHERE OptionGroupDesc='" & GrpName & "'"
		set rstemp=conntemp.execute(query)
		XMLcheckOptGrp=rstemp("idOptionGroup")
	else
		XMLcheckOptGrp=rstemp("idOptionGroup")
	end if

	set rstemp=nothing
	call closedb()

End Function

Function XMLcheckAttr(IDGrp,AttrName)
Dim IDOption
Dim rstemp,query

	call opendb()

	query="SELECT idOption FROM options WHERE optionDescrip like '" & AttrName & "'"
	set rstemp=conntemp.execute(query)

	if rstemp.eof then
		query="insert into options (optionDescrip) values ('" & AttrName & "')"
		set rstemp=conntemp.execute(query)
		query="SELECT idOption FROM options WHERE optionDescrip like '" & AttrName & "'"
		set rstemp=conntemp.execute(query)
		IDOption=rstemp("idOption")
	else
		IDOption=rstemp("idOption")
	end if
	set rstemp=nothing

	query="SELECT idoption FROM optGrps WHERE idoption=" & IDOption & " AND idOptionGroup=" & IDGrp
	set rstemp=connTemp.execute(query)

	if rstemp.eof then
		query="insert into optGrps (idOptionGroup,idoption) values (" & IDGrp & "," & IDOption & ")"
		set rstemp=conntemp.execute(query)
	end if
	set rstemp=nothing
	
	call closedb()
	XMLcheckAttr=IDOption
	
End Function

Sub XMLcheckPrdGrp(IDPrd,IDGrp,GrpReq,GrpOrder)
Dim query,rstemp

	call opendb()

	query="SELECT idOptionGroup FROM pcProductsOptions WHERE idProduct=" & IDPrd & " AND idOptionGroup=" & IDGrp & ";"
	set rstemp=conntemp.execute(query)

	if rstemp.eof then
		query="INSERT INTO pcProductsOptions (idProduct,idOptionGroup,pcProdOpt_Required,pcProdOpt_order) VALUES (" & IDPrd & "," & IDGrp & "," & GrpReq & "," & GrpOrder & ");"
		set rstemp=conntemp.execute(query)	
		BackupStr=BackupStr & "DELPRDGRP" & chr(9) & IDPrd & chr(9) & IDGrp & vbcrlf
		set rstemp=nothing
		call closedb()
	else
		Call XMLBackUpPrdData(4,IDGrp)
		call opendb()
		query="UPDATE pcProductsOptions SET idProduct=" & IDPrd & ",idOptionGroup=" & IDGrp & ",pcProdOpt_Required=" & GrpReq & ",pcProdOpt_order=" & GrpOrder & " WHERE idProduct=" & IDPrd & " AND idOptionGroup=" & IDGrp & ";"
		set rstemp=conntemp.execute(query)
		set rstemp=nothing
		call closedb()
	end if
	
End Sub

Sub XMLImUpOptGrp(IDPrd,IDGrp,IDOpt,OptPrice,OptWPrice,DOrder,InActive)
Dim rstemp,query,tmp_idoptoptgrp

	call opendb()
	
	query="SELECT idoptoptgrp FROM options_optionsGroups WHERE IDProduct=" & IDPrd & " AND idOptionGroup=" & IDGrp & " AND idOption=" & IDOpt & ";"
	set rstemp=connTemp.execute(query)

	if rstemp.eof then
		query="INSERT INTO options_optionsGroups (IDProduct,idOptionGroup,idOption,price,Wprice,sortOrder,InActive) VALUES (" & IDPrd & "," & IDGrp & "," & IDOpt & "," & OptPrice & "," & OptWPrice & "," & DOrder & "," & InActive & ")"
		set rstemp=connTemp.execute(query)
		BackupStr=BackupStr & "DELPRDOPT" & chr(9) & IDPrd & chr(9) & IDGrp & chr(9) & IDOpt & vbcrlf
		set rstemp=nothing
		call closedb()
	else
		tmp_idoptoptgrp=rstemp("idoptoptgrp")
		set rstemp=nothing
		Call XMLBackUpPrdData(3,tmp_idoptoptgrp)
		call opendb()
		query="UPDATE options_optionsGroups SET price=" & OptPrice & ",Wprice=" & OptWprice & ",sortOrder=" & DOrder & ",InActive=" & InActive & " WHERE idoptoptgrp=" & tmp_idoptoptgrp
		set rstemp=connTemp.execute(query)
		set rstemp=nothing
		call closedb()
	end if
	
End Sub

Sub XMLAddUpdOptGroup()
Dim attNode,subNode,ChildNodes,rNode,subNode1,childNotes1
Dim tmpNodeName,tmpNodeValue
Dim pcv_IDGrp1,pcv_IDOpt1
	
	Set rNode=iRoot.selectNodes(cm_product & "/" & prdOptionGroup_name)
	For Each attNode In rNode
		If attNode.Text<>"" then
			Set ChildNodes = attNode.childNodes
			groupRequired_value=0
			groupOrder_value=0
			groupName_value=""
			groupName_ex=0
			option_ex=0
		
			For Each subNode In ChildNodes
				tmpNodeName=subNode.NodeName
				tmpNodeValue=subNode.Text
				Select Case tmpNodeName
					Case groupName_name:
						if tmpNodeValue<>"" then
							groupName_ex=1
							groupName_value=tmpNodeValue
						end if
					Case groupRequired_name:
						if tmpNodeValue<>"" then
							if IsNumeric(tmpNodeValue) then
								groupRequired_ex=1
								groupRequired_value=tmpNodeValue
								if groupRequired_value<>0 then
									groupRequired_value=1
								end if
							end if
						end if
					Case groupOrder_name:
						if tmpNodeValue<>"" then
							if IsNumeric(tmpNodeValue) then
								groupOrder_ex=1
								groupOrder_value=tmpNodeValue
							end if
						end if
				End Select
			Next
			
			if groupName_ex=1 then
				pcv_IDGrp1=XMLcheckOptGrp(groupName_value)
				Call XMLcheckPrdGrp(prdID_value,pcv_IDGrp1,groupRequired_value,groupOrder_value)
			end if
			
			For Each subNode In ChildNodes
				tmpNodeName=subNode.NodeName
				tmpNodeValue=subNode.Text
				Select Case tmpNodeName
					Case option_name:
						if tmpNodeValue<>"" then
							Set ChildNodes1 = subNode.childNodes
							optName_ex=0
							optName_value=""
							optPrice_value=0
							optWPrice_value=0
							optOrder_value=0
							optInactive_value=0
							For Each subNode1 In ChildNodes1
								tmpNodeName=subNode1.NodeName
								tmpNodeValue=subNode1.Text
								Select Case tmpNodeName
									Case optName_name:
										if tmpNodeValue<>"" then
											optName_ex=1
											optName_value=getUserInputNew(tmpNodeValue,0)
										end if
									Case optPrice_name:
										if tmpNodeValue<>"" then
											if IsNumeric(tmpNodeValue) then
												optPrice_ex=1
												optPrice_value=tmpNodeValue
											end if
										end if
									Case optWPrice_name:
										if tmpNodeValue<>"" then
											if IsNumeric(tmpNodeValue) then
												optWPrice_ex=1
												optWPrice_value=tmpNodeValue
											end if
										end if
									Case optOrder_name:
										if tmpNodeValue<>"" then
											if IsNumeric(tmpNodeValue) then
												optOrder_ex=1
												optOrder_value=tmpNodeValue
											end if
										end if
									Case optInactive_name:
										if tmpNodeValue<>"" then
											if IsNumeric(tmpNodeValue) then
												optInactive_ex=1
												optInactive_value=tmpNodeValue
												if optInactive_value<>0 then
													optInactive_value=1
												end if
											end if
										end if
								End Select
							Next
							if optName_ex=1 then
								pcv_IDOpt1=XMLcheckAttr(pcv_IDGrp1,optName_value)
								Call XMLImUpOptGrp(prdID_value,pcv_IDGrp1,pcv_IDOpt1,optPrice_value,optWPrice_value,optOrder_value,optInactive_value)
							else
								call XMLcreateError(120,cm_errorStr_120 & groupName_value)
							end if
						end if
				End Select
			Next
			
		End if
	Next

End Sub

function XMLCheckPrdCustomField(cfid,cfname)
Dim query,rstemp,rstemp1
	if cfid<>"0" then
		query="SELECT idSearchField FROM pcSearchFields WHERE idSearchField=" & cfid & ";"
	else
		query="SELECT idSearchField FROM pcSearchFields WHERE pcSearchFieldName like '" & cfname & "'"
	end if
	set rstemp=conntemp.execute(query)
	
	if rstemp.eof then
		query="INSERT INTO pcSearchFields (pcSearchFieldName,pcSearchFieldShow,pcSearchFieldOrder,pcSearchFieldCPShow,pcSearchFieldSearch,pcSearchFieldCPSearch) VALUES ('" & cfname & "',1,0,1,1,1)"
		set rstemp1=conntemp.execute(query)
		query="SELECT idSearchField FROM pcSearchFields WHERE pcSearchFieldName like '" & cfname & "'"
		set rstemp1=conntemp.execute(query)
		XMLCheckPrdCustomField=rstemp1("idSearchField")
	else
		cfid=rstemp("idSearchField")
		if cfname<>"" then
			query="UPDATE pcSearchFields SET pcSearchFieldName='" & cfname & "' WHERE idSearchField=" & cfid & ";"
			set rstemp1=conntemp.execute(query)
		end if
		XMLCheckPrdCustomField=cfid
	end if
	set rstemp=nothing
	set rstemp1=nothing
end function

function XMLCheckPrdCustomFieldValue(idcustom,searchvalue)
Dim query,rstemp,rstemp1
	query="SELECT idSearchData FROM pcSearchData WHERE idSearchField=" & idcustom & " AND pcSearchDataName like '" & searchvalue & "'"
	set rstemp=conntemp.execute(query)
	
	if rstemp.eof then
		query="INSERT INTO pcSearchData (idSearchField,pcSearchDataName,pcSearchDataOrder) VALUES (" & idcustom & ",'" & searchvalue & "',0)"
		set rstemp1=conntemp.execute(query)
		query="SELECT idSearchData FROM pcSearchData WHERE idSearchField=" & idcustom & " AND pcSearchDataName like '" & searchvalue & "'"
		set rstemp1=conntemp.execute(query)
		XMLCheckPrdCustomFieldValue=rstemp1("idSearchData")
	else
		XMLCheckPrdCustomFieldValue=rstemp("idSearchData")
	end if
	set rstemp=nothing
	set rstemp1=nothing
end function

Sub XMLAddUpdPrdCF()
Dim attNode,subNode,ChildNodes,rNode
Dim CFCount,tmpNodeName,tmpNodeValue
Dim tmp1,query,rstemp
	
	CFCount=0
	tmp1=""
	
	call opendb()
	
	Set rNode=iRoot.selectNodes(cm_product & "/" & prdCustomField_name)
	For Each attNode In rNode
		If attNode.Text<>"" then
			Set ChildNodes = attNode.childNodes
			cfName_ex=0
			cfValue_ex=0
			cfID_value=0
			cfName_value=""
			cfValue_value=""
		
			For Each subNode In ChildNodes
				tmpNodeName=subNode.NodeName
				tmpNodeValue=subNode.Text
				Select Case tmpNodeName
					Case cfID_name:
						if tmpNodeValue<>"" then
							if IsNumeric(tmpNodeValue) AND tmpNodeValue>"0" then
								cfID_ex=1
								cfID_value=tmpNodeValue
							end if
						end if
					Case cfName_name:
						cfName_ex=1
						cfName_value=getUserInputNew(tmpNodeValue,0)
					Case cfValue_name:
						cfValue_ex=1
						cfValue_value=getUserInputNew(tmpNodeValue,0)
				End Select
			Next
				
			if ((cfID_ex=1) OR (cfName_ex=1)) AND (cfValue_ex=1) then
				cfID_value=XMLCheckPrdCustomField(cfID_value,cfName_value)
				cfValueID_value=XMLCheckPrdCustomFieldValue(cfID_value,cfValue_value)
				if (cfID_value>"0") AND (cfValueID_value>"0") then
					query="DELETE FROM pcSearchFields_Products WHERE idproduct=" & prdID_value & " AND idSearchData IN (SELECT idSearchData FROM pcSearchData WHERE idSearchField=" & cfID_value & ");"
					Set rstemp=conntemp.execute(query)
					Set rstemp=nothing

					query="INSERT INTO pcSearchFields_Products (idproduct,idSearchData) VALUES (" & prdID_value & "," & cfValueID_value & ");"
					Set rstemp=conntemp.execute(query)
					Set rstemp=nothing
				end if
			end if
			
		End if
	Next
	
	call closedb()

End Sub

Sub XMLUpdDownloadInfor()
Dim query,rstemp
	
	If prdDownloadable_value=0 then
		Call XMLBackUpPrdData(1,0)
		call opendb()
		query="DELETE FROM DProducts WHERE idproduct=" & prdID_value & ";"
		set rstemp=connTemp.execute(query)
		set rstemp=nothing
		call closedb()
	Else
		If prdDownloadable_value=1 AND prdDownloadInfo_ex=1 then
			Call XMLBackUpPrdData(1,0)
			call opendb()
			query="DELETE FROM DProducts WHERE idproduct=" & prdID_value & ";"
			set rstemp=connTemp.execute(query)
			set rstemp=nothing
					
			query="INSERT INTO DProducts (IdProduct,ProductURL,URLExpire,ExpireDays,License,LocalLG,RemoteLG,LicenseLabel1,LicenseLabel2,LicenseLabel3,LicenseLabel4,LicenseLabel5,AddToMail) values (" & prdID_value & ",'" & diFileLocation_value & "'," & diURLExpire_value & "," & diURLExpDays_value & "," & diUseLicenseGen_value & ",'" & diLocalGen_value & "','" & diRemoteGen_value & "','" & diLFLabel1_value & "','" & diLFLabel2_value & "','" & diLFLabel3_value & "','" & diLFLabel4_value & "','" & diLFLabel5_value & "','" & diAddMsg_value & "')"
			query=replace(query,chr(34),"&quot;")
			set rstemp=conntemp.execute(query)
			set rstemp=nothing
			call closedb()
		end if
	End if
	
End Sub

Sub XMLUpdGCInfor()
Dim query,rstemp
	
	If prdGiftCertificate_value=0 then
		Call XMLBackUpPrdData(2,0)
		call opendb()
		query="DELETE FROM pcGC WHERE pcGC_IDProduct=" & prdID_value & ";"
		set rstemp=connTemp.execute(query)
		set rstemp=nothing
		call closedb()
	Else
		If prdGiftCertificate_value=1 AND prdGCInfo_ex=1 then
			Call XMLBackUpPrdData(2,0)
			call opendb()
			query="DELETE FROM pcGC WHERE pcGC_IDProduct=" & prdID_value & ";"
			set rstemp=connTemp.execute(query)
			set rstemp=nothing
			if giExpDate_value<>"" then
				if SQL_Format="1" then
					giExpDate_value=(day(giExpDate_value)&"/"&month(giExpDate_value)&"/"&year(giExpDate_value))
				else
					giExpDate_value=(month(giExpDate_value)&"/"&day(giExpDate_value)&"/"&year(giExpDate_value))
				end if
			end if
			if scDB="SQL" then
				query="INSERT INTO pcGC (pcGC_IDProduct,pcGC_Exp,pcGC_EOnly,pcGC_CodeGen,pcGC_ExpDate,pcGC_ExpDays,pcGC_GenFile) values (" & prdID_value & "," & giExpire_value & "," & giEOnly_value & "," & giUseGen_value & ",'" & giExpDate_value & "'," & giExpNDays_value & ",'" & giCustomGen_value & "');"
			else
				query="INSERT INTO pcGC (pcGC_IDProduct,pcGC_Exp,pcGC_EOnly,pcGC_CodeGen,pcGC_ExpDate,pcGC_ExpDays,pcGC_GenFile) values (" & prdID_value & "," & giExpire_value & "," & giEOnly_value & "," & giUseGen_value & ",#" & giExpDate_value & "#," & giExpNDays_value & ",'" & giCustomGen_value & "');"
			end if
			query=replace(query,chr(34),"&quot;")
			set rstemp=conntemp.execute(query)
			set rstemp=nothing
			call closedb()
		end if
	End if
	
End Sub

Sub XMLCheckIsDropShipper()
Dim query,rs1

	call opendb()

	if prdSupplierID_value>0 then
		query="SELECT pcSupplier_IsDropShipper FROM pcSuppliers WHERE pcSupplier_ID=" & prdSupplierID_value & ";"
		set rs1=conntemp.execute(query)
		if not rs1.eof then
			prdIsDropShipper_value=rs1("pcSupplier_IsDropShipper")
			if prdIsDropShipper_value="1" then
				prdDropShipperID_value=prdSupplierID_value
			end if
		end if
		set rs1=nothing
	end if

	call closedb()
	
End Sub

Function XMLCheckBrand()
Dim query,rstemp,rstemp1

	call opendb()
	
	if CLng(brandID_value)>=0 then
		query="Select idBrand from Brands where idBrand=" & brandID_value & ";"
	else
		query="Select idBrand from Brands where BrandName like '" & brandName_value & "'"
	end if
	set rstemp=conntemp.execute(query)
	
	if rstemp.eof then
		if brandName_value<>"" then
			if brandLogo_value="" then
				brandLogo_value="no_image.gif"
			end if
			query="insert into Brands (BrandName,BrandLogo) values ('" & brandName_value & "','" & brandLogo_value & "')"
			set rstemp1=conntemp.execute(query)
			set rstemp1=nothing
			query="Select idBrand from Brands where BrandName='" & brandName_value & "';"
			set rstemp1=conntemp.execute(query)
			brandID_value=rstemp1("IDBrand")
			set rstemp1=nothing
		else
			brandID_value=0
		end if
	else
		brandID_value=rstemp("IDBrand")
		set rstemp=nothing
		if brandName_value & brandLogo_value<>"" then
			query="UPDATE Brands SET BrandName='" & brandName_value & "',BrandLogo='" & brandLogo_value & "' WHERE idBrand=" & brandID_value & ";"
			set rstemp1=conntemp.execute(query)
			set rstemp1=nothing
		end if
	end if
	set rstemp=nothing
	
	XMLCheckBrand=brandID_value
	call closedb()
	
End Function

Sub RunAddProduct()
Dim tmp_IsBTO,tmp_IsItem,dtTodaysDate,tmp_str1,tmp_str2,unitslb
Dim query,rstemp

	tmp_IsBTO=0
	tmp_IsItem=0
	Select Case prdType_value
		Case 1:	tmp_IsBTO=-1
		Case 2: tmp_IsItem=-1
	End Select
	
	if prdBrand_ex=1 then
		brandID_value=XMLCheckBrand()
	end if
	Call XMLCheckIsDropShipper()

	dtTodaysDate=Date()
	if SQL_Format="1" then
		dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
	else
		dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
	end if
	
	'BTO-S
		tmp_str1=""
		tmp_str2=""
		if scBTO=1 then
			tmp_str1=tmp_str1 & ",pcprod_hidebtoprice"
			tmp_str2=tmp_str2 & "," & prdHideBTOPrices_value 
			tmp_str1=tmp_str1 & ",pcprod_HideDefConfig"
			tmp_str2=tmp_str2 & "," & prdHideDefaultConfig_value
			tmp_str1=tmp_str1 & ",NoPrices"
			tmp_str2=tmp_str2 & "," & prdDisallowPurchasing_value
			tmp_str1=tmp_str1 & ",pcProd_SkipDetailsPage"
			tmp_str2=tmp_str2 & "," & prdSkipPrdPage_value 
		end if
	'BTO-E
	
	call opendb()
	
	unitslb=UnitsToPound_value
	
	if prdOversize_ex=1 AND prdOversize_value<>"NO" AND prdWeight_ex=0 then
		prdOversize_value=prdOversize_value & "0"
	end if
	
	if scDB="SQL" then
		query="INSERT INTO products (IDBrand,sku, description, details, price, listPrice, bToBPrice, imageUrl, listhidden, hotDeal,iRewardPoints, weight, stock, active,showInHome, idSupplier, smallImageUrl,largeImageUrl, notax, noshipping,formquantity,emailtext,sDesc,nostock,noshippingtext, pcprod_EnteredOn,pcprod_qtyvalidate,pcprod_minimumqty,cost,pcProd_BackOrder,pcProd_ShipNDays,pcProd_NotifyStock,pcProd_ReorderLevel,pcSupplier_ID,pcProd_IsDropShipped,pcDropShipper_ID,pcprod_GC,pcProd_MetaTitle,pcProd_MetaDesc,pcProd_MetaKeywords" & tmp_str1 & ",pcprod_QtyToPound,serviceSpec,configOnly,downloadable,OverSizeSpec) VALUES (" & brandID_value & ",'" &prdSKU_value& "','" &prdName_value& "','" & prdDesc_value & "'," &prdPrice_value& "," &prdListPrice_value& "," &prdWPrice_value& ",'" &prdImg_value& "'," & prdShowSavings_value & "," & prdSpecial_value & "," & prdRewardPoints_value & "," &prdWeight_value& "," &prdStock_value& "," &prdActive_value& "," & prdFeatured_value & ",10,'" &prdSmallImg_value& "','"&prdLargeImg_value&"',"&prdNoTax_value&","&prdNoShippingCharge_value&"," & prdNotForSale_value & ",'" & prdNotForSaleMsg_value & "','" & prdSDesc_value & "'," & prdDisregardStock_value & "," & prdDisplayNoShipText_value & ",'"&dtTodaysDate&"'," & prdPurchaseMulti_value & "," & prdMinimumQty_value & "," & prdCost_value & "," & prdBackOrder_value & "," & prdShipNDays_value & "," & prdLowStockNotice_value & "," & prdReorderLevel_value & "," & prdSupplierID_value & "," & prdIsDropShipped_value & "," & prdDropShipperID_value & "," & prdGiftCertificate_value & ",'" & mtTitle_value & "','" & mtDesc_value & "','" & mtKeywords_value & "'" & tmp_str2 & "," & unitslb & "," & tmp_IsBTO & "," & tmp_IsItem & "," & prdDownloadable_value & ",'" & prdOversize_value & "')"
	else
		query="INSERT INTO products (IDBrand,sku, description, details, price, listPrice, bToBPrice, imageUrl, listhidden, hotDeal,iRewardPoints, weight, stock, active,showInHome, idSupplier, smallImageUrl,largeImageUrl, notax, noshipping,formquantity,emailtext,sDesc,nostock,noshippingtext, pcprod_EnteredOn,pcprod_qtyvalidate,pcprod_minimumqty,cost,pcProd_BackOrder,pcProd_ShipNDays,pcProd_NotifyStock,pcProd_ReorderLevel,pcSupplier_ID,pcProd_IsDropShipped,pcDropShipper_ID,pcprod_GC,pcProd_MetaTitle,pcProd_MetaDesc,pcProd_MetaKeywords" & tmp_str1 & ",pcprod_QtyToPound,serviceSpec,configOnly,downloadable,OverSizeSpec) VALUES (" & brandID_value & ",'" &prdSKU_value& "','" &prdName_value& "','" & prdDesc_value & "'," &prdPrice_value& "," &prdListPrice_value& "," &prdWPrice_value& ",'" &prdImg_value& "'," & prdShowSavings_value & "," & prdSpecial_value & "," & prdRewardPoints_value & "," &prdWeight_value& "," &prdStock_value& "," &prdActive_value& "," & prdFeatured_value & ",10,'" &prdSmallImg_value& "','"&prdLargeImg_value&"',"&prdNoTax_value&","&prdNoShippingCharge_value&"," & prdNotForSale_value & ",'" & prdNotForSaleMsg_value & "','" & prdSDesc_value & "'," & prdDisregardStock_value & "," & prdDisplayNoShipText_value & ",#"&dtTodaysDate&"#," & prdPurchaseMulti_value & "," & prdMinimumQty_value & "," & prdCost_value & "," & prdBackOrder_value & "," & prdShipNDays_value & "," & prdLowStockNotice_value & "," & prdReorderLevel_value & "," & prdSupplierID_value & "," & prdIsDropShipped_value & "," & prdDropShipperID_value & "," & prdGiftCertificate_value & ",'" & mtTitle_value & "','" & mtDesc_value & "','" & mtKeywords_value & "'" & tmp_str2 & "," & unitslb & "," & tmp_IsBTO & "," & tmp_IsItem & "," & prdDownloadable_value & ",'" & prdOversize_value & "')"
	end if
	query=replace(query,chr(34),"&quot;")
	set rstemp=conntemp.execute(query)
	
	query="SELECT idproduct FROM Products WHERE sku like '" & prdSKU_value & "' ORDER BY idproduct DESC;"
	set rstemp=connTemp.execute(query)
	if not rstemp.eof then
		prdID_value=rstemp("idproduct")
	end if
	set rstemp=nothing
		
	call closedb()
	
	if prdDownloadable_ex=1 then
		call XMLUpdDownloadInfor()
	end if
	
	if prdGiftCertificate_ex=1 then
		call XMLUpdGCInfor()
	end if
	
	if prdCustomField_ex=1 then
		Call XMLAddUpdPrdCF()
	end if
	
	call opendb()
	if (clng(prdSupplierID_value)>0) OR (clng(prdDropShipperID_value)>0) then
		query="DELETE FROM pcDropShippersSuppliers WHERE idproduct=" & prdID_value
		set rstemp=connTemp.execute(query)
		set rstemp=nothing
		query="INSERT INTO pcDropShippersSuppliers (idproduct,pcDS_IsDropShipper) VALUES (" & prdID_value & "," & prdIsDropShipper_value & ")"
		set rstemp=connTemp.execute(query)
		set rstemp=nothing
	end if
	call closedb()
	
	if prdOptionGroup_ex=1 then
		Call XMLAddUpdOptGroup()
	end if
	
	Call XMLAddUpdCategories(0)
	
	BackupStr="DELPRD" & chr(9) & prdID_value & vbcrlf
	
End Sub

Sub RunUpdProduct()
Dim tmp_IsBTO,tmp_IsItem,tmp1
Dim query,rstemp

	tmp_IsBTO=0
	tmp_IsItem=0
	Select Case prdType_value
		Case 1:	tmp_IsBTO=-1
		Case 2: tmp_IsItem=-1
	End Select
	if prdBrand_ex=1 then
		brandID_value=XMLCheckBrand()
	end if
	Call XMLCheckIsDropShipper()
	
	Call XMLBackUpPrdData(0,0)

	tmp1=""
	if prdBrand_ex=1 then
		tmp1=tmp1 & ",IDBrand=" & brandID_value
	end if
	if prdListPrice_ex=1 then
		tmp1=tmp1 & ",listPrice=" & prdListPrice_value
	end if
	if prdWPrice_ex=1 then
		tmp1=tmp1 & ", bToBPrice=" & prdWPrice_value
	end if
	if prdWeight_ex=1 then
		tmp1=tmp1 & ", weight=" & prdWeight_value
	end if
	if UnitsToPound_ex=1 then
		tmp1=tmp1 & ", pcprod_QtyToPound=" & UnitsToPound_value
	end if
	if prdStock_ex=1 then
		tmp1=tmp1 & ", stock=" & prdStock_value
	end if
	if prdSmallImg_ex=1 then
		tmp1=tmp1 & ",smallImageUrl='" & prdSmallImg_value & "'"
	end if
	if prdImg_ex=1 then
		tmp1=tmp1 & ", imageUrl='" & prdImg_value & "'"
	end if
	if prdLargeImg_ex=1 then
		tmp1=tmp1 & ",largeImageUrl='" & prdLargeImg_value & "'"
	end if
	if prdActive_ex=1 then
		tmp1=tmp1 & ", active = " & prdActive_value
	'else
	'	tmp1=tmp1 & ", active = -1"
	end if
	if prdFeatured_ex=1 then
		tmp1=tmp1 & ", showInHome = " & prdFeatured_value
	end if
	
	if mtTitle_ex=1 then
		tmp1=tmp1 & ", pcProd_MetaTitle = '" & mtTitle_value & "'"
	end if
	
	if mtDesc_ex=1 then
		tmp1=tmp1 & ", pcProd_MetaDesc = '" & mtDesc_value & "'"
	end if
	
	if mtKeywords_ex=1 then
		tmp1=tmp1 & ", pcProd_MetaKeywords = '" & mtKeywords_value & "'"
	end if
	
	'BTO-S
	if scBTO=1 then
	
		if prdHideBTOPrices_ex=1 then
			tmp1=tmp1 & ", pcprod_hidebtoprice = " & prdHideBTOPrices_value 
		end if
		if prdHideDefaultConfig_ex=1 then
			tmp1=tmp1 & ", pcprod_HideDefConfig = " & prdHideDefaultConfig_value
		end if
		if prdDisallowPurchasing_ex=1 then
			tmp1=tmp1 & ", NoPrices = " & prdDisallowPurchasing_value
		end if
		if prdSkipPrdPage_ex=1 then
			tmp1=tmp1 & ", pcProd_SkipDetailsPage = " & prdSkipPrdPage_value
		end if
	
	end if
	'BTO-E
	
	if prdGiftCertificate_ex=1 then
		tmp1=tmp1 & ", pcprod_GC = " & prdGiftCertificate_value 
	end if
	
	'Start SDBA
	if prdCost_ex=1 then
		tmp1=tmp1 & ", cost = " & prdCost_value 
	end if
	
	if prdBackOrder_ex=1 then
		tmp1=tmp1 & ", pcProd_BackOrder = " & prdBackOrder_value
	end if

	if prdShipNDays_ex=1 then
		tmp1=tmp1 & ", pcProd_ShipNDays = " & prdShipNDays_value 
	end if
	
	if prdLowStockNotice_ex=1 then
		tmp1=tmp1 & ", pcProd_NotifyStock = " & prdLowStockNotice_value 
	end if
	
	if prdReorderLevel_ex=1 then
		tmp1=tmp1 & ", pcProd_ReorderLevel = " & prdReorderLevel_value 
	end if
	
	if prdIsDropShipped_ex=1 then
		tmp1=tmp1 & ", pcProd_IsDropShipped = " & prdIsDropShipped_value 
	end if
	
	if prdSupplierID_ex=1 then
		tmp1=tmp1 & ", pcSupplier_ID = " & prdSupplierID_value 
	end if
	
	if (prdDropShipperID_ex=1) or (prdIsDropShipper_value="1") then
		tmp1=tmp1 & ", pcDropShipper_ID = " & prdDropShipperID_value 
	end if
	'End SDBA
	
	if prdShowSavings_ex=1 then
		tmp1=tmp1 & ", listhidden=" & prdShowSavings_value
	end if
	if prdSpecial_ex=1 then
		tmp1=tmp1 & ", hotDeal=" & prdSpecial_value
	end if
	if prdRewardPoints_ex=1 then
		tmp1=tmp1 & ",iRewardPoints=" & prdRewardPoints_value
	end if
	if prdNoTax_ex=1 then
		tmp1=tmp1 & ",notax=" & prdNoTax_value
	end if
	if prdNoShippingCharge_ex=1 then
		tmp1=tmp1 & ", noshipping=" & prdNoShippingCharge_value
	end if
	if prdNotForSale_ex=1 then
		tmp1=tmp1 & ",formquantity=" & prdNotForSale_value
	end if
	if prdNotForSaleMsg_ex=1 then
		tmp1=tmp1 & ",emailtext='" & prdNotForSaleMsg_value & "'"
	end if
	if prdName_ex=1 then
		tmp1=tmp1 & ",description='" & prdName_value & "'"
	end if
	
	if prdDesc_ex=1 then
		tmp1=tmp1 & ", details='" & prdDesc_value & "'"
	end if
	
	if prdSDesc_ex=1 then
		tmp1=tmp1 & ", sDesc='" & prdSDesc_value & "'"
	end if
		
	if prdPrice_ex=1 then
		tmp1=tmp1 & ", price=" & prdPrice_value
	end if
	
	if prdDisregardStock_ex=1 then
		tmp1=tmp1 & ", noStock=" & prdDisregardStock_value
	end if
	
	if prdDisplayNoShipText_ex=1 then
		tmp1=tmp1 & ", noshippingtext=" & prdDisplayNoShipText_value
	end if
	
	if prdMinimumQty_ex=1 then
		tmp1=tmp1 & ",pcprod_minimumqty=" & prdMinimumQty_value
	end if
	
	if prdPurchaseMulti_ex=1 then
		tmp1=tmp1 & ",pcprod_qtyvalidate=" & prdPurchaseMulti_value
	end if
	
	if prdDownloadable_ex=1 then
		tmp1=tmp1 & ",Downloadable=" & prdDownloadable_value
	end if
	
	if prdType_ex=1 then
		tmp1=tmp1 & ",serviceSpec=" & tmp_IsBTO & ",configOnly=" & tmp_IsItem
	end if
	
	if prdCreatedDate_ex=1 then
		dtTodaysDate=prdCreatedDate_value
		if SQL_Format="1" then
			dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
		else
			dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
		end if
		if scDB="SQL" then
			tmp1=tmp1 & ",pcprod_EnteredOn='" & dtTodaysDate & "'"
		else
			tmp1=tmp1 & ",pcprod_EnteredOn=#" & dtTodaysDate & "#"
		end if
	end if
	
	call opendb()
	
	if prdID_value<>"" then
		query="update products set removed=0" & tmp1 & " where idproduct=" & prdID_value & ";"
	end if

	query=replace(query,chr(34),"&quot;")
	set rstemp=conntemp.execute(query)
	set rstemp=nothing
	
	if prdID_ex=0 then
		query="SELECT idproduct FROM Products WHERE sku like '" & prdSKU_value & "' ORDER BY idproduct DESC;"
		set rstemp=connTemp.execute(query)
		if not rstemp.eof then
			prdID_value=rstemp("idproduct")
		end if
		set rstemp=nothing
	end if
	
	if prdOversize_ex=1 AND prdOversize_value<>"NO" AND prdWeight_ex=0 then
		query="SELECT weight FROM Products WHERE idproduct=" & prdID_value & ";"
		set rstemp=connTemp.execute(query)
		if not rstemp.eof then
			tmpWeightvalue=rstemp("weight")
		end if
		set rstemp=nothing
		prdOversize_value=prdOversize_value & tmpWeightvalue
	end if
	
	if prdOversize_ex=1 then
		query="UPDATE products SET OverSizeSpec=" & prdOversize_value & " WHERE idproduct=" & prdID_value & ";"
		set rstemp=connTemp.execute(query)
		set rstemp=nothing
	end if
		
	call closedb()
	
	if prdDownloadable_ex=1 then
		call XMLUpdDownloadInfor()
	end if
	
	if prdGiftCertificate_ex=1 then
		call XMLUpdGCInfor()
	end if
	
	if prdCustomField_ex=1 then
		Call XMLAddUpdPrdCF()
	end if
	
	call opendb()
	if (clng(prdSupplierID_value)>0) OR (clng(prdDropShipperID_value)>0) then
		query="DELETE FROM pcDropShippersSuppliers WHERE idproduct=" & prdID_value
		set rstemp=connTemp.execute(query)
		set rstemp=nothing
		query="INSERT INTO pcDropShippersSuppliers (idproduct,pcDS_IsDropShipper) VALUES (" & prdID_value & "," & prdIsDropShipper_value & ")"
		set rstemp=connTemp.execute(query)
		set rstemp=nothing
	end if
	call closedb()
	
	if prdOptionGroup_ex=1 then
		Call XMLAddUpdOptGroup()
	end if
	
	Call XMLAddUpdCategories(1)
	
End Sub

Sub RunAddUpdProduct(requestType)
Dim query,rs,tmpBackUp
Dim requestKey,fso,afi
	
	call opendb()
	if requestType=0 then
		query="SELECT idproduct FROM Products WHERE sku like '" & prdSKU_value & "' AND removed=0;"
		set rs=connTemp.execute(query)
		
		if not rs.eof then
			set rs=nothing
			call closedb()
			call XMLcreateError(118,cm_errorStr_118 & cm_errorStr_118b & prdSKU_value & cm_errorStr_118c)
			call returnXML()
		end if
		set rs=nothing
	else	
		
		if ImportField_value = 1 then
			query="SELECT idproduct FROM Products WHERE sku like '" & prdSKU_value & "' AND removed=0;"			
			set rstemp=conntemp.execute(query)			
			if rstemp.eof then
				set rstemp=nothing
				call closedb()
				call XMLcreateError(118,cm_errorStr_118 & cm_errorStr_118b & prdSKU_value & cm_errorStr_118d)
				call returnXML()
			else
				prdID_value=rstemp("idproduct")
			end if
			set rstemp=nothing	
		end if	
			
		if prdID_value > 0 then
			query="SELECT idproduct FROM Products WHERE idproduct=" & prdID_value & " AND removed=0;"
			set rs=connTemp.execute(query)
		
			if rs.eof then
				set rs=nothing
				call closedb()
				call XMLcreateError(118,cm_errorStr_118 & cm_errorStr_118a & prdID_value & cm_errorStr_118d)
				call returnXML()
			end if
			set rs=nothing
		end if		
		
	end if		

	query="SELECT Products.sku FROM Products WHERE idproduct=" & prdID_value & " AND removed=0;"
	set rs=connTemp.execute(query)		
	if not rs.eof then
		prdSKU_value=rs("sku")
	end if
	set rs=nothing
	call closedb()	
	
	if requestType=0 then
		Call RunAddProduct()
	else
		Call RunUpdProduct()
	end if
	
	if BackupStr<>"" then
		tmpBackUp=1
	else
		tmpBackUp=0
	end if
	
	IF cm_LogTurnOn=1 THEN
		if requestType=0 then
			requestKey=CreateRequestRecord(pcv_PartnerID,9,prdID_value,tmpBackUp,0,0,0,0)
			cm_requestKey_value=requestKey
		else
			requestKey=CreateRequestRecord(pcv_PartnerID,11,prdID_value,tmpBackUp,0,0,0,0)
			cm_requestKey_value=requestKey
		end if
	
		Set tmpNode=oXML.createNode(1,cm_requestKey_name,"")
		tmpNode.Text=requestKey
		oRoot.appendChild(tmpNode)
	END IF
	
	if xmlHaveErrors=0 then
		Set tmpNode=oXML.createNode(1,cm_requestStatus_name,"")
		tmpNode.Text=cm_SuccessCode
		oRoot.appendChild(tmpNode)
	else
		oRoot.selectSingleNode(cm_requestStatus_name).Text=cm_HalfSuccessCode
	end if

	Set tmpNode=oXML.createNode(1,prdID_name,"")
	tmpNode.Text=prdID_value
	oRoot.appendChild(tmpNode)
	
	Set tmpNode=oXML.createNode(1,prdSKU_name,"")
	tmpNode.Text=prdSKU_value
	oRoot.appendChild(tmpNode)
	
	if BackupStr<>"" then
		Set fso=Server.CreateObject("Scripting.FileSystemObject")
		Set afi=fso.CreateTextFile(server.MapPath(".") & "\logs\" & requestKey & ".txt",True)
		afi.Write(BackupStr)
		afi.Close
		Set afi=nothing
		Set fso=nothing
	end if
	
End Sub

%>

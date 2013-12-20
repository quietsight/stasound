<%  
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

Dim pcv_intOptionsExist


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Are Options
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function pcf_CheckForOptions(ID)	
	'//Product ID
	pcv_str_CheckOptionsID=ID
	'// CHECK FOR OPTIONS
	' TABLES: products, pcProductsOptions, optionsgroups, ptions_optionsGroups
	query = 		"SELECT DISTINCT optionsGroups.OptionGroupDesc, pcProductsOptions.pcProdOpt_ID, pcProductsOptions.idOptionGroup, pcProductsOptions.pcProdOpt_Required, pcProductsOptions.pcProdOpt_Order "
	query = query & "FROM products "
	query = query & "INNER JOIN ( "
	query = query & "pcProductsOptions INNER JOIN ( "
	query = query & "optionsgroups "
	query = query & "INNER JOIN options_optionsGroups "
	query = query & "ON optionsgroups.idOptionGroup = options_optionsGroups.idOptionGroup "
	query = query & ") ON optionsGroups.idOptionGroup = pcProductsOptions.idOptionGroup "
	query = query & ") ON products.idProduct = pcProductsOptions.idProduct "
	query = query & "WHERE products.idProduct=" & pcv_str_CheckOptionsID &" "
	query = query & "AND options_optionsGroups.idProduct=" & pcv_str_CheckOptionsID &" "
	query = query & "ORDER BY pcProductsOptions.pcProdOpt_Order;"
	set rsCheckOptions=server.createobject("adodb.recordset")
	set rsCheckOptions=conntemp.execute(query)	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsCheckOptions=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	If NOT rsCheckOptions.eof Then
		pcf_CheckForOptions=1
	Else
		pcf_CheckForOptions=2
	End If	
	
	set rsCheckOptions = nothing
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Are Options
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Required Options
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function pcf_CheckForReqOptions(ID)	
	'//Product ID
	pcv_str_CheckOptionsID=ID
	'// CHECK FOR REQUIRED OPTIONS
	' TABLES: products, pcProductsOptions, optionsgroups, ptions_optionsGroups
	query = 		"SELECT DISTINCT optionsGroups.OptionGroupDesc, pcProductsOptions.pcProdOpt_ID, pcProductsOptions.idOptionGroup, pcProductsOptions.pcProdOpt_Required, pcProductsOptions.pcProdOpt_Order "
	query = query & "FROM products "
	query = query & "INNER JOIN ( "
	query = query & "pcProductsOptions INNER JOIN ( "
	query = query & "optionsgroups "
	query = query & "INNER JOIN options_optionsGroups "
	query = query & "ON optionsgroups.idOptionGroup = options_optionsGroups.idOptionGroup "
	query = query & ") ON optionsGroups.idOptionGroup = pcProductsOptions.idOptionGroup "
	query = query & ") ON products.idProduct = pcProductsOptions.idProduct "
	query = query & "WHERE products.idProduct=" & pcv_str_CheckOptionsID &" "
	query = query & "AND options_optionsGroups.idProduct=" & pcv_str_CheckOptionsID &" "
	query = query & "ORDER BY pcProductsOptions.pcProdOpt_Order;"
	set rsCheckOptions=server.createobject("adodb.recordset")
	set rsCheckOptions=conntemp.execute(query)	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsCheckOptions=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	If NOT rsCheckOptions.eof Then
		pcf_CheckForReqOptions=2
		do while not rsCheckOptions.eof
			if clng(rsCheckOptions("pcProdOpt_Required"))<>0 then
				pcf_CheckForReqOptions=1
				Exit Function
			end if
			rsCheckOptions.MoveNext
		loop
	Else
		pcf_CheckForReqOptions=2
	End If	
	
	set rsCheckOptions = nothing
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Required Options
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Required Input Fields
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function pcf_CheckForReqInputFields(ID)	
	'//Product ID
	pcv_str_CheckID=ID
	'// CHECK FOR REQUIRED INPUT FIELDS
	' TABLES: products
	query ="SELECT x1req,x2req,x3req FROM products WHERE idProduct=" & pcv_str_CheckID
	set rsCheckInput=server.createobject("adodb.recordset")
	set rsCheckInput=conntemp.execute(query)	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsCheckInput=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	If NOT rsCheckInput.eof Then
		pcf_CheckForReqInputFields=2
		if clng(rsCheckInput("x1req"))+clng(rsCheckInput("x2req"))+clng(rsCheckInput("x3req"))<>0 then
			pcf_CheckForReqInputFields=1
		end if
	Else
		pcf_CheckForReqInputFields=2
	End If	
	
	set rsCheckInput = nothing
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Required Input Fields
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Required Accessories
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Function pcf_CheckForReqAccessories(ID)
	'//Product ID
	IdProduct=ID
	'// No Required Accessories
	pcf_CheckForReqAccessories=2
	query="SELECT cs_relationships.idproduct, cs_relationships.idrelation, cs_relationships.cs_type, cs_relationships.isRequired, products.servicespec, products.description FROM cs_relationships INNER JOIN products ON cs_relationships.idrelation=products.idProduct WHERE (((cs_relationships.idproduct)="& IdProduct &") AND ((products.active)=-1) AND ((products.removed)=0)) ORDER BY cs_relationships.num,cs_relationships.idrelation;"
	set rsCheckRequiredAccessories=server.createobject("adodb.recordset")
	set rsCheckRequiredAccessories=conntemp.execute(query)	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsCheckInput=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if	
	
	If NOT rsCheckRequiredAccessories.EOF Then
		pcArray_RequiredAccessories=rsCheckRequiredAccessories.GetRows()
		pcv_intAccessoriesCount=UBound(pcArray_RequiredAccessories,2)+1
	End If
	set rsCheckRequiredAccessories=nothing
	
	AccessoriesCnt=Cint(0)	
	do while (AccessoriesCnt < pcv_intAccessoriesCount)
	
		cs_pBTOCnt=Cint(0)
		cs_pOptCnt=Cint(0)
		cs_pAddtoCart=Cint(0)
		
		pidrelation=pcArray_RequiredAccessories(1,AccessoriesCnt) '// rsCheckRequiredAccessories("idrelation")
		pcsType=pcArray_RequiredAccessories(2,AccessoriesCnt) '// rsCheckRequiredAccessories("cs_type")	
		cs_pserviceSpec=pcArray_RequiredAccessories(4,AccessoriesCnt) '// rsCheckRequiredAccessories("servicespec")
		pcv_strIsRequiredAccessory=pcArray_RequiredAccessories(3,AccessoriesCnt) '// rsCheckRequiredAccessories("isRequired")
		
		AccessoriesCnt=AccessoriesCnt+1

		'// CHECK IF THIS PRODUCT HAS AT LEAST ONE ACTIVE CATEGORY									
		pcv_intCategoryActive=1									
		If Session("customerType")=1 Then
			pcv_strCSTemp=""
		else
			pcv_strCSTemp=" AND pccats_RetailHide<>1 "
		end if									
		query="SELECT categories_products.idProduct "
		query=query+"FROM categories_products " 
		query=query+"INNER JOIN categories "
		query=query+"ON categories_products.idCategory = categories.idCategory "
		query=query+"WHERE categories_products.idProduct="& pidrelation &" AND iBTOhide=0 " & pcv_strCSTemp & " "
		query=query+"ORDER BY priority, categoryDesc ASC;"	
		set rsCheckCategory=server.CreateObject("ADODB.RecordSet")
		set rsCheckCategory=conntemp.execute(query)									
		If rsCheckCategory.eof Then
			pcv_intCategoryActive=2
		End If
		set rsCheckCategory=nothing
		
		'// CHECK FOR REQUIRED OPTIONS							
		pcv_intOptionsExist=pcf_CheckForReqOptions(pidrelation) '// check options function (1=YES, 2=NO)
		
		'// CHECK FOR REQUIRED INPUT FIELDS
		if pcv_intOptionsExist=2 then
			pcv_intOptionsExist=pcf_CheckForReqInputFields(pidrelation)
		end if
				
		if cs_pserviceSpec=true OR (pcv_intOptionsExist = 1) then
			If pcv_intOptionsExist = 1 Then
				cs_pOptCnt=cs_pOptCnt+1
			Else
				cs_pBTOCnt=cs_pBTOCnt+1
			End If
		End If
		
		' If item is either BTO or have options or is within Hidden Category,
		' do not require item (bundle) or 
		' do not require checkbox (accessory) 
		' as these will not be shown on page 
		if cint(cs_pOptCnt) + cint(cs_pBTOCnt) = 0 then 
			cs_pAddtoCart = 1
		end if		
		
		if ((cs_pAddtoCart=1 AND pcsType<>"Accessory") OR (pcsType="Accessory")) AND (pcv_intCategoryActive=1) then            		   
			if pcv_strIsRequiredAccessory <> 0 then
				pcf_CheckForReqAccessories=1
				Exit Function						
			end if
		end if
		
		'// Clear Variables
		cs_pBTOCnt=Cint(0)
		cs_pOptCnt=Cint(0)
		cs_pAddtoCart=Cint(0)
		cs_pserviceSpec=""
		pidrelation=Cint(0)
		pcsType=""
		
	loop

End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Required Accessories
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  if wholesale allowed, check if customer is also wholesale
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Function pcf_WholesaleCustomerAllowed
	if scorderlevel="1" AND session("customerType")=1 then
		pcf_WholesaleCustomerAllowed = true
	Else
		pcf_WholesaleCustomerAllowed = false
	End If
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  if wholesale allowed, check if customer is also wholesale
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Check if out of stock purchase allowed
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Function pcf_OutStockPurchaseAllow
	If (scOutofStockPurchase=-1 AND CLng(pStock)<1 AND pserviceSpec=0 AND pNoStock=0 AND pcv_intBackOrder=0) OR (pserviceSpec<>0 AND scOutofStockPurchase=-1 AND iBTOOutofstockpurchase=-1 AND CLng(pStock)<1 AND pNoStock=0 AND pcv_intBackOrder=0) Then
		pcf_OutStockPurchaseAllow = false
	Else
		pcf_OutStockPurchaseAllow = true
	End If
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Check if out of stock purchase allowed
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  QTY Minimums
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function pcf_CheckMinQty(ID)	
	'//Product ID
	pcv_str_CheckID=ID
	pcf_CheckMinQty=2
	query="SELECT products.pcprod_MinimumQty FROM products WHERE idProduct=" & pcv_str_CheckID & " AND configOnly=0 AND removed=0;" 
	set rsQtyMinimum=server.createobject("adodb.recordset")
	set rsQtyMinimum=conntemp.execute(query)	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsQtyMinimum=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if	
	If NOT rsQtyMinimum.eof Then
		pcv_lngMinimumQty=rsQtyMinimum("pcprod_MinimumQty")
		if isNull(pcv_lngMinimumQty) OR pcv_lngMinimumQty="" then
			pcv_lngMinimumQty="0"
		end if
		if pcv_lngMinimumQty <> 0 then
			pcf_CheckMinQty=1
		end if
	Else
		pcf_CheckMinQty=2
	End If		
	set rsQtyMinimum = nothing
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  QTY Minimums
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
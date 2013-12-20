<%ReDim PromoArr(100,5)
PromoIndex=0
PrdPromoArr=""
PrdPromoCount=0
CatPromoArr=""
CatPromoCount=0

'Include Option Prices to Total Product Price to calculate Product Promotion
PR_IncOptPrice=1 '//Defaut=1

PromoMsgStr=""
tmpShowList="***"

k=0
m=0

PromoIndex=0

pcTmpArr=Session("pcCartSession")
pcTmpIndex=Session("pcCartIndex")

call opendb()

pcv_HaveCatPromotions=0

query="SELECT pcCatPro_id,idcategory,pcCatPro_QtyTrigger,pcCatPro_DiscountType,pcCatPro_DiscountValue,pcCatPro_ApplyUnits,pcCatPro_PromoMsg,pcCatPro_ConfirmMsg,pcCatPro_SDesc FROM pcCatPromotions;"
set rsQ=connTemp.execute(query)
if not rsQ.eof then
	pcv_HaveCatPromotions=1
	CatPromoArr=rsQ.getRows()
	set rsQ=nothing
	CatPromoCount=ubound(CatPromoArr,2)
end if
set rsQ=nothing

if pcv_HaveCatPromotions=1 then
	For m=0 to CatPromoCount
		'call CalCatPromo(m)
	Next
end if

pcv_HavePrdPromotions=0

query="SELECT pcPrdPro_id,idproduct,pcPrdPro_QtyTrigger,pcPrdPro_DiscountType,pcPrdPro_DiscountValue,pcPrdPro_ApplyUnits,pcPrdPro_PromoMsg,pcPrdPro_ConfirmMsg,pcPrdPro_SDesc,pcPrdPro_IncExcCust,pcPrdPro_IncExcCPrice,pcPrdPro_RetailFlag,pcPrdPro_WholesaleFlag FROM pcPrdPromotions WHERE pcPrdPro_Inactive=0;"
set rsQ=connTemp.execute(query)
if not rsQ.eof then
	pcv_HavePrdPromotions=1
	PrdPromoArr=rsQ.getRows()
	set rsQ=nothing
	PrdPromoCount=ubound(PrdPromoArr,2)
end if
set rsQ=nothing

pcv_strUsedPrdPromo = ""
if pcv_HavePrdPromotions=1 then
	For k=1 to pcTmpIndex '// For each product in cart array
		For m=0 to PrdPromoCount '// For each promo available
			'// If Product ID = Promo Product ID
			if clng(CheckParentPrd(k))=clng(PrdPromoArr(1,m)) then
				call CalPrdPromo(k,m) '// Apply Promo
				pcv_strUsedPrdPromo = pcv_strUsedPrdPromo & CheckParentPrd(k) & "_" & m & "," '// Keep track for cleaning sessions
			end if
		Next
	Next
end if

Session("pcPromoSession")=PromoArr
Session("pcPromoIndex")=PromoIndex

Function CheckParentPrd(tmpIdx)
  Dim pcv_tmpPPrd
  if (pcTmpArr(tmpIdx,32)<>"") then
	  pcv_tmpPPrd=split(pcTmpArr(tmpIdx,32),"$$")
	  CheckParentPrd=pcv_tmpPPrd(ubound(pcv_tmpPPrd))
  else
	  CheckParentPrd=pcTmpArr(tmpIdx,0)
  end if
End Function

Function CheckSubCatList(tmpIdParent)
	Dim query,rsQ,tmpPcCatArr,tmpPcCatCount,j,tmpCatStr
	tmpCatStr=""
	query="SELECT idcategory FROM Categories WHERE idParentCategory=" & tmpIdParent & ";"
	set rsQ=connTemp.execute(query)
	if not rsQ.eof then
		tmpPcCatArr=rsQ.getRows()
		set rsQ=nothing
		tmpPcCatCount=ubound(tmpPcCatArr,2)
		For j=0 to tmpPcCatCount
			tmpCatStr=tmpCatStr & "," & tmpPcCatArr(0,j)
			tmpCatStr=tmpCatStr & CheckSubCatList(tmpPcCatArr(0,j))
		Next
	end if
	set rsQ=nothing
	CheckSubCatList=tmpCatStr
End Function

Function GetAvailableUnits(tmpC,tmpP)
	Dim pcv_intTrueCount,pcv_intBaseID,pcv_CheapestPrice, pcv_tmpCheapestPrice
	pcv_intTrueCount = 0
	pcv_CheapestPrice = cdbl(0)
	pcv_tmpCheapestPrice = cdbl(0)
	pcv_intBaseID = CheckParentPrd(tmpC)
	For UnitCount=1 to pcTmpIndex
		If cdbl(pcv_intBaseID) = cdbl(CheckParentPrd(UnitCount)) Then
			pcv_intTrueCount = cdbl(pcv_intTrueCount) + cdbl(pcTmpArr(UnitCount,2))
			'// Find Cheapest Unit
			if PR_IncOptPrice=1 then
				pcv_tmpCheapestPrice=(pcTmpArr(UnitCount,3)+pcTmpArr(UnitCount,5)+round((pcTmpArr(UnitCount,31)/pcTmpArr(UnitCount,2)),2))
			else
				pcv_tmpCheapestPrice=(pcTmpArr(UnitCount,3)+round((pcTmpArr(UnitCount,31)/pcTmpArr(UnitCount,2)),2))
			end if
			if (pcv_tmpCheapestPrice<pcv_CheapestPrice) OR (pcv_CheapestPrice=0) then
				pcv_CheapestPrice = pcv_tmpCheapestPrice
				Session("UAP_" & pcv_intBaseID & "_" & tmpP) = pcv_CheapestPrice
			end if
		End If
	Next
	GetAvailableUnits = pcv_intTrueCount
	'GetRealQty = clng(pcTmpArr(tmpP,2))
End Function

Sub CalPrdPromo(tmpC,tmpP)
	Dim query,rsQ,j
	Dim pcPrdTargetArr,PrdTargetCount,pcv_HavePrdTargets,tmpPrdTargetStr
	Dim pcCatTargetArr,CatTargetCount,pcv_HaveCatTargets,tmpCatTargetStr
	Dim tmpU,tmpTotal,tmpAll,tmpHas
	Dim tmpIDProduct
	
	tmpPrdTargetStr=""
	tmpCatTargetStr=""
	tmpIDCode=PrdPromoArr(0,tmpP) '// Promo ID
	tmpIDProduct=PrdPromoArr(1,tmpP)  '// Product ID
	tmpProductQty=GetAvailableUnits(tmpC,tmpP)  '// Qty in Cart that are eligible for Promo
	tmpQtyTrigger=clng(PrdPromoArr(2,tmpP)) '// Promo Trigger
	tmpDiscountType=PrdPromoArr(3,tmpP) '// Discount Type
	tmpDiscountValue=PrdPromoArr(4,tmpP) '// Discount Value
	tmpApplyUnits=PrdPromoArr(5,tmpP) '// Apply to "n" Units
	tmpConfirmMsg=PrdPromoArr(7,tmpP)
	tmpDescMsg=PrdPromoArr(8,tmpP)
	pcIncExcCust=PrdPromoArr(9,tmpP)
	pcIncExcCPrice=PrdPromoArr(10,tmpP)
	pcv_retail=PrdPromoArr(11,tmpP)
	pcv_wholeSale=PrdPromoArr(12,tmpP)
	
	pcv_Filters=0
	pcv_FResults=0
	
	'Filter by Customers
	pcv_CustFilter=0
	query="select IDCustomer from PcPPFCusts where pcPrdPro_id=" & tmpIDCode
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if	
	if not rs.eof then
		pcv_Filters=pcv_Filters+1
		pcv_CustFilter=1
	end if
	set rs=nothing	
	if pcv_CustFilter=1 then		
		query="select IDCustomer from PcPPFCusts where pcPrdPro_id=" & tmpIDCode & " and IDCustomer=" & session("IDCustomer")
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if	
		if not rs.eof then
			if (pcIncExcCust="0") then
				pcv_FResults=pcv_FResults+1
			end if
		else
			if (pcIncExcCust="1") then
				pcv_FResults=pcv_FResults+1
			end if
		end if
		set rs=nothing	
	end if
	'End of Filter by Customers
	
	
	'Filter by Customer Categories
	pcv_CustCatFilter=0	
	query="select idCustomerCategory from pcPPFCustPriceCats where pcPrdPro_id=" & tmpIDCode
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if	
	if not rs.eof then
		pcv_Filters=pcv_Filters+1
		pcv_CustCatFilter=1
	end if
	set rs=nothing	
	if pcv_CustCatFilter=1 then		
		query="select pcPPFCustPriceCats.idCustomerCategory from pcPPFCustPriceCats, Customers where pcPPFCustPriceCats.pcPrdPro_id=" & tmpIDCode & " and pcPPFCustPriceCats.idCustomerCategory = Customers.idCustomerCategory and Customers.idcustomer=" & session("IDCustomer")
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if	
		if not rs.eof then
			if (pcIncExcCPrice="0") then
				pcv_FResults=pcv_FResults+1
			end if
		else
			if (pcIncExcCPrice="1") then
				pcv_FResults=pcv_FResults+1
			end if
		end if
		set rs=nothing	
	end if
	'End of Filter by Customer Categories



	'Check to see if promotion is filtered by reatil or wholesale.
	if (pcv_retail ="0" and pcv_wholeSale ="1") or (pcv_retail ="1" and pcv_wholeSale ="0") Then
		pcv_Filters=pcv_Filters+1
		if pcv_wholeSale = "1" and session("customertype") = 1 then
			pcv_FResults=pcv_FResults+1	
		end if 
		if pcv_retail = "1" and session("customertype") <> 1 Then
			pcv_FResults=pcv_FResults+1
		end if    
	end if
	'End Check to see if promotion is filtered by reatil or wholesale.
	
	IF pcv_Filters=pcv_FResults THEN
	
		if Instr(tmpShowList,"***" & PrdPromoArr(1,tmpP) & "***")=0 then
			tmpShowList=tmpShowList & PrdPromoArr(1,tmpP) & "***"
			PromoMsgStr=PromoMsgStr & "<li>" & PrdPromoArr(6,tmpP) & "</li>"
		end if
		
		'Have Trigger
		if (tmpProductQty-tmpQtyTrigger>=0) AND (Session("UA_" & CheckParentPrd(tmpC) & "_" & tmpP) <> "DONE") then
	
			'Get Target Products
			query="SELECT idproduct FROM pcPPFProducts WHERE pcPrdPro_ID=" & tmpIDCode & ";"
			set rsQ=connTemp.execute(query)
			pcv_HavePrdTargets=0
			if not rsQ.eof then
				pcPrdTargetArr=rsQ.getRows()
				set rsQ=nothing
				PrdTargetCount=ubound(pcPrdTargetArr,2)
				For j=0 to PrdTargetCount
					tmpPrdTargetStr=tmpPrdTargetStr & "**" & pcPrdTargetArr(0,j)
				Next
				tmpPrdTargetStr=tmpPrdTargetStr & "**"
				pcv_HavePrdTargets=1
			end if
			set rsQ=nothing
			
			'Get Target Categories
			query="SELECT idcategory,pcPPFCats_IncSubCats FROM pcPPFCategories WHERE pcPrdPro_ID=" & tmpIDCode & ";"
			set rsQ=connTemp.execute(query)
			pcv_HaveCatTargets=0
			if not rsQ.eof then
				pcCatTargetArr=rsQ.getRows()
				set rsQ=nothing
				CatTargetCount=ubound(pcCatTargetArr,2)
				For j=0 to CatTargetCount
					if tmpCatTargetStr<>"" then
						tmpCatTargetStr=tmpCatTargetStr & ","
					end if
					tmpCatTargetStr=tmpCatTargetStr & pcCatTargetArr(0,j)
					if pcCatTargetArr(1,j)="1" then
						tmpCatTargetStr=tmpCatTargetStr & CheckSubCatList(pcCatTargetArr(0,j))
					end if
				Next
				pcv_HaveCatTargets=1
			end if
			set rsQ=nothing

			'Filter by Target Products
			IF pcv_HavePrdTargets=1 OR (pcv_HavePrdTargets=0 AND pcv_HaveCatTargets=0) THEN

				'// Apply the Promo to all available Qty until you run out
				Do while (tmpProductQty>0) AND (clng(tmpProductQty)-clng(tmpQtyTrigger)>=0)					
					tmpProductQty=tmpProductQty-clng(tmpQtyTrigger)
					
					tmpU=tmpApplyUnits '// # of Units
					if tmpApplyUnits=0 then
						tmpAll=1 '// Apply to all additional
					else
						tmpAll=0 '// Appy to "n" units
					end if		
							
					tmpTotal=0
					tmpQtyTotal=0
					
					For j=1 to pcTmpIndex '// For each item in cart array
	
						'// If (Not Deleted) AND (Qty more than Zero) AND ((No Target Product) OR (Matching Target Product))		
						if pcTmpArr(j,10)="0" AND tmpProductQty>"0" AND ((tmpPrdTargetStr="") OR (InStr(tmpPrdTargetStr,"**" & CheckParentPrd(j) & "**")>0)) then	
							if tmpAll=1 then
								
								'// APPLY TO ALL
								tmpTotal=tmpTotal+(tmpProductQty*Session("UAP_" & CheckParentPrd(tmpC) & "_" & tmpP))
								tmpQtyTotal=tmpQtyTotal+tmpProductQty
								tmpProductQty=0
								
							else
	
								'// APPLY TO 'N' UNITS														
								if tmpU-clng(tmpProductQty)>=0 then '// If all "n" units are available
									tmpTotal=tmpTotal+(tmpProductQty*Session("UAP_" & CheckParentPrd(tmpC) & "_" & tmpP))
									tmpU=tmpU-clng(tmpProductQty)
									tmpProductQty=0								
								else '// Not all "n" units are left
									tmpTotal=tmpTotal+(tmpU*Session("UAP_" & CheckParentPrd(tmpC) & "_" & tmpP))
									tmpProductQty=tmpProductQty-tmpU
									tmpU=0
								end if
	
							end if
							
						end if
						
						'// Check if we are done
						if tmpU=0 AND tmpAll=0 then
							exit for
						end if
					Next
					
					'// Prior to exit set a session flag so this promo cannot be applied to the same ID again
					Session("UA_" & CheckParentPrd(tmpC) & "_" & tmpP) = "DONE" '// UA_ProductID_PromoID
	
					'// Apply the Total Promo
					if tmpTotal>0 then
						if tmpDiscountType="1" then
							if tmpQtyTotal>0 then
								tmpTotal=tmpDiscountValue*tmpQtyTotal
							else
								tmpTotal=tmpDiscountValue*(tmpApplyUnits-tmpU)
							end if
						else
							tmpTotal=Round((tmpTotal*tmpDiscountValue)/100,2)
						end if
						tmpHas=0
						For j=1 to PromoIndex
							if PromoArr(j,0)="P" & tmpIDCode then
								PromoArr(j,2)=PromoArr(j,2)+tmpTotal
								tmpHas=1
								exit for
							end if
						Next
						if tmpHas=0 then
							PromoIndex=PromoIndex+1
							j=PromoIndex
							PromoArr(j,0)="P" & tmpIDCode
							PromoArr(j,1)=tmpConfirmMsg
							PromoArr(j,2)=tmpTotal
							PromoArr(j,3)=tmpDescMsg
							PromoArr(j,4)=tmpIDProduct
						end if
					end if
				Loop
			END IF 'Filter by Target Products
			
			'Filter by Target Categories
			IF pcv_HaveCatTargets=1 THEN
				Do while (pcTmpArr(tmpC,2)>0) AND (clng(pcTmpArr(tmpC,2))-clng(tmpQtyTrigger)>=0)
					pcTmpArr(tmpC,2)=pcTmpArr(tmpC,2)-clng(tmpQtyTrigger)
					tmpU=tmpApplyUnits
					if tmpApplyUnits=0 then
						tmpAll=1
					else
						tmpAll=0
					end if
					tmpTotal=0
					tmpQtyTotal=0
					For j=1 to pcTmpIndex
						if pcTmpArr(j,10)="0" AND pcTmpArr(j,2)>"0" then
							query="SELECT idcategory FROM categories_products WHERE idproduct=" & pcTmpArr(j,0) & " AND idcategory IN (" & tmpCatTargetStr & ");"
							set rsQ=connTemp.execute(query)
							if not rsQ.eof then
								set rsQ=nothing
								if tmpAll=1 then
									if PR_IncOptPrice=1 then
										tmpTotal=tmpTotal+(pcTmpArr(j,2)*(pcTmpArr(j,3)+pcTmpArr(j,5)))
									else
										tmpTotal=tmpTotal+(pcTmpArr(j,2)*pcTmpArr(j,3))
									end if
									tmpQtyTotal=tmpQtyTotal+pcTmpArr(j,2)
									pcTmpArr(j,2)=0
								else
									if tmpU-clng(pcTmpArr(j,2))>=0 then
										if PR_IncOptPrice=1 then
											tmpTotal=tmpTotal+(pcTmpArr(j,2)*(pcTmpArr(j,3)+pcTmpArr(j,5)))
										else
											tmpTotal=tmpTotal+(pcTmpArr(j,2)*pcTmpArr(j,3))
										end if
										tmpU=tmpU-clng(pcTmpArr(j,2))
										pcTmpArr(j,2)=0
									else
										if PR_IncOptPrice=1 then
											tmpTotal=tmpTotal+(tmpU*(pcTmpArr(j,3)+pcTmpArr(j,5)))
										else
											tmpTotal=tmpTotal+(tmpU*pcTmpArr(j,3))
										end if
										pcTmpArr(j,2)=pcTmpArr(j,2)-tmpU
										tmpU=0
									end if
								end if
							end if
							set rsQ=nothing
						end if
						if tmpU=0 AND tmpAll=0 then
							exit for
						end if
					Next
					if tmpTotal>0 then
						if tmpDiscountType="1" then
							if tmpQtyTotal>0 then
								tmpTotal=tmpDiscountValue*tmpQtyTotal
							else
								tmpTotal=tmpDiscountValue*(tmpApplyUnits-tmpU)
							end if
						else
							tmpTotal=Round((tmpTotal*tmpDiscountValue)/100,2)
						end if
						tmpHas=0
						For j=1 to PromoIndex
							if PromoArr(j,0)="P" & tmpIDCode then
								PromoArr(j,2)=PromoArr(j,2)+tmpTotal
								tmpHas=1
								exit for
							end if
						Next
						if tmpHas=0 then
							PromoIndex=PromoIndex+1
							j=PromoIndex
							PromoArr(j,0)="P" & tmpIDCode
							PromoArr(j,1)=tmpConfirmMsg
							PromoArr(j,2)=tmpTotal
							PromoArr(j,3)=tmpDescMsg
							PromoArr(j,4)=tmpIDProduct
						end if
					end if
				Loop
			END IF 'Filter by Target Categories
			
		end if 'Have Trigger

	END IF '// IF pcv_Filters=pcv_FResults THEN
End Sub

Function CheckCatTrigger(tmpIDCat,QtyTrigger,istyle,cindex)
	Dim query,rsQ,j,tmpQty

	tmpQty=QtyTrigger
	For j=1 to pcTmpIndex
		if pcTmpArr(j,10)="0" AND pcTmpArr(j,2)>"0" then
			query="SELECT idcategory FROM categories_products WHERE idproduct=" & pcTmpArr(j,0) & " AND idcategory=" & tmpIDCat & ";"
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				set rsQ=nothing
				if istyle="0" then
					PromoMsgStr=PromoMsgStr & "<li>" & CatPromoArr(6,cindex) & "</li>"
				end if
				if clng(tmpQty)-clng(pcTmpArr(j,2))>=0 then
					tmpQty=clng(tmpQty)-clng(pcTmpArr(j,2))
					if istyle="1" then
						pcTmpArr(j,2)=0
					end if
				else
					if istyle="1" then
						pcTmpArr(j,2)=pcTmpArr(j,2)-tmpQty
					end if
					tmpQty=0
				end if
			end if
			set rsQ=nothing
		end if
		if tmpQty=0 then
			exit for
		end if
	Next
	if tmpQty=0 then
		CheckCatTrigger=true
	else
		CheckCatTrigger=false
	end if
End Function

Sub CalCatPromo(tmpP)
Dim query,rsQ,j
Dim pcPrdTargetArr,PrdTargetCount,pcv_HavePrdTargets,tmpPrdTargetStr
Dim pcCatTargetArr,CatTargetCount,pcv_HaveCatTargets,tmpCatTargetStr
Dim tmpU,tmpTotal,tmpAll,tmpHas
	tmpPrdTargetStr=""
	tmpCatTargetStr=""
	tmpIDCode=CatPromoArr(0,tmpP)
	tmpIDCat=CatPromoArr(1,tmpP)
	tmpQtyTrigger=clng(CatPromoArr(2,tmpP))
	tmpDiscountType=CatPromoArr(3,tmpP)
	tmpDiscountValue=CatPromoArr(4,tmpP)
	tmpApplyUnits=CatPromoArr(5,tmpP)
	tmpConfirmMsg=CatPromoArr(7,tmpP)
	tmpDescMsg=CatPromoArr(8,tmpP)
	
	'Have Trigger
	if CheckCatTrigger(tmpIDCat,tmpQtyTrigger,0,tmpP) then
		'Get Target Products
		query="SELECT idproduct FROM pcCPFProducts WHERE pcCatPro_ID=" & tmpIDCode & ";"
		set rsQ=connTemp.execute(query)
		pcv_HavePrdTargets=0
		if not rsQ.eof then
			pcPrdTargetArr=rsQ.getRows()
			set rsQ=nothing
			PrdTargetCount=ubound(pcPrdTargetArr,2)
			For j=0 to PrdTargetCount
				tmpPrdTargetStr=tmpPrdTargetStr & "**" & pcPrdTargetArr(0,j)
			Next
			tmpPrdTargetStr=tmpPrdTargetStr & "**"
			pcv_HavePrdTargets=1
		end if
		set rsQ=nothing
		
		'Get Target Categories
		query="SELECT idcategory,pcCPFCats_IncSubCats FROM pcCPFCategories WHERE pcCatPro_ID=" & tmpIDCode & ";"
		set rsQ=connTemp.execute(query)
		pcv_HaveCatTargets=0
		if not rsQ.eof then
			pcCatTargetArr=rsQ.getRows()
			set rsQ=nothing
			CatTargetCount=ubound(pcCatTargetArr,2)
			For j=0 to CatTargetCount
				if tmpCatTargetStr<>"" then
					tmpCatTargetStr=tmpCatTargetStr & ","
				end if
				tmpCatTargetStr=tmpCatTargetStr & pcCatTargetArr(0,j)
				if pcCatTargetArr(1,j)="1" then
					tmpCatTargetStr=tmpCatTargetStr & CheckSubCatList(pcCatTargetArr(0,j))
				end if
			Next
			pcv_HaveCatTargets=1
		end if
		set rsQ=nothing
		
		
		'Filter by Target Products
		IF pcv_HavePrdTargets=1 OR (pcv_HavePrdTargets=0 AND pcv_HaveCatTargets=0) THEN
			Do while CheckCatTrigger(tmpIDCat,tmpQtyTrigger,1,tmpP)
				tmpU=tmpApplyUnits
				if tmpApplyUnits=0 then
					tmpAll=1
				else
					tmpAll=0
				end if
				tmpTotal=0
				tmpQtyTotal=0
				For j=1 to pcTmpIndex
					if pcTmpArr(j,10)="0" AND pcTmpArr(j,2)>"0" AND ((tmpPrdTargetStr="") OR (InStr(tmpPrdTargetStr,"**" & pcTmpArr(j,0) & "**")>0)) then
						if tmpAll=1 then
							if PR_IncOptPrice=1 then
								tmpTotal=tmpTotal+(pcTmpArr(j,2)*(pcTmpArr(j,3)+pcTmpArr(j,5)))
							else
								tmpTotal=tmpTotal+(pcTmpArr(j,2)*pcTmpArr(j,3))
							end if
							tmpQtyTotal=tmpQtyTotal+pcTmpArr(j,2)
							pcTmpArr(j,2)=0
						else
							if tmpU-clng(pcTmpArr(j,2))>=0 then
								if PR_IncOptPrice=1 then
									tmpTotal=tmpTotal+(pcTmpArr(j,2)*(pcTmpArr(j,3)+pcTmpArr(j,5)))
								else
									tmpTotal=tmpTotal+(pcTmpArr(j,2)*pcTmpArr(j,3))
								end if
								tmpU=tmpU-clng(pcTmpArr(j,2))
								pcTmpArr(j,2)=0
							else
								if PR_IncOptPrice=1 then
									tmpTotal=tmpTotal+(tmpU*(pcTmpArr(j,3)+pcTmpArr(j,5)))
								else
									tmpTotal=tmpTotal+(tmpU*pcTmpArr(j,3))
								end if
								pcTmpArr(j,2)=pcTmpArr(j,2)-tmpU
								tmpU=0
							end if
						end if
					end if
					if tmpU=0 AND tmpAll=0 then
						exit for
					end if
				Next
				if tmpTotal>0 then
					if tmpDiscountType="1" then
						if tmpQtyTotal>0 then
							tmpTotal=tmpDiscountValue*tmpQtyTotal
						else
							tmpTotal=tmpDiscountValue*(tmpApplyUnits-tmpU)
						end if
					else
						tmpTotal=Round((tmpTotal*tmpDiscountValue)/100,2)
					end if
					tmpHas=0
					For j=1 to PromoIndex
						if PromoArr(j,0)="C" & tmpIDCode then
							PromoArr(j,2)=PromoArr(j,2)+tmpTotal
							tmpHas=1
							exit for
						end if
					Next
					if tmpHas=0 then
						PromoIndex=PromoIndex+1
						j=PromoIndex
						PromoArr(j,0)="C" & tmpIDCode
						PromoArr(j,1)=tmpConfirmMsg
						PromoArr(j,2)=tmpTotal
						PromoArr(j,3)=tmpDescMsg
					end if
				end if
			Loop
		END IF 'Filter by Target Products
		
		'Filter by Target Categories
		IF pcv_HaveCatTargets=1 THEN
			Do while CheckCatTrigger(tmpIDCat,tmpQtyTrigger,1,tmpP)
				tmpU=tmpApplyUnits
				if tmpApplyUnits=0 then
					tmpAll=1
				else
					tmpAll=0
				end if
				tmpTotal=0
				tmpQtyTotal=0
				For j=1 to pcTmpIndex
					if pcTmpArr(j,10)="0" AND pcTmpArr(j,2)>"0" then
						query="SELECT idcategory FROM categories_products WHERE idproduct=" & pcTmpArr(j,0) & " AND idcategory IN (" & tmpCatTargetStr & ");"
						set rsQ=connTemp.execute(query)
						if not rsQ.eof then
							set rsQ=nothing
							if tmpAll=1 then
								if PR_IncOptPrice=1 then
									tmpTotal=tmpTotal+(pcTmpArr(j,2)*(pcTmpArr(j,3)+pcTmpArr(j,5)))
								else
									tmpTotal=tmpTotal+(pcTmpArr(j,2)*pcTmpArr(j,3))
								end if
								tmpQtyTotal=tmpQtyTotal+pcTmpArr(j,2)
								pcTmpArr(j,2)=0
							else
								if tmpU-clng(pcTmpArr(j,2))>=0 then
									if PR_IncOptPrice=1 then
										tmpTotal=tmpTotal+(pcTmpArr(j,2)*(pcTmpArr(j,3)+pcTmpArr(j,5)))
									else
										tmpTotal=tmpTotal+(pcTmpArr(j,2)*pcTmpArr(j,3))
									end if
									tmpU=tmpU-clng(pcTmpArr(j,2))
									pcTmpArr(j,2)=0
								else
									if PR_IncOptPrice=1 then
										tmpTotal=tmpTotal+(tmpU*(pcTmpArr(j,3)+pcTmpArr(j,5)))
									else
										tmpTotal=tmpTotal+(tmpU*pcTmpArr(j,3))
									end if
									pcTmpArr(j,2)=pcTmpArr(j,2)-tmpU
									tmpU=0
								end if
							end if
						end if
						set rsQ=nothing
					end if
					if tmpU=0 AND tmpAll=0 then
						exit for
					end if
				Next
				if tmpTotal>0 then
					if tmpDiscountType="1" then
						if tmpQtyTotal>0 then
							tmpTotal=tmpDiscountValue*tmpQtyTotal
						else
							tmpTotal=tmpDiscountValue*(tmpApplyUnits-tmpU)
						end if
					else
						tmpTotal=Round((tmpTotal*tmpDiscountValue)/100,2)
					end if
					tmpHas=0
					For j=1 to PromoIndex
						if PromoArr(j,0)="C" & tmpIDCode then
							PromoArr(j,2)=PromoArr(j,2)+tmpTotal
							tmpHas=1
							exit for
						end if
					Next
					if tmpHas=0 then
						PromoIndex=PromoIndex+1
						j=PromoIndex
						PromoArr(j,0)="C" & tmpIDCode
						PromoArr(j,1)=tmpConfirmMsg
						PromoArr(j,2)=tmpTotal
						PromoArr(j,3)=tmpDescMsg
					end if
				end if
			Loop
		END IF 'Filter by Target Categories
		
	end if 'Have Trigger
End Sub

'// Clean Up Sessions
pcArray_UsedPrdPromo = split(pcv_strUsedPrdPromo,",")
For PrdPromoCount=0 to ubound(pcArray_UsedPrdPromo)
	Session("UA_" & pcArray_UsedPrdPromo(PrdPromoCount)) = ""
	Session("UAP_" & pcArray_UsedPrdPromo(PrdPromoCount)) = ""
Next
%>



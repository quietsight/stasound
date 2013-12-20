<%	
	
Sub RelCalCS(cartIdx,IDParent,IDBundle,ctype)
Dim query,rs,y,query1
call opendb()
if ctype=0 then 'Parent
	query1="cs_relationships.idproduct=products.idProduct"
else 'Bundle
	query1="cs_relationships.idrelation=products.idProduct"
end if
query="SELECT cs_relationships.idproduct, cs_relationships.idrelation, cs_relationships.cs_type, cs_relationships.discount, cs_relationships.ispercent,cs_relationships.isRequired, products.servicespec, products.price, products.description, products.bToBprice, products.serviceSpec, products.noprices FROM cs_relationships INNER JOIN products ON " & query1 & " WHERE cs_relationships.idrelation=" & IDBundle & " AND cs_relationships.idproduct="&IDParent&" AND products.active=-1 AND products.removed=0 ORDER BY cs_relationships.num,cs_relationships.idrelation;"
set rs=conntemp.execute(query)
if not rs.eof then
	pcArray_CSRelations = rs.getRows()
	set rs=nothing
	pcv_intProductCount = UBound(pcArray_CSRelations,2)
	
	For y=0 to pcv_intProductCount
					if ctype=0 then 'Parent
						pidrelation=pcArray_CSRelations(0,y) '// rs("idproduct")
					else 'Bundle
						pidrelation=pcArray_CSRelations(1,y) '// rs("idrelation")
					end if
					pcsType=pcArray_CSRelations(2,y) '// rs("cs_type")			
					pDiscount=pcArray_CSRelations(3,y) '// rs("discount")
					pIsPercent=pcArray_CSRelations(4,y) '// rs("isPercent")
					pcv_strIsRequired=pcArray_CSRelations(5,y) '// rs("isRequired")
					cs_pserviceSpec=pcArray_CSRelations(6,y) '// rs("servicespec")
					
					ppPrice=pcArray_CSRelations(7,y) '// rs("price")
					
					if pcArray_CSRelations(9,y)>"0" then
						ppBPrice=pcArray_CSRelations(9,y)
					else
						ppBPrice=ppPrice
					end if
					
					cs_pserviceSpec=pcArray_CSRelations(10,y)
					if cs_pserviceSpec="" OR IsNull(cs_pserviceSpec) then
						cs_pserviceSpec=0
					end if
					cs_pnoprices=pcArray_CSRelations(11,y)
					if cs_pnoprices="" OR IsNull(cs_pnoprices) then
						cs_pnoprices=0
					end if
					
					tmp_pidProduct=pidProduct
					tmp_pPrice=pPrice
					tmp_pPrice1=pPrice1
					tmp_pBtoBPrice=pBtoBPrice
					tmp_pBtoBPrice1=pBtoBPrice1
					tmp_pnoprices=pnoprices
					tmp_pserviceSpec=pserviceSpec
					
					pidProduct=pidrelation
					pPrice=ppPrice
					pBtoBPrice=ppBPrice
					pnoprices=cs_pnoprices
					pserviceSpec=cs_pserviceSpec
					
					ppPrice1=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,0)
					ppBtoBPrice1=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,1)
					if session("customertype")=1 and ppBtoBPrice1>0 then
						ppPrice1=ppBtoBPrice1
					end if
					
					pidProduct=tmp_pidProduct
					pPrice=tmp_pPrice
					pPrice1=tmp_pPrice1
					pBtoBPrice=tmp_pBtoBPrice
					pBtoBPrice1=tmp_pBtoBPrice1
					pnoprices=tmp_pnoprices
					pserviceSpec=tmp_pserviceSpec
					
										
					if pIsPercent<>0 then
						ppPrice1 = CDbl(ppPrice1*(pDiscount/100))
					else
						if ctype=0 then
							ppPrice1=0
						else
							ppPrice1 = CDbl(pDiscount)
						end if
					end if
	
	Next
	pcCartArray(cartIdx,28)=ppPrice1
end if
set rs=nothing
call closedb()
End Sub

	for f=1 to ppcCartIndex
		if pcCartArray(f,10)=0 AND pcCartArray(f,27)<>"" AND pcCartArray(f,27)<>"0" then
				if pcCartArray(f,27)="-1" AND pcCartArray(f,8)>"0" then
					call RelCalCS(f,pcCartArray(f,0),pcCartArray(f,8),0)
				else
					if pcCartArray(f,27)>"0" then
						call RelCalCS(f,pcCartArray(pcCartArray(f,27),0),pcCartArray(f,0),1)
					end if
				end if
		end if
	next

call closedb()
%>
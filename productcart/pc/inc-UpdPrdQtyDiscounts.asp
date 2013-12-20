<%	
	Function CheckSamePrd(m,n)
	Dim tmpRe,y
		tmpRe=true
		For y=0 to 32
			if (y<>0) AND (y<>2) AND (y<>3) AND (y<>15) AND (y<>17) AND (y<>21) AND (pcCartArray(m,y) & ""<>pcCartArray(n,y) & "") then
				tmpRe=false
				exit for
			end if
		Next
		CheckSamePrd=tmpRe
	End Function
	
	pcv_SourcePrd=0
	'tmpSaveSource="||"
	for f=1 to ppcCartIndex
		if pcCartArray(f,10)=0 then
				pcv_SourcePrd=pcCartArray(f,0)
				'if instr(tmpSaveSource,"||" & pcv_SourcePrd & "||")=0 then
					'tmpSaveSource=tmpSaveSource & pcv_SourcePrd & "||"
					tmpQuantity=0
					For k=1 to ppcCartIndex
						if pcCartArray(k,10)=0 then
							if (pcCartArray(k,0)=pcv_SourcePrd) AND CheckSamePrd(f,k) then
								tmpQuantity=tmpQuantity+cLng(pcCartArray(k,2))
							end if
						end if
					Next
	
					disTotalQuantity=tmpQuantity
					
					query="SELECT discountPerUnit,discountPerWUnit,percentage,baseproductonly FROM discountsPerQuantity WHERE idProduct=" & pcv_SourcePrd & ";"
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=conntemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rstemp=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					if not rstemp.eof then
						For k=1 to ppcCartIndex
							if pcCartArray(k,10)=0 then
								if (pcCartArray(k,0)=pcv_SourcePrd) AND CheckSamePrd(f,k) then
									pcCartArray(k,15)=0
									pcCartArray(k,3)=pcCartArray(k,17)
								end if
							end if
						Next
					end if
					
					query="SELECT discountPerUnit,discountPerWUnit,percentage,baseproductonly FROM discountsPerQuantity WHERE idProduct=" & pcv_SourcePrd & " AND quantityFrom<=" &disTotalQuantity& " AND quantityUntil>=" &disTotalQuantity
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=conntemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rstemp=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					
					if not rstemp.eof and err.number<>9 then
						' there are quantity discounts defined for that quantity 
						pDiscountPerUnit=rstemp("discountPerUnit")
						pDiscountPerWUnit=rstemp("discountPerWUnit")
						pPercentage=rstemp("percentage")
						pbaseproductonly=rstemp("baseproductonly")
		
						set rstemp=nothing
						
						IF pcQDiscountType="1" THEN
							'====================
							' get original price
							'====================
							query="SELECT price, bToBPrice FROM products WHERE idProduct=" &pcv_SourcePrd
							set rstemp=conntemp.execute(query)
							
							if not rstemp.eof then
								tempprice=rstemp("price")
								tempbToBPrice=rstemp("bToBPrice")
							end if
				
							set rstemp = nothing
				
							'Check if this customer is logged in with a customer category
							if session("customerCategory")<>0 then
								query="SELECT pcCC_Price FROM pcCC_Pricing WHERE idcustomerCategory="&session("customerCategory")&" AND idProduct="&pcv_SourcePrd&";"
								set rs=server.CreateObject("ADODB.RecordSet")
								set rs=conntemp.execute(query)
					
								if NOT rs.eof then
									strcustomerCategory="YES"
									dblpcCC_Price=rs("pcCC_Price")
									dblpcCC_Price=pcf_Round(dblpcCC_Price, 2)
								else
									strcustomerCategory="NO"
								end if
					
								set rs=nothing
							end if
					
							if (tempbToBPrice<>0) then
								tempBPrice=tempbToBPrice
							else
								tempBPrice=tempprice
							end if

							if session("customerType")=1 then
								pPrice=tempBPrice
							else
								pPrice=tempprice
							end if 

							if session("customerCategoryType")="ATB" then
								if session("ATBCustomer")=1 AND session("ATBPercentOff")=1 then
									tempBPrice=tempBPrice-(pcf_Round(tempBPrice*(cdbl(session("ATBPercentage"))/100),2))
									pPrice=tempBPrice
								end if
								if session("ATBCustomer")=1 AND session("ATBPercentOff")=0 then
									tempprice=tempprice-(pcf_Round(tempprice*(cdbl(session("ATBPercentage"))/100),2))
									pPrice=tempprice
								end if					
							end if
				
							'if strcustomerCategory="YES" AND dblpcCC_Price>0 then
							if strcustomerCategory="YES" then
								pcCartArray(ppcCartIndex,3)=dblpcCC_Price
								pPrice=dblpcCC_Price
							end if
				
						END IF 'pcQDiscountType="1"
						
						For k=1 to ppcCartIndex
						if pcCartArray(k,10)=0 then
							if (pcCartArray(k,0)=pcv_SourcePrd) AND CheckSamePrd(f,k) then
								if session("customerType")=1 then
									pcCartArray(k,18)=1
								else
									pcCartArray(k,18)=0
								end if
		
								pOrigPrice=pcCartArray(k,17)
								pcCartArray(k,15)=0
								pTotalQuantity=pcCartArray(k,2)
								'reset price for source product
								pcCartArray(k,3)=pOrigPrice
								if pcQDiscountType<>"1" then
									pOrigPrice=pOrigPrice-(pcCartArray(k,30)/pTotalQuantity)
								else
									pOrigPrice=pPrice
								end if
								
								if session("customerType")<>1 then  'customer is a normal user
									if pPercentage="0" then 
										pcCartArray(k,3)=pcCartArray(k,3) - pDiscountPerUnit  'Price - discount per unit
										pcCartArray(k,15)=pcCartArray(k,15) + (pDiscountPerUnit * pTotalQuantity)  'running total of discounts
									else
										if pbaseproductonly="-1" then
											pcCartArray(k,3)=pcCartArray(k,3) - ((pDiscountPerUnit/100) * pOrigPrice)
										else
											pcCartArray(k,3)=pcCartArray(k,3) - ((pDiscountPerUnit/100) * (pOrigPrice+pcCartArray(k,5)))
										end if
										if pbaseproductonly="-1" then
											pcCartArray(k,15)=pcCartArray(k,15) + (((pDiscountPerUnit/100) * pOrigPrice) * pTotalQuantity)
										else
											pcCartArray(k,15)=pcCartArray(k,15) + (((pDiscountPerUnit/100) * (pOrigPrice+pcCartArray(k,5))) * pTotalQuantity)
										end if
									end if
								else  'customer is a wholesale customer
									if pPercentage="0" then 
										pcCartArray(k,3)=pcCartArray(k,3) - pDiscountPerWUnit
										pcCartArray(k,15)=pcCartArray(k,15) + (pDiscountPerWUnit * pTotalQuantity)
									else
										if pbaseproductonly="-1" then
											pcCartArray(k,3)=pcCartArray(k,3) - ((pDiscountPerWUnit/100) * pOrigPrice)
										else
											pcCartArray(k,3)=pcCartArray(k,3) - ((pDiscountPerWUnit/100) * (pOrigPrice+pcCartArray(k,5)))
										end if
										if pbaseproductonly="-1" then
											pcCartArray(k,15)=pcCartArray(k,15) + (((pDiscountPerWUnit/100) * pOrigPrice)* pTotalQuantity)
										else
											pcCartArray(k,15)=pcCartArray(k,15) + (((pDiscountPerWUnit/100) * (pOrigPrice+pcCartArray(k,5))) * pTotalQuantity)
										end if
									end if
								end if
							end if
						end if
						Next
					end if
					set rstemp=nothing
				'end if
		end if
	next
%>
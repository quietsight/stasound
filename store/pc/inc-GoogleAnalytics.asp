<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'// GOOGLE ANALYTICS
'// E-commerce transaction tracking
'// http://www.google.com/support/analytics/bin/answer.py?answer=27203&topic=7282

'// Check whether the code should be run
'// This variable is set on orderComplete.asp
'// The code is ONLY run when an order has been completed
Function pcf_GoogleEscape(addItem)
	Dim pcv_GoogleEscape
	pcv_GoogleEscape = addItem
	if len(addItem)>0 then
		pcv_GoogleEscape=replace(pcv_GoogleEscape,"'","\'")
		pcv_GoogleEscape=replace(pcv_GoogleEscape,"""","\""")
	end if
	pcf_GoogleEscape=pcv_GoogleEscape
End Function

IF pcGAorderComplete=1 AND (trim(scGoogleAnalytics)<>"" AND NOT IsNull(scGoogleAnalytics)) THEN ' 1

		'on error resume next
		
		dim pGetItems, pcGAtransaction, pcGAtransactionItems
		
		pGetItems = 1
		
		call openDb()
		
		'// Google Analytics
		'// STEP 1: GENERATE ORDER INFO LINE
		'// STEP 2: GENERATE ITEM INFO LINES
		
		'// STEP 1 - START
		
			'// Get order info from db
		
			query="SELECT city, state, stateCode, CountryCode, shipmentDetails, idAffiliate, taxAmount, total, taxDetails, ord_VAT FROM orders WHERE idOrder=" & pOID
			set rs=Server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)
			
			if err.number <> 0 then
				set rs=nothing
				pGetItems = 0
				call closeDb()
			end If
			
			if rs.eof then
				set rs=nothing
				pGetItems = 0
				call closeDb()
			end if
						
			IF pGetItems<>0 THEN ' 2 - If database errors, skip everything
		
				'// Transaction line variables per Google Analytics documentation
				'// [order-id]  Your internal unique order id number  
				'// [affiliation]  Optional partner or store affilation  
				'// [total]  Total dollar amount of the transaction  
				'// [tax]  Tax amount of the transaction  
				'// [shipping]  The shipping amount of the transaction  
				'// [city]  City to correlate the transaction with  
				'// [state/region]  State or province  
				'// [country]  Country
				
				pIdOrder = pOID
				
				'// Gather affiliate information
				pidAffiliate=rs("idaffiliate")
					If pidaffiliate>"1" then
						query="SELECT affiliateName, affiliateCompany FROM affiliates WHERE idAffiliate =" & pidAffiliate
						Set rsTemp=Server.CreateObject("ADODB.Recordset")
						Set rsTemp=connTemp.execute(query)
						paffiliateName = rsTemp("affiliateName")
						paffiliateCompany = rsTemp("affiliateCompany")
							if trim(paffiliateCompany)<>"" then
								paffiliateName = paffiliateName & "(" & paffiliateCompany & ")"
							end if
						paffiliateName = replace(paffiliateName,"|","-")
						Set rsTemp = nothing
					else
						paffiliateName = "N/A"
					end if
					
				'// Order Total
				ptotal=rs("total")
					
				'// Gather tax information to determine total tax amount
				ptaxAmount=rs("taxAmount")
				ptaxDetails=rs("taxDetails")
				pord_VAT=rs("ord_VAT")
				if pord_VAT>0 then
					ptaxAmount=pord_VAT
				else
					if isNull(ptaxDetails) OR trim(ptaxDetails)="" then
						ptaxAmount=ptaxAmount
					else
						taxArray=split(ptaxDetails,",")
						ptaxTotal=0
						for i=0 to (ubound(taxArray)-1)
							taxDesc=split(taxArray(i),"|")
							ptaxTotal=ptaxTotal+taxDesc(1)
						next 
						ptaxAmount=ptaxTotal
					end if
				end if
				
				'// Gather shipping information to determine total shipping amount
				pshipmentDetails=rs("shipmentDetails")
					pTotalShipping=0
					shipping=split(pshipmentDetails,",")
					if ubound(shipping)>1 then
						if NOT isNumeric(trim(shipping(2))) then
							pTotalShipping=0
						else
							Postage=trim(shipping(2))
							if ubound(shipping)=>3 then
								serviceHandlingFee=trim(shipping(3))
								if NOT isNumeric(serviceHandlingFee) then
									serviceHandlingFee=0
								end if
							else
								serviceHandlingFee=0
							end if
							if serviceHandlingFee<>0 then
								pTotalShipping=CDbl(Postage)+CDbl(serviceHandlingFee)
							else
								pTotalShipping=CDbl(Postage)
							end if
						end if
					end if
						
				'// Gather order location information
				pcity=rs("city")
					pcity = replace(pcity,"|","-")
				pstate=rs("state")
				pstateCode=rs("stateCode")
					if trim(pstateCode)="" then
						pstateCode=pstate
					end if
				pCountryCode=rs("CountryCode")
				
				set rs = nothing
				
				'// Transaction line example per Google Analytics documentation
				pcGAtransaction = 					"_gaq.push(['_addTrans', " & VbCrLf
    			pcGAtransaction = pcGAtransaction & "'" & scpre+int(pIdOrder) & "',           	// order ID - required " & VbCrLf
    			pcGAtransaction = pcGAtransaction & "'" & pcf_GoogleEscape(paffiliateName) & "',  				// affiliation or store name " & VbCrLf
    			pcGAtransaction = pcGAtransaction & "'" & ptotal & "',          				// total - required " & VbCrLf
    			pcGAtransaction = pcGAtransaction & "'" & ptaxAmount & "',           			// tax " & VbCrLf
    			pcGAtransaction = pcGAtransaction & "'" & pTotalShipping & "',              	// shipping " & VbCrLf
    			pcGAtransaction = pcGAtransaction & "'" & pcf_GoogleEscape(pcity) & "',       					// city " & VbCrLf
    			pcGAtransaction = pcGAtransaction & "'" & pcf_GoogleEscape(pstateCode) & "',     					// state or province " & VbCrLf
    			pcGAtransaction = pcGAtransaction & "'" & pCountryCode & "'             		// country " & VbCrLf
    			pcGAtransaction = pcGAtransaction & "]); " & VbCrLf
    			pcGAtransaction = pcGAtransaction & VbCrLf
			
			END IF ' 2 - End If database errors, skip everything
			call closeDb()			
		'// STEP 1 - END
		
		'// STEP 2 - START
										
			'// Gather item info from database
			call openDb()		
			query="SELECT ProductsOrdered.idProduct, ProductsOrdered.quantity, ProductsOrdered.unitPrice, products.description, products.sku FROM ProductsOrdered, products WHERE ProductsOrdered.idProduct=products.idProduct AND ProductsOrdered.idOrder=" & pOID
			set rs=Server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)
			
			if err.number <> 0 then
				set rs=nothing
				pGetItems = 0
				call closeDb()
			end If
			
			if rs.eof then
				set rs=nothing
				pGetItems = 0
				call closeDb()
			end if
					
			IF pGetItems<>0 THEN ' 3 - If database errors, skip everything
					
				'// Item line variables  
				'// [order-id]  Your internal unique order id number (should be same as transaction line)  
				'// [sku/code]  Product SKU code  
				'// [product name]  Product name or description  
				'// [category]  Category of the product or variation  
				'// [price]  Unit-price of the product  
				'// [quantity]  Quantity ordered
				
				pcGAtransactionItems = ""

				Do While Not rs.eof
					
					pIdProduct = rs("idProduct")			
					pSKU = rs("sku")
						pSKU = replace(pSKU,"|","-")
					pName = rs("description")
						pName = replace(pName,"|","-")
					pUnitPrice = rs("unitPrice")
					pQuantity = rs("quantity")
					
						'// Find category information
						query="SELECT idCategory FROM categories_products WHERE idProduct ="& pIdProduct
						set rsTemp=server.CreateObject("ADODB.RecordSet")
						set rsTemp=conntemp.execute(query)
						if not rsTemp.eof then
							idCategory=rsTemp("idCategory")
							query="SELECT categoryDesc FROM categories WHERE idCategory =" & idCategory
							set rsTemp=conntemp.execute(query)
							if err.number <> 0 then
								set rsTemp=nothing
								pCategory = "NA"
							end If
							if rsTemp.eof then
								response.write "yes"
								response.End()
								set rsTemp=nothing
								pCategory = "NA"
							end if
							pCategory = rsTemp("categoryDesc")
						else
							pCategory = "NA"
						end if
						pCategory = replace(pCategory,"|","-")
						set rsTemp=nothing
				
					'// Item line example per Google Analytics documentation
					pcGAtransactionItems = pcGAtransactionItems & "_gaq.push(['_addItem', " & VbCrLf
					pcGAtransactionItems = pcGAtransactionItems & "  '" & scpre+int(pIdOrder) & "', 	// order ID - required " & VbCrLf
					pcGAtransactionItems = pcGAtransactionItems & "  '" & pcf_GoogleEscape(pSKU) & "',           			// SKU/code " & VbCrLf
					pcGAtransactionItems = pcGAtransactionItems & "  '" & pcf_GoogleEscape(pName) & "',        			// product name " & VbCrLf
					pcGAtransactionItems = pcGAtransactionItems & "  '" & pcf_GoogleEscape(pCategory) & "',   			// category or variation " & VbCrLf
					pcGAtransactionItems = pcGAtransactionItems & "  '" & pUnitPrice & "',          	// unit price - required " & VbCrLf
					pcGAtransactionItems = pcGAtransactionItems & "  '" & pQuantity & "'               	// quantity - required " & VbCrLf
    				pcGAtransactionItems = pcGAtransactionItems & "]); " & VbCrLf
    				pcGAtransactionItems = pcGAtransactionItems & VbCrLf

				
				rs.movenext
				loop
				set rs = nothing
				
			END IF ' 3 - End If database errors, skip everything
		
		'// STEP 2 - END
		
		call closeDb()
		
		'// Write the hidden form
		%>
		<script type="text/javascript">
			<%=pcGAtransaction%>
			<%=pcGAtransactionItems%>
 			_gaq.push(['_trackTrans']); //submits transaction to the Analytics servers
		</script> 

<%
END IF ' 1

IF trim(scGoogleAnalytics)<>"" AND NOT IsNull(scGoogleAnalytics) THEN
%>
<form style="display:none;" method="post" name="gcaform">
<input type="hidden" id="analyticsdata" name="analyticsdata" value="">
</form>
<script type="text/javascript"> 
function setGoogleCheckout() {	
	_gaq.push(function() {var pageTracker = _gaq._getAsyncTracker(); setUrchinInputCode(pageTracker);});	
	var strUrl = "<%=buttonAction%>";	
	document.gcaform.action = strUrl;
	document.gcaform.submit();
} 
</script> 
<script src="https://checkout.google.com/files/digital/ga_post.js" type="text/javascript"></script>
<% 
END IF
%>
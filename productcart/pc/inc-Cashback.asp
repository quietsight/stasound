<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'// BING CASHBACK
IF LSCB_STATUS = "1" AND LSCB_KEY <>"" THEN ' 1

	%><!-- BING Cashback --><%
	'// Check for a cashback referral
	'// If referral set the affiliate session
	'// Do not clear session on order complete as there may be multiple purchases
	pcv_strIsCashback = getUserInput(request("cashback"),1)
	If len(pcv_strIsCashback)>0 Then
		Session("cashback")=1
	End If
	
	
	'// Check whether the code should be run
	'// This variable is set on orderComplete.asp
	'// The code is ONLY run when an order has been completed
	
	If pcv_intOrderComplete=1 AND Session("cashback")=1 Then ' 2
			
			dim pGetItemsCB, pcCBtransaction
			
			pGetItemsCB = 1
			
			call openDb()
			
			'// STEP 1: GENERATE CASHBACK CONSTANTS
			'// STEP 2: GENERATE ITEM INFO LINES
			
			'// STEP 1 - START
			
				'// Get Cashback info from db
			
				query="SELECT city, state, stateCode, CountryCode, shipmentDetails, idAffiliate, taxAmount, total, taxDetails, ord_VAT FROM orders WHERE idOrder=" & pOID
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=conntemp.execute(query)
				
				if err.number <> 0 then
					set rs=nothing
					pGetItemsCB = 0
					call closeDb()
				end If
				
				if rs.eof then
					set rs=nothing
					pGetItemsCB = 0
					call closeDb()
				end if			
	
					
			'// STEP 1 - END
			
			'// STEP 2 - START
											
				'// Gather item info from database
				
				query="SELECT ProductsOrdered.idProduct, ProductsOrdered.quantity, ProductsOrdered.unitPrice, products.description, products.sku FROM ProductsOrdered, products WHERE ProductsOrdered.idProduct=products.idProduct AND ProductsOrdered.idOrder=" & pOID
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=conntemp.execute(query)
				
				if err.number <> 0 then
					set rs=nothing
					pGetItemsCB = 0
					call closeDb()
				end If
				
				if rs.eof then
					set rs=nothing
					pGetItemsCB = 0
					call closeDb()
				end if
						
				IF pGetItemsCB<>0 THEN ' 3 - If database errors, skip everything
						
					'// Item line variables  
					'// [order-id]  Your internal unique order id number
					'// [price]  Unit-price of the product  
					'// [quantity]  Quantity ordered
					%>
					<!-- Begin BING cashback Tracking Pixel Code -->
					<script type='text/javascript'>
					<!--
						var jf_merchant_id = '<%=LSCB_KEY%>';
						var jf_merchant_order_num = '<%=scpre+int(pOID)%>';
						var jf_purchased_items = new Array(); 
						<%
						Do While Not rs.eof
							
							pIdProduct = rs("idProduct")			
							pSKU = rs("sku")
								pSKU = replace(pSKU,"|","-")
							pName = rs("description")
								pName = replace(pName,"|","-")
							pUnitPrice = rs("unitPrice")
							pQuantity = rs("quantity")
							%>
							
							// add cart item
							var jf_item = new Object();
							jf_item.mpi = '<%=pIdProduct%>';
							jf_item.price = '<%=pUnitPrice%>';
							jf_item.quantity = <%=pQuantity%>;
							jf_purchased_items.push(jf_item);
	
							<%					
						rs.movenext
						loop
						set rs = nothing
						%>
						//-->
						</script>
						<script type='text/javascript' src='https://www.jellyfish.com/javascripts/1x1tracking.js'></script>
						<!-- End Cashback Tracking Pixel Code -->
						<%
						'// Clear the Cashback Session
						'Session("cashback")=""
						
				END IF ' 3 - End If database errors, skip everything
			
			'// STEP 2 - END
			
			call closeDb()
	End If ' 2
	
END IF ' 1
%>
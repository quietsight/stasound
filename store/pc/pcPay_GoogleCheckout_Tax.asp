<%
'--------------------------------------------------------------
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<% 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// Start: Variables to Find and Set
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
pcDblServiceHandlingFee = 0			
pcDblIncHandlingFee = 0
discountTotal = ccur(0)
CatDiscTotal=0
taxPrdAmount=0
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// END: Variables to Find and Set
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
'///////////////////////////////////////////////////////////////////////////////////////
'// START: TAX ADJUSTMENTS
'///////////////////////////////////////////////////////////////////////////////////////
if zipCnt=1 then
		
		
	'*******************************************************************************
	'// START: TAX PER PRODUCT
	'*******************************************************************************		
	ppcCartIndex=Session("pcCartIndex")

	for f=1 to pcCartIndex
		if pcCartArray(f,10)=0 then
			query="SELECT taxPerProduct FROM taxPrd WHERE ((stateCode='" &StateCode& "' AND stateCodeEq=-1) OR (stateCode IS NULL) OR (stateCode<>'" &StateCode& "' AND stateCodeEq=0)) AND ((CountryCode='"&CountryCode&"' AND CountryCodeEq=-1) OR (CountryCode IS NULL) OR (CountryCode<>'" &CountryCode& "' AND CountryCodeEq=0)) AND ((zip='" &zip& "' AND zipEq=-1) OR (zip IS NULL) OR (zip<>'" &zip& "' AND zipEq=0)) AND idProduct=" & pcCartArray(f,0)
			set rsTax=server.CreateObject("ADODB.RecordSet")
			set rsTax=conntemp.execute(query)
			taxPrdArray=0
			if NOT rsTax.eof then
				do until rsTax.eof 
					taxPrdAmount=taxPrdAmount+(rsTax("taxPerProduct") * ( pcCartArray(f,2) * (pcCartArray(f,5) + pcCartArray(f,3)) )) 
					taxPrdArray=1  
				 rsTax.movenext
				loop
			end if
			set rsTax=nothing
			pcCartArray(f,24)=taxPrdArray
		end if 
	next
	'*******************************************************************************
	'// END: TAX PER PRODUCT
	'*******************************************************************************




	'*******************************************************************************
	'// START: TAX PER SHIP METHOD 
	'*******************************************************************************
	
	'// Calculate total price of the order, total weight and product total quantities

	'// Get Taxable Total	
	pshippingStateCode=Session("pcSFStateCode")
	pshippingCountryCode=Session("pcSFCountryCode")
	pTaxableTotal=ccur(calculateTaxableTotal(pcCartArray, ppcCartIndex))
	
	'// This will probably come from Google Constants to save query time
	pPaymentPriceToAdd=GOOGLEPRICETOADD
	pPaymentpercentageToAdd=GOOGLEPERCENTAGETOADD
	
	'// Add payment amount
	if ccur(pPaymentPriceToAdd)<>0 or ccur(pPaymentpercentageToAdd)<>0 then 
		tempTaxPercentageToAdd=(pPaymentpercentageToAdd*pTaxableTotal/100)
		tempTaxPercentageToAdd=roundTo(tempTaxPercentageToAdd,.01)
		taxPaymentTotal=pPaymentPriceToAdd + tempTaxPercentageToAdd 'processing fees on taxable total (only if percentage) 
		paymentTotal=pPaymentPriceToAdd + tempPercentageToAdd '// this is used for Discount Codes          
	end if			
	
	'// Start Reward Points
	'//
	
	'// Shipping Price to Add for Tax
	If Session(pshipDetailsArray2 & "_handling")<>"" AND isNULL(Session(pshipDetailsArray2 & "_handling"))=False Then
		pcvTempPriceToAdd = Session(pshipDetailsArray2) - Session(pshipDetailsArray2 & "_handling")
	Else
		pcvTempPriceToAdd = Session(pshipDetailsArray2)
	End If
	pcShipmentPriceToAdd = pcvTempPriceToAdd
	if pcShipmentPriceToAdd>0 then 
		pcDblShipmentTotal=pcShipmentPriceToAdd   
	else
		pcDblShipmentTotal=0
	end if
	'// Handling Price to Add for Tax
	pcHandlingPriceToAdd=Session(pshipDetailsArray2 & "_handling")
	if pcHandlingPriceToAdd<>"" then
		pcDblServiceHandlingFee=pcHandlingPriceToAdd
	else
		pcDblServiceHandlingFee=0
	end if
	
	
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// Start: Include Payment/ Shipping charges for Tax calculation
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	taxCalAmt=0			
	if TAX_SHIPPING_ALONE="NA" then
		If pTaxonCharges=1 then
			taxCalAmt=taxCalAmt+pcDblShipmentTotal
		End If
		If pTaxonFees=1 then
			taxCalAmt=taxCalAmt+pcDblServiceHandlingFee
		End If
	else
		if pcDblIncHandlingFee=0 then
			if TAX_SHIPPING_ALONE="Y" AND TAX_SHIPPING_AND_HANDLING_TOGETHER="Y" then
				taxCalAmt=taxCalAmt+pcDblShipmentTotal
				taxCalAmt=taxCalAmt+pcDblServiceHandlingFee
			end if
			if TAX_SHIPPING_ALONE="N" AND TAX_SHIPPING_AND_HANDLING_TOGETHER="Y" then
				taxCalAmt=taxCalAmt+pcDblServiceHandlingFee
			end if
		else
			if TAX_SHIPPING_AND_HANDLING_TOGETHER="Y" then
				taxCalAmt=taxCalAmt+pcDblShipmentTotal
				taxCalAmt=taxCalAmt+pcDblServiceHandlingFee
			else
				taxCalAmt=taxCalAmt+pcDblServiceHandlingFee
			end if
		end if
	end if
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// END: Include Payment/ Shipping charges for Tax calculation
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// Start: More
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	if int(taxCalAmt)=0 AND int(pTaxableTotal)=0 then
		ptaxLocAmount=0
	else
		
		'GGG Add-on start
		if Session("Cust_GW")="1" then
			pTaxableTotal=pTaxableTotal+ccur(GWTotal)
		end if
		'GGG Add-on end
		
		'if VAT
		VATTotal=0
		if taxPaymentTotal="" then
			taxPaymentTotal=0
		end if
		if ptaxVAT="1" then
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// Customer is using VAT SETTINGS
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			IF GCAmount=0 then
				VatTaxedAmount=(pTaxableTotal+taxCalAmt+taxPaymentTotal-discountTotal-CatDiscTotal)
			Else
				VatTaxedAmount=(pTaxableTotal+taxCalAmt+taxPaymentTotal)
			End if
			noVATTotal=VatTaxedAmount/(1+(ptaxVATrate/100))
			noVATTotal=RoundTo(noVATTotal,.01)
			VATTotal=RoundTo(VatTaxedAmount-noVATTotal,.01)
			VATTaxedAmount=RoundTo(VatTaxedAmount,.01)
			if VATTaxedAmount<0 then
				VATTaxedAmount=0
			end if
			
			'// set total
			ptaxLocAmount = VATTotal
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// END VAT SETTINGS
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
		else
			if ptaxseparate="1" then
				
				ptaxLocAmount="1"
				
				for y = 1 to session("taxCnt")					
					ptaxLoc=session("tax"&y)
					'GGG Add-on start
					IF GCAmount=0 then
						tempTAmt=((pTaxableTotal+taxCalAmt+taxPaymentTotal-discountTotal-CatDiscTotal) * ptaxLoc)
					ELSE
						tempTAmt=((pTaxableTotal+taxCalAmt+taxPaymentTotal) * ptaxLoc)
					END IF
					'GGG Add-on end
					
					tempTAmt=roundTo(tempTAmt,.01)
					if tempTAmt<0 then
						tempTAmt=0
					end if
					ptaxLocAmount=ptaxLocAmount+tempTAmt
				next
				
			else
				
				'GGG Add-on start
				IF GCAmount=0 then
					ptaxLocAmount=((pTaxableTotal+taxCalAmt+taxPaymentTotal-discountTotal-CatDiscTotal) * ptaxLoc)
					ptaxLocNoShipAmount=((pTaxableTotal+taxPaymentTotal-discountTotal-CatDiscTotal) * ptaxLoc)
				ELSE
					ptaxLocAmount=((pTaxableTotal+taxCalAmt+taxPaymentTotal) * ptaxLoc)
					ptaxLocNoShipAmount=((pTaxableTotal+taxPaymentTotal) * ptaxLoc)
					END IF				
				'GGG Add-on end
				
				'// Tax with Shipping
				ptaxLocAmount=RoundTo(ptaxLocAmount,.01)
				if ptaxLocAmount<0 then
					ptaxLocAmount=0
				end if
			
				'// Tax without Shipping
				ptaxLocNoShipAmount=RoundTo(ptaxLocNoShipAmount,.01)
				if ptaxLocNoShipAmount<0 then
					ptaxLocNoShipAmount=0
				end if

			end if		
		end if
	end if
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// END: More
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						
	'// The total tax value [Note: find when this is needed "VATTotal"]
	pTaxAmount = taxPrdAmount + ptaxLocAmount 	
	pTaxNoShipAmount = taxPrdAmount + ptaxLocNoShipAmount		
	
	'*******************************************************************************
	'// END: TAX PER SHIP METHOD 
	'*******************************************************************************
	
	Session(pshipDetailsArray2 & "_tax") = pcf_CurrencyField(pTaxAmount)
	Session(pshipDetailsArray2 & "_tax2") = pcf_CurrencyField(pTaxNoShipAmount)				
	
end if
'///////////////////////////////////////////////////////////////////////////////////////
'// END: TAX ADJUSTMENTS
'///////////////////////////////////////////////////////////////////////////////////////
%>
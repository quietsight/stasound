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
'///////////////////////////////////////////////////////////////////////////////////
'//  START: ORDER PROCESSING
'///////////////////////////////////////////////////////////////////////////////////

'***********************************************************************************
' START: Get info we need... segway into PC "SaveOrd.asp" code for managability
'***********************************************************************************
pcv_strFirstName=replace(pcv_strFirstName,"''","'")
pcv_strLastName=replace(pcv_strLastName,"''","'")
pcv_strBillingCompanyName=replace(pcv_strBillingCompanyName,"''","'")
pcv_strBillingAddress1=replace(pcv_strBillingAddress1,"''","'")
pcv_strBillingAddress2=replace(pcv_strBillingAddress2,"''","'")
pcv_strBillingCity=replace(pcv_strBillingCity,"''","'")

'// Billing
pFirstName=getUserInput(pcv_strFirstName,0)
pLastName=getUserInput(pcv_strLastName,0)
pCustomerCompany=getUserInput(pcv_strBillingCompanyName,100)
pPhone=pcv_strBillingPhone
pEmail=pcv_strBillingEmail
pAddress=getUserInput(pcv_strBillingAddress1,0)
pZip=pcv_strBillingPostalCode
pStateCode=pcv_BillingStateCode
pState=getUserInput(pcv_BillingState,0)
pCity=getUserInput(pcv_strBillingCity,0)
pCountryCode=pcv_strBillingCountryCode
pAddress2=getUserInput(pcv_strBillingAddress2,0)
pFax=pcv_strBillingFax

'// Shipping
pcv_strShippingFirstName = Left(pcv_strShippingContactName,(instr(pcv_strShippingContactName," ")-1))
pcv_strShippingLastName =  Right(pcv_strShippingContactName,(len(pcv_strShippingContactName)-instr(pcv_strShippingContactName," ")))	
pShippingFirstName=getUserInput(pcv_strShippingFirstName,0)
pShippingLastName=getUserInput(pcv_strShippingLastName,0)
pShippingCompany=getUserInput(pcv_strShippingCompanyName,0)
pShippingAddress=getUserInput(pcv_strShippingAddress1,0)
pShippingAddress2=getUserInput(pcv_strShippingAddress2,0)
pShippingCity=getUserInput(pcv_strShippingCity,0)
pShippingStateCode=getUserInput(pcv_ShippingStateCode,0)
pShippingState=getUserInput(pcv_ShippingState,0)
pShippingZip=getUserInput(pcv_strShippingPostalCode,0)
pShippingCountryCode=getUserInput(pcv_strShippingCountryCode,0)
pShippingPhone=getUserInput(pcv_strShippingPhone,0)

if pZip="" then
	pZip="NA"
end if
if pShippingZip="" then
	pShippingZip="NA"
end if
'***********************************************************************************
' END: Get info from sessions and customers
'***********************************************************************************



'***********************************************************************************
' START: Order Information
'***********************************************************************************
'// Residential Flag (1 or 0)
if isNULL(ShippingResidential) then
	pOrdShipType=1
else
	pOrdShipType=getUserInput(ShippingResidential,0)
end if

if pOrdShipType="0" then 'flagged as commercial
	pOrdShipType=1 'enter it as commercial
else
	pOrdShipType=0 'residential (default)
end if

'// Package Details
pOrdPackageNum = pcv_intPackageNum '// We will grab this from the customer sessions table, which we saved at Start
if pOrdPackageNum="" then
	pOrdPackageNum=1
end if

'// Misc.
pShippingNickName="Google Checkout Default"
pIDRefer=0 '// This will be "0" with Google
pRewardsBalance=0 '// The balance remaining on the rewards account

'// Tax Information
pTaxShippingAlone=""
pTaxShipppingAndHandlingTogether=""
pTaxCountyCode=""
pTaxProductAmount=""

'// Shipping Data
query="SELECT serviceDescription, serviceHandlingFee, serviceShowHandlingFee, serviceCode FROM shipService WHERE (((shipService.serviceActive)=-1)) ORDER BY shipService.servicePriority;"
set rsShippingObj=Server.CreateObject("ADODB.RecordSet")
set rsShippingObj=connTemp.execute(query)
if NOT rsShippingObj.eof then
	do while NOT rsShippingObj.eof
		pcv_strServiceDescription = rsShippingObj("serviceDescription")
		pcv_strServiceDescription= replace(pcv_strServiceDescription,"<sup>&reg;</sup>","")
		pcv_strServiceDescription= replace(pcv_strServiceDescription,"&reg;","")
		pcv_strServiceDescription= replace(pcv_strServiceDescription,"<sup>SM</sup>","")
		if pcv_strServiceDescription = pcv_strShippingName then
			pcv_strServiceHandlingFee = rsShippingObj("serviceHandlingFee")
			pcv_strServiceShowHandlingFee = rsShippingObj("serviceShowHandlingFee")
			pcv_strServiceCode = rsShippingObj("serviceCode")
			'exit do
		end if
	rsShippingObj.movenext
	loop
	set rsShippingObj=nothing
end if
if pcv_strServiceHandlingFee="" then
	pcv_strServiceHandlingFee = "0"
end if
if pcv_strServiceShowHandlingFee="" then
	pcv_strServiceShowHandlingFee = "0"
end if
if pcv_strServiceCode="" then
	pcv_strServiceCode = "NA"
end if	
'response.end

if instr(pcv_strShippingName,"FEDEX")>0 then
	pcv_strShippingCarrier="FedEx"
elseif instr(pcv_strShippingName,"UPS")>0 then
	pcv_strShippingCarrier="UPS"
else
	pcv_strShippingCarrier="Other"
end if

'// We over-ride the Handling Fee because it has already been added into the Shipping Cost at Google
pcv_strServiceHandlingFee=0
pShipping=pcv_strShippingCarrier &","& pcv_strShippingName &","& pcv_strShippingCost &","& pcv_strServiceHandlingFee &","& pcv_strServiceShowHandlingFee &","& pcv_strServiceCode &"" '// Shipping Array
pComments="" '// This will be empty with Google
pShippingReferenceId="0" '// This will be "0" with Google
pShippingFax="" '// This will be empty with Google
pShippingEmail="" '// This will be empty with Google
pShippingFullName=pShippingFirstName& " "&pShippingLastName

'// Discount Flag
pDiscountUsed=""
'***********************************************************************************
' END: Order Information
'***********************************************************************************




'***********************************************************************************
' START: ADDITIONAL ORDER INFORMATION
'***********************************************************************************
'// Discountcode, chkPayment, discountc, idPayment, taxDetailsString, VATTotal, discountAmount
pIdPayment= 0 '// This will always indicate Google at zero
pVATTotal= ""
pDiscountCode = Session("DiscountCode") 
intCodeCnt = (Session("TotalCodesUsed")-1) ' The number of discount codes used
discountAmount = Session("DiscountTotal")
ptaxDetailsString = "" ' The tax details as a string

'// Order Total
pTotal=pcv_strOrderTotal

'// GGG Add-on Total
pGWTotal= 0

'// Tax Total
pTaxAmount = pcv_strMerchantCalculationTax 'session("taxAmount")

'// Rewards Total
If session("pcSFRewardsDollarValue")<>"" then
	piRewardValue = session("pcSFRewardsDollarValue")
	session("pcSFRewardsDollarValue")=""
End if
'***********************************************************************************
' END: ADDITIONAL ORDER INFORMATION
'***********************************************************************************





'***********************************************************************************
' START: REWARD POINTS
'***********************************************************************************
'If it is NOT the First purchase by this visitor the Session var is null
pRewardReferral=0
pRewardRefId=0

If Session("ContinueRef")<>"" then
	If Session("ContinueRef") > 0 And RewardsReferral=1 Then
		pRewardRefId=Session("ContinueRef")
		If RewardsFlat=1 Then
			pRewardReferral=RewardsFlatValue
		End If
		If RewardsPerc=1 Then
			pRewardReferral=(pOrderTotal * (RewardsPercValue / 100))
		End If
	End If
End If	
'End Referral Rewards
'***********************************************************************************
' END: REWARD POINTS
'***********************************************************************************



'***********************************************************************************
' START: AFFILIATE INFO
'***********************************************************************************
session("idAffiliate")=pcv_strAffiliateID 
if NOT len(session("idAffiliate"))>0 then
	session("idAffiliate") = Cint(1)
end if
'***********************************************************************************
' END: AFFILIATE INFO
'***********************************************************************************




'***********************************************************************************
' START: REBUILD THE CART
'***********************************************************************************

'// Start Sessions
'session("idPCStore")= scID
session("idCustomer")=pcv_intCustomerId   
session("language")=Cstr("english")
session("pcCartIndex")=Cint(0)

	
'// RePopulate the Shopping Cart
pcs_RestoreCartArray		
session("pcCartSession")=pcCartArray 
pcCartIndex=f-1
session("pcCartIndex")=pcCartIndex
ppcCartIndex=Session("pcCartIndex")
'***********************************************************************************
' END: REBUILD THE CART
'***********************************************************************************



'***********************************************************************************
' START: AFFILIATES
'***********************************************************************************
'// Retrieve affiliate ID from session
pIdAffiliate=session("idAffiliate")
if pIdAffiliate="" then
	pIdAffiliate=1
end if

IF pIdAffiliate<>1 THEN
	'// Determine whether the affiliate should be associated with this order
	pcInt_AllowedAffOrders=scAllowedAffOrders
	'// If 0, then unlimited orders are allowed.
	If pcInt_AllowedAffOrders = 0 then
		pcInt_AffiliateOK = 1
	else
		'// Find out if the same affiliate has referred this customer before
		query="SELECT idOrder FROM orders WHERE idAffiliate="&pIdAffiliate&" AND idCustomer="&session("idCustomer")&" AND orderStatus>1 AND orderStatus<>5"
		set rs=server.CreateObject("ADODB.RecordSet")
		rs.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText
		totalAffiliateOrders=clng(rs.RecordCount)
		set rs=nothing
		'// Check the number of orders against the max that the affiliate can earn
		'// commissions on
		if clng(totalAffiliateOrders) <= clng(pcInt_AllowedAffOrders) then
			pcInt_AffiliateOK = 1
		else
			pcInt_AffiliateOK = 0
		end if
	end if
		
		'// Check for customer type and exclude wholesale customers, if feature is active
		Dim pcInt_ExcludeWholesaleAff
		pcInt_ExcludeWholesaleAff = scExcludeWholesaleAff
		if pcInt_ExcludeWholesaleAff="" or not validNum(pcInt_ExcludeWholesaleAff) then pcInt_ExcludeWholesaleAff = 1
		if session("customerType")=1 and scExcludeWholesaleAff="1" then pcInt_AffiliateOK = 0
		
		if pcInt_AffiliateOK=0 then
			pIdAffiliate=1
		end if
	
		'// START Troubleshooting Area: write useful affiliate variables to the page
		'response.write "totalAffiliateOrders="&totalAffiliateOrders&"<br>"
		'response.write "pcInt_AffiliateOK="&pcInt_AffiliateOK&"<br>"
		'response.write "pcInt_ExcludeWholesaleAff="&pcInt_ExcludeWholesaleAff&"<br>"
		'response.write "customerType="&session("customerType")&"<br>"
		'response.write "pIdAffiliate="&pIdAffiliate
		'response.End()

END IF
'***********************************************************************************
' END: AFFILIATES
'***********************************************************************************



'***********************************************************************************
' START: VARIABLES
'***********************************************************************************
'// Details
pDetails=Cstr("")
'// Totals
pSubtotal=Cdbl(calculateCartTotal(pcCartArray, ppcCartIndex))
pCartTotalWeight=int(calculateCartWeight(pcCartArray, ppcCartIndex))
pCartQuantity=int(calculateCartQuantity(pcCartArray, ppcCartIndex))
pShipWeight=Cdbl(calculateShipWeight(pcCartArray, ppcCartIndex))
pAffiliateSubTotal=pSubtotal
'// Date
pDateOrder=Date()
if SQL_Format="1" then
	pDateOrder=Day(pDateOrder)&"/"&Month(pDateOrder)&"/"&Year(pDateOrder)
else
	pDateOrder=Month(pDateOrder)&"/"&Day(pDateOrder)&"/"&Year(pDateOrder)
end if
'// State and Province
If pStateCode <> "" and (pCountryCode="US" or pCountryCode="CA") then
	pState=""
end if
If pShippingStateCode <> "" and (pShippingCountryCode="US" or pShippingCountryCode="CA") then
	pShippingState=""
end if
'***********************************************************************************
' END: VARIABLES
'***********************************************************************************




'***********************************************************************************
' START: COMPILE MEMO FIELD
'***********************************************************************************
for f=1 to ppcCartIndex 
	' if item is not deleted from cart 
	if pcCartArray(f,10) = 0 then 
		tempAmt=Cdbl( pcCartArray(f,2) * (pcCartArray(f,5)+pcCartArray(f,3)) )
		if scDecSign="," then
			tempAmt=replace(tempAmt,",",".")
		end if
		pDetails	= pDetails & "  Amount: ||"& tempAmt & " Qty:" &pcCartArray(f,2)& "  SKU #:" &pcCartArray(f,7) & " - " &pcCartArray(f,1)& " " & pcCartArray(f,4) & Vbcrlf      
		pDetails = replace(pDetails,"'","''")
		pDetails=replace(pDetails,"''''","''")    
	end if ' item deleted
next
'***********************************************************************************
' END: COMPILE MEMO FIELD
'***********************************************************************************




'***********************************************************************************
' START: SHIPMENT DATA
'***********************************************************************************
If Session("nullShipper")="Yes" then
	pShipmentDesc=ship_dictLanguage.Item(Session("language")&"_noShip_a")
	pShipmentPriceToAdd="0"
else
	if Session("nullShipRates")="Yes" then
		pShipmentDesc=ship_dictLanguage.Item(Session("language")&"_noShip_b")
		pShipmentPriceToAdd="0"
	else
		shipping=split(pShipping,",")
	
		Shipper=shipping(0)
		Service=shipping(1)
		Postage=shipping(2)
		pShipmentDesc=Shipper&" - "&Service
		pShipmentPriceToAdd=Postage

		if ubound(shipping)=>3 then
			pserviceHandlingFee=shipping(3)
		else
			pserviceHandlingFee="0"
		end if	
	end if
end if
'***********************************************************************************
' END: SHIPMENT DATA
'***********************************************************************************




'***********************************************************************************
' START: SHIPMENT TOTALS
'***********************************************************************************
if pShipmentPriceToAdd>0 then 
	shipmentTotal=pShipmentPriceToAdd       
end if
pPSubtotal=pSubtotal
pSubtotal=pSubtotal + shipmentTotal+pserviceHandlingFee
pSRF="0"
If Session("nullShipper")="Yes" then
	pShipmentDetails=ship_dictLanguage.Item(Session("language")&"_noShip_a")
else
	if Session("nullShipRates")="Yes" then
		pSRF="1"
		pShipmentDetails=ship_dictLanguage.Item(Session("language")&"_noShip_b")
	else
		pShipmentDetails=pShipping
		pShipmentDetails=replace(pShipmentDetails,"<SUP>SM</SUP>","")
		pShipmentDetails=replace(pShipmentDetails,"<SUP>&reg;</SUP>","")
		pShipmentDetails=replace(pShipmentDetails,"&reg;","")
	end if
end if
'***********************************************************************************
' END: SHIPMENT TOTALS
'***********************************************************************************




'***********************************************************************************
' START: PAYMENT DETAILS
'***********************************************************************************
if pIdPayment=0 then
	
	pPaymentDetails = "Google Checkout || 0.00"
	pPaymentDesc="Google Checkout"
	pPaymentPriceToAdd=GOOGLEPRICETOADD
	pPaymentpercentageToAdd=GOOGLEPERCENTAGETOADD


	' calculate payment price
	if Cdbl(pPaymentPriceToAdd)<>0 or Cdbl(pPaymentpercentageToAdd)<>0 then 
		tempPercentageToAdd=(pPaymentpercentageToAdd*pPSubtotal/100)
		tempPercentageToAdd=roundTo(tempPercentageToAdd,.01)
		paymentTotal=pPaymentPriceToAdd + tempPercentageToAdd
	end if

	pSubtotal=pSubtotal + paymentTotal
	
	pPaymentDetails = pPaymentDesc & " || "& paymentTotal
	if scDecSign="," then
		pPaymentDetails=replace(pPaymentDetails,",",".")
	end if
	
end if
'***********************************************************************************
' END: PAYMENT DETAILS
'***********************************************************************************




'***********************************************************************************
' START: DISCOUNTS
'***********************************************************************************
if pDiscountCode<>"" then 

	myTest=0
	pDiscountDetails=Cstr("")
	discountTotal=Cdbl(0)

	if instr(pDiscountCode,",")>0 then
		myTest=0
	else
	 	myTest=1
	end if

	'// There are discount code(s)
	IF myTest=0 THEN

		DiscountCodeArry=Split(pDiscountCode,",")
		DiscountAmountArry=split(discountAmount,",")
	
		dim intDiscountUsedCnt
		intDiscountUsedCnt=0
		intDiscountArryCnt=0

		For i=0 to intCodeCnt

			pTempDiscCode=DiscountCodeArry(i)
			pTempDiscAmount=DiscountAmountArry(i)

			if pTempDiscCode<>"" then
			
				query="SELECT quantityFrom,quantityUntil,weightFrom,weightUntil,priceFrom,priceUntil,idDiscount,oneTime,discountDesc,priceToDiscount,percentageToDiscount FROM discounts WHERE discountcode='" & pTempDiscCode &"'"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				if NOT rs.eof then
					intQuantityFrom=rs("quantityFrom")
					intQuantityUntil=rs("quantityUntil")
					intWeightFrom=rs("weightFrom")
					intWeightUntil=rs("weightUntil")
					dblPriceFrom=rs("priceFrom")
					dblPriceUntil=rs("priceUntil")
					pIdDiscount=rs("idDiscount")
					pOneTime=rs("oneTime")
					pDiscountDesc=rs("discountDesc")
					pPriceToDiscount=rs("priceToDiscount")
					pPercentageToDiscount=rs("percentageToDiscount")
				end if 
		 
				if pPriceToDiscount>0 or ppercentageToDiscount>0 then 
					discountTotal=Session("DiscountCodeTotal")        
				end if 
				
				if intDiscountArryCnt=0 then
					pDiscountDetails=pDiscountDetails&pDiscountDesc & " - || "& pTempDiscAmount
					intDiscountArryCnt=intDiscountArryCnt+1
				else
					pDiscountDetails=pDiscountDetails&","&pDiscountDesc & " - || "& pTempDiscAmount
					intDiscountArryCnt=intDiscountArryCnt+1
				end if
				
			end if '// if pTempDiscCode<>"" then
		Next
		
	ELSE
	
		' GGG
		GCAmount=0

		query="SELECT pcGCOrdered.pcGO_ExpDate,pcGCOrdered.pcGO_Amount,pcGCOrdered.pcGO_Status,products.Description FROM pcGCOrdered,products WHERE pcGCOrdered.pcGO_GcCode='"&pDiscountCode&"' AND products.idproduct=pcGCOrdered.pcGO_IDProduct"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		IF rs.eof then
			pDiscountCode=""		
		ELSE
			mTest=0
			pGCExpDate=rs("pcGO_ExpDate")
			pGCAmount=rs("pcGO_Amount")
			if pGCAmount<>"" then
			else
				pGCAmount=0
			end if
			pGCStatus=rs("pcGO_Status")
			pDiscountDesc=rs("Description")
		
			if cdbl(pGCAmount)<=0 then
				mTest=1
				pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5")
			end if
			if cint(pGCStatus)<>1 then
				mTest=1
				pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_4")
			end if
			if year(pGCExpDate)<>"1900" AND Date()>pGCExpDate then
				mTest=1
				pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5")
			end if
	
			if mTest=0 then
				'// Have Available Amount			
				GCAmount=cdbl(Session("DiscountCodeTotal"))
				if GCAmount<>"" then
				else
					GCAmount=0
				end if		
				pGCAmount=pGCAmount-GCAmount
				if pGCAmount<0 then
					pGCAmount=0
				end if
				if pGCAmount=0 then
					pGCStatus=0
				end if				
			end if
		END IF 'rs.eof
		
	END IF

	discountTotal=Cdbl(0)
	 
	if pPriceToDiscount>0 or ppercentageToDiscount>0 or GCAmount>0 then 
		discountTotal=Session("DiscountCodeTotal")        
	end if 
 
	pSubtotal = pSubtotal - discountTotal
	
	if GCAmount>0 then
		pDiscountDetails=dictLanguage.Item(Session("language")&"_saveorder_10")
	end if
		
else

 	pDiscountDetails=dictLanguage.Item(Session("language")&"_saveorder_10")
	
end if
'***********************************************************************************
' END: DISCOUNTS
'***********************************************************************************




'***********************************************************************************
' START: AFFILIATE DETAILS
'***********************************************************************************
pAffiliateValid =Cint(1)
pAffiliatePay=0
if pIdAffiliate<>1 then
 	query="SELECT commission FROM affiliates WHERE idAffiliate=" &pIdAffiliate
	set rs=server.CreateObject("ADODB.RecordSet")
 	set rs=conntemp.execute(query) 
 	if rs.eof then
  		pAffiliateValid=0
 	else
		'// Calculate Affiliate Pay
		pAffiliateSubTotal=pAffiliateSubTotal - discountTotal
		pAffiliatePay=pAffiliateSubTotal * (rs("commission")/100)
	end if
	set rs=nothing
end if 
if pAffiliateValid=0 then
	pIdAffiliate=1
	pAffiliatePay=0
end if
'***********************************************************************************
' END: AFFILIATE DETAILS
'***********************************************************************************




'***********************************************************************************
' START: REWARD DETAILS
'***********************************************************************************
If piRewardValue="" then
	piRewardValue="0"
End If
If Session("pcSFUseRewards")="" then
	Session("pcSFUseRewards")="0"
End If

'// Save order temporarily
IDrefer=session("IDrefer")
if isNull(IDrefer) OR IDrefer="" then
	IDrefer="0"
end if

pord_DeliveryDate=session("DF1") & " " & session("TF1")
pord_DeliveryDate=trim(pord_DeliveryDate)
if not isDate(pord_DeliveryDate) then
	pord_DeliveryDate=""
end if

pord_OrderName=session("pord_OrderName")

pcv_CatDiscounts=Session("CatDiscTotal")
if isNull(pcv_CatDiscounts) OR pcv_CatDiscounts="" then
	pcv_CatDiscounts="0"
end if
'***********************************************************************************
' END: REWARD DETAILS
'***********************************************************************************





'***********************************************************************************
' START: GENERATE ORDER UPDATE QUERY
'***********************************************************************************
strUpdateQuery="UPDATE orders SET IDrefer=" & IDrefer & ","
if scDB="SQL" then
	strUpdateQuery=strUpdateQuery&"orderDate='" & pDateOrder  & "',"
else
	strUpdateQuery=strUpdateQuery&"orderDate=#" & pDateOrder  & "#,"
end if	

strUpdateQuery=strUpdateQuery&"idCustomer=" & int(Session("idCustomer"))& ", details='" &pDetails &"', total=" &replacecomma(pTotal)& ", taxAmount=" &replacecomma(pTaxAmount)& ", comments='" &pComments& "', address='" &paddress & "', zip='" &pzip& "',state='" &pState& "',stateCode='" &pStateCode& "',city='" &pCity& "',CountryCode='" &pCountryCode& "',shippingAddress='" &pShippingAddress & "',shippingZip='" &pShippingZip& "',shippingState='" &pShippingState& "',shippingStateCode='" &pShippingStateCode& "', shippingCity='" &pShippingCity& "', shippingCountryCode='" &pShippingCountryCode& "',shipmentDetails='" &pShipmentDetails& "', paymentDetails='" &replace(pPaymentDetails,"'","''")& "',discountDetails='" &replace(pDiscountDetails,"'","''")& "',randomNumber=" & session("pcSFIdDbSession") & ",orderStatus=1,pcOrd_shippingPhone=' " &pShippingPhone& "',idAffiliate=" &pIdAffiliate& ", affiliatePay=" &replacecomma(pAffiliatePay)&",shippingFullName='"&pShippingFullName&"', iRewardPoints="& 0 &",iRewardValue= " &piRewardValue&", iRewardRefid=" &pRewardRefId&", iRewardPointsRef=" &pRewardReferral&", iRewardPointsCustAccrued=" & 0 &", address2='" &paddress2 & "', shippingCompany='" &pShippingCompany & "', shippingAddress2='" &pShippingAddress2 & "',taxDetails='"&ptaxDetailsString&"',SRF="&pSRF&",ordShipType="&pOrdShipType&", ordPackageNum="&pOrdPackageNum&", ord_OrderName='"&pord_OrderName&"'"
if DFShow="1"  and pord_DeliveryDate <> "" then
	if scDB="SQL" then
		strUpdateQuery=strUpdateQuery&",ord_DeliveryDate='" & pord_DeliveryDate  & "'"
	else
		strUpdateQuery=strUpdateQuery&",ord_DeliveryDate=#" & pord_DeliveryDate  & "#"
	end if
end if
if pVATTotal="" then
	pVATTotal=0
end if
'GGG Add-on start
if GCAmount<>"" then
else
	GCAmount=0
end if
if session("Cust_IDEvent")<>"" then
	gIDEvent=session("Cust_IDEvent")
else
	gIDEvent="0"
end if
pcv_GcReName=session("Cust_GcReName")
pcv_GcReEmail=session("Cust_GcReEmail")
pcv_GcReMsg=session("Cust_GcReMsg")
if GCAmount=0 then
	pDiscountCode=""
end if
'GGG Add-on end
strUpdateQuery=strUpdateQuery&",ord_VAT="&replacecomma(pVATTotal)&",pcord_CatDiscounts=" & pcv_CatDiscounts & ",pcOrd_DiscountsUsed='"&pDiscountUsed&"',pcOrd_GcCode='" & pDiscountCode & "',pcOrd_GcUsed=" & GCAmount & ",pcOrd_GCs=0,pcOrd_IDEvent=" & gIDEvent & ",pcOrd_GWTotal=" & pGWTotal & ",pcOrd_GcReName='" & pcv_GcReName & "',pcOrd_GcReEmail='" & pcv_GcReEmail & "',pcOrd_GcReMsg='" & pcv_GcReMsg & "',pcOrd_shippingFax='"&pShippingFax&"', pcOrd_ShippingEmail='"&pShippingEmail&"', pcOrd_ShipWeight="&pShipWeight&" "
'***********************************************************************************
' END: GENERATE ORDER UPDATE QUERY
'***********************************************************************************



'***********************************************************************************
' START: GENERATE ORDER INSERT QUERY
'***********************************************************************************
strInsertQuery="INSERT INTO orders (pcOrd_GoogleIDOrder, IDrefer,orderDate,idCustomer, details, total, taxAmount, comments, address, zip, state, stateCode, city, CountryCode, shippingAddress, shippingZip, shippingState, shippingStateCode, shippingCity, shippingCountryCode, shipmentDetails, paymentDetails, discountDetails, randomNumber, orderStatus, pcOrd_shippingPhone, idAffiliate, affiliatePay,shippingFullName, iRewardPoints, iRewardValue,iRewardRefid,iRewardPointsRef,iRewardPointsCustAccrued, address2, shippingCompany, shippingAddress2,taxDetails,SRF,ordShipType, ordPackageNum, ord_OrderName"
if DFShow="1"  and pord_DeliveryDate <> "" then
	strInsertQuery=strInsertQuery&",ord_DeliveryDate"
end if
strInsertQuery=strInsertQuery&",ord_VAT,pcord_CatDiscounts,pcOrd_DiscountsUsed,pcOrd_GcCode,pcOrd_GcUsed,pcOrd_GCs,pcOrd_IDEvent,pcOrd_GWTotal,pcOrd_GcReName,pcOrd_GcReEmail,pcOrd_GcReMsg,pcOrd_shippingFax, pcOrd_ShippingEmail, pcOrd_ShipWeight) VALUES ('" & pcv_strOrderNumber & "', " & IDrefer & ","
if scDB="SQL" then
	strInsertQuery=strInsertQuery&"'" & pDateOrder  & "'"
else
	strInsertQuery=strInsertQuery&"#" & pDateOrder  & "#"
end if

strInsertQuery=strInsertQuery&"," & int(Session("idCustomer"))& ",'" &pDetails &"'," &replacecomma(pTotal)& "," &replacecomma(pTaxAmount)& ",'" &pComments& "','" &paddress & "','" &pzip& "','" &pState& "','" &pStateCode& "','" &pCity& "','" &pCountryCode& "','" &pShippingAddress & "','" &pShippingZip& "','" &pShippingState& "','" &pShippingStateCode& "','" &pShippingCity& "','" &pShippingCountryCode& "','" &pShipmentDetails& "','" &replace(pPaymentDetails,"'","''")& "','" &replace(pDiscountDetails,"'","''")& "',"& session("pcSFIdDbSession") &", 1,' " &pShippingPhone& "'," &pIdAffiliate& ", " &replacecomma(pAffiliatePay)&",'"&pShippingFullName&"', "& 0 &", " &piRewardValue&", " &pRewardRefId&", " &pRewardReferral&", "& 0 &", '" &paddress2 & "', '" &pShippingCompany & "', '" &pShippingAddress2 & "','"&ptaxDetailsString&"',"&pSRF&","&pOrdShipType&","&pOrdPackageNum&",'"&pord_OrderName&"'"
if DFShow="1" and pord_DeliveryDate <> "" then
	if scDB="SQL" then
		strInsertQuery=strInsertQuery&",'" & pord_DeliveryDate  & "'"
	else
		strInsertQuery=strInsertQuery&",#" & pord_DeliveryDate  & "#"
	end if
end if
if pVATTotal="" then
	pVATTotal=0
end if
strInsertQuery=strInsertQuery&","&replacecomma(pVATTotal)&"," & pcv_CatDiscounts & ",'"&pDiscountUsed&"','" & pDiscountCode & "'," & GCAmount & ",0," & gIDEvent & "," & pGWTotal & ",'" & pcv_GcReName & "','" & pcv_GcReEmail & "','" & pcv_GcReMsg & "', '"&pShippingFax&"', '"&pShippingEmail&"', "&pShipWeight&")"
'***********************************************************************************
' END: GENERATE ORDER INSERT QUERY
'***********************************************************************************



'***********************************************************************************
' START: RUN ORDER QUERY
'***********************************************************************************		

'// Update Order
strUpdateQuery=strUpdateQuery&"WHERE pcOrd_GoogleIDOrder='"& pcv_strOrderNumber &"';"

'// Check if Order is already saved
query="SELECT idOrder FROM orders WHERE pcOrd_GoogleIDOrder='"& pcv_strOrderNumber & "' AND idCustomer=" & session("idCustomer") & ";"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

If NOT rs.eof Then
	
	session("idOrderSaved")=rs("idOrder")
	
	'// Update the Order
	set rs=conntemp.execute(strUpdateQuery)
	
	'// Delete Products Ordered
	query="DELETE FROM ProductsOrdered WHERE idOrder="& session("idOrderSaved") &";"
	set rs=conntemp.execute(query)
	
Else

	'// Insert Order
	set rs=conntemp.execute(strInsertQuery)
	
End If

set rs=nothing

'***********************************************************************************
' END: RUN ORDER QUERY
'***********************************************************************************


    '// <new-order-notification>
    sendNotificationAcknowledgment


'***********************************************************************************
' START: GET ORDER ID
'***********************************************************************************
query="SELECT idOrder FROM orders WHERE pcOrd_GoogleIDOrder='" & pcv_strOrderNumber & "' AND idCustomer=" & session("idCustomer") & ";"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)		
	pIdorder=rs("idOrder")
	session("idOrderSaved")=pIdorder
set rs=nothing
'***********************************************************************************
' END: GET ORDER ID
'***********************************************************************************




'***********************************************************************************
' START: ORDER STATUS
'***********************************************************************************
gwAuthCode=pcv_strBuyerID
gwTransID=pcv_strOrderNumber
paymentCode= "Google"

Todaydate=Date()
if SQL_Format="1" then
	Todaydate=Day(Todaydate)&"/"&Month(Todaydate)&"/"&Year(Todaydate)
else
	Todaydate=Month(Todaydate)&"/"&Day(Todaydate)&"/"&Year(Todaydate)
end if
pOrderTime=Now()

if scDB="Access" then
	query="UPDATE orders SET pcOrd_PaymentStatus=0,orderstatus=2, processDate=#"&Todaydate&"#,gwAuthCode='"&gwAuthCode&"',gwTransID='"&gwTransID&"',paymentCode='"&paymentCode&"',pcOrd_Payer='"& session("idCustomer") &"', pcOrd_Time=#"&pOrderTime&"# WHERE idOrder=" & pIdOrder
else
	query="UPDATE orders SET pcOrd_PaymentStatus=0,orderstatus=2, processDate='"&Todaydate&"',gwAuthCode='"&gwAuthCode&"',gwTransID='"&gwTransID&"',paymentCode='"&paymentCode&"',pcOrd_Payer='"& session("idCustomer") &"', pcOrd_Time='"&pOrderTime&"' WHERE idOrder=" & pIdOrder
end if

set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
set rs=nothing
'***********************************************************************************
' END: ORDER STATUS
'***********************************************************************************




'***********************************************************************************
' START: SAVE PRODUCTS ORDERED
'***********************************************************************************
for f=1 to ppcCartIndex 
 	if pcCartArray(f,10)=0 then     
    
		if pcCartArray(f,11)="" or isNull(pcCartArray(f,11)) then
			pcCartArray(f,11)="NULL"
		end if
	  
		if pcCartArray(f,12)="" or isNull(pcCartArray(f,12)) then
			pcCartArray(f,12)="NULL"
		end if
	  
		if pcCartArray(f,14)="" or isNull(pcCartArray(f,14)) then
			pcCartArray(f,14)=0
		end if
	 
		' replace , by .
		pcCartArray(f,14)=replace(pcCartArray(f,14),",",".")
		if pcCartArray(f,16)<>"" or pcCartArray(f,15)<>"0" then
			tempVar1=(pcCartArray(f,5) + pcCartArray(f,17))
		else
			tempVar1=(pcCartArray(f,5) + pcCartArray(f,3))
		end if
		
		If pcCartArray(f,16)="" then
			pcCartArray(f,16)=0
		end If
		
		if pcCartArray(f,15)<>"" then
			QDiscounts=pcCartArray(f,15)
		else
			QDiscounts="0"
		end if
		if pcCartArray(f,30)<>"" then
			ItemsDiscounts=pcCartArray(f,30)
		else
			ItemsDiscounts="0"
		end if
		
		'GGG Add-on start
		if pcCartArray(f,33)<>"" then
		geID=pcCartArray(f,33)
		else
		geID="0"
		end if
		
		if pcCartArray(f,34)<>"" then
		pGWOpt=pcCartArray(f,34)
		else
		pGWOpt="0"
		end if
		
		if pcCartArray(f,35)<>"" then
			pGWOptText=Server.HTMLEncode(pcCartArray(f,35))
			pGWOptText=replace(pGWOptText,"'","''")
			if len(pGWOptText)>240 then
				pGWOptText=left(pGWOptText,240)
			end if
		else
			pGWOptText=""
		end if
		
		if pGWOpt<>"0" then
			query="select pcGW_OptPrice from pcGWOptions where pcGW_IDOpt=" & pGWOpt
			set rsG=connTemp.execute(query)
			pGWPrice=rsG("pcGW_OptPrice")
			if pGWPrice<>"" then
			else
				pGWPrice="0"
			end if
		else
			pGWPrice="0"
		end if
		'GGG Add-on end
		
		pcv_xdetails=pcCartArray(f,21)
		if pcv_xdetails<>"" then
			pcv_xdetails=replace(pcv_xdetails,"<br>","|")
			pcv_xdetails=replace(pcv_xdetails,"'","''")
			pcv_xdetails=replace(pcv_xdetails,"''''","''")
		end if
		
		'// Start SDBA
		query="SELECT serviceSpec,stock,nostock,pcProd_BackOrder,pcDropShipper_ID FROM Products WHERE idproduct=" & pcCartArray(f,0)
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)		
		if not rs.eof then
			pcv_serviceSpec=rs("serviceSpec")
			if IsNull(pcv_pserviceSpec) or pcv_pserviceSpec="" then
				pcv_pserviceSpec="0"
			end if
			pcv_Stock=rs("stock")
			if IsNull(pcv_Stock) or pcv_Stock="" then
				pcv_Stock="0"
			end if
			pcv_NoStock=rs("nostock")
			if IsNull(pcv_NoStock) or pcv_NoStock="" then
				pcv_NoStock="0"
			end if
			pcv_intBackOrder=rs("pcProd_BackOrder")
			if IsNull(pcv_intBackOrder) or pcv_intBackOrder="" then
				pcv_intBackOrder="0"
			end if
			pcv_IDDropShipper=rs("pcDropShipper_ID")
			if IsNull(pcv_IDDropShipper) or pcv_IDDropShipper="" then
				pcv_IDDropShipper="0"
			end if
		else
			pcv_pserviceSpec="0"
			pcv_Stock="0"
			pcv_NoStock="0"
			pcv_intBackOrder="0"
			pcv_IDDropShipper="0"
		end if
		set rs=nothing				
		If (scOutofStockPurchase=-1 AND CLng(pcv_Stock)<1 AND pcv_serviceSpec=0 AND pcv_NoStock=0 AND pcv_intBackOrder=1) OR (pcv_serviceSpec<>0 AND scOutofStockPurchase=-1 AND iBTOOutofstockpurchase=-1 AND CLng(pcv_Stock)<1 AND pcv_NoStock=0 AND pcv_intBackOrder=1) Then
			tmp_BackOrder="1"
		Else
			tmp_BackOrder="0"
		End if
		'// End SDBA
		
		'// Insert the Line Item
		query="INSERT INTO ProductsOrdered (idOrder, idProduct, quantity, unitPrice, unitCost, idconfigSession, xfdetails, QDiscounts,ItemsDiscounts, pcPackageInfo_ID, pcDropShipper_ID, pcPrdOrd_Shipped, pcPrdOrd_BackOrder, pcPrdOrd_SelectedOptions, pcPrdOrd_OptionsPriceArray, pcPrdOrd_OptionsArray, pcPO_EPID,pcPO_GWOpt, pcPO_GWNote, pcPO_GWPrice) VALUES (" &pIdOrder& "," &pcCartArray(f,0)& "," &pcCartArray(f,2)& "," & replacecomma(tempVar1) & "," & replacecomma(pcCartArray(f,14))& "," &pcCartArray(f,16)& ",'" &pcv_xdetails& "'," & QDiscounts & "," & ItemsDiscounts & ",0," & pcv_IDDropShipper & ",0," & tmp_BackOrder & ",'" & replace(pcCartArray(f,11),"'","''") & "','" & replace(pcCartArray(f,25),"'","''") & "','" & replace(pcCartArray(f,4),"'","''") &"'," & geID & "," & pGWOpt & ",'" & pGWOptText & "'," & pGWPrice & ")"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)  
		set rs=nothing
	end if 
next 
'***********************************************************************************
' END: SAVE PRODUCTS ORDERED
'***********************************************************************************


'///////////////////////////////////////////////////////////////////////////////////
'//  END: ORDER PROCESSING
'///////////////////////////////////////////////////////////////////////////////////



'// SET THESE
pIdOrder=session("idOrderSaved")


'***********************************************************************************
' START: GET ORDER DETAILS
'***********************************************************************************
query="SELECT orders.idcustomer, orders.address, orders.City, orders.StateCode, orders.State, orders.zip, orders.CountryCode, orders.shippingAddress, orders.shippingCity, orders.shippingStateCode, orders.shippingState, orders.shippingZip,  orders.shippingCountryCode, orders.pcOrd_shippingPhone, orders.ShipmentDetails, orders.PaymentDetails, orders.discountDetails, orders.taxAmount, orders.total, orders.comments, orders.ShippingFullName, orders.address2, orders.ShippingCompany, orders.ShippingAddress2, orders.taxDetails, orders.orderstatus, orders.iRewardPoints, orders.iRewardValue, orders.iRewardRefId, orders.iRewardPointsRef, orders.iRewardPointsCustAccrued, orders.ordPackageNum, customers.phone, orders.ord_DeliveryDate, orders.ord_VAT, orders.pcOrd_DiscountsUsed, orders.pcOrd_Payer FROM orders, customers WHERE orders.idcustomer=customers.idcustomer AND orders.idOrder=" & pIdOrder
set rsObjOrder=server.CreateObject("ADODB.RecordSet")
set rsObjOrder=conntemp.execute(query)

pidcustomer=rsObjOrder("idcustomer")
paddress=rsObjOrder("address")
pCity=rsObjOrder("City")
pStateCode=rsObjOrder("StateCode")
pState=rsObjOrder("State")
if isNULL(pStateCode) OR pStateCode="" then
	pStateCode=pState
end if
pzip=rsObjOrder("zip")
pCountryCode=rsObjOrder("CountryCode")
pshippingAddress=rsObjOrder("shippingAddress")
pshippingCity=rsObjOrder("shippingCity")
pshippingStateCode=rsObjOrder("shippingStateCode")
pshippingState=rsObjOrder("shippingState")
if isNULL(pshippingStateCode) OR pshippingStateCode="" then
	pshippingStateCode=pshippingState
end if
pshippingZip=rsObjOrder("shippingZip")
pshippingCountryCode=rsObjOrder("shippingCountryCode")
pshippingPhone=rsObjOrder("pcOrd_shippingPhone")
pShipmentDetails=rsObjOrder("ShipmentDetails")
pPaymentDetails=rsObjOrder("PaymentDetails")
pdiscountDetails=rsObjOrder("discountDetails")
ptaxAmount=rsObjOrder("taxAmount")
ptotal=rsObjOrder("total")
pcomments=rsObjOrder("comments")
pShippingFullName=rsObjOrder("ShippingFullName")
paddress2=rsObjOrder("address2")
pShippingCompany=rsObjOrder("ShippingCompany")
pShippingAddress2=rsObjOrder("ShippingAddress2")
ptaxDetails=rsObjOrder("taxDetails")
pCurOrderStatus=rsObjOrder("orderStatus")
piRewardPoints=rsObjOrder("iRewardPoints")
piRewardValue=rsObjOrder("iRewardValue")
piRewardRefId=rsObjOrder("iRewardRefId")
piRewardPointsRef=rsObjOrder("iRewardPointsRef") 
piRewardPointsCustAccrued=rsObjOrder("iRewardPointsCustAccrued")
pOrdPackageNum=rsObjOrder("ordPackageNum")
pPhone=rsObjOrder("phone")
pord_DeliveryDate=rsObjOrder("ord_DeliveryDate")
pord_DeliveryDate=showDateFrmt(pord_DeliveryDate)
pord_VAT=rsObjOrder("ord_VAT")
strPcOrd_DiscountsUsed=rsObjOrder("pcOrd_DiscountsUsed")
pcOrd_Payer=rsObjOrder("pcOrd_Payer")
set rsObjOrder=nothing
'***********************************************************************************
' END: GET ORDER DETAILS
'***********************************************************************************

ppStatus=0 '// This will allow the code below to execute.


'***********************************************************************************
' START: CUSTOMER ID
'***********************************************************************************
pName= pFirstName
pLName= pLastName 
pIdCustomer = pcv_intCustomerId
'***********************************************************************************
' END: CUSTOMER ID
'***********************************************************************************




'***********************************************************************************
' START: ITERATE THROUGH ORDER ITEMS
'***********************************************************************************
query="SELECT idProduct,quantity,idconfigSession FROM ProductsOrdered WHERE idOrder=" & pIdOrder
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)  
	
do while not rs.eof  
	pIdProduct=rs("idProduct")
	pQuantity=rs("quantity")
	idconfig=rs("idconfigSession")
	
	'// Check if stock is ignored, or not
	query="SELECT noStock FROM products WHERE idProduct=" & pIdProduct
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=conntemp.execute(query)   
	pNoStock=rstemp("noStock")

	query="SELECT stock, sales, description FROM products WHERE idProduct=" & pIdProduct
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=conntemp.execute(query)       
	pDescription=rstemp("description")
	
	if pNoStock=0 then
		'// Decrease stock 
		if ppStatus=0 then
			query="UPDATE products SET stock=stock-" & pQuantity &" WHERE idProduct=" & pIdProduct
			set rstemp=server.CreateObject("ADODB.RecordSet")
			set rsTemp=conntemp.execute(query) 
			
			'// Update BTO Items & Additional Charges stock and sales 
			IF (idconfig<>"") and (idconfig<>"0") then
				query="select stringProducts,stringQuantity,stringCProducts from configSessions where idconfigSession=" & idconfig
				set rs1=server.CreateObject("ADODB.RecordSet")
				set rs1=conntemp.execute(query)
				stringProducts=rs1("stringProducts")
				stringQuantity=rs1("stringQuantity")
				stringCProducts=rs1("stringCProducts")
				if (stringProducts<>"") and (stringProducts<>"na") then
					PrdArr=split(stringProducts,",")
					QtyArr=split(stringQuantity,",")
					
					for k=lbound(PrdArr) to ubound(PrdArr)
						if (PrdArr(k)<>"") and (PrdArr(k)<>"0") then
							query="UPDATE products SET stock=stock-" &QtyArr(k)*pQuantity&",sales=sales+" &QtyArr(k)*pQuantity&" WHERE idProduct=" &PrdArr(k)
							set rs1=conntemp.execute(query)
						end if
					next
				end if
				if (stringCProducts<>"") and (stringCProducts<>"na") then
					CPrdArr=split(stringCProducts,",")
					
					for k=lbound(CPrdArr) to ubound(CPrdArr)
						if (CPrdArr(k)<>"") and (CPrdArr(k)<>"0") then
							query="UPDATE products SET stock=stock-" &pQuantity&",sales=sales+" &pQuantity&" WHERE idProduct=" &CPrdArr(k)
							set rs1=conntemp.execute(query)
						end if
					next
				end if
			END IF
			'// End Update BTO Items & Additional Charges stock and sales 
			
		end if
	end if
				 
	'// Update sales 
	if ppStatus=0 then  
		query="UPDATE products SET sales=sales+" &pQuantity&" WHERE idProduct=" &pIdProduct
		set rstemp=server.CreateObject("ADODB.RecordSet")
		set rstemp=conntemp.execute(query)  
		set rstemp=nothing 
	end if 
	
	
	rs.movenext	
loop
set rs=nothing
set rstemp=nothing
set rs1=nothing
'***********************************************************************************
' END: ITERATE THROUGH ORDER ITEMS
'***********************************************************************************




'***********************************************************************************
' START: REWARD POINTS
'***********************************************************************************
qry_ID=pIdOrder
If piRewardPoints > 0 Then
	if ppStatus=0 then
		'// Even if pending, if a customer uses pts, they must be held as substracted until order is canceled.
		query="SELECT iRewardPointsUsed, idCustomer FROM customers WHERE idCustomer=" & pIdCustomer
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		iRewardPointsUsed=rs("iRewardPointsUsed")
		If IsNull(iRewardPointsUsed) OR iRewardPointsUsed="" Then
			iRewardPointsUsed=0
		end if		
		query = "SELECT iRewardValue FROM orders WHERE idOrder=" & qry_ID
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		piRewardValue=rs("iRewardValue")
		iNewUsed = iRewardPointsUsed + piRewardPoints		
		query = "UPDATE customers SET iRewardPointsUsed=" & iNewUsed & " WHERE idCustomer=" & pIdCustomer
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=nothing
	end if
End If
'***********************************************************************************
' END: REWARD POINTS
'***********************************************************************************





'///////////////////////////////////////////////////////////////////////////////////
'// START: EMAILS
'///////////////////////////////////////////////////////////////////////////////////


'***********************************************************************************
' START: ONE TIME DISCOUNTS
'***********************************************************************************
If strPcOrd_DiscountsUsed<>"" then
	if instr(strPcOrd_DiscountsUsed,",") then
		pDiscountUsedArray=split(strPcOrd_DiscountsUsed,",")
		tempCnt=Ubound(strPcOrd_DiscountsUsed)
	else
		tempCnt=0
	end if
	for i=0 to tempCnt
		if tempCnt=0 then
			pDiscountUsedVar=strPcOrd_DiscountsUsed
		else
			pDiscountUsedVar=pDiscountUsedArray(i)
		end if
		query="INSERT INTO used_discounts (idDiscount, idcustomer) VALUES ("&pDiscountUsedVar&","&pIdCustomer&");"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		set rs=nothing
	Next
end if
'***********************************************************************************
' END: ONE TIME DISCOUNTS
'***********************************************************************************


'***********************************************************************************
' START: SDBA - Send Low Inventory Notification
'***********************************************************************************
%>
<!--#include file="inc_StockNotificationEmail.asp"-->
<%
'***********************************************************************************
' END: SDBA - Send Low Inventory Notification
'***********************************************************************************


'///////////////////////////////////////////////////////////////////////////////////
'// END: EMAILS
'///////////////////////////////////////////////////////////////////////////////////





'***********************************************************************************
' START: CLEAR DATA
'***********************************************************************************
Session("pcCartIndex")=Cint(0)
session("iOrderTotal")=""
session("continueRef")=""
session("pcSFCartRewards")=Cint(0)
session("pcSFUseRewards")=Cint(0)
session("IDRefer")=""
session("specialdiscount")=""
session("EPN_idOrder")=""
session("pc_pidOrder")=""
session("GWAuthCode")=""
session("GWTransId")=""
session("GWPaymentId")=""
session("GWTransType")=""
session("GWOrderId")=""
session("GWSessionID")=""
session("GWOrderDone")=""
'GGG Add-on start
session("Cust_BuyGift")=""
session("Cust_IDEvent")=""
'GGG Add-on end
'***********************************************************************************
' END: CLEAR DATA
'***********************************************************************************

'///////////////////////////////////////////////////////////////////////////////////
'//  END: ORDER STATUS AND PAYMENT STATUS
'///////////////////////////////////////////////////////////////////////////////////
%>
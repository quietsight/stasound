<%

'///// START - EXISTING NAMES /////

'***** START - Encoding Variables *****
IsEncodingIssues=1
'***** END - Encoding Variables *****

'***** START - Search Variables *****
srcFromDate_ex=0
srcToDate_ex=0
srcHideExported_ex=0
cm_products_ex=0
cm_product_ex=0
cm_customer_ex=0
cm_partnerID_ex=0
cm_partnerPassword_ex=0
cm_partnerKey_ex=0
cm_callbackURL_ex=0
cm_methodName_ex=0
cm_filters_ex=0
cm_requests_ex=0
cm_resultCount_ex=0
cm_requestKey_ex=0

'***** START - Search Products Variables *****
srcCategoryID_ex=0
srcCFieldID_ex=0
srcCFieldValue_ex=0
srcPriceFrom_ex=0
srcPriceTo_ex=0
srcInStock_ex=0
srcSKU_ex=0
srcBrandID_ex=0
srcKeyword_ex=0
srcExactPhrase_ex=0
srcIncInactive_ex=0
srcIncDeleted_ex=0
srcIncNormal_ex=0
srcIncBTO_ex=0
srcIncBTOItems_ex=0
srcSpecial_ex=0
srcFeatured_ex=0
srcSort_ex=0
'***** END - Search Products Variables *****

'***** START - Search Customers Variables *****
srcFirstName_ex=0
srcLastName_ex=0
srcCompany_ex=0
srcEmail_ex=0
srcCity_ex=0
srcCountryCode_ex=0
srcPhone_ex=0
srcCustomerType_ex=0
srcPricingCategory_ex=0
srcCustomerField_ex=0
srcIncLocked_ex=0
srcIncSuspended_ex=0
'***** END - Search Customers Variables *****

'***** START - Search Orders Variables *****
srcCustomerID_ex=0
srcPricingCatID_ex=0
srcOrderStatus_ex=0
srcPaymentStatus_ex=0
srcPaymentType_ex=0
srcShippingType_ex=0
srcStateCode_ex=0
srcDiscountCode_ex=0
srcPrdOrderedID_ex=0
'***** END - Search Orders Variables *****

'***** END - Search Variables *****

'***** START - Product Infor Variables *****
prdID_ex=0
prdSKU_ex=0
prdName_ex=0
prdDesc_ex=0
prdSDesc_ex=0
prdType_ex=0
prdPrice_ex=0
prdListPrice_ex=0
prdWPrice_ex=0
prdWeight_ex=0
Pounds_ex=0
Ounces_ex=0
Kgs_ex=0
Grams_ex=0
UnitsToPound_ex=0
prdStock_ex=0
prdCategory_ex=0
catID_ex=0
catName_ex=0
catLDesc_ex=0
catSDesc_ex=0
catImg_ex=0
catLargeImg_ex=0
catParentID_ex=0
catParentName_ex=0
prdBrand_ex=0
brandID_ex=0
brandName_ex=0
brandLogo_ex=0
prdSmallImg_ex=0
prdImg_ex=0
prdLargeImg_ex=0
prdActive_ex=0
prdShowSavings_ex=0
prdSpecial_ex=0
prdFeatured_ex=0
prdOptionGroup_ex=0
groupName_ex=0
groupRequired_ex=0
groupOrder_ex=0
option_ex=0
optName_ex=0
optPrice_ex=0
optWPrice_ex=0
optOrder_ex=0
optInactive_ex=0
prdRewardPoints_ex=0
prdNoTax_ex=0
prdNoShippingCharge_ex=0
prdNotForSale_ex=0
prdNotForSaleMsg_ex=0
prdDisregardStock_ex=0
prdDisplayNoShipText_ex=0
prdMinimumQty_ex=0
prdPurchaseMulti_ex=0
prdOversize_ex=0
osWidth_ex=0
osHeight_ex=0
osLength_ex=0
osX_ex=0
osWeight_ex=0
prdCost_ex=0
prdBackOrder_ex=0
prdShipNDays_ex=0
prdLowStockNotice_ex=0
prdReorderLevel_ex=0
prdIsDropShipped_ex=0
prdSupplierID_ex=0
prdDropShipperID_ex=0
prdMetaTags_ex=0
mtTitle_ex=0
mtDesc_ex=0
mtKeywords_ex=0
prdDownloadable_ex=0
prdDownloadInfo_ex=0
diFileLocation_ex=0
diURLExpire_ex=0
diURLExpDays_ex=0
diUseLicenseGen_ex=0
diLocalGen_ex=0
diRemoteGen_ex=0
diLFLabel1_ex=0
diLFLabel2_ex=0
diLFLabel3_ex=0
diLFLabel4_ex=0
diLFLabel5_ex=0
diAddMsg_ex=0
prdGiftCertificate_ex=0
prdGCInfo_ex=0
giExpire_ex=0
giEOnly_ex=0
giUseGen_ex=0
giExpDate_ex=0
giExpNDays_ex=0
giCustomGen_ex=0
prdHideBTOPrices_ex=0
prdHideDefaultConfig_ex=0
prdDisallowPurchasing_ex=0
prdSkipPrdPage_ex=0
prdCustomField_ex=0
cfID_ex=0
cfName_ex=0
cfValue_ex=0
prdCreatedDate_ex=0
'***** END - Product Infor Variables *****

'***** START - Customer Infor Variables *****
custID_ex=0
custFirstName_ex=0
custLastName_ex=0
custName_ex=0
custEmail_ex=0
custPassword_ex=0
custType_ex=0
custCompany_ex=0
custPhone_ex=0
custFax_ex=0
custBillingAddress_ex=0
custShippingAddress_ex=0
custAddress_ex=0
custAddress2_ex=0
custCity_ex=0
custStateCode_ex=0
custProvince_ex=0
custZip_ex=0
custCountryCode_ex=0
custShipCompany_ex=0
custShipAddress_ex=0
custShipAddress2_ex=0
custShipCity_ex=0
custShipStateCode_ex=0
custShipProvince_ex=0
custShipZip_ex=0
custShipCountryCode_ex=0
custShipAddressID_ex=0
custShipNickName_ex=0
custShipFirstName_ex=0
custShipLastName_ex=0
custShipEmail_ex=0
custShipPhone_ex=0
custShipFax_ex=0
custRPBalance_ex=0
custRPUsed_ex=0
custPricingCategory_ex=0
pricingCatID_ex=0
pricingCatName_ex=0
custField_ex=0
fieldID_ex=0
fieldName_ex=0
fieldValue_ex=0
custNewsletter_ex=0
custTotalOrders_ex=0
custTotalSales_ex=0
custCreatedDate_ex=0
custStatus_ex=0
'***** END - Customer Infor Variables *****

'***** START - Order Infor Variables *****
ordID_ex=0
ordName_ex=0
ordDate_ex=0
ordCustDetails_ex=0
ordTotal_ex=0
ordProcessedDate_ex=0
ordShippingAddress_ex=0
ordShipName_ex=0
ordShipCompany_ex=0
ordShipAddress_ex=0
ordShipAddress2_ex=0
ordShipCity_ex=0
ordShipStateCode_ex=0
ordShipProvince_ex=0
ordShipZip_ex=0
ordShipCountryCode_ex=0
ordShipPhone_ex=0
ordShipFax_ex=0
ordShipEmail_ex=0
ordShipDetails_ex=0
shipType_ex=0
shipFees_ex=0
handlingFees_ex=0
packageCount_ex=0
shipDate_ex=0
shipVia_ex=0
trackingNumber_ex=0
packageInfo_ex=0
pkgID_ex=0
pkgShipMethod_ex=0
pkgShipDate_ex=0
pkgComment_ex=0
pkgTrackingNumber_ex=0
ordDeliveryDate_ex=0
ordStatus_ex=0
ordPaymentStatus_ex=0
ordPayDetails_ex=0
paymentMethod_ex=0
paymentFees_ex=0
authorizationCode_ex=0
transactionID_ex=0
paymentGateway_ex=0
ordAffiliate_ex=0
affID_ex=0
affName_ex=0
affPayment_ex=0
ordRP_ex=0
ordAccruedRP_ex=0
ordReferrer_ex=0
refID_ex=0
refName_ex=0
ordRMACredit_ex=0
ordTaxAmount_ex=0
ordTaxDetails_ex=0
taxName_ex=0
taxAmount_ex=0
ordVAT_ex=0
ordDiscountDetails_ex=0
discountName_ex=0
discountAmount_ex=0
GiftCertificate_ex=0
GiftCertificateUsed_ex=0
ordCatDiscounts_ex=0
ordCustomerComments_ex=0
ordAdminComments_ex=0
ordReturnDate_ex=0
ordReturnReason_ex=0
ordShoppingCart_ex=0
products_ex=0
prdUnitPrice_ex=0
prdQuantity_ex=0
prdBTOConfig_ex=0
itemID_ex=0
itemName_ex=0
itemSKU_ex=0
itemCategoryID_ex=0
itemCategoryName_ex=0
itemQuantity_ex=0
itemAddPrice_ex=0
prdOption_ex=0
prdQtyDiscounts_ex=0
prdItemDiscounts_ex=0
prdGiftWrapping_ex=0
gwID_ex=0
gwName_ex=0
gwPrice_ex=0
packageID_ex=0
prdTotalPrice_ex=0
'***** END - Order Infor Variables *****

'***** START - Common Import Variables *****
ImportField_name_ex=0
'***** END - Common Import Variables *****

'///// END - EXISTING NAMES /////




'///// START - VALUES /////

BackupStr=""
xmlHaveErrors=0

'***** START - Search Variables *****
srcFromDate_value=""
srcToDate_value=""
srcHideExported_value=""
cm_partnerID_value=""
cm_partnerPassword_value=""
cm_partnerKey_value=""
cm_callbackURL_value=""
cm_methodName_value=""
cm_resultCount_value=""
cm_requestKey_value=""

'***** START - Search Products Variables *****
srcCategoryID_value=""
srcCFieldID_value=""
srcCFieldValue_value=""
srcPriceFrom_value=""
srcPriceTo_value=""
srcInStock_value=""
srcSKU_value=""
srcBrandID_value=""
srcKeyword_value=""
srcExactPhrase_value=""
srcIncInactive_value=""
srcIncDeleted_value=""
srcIncNormal_value=""
srcIncBTO_value=""
srcIncBTOItems_value=""
srcSpecial_value=""
srcFeatured_value=""
srcSort_value=""
'***** END - Search Products Variables *****

'***** START - Search Customers Variables *****
srcFirstName_value=""
srcLastName_value=""
srcCompany_value=""
srcEmail_value=""
srcCity_value=""
srcCountryCode_value=""
srcPhone_value=""
srcCustomerType_value=""
srcPricingCategory_value=""
srcCustomerField_value=""
srcIncLocked_value=""
srcIncSuspended_value=""
'***** END - Search Customers Variables *****

'***** START - Search Orders Variables *****
srcCustomerID_value=""
srcPricingCatID_value=""
srcOrderStatus_value=""
srcPaymentStatus_value=""
srcPaymentType_value=""
srcShippingType_value=""
srcStateCode_value=""
srcDiscountCode_value=""
srcPrdOrderedID_value=""
'***** END - Search Orders Variables *****

'***** END - Search Variables *****

'***** START - Product Infor Variables *****
prdID_value=""
prdSKU_value=""
prdName_value=""
prdDesc_value=""
prdSDesc_value=""
prdType_value=""
prdPrice_value=""
prdListPrice_value=""
prdWPrice_value=""
prdWeight_value=""
Pounds_value=""
Ounces_value=""
Kgs_value=""
Grams_value=""
UnitsToPound_value=""
prdStock_value=""
prdCategory_value=""
catID_value=""
catName_value=""
catLDesc_value=""
catSDesc_value=""
catImg_value=""
catLargeImg_value=""
catParentID_value=""
catParentName_value=""
prdBrand_value=""
brandID_value=""
brandName_value=""
brandLogo_value=""
prdSmallImg_value=""
prdImg_value=""
prdLargeImg_value=""
prdActive_value=""
prdShowSavings_value=""
prdSpecial_value=""
prdFeatured_value=""
prdOptionGroup_value=""
groupName_value=""
groupRequired_value=""
groupOrder_value=""
option_value=""
optName_value=""
optPrice_value=""
optWPrice_value=""
optOrder_value=""
optInactive_value=""
prdRewardPoints_value=""
prdNoTax_value=""
prdNoShippingCharge_value=""
prdNotForSale_value=""
prdNotForSaleMsg_value=""
prdDisregardStock_value=""
prdDisplayNoShipText_value=""
prdMinimumQty_value=""
prdPurchaseMulti_value=""
prdOversize_value=""
osWidth_value=""
osHeight_value=""
osLength_value=""
osX_value=""
osWeight_value=""
prdCost_value=""
prdBackOrder_value=""
prdShipNDays_value=""
prdLowStockNotice_value=""
prdReorderLevel_value=""
prdIsDropShipped_value=""
prdSupplierID_value=""
prdDropShipperID_value=""

prdIsDropShipper_value=0

prdMetaTags_value=""
mtTitle_value=""
mtDesc_value=""
mtKeywords_value=""
prdDownloadable_value=""
prdDownloadInfo_value=""
diFileLocation_value=""
diURLExpire_value=""
diURLExpDays_value=""
diUseLicenseGen_value=""
diLocalGen_value=""
diRemoteGen_value=""
diLFLabel1_value=""
diLFLabel2_value=""
diLFLabel3_value=""
diLFLabel4_value=""
diLFLabel5_value=""
diAddMsg_value=""
prdGiftCertificate_value=""
prdGCInfo_value=""
giExpire_value=""
giEOnly_value=""
giUseGen_value=""
giExpDate_value=""
giExpNDays_value=""
giCustomGen_value=""
prdHideBTOPrices_value=""
prdHideDefaultConfig_value=""
prdDisallowPurchasing_value=""
prdSkipPrdPage_value=""
prdCustomField_value=""
cfID_value=""
cfName_value=""
cfValue_value=""
prdCreatedDate_value=""
'***** END - Product Infor Variables *****

'***** START - Customer Infor Variables *****
custID_value=""
custFirstName_value=""
custLastName_value=""
custName_value=""
custEmail_value=""
custPassword_value=""
custType_value=""
custCompany_value=""
custPhone_value=""
custFax_value=""
custBillingAddress_value=""
custShippingAddress_value=""
custAddress_value=""
custAddress2_value=""
custCity_value=""
custStateCode_value=""
custProvince_value=""
custZip_value=""
custCountryCode_value=""
custShipCompany_value=""
custShipAddress_value=""
custShipAddress2_value=""
custShipCity_value=""
custShipStateCode_value=""
custShipProvince_value=""
custShipZip_value=""
custShipCountryCode_value=""
custShipAddressID_value=""
custShipNickName_value=""
custShipFirstName_value=""
custShipLastName_value=""
custShipEmail_value=""
custShipPhone_value=""
custShipFax_value=""
custRPBalance_value=""
custRPUsed_value=""
custPricingCategory_value=""
pricingCatID_value=""
pricingCatName_value=""
custField_value=""
fieldID_value=""
fieldName_value=""
fieldValue_value=""
custNewsletter_value=""
custTotalOrders_value=""
custTotalSales_value=""
custCreatedDate_value=""
custStatus_value=""
'***** END - Customer Infor Variables *****

'***** START - Order Infor Variables *****
ordID_value=""
ordName_value=""
ordDate_value=""
ordCustDetails_value=""
ordTotal_value=""
ordProcessedDate_value=""
ordShippingAddress_value=""
ordShipName_value=""
ordShipCompany_value=""
ordShipAddress_value=""
ordShipAddress2_value=""
ordShipCity_value=""
ordShipStateCode_value=""
ordShipProvince_value=""
ordShipZip_value=""
ordShipCountryCode_value=""
ordShipPhone_value=""
ordShipFax_value=""
ordShipEmail_value=""
ordShipDetails_value=""
shipType_value=""
shipFees_value=""
handlingFees_value=""
packageCount_value=""
shipDate_value=""
shipVia_value=""
trackingNumber_value=""
packageInfo_value=""
pkgID_value=""
pkgShipMethod_value=""
pkgShipDate_value=""
pkgComment_value=""
pkgTrackingNumber_value=""
ordDeliveryDate_value=""
ordStatus_value=""
ordPaymentStatus_value=""
ordPayDetails_value=""
paymentMethod_value=""
paymentFees_value=""
authorizationCode_value=""
transactionID_value=""
paymentGateway_value=""
ordAffiliate_value=""
affID_value=""
affName_value=""
affPayment_value=""
ordRP_value=""
ordAccruedRP_value=""
ordReferrer_value=""
refID_value=""
refName_value=""
ordRMACredit_value=""
ordTaxAmount_value=""
ordTaxDetails_value=""
taxName_value=""
taxAmount_value=""
ordVAT_value=""
ordDiscountDetails_value=""
discountName_value=""
discountAmount_value=""
GiftCertificate_value=""
GiftCertificateUsed_value=""
ordCatDiscounts_value=""
ordCustomerComments_value=""
ordAdminComments_value=""
ordReturnDate_value=""
ordReturnReason_value=""
ordShoppingCart_value=""
products_value=""
prdUnitPrice_value=""
prdQuantity_value=""
prdBTOConfig_value=""
itemID_value=""
itemName_value=""
itemSKU_value=""
itemCategoryID_value=""
itemCategoryName_value=""
itemQuantity_value=""
itemAddPrice_value=""
prdOption_value=""
ordOptID_value=""
prdQtyDiscounts_value=""
prdItemDiscounts_value=""
prdGiftWrapping_value=""
gwID_value=""
gwName_value=""
gwPrice_value=""
packageID_value=""
prdTotalPrice_value=""
'***** END - Order Infor Variables *****

'***** START - Common Import Variables *****
ImportField_value=""
'***** END - Common Import Variables *****

'///// END - VALUES /////

%>
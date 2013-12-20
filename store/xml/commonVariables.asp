<%
'///// START - NAMES /////

'***** START - Search Variables *****"
private const srcFromDate_name="FromDate"
private const srcToDate_name="ToDate"
private const srcHideExported_name="HideExported"
private const cm_products="Products"
private const cm_product="Product"
private const cm_customers="Customers"
private const cm_customer="Customer"
private const cm_orders="Orders"
private const cm_order="Order"

private const cm_ProductCartResponse_name="ProductCartResponse"
private const cm_ExportedFlag_name="ExportedFlag"

private const cm_SearchProductsRequest_name="SearchProductsRequest"
private const cm_SearchProductsResponse_name="SearchProductsResponse"
private const cm_SearchCustomersRequest_name="SearchCustomersRequest"
private const cm_SearchCustomersResponse_name="SearchCustomersResponse"
private const cm_SearchOrdersRequest_name="SearchOrdersRequest"
private const cm_SearchOrdersResponse_name="SearchOrdersResponse"

private const cm_GetProductDetailsRequest_name="GetProductDetailsRequest"
private const cm_GetProductDetailsResponse_name="GetProductDetailsResponse"
private const cm_GetCustomerDetailsRequest_name="GetCustomerDetailsRequest"
private const cm_GetCustomerDetailsResponse_name="GetCustomerDetailsResponse"
private const cm_GetOrderDetailsRequest_name="GetOrderDetailsRequest"
private const cm_GetOrderDetailsResponse_name="GetOrderDetailsResponse"

private const cm_NewProductsRequest_name="NewProductsRequest"
private const cm_NewProductsResponse_name="NewProductsResponse"
private const cm_NewCustomersRequest_name="NewCustomersRequest"
private const cm_NewCustomersResponse_name="NewCustomersResponse"
private const cm_NewOrdersRequest_name="NewOrdersRequest"
private const cm_NewOrdersResponse_name="NewOrdersResponse"

private const cm_AddProductRequest_name="AddProductRequest"
private const cm_AddProductResponse_name="AddProductResponse"
private const cm_AddCustomerRequest_name="AddCustomerRequest"
private const cm_AddCustomerResponse_name="AddCustomerResponse"

private const cm_UpdateProductRequest_name="UpdateProductRequest"
private const cm_UpdateProductResponse_name="UpdateProductResponse"
private const cm_UpdateCustomerRequest_name="UpdateCustomerRequest"
private const cm_UpdateCustomerResponse_name="UpdateCustomerResponse"

private const cm_UndoRequest_name="UndoRequest"
private const cm_UndoResponse_name="UndoResponse"

private const cm_MarkAsExportedRequest_name="SetExportFlagRequest"
private const cm_MarkAsExportedResponse_name="SetExportFlagResponse"

private const cm_partnerID_name="PartnerID"
private const cm_partnerPassword_name="PartnerPassword"
private const cm_partnerKey_name="PartnerKey"
private const cm_callbackURL_name="CallbackURL"

private const cm_filters_name="Filters"
private const cm_requests_name="Requests"
private const cm_request_name="Request"
private const cm_requestAll_name="All"
private const cm_requestDefault_name="Default"
private const cm_errorList_name="ErrorList"
private const cm_errorCode_name="ErrorCode"
private const cm_errorDesc_name="ErrorDesc"
private const cm_resultCount_name="Count"
private const cm_requestKey_name="RequestKey"
private const cm_requestStatus_name="RequestStatus"

'***** START - Search Products Variables *****"
private const srcCategoryID_name="CategoryID"
private const srcCFieldID_name="CustomFieldID"
private const srcCFieldValue_name="CustomFieldValue"
private const srcPriceFrom_name="PriceFrom"
private const srcPriceTo_name="PriceTo"
private const srcInStock_name="InStock"
private const srcSKU_name="SKU"
private const srcBrandID_name="BrandID"
private const srcKeyword_name="Keyword"
private const srcExactPhrase_name="ExactPhrase"
private const srcIncInactive_name="Inactive"
private const srcIncDeleted_name="Deleted"
private const srcIncNormal_name="Standard"
private const srcIncBTO_name="BTO"
private const srcIncBTOItems_name="BTOItem"
private const srcSpecial_name="Special"
private const srcFeatured_name="Featured"
private const srcSort_name="SortOrder"
'***** END - Search Products Variables *****"

'***** START - Search Customers Variables *****"
private const srcFirstName_name="FirstName"
private const srcLastName_name="LastName"
private const srcCompany_name="Company"
private const srcEmail_name="Email"
private const srcCity_name="City"
private const srcCountryCode_name="CountryCode"
private const srcPhone_name="Phone"
private const srcCustomerType_name="CustomerType"
private const srcPricingCategory_name="PricingCategoryID"
private const srcCustomerField_name="CustomerField"
private const srcIncLocked_name="Locked"
private const srcIncSuspended_name="Suspended"
'***** END - Search Customers Variables *****"

'***** START - Search Orders Variables *****"
private const srcCustomerID_name="CustomerID"
private const srcPricingCatID_name="PricingCategoryID"
private const srcOrderStatus_name="OrderStatus"
private const srcPaymentStatus_name="PaymentStatus"
private const srcPaymentType_name="PaymentType"
private const srcStateCode_name="StateCode"
private const srcDiscountCode_name="DiscountCode"
private const srcPrdOrderedID_name="ProductOrderedID"

'//Note: Use the same "srcCustomerType_name" and "srcCountryCode_name" with search customers

'***** END - Search Orders Variables *****"

'***** END - Search Variables *****"

'***** START - Product Info Variables *****"
private const prdID_name="ProductID"
private const prdSKU_name="SKU"
private const prdName_name="Name"
private const prdDesc_name="Description"
private const prdSDesc_name="ShortDescription"
private const prdType_name="Type"
private const prdPrice_name="Price"
private const prdListPrice_name="ListPrice"
private const prdWPrice_name="WholesalePrice"
private const prdWeight_name="Weight"
private const Pounds_name="Pounds"
private const Ounces_name="Ounces"
private const Kgs_name="Kilograms"
private const Grams_name="Grams"
private const UnitsToPound_name="UnitsToPound"
private const prdStock_name="Stock"
private const prdCategory_name="Category"
private const catID_name="ID"
private const catName_name="Name"
private const catLDesc_name="Description"
private const catSDesc_name="ShortDescription"
private const catImg_name="Image"
private const catLargeImg_name="LargeImage"
private const catParentID_name="ParentID"
private const catParentName_name="ParentName"
private const prdBrand_name="Brand"
private const brandID_name="ID"
private const brandName_name="Name"
private const brandLogo_name="Logo"
private const prdSmallImg_name="SmallImage"
private const prdImg_name="Image"
private const prdLargeImg_name="LargeImage"
private const prdActive_name="IsActive"
private const prdShowSavings_name="ShowSavings"
private const prdSpecial_name="Special"
private const prdFeatured_name="Featured"
private const prdOptionGroup_name="OptionGroup"
private const groupName_name="Name"
private const groupRequired_name="Required"
private const groupOrder_name="SortOrder"
private const option_name="Option"
private const optName_name="Name"
private const optPrice_name="Price"
private const optWPrice_name="WholesalePrice"
private const optOrder_name="SortOrder"
private const optInactive_name="Inactive"
private const prdRewardPoints_name="RewardPoints"
private const prdNoTax_name="NoTax"
private const prdNoShippingCharge_name="NoShippingCharge"
private const prdNotForSale_name="NotForSale"
private const prdNotForSaleMsg_name="NotForSaleMessage"
private const prdDisregardStock_name="DisregardStock"
private const prdDisplayNoShipText_name="DisplayNoShippingText"
private const prdMinimumQty_name="MinimumQuantity"
private const prdPurchaseMulti_name="PurchaseMultiple"
private const prdOversize_name="Oversized"
private const osWidth_name="Width"
private const osHeight_name="Height"
private const osLength_name="Length"
private const prdCost_name="Cost"
private const prdBackOrder_name="AllowBackOrder"
private const prdShipNDays_name="BackOrderNDays"
private const prdLowStockNotice_name="LowStockNotice"
private const prdReorderLevel_name="ReorderLevel"
private const prdIsDropShipped_name="IsDropShipped"
private const prdSupplierID_name="SupplierID"
private const prdDropShipperID_name="DropShipperID"
private const prdMetaTags_name="MetaTags"
private const mtTitle_name="Title"
private const mtDesc_name="Description"
private const mtKeywords_name="Keywords"
private const prdDownloadable_name="Downloadable"
private const prdDownloadInfo_name="DownloadDetails"
private const diFileLocation_name="FileLocation"
private const diURLExpire_name="Expires"
private const diURLExpDays_name="ExpiresNDays"
private const diUseLicenseGen_name="DeliverLicense"
private const diLocalGen_name="LocalFile"
private const diRemoteGen_name="RemoteFile"
private const diLFLabel1_name="LicenseField1"
private const diLFLabel2_name="LicenseField2"
private const diLFLabel3_name="LicenseField3"
private const diLFLabel4_name="LicenseField4"
private const diLFLabel5_name="LicenseField5"
private const diAddMsg_name="Message"
private const prdGiftCertificate_name="GiftCertificate"
private const prdGCInfo_name="GiftCertificateDetails"
private const giExpire_name="Expires"
private const giEOnly_name="ElectronicOnly"
private const giUseGen_name="Generator"
private const giExpDate_name="ExpiresDate"
private const giExpNDays_name="ExpiresNDays"
private const giCustomGen_name="CustomFileLocation"
private const prdHideBTOPrices_name="HideBTOPrices"
private const prdHideDefaultConfig_name="HideDefaultConfiguration"
private const prdDisallowPurchasing_name="DisallowPurchasing"
private const prdSkipPrdPage_name="SkipProductPage"
private const prdCustomField_name="CustomField"
private const cfID_name="ID"
private const cfName_name="Name"
private const cfValue_name="Value"
private const prdCreatedDate_name="DateCreated"
'***** END - Product Info Variables *****"

'***** START - Customer Info Variables *****"
private const custID_name="CustomerID"
private const custFirstName_name="FirstName"
private const custLastName_name="LastName"
private const custName_name="Name"
private const custEmail_name="Email"
private const custPassword_name="Password"
private const custType_name="Type"
private const custCompany_name="Company"
private const custPhone_name="Phone"
private const custFax_name="Fax"
private const custBillingAddress_name="BillingAddress"
private const custShippingAddress_name="ShippingAddress"
private const custAddress_name="Address"
private const custAddress2_name="Address2"
private const custCity_name="City"
private const custStateCode_name="StateCode"
private const custProvince_name="Province"
private const custZip_name="PostalCode"
private const custCountryCode_name="CountryCode"
private const custShipCompany_name="Company"
private const custShipAddress_name="Address"
private const custShipAddress2_name="Address2"
private const custShipCity_name="City"
private const custShipStateCode_name="StateCode"
private const custShipProvince_name="Province"
private const custShipZip_name="PostalCode"
private const custShipCountryCode_name="CountryCode"
private const custShipAddressID_name="ID"
private const custShipNickName_name="NickName"
private const custShipFirstName_name="FirstName"
private const custShipLastName_name="LastName"
private const custShipEmail_name="Email"
private const custShipPhone_name="Phone"
private const custShipFax_name="Fax"
private const custRPBalance_name="RewardPoints"
private const custRPUsed_name="RewardPointsUsed"
private const custPricingCategory_name="CustomerCategoryPricing"
private const pricingCatID_name="ID"
private const pricingCatName_name="Name"
private const custField_name="CustomerField"
private const fieldID_name="ID"
private const fieldName_name="Name"
private const fieldValue_name="Value"
private const custNewsletter_name="Newsletter"
private const custTotalOrders_name="TotalOrders"
private const custTotalSales_name="TotalSales"
private const custCreatedDate_name="DateCreated"
private const custStatus_name="Status"
'***** END - Customer Info Variables *****"

'***** START - Order Info Variables *****"
private const ordID_name="OrderID"
private const ordName_name="Name"
private const ordDate_name="Date"
'// Note: Use the same "custID_name" with customers
private const ordCustDetails_name="BillingAddress"
'// Note: Use the same from "custName_name" to "custCountryCode_name" with customers
private const ordTotal_name="Total"
private const ordProcessedDate_name="ProcessedDate"
private const ordShippingAddress_name="ShippingAddress"
private const ordShipName_name="FullName"
private const ordShipCompany_name="Company"
private const ordShipAddress_name="Address"
private const ordShipAddress2_name="Address2"
private const ordShipCity_name="City"
private const ordShipStateCode_name="StateCode"
private const ordShipProvince_name="Province"
private const ordShipZip_name="PostalCode"
private const ordShipCountryCode_name="CountryCode"
private const ordShipPhone_name="Phone"
private const ordShipFax_name="Fax"
private const ordShipEmail_name="Email"
private const ordShipDetails_name="ShippingDetails"
private const shipMethod_name="Method"
private const shipType_name="Type"
private const shipFees_name="Fees"
private const handlingFees_name="Handling"
private const packageCount_name="PackageCount"
private const shipWeight_name="Weight"
'// Note: Use the same "Pounds_name","Ounces_name","Kgs_name","Grams_name" with products
private const packageInfo_name="PackageDetails"
private const pkgID_name="ID"
private const pkgShipMethod_name="Method"
private const pkgShipDate_name="Date"
private const pkgComment_name="Comment"
private const pkgTrackingNumber_name="TrackingNumber"
private const ordDeliveryDate_name="DeliveryDate"
private const ordStatus_name="FulfillmentStatus"
private const ordPaymentStatus_name="PaymentStatus"
private const ordPayDetails_name="PaymentDetails"
private const paymentMethod_name="Method"
private const paymentFees_name="Fees"
private const authorizationCode_name="AuthorizationCode"
private const transactionID_name="TransactionID"
private const paymentGateway_name="PaymentGateway"
private const ordAffiliate_name="Affiliate"
private const affID_name="ID"
private const affName_name="Name"
private const affPayment_name="Amount"
private const ordRP_name="RewardPoints"
private const ordAccruedRP_name="RewardPointsAccrued"
private const ordReferrer_name="Referrer"
private const refID_name="ID"
private const refName_name="Name"
private const ordRMACredit_name="RMACredit"
private const ordTaxAmount_name="TotalTax"
private const ordTaxDetails_name="TaxDetails"
private const taxName_name="Name"
private const taxAmount_name="Amount"
private const ordVAT_name="TotalVAT"
private const ordDiscountDetails_name="DiscountDetails"
private const ordDiscountDetail_name="DiscountDetail"
private const discountName_name="Name"
private const discountAmount_name="Amount"
private const giftCertificate_name="GiftCertificate"
private const giftCertificateUsed_name="GiftCertificateUsed"
private const ordCatDiscounts_name="CategoryDiscount"
private const ordCustomerComments_name="CustomerComments"
private const ordAdminComments_name="AdminComments"
private const ordReturnDate_name="ReturnDate"
private const ordReturnReason_name="ReturnReason"
private const ordShoppingCart_name="ShoppingCart"
private const products_name="ProductsOrdered"
'//Note: Use the same "prdID_name","prdSKU_name" and "prdName_name" with products
private const prdUnitPrice_name="UnitPrice"
private const prdQuantity_name="Quantity"
private const prdBTOConfig_name="BTOConfiguration"
private const btoItem_name="BTOItem"
private const itemID_name="ID"
private const itemName_name="Name"
private const itemSKU_name="SKU"
private const itemCategoryID_name="CategoryID"
private const itemCategoryName_name="CategoryName"
private const itemQuantity_name="Quantity"
private const itemAddPrice_name="AddPrice"
private const prdOption_name="Option"
private const ordOptID_name="ID"
'//Note: Use the same "optName_name" and "optPrice_name" with products
private const prdQtyDiscounts_name="QuantityDiscount"
private const prdItemDiscounts_name="ItemDiscount"
private const prdGiftWrapping_name="GiftWrapping"
private const gwID_name="ID"
private const gwName_name="Name"
private const gwPrice_name="Price"
private const packageID_name="PackageID"
private const prdTotalPrice_name="Total"
'***** END - Order Info Variables *****"

'***** START - Export Variables *****
private const ExportedFlag_name="ExportedFlag"
private const ExportedID_name="ExportedID"
private const ExportedIDType_name="IDType"
'***** END - Export Variables *****

'***** START - Common Import Variables *****
private const ImportField_name="MapToField"
'***** END - Common Import Variables *****

'///// END - NAMES /////



'///// START - XML /////

'***** START - Encoding Variables *****
private const xmldocument_encoding="text/xml; charset:ISO-8859-1;"
private const xmlInit_encoding="ISO-8859-1"
'***** END - Encoding Variables *****

'///// END - XML /////



'///// START - EXISTING NAMES /////

'***** START - Export Variables *****
Dim ExportedFlag_ex
Dim ExportedID_ex
Dim ExportedIDType_ex
'***** END - Export Variables *****

'***** START - Search Variables *****
Dim srcFromDate_ex
Dim srcToDate_ex
Dim srcHideExported_ex
Dim cm_products_ex
Dim cm_product_ex
Dim cm_customer_ex
Dim cm_partnerID_ex
Dim cm_partnerPassword_ex
Dim cm_partnerKey_ex
Dim cm_callbackURL_ex
Dim cm_methodName_ex
Dim cm_filters_ex
Dim cm_requests_ex
Dim cm_resultCount_ex
Dim cm_requestKey_ex

'***** START - Search Products Variables *****
Dim srcCategoryID_ex
Dim srcCFieldID_ex
Dim srcCFieldValue_ex
Dim srcPriceFrom_ex
Dim srcPriceTo_ex
Dim srcInStock_ex
Dim srcSKU_ex
Dim srcBrandID_ex
Dim srcKeyword_ex
Dim srcExactPhrase_ex
Dim srcIncInactive_ex
Dim srcIncDeleted_ex
Dim srcIncNormal_ex
Dim srcIncBTO_ex
Dim srcIncBTOItems_ex
Dim srcSpecial_ex
Dim	srcFeatured_ex
Dim srcSort_ex
'***** END - Search Products Variables *****

'***** START - Search Customers Variables *****
Dim srcFirstName_ex
Dim srcLastName_ex
Dim srcCompany_ex
Dim srcEmail_ex
Dim srcCity_ex
Dim srcCountryCode_ex
Dim srcPhone_ex
Dim srcCustomerType_ex
Dim srcPricingCategory_ex
Dim srcCustomerField_ex
Dim srcIncLocked_ex
Dim srcIncSuspended_ex
'***** END - Search Customers Variables *****

'***** START - Search Orders Variables *****
Dim srcCustomerID_ex
Dim srcPricingCatID_ex
Dim srcOrderStatus_ex
Dim srcPaymentStatus_ex
Dim srcPaymentType_ex
Dim srcShippingType_ex
Dim srcStateCode_ex
Dim srcDiscountCode_ex
Dim srcPrdOrderedID_ex
'***** END - Search Orders Variables *****

'***** END - Search Variables *****

'***** START - Product Info Variables *****
Dim prdID_ex
Dim prdSKU_ex
Dim prdName_ex
Dim prdDesc_ex
Dim prdSDesc_ex
Dim prdType_ex
Dim prdPrice_ex
Dim prdListPrice_ex
Dim prdWPrice_ex
Dim prdWeight_ex
Dim Pounds_ex
Dim Ounces_ex
Dim Kgs_ex
Dim Grams_ex
Dim UnitsToPound_ex
Dim prdStock_ex
Dim prdCategory_ex
Dim catID_ex
Dim catName_ex
Dim catLDesc_ex
Dim catSDesc_ex
Dim catImg_ex
Dim catLargeImg_ex
Dim catParentID_ex
Dim catParentName_ex
Dim prdBrand_ex
Dim brandID_ex
Dim brandName_ex
Dim brandLogo_ex
Dim prdSmallImg_ex
Dim prdImg_ex
Dim prdLargeImg_ex
Dim prdActive_ex
Dim prdShowSavings_ex
Dim prdSpecial_ex
Dim prdFeatured_ex
Dim prdOptionGroup_ex
Dim groupName_ex
Dim groupRequired_ex
Dim groupOrder_ex
Dim option_ex
Dim optName_ex
Dim optPrice_ex
Dim optWPrice_ex
Dim optOrder_ex
Dim optInactive_ex
Dim prdRewardPoints_ex
Dim prdNoTax_ex
Dim prdNoShippingCharge_ex
Dim prdNotForSale_ex
Dim prdNotForSaleMsg_ex
Dim prdDisregardStock_ex
Dim prdDisplayNoShipText_ex
Dim prdMinimumQty_ex
Dim prdPurchaseMulti_ex
Dim prdOversize_ex
Dim osWidth_ex
Dim osHeight_ex
Dim osLength_ex
Dim osX_ex
Dim osWeight_ex
Dim prdCost_ex
Dim prdBackOrder_ex
Dim prdShipNDays_ex
Dim prdLowStockNotice_ex
Dim prdReorderLevel_ex
Dim prdIsDropShipped_ex
Dim prdSupplierID_ex
Dim prdDropShipperID_ex
Dim prdMetaTags_ex
Dim mtTitle_ex
Dim mtDesc_ex
Dim mtKeywords_ex
Dim prdDownloadable_ex
Dim prdDownloadInfo_ex
Dim diFileLocation_ex
Dim diURLExpire_ex
Dim diURLExpDays_ex
Dim diUseLicenseGen_ex
Dim diLocalGen_ex
Dim diRemoteGen_ex
Dim diLFLabel1_ex
Dim diLFLabel2_ex
Dim diLFLabel3_ex
Dim diLFLabel4_ex
Dim diLFLabel5_ex
Dim diAddMsg_ex
Dim prdGiftCertificate_ex
Dim prdGCInfo_ex
Dim giExpire_ex
Dim giEOnly_ex
Dim giUseGen_ex
Dim giExpDate_ex
Dim giExpNDays_ex
Dim giCustomGen_ex
Dim prdHideBTOPrices_ex
Dim prdHideDefaultConfig_ex
Dim prdDisallowPurchasing_ex
Dim prdSkipPrdPage_ex
Dim prdCustomField_ex
Dim cfID_ex
Dim cfName_ex
Dim cfValue_ex
Dim prdCreatedDate_ex
'***** END - Product Info Variables *****

'***** START - Customer Info Variables *****
Dim custID_ex
Dim custFirstName_ex
Dim custLastName_ex
Dim custName_ex
Dim custEmail_ex
Dim custPassword_ex
Dim custType_ex
Dim custCompany_ex
Dim custPhone_ex
Dim custFax_ex
Dim custBillingAddress_ex
Dim custShippingAddress_ex
Dim custAddress_ex
Dim custAddress2_ex
Dim custCity_ex
Dim custStateCode_ex
Dim custProvince_ex
Dim custZip_ex
Dim custCountryCode_ex
Dim custShipCompany_ex
Dim custShipAddress_ex
Dim custShipAddress2_ex
Dim custShipCity_ex
Dim custShipStateCode_ex
Dim custShipProvince_ex
Dim custShipZip_ex
Dim custShipCountryCode_ex
Dim custShipAddressID_ex
Dim custShipNickName_ex
Dim custShipFirstName_ex
Dim custShipLastName_ex
Dim custShipEmail_ex
Dim custShipPhone_ex
Dim custShipFax_ex
Dim custRPBalance_ex
Dim custRPUsed_ex
Dim custPricingCategory_ex
Dim pricingCatID_ex
Dim pricingCatName_ex
Dim custField_ex
Dim fieldID_ex
Dim fieldName_ex
Dim fieldValue_ex
Dim custNewsletter_ex
Dim custTotalOrders_ex
Dim custTotalSales_ex
Dim custCreatedDate_ex
Dim custStatus_ex
'***** END - Customer Info Variables *****

'***** START - Order Info Variables *****
Dim ordID_ex
Dim ordName_ex
Dim ordDate_ex
Dim ordCustDetails_ex
Dim ordTotal_ex
Dim ordProcessedDate_ex
Dim ordShippingAddress_ex
Dim ordShipName_ex
Dim ordShipCompany_ex
Dim ordShipAddress_ex
Dim ordShipAddress2_ex
Dim ordShipCity_ex
Dim ordShipStateCode_ex
Dim ordShipProvince_ex
Dim ordShipZip_ex
Dim ordShipCountryCode_ex
Dim ordShipPhone_ex
Dim ordShipFax_ex
Dim ordShipEmail_ex
Dim ordShipDetails_ex
Dim shipType_ex
Dim shipFees_ex
Dim handlingFees_ex
Dim packageCount_ex
Dim shipDate_ex
Dim shipVia_ex
Dim trackingNumber_ex
Dim packageInfo_ex
Dim pkgID_ex
Dim pkgShipMethod_ex
Dim pkgShipDate_ex
Dim pkgComment_ex
Dim pkgTrackingNumber_ex
Dim ordDeliveryDate_ex
Dim ordStatus_ex
Dim ordPaymentStatus_ex
Dim ordPayDetails_ex
Dim paymentMethod_ex
Dim paymentFees_ex
Dim authorizationCode_ex
Dim transactionID_ex
Dim paymentGateway_ex
Dim ordAffiliate_ex
Dim affID_ex
Dim affName_ex
Dim affPayment_ex
Dim ordRP_ex
Dim ordAccruedRP_ex
Dim ordReferrer_ex
Dim refID_ex
Dim refName_ex
Dim ordRMACredit_ex
Dim ordTaxAmount_ex
Dim ordTaxDetails_ex
Dim taxName_ex
Dim taxAmount_ex
Dim ordVAT_ex
Dim ordDiscountDetails_ex
Dim discountName_ex
Dim discountAmount_ex
Dim GiftCertificate_ex
Dim GiftCertificateUsed_ex
Dim ordCatDiscounts_ex
Dim ordCustomerComments_ex
Dim ordAdminComments_ex
Dim ordReturnDate_ex
Dim ordReturnReason_ex
Dim ordShoppingCart_ex
Dim products_ex
Dim prdUnitPrice_ex
Dim prdQuantity_ex
Dim prdBTOConfig_ex
Dim itemID_ex
Dim itemName_ex
Dim itemSKU_ex
Dim itemCategoryID_ex
Dim itemCategoryName_ex
Dim itemQuantity_ex
Dim itemAddPrice_ex
Dim prdOption_ex
Dim ordOptID_ex
Dim prdQtyDiscounts_ex
Dim prdItemDiscounts_ex
Dim prdGiftWrapping_ex
Dim gwID_ex
Dim gwName_ex
Dim gwPrice_ex
Dim packageID_ex
Dim prdTotalPrice_ex
'***** END - Order Info Variables *****

'///// END - EXISTING NAMES /////




'///// START - VALUES /////

Dim BackupStr
Dim xmlHaveErrors

'***** START - Export Variables *****
Dim ExportedFlag_value
Dim ExportedID_value
Dim ExportedIDType_value
'***** END - Export Variables *****

'***** START - Search Variables *****
Dim srcFromDate_value
Dim srcToDate_value
Dim srcHideExported_value
Dim cm_partnerID_value
Dim cm_partnerPassword_value
Dim cm_partnerKey_value
Dim cm_callbackURL_value
Dim cm_methodName_value
Dim cm_resultCount_value
Dim cm_requestKey_value

'***** START - Search Products Variables *****
Dim srcCategoryID_value
Dim srcCFieldID_value
Dim srcCFieldValue_value
Dim srcPriceFrom_value
Dim srcPriceTo_value
Dim srcInStock_value
Dim srcSKU_value
Dim srcBrandID_value
Dim srcKeyword_value
Dim srcExactPhrase_value
Dim srcIncInactive_value
Dim srcIncDeleted_value
Dim srcIncNormal_value
Dim srcIncBTO_value
Dim srcIncBTOItems_value
Dim srcSpecial_value
Dim	srcFeatured_value
Dim srcSort_value
'***** END - Search Products Variables *****

'***** START - Search Customers Variables *****
Dim srcFirstName_value
Dim srcLastName_value
Dim srcCompany_value
Dim srcEmail_value
Dim srcCity_value
Dim srcCountryCode_value
Dim srcPhone_value
Dim srcCustomerType_value
Dim srcPricingCategory_value
Dim srcCustomerField_value
Dim srcIncLocked_value
Dim srcIncSuspended_value
'***** END - Search Customers Variables *****

'***** START - Search Orders Variables *****
Dim srcCustomerID_value
Dim srcPricingCatID_value
Dim srcOrderStatus_value
Dim srcPaymentStatus_value
Dim srcPaymentType_value
Dim srcShippingType_value
Dim srcStateCode_value
Dim srcDiscountCode_value
Dim srcPrdOrderedID_value
'***** END - Search Orders Variables *****

'***** END - Search Variables *****

'***** START - Product Info Variables *****
Dim prdID_value
Dim prdSKU_value
Dim prdName_value
Dim prdDesc_value
Dim prdSDesc_value
Dim prdType_value
Dim prdPrice_value
Dim prdListPrice_value
Dim prdWPrice_value
Dim prdWeight_value
Dim Pounds_value
Dim Ounces_value
Dim Kgs_value
Dim Grams_value
Dim UnitsToPound_value
Dim prdStock_value
Dim prdCategory_value
Dim catID_value
Dim catName_value
Dim catLDesc_value
Dim catSDesc_value
Dim catImg_value
Dim catLargeImg_value
Dim catParentID_value
Dim catParentName_value
Dim prdBrand_value
Dim brandID_value
Dim brandName_value
Dim brandLogo_value
Dim prdSmallImg_value
Dim prdImg_value
Dim prdLargeImg_value
Dim prdActive_value
Dim prdShowSavings_value
Dim prdSpecial_value
Dim prdFeatured_value
Dim prdOptionGroup_value
Dim groupName_value
Dim groupRequired_value
Dim groupOrder_value
Dim option_value
Dim optName_value
Dim optPrice_value
Dim optWPrice_value
Dim optOrder_value
Dim optInactive_value
Dim prdRewardPoints_value
Dim prdNoTax_value
Dim prdNoShippingCharge_value
Dim prdNotForSale_value
Dim prdNotForSaleMsg_value
Dim prdDisregardStock_value
Dim prdDisplayNoShipText_value
Dim prdMinimumQty_value
Dim prdPurchaseMulti_value
Dim prdOversize_value
Dim osWidth_value
Dim osHeight_value
Dim osLength_value
Dim osX_value
Dim osWeight_value
Dim prdCost_value
Dim prdBackOrder_value
Dim prdShipNDays_value
Dim prdLowStockNotice_value
Dim prdReorderLevel_value
Dim prdIsDropShipped_value
Dim prdSupplierID_value
Dim prdDropShipperID_value

Dim prdIsDropShipper_value

Dim prdMetaTags_value
Dim mtTitle_value
Dim mtDesc_value
Dim mtKeywords_value
Dim prdDownloadable_value
Dim prdDownloadInfo_value
Dim diFileLocation_value
Dim diURLExpire_value
Dim diURLExpDays_value
Dim diUseLicenseGen_value
Dim diLocalGen_value
Dim diRemoteGen_value
Dim diLFLabel1_value
Dim diLFLabel2_value
Dim diLFLabel3_value
Dim diLFLabel4_value
Dim diLFLabel5_value
Dim diAddMsg_value
Dim prdGiftCertificate_value
Dim prdGCInfo_value
Dim giExpire_value
Dim giEOnly_value
Dim giUseGen_value
Dim giExpDate_value
Dim giExpNDays_value
Dim giCustomGen_value
Dim prdHideBTOPrices_value
Dim prdHideDefaultConfig_value
Dim prdDisallowPurchasing_value
Dim prdSkipPrdPage_value
Dim prdCustomField_value
Dim cfID_value
Dim cfName_value
Dim cfValue_value
Dim prdCreatedDate_value
'***** END - Product Info Variables *****

'***** START - Customer Info Variables *****
Dim custID_value
Dim custFirstName_value
Dim custLastName_value
Dim custName_value
Dim custEmail_value
Dim custPassword_value
Dim custType_value
Dim custCompany_value
Dim custPhone_value
Dim custFax_value
Dim custBillingAddress_value
Dim custShippingAddress_value
Dim custAddress_value
Dim custAddress2_value
Dim custCity_value
Dim custStateCode_value
Dim custProvince_value
Dim custZip_value
Dim custCountryCode_value
Dim custShipCompany_value
Dim custShipAddress_value
Dim custShipAddress2_value
Dim custShipCity_value
Dim custShipStateCode_value
Dim custShipProvince_value
Dim custShipZip_value
Dim custShipCountryCode_value
Dim custShipAddressID_value
Dim custShipNickName_value
Dim custShipFirstName_value
Dim custShipLastName_value
Dim custShipEmail_value
Dim custShipPhone_value
Dim custShipFax_value
Dim custRPBalance_value
Dim custRPUsed_value
Dim custPricingCategory_value
Dim pricingCatID_value
Dim pricingCatName_value
Dim custField_value
Dim fieldID_value
Dim fieldName_value
Dim fieldValue_value
Dim custNewsletter_value
Dim custTotalOrders_value
Dim custTotalSales_value
Dim custCreatedDate_value
Dim custStatus_value
'***** END - Customer Info Variables *****

'***** START - Order Info Variables *****
Dim ordID_value
Dim ordName_value
Dim ordDate_value
Dim ordCustDetails_value
Dim ordTotal_value
Dim ordProcessedDate_value
Dim ordShippingAddress_value
Dim ordShipName_value
Dim ordShipCompany_value
Dim ordShipAddress_value
Dim ordShipAddress2_value
Dim ordShipCity_value
Dim ordShipStateCode_value
Dim ordShipProvince_value
Dim ordShipZip_value
Dim ordShipCountryCode_value
Dim ordShipPhone_value
Dim ordShipFax_value
Dim ordShipEmail_value
Dim ordShipDetails_value
Dim shipType_value
Dim shipFees_value
Dim handlingFees_value
Dim packageCount_value
Dim shipDate_value
Dim shipVia_value
Dim trackingNumber_value
Dim packageInfo_value
Dim pkgID_value
Dim pkgShipMethod_value
Dim pkgShipDate_value
Dim pkgComment_value
Dim pkgTrackingNumber_value
Dim ordDeliveryDate_value
Dim ordStatus_value
Dim ordPaymentStatus_value
Dim ordPayDetails_value
Dim paymentMethod_value
Dim paymentFees_value
Dim authorizationCode_value
Dim transactionID_value
Dim paymentGateway_value
Dim ordAffiliate_value
Dim affID_value
Dim affName_value
Dim affPayment_value
Dim ordRP_value
Dim ordAccruedRP_value
Dim ordReferrer_value
Dim refID_value
Dim refName_value
Dim ordRMACredit_value
Dim ordTaxAmount_value
Dim ordTaxDetails_value
Dim taxName_value
Dim taxAmount_value
Dim ordVAT_value
Dim ordDiscountDetails_value
Dim discountName_value
Dim discountAmount_value
Dim GiftCertificate_value
Dim GiftCertificateUsed_value
Dim ordCatDiscounts_value
Dim ordCustomerComments_value
Dim ordAdminComments_value
Dim ordReturnDate_value
Dim ordReturnReason_value
Dim ordShoppingCart_value
Dim products_value
Dim prdUnitPrice_value
Dim prdQuantity_value
Dim prdBTOConfig_value
Dim itemID_value
Dim itemName_value
Dim itemSKU_value
Dim itemCategoryID_value
Dim itemCategoryName_value
Dim itemQuantity_value
Dim itemAddPrice_value
Dim prdOption_value
Dim ordOptID_value
Dim prdQtyDiscounts_value
Dim prdItemDiscounts_value
Dim prdGiftWrapping_value
Dim gwID_value
Dim gwName_value
Dim gwPrice_value
Dim packageID_value
Dim prdTotalPrice_value
'***** END - Order Info Variables *****

Dim xmlRequestID_value
Dim xmlRequestKey_value
Dim xmlRequestType_value
Dim	xmlBackup_value
Dim xmlUndo_value


'***** START - Export Variables *****
Dim pcv_strSummary
Dim pcv_CountCompleted
Dim pcv_CountTotal
'***** END - Export Variables *****


'***** START - Encoding Variables *****
Dim IsEncodingIssues
'***** END - Encoding Variables *****

'///// END - VALUES /////

%>

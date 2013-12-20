<%@ LANGUAGE = VBScript.Encode %>
<%'ProductCart XML Gateway

Session.LCID = 1033

Dim iXML,oXML,iRoot,oRoot,iNode,oNode,eNode,statNode,tmpNode,fNode,rNode
Dim connTemp,rs,query
Dim pcv_PartnerID
Dim cm_LogTurnOn,cm_LogErrors,cm_CaptureRequest,cm_CaptureResponse,cm_EnforceHTTPs,cm_ExportAdmin
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="commonVariables.asp"-->
<!--#include file="commonValues.asp"-->
<!--#include file="errorlist.asp"-->
<!--#include file="commonFunctions.asp"-->
<!--#include file="productFunctions.asp"-->
<!--#include file="customerFunctions.asp"-->
<!--#include file="orderFunctions.asp"-->
<%

	call GetXMLSettings()

	Set iXML=Server.CreateObject("MSXML2.DOMDocument"&scXML)	
	
	call InitResponseDocument(cm_ProductCartResponse_name)

	iXML.async=false
	iXML.load(Request)
	
	call CheckHTTPHeaders()
	
	call CheckValidXMLDocument()
	
	Set iRoot=iXML.documentElement
	
	call CheckCommonRequiredXMLTags()
	
	Select Case cm_methodName_value
		Case cm_SearchProductsRequest_name:
			Call CheckSrcProductsTags()
		Case cm_SearchCustomersRequest_name:
			Call CheckSrcCustomersTags()
		Case cm_SearchOrdersRequest_name:
			Call CheckSrcOrdersTags()
		Case cm_GetProductDetailsRequest_name:
			Call CheckGetProductDetailsTags()
		Case cm_GetCustomerDetailsRequest_name:
			Call CheckGetCustomerDetailsTags()
		Case cm_GetOrderDetailsRequest_name:
			Call CheckGetOrderDetailsTags()		
		Case cm_NewProductsRequest_name:
			Call CheckNewProductsTags()
		Case cm_NewCustomersRequest_name:
			Call CheckNewCustomersTags()
		Case cm_NewOrdersRequest_name:
			Call CheckNewOrdersTags()
		Case cm_AddProductRequest_name:
			Call CheckAddUpdProduct(0)
		Case cm_AddCustomerRequest_name:
			Call CheckAddUpdCustomer(0)
		Case cm_UpdateProductRequest_name:
			Call CheckAddUpdProduct(1)
		Case cm_UpdateCustomerRequest_name:
			Call CheckAddUpdCustomer(1)
		Case cm_UndoRequest_name:
			Call CheckUndoRequestTags()
		Case cm_MarkAsExportedRequest_name:
			Call CheckMarkAsExportedRequestTags()
	End Select
	
	call CheckValidPartner()
	
	Select Case cm_methodName_value
		Case cm_SearchProductsRequest_name:
			Call RunSrcProducts()
		Case cm_SearchCustomersRequest_name:
			Call RunSrcCustomers()
		Case cm_SearchOrdersRequest_name:
			Call RunSrcOrders()
		Case cm_GetProductDetailsRequest_name:
			Call RunGetProductDetails()
		Case cm_GetCustomerDetailsRequest_name:
			Call RunGetCustomerDetails()
		Case cm_GetOrderDetailsRequest_name:
			Call RunGetOrderDetails()
		Case cm_NewProductsRequest_name:
			Call RunNewProducts()
		Case cm_NewCustomersRequest_name:
			Call RunNewCustomers()
		Case cm_NewOrdersRequest_name:
			Call RunNewOrders()
		Case cm_AddProductRequest_name:
			Call RunAddUpdProduct(0)
		Case cm_AddCustomerRequest_name:
			Call RunAddUpdCustomer(0)
		Case cm_UpdateProductRequest_name:
			Call RunAddUpdProduct(1)
		Case cm_UpdateCustomerRequest_name:
			Call RunAddUpdCustomer(1)
		Case cm_UndoRequest_name:
			Call RunUndoRequest()
		Case cm_MarkAsExportedRequest_name:
			Call RunMarkAsExportedRequest()
	End Select
	
	call returnXML()

%>
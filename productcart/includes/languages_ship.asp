<%
if session("language")="" then
   session("language")="english"
end if

dim ship_dictLanguage  
set ship_dictLanguage=CreateObject("Scripting.Dictionary")

' english
ship_dictLanguage.Add "english_shiplogin_a", "Please wait while we gather shipping information for your order"
ship_dictLanguage.Add "english_chooseShpmnt_a", "Service Type"
ship_dictLanguage.Add "english_chooseShpmnt_b", "Delivery Time"
ship_dictLanguage.Add "english_chooseShpmnt_c", "Rate"
ship_dictLanguage.Add "english_chooseShpmnt_d", "No shipping rates were returned for your order. The shipping provider's server may be unavailable at this time. Please try again in a few minutes. If the problem persists, please contact the store administrator."
ship_dictLanguage.Add "english_chooseShpmnt_e", "*Free shipping applies only to the services that are marked as &quot;Free&quot;."
ship_dictLanguage.Add "english_chooseShpmnt_f", "Free*"
ship_dictLanguage.Add "english_chooseShpmnt_g", "We were not able to calculate shipping rates for your order at this time. If you continue the checkout process, we will contact you after we receive your order to discuss shipping options with you. Shipping charges will be added to your order at that time."
ship_dictLanguage.Add "english_chooseShpmnt_h", "This order will be shipped in "
ship_dictLanguage.Add "english_chooseShpmnt_i", " package(s)"

ship_dictLanguage.Add "english_login_a", "Not Available"
ship_dictLanguage.Add "english_login_b", "Shipping Information"
ship_dictLanguage.Add "english_login_c", "Residence"
ship_dictLanguage.Add "english_login_d", "Business"
ship_dictLanguage.Add "english_login_e", "You have "
ship_dictLanguage.Add "english_login_f", ". That means "
ship_dictLanguage.Add "english_login_g", " to use towards this purchase!"
ship_dictLanguage.Add "english_login_h", "Will this order be shipped to a business or residence?"

ship_dictLanguage.Add "english_custRewards_a", "Your current "
ship_dictLanguage.Add "english_custRewards_b", " balance is "
ship_dictLanguage.Add "english_custRewards_c", "This translates into "
ship_dictLanguage.Add "english_custRewards_d", " that you can use for purchases in this store."
ship_dictLanguage.Add "english_custRewards_e", "You can use your "
ship_dictLanguage.Add "english_custRewards_f", " when placing an order by entering the desired amount on the checkout page."
ship_dictLanguage.Add "english_custRewards_g", "To date, you have accrued a total of "
ship_dictLanguage.Add "english_custRewards_h", "To date, you have used a total of "

ship_dictLanguage.Add "english_viewCart_a", "Total Weight: "
ship_dictLanguage.Add "english_viewCart_b", "Estimated Shipping Charges"
ship_dictLanguage.Add "english_viewCart_c", "Weight: "

ship_dictLanguage.Add "english_noShip_a", "No shipping charge (or no shipping required)." 'do not use more then one comma in this line.
ship_dictLanguage.Add "english_noShip_b", "Shipping charges to be determined." 'do not use more then one comma in this line.

ship_dictLanguage.Add "english_chooseShpmnt_j", "Other Shipping Options" 
ship_dictLanguage.Add "english_chooseShpmnt_k", "UPS" 
ship_dictLanguage.Add "english_chooseShpmnt_l", "USPS"
ship_dictLanguage.Add "english_chooseShpmnt_m", "FedEx"
ship_dictLanguage.Add "english_chooseShpmnt_n", "Canada Post"

' START new text strings in v3
	ship_dictLanguage.Add "english_chooseShpmnt_10", "Choose Shipping Provider"
	ship_dictLanguage.Add "english_chooseShpmnt_11", "Select a Shipping Option"
	
	'Start SDBA
	ship_dictLanguage.Add "english_dropshipping_msg", "Please note that this order might be shipped from multiple warehouses and therefore arrive in separate packages."
	
	ship_dictLanguage.Add "english_partship_sbj_1", "Order ID #<ORDER_ID> - Partially Shipped"
	ship_dictLanguage.Add "english_partship_sbj_1a", "Order ID #<ORDER_ID> - Drop-Shipper Comments"
	ship_dictLanguage.Add "english_partship_msg_1", "Please note that your order #<ORDER_ID> has only been partially shipped. You will receive another notification from us when the rest of your order is shipped."
	ship_dictLanguage.Add "english_partship_msg_1a", "Order #<ORDER_ID> has been partially shipped."
	ship_dictLanguage.Add "english_partship_msg_2", "List of shipped product(s):"
	ship_dictLanguage.Add "english_partship_msg_3", "Package Information:"
	ship_dictLanguage.Add "english_partship_msg_4", "Shipment Method:"
	ship_dictLanguage.Add "english_partship_msg_5", "Tracking Number:"
	ship_dictLanguage.Add "english_partship_msg_6", "Shipped Date:"
	ship_dictLanguage.Add "english_partship_msg_7", "Comments:"
	ship_dictLanguage.Add "english_partship_msg_8", "Please Note: this shipment completes your order. According to our records, all products have been shipped and should arrive shortly (delivery time changes depending on the shipping option you selected when you placed your order). If you do not receive one or more of the items that you have ordered, please contact us."
	ship_dictLanguage.Add "english_partship_msg_8a", "Please Note: this is the last package that has been shipped. You can click the link below to obtain additional information on Order #<ORDER_ID>:"
	ship_dictLanguage.Add "english_partship_sbj_9", "Order ID #<ORDER_ID> - Shipment Completed"

	
	ship_dictLanguage.Add "english_custconfirm_msg_1", "Do you want one or separate shipments? Please click on the link below to let us know."
	
	ship_dictLanguage.Add "english_admconfirm_msg_1", "BACK-ORDERED PRODUCTS NOTICE"
	ship_dictLanguage.Add "english_admconfirm_msg_2", "This order contains products that are back-ordered. The products are:"
	
	ship_dictLanguage.Add "english_notifyseparate_sbj_1", "Order ID #<ORDER_ID> - Shipping Preference Notification"
	ship_dictLanguage.Add "english_notifyseparate_msg_1", "The customer indicated they he/she wants to receive separate shipments. Thefore, you should ship the products that are immediately available. You can click the link below to manage the Order #<ORDER_ID>:"
	ship_dictLanguage.Add "english_notifyseparate_msg_2", "The customer indicated they he/she wants to wait until all products are in stock and receive only one shipment. You can click the link below to manage the Order #<ORDER_ID>:"
	
	ship_dictLanguage.Add "english_sds_notifyorder_1", "Products List:"
	ship_dictLanguage.Add "english_sds_notifyorder_2", "Ship products directly to the customer address:"
	ship_dictLanguage.Add "english_sds_notifyorder_2a", "Ship products to our store address:"
	ship_dictLanguage.Add "english_sds_notifyorder_3", "Shipping Method:"
	
	ship_dictLanguage.Add "english_sds_notifycanceldorder_1", "The order #<ORDER_ID> has been cancelled - <STORE_NAME>"
	ship_dictLanguage.Add "english_sds_notifycanceldorder_2", "Dear <DROP_SHIPPER_COMPANY> <DROP_SHIPPER_NAME>,<br><br>The order #<ORDER_ID> has been cancelled. Please do not ship any products in this order.<br>"
			
	'End SDBA
	
' END new text strings in v3

' START new text strings in v3.1
	ship_dictLanguage.Add "english_xmlPrdInfo_kg", "kg"
	ship_dictLanguage.Add "english_xmlPrdInfo_g", "g"
	ship_dictLanguage.Add "english_xmlPrdInfo_lbs", "lbs"
	ship_dictLanguage.Add "english_xmlPrdInfo_ozs", "ozs"
' END new text strings in v3.1

' START new text strings in v3.5
	ship_dictLanguage.Add "english_partship_msg_9", "Tracking Link: "
	ship_dictLanguage.Add "english_orderVerify_msg_1", "You cannot checkout because the order amount has changed and free shipping is not available with the new order amount. Please click on the Back button above to recalculate shipping charges"
' END new text strings in v3.5

' end language definitions
function clearShipLanguage()
 ' clear the dictionary.
 on error resume next
 clearShipLanguage 		= ship_dictLanguage.removeAll   
 set clearShipLanguage 	= nothing
end function
%>


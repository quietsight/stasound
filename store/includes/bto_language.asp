<%
if session("language")="" then
   session("language")="english"
end if

dim bto_dictLanguage  
set bto_dictLanguage = CreateObject("Scripting.Dictionary")

' english
bto_dictLanguage.Add "english_configurePrd_1",  "Configure: "    
bto_dictLanguage.Add "english_configurePrd_2",  "Base Price: "    
bto_dictLanguage.Add "english_configurePrd_3",  "Customized Price: "    
bto_dictLanguage.Add "english_configurePrd_4",  "Default Price: "    
bto_dictLanguage.Add "english_configurePrd_5",  "Customizations: "    
bto_dictLanguage.Add "english_configurePrd_6",  "Price as Customized:"
bto_dictLanguage.Add "english_configurePrd_8",  "Items Discounts:"
bto_dictLanguage.Add "english_configurePrd_12",  "Total"
bto_dictLanguage.Add "english_configurePrd_13",  "Additional Charges"
bto_dictLanguage.Add "english_configurePrd_14",  "Additional Charges:"
bto_dictLanguage.Add "english_configurePrd_15",  "Quantity: "
bto_dictLanguage.Add "english_configurePrd_16",  "Customizations (<a href='Reconfigure.asp?pcCartIndex=" & request.QueryString("pcCartIndex") & "&N=1&idconf=" & request("idconf") & "&act=" & request("act") & "'>Edit</a>): "
bto_dictLanguage.Add "english_configurePrd_17",  "Customizations (<a href='Reconfigure.asp?pcCartIndex=" & request.QueryString("pcCartIndex") & "&N=0'>Edit</a>): "
bto_dictLanguage.Add "english_configurePrd_18",  "Customizations (<a href='bto_Reconfigure.asp?ido=" & request.QueryString("ido") & "&idp="&request.QueryString("idp")&"&configSession="&request.QueryString("configSession")&"'>Edit</a>): "
bto_dictLanguage.Add "english_configurePrd_18A",  "Customizations (<a href='bto_Reconfigure.asp?idproduct=" & request.QueryString("idproduct") & "&idconf="&request.QueryString("idconf")&"&idquote="&request.QueryString("idquote")&"'>Edit</a>): "
bto_dictLanguage.Add "english_configurePrd_19",  "There are no configurable options for this product at this time. Please contact the store administrator."
bto_dictLanguage.Add "english_configurePrd_20",  "None"
bto_dictLanguage.Add "english_configurePrd_21",  "Please note: you can select a maximum of "
bto_dictLanguage.Add "english_configurePrd_21a",  " items."

bto_dictLanguage.Add "english_instConfiguredPrd_1", "You are trying to add an item to the cart that has not been properly configured. You must have at least one option chosen to customize this product."

bto_dictLanguage.Add "english_viewcart_1", "Customizations: "
bto_dictLanguage.Add "english_viewcart_2", "No customizations."
bto_dictLanguage.Add "english_viewcart_3", "Additional Charges: "
bto_dictLanguage.Add "english_viewcart_4", "QTY: "

bto_dictLanguage.Add "english_CustviewPastD_1", "Customizations: "
bto_dictLanguage.Add "english_CustviewPastD_2", "Items Discounts:"
bto_dictLanguage.Add "english_CustviewPastD_3", "Quantity Discounts:"
bto_dictLanguage.Add "english_CustviewPastD_4", "Product Subtotal:"
bto_dictLanguage.Add "english_CustviewPastD_5", "Additional Charges:"

bto_dictLanguage.Add "english_Custquotesview_1", "Saved Quotes"
bto_dictLanguage.Add "english_Custquotesview_2", "SKU"
bto_dictLanguage.Add "english_Custquotesview_3", "Description"
bto_dictLanguage.Add "english_Custquotesview_4", "Quote"
bto_dictLanguage.Add "english_Custquotesview_5", "There are no items in your saved quotes list"
bto_dictLanguage.Add "english_Custquotesview_6", "This item is no longer available for purchase."
bto_dictLanguage.Add "english_Custquotesview_7", "View printer-friendly version"
bto_dictLanguage.Add "english_Custquotesview_8", "This quote was submitted to"
bto_dictLanguage.Add "english_Custquotesview_9", "on"
bto_dictLanguage.Add "english_Custquotesview_10", "The discount code or gift certificate that you have entered might have expired, or might have already been used, or cannot be applied to your quote (e.g. product quantity too low, total price too low). If you have any questions, please contact us."   
bto_dictLanguage.Add "english_Custquotesview_11", "Your quote was updated successfully!"
bto_dictLanguage.Add "english_Custquotesview_12", "Discount Code"
bto_dictLanguage.Add "english_Custquotesview_13", "Description"
bto_dictLanguage.Add "english_Custquotesview_14", "Details"
bto_dictLanguage.Add "english_Custquotesview_15", "(includes discount by code)"
bto_dictLanguage.Add "english_Custquotesview_16", "Amount"
bto_dictLanguage.Add "english_Custquotesview_17", "Error"
bto_dictLanguage.Add "english_Custquotesview_18", "Created On: "
bto_dictLanguage.Add "english_Custquotesview_19", "Submitted On: "

bto_dictLanguage.Add "english_printableQuote_6", "QTY"
bto_dictLanguage.Add "english_printableQuote_1", "SKU"
bto_dictLanguage.Add "english_printableQuote_2", "Description"
bto_dictLanguage.Add "english_printableQuote_3", "Price"
bto_dictLanguage.Add "english_printableQuote_4", "Customizations:"
bto_dictLanguage.Add "english_printableQuote_5", "Additional Charges:"

bto_dictLanguage.Add "english_custPref_1", "View Saved Quotes"

' START new in v3
bto_dictLanguage.Add "english_configurePrd_50", "Please enter your Discount or Gift Certificate Code (if any):"
bto_dictLanguage.Add "english_quotenotice_1", "Your quote has been finalized"
bto_dictLanguage.Add "english_quotenotice_2", "The quote #"
bto_dictLanguage.Add "english_quotenotice_3", " has been finalized. Please log-in to place an order."
' END new in v3

' Conflict Management - START
bto_dictLanguage.Add "english_btocm_1", "You cannot select "
bto_dictLanguage.Add "english_btocm_2", ", when you also select "
bto_dictLanguage.Add "english_btocm_3", "You cannot select any items from the category "
bto_dictLanguage.Add "english_btocm_4", "You must select "
bto_dictLanguage.Add "english_btocm_4a", ". So we changed the selection from "
bto_dictLanguage.Add "english_btocm_4b", ". So we selected "
bto_dictLanguage.Add "english_btocm_4c", " to "
bto_dictLanguage.Add "english_btocm_4d", " in the category "
bto_dictLanguage.Add "english_btocm_5", "You must deselect "
bto_dictLanguage.Add "english_btocm_5a", " before you can select "
bto_dictLanguage.Add "english_btocm_5b", " because they are not compatible."
bto_dictLanguage.Add "english_btocm_6", "Please select an item in the category "
bto_dictLanguage.Add "english_btocm_7", "You cannot select any items in the category "
bto_dictLanguage.Add "english_btocm_8a", " is not compatible with "
bto_dictLanguage.Add "english_btocm_8b", ", so we changed that selection to "
bto_dictLanguage.Add "english_btocm_8c", ". There might be other items in the "
bto_dictLanguage.Add "english_btocm_8d", " category that are compatible with "
bto_dictLanguage.Add "english_btocm_8e", ". Review that category to confirm or change the selection."
bto_dictLanguage.Add "english_btocm_8f", " in the "
bto_dictLanguage.Add "english_btocm_8g", " category, so we unchecked that selection."
bto_dictLanguage.Add "english_btocm_9", " from "
bto_dictLanguage.Add "english_btocm_10", "an item from the category "
bto_dictLanguage.Add "english_btocm_11", ". So we unselected all items in this category."
bto_dictLanguage.Add "english_btocm_11a", "any items from the category "
' Conflict Management - END

' end language definitions
function clearbtoLanguage()
 ' clear the dictionary.
 on error resume next
 clearbtoLanguage 		= bto_dictLanguage.removeAll   
 set clearbtoLanguage 	= nothing
end function
%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<% 
dim query, conntemp, rstemp
call openDb()

add2="images/sample/pc_button_add.gif"
addtocart="images/sample/pc_button_add.gif"
addtowl="images/sample/pc_button_wishlist.gif"
checkout="images/sample/pc_button_checkout.gif"
pcancel="images/sample/pc_button_cancel.gif"
continueshop="images/sample/pc_button_continue_shop.gif"
morebtn="images/sample/pc_button_details.gif"
login="images/sample/pc_button_login.gif"
submit="images/sample/pc_button_continue.gif"
recalculate="images/sample/pc_button_recalculate.gif"
register="images/sample/pc_button_register.gif"
remove="images/sample/pc_button_remove.gif"
login_checkout=""
register_checkout=""
back= "images/sample/pc_button_back.gif"
viewcartbtn= "images/sample/pc_button_viewcart.gif"
checkoutbtn= "images/sample/pc_button_tellafriend.gif"
'BTO ADDON-S
If scBTO=1 then
	customize= "images/sample/pc_button_customize.gif"
	reconfigure= "images/sample/pc_button_reconfig.gif"
	resetdefault= "images/sample/pc_button_reset.gif"
	savequote= "images/sample/pc_button_quote.gif"
	revorder= "images/sample/pc_button_review_order.gif"
	submitquote="images/sample/pc_button_submit_quote.gif"
	pcv_requestQuote="images/sample/pc_button_request_quote.gif"
End If
'BTO ADDON-E
pcv_placeOrder="images/sample/pc_button_placeOrder.gif"
pcv_checkoutWR="images/sample/pc_button_checkoutwr.gif"
pcv_processShip="images/sample/pc_button_shipment_process.gif"
pcv_finalShip="images/sample/pc_button_shipment_finalize.gif"
pcv_backtoOrder="images/sample/pc_button_back_to_order_details.gif"
pcv_previous="images/sample/pc_button_previous.gif"
pcv_next="images/sample/pc_button_next.gif"

'GGG Add-on start

	pcrereg="images/sample/ggg_button_create.gif"
	pdelreg="images/sample/ggg_button_delete.gif"
	paddreg="images/sample/ggg_button_add.gif"
	pupdreg="images/sample/ggg_button_update.gif"
	psendmsgs="images/sample/ggg_button_send.gif"
	pretreg="images/sample/ggg_button_return.gif"

'GGG Add-on end

yellowupd="images/sample/pc_button_update.gif"
pcv_strSaveCart="images/sample/pc_save_cart.gif"

query="UPDATE layout SET add2='"& add2 &"',addtocart='"& addtocart &"',addtowl='"& addtowl &"',checkout='"& checkout &"',cancel='"& pcancel &"',continueshop='"& continueshop &"',morebtn='"& morebtn &"',login='"& login &"',submit='"& submit &"',recalculate='"& recalculate &"',register='"& register &"',remove='"&  remove &"',login_checkout='.',back='"& back &"',register_checkout='.',viewcartbtn='"& viewcartbtn &"',checkoutbtn='"& checkoutbtn &"'"
'BTO ADDON-S
If scBTO=1 then
	query=query&",customize='"& customize &"',[reconfigure]='"& reconfigure &"',resetdefault='"& resetdefault &"',savequote='"& savequote &"',revorder='"& revorder&"',submitquote='"& submitquote & "',pcLO_requestQuote='" & pcv_requestQuote & "'"
End If
'BTO ADDON-E

'GGG Add-on start

query=query&",CreRegistry='"& pcrereg &"',DelRegistry='"& pdelreg &"',AddToRegistry='"& paddreg &"',UpdRegistry='"& pupdreg &"',SendMsgs='"& psendmsgs&"',RetRegistry='"& pretreg & "'"

'GGG Add-on end

query=query&",pcLO_placeOrder='" & pcv_placeOrder & "',pcLO_checkoutWR='" & pcv_checkoutWR & "',pcLO_processShip='" & pcv_processShip & "',pcLO_finalShip='" & pcv_finalShip & "',pcLO_backtoOrder='" & pcv_backtoOrder & "',pcLO_Previous='" & pcv_previous & "',pcLO_Next='" & pcv_next & "',pcLO_Update='" & yellowupd & "',pcLO_Savecart='" & pcv_strSaveCart & "' WHERE id=2;"

set rstemp=Server.CreateObject("ADODB.Recordset")     
set rstemp=conntemp.execute(query)

if err.number <> 0 then
	set rstemp=nothing
	call closeDb()
  	response.write "Error on setBtnDefault.asp: "&Err.Description
end If 

set rstemp=nothing
call closeDb()
response.redirect "AdminButtons.asp" 
%>
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
dim query, conntemp, rs
call openDb()

padd2=Request.QueryString("File1")
paddtocart=Request.QueryString("File2")
paddtowl=Request.QueryString("File3")
pcheckout=Request.QueryString("File4")
pcancel=Request.QueryString("File5")
pcontinueshop=Request.QueryString("File6")
pmorebtn=Request.QueryString("File7")
plogin=Request.QueryString("File8")
precalculate=Request.QueryString("File10")
pregister=Request.QueryString("File11")
premove=Request.QueryString("File12")
plogin_checkout=Request.QueryString("File13")
pback=Request.QueryString("File14")
pregister_checkout=Request.QueryString("File15")
pviewcartbtn=Request.QueryString("File16")
pcheckoutbtn=Request.QueryString("File17")
'BTO ADDON-S
if scBTO=1 then
	pcustomize=Request.QueryString("File18")
	preconfigure=Request.QueryString("File19")
	presetdefault=Request.QueryString("File20")
	psavequote=Request.QueryString("File21")
	prevorder=Request.QueryString("File22")
	psubmitquote=Request.QueryString("File23")
	pcv_requestQuote=Request.QueryString("File24")
end if
'BTO ADDON-E

pcv_placeOrder=Request.QueryString("File25")
pcv_checkoutWR=Request.QueryString("File26")
pcv_processShip=Request.QueryString("File27")
pcv_finalShip=Request.QueryString("File28")
pcv_backtoOrder=Request.QueryString("File29")
pcv_previous=Request.QueryString("File30")
pcv_next=Request.QueryString("File31")

'GGG Add-on start

	pcrereg=Request.QueryString("File32")
	pdelreg=Request.QueryString("File33")
	paddreg=Request.QueryString("File34")
	pupdreg=Request.QueryString("File35")
	psendmsgs=Request.QueryString("File36")
	pretreg=Request.QueryString("File37")

'GGG Add-on end

yellowupd=Request.QueryString("File38")
savecart=Request.QueryString("File39")
psubmit=yellowupd

cma="0"
query="UPDATE layout SET "

'placeOrder
If pcv_placeOrder <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	pcv_placeOrder="images/pc/"& pcv_placeOrder
	query=query &"pcLO_placeOrder='"& pcv_placeOrder &"'"
End If

'checkoutWR
If pcv_checkoutWR <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	pcv_checkoutWR="images/pc/"& pcv_checkoutWR
	query=query &"pcLO_checkoutWR='"& pcv_checkoutWR &"'"
End If

'processShip
If pcv_processShip <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	pcv_processShip="images/pc/"& pcv_processShip
	query=query &"pcLO_processShip='"& pcv_processShip &"'"
End If

'finalShip
If pcv_finalShip <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	pcv_finalShip="images/pc/"& pcv_finalShip
	query=query &"pcLO_finalShip='"& pcv_finalShip &"'"
End If

'backtoOrder
If pcv_backtoOrder <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	pcv_backtoOrder="images/pc/"& pcv_backtoOrder
	query=query &"pcLO_backtoOrder='"& pcv_backtoOrder &"'"
End If

'Previous
If pcv_previous <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	pcv_previous="images/pc/"& pcv_previous
	query=query &"pcLO_Previous='"& pcv_previous &"'"
End If

'Next
If pcv_next <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	pcv_next="images/pc/"& pcv_next
	query=query &"pcLO_Next='"& pcv_next &"'"
End If

'recalculate
If precalculate <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim precalculate2
	precalculate2="images/pc/"&precalculate
	query=query &"recalculate='"& precalculate2 &"'"
End If
'continueshop
If pcontinueshop <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim pcontinueshop2
	pcontinueshop2="images/pc/"&pcontinueshop
	query=query &"continueshop='"& pcontinueshop2 &"'"
End If
'checkout
If pcheckout <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim pcheckout2
	pcheckout2="images/pc/"&pcheckout
	query=query &"checkout='"& pcheckout2 &"'"
End If
'submit
If psubmit <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim psubmit2
	psubmit2="images/pc/"&psubmit
	query=query &"submit='"& psubmit2 &"'"
End If
'morebtn
If pmorebtn <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim pmorebtn2
	pmorebtn2="images/pc/"&pmorebtn
	query=query &"morebtn='"& pmorebtn2 &"'"
End If
'viewcartbtn
If pviewcartbtn <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim pviewcartbtn2
	pviewcartbtn2="images/pc/"&pviewcartbtn
	query=query &"viewcartbtn='"&  pviewcartbtn2 &"'"
End If
'checkoutbtn
If pcheckoutbtn <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim pcheckoutbtn2
	pcheckoutbtn2="images/pc/"&pcheckoutbtn
	query=query &"checkoutbtn='"&  pcheckoutbtn2 &"'"
End If
'addtocart
If paddtocart <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim paddtocart2
	paddtocart2="images/pc/"&paddtocart
	query=query &"addtocart='"& paddtocart2 &"'"
End If
'addtowl
If paddtowl <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim paddtowl2
	paddtowl2="images/pc/"&paddtowl
	query=query &"addtowl='"& paddtowl2 &"'"
End If
'register
If pregister <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim pregister2
	pregister2="images/pc/"&pregister
	query=query &"register='"& pregister2 &"'"
End If
'cancel
If pcancel <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim pcancel2
	pcancel2="images/pc/"&pcancel
	query=query &"cancel='"& pcancel2 &"'"
End If
'remove
If premove <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim premove2
	premove2="images/pc/"&premove
	query=query &"remove='"&  premove2 &"'"
End If
'login_checkout
If plogin_checkout <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim plogin_checkout2
	plogin_checkout2="images/pc/"&plogin_checkout
	query=query &"login_checkout='"& plogin_checkout2 &"'"
End If
'register_checkout
If pregister_checkout <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim pregister_checkout2
	pregister_checkout2="images/pc/"&pregister_checkout
	query=query &"register_checkout='"& pregister_checkout2 &"'"
End If
'BTO ADDON-S
if scBTO=1 then

	'requestQuote
	If pcv_requestQuote <> "" Then
		If cma="0" Then
		Else
			query=query & ","
		End If
		cma="1"
		pcv_requestQuote="images/pc/"& pcv_requestQuote
		query=query &"pcLO_requestQuote='"& pcv_requestQuote &"'"
	End If

	'resetdefault
	If presetdefault <> "" Then
		If cma="0" Then
		Else
			query=query & ","
		End If
		cma="1"
		Dim presetdefault2
		presetdefault2="images/pc/"&presetdefault
		query=query &"resetdefault='"& presetdefault2 &"'"
	End If
	'review & order
	If prevorder <> "" Then
		If cma="0" Then
		Else
			query=query & ","
		End If
		cma="1"
		Dim prevorder2
		prevorder2="images/pc/"&prevorder
		query=query &"revorder='"& prevorder2 &"'"
	End If
	'submit quote
	If psubmitquote <> "" Then
		If cma="0" Then
		Else
			query=query & ","
		End If
		cma="1"
		Dim psubmitquote2
		psubmitquote2="images/pc/"&psubmitquote
		query=query &"submitquote='"& psubmitquote2 &"'"
	End If
	'customize
	If pcustomize <> "" Then
		If cma="0" Then
		Else
			query=query & ","
		End If
		cma="1"
		Dim pcustomize2
		pcustomize2="images/pc/"&pcustomize
		query=query &"customize='"& pcustomize2 &"'"
	End If
	'reconfigure
	If preconfigure <> "" Then
		If cma="0" Then
		Else
			query=query & ","
		End If
		cma="1"
		Dim preconfigure2
		preconfigure2="images/pc/"&preconfigure
		query=query &"[reconfigure]='"& preconfigure2 &"'"
	End If
	'savequote
	If psavequote <> "" Then
		If cma="0" Then
		Else
			query=query & ","
		End If
		cma="1"
		Dim psavequote2
		psavequote2="images/pc/"&psavequote
		query=query &"savequote='"& psavequote2 &"'"
	End If
End If
'BTO ADDON-E
'add2
If padd2 <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	padd2="images/pc/"&padd2
	query=query &"add2='"& padd2 &"'"
End If
'login
If plogin <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim plogin2
	plogin2="images/pc/"&plogin
	query=query &"login='"& plogin2 &"'"
End If
'back
If pback <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim pback2
	pback2="images/pc/"&pback
	query=query &"back='"& pback2 &"'"
End If

'GGG Add-on start
'Create Registry
If pcrereg <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim pcrereg2
	pcrereg2="images/pc/"&pcrereg
	query=query &"CreRegistry='"& pcrereg2 &"'"
End If

'Delete Registry
If pdelreg <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim pdelreg2
	pdelreg2="images/pc/"&pdelreg
	query=query &"DelRegistry='"& pdelreg2 &"'"
End If

'Add to Registry
If paddreg <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim paddreg2
	paddreg2="images/pc/"&paddreg
	query=query &"AddToRegistry='"& paddreg2 &"'"
End If

'Update Registry
If pupdreg <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim pupdreg2
	pupdreg2="images/pc/"&pupdreg
	query=query &"UpdRegistry='"& pupdreg2 &"'"
End If

'Send Messages
If psendmsgs <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim psendmsgs2
	psendmsgs2="images/pc/"&psendmsgs
	query=query &"SendMsgs='"& psendmsgs2 &"'"
End If

'Return to Registry
If pretreg <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim pretreg2
	pretreg2="images/pc/"&pretreg
	query=query &"RetRegistry='"& pretreg2 &"'"
End If

'GGG Add-on end

'Yellow "Continue" Button
If yellowupd <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim yellowupd2
	yellowupd2="images/pc/"&yellowupd
	query=query &"pcLO_Update='"& yellowupd2 &"'"
End If

'Save Cart Button
If savecart <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	Dim savecart2
	savecart2="images/pc/"&savecart
	query=query &"pcLO_Savecart='"& savecart2 &"'"
End If

query=query &" WHERE id=2"

set rs=Server.CreateObject("ADODB.Recordset")
set rs=conntemp.execute(query)
if err.number <> 0 then
	set rs=nothing
	call closeDb()
	response.write "Error on dbButtons.asp: "&Err.Description
end If
set rs=nothing
call closeDb()
s=request.querystring("s")
msg=request.querystring("msg")
response.redirect "AdminButtons.asp?msg="&Server.URLEncode(msg)&"&s="&s
%>
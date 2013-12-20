<%@ LANGUAGE="VBSCRIPT" %>
<% 'OPTION EXPLICIT %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "Custwl.asp"
' This page checks if a product is already saved. If not it adds the product to the wishlist.
'
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="CustLIv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#INCLUDE FILE="../includes/stringfunctions.asp"-->  
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/currencyformatinc.asp" --> 
<!--#include file="../includes/rc4.asp" --> 
<!--#include FILE="../includes/ErrorHandler.asp"-->
<%
Response.Buffer = True

Dim conntemp, mysql, rstemp, query
Dim pcv_strSelectedOptions, pcvstrTmpOptionGroup, xOptionGroupCount, pcv_intOptionGroupCount

'*****************************************************************************************************
' START: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************
%>
<!--#include file="pcStartSession.asp"-->
<%
'*****************************************************************************************************
' END: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Page On-Load
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
call openDb()

session("redirectUrlLI")=request("redirectUrl") 

'// 1) ViewPrd.asp Page Only

pIdCustomer=session("idcustomer")
pIdProduct=getUserInput(request.querystring("idProduct"),0)
if not validNum(pIdProduct) then
   response.redirect "msg.asp?message=210" 
end if

'--> Get the options from the querysting
pcv_intOptionGroupCount = getUserInput(request.querystring("OptionGroupCount"),0)
if ((IsNull(pcv_intOptionGroupCount)) OR (pcv_intOptionGroupCount="") OR (not validNum(pcv_intOptionGroupCount))) then
	pcv_intOptionGroupCount = 0
end if
pcv_intOptionGroupCount = cint(pcv_intOptionGroupCount)

xOptionGroupCount = 0
pcv_strSelectedOptions = ""
do until xOptionGroupCount = pcv_intOptionGroupCount	
	xOptionGroupCount = xOptionGroupCount + 1
	pcvstrTmpOptionGroup = request.querystring("idOption"&xOptionGroupCount)
	if pcvstrTmpOptionGroup <> "" then			
		pcv_strSelectedOptions = pcv_strSelectedOptions & pcvstrTmpOptionGroup & chr(124)	
	end if	
loop
' trim the last pipe if there is one
xStringLength = len(pcv_strSelectedOptions)
if xStringLength>0 then
	pcv_strSelectedOptions = left(pcv_strSelectedOptions,(xStringLength-1))
end if

'// 2) Additional variables from BTO Configure page
pidConfigWishlistSession=request.querystring("idConf")
pdiscountcode=request("discountcode")
if pdiscountcode="" then
	pdiscountcode="0"
else
	pdiscountcode=replace(pdiscountcode,"'","''")
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Page On-Load
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: BTO ADDON
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If scBTO=1 then
	if pidConfigWishlistSession="" then
		pidConfigWishlistSession="0"
	end if
	if pidConfigWishlistSession<>"0" then		
		query="SELECT idproduct FROM configWishlistSessions WHERE idconfigWishlistSession=" &pidConfigWishlistSession
		set rstemp=server.CreateObject("ADODB.RecordSet")
		set rstemp=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rstemp=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if		
		pIdProduct=rstemp("idProduct")		
		set rstemp=nothing
	end if
end if
'BTO ADDON-E
if pidConfigWishlistSession="" then
	pidConfigWishlistSession="0"
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: BTO ADDON
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: check if that item w/ same options exists in wish list already.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if pidConfigWishlistSession="0" then '// If its not configurable then check if it exists
	query="SELECT idproduct, pcwishList_OptionsArray FROM wishlist WHERE idcustomer=" & pIdCustomer & " "
	query = query & "AND idProduct=" &pIdProduct& " "
	query = query & "AND idconfigWishlistSession=0; "
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=conntemp.execute(query)	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if	
	if NOT rstemp.eof then
		do until rstemp.eof
			pcv_wishList_OptionsArray=rstemp("pcwishList_OptionsArray")
			if isNULL(pcv_wishList_OptionsArray)=True then
				pcv_wishList_OptionsArray=""
			end if
			if pcv_wishList_OptionsArray = pcv_strSelectedOptions then
				set rstemp=nothing
				call closeDb()
				response.redirect "msg.asp?message=36"
			end if
			rstemp.movenext
		loop
	end if
	set rstemp=nothing	
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: check if that item w/ same options exists in wish list already.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: INSERT INTO wishlist
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if pidConfigWishlistSession<>"0" then '// if its standard
	query =			"INSERT INTO wishlist (idCustomer, idProduct, idconfigWishlistSession, discountcode) "
	query = query & "VALUES ("&pIdCustomer&","&pIdProduct&", "&pidConfigWishlistSession&",'" & pdiscountcode & "')"	
else '// if its BTO (no options)
	query =			"INSERT INTO wishlist (idCustomer, idProduct, idconfigWishlistSession, discountcode, pcwishList_OptionsArray) "
	query = query & "VALUES (" &pIdCustomer& "," &pIdProduct& ",0,'" & pdiscountcode & "','" & pcv_strSelectedOptions & "')"	
end if
set rstemp=server.CreateObject("ADODB.RecordSet")
set rstemp=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
set rstemp=nothing
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: INSERT INTO wishlist
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Get Service Spec and Redirect
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
query="SELECT products.serviceSpec FROM products WHERE idproduct=" & pIdProduct
set rstemp=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

pService=rs("serviceSpec")

set rstemp=nothing
call closeDb()

if (pidConfigWishlistSession<>"0") or (pService=-1) then
	response.redirect "Custquotesview.asp?msg="&Server.Urlencode(dictLanguage.Item(Session("language")&"_Custwl_3"))
else
	response.redirect "Custquotesview.asp?msg="&Server.Urlencode(dictLanguage.Item(Session("language")&"_Custwl_3"))
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Get Service Spec and Redirect
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>


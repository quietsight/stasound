<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add custom field to products" %>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<% 'on error resume next
Dim connTemp, rs, query
  
  pcIntBrandID=request("idbrand")
  if not validNum(pcIntBrandID) then
	msg=Server.URLEncode("Product assignment failed due to invalid brand ID.")
	response.Redirect("BrandsManage.asp?message="&msg)
  end if

	If (request("prdlist")<>"") and (request("prdlist")<>",") then
	prdlist=split(request("prdlist"),",")
		call openDb()
		For i=lbound(prdlist) to (ubound(prdlist)-1)
		id=prdlist(i)
			IF validNum(id) THEN
					query="UPDATE products SET IDbrand="& pcIntBrandID &"  WHERE idproduct="& id
					Set rstemp=conntemp.execute(query)
					Set rstemp=nothing
			ELSE
				msg=Server.URLEncode("Product assignment failed due to invalid product ID.")
				response.Redirect("BrandsManage.asp?message="&msg)
			END IF
		next
		call closeDb()
	End if 'have prdlist

session("admin_useExist")=0
session("admin_idcustom")=0
session("admin_skeyword")=""

msg=Server.URLEncode("The selected products were successfully assigned to this brand.")
response.Redirect("BrandsProducts.asp?idbrand="&pcIntBrandID&"&s=1&message="&msg)

%>
<!--#include file="AdminFooter.asp"-->
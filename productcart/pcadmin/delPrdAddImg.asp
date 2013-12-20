<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% 
pageTitle="Delete Additional Product Images"

dim query, conntemp, rstemp

' form parameters
pIdProduct=request.Querystring("idProduct")
pcProdImage_Url=request.Querystring("timg")
pcProdImage_LargeUrl=request.Querystring("dimg")
pcProdImage_ID=request.Querystring("pid")
redir=request.Querystring("redir")

if redir="" then
    redir = "modifyProduct.asp"
end If

if request.Querystring("iPageCurrent")="" then
    iPageCurrent=1 
else
    iPageCurrent=Request.Querystring("iPageCurrent")
end If

if not validNum(pIdProduct) then
  response.redirect "msg.asp?message=2"
end if

call openDb()

query="DELETE FROM pcProductsImages WHERE pcProdImage_ID=" &pcProdImage_ID
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(query)
if err.number <> 0 then
	pcErrDescription=err.description
	set rstemp=nothing
	call closeDb()
    response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&pcErrDescription) 
end If
set rstemp=nothing
call closeDb()
%>
<!--#include file="AdminHeader.asp"-->
<div class="pcCPmessageSuccess">The following images are no longer associated with this product:
	<br /><br /><%response.write pcProdImage_Url%>
	<br /><%response.write pcProdImage_LargeUrl%>
	<br /><br /><a href="<%=redir %>?idproduct=<%= pIdProduct %>&iPageCurrent=<%=iPageCurrent%>&tab=tab3#TabbedPanels1">Back</a>
</div>
<!--#include file="AdminFooter.asp"-->
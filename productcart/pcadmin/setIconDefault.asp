<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%PmAdmin=1%><!--#include file="adminv.asp"-->
<% 
dim query, conntemp, rstemp
call openDb()

erroricon="images/sample/pc_icon_error.gif"
requiredicon="images/sample/pc_icon_required.gif"
errorfieldicon="images/sample/pc_icon_errorfield.gif"
previousicon="images/sample/pc_icon_prev.gif"
nexticon="images/sample/pc_icon_next.gif"
discount="images/sample/pc_icon_discount.gif"
zoom="images/sample/pc_icon_zoom.gif"
arrowUp="images/sample/up-arrow.gif"
arrowDown="images/sample/down-arrow.gif"

query="UPDATE icons SET erroricon='"& erroricon &"',requiredicon='"& requiredicon &"',errorfieldicon='"& errorfieldicon &"',previousicon='"& previousicon &"',nexticon='"& nexticon &"',zoom='"& zoom &"',discount='"& discount &"',arrowUp='"& arrowUp &"',arrowDown='"& arrowDown &"' WHERE id=1"

set rstemp=Server.CreateObject("ADODB.Recordset")     
set rstemp=conntemp.execute(query)

if err.number <> 0 then
    response.write "Error: "&Err.Description
end If 
set rstemp=nothing
call closeDb()
response.redirect "AdminIcons.asp?s=1&msg=" & Server.URLEncode("Store icons were reset to the default image that ships with ProductCart. No images were deleted from the server.") 
%>
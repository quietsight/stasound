<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=8%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% pageTitle="Delete Affiliate from Database" %>
<% 
dim query, conntemp, rs

'// retreive form parameters
pcv_Idaffiliate=request.Querystring("idaffiliate")

if trim(pcv_Idaffiliate)="" then
   response.redirect "msg.asp?message=7"
end if

call openDb()

'// check if affiliate is associated with any orders, if so, make inactive and show message to admin
query="SELECT * FROM orders WHERE idaffiliate=" &pcv_Idaffiliate
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if rs.eof then
	'// Not associated with any orders, delete affiliate
	query="DELETE FROM affiliates WHERE idaffiliate=" &pcv_Idaffiliate
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	pcv_message="The affiliate was successfully deleted from the database."
	if err.number <> 0 then
		response.redirect "techErr.asp?error="& Server.Urlencode("<b><font color=#ff0000>Error occurred while trying to delete affiliate. [deAffb.asp]: </font></b>"&Err.Description) 
	end If
	set rs=nothing
	call closeDb()
	response.redirect "Adminaffiliates.asp?s=1&msg="& server.URLEncode(pcv_message)
else
	query="UPDATE affiliates SET pcAff_Active=0 WHERE idaffiliate=" &pcv_Idaffiliate
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	pcv_message="The affiliate you chose to delete is associated with one or more orders placed on the store. Therefore, it cannot be deleted. Instead, you can make this affiliate inactive."
	set rs=nothing
	call closeDb()
	response.redirect "Adminaffiliates.asp?msg="& server.URLEncode(pcv_message)
end if
%>
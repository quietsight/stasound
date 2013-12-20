<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Refund an Endicia Postage Label" %>
<% response.Buffer=true %>
<% section="orders" %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/currencyformatinc.asp" -->
<!--#include file="../includes/EndiciaFunctions.asp"-->
<!--#include file="AdminHeader.asp"-->
<%Dim connTemp,rs,query
Dim pcv_IDOrder,PackID,pcv_TrackNum,pcv_IsPIC

pcPageName="EDC_manage.asp"

call opendb()

call GetEDCSettings()

if request("id")="" then
	response.redirect "menu.asp"
end if

PackID=request("id")

if Not IsNumeric(PackID) then
	response.redirect "menu.asp"
end if

EDC_ErrMsg=""
EDC_SuccessMsg=""
msg=""

query="SELECT idOrder,pcPackageInfo_TrackingNumber,pcPackageInfo_EndiciaIsPIC FROM pcPackageInfo WHERE pcPackageInfo_ID=" & PackID & ";"
set rsQ=connTemp.execute(query)

if not rsQ.eof then
	pcv_IDOrder=rsQ("idOrder")
	pcv_TrackNum=rsQ("pcPackageInfo_TrackingNumber")
	pcv_IsPIC=rsQ("pcPackageInfo_EndiciaIsPIC")
else
	msg="The Package Information cannot be found"
	msgType=0	
end if
set rsQ=nothing

IF msg="" THEN
	tmpWebEDC=EDCURLSpc & "&method=RefundRequest"
	tmpXML=""
	tmpXML="<?xml version=""1.0"" encoding=""utf-8""?>"
	tmpXML=tmpXML & "<RefundRequest>"
	tmpXML=tmpXML & "<AccountID>" & EDCUserID & "</AccountID>"
	tmpXML=tmpXML & "<PassPhrase>" & EDCPassP & "</PassPhrase>"
	if DeveloperTest<>"" then
		tmpXML=tmpXML & "<Test>" & DeveloperTest& "</Test>"
	end if
	tmpXML=tmpXML & "<RefundList>"
	if pcv_IsPIC="1" then
		tmpXML=tmpXML & "<PICNumber>" & pcv_TrackNum & "</PICNumber>"
	else
		tmpXML=tmpXML & "<CustomsID>" & pcv_TrackNum & "</CustomsID>"
	end if
	tmpXML=tmpXML & "</RefundList>"
	tmpXML=tmpXML & "</RefundRequest>"
	tmpXML="XMLInput=" & Server.URLEncode(tmpXML)
	tmpWebEDC=tmpWebEDC & "&" & tmpXML
	result=ConnectServer(tmpWebEDC,"GET","","","")
	IF result="ERROR" or result="TIMEOUT" THEN
		msg="Cannot connect to Endicia Label Server"
		msgType=0
	ELSE
		tmpCode=FindStatusCode(result)
		if tmpCode="0" then
			tmpAppro=FindXMLValue(result,"IsApproved")
			if tmpAppro="YES" then
				msg="Your refund request was approved!<br>Returned Message: " & FindXMLValue(result,"ErrorMsg")
				msgType=1
				query="DELETE FROM pcPackageInfo WHERE pcPackageInfo_ID=" & PackID & ";"
				set rsQ=connTemp.execute(query)
				set rsQ=nothing
				query="UPDATE ProductsOrdered SET pcPackageInfo_ID=0,pcPrdOrd_Shipped=0 WHERE pcPackageInfo_ID=" & PackID & ";"
				set rsQ=connTemp.execute(query)
				set rsQ=nothing
				query="SELECT idProduct FROM ProductsOrdered WHERE idOrder=" & pcv_IDOrder & " AND pcPrdOrd_Shipped=1;"
				set rsQ=connTemp.execute(query)
				pcv_OrdStatus="3"
				if not rsQ.eof then
					pcv_OrdStatus="7"
				end if
				set rsQ=nothing
				query="UPDATE Orders Set orderStatus=" & pcv_OrdStatus & " WHERE idOrder=" & pcv_IDOrder & ";"
				set rsQ=connTemp.execute(query)
				set rsQ=nothing
				call SaveTrans(tmpXML,result,1,7)
			else
				msg=FindXMLValue(result,"ErrorMsg")
				EDC_ErrMsg=msg
				msgType=0
				call SaveTrans(tmpXML,result,0,7)
			end if
		else
			msg=FindXMLValue(result,"ErrorMsg")
			EDC_ErrMsg=msg
			msgType=0
			call SaveTrans(tmpXML,result,0,7)
		end if
	END IF
END IF
%>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>
<br>
<br>
&nbsp;&nbsp;&nbsp;<a href="OrdDetails.asp?id=<%=pcv_IDOrder%>">Back to order details &gt;&gt;</a><br><br>
<%call closedb()%>
<!--#include file="AdminFooter.asp"-->
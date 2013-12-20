<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/validation.asp"--> 
<!--#include file="../includes/securitysettings.asp" -->
<!--#include file="../includes/emailsettings.asp"--> 
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="pcStartSession.asp" -->
<!--#include file="pcVerifySession.asp"-->
<% Dim query,rs,connTemp
call openDb()

pcErrMsg=""

pCustomerEmail=getUserInput(request("e"),0)
pConCode=getUserInput(request("c"),0)

if pCustomerEmail="" then
	pcErrMsg=pcErrMsg & "<li>" & dictLanguage.Item(Session("language")&"_opc_consconf_1") & "</li>"
end if

if pConCode="" then
	pcErrMsg=pcErrMsg & "<li>" & dictLanguage.Item(Session("language")&"_opc_consconf_2") & "</li>"
end if

'// Get Target Customer
IF pcErrMsg="" THEN

	query="SELECT idCustomer,customerType,idCustomerCategory FROM Customers WHERE [email] like '" & pCustomerEmail & "' AND pcCust_ConsolidateStr like '" & pConCode & "';"
	set rs=connTemp.execute(query)
	if rs.eof then
		pcErrMsg=pcErrMsg & "<li>" & dictLanguage.Item(Session("language")&"_opc_consconf_3") & "</li>"
	else
		pcv_idtarget=rs("idCustomer")
		pcv_customerType=rs("customerType")
		pcv_idCustomerCategory=rs("idCustomerCategory")
	end if
	set rs=nothing

END IF


'// Start Consolidation
IF pcErrMsg="" THEN

	query="SELECT idCustomer FROM Customers WHERE [email] like '" & pCustomerEmail & "' AND idCustomer<>" & pcv_idtarget 
	set rs=connTemp.execute(query)
	intCount=-1
	if not rs.eof then
		pcArr=rs.getRows()
		set rs=nothing
		intCount=ubound(pcArr,2)
	end if
	set rs=nothing
	
	For i=0 to intCount
		pcv_idcustomer=pcArr(0,i)
		
		query="UPDATE ORDERS SET idcustomer=" & pcv_idtarget & " WHERE idcustomer=" & pcv_idcustomer
		set rs=connTemp.execute(query)
		set rs=nothing

		query="UPDATE DPRequests SET idcustomer=" & pcv_idtarget & " WHERE idcustomer=" & pcv_idcustomer
		set rs=connTemp.execute(query)
		set rs=nothing

		query="UPDATE authorders SET idcustomer=" & pcv_idtarget & " WHERE idcustomer=" & pcv_idcustomer
		set rs=connTemp.execute(query)
		set rs=nothing
		
		query="UPDATE pcPay_EIG_Authorize SET idcustomer=" & pcv_idtarget & " WHERE idcustomer=" & pcv_idcustomer
		set rs=connTemp.execute(query)
		set rs=nothing
		
		query="UPDATE pfporders SET idcustomer=" & pcv_idtarget & " WHERE idcustomer=" & pcv_idcustomer
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=nothing
		
		query="UPDATE netbillorders SET idcustomer=" & pcv_idtarget & " WHERE idcustomer=" & pcv_idcustomer
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=nothing
		
		'//Trap errors if Payment tables do not exist
		on error resume next
		
		query="UPDATE pcPay_LinkPointAPI SET idcustomer=" & pcv_idtarget & " WHERE idcustomer=" & pcv_idcustomer
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=nothing
		
		if err.number<>0 then
			err.clear
		end if
		
		query="UPDATE pcPay_PayPal_Authorize SET idcustomer=" & pcv_idtarget & " WHERE idcustomer=" & pcv_idcustomer
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=nothing
		
		if err.number<>0 then
			err.clear
		end if
		
		query="UPDATE pcPay_USAePay_Orders SET idcustomer=" & pcv_idtarget & " WHERE idcustomer=" & pcv_idcustomer
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=nothing
		
		if err.number<>0 then
			err.clear
		end if
		
		query="UPDATE pcPay_eMerch_Orders SET idcustomer=" & pcv_idtarget & " WHERE idcustomer=" & pcv_idcustomer
		set rs=connTemp.execute(query)
		set rs=nothing
		
		if err.number<>0 then
			err.clear
		end if
		
		on error goto 0
		
		query="UPDATE pcPPFCusts SET idcustomer=" & pcv_idtarget & " WHERE idcustomer=" & pcv_idcustomer
		set rs=connTemp.execute(query)
		set rs=nothing
		
		query="UPDATE wishList set idCustomer=" & pcv_idtarget & " WHERE idCustomer=" & pcv_idcustomer
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=nothing

		query="SELECT iRewardPointsAccrued,iRewardPointsUsed FROM customers WHERE idcustomer=" & pcv_idcustomer
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		
		pcv_RP=0
		pcv_RPU=0

		if not rs.eof then
			pcv_RP=rs("iRewardPointsAccrued")
			pcv_RPU=rs("iRewardPointsUsed")

			query="UPDATE customers SET iRewardPointsAccrued=0,iRewardPointsUsed=0 WHERE idcustomer=" & pcv_idcustomer
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
		end if
		set rs=nothing

		query="UPDATE customers SET iRewardPointsAccrued=iRewardPointsAccrued+" &pcv_RP & ",iRewardPointsUsed=iRewardPointsUsed+" & pcv_RPU & " WHERE idcustomer=" & pcv_idtarget
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=nothing
		
		'// Remove no-needed records
		query="DELETE FROM pcTaxEptCust WHERE idcustomer=" & pcv_idcustomer
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=nothing
		
		query="DELETE FROM recipients WHERE idcustomer=" & pcv_idcustomer
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=nothing
		
		query="DELETE FROM used_discounts WHERE idcustomer=" & pcv_idcustomer
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=nothing
		
		query="DELETE FROM pcCustomerFieldsValues WHERE idcustomer=" & pcv_idcustomer
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=nothing
		
		query="DELETE FROM pcCustomerSessions WHERE idcustomer=" & pcv_idcustomer
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=nothing
		
		query="DELETE FROM pcCustomerTermsAgreed WHERE idcustomer=" & pcv_idcustomer
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=nothing
		
		on error resume next '// required
		query="DELETE FROM pcQB_CMaps WHERE idcustomer=" & pcv_idcustomer
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=nothing
		error.clear '// required
		
		query="SELECT SavedCartID FROM pcSavedCarts WHERE idcustomer=" & pcv_idcustomer
		set rsQ=server.CreateObject("ADODB.RecordSet")
		set rsQ=connTemp.execute(query)
		if not rsQ.eof then
			tmpArr=rsQ.getRows()
			set rsQ=nothing
			intC=ubound(tmpArr,2)
			For k=0 to intC
				IDSC=tmpArr(0,k)
				query="DELETE FROM pcSavedCartArray WHERE SavedCartID=" & IDSC & ";"
				set rsQ2=server.CreateObject("ADODB.RecordSet")
				set rsQ2=connTemp.execute(query)
				set rsQ2=nothing
			Next
		end if
		set rsQ=nothing
		
		query="DELETE FROM pcSavedCarts WHERE idcustomer=" & pcv_idcustomer
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=nothing
		
	Next
	
	query="UPDATE Customers SET pcCust_Guest=0,pcCust_ConsolidateStr='' WHERE idCustomer=" & pcv_idtarget & ";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	set rs=nothing

	query="DELETE FROM Customers WHERE [email] like '" & pCustomerEmail & "' AND idCustomer<>" & pcv_idtarget 
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	set rs=nothing
	
	session("idCustomer")=pcv_idtarget
	session("CustomerGuest")=0
	Session("customerType")=pcv_customerType
	session("customerCategory")=pcv_idCustomerCategory
	Session("SFStrRedirectUrl")="CustPref.asp"
	
END IF

if pcErrMsg<>"" then
	pcErrMsg=dictLanguage.Item(Session("language")&"_opc_consconf_4") & "<br><ul>" & pcErrMsg & "</ul>"
end if

call closedb()
%>
<!--#include file="header.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<div id="pcMain">
	<table class="pcMainTable">
	<tr>
		<td>
			<%if pcErrMsg<>"" then%>
				<div class="pcErrorMessage">
					<%=pcErrMsg%>
				</div>
			<%else%>
				<div class="pcSuccessMessage">
					<%=dictLanguage.Item(Session("language")&"_opc_consconf_5")%>
				</div>
				<br />
				<input type="button" name="go" value="<%=dictLanguage.Item(Session("language")&"_opc_consconf_6")%>" onclick="location='login.asp?lmode=2';" class="submit2" />
			<%end if%>
		</td>
	</tr>
	</table>
</div>
<%call closedb()%>
<!--#include file="footer.asp"-->	
		
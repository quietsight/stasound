<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
pageTitle="Consolidate Customer Accounts" 
pageIcon="pcv4_icon_people.png"
section="mngAcc" 
%>
<%PmAdmin=7%><!--#include file="adminv.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/validation.asp" --> 
<!--#include file="../includes/languages.asp" --> 
<!--#include file="AdminHeader.asp"-->
<%
pcv_idcustomer=request("idcustomer")
if not validNum(pcv_idcustomer) then pcv_idcustomer=0

pcv_idtarget=request("idtarget")
if not validNum(pcv_idtarget) then pcv_idtarget=0

if (pcv_idcustomer=0) or (pcv_idtarget=0) then
	response.redirect "viewcusta.asp"
end if

dim connTemp, query
call opendb()

query="UPDATE ORDERS SET idcustomer=" & pcv_idtarget & " WHERE idcustomer=" & pcv_idcustomer
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
set rs=nothing

query="UPDATE DPRequests SET idcustomer=" & pcv_idtarget & " WHERE idcustomer=" & pcv_idcustomer
set rs=server.CreateObject("ADODB.RecordSet")
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
		
query="SELECT email,[name],lastname,iRewardPointsAccrued,iRewardPointsUsed FROM customers WHERE idcustomer=" & pcv_idcustomer
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if not rs.eof then
	pcv_email1=rs("email")
	pcv_custname1=rs("name") & " " & rs("lastname")
	pcv_RP=rs("iRewardPointsAccrued")
	pcv_RPU=rs("iRewardPointsUsed")

	query="UPDATE customers SET iRewardPointsAccrued=0,iRewardPointsUsed=0 WHERE idcustomer=" & pcv_idcustomer
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
end if
set rs=nothing

query="SELECT email,[name],lastname FROM customers WHERE idcustomer=" & pcv_idtarget
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if not rs.eof then
	pcv_email2=rs("email")
	pcv_custname2=rs("name") & " " & rs("lastname")
	query="UPDATE customers SET iRewardPointsAccrued=iRewardPointsAccrued+" &pcv_RP & ",iRewardPointsUsed=iRewardPointsUsed+" & pcv_RPU & " WHERE idcustomer=" & pcv_idtarget
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
end if
set rs=nothing

call closedb()%>

<form method="post" name="form1" class="pcForms">
	<table class="pcCPcontent">
		<tr> 
			<td>
            
            <div class="pcCPmessageSuccess">
                Customers accounts consolidated successfully!
                <br /><br />
                <span style="font-style: normal;">All orders <u>from</u> the customer account: <b><%=pcv_custname1 & " - " & pcv_email1%></b> were moved <u>to</u> the account: <b><%=pcv_custname2 & " - " & pcv_email2%></b> - <a href="viewCustOrders.asp?idcustomer=<%=pcv_idtarget%>" target="_blank">View Orders</a>. 
                    <%
                    if RewardsActive=1 then
                    %>
                    <br /><br /><%=RewardsLabel%> - if any - were also moved to this second account.
                    <%
                    end if
                    %>
                 </span>
             </div>
             
             <div class="pcCPmessage"><span style="font-style: normal;">Would you like to <strong>remove</strong> the customer account <%=pcv_custname1 & " - " & pcv_email1%> (Customer ID: <%=pcv_idcustomer%>), since it no longer has any orders associated with it? If &quot;Yes&quot;, <a href="javascript:if (confirm('Are you sure to want to continue?')) location='delCustomer.asp?idcustomer=<%=pcv_idcustomer%>'">click here</a>.</span></div>
             
            </td>
		</tr>
		<tr> 
			<td>&nbsp;</td>
		</tr>
		<tr> 
			<td align="center">
			<input type="button" name="viewOrders" value="View Orders for Consolidated Account" onClick="location.href='viewCustOrders.asp?idcustomer=<%=pcv_idtarget%>'">
			&nbsp;
            <input type="button" name="viewCust" value="Locate Another Customer" onClick="location.href='viewCusta.asp'">
			</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->
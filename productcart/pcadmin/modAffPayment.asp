<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=8%><!--#include file="adminv.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"-->
<%
Dim connTemp,rstemp,mySQL 
call opendb()
IDPayment=request("IDPayment")
idaffiliate=request("idaffiliate")

if request("ac")="update" then
PayDate=request("PayDate")
PayAmount=request("PayAmount")
PayStatus=request("PayStatus")

if scDB="Access" then
mySQL="update pcAffiliatesPayments  set pcAffpay_Amount=" & PayAmount & ",pcAffpay_PayDate=#" & payDate & "#,pcAffpay_Status='" & payStatus & "' where pcAffpay_IDPayment=" & IDPayment & " and pcAffpay_IDAffiliate=" & IDAffiliate
else
mySQL="update pcAffiliatesPayments  set pcAffpay_Amount=" & PayAmount & ",pcAffpay_PayDate='" & payDate & "',pcAffpay_Status='" & payStatus & "' where pcAffpay_IDPayment=" & IDPayment & " and pcAffpay_IDAffiliate=" & IDAffiliate
end if
set rstemp=connTemp.execute(mySQL)
set rstemp=nothing
call closeDb()
response.redirect "ReportAffiliateHistory.asp?IDAffiliate=" & idAffiliate
end if
%> 
<html>
<head>
</head>
<body>

<TABLE width="94%" border="0" cellpadding="4" cellspacing="0" align="center">
        
        <%
        Dim tempId
				tempId=0
				' Our Connection Object
				Dim con
				Set con=CreateObject("ADODB.Connection")
				con.Open scDSN 
	
				' Choose the records to display
				affVar=Request.QueryString("idaffiliate")
				If affVar="" OR affVar="0" OR affVar="1" then
					response.write "<tr><td><font face=""Arial, Helvetica, sans-serif"" size=""2"">You must specify an affiliate to be able to generate a report. You can do this by either entering an ID or choosing one from the drop-down list. <a href=srcOrdByDate.asp><font color=blue>Click here</font></a> to go back.</font><br></td></tr></table></td></tr></table></body></html>"
					response.end
				End If

				query="SELECT * FROM Orders WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12))"
				query=query&" AND idaffiliate="& affVar &" ORDER BY orders.orderDate desc"
				' Our Recordset Object
				Dim rs
				Set rs=CreateObject("ADODB.Recordset")
				rs.CursorLocation=adUseClient
				rs.Open query, scDSN , 3, 3
			
				' If the returning recordset is not empty
				If rs.EOF Then %>
					<tr> 
						<td height="21" colspan="5"><font face="Arial, Helvetica, sans-serif" size="2">Sorry, 
							no records match your query</font></td>
					</tr>
      	<% Else %>
					<%
					MySQL="SELECT idaffiliate,affiliateemail,affiliateName,affiliateAddress,affiliateAddress2,affiliatecity,affiliatestate,affiliatezip,affiliateCountryCode FROM affiliates WHERE idaffiliate="& affVar
					Set rsObjAff=CreateObject("ADODB.Recordset")
					rsObjAff.Open MySQL, scDSN , 3, 3
					%>
						<tr> 
							<td colspan="5"><font face="Arial, Helvetica, sans-serif" size="2" color="#000066"><b><img src="images/pc_individual.gif" width="14" height="15"> 
								Affiliate ID: <%=affVar%> </b></font></td>
						</tr>
						<tr> 
							<td colspan="5"><a href=mailto="<%=rsObjAff("affiliateemail")%>"><font face="Arial, Helvetica, sans-serif" size="2"><%=rsObjAff("affiliateName")%></font></a></td>
						</tr>
						<tr> 
							<td colspan="5"><font face="Arial, Helvetica, sans-serif" size="2"><%=rsObjAff("affiliateAddress")%> 
								<% if len(rsObjAff("affiliateaddress2"))>0 then%>
									<%response.write " ," & rsObjAff("affiliateaddress2")%>
								<% end If %>
								</font></td>
						</tr>
						<tr> 
							<td colspan="5"><font face="Arial, Helvetica, sans-serif" size="2"><%=rsObjAff("affiliatecity")%>, <%=rsObjAff("affiliatestate")%>&nbsp;<%=rsObjAff("affiliatezip")%></font></td>
						</tr>
						<tr> 
							<td colspan="5"><font face="Arial, Helvetica, sans-serif" size="2"><%=rsObjAff("affiliateCountryCode")%></font></td>
						</tr>
</TABLE>
                      <%mySQL="SELECT pcAffpay_idpayment,pcAffpay_Amount,pcAffpay_PayDate,pcAffpay_Status FROM pcAffiliatesPayments WHERE pcAffpay_idAffiliate=" & affVar & " and pcAffpay_IDPayment=" & IDPayment
						set rstemp=connTemp.execute(mySQL)
						if rstemp.eof then%>
						<font face="Arial, Helvetica, sans-serif" size="2" color="#FF0000"><br><b>This payment history not found.</b></font><br><br>
                      <%else
                      PaidAmount=rstemp("pcAffpay_Amount")
                      PaidDate=rstemp("pcAffpay_PayDate")
                      PaidStatus=rstemp("pcAffpay_Status")%>
<script language="JavaScript">
<!--
function isDigit(s)
{
var test=""+s;
if(test=="."||test==","||test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
		{
		return(true) ;
		}
		return(false);
	}
	
function allDigit(s)
	{
		var test=""+s ;
		for (var k=0; k <test.length; k++)
		{
			var c=test.substring(k,k+1);
			if (isDigit(c)==false)
			{
				return (false);
			}
		}
		return (true);
	}
		
function Form1_Validator(theForm)
{

	if (theForm.paydate.value == "")
  	{
			 alert("Please enter a value for this field.");
		    theForm.paydate.focus();
		    return (false);
	}
	if (theForm.payamount.value == "")
  	{
			 alert("Please enter a value for this field.");
		    theForm.payamount.focus();
		    return (false);
	}	
	if (theForm.payamount.value == "0")
	{
    alert("Please enter a value greater than zero for this field.");
    theForm.payamount.focus();
    return (false);
	}	
	if (allDigit(theForm.payamount.value) == false)
	{
    alert("Please enter a number for this field.");
    theForm.payamount.focus();
    return (false);
	}

return (true);
}
//-->
                      </script>
                    <form name="form1" action="modAffPayment.asp?ac=update&idAffiliate=<%=affVar%>&IDPayment=<%=IDPayment%>" method="post" onSubmit="return Form1_Validator(this)">
                    <table cellspacing="0" cellpadding="3" border="0" style="border-collapse: collapse" bordercolor="#111111" width="397">
                    <tr><td colspan="2"><font face="Arial, Helvetica, sans-serif" size="2"><b>Edit/Update payment: #<%=IDPayment%></b></font></td></tr>
                    <tr><td width="25%"><font face="Arial, Helvetica, sans-serif" size="2">Date:</font></td><td width="75">
                      <input type="text" name="paydate" size="30" value="<%=PaidDate%>"></td></tr>
                    <tr><td width="25%"><font face="Arial, Helvetica, sans-serif" size="2">Payment Amount:</font></td><td width="75">
                      <input type="text" name="payamount" size="30" value="<%=PaidAmount%>"></td></tr>
                    <tr><td width="25%"><font face="Arial, Helvetica, sans-serif" size="2">Payment Status:</font></td><td width="75">
                      <input type="text" name="paystatus" size="30" value="<%=PaidStatus%>"></td></tr>
                    <tr><td width="25%">&nbsp;</td><td width="75">
                      <input type=submit name="submit" value="Update this payment" class="ibtnGrey"></td></tr>
                    </table>
                    </form>
                    <br><br>
<%end if%>
<%end if%>
<%call closedb()%>
</body></html>
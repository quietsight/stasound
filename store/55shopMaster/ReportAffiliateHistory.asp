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
<!--#include file="../includes/dateinc.asp"--> 
<!--#include file="../includes/SQLFormat.txt"-->
<html>
<head>
<title>Managing Affiliates - Payments and Commissions Report</title>
</head>
<body onUnload="windowPrint.close()">

<TABLE width="94%" border="0" cellpadding="4" cellspacing="0" align="center">
        <%Dim connTemp,rstemp,mySQL 
        call opendb()
        
        if request("ac")="delpayment" then
        IDPayment=request("IDPayment")
        if IDPayment<>"" then
        mySQL="delete from pcAffiliatesPayments where pcAffpay_IDPayment=" & IDPayment
        set rstemp=connTemp.execute(mySQL)
        end if
        end if
        
        if request("ac")="addpayment" then
        PayDateVar=request("PayDate")
		'// Format for DB input
		PayDate=PayDateVar
		if PayDateVar<>"" then
			if scDateFrmt="DD/MM/YY" AND SQL_Format="0" then
				PayDateArray=split(PayDateVar,"/")
				PayDate=(PayDateArray(1)&"/"&PayDateArray(0)&"/"&PayDateArray(2))
			end if
		end if
        PayAmount=request("PayAmount")
        PayStatus=request("PayStatus")
        idaffiliate=request("idaffiliate")
        
        mySQL="insert into pcAffiliatesPayments (pcAffpay_idAffiliate,pcAffpay_Amount,pcAffpay_PayDate,pcAffpay_Status) values (" & IDAffiliate & "," & PayAmount & ",'" & PayDate & "','" & PayStatus & "')"
        set rstemp=connTemp.execute(mySQL)
        end if
        
        Dim tempId
				tempId=0
				' Our Connection Object
				Dim con
				Set con=CreateObject("ADODB.Connection")
				con.Open scDSN 
	
				' Choose the records to display
				affVar=Request("idaffiliate")
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
							<td colspan="5"><font face="Arial, Helvetica, sans-serif" size="2" color="#000066"><img src="images/pc_individual.gif" width="14" height="15">&nbsp;<b>Affiliate Name: <a href="mailto:<%=rsObjAff("affiliateemail")%>"><%=rsObjAff("affiliateName")%></a></b> (Affiliate ID: <%=affVar%>)</font></td>
						</tr>
						<tr> 
							<td colspan="5"><font face="Arial, Helvetica, sans-serif" size="2"><%=rsObjAff("affiliateAddress")%> 
								<% if len(rsObjAff("affiliateaddress2"))>0 then%>
									<%response.write ", " & rsObjAff("affiliateaddress2")%>
								<% end If %></font></td>
						</tr>
						<tr> 
							<td colspan="5"><font face="Arial, Helvetica, sans-serif" size="2"><%=rsObjAff("affiliatecity")%>, <%=rsObjAff("affiliatestate")%>&nbsp;<%=rsObjAff("affiliatezip")%>, <%=rsObjAff("affiliateCountryCode")%></font></td>
						</tr>
				<%
				mySQL="SELECT SUM(affiliatePay) AS AfftotalSum FROM orders WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND idAffiliate=" & affVar
				set rstemp=connTemp.execute(mySQL)
				AffTotalSum=rstemp("AfftotalSum")
				if AffTotalSum<>"" then
				else
				AffTotalSum=0
				end if
				mySQL="SELECT SUM(pcAffpay_Amount) AS AfftotalPaid FROM pcAffiliatesPayments WHERE pcAffpay_idAffiliate=" & affVar
				set rstemp=connTemp.execute(mySQL)
				AffTotalPaid=rstemp("AfftotalPaid")
				if AffTotalPaid<>"" then
				else
				AffTotalPaid=0
				end if
				CurrentBalance=AffTotalSum-AffTotalPaid
				%>
				<tr> 
					<td colspan="5"><font face="Arial, Helvetica, sans-serif" size="2"><b>Current balance: <font color="#FF0000"><%=scCursign%><%=money(CurrentBalance)%></font></b></font><br><br></td>
				</tr>
				<tr bgcolor="#e1e1e1"> 
					<td colspan="5"><font face="Arial, Helvetica, sans-serif" size="2"><b>Payment History:</b></font></td>
				</tr>
				<tr> 
					<td colspan="5">
                    <table border="0" cellpadding="3" cellspacing="5" width="100%" id="AutoNumber1">
                      <tr >
                        <td width="15%"><font face="Arial, Helvetica, sans-serif" size="2"><b>Date</b></font></td>
                        <td width="15%"><font face="Arial, Helvetica, sans-serif" size="2"><b>Amount</b></font></td>
                        <td width="60%"><font face="Arial, Helvetica, sans-serif" size="2"><b>Status</b></font></td>
                        <td width="10%"><font face="Arial, Helvetica, sans-serif" size="2"><b>Actions</b></font></td>
                      </tr>
                      <%mySQL="SELECT pcAffpay_idpayment,pcAffpay_Amount,pcAffpay_PayDate,pcAffpay_Status FROM pcAffiliatesPayments WHERE pcAffpay_idAffiliate=" & affVar & " ORDER BY pcAffpay_PayDate Desc;"
						set rstemp=connTemp.execute(mySQL)
						if rstemp.eof then%>
						<tr>
                        <td width="33%" colspan="4"><font face="Arial, Helvetica, sans-serif" size="2" color="#FF0000"><br><b>No payment history.</b></font><br><br></td>
                      </tr>
                      <%else
                      do while not rstemp.eof
                      IDpayment=rstemp("pcAffpay_idpayment")
                      PaidAmount=rstemp("pcAffpay_Amount")
                      PaidDate=rstemp("pcAffpay_PayDate")
					  '// Format date
					  PaidDate=ShowDateFrmt(PaidDate)
                      PaidStatus=rstemp("pcAffpay_Status")%>
                      <tr>
                        <td><font face="Arial, Helvetica, sans-serif" size="2"><%=PaidDate%></font></td>
                        <td><p align="right"><font face="Arial, Helvetica, sans-serif" size="2"><%=scCurSign%>&nbsp;<%=money(PaidAmount)%></font></td>
                        <td><font face="Arial, Helvetica, sans-serif" size="2"><%=PaidStatus%></font></td>
                        <td nowrap><font face="Arial, Helvetica, sans-serif" size="2"><a href="modAffpayment.asp?IDpayment=<%=IDPayment%>&idAffiliate=<%=affVar%>">Edit/Update</a>&nbsp;<a href="javascript:if (confirm('You are about to remove this payment from your database. Are you sure you want to complete this action?')) location='ReportAffiliateHistory.asp?ac=delpayment&IDpayment=<%=IDPayment%>&idAffiliate=<%=affVar%>'">Delete</a></font></td>
                      </tr>
                      <%rstemp.MoveNext
                      loop%>
                      <tr>
                        <td>
                        <p align="right"><font face="Arial, Helvetica, sans-serif" size="2"><b>Total Paid:</b></font></td>
                        <td><p align="right"><font face="Arial, Helvetica, sans-serif" size="2"><b><%=scCurSign%>&nbsp;<%=money(AffTotalPaid)%></b></font></td>
                        <td><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font></td>
                        <td></td>
                      </tr>
                      <%end if%>
                    </table>
                    <br>
<%if Cdbl(CurrentBalance)>0 then%>
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
										<hr noshade color="#e1e1e1">
                    <form name="form1" action="ReportAffiliateHistory.asp?ac=addpayment&idAffiliate=<%=affVar%>" method="post" onSubmit="return Form1_Validator(this)">
                    <table cellspacing="0" cellpadding="3" border="0" width="397">
                    <tr><td colspan="2"><font face="Arial, Helvetica, sans-serif" size="2"><b>Add new payment:</b></font></td></tr>
                    <tr><td width="25%"><font face="Arial, Helvetica, sans-serif" size="2">Date:</font></td><td width="75">
                      <input type="text" name="paydate" size="30" value="<%=ShowDateFrmt(Date())%>"></td></tr>
                    <tr><td width="25%"><font face="Arial, Helvetica, sans-serif" size="2">Payment Amount:</font></td><td width="75">
                      <input type="text" name="payamount" size="30" value="0"></td></tr>
                    <tr><td width="25%"><font face="Arial, Helvetica, sans-serif" size="2">Payment Status:</font></td><td width="75">
                      <input type="text" name="paystatus" size="30"></td></tr>
                    <tr><td width="25%">&nbsp;</td><td width="75"><input type=submit name="submit" value="Add payment" class="ibtnGrey"></td></tr>
                    </table>
                    </form>
										<hr noshade color="#e1e1e1">
                    <br>
<%end if%>
                    </td>
				</tr>
				<tr> 
					<td colspan="5"><font face="Arial, Helvetica, sans-serif" size="2"> 
						<% Response.Write "Total Sales Records Found : " & rs.RecordCount & "<br><br>" %>
						</font></td>
				</tr>
				<tr> 
					<td width="12%" bgcolor="#e1e1e1"> <div align="left"><b><font face="Arial, Helvetica, sans-serif" size="2">Date</font></b></div></td>
					<td width="31%" bgcolor="#e1e1e1"> <div align="left"><b><font face="Arial, Helvetica, sans-serif" size="2">Customer</font></b></div></td>
					<td width="17%" bgcolor="#e1e1e1"> <div align="right"><b><font face="Arial, Helvetica, sans-serif" size="2">Total 
							of Sale</font></b></div></td>
					<td bgcolor="#e1e1e1" width="15%"> <div align="right"><b><font face="Arial, Helvetica, sans-serif" size="2">Tax</font></b></div></td>
					<td width="25%" bgcolor="#e1e1e1"> <div align="right"><b><font face="Arial, Helvetica, sans-serif" size="2">Commission</font></b></div></td>
				</tr>
				<% 
				gTotalsales=0
				gTotaltaxes=0
				gTotalcom=0
				do until rs.EOF
					MySQL="SELECT idaffiliate,affiliateemail,affiliateName,affiliateAddress,affiliateAddress2,affiliatecity,affiliatestate,affiliatezip,affiliateCountryCode FROM affiliates WHERE idaffiliate="& rs("idaffiliate")
					Set rsObjAff=CreateObject("ADODB.Recordset")
					rsObjAff.Open MySQL, scDSN , 3, 3
					
					querySTR="SELECT * FROM ProductsOrdered WHERE idorder="& rs("idorder")
					Set rsSTR=CreateObject("ADODB.Recordset")
					rsSTR.CursorLocation=adUseClient
					rsSTR.Open querySTR, scDSN , 3, 3
					bOrderTotal=0
					do until rsSTR.eof
						querySTR="SELECT name,lastname FROM customers WHERE idcustomer="& rs("idcustomer")
						Set rsCust=CreateObject("ADODB.Recordset")
						rsCust.CursorLocation=adUseClient
						rsCust.Open querySTR, scDSN , 3, 3
						CustName=rsCust("name")& " "&rsCust("lastname")
						rsCust.Close
						set rsCust=nothing
						unitTotal=rsSTR("unitPrice")
						quantity=rsSTR("quantity")
						bOrderTotal=0 + (unitTotal * quantity)
					rsSTR.moveNext
					loop
					rsSTR.Close
					set rsSTR=nothing
					gTotalsales=gTotalsales + rs("total")
					gTotaltaxes=gTotaltaxes + rs("taxAmount") 
					%>
					<% '// Format Date
					dtOrderDate=rs("orderDate")
					dtOrderDate=ShowDateFrmt(dtOrderDate) %>
        <tr> 
          <td height="21" width="12%"> <div align="left"><font face="Arial, Helvetica, sans-serif" size="2"><%=dtOrderDate%></font></div></td>
          <td height="21" width="31%"> <div align="left"><font face="Arial, Helvetica, sans-serif" size="2"><%=CustName%></font></div></td>
          <td height="21" width="17%"> <div align="right"><font face="Arial, Helvetica, sans-serif" size="2"><%=scCurSign&money(rs("total"))%></font></div></td>
          <td height="21" width="15%"> <div align="right"><font face="Arial, Helvetica, sans-serif" size="2"><%=scCurSign&money(rs("taxAmount"))%></font></div></td>
          <td height="21" width="25%"> <div align="right"><font face="Arial, Helvetica, sans-serif" size="2"><%=scCurSign&money(rs("affiliatePay"))%></font></div></td>
        </tr>
        <% gTotalcomm=gTotalcomm + rs("affiliatePay") %>
			<% rs.MoveNext
			loop
		End If %>
        <tr> 
          <td height="21" colspan="5"> <hr width="100%" size="1" noshade> </td>
        </tr>
        <tr> 
          <td height="14" colspan="3"><div align="right"><font size="2" face="Arial, Helvetica, sans-serif"><b>Total 
              of all Sales</b></font></div></td>
          <td height="14" width="15%"><div align="right"><font size="2" face="Arial, Helvetica, sans-serif"><b>Tax 
              Amount</b></font></div></td>
          <td height="14" width="25%"><div align="right"><font face="Arial, Helvetica, sans-serif" size="2"><b>Total 
              Commissions</b></font></div></td>
        </tr>
        <tr> 
          <td height="21" colspan="2">&nbsp;</td>
          <td height="21" width="17%"><div align="right"><font size="2" face="Arial, Helvetica, sans-serif"><b><font color="#FF0000"><%=scCurSign&money(gTotalsales)%></font></b></font></div></td>
          <td height="21" width="15%"><div align="right"><font size="2" face="Arial, Helvetica, sans-serif"><b><font color="#FF0000"><%=scCurSign&money(gTotaltaxes)%></font></b></font></div></td>
          <td height="21" width="25%"><div align="right"><font size="2" color="#FF0000" face="Arial, Helvetica, sans-serif"><b><%=scCurSign&money(gTotalcomm)%></b></font></div></td>
        </tr>
      </table>


<%	' Done. Now release Objects
	con.Close
	Set con=Nothing
	Set rs=Nothing
%>
<%call closedb()%>
</body>
</html>
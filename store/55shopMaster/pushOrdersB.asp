<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
pageTitle="Consolidate Customer Accounts" 
pageIcon="pcv4_icon_people.png"
section="mngAcc" 
%>
<%PmAdmin=7%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/encrypt.asp"-->
<!--#include file="../includes/rc4.asp"-->
<%

strORD="lastName"
strSort="ASC"

pMode=request.Querystring("mode")
pcv_customerName=request("customerName")
pcv_idcustomer=request("idcustomer")
if pcv_idcustomer="" then
	pcv_idcustomer = request.QueryString("pidCustomer")
end if
if not validNum(pcv_idcustomer) then response.Redirect "viewCusta.asp"
if pcv_customerName="" then  response.Redirect "pushOrdersA.asp?idcustomer=" & pcv_idcustomer & "&msg=" & server.URLEncode("Please enter the customer's name, company name, or e-mail address.")

dim query, conntemp, rs

call openDb()

pcv_customerName=replace(pcv_customerName,"'","''")
customerNameArray=split(pcv_customerName," ")
if ubound(customerNameArray)>0 then
	query="SELECT LastName,[name],customerCompany,phone,customerType,idcustomer,email FROM customers WHERE (idcustomer <> "& pcv_idcustomer & " AND "
	for i=0 to ubound(customerNameArray)
		if i>0 then
			query=query&"OR (idcustomer <> "& pcv_idcustomer & " AND "
		end if
		query=query&"([name] LIKE '%" &customerNameArray(i)& "%' OR LastName LIKE '%" &customerNameArray(i)& "%' OR customerCompany LIKE '%" &customerNameArray(i)& "%' OR email LIKE '%" &customerNameArray(i)& "%')) "
	next
	query=query&" ORDER BY "& strORD &" "& strSort
else
	query="SELECT LastName,[name],customerCompany,phone,customerType,idcustomer,email FROM customers WHERE (idcustomer <> "& pcv_idcustomer & " AND ([name] LIKE '%" &pcv_customerName& "%' OR LastName LIKE '%" &pcv_customerName& "%' OR customerCompany LIKE '%" &pcv_customerName& "%' OR email LIKE '%" &pcv_customerName & "%')) ORDER BY "& strORD &" "& strSort
end if

Set rs=Server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query)
	
if rs.eof then
	set rs=nothing
	call closedb()
	response.redirect "msg.asp?message=26"
end if %>

<!--#include file="AdminHeader.asp"-->
<script>
function Form1_Validator(theForm)
	{
		if (theForm.idtarget.value == "")
			{
				alert("Please select a customer account.");
				return (false);
			}
		return(true);
	}
</script>
<form name="form1" method="post" action="pushOrdersC.asp?action=push" onSubmit="return Form1_Validator(this)" class="pcForms">
	<input type="hidden" name="idcustomer" value="<%=pcv_idcustomer%>">
	<input type="hidden" name="idtarget" value="">
	<table class="pcCPcontent">          
		<tr>             
			<td>
			<% 
			query="SELECT email,[name],lastname FROM customers WHERE idcustomer=" & pcv_idcustomer
			set rstemp=server.CreateObject("ADODB.RecordSet")
			set rstemp=connTemp.execute(query)
			
			if not rs.eof then
				pcv_email=rstemp("email")
				pcv_name=rstemp("name")
				pcv_lastname=rstemp("lastname")
			end if
			set rstemp=nothing
			%>
			All orders currently under &quot;<b><%=pcv_name & " " & pcv_lastname & " - " & pcv_email%></b>&quot; will be moved to the account you select below.
			</td>
		</tr>
        <tr>
            <td>
                <table class="pcCPcontent">
                    <tr> 
                        <td colspan="5" class="pcCPspacer"></td>
                    </tr>
                    <tr> 
                        <th nowrap>&nbsp;</th>
                        <th nowrap>Name</th>
                        <th nowrap>Company</th>
                        <th nowrap>E-mail</th>
                        <th nowrap>Type</th>
                    </tr>
                    <tr> 
                        <td colspan="5" class="pcCPspacer"></td>
                    </tr>
                    <% Dim Count
                    Count=0
                    pcArray=rs.getRows()
                    intCount=Ubound(pcArray,2)
                    set rs=nothing
                    For i=0 to intCount
                        count=count + 1
                        pLastName=pcArray(0,i)
                        pname=pcArray(1,i)
                        pcustomerCompany=pcArray(2,i)
                        pphone=pcArray(3,i)
                        pcustomerType=pcArray(4,i)
                        pidcustomer=pcArray(5,i)
                        pemail=pcArray(6,i)
                        %>
                  
                        <tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
                            <td align="right"><input type="radio" name="ID" value="<%=count%>" onclick="document.form1.idtarget.value='<%=pidcustomer%>';" class="clearBorder"></td>
                            <td>
                            <a href="modcusta.asp?idcustomer=<%=pidcustomer%>" target="_blank"><%response.write (pLastName&", "&pname)%></a>
                            <% if pcustomerType="3" then%>
                            <img src="images/pcadmin_lockedaccount.jpg" width="12" height="12"> 
                            <% end if %>
                            </td>
                            <td align="left"><%response.write (pcustomerCompany)%></td>
                            <td nowrap><%=pemail%></td>
                            <td nowrap class="pcSmallText"><% if pcf_GetCustType(pidcustomer)=0 then%>Registered customer<% else %>Guest account<% end if %></td>
                        </tr>
                    <% Next
					call closeDb()
					%>
                </table>
            </td>
        </tr>
        <tr>
            <td align="center"><br>
            <input type="submit" name="submit" value="Move orders to selected account" class="submit2">&nbsp;
            <input type="button" name="Button" value="Back" onClick="javascript:history.back()">
            </td>
        </tr>
    </table>
</form>
<!--#include file="AdminFooter.asp"-->
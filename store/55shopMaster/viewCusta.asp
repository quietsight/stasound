<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="Locate a Customer" 
pageIcon="pcv4_icon_people.png"
pcStrPageName="viewCusta.asp"
section="mngAcc" 
%>
<%PmAdmin="7*9*"%><!--#include file="adminv.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/stringfunctions.asp" -->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%Dim connTemp,rs,query

call opendb()
%>
<!--#include file="pcCharts.asp"-->
<form method="post" name="listCust" action="viewCustb.asp" class="pcForms">
	<table class="pcCPcontent" style="margin: 0 0 20px 0;">
		<%
				strAction = request.QueryString("action")
				pcv_IDCustomer = request.QueryString("idCustomer")
				if strAction = "added" then %>
				<tr>
					<td colspan="4" align="center">
						<div class="pcCPmessageSuccess">
							New customer added successfully
							<%if request.QueryString("mailup")="0" then%>
							<br>
							<%=dictLanguage.Item(Session("language")&"_MailUp_SynNote1A")%>
							<%end if%>
							<% if validNum(pcv_IDCustomer) then %>
							<br /><br />
							<a href="modCusta.asp?idcustomer=<%=pcv_IDCustomer%>">View &amp; edit this new customer &gt;&gt;</a>
							<% end if %>
						</div>
					</td>
				</tr>
				<tr>
					<td colspan="4" class="pcCPspacer"></td>
				</tr>
		<% else %>
        <tr>
            <td colspan="4" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
        <% end if %>
		<tr> 
			<th colspan="4">Find customers using one or more of the fields below</th>
		</tr>
		<tr>
			<td colspan="4" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td align="right" valign="top">E-mail address:</td>
			<td valign="top"><input type="text" name="key4" size="40" value=""></td>
			<td align="right" nowrap valign="top">Customer Type</td>
			<td valign="top">
				<select name="customerType">
					<option value='' selected>All</option>
					<option value='0'>Retail Customer</option>
					<option value='1'>Wholesale Customer</option>
					<% 'if there are pricing categories - List them here
					
					query="SELECT idcustomerCategory, pcCC_Name, pcCC_CategoryType FROM pcCustomerCategories;"
					SET rs=Server.CreateObject("ADODB.RecordSet")
					SET rs=conntemp.execute(query)
					if NOT rs.eof then 
						do until rs.eof 
							intIdcustomerCategory=rs("idcustomerCategory")
							strpcCC_Name=rs("pcCC_Name")
							%>
							<option value='CC_<%=intIdcustomerCategory%>'><%=strpcCC_Name%></option>
							<% rs.moveNext
						loop
					end if
					SET rs=nothing
					call closedb()
					%>
				</select>
			</td>
    	</tr>
        
		<tr> 
			<td align="right" valign="top" width="15%" nowrap>First Name:</td>
			<td width="35%" nowrap valign="top"><input type="text" name="key1" size="40" value=""></td>
			<td align="right" valign="top" width="15%" nowrap>Last Name:</td>
			<td width="35%" nowrap valign="top"><input type="text" name="key2" size="40" value=""></td>
		</tr>
    
		<tr> 
			<td align="right" valign="top">Company:</td>
			<td valign="top"><input type="text" name="key3" size="40" value=""></td>
			<td align="right" valign="top">Phone number:</td>
			<td valign="top"><input type="text" name="key6" size="40" value=""></td>
		</tr>
    
  		<tr> 
			<td align="right" valign="top">City:</td>
			<td valign="top"><input type="text" name="key5" size="40" value=""></td>
			<td align="right" valign="top">State Code:</td>
			<td valign="top"><input type="text" name="key7" size="40" value=""></td>
		</tr>
		
		<tr> 
			<td align="right" valign="top">Province:</td>
			<td valign="top"><input type="text" name="key8" size="40" value=""></td>
			<td align="right" valign="top">Zip Code:</td>
			<td valign="top"><input type="text" name="key9" size="40" value=""></td>
		</tr>
		
		<tr> 
			<td align="right" valign="top">Country Code:</td>
			<td colspan="3" valign="top">
            	<input type="text" name="key10" size="40" value="">
                <input type="checkbox" name="key11" value="1" class="clearBorder"> Exclude customers in the country specified.
			</td>
		</tr>
    
		<tr> 
			<td colspan="4"><hr></td>
		</tr>
		<tr> 
			<td colspan="4" align="center">
				<input type="submit" name="srcView" value="Search" class="submit2">
				&nbsp;<input type="button" value="View All" onclick="location.href='viewCustb.asp?mode=ALL'">
				&nbsp;<input type="button" value="View Last 10" onclick="location.href='viewCustb.asp?mode=LAST'">
				&nbsp;<input type="button" value="Add New" onclick="location.href='instCusta.asp'">
				&nbsp;<input type="button" value="Back" onClick="javascript:history.back()">
			</td>
		</tr>
	</table>
    <% if scGuestCheckoutOpt<>"2" then %>
	<table class="pcCPcontent" style="margin: 10px 0;">
		<tr> 
			<th>Customer registrations and sales for last 30 days</th>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr>
		<td>
			<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="50%" valign="top">
					<div id="chartNewCusts30days" style="height:250px; "></div>
				</td>
				<td width="50%" valign="top">
					<div id="chartTop10Custs30days" style="height:250px; "></div>
				</td>
			</tr>
			</table>
			<%
			call pcs_NewCusts30Days("chartNewCusts30days")
			call pcs_Top10Custs30Days("chartTop10Custs30days")
			%>
		</td>
		</tr>
	</table>
    <% end if %>
</form>
<!--#include file="AdminFooter.asp"-->
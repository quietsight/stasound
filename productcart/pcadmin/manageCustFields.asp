<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Special Customer Fields" %>
<% Section="layout" %>
<%PmAdmin=7%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<% 
Dim connTemp,rs,query,rstemp

if request("submit")<>"" then
	Count=request("Count")
	if Count="" then
		Count="0"
	end if
	call opendb()
	For i=1 to Count
		cfOrder=request("cfOrder" & i)
		if cfOrder="" then
			cfOrder=0
		end if
		cfID=request("cfID" & i)
		query="UPDATE pcCustomerFields SET pcCField_Order=" & cfOrder & " WHERE pcCField_ID=" & cfID
		response.write query
		set rstemp=Server.CreateObject("ADODB.Recordset")
		set rstemp=connTemp.execute(query)
		set rstemp=nothing
	Next
	call closedb()
	response.redirect "manageCustFields.asp?s=1&message=5"
end if

call opendb()
query="SELECT pcCField_ID, pcCField_Name, pcCField_FieldType, pcCField_Required, pcCField_PricingCategories, pcCField_Order FROM pcCustomerFields ORDER BY pcCField_Order, pcCField_Name ASC;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

pcv_HaveRecord=0
if not rs.eof then
	pcv_HaveRecord=1
	pcArr=rs.GetRows()
	pcv_intCount=ubound(pcArr,2)
end if

set rs=nothing

%>

<% ' START get message, if any
msg=""
Select Case request("message")
	Case 2:	msg="The Special Customer Field was removed successfully!"
	Case 3:	msg="Special Customer Field field was added successfully!"
	Case 4:	msg="Special Customer Field updated successfully!"
	case 5: msg="Special Customer Fields ordered successfully."
End Select
msgtype=1
	' END get message, if any
%>
	 
<form action="manageCustFields.asp" class="pcForms">
	<table class="pcCPcontent">
        <tr>
            <td colspan="3" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
	<%
		IF pcv_HaveRecord="0" THEN
	%>
		<tr>
			<td colspan="6" align="center"><div class="pcCPmessage">No Special Customer Field was found.</div></td>
		</tr>
	<%
		' END show messages
		ELSE
	%>
		<tr>
        	<th valign="top" nowrap width="2%">Order</th>
			<th valign="top" nowrap width="58%">Field Name</th>
			<th valign="top" nowrap width="10%">Type</th>
			<th valign="top" nowrap width="10%">Required</th>
			<th valign="top" colspan="2" width="20%">Pricing Categories</th>
		</tr>
		<tr>
			<td colspan="6" class="pcCPspacer"></td>
		</tr>
		<%
		For i=0 to pcv_intCount
		Count=Count+1
		%>
		<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
        	<td>
            	<input type="text" name="cfOrder<%=Count%>" value="<%=pcArr(5,i)%>" size="3">
                <input type="hidden" name="cfID<%=Count%>" value="<%=pcArr(0,i)%>">
            </td>
	    	<td><a href="addmodCustField.asp?idcustfield=<%=pcArr(0,i)%>"><%=pcArr(1,i)%></a></td>
			<td><%if pcArr(2,i)="1" then%>Checkbox<%else%>Input Field<%end if%></td>
			<td><%if pcArr(3,i)="1" then%>Yes<%else%>No<%end if%></td>
			<td><%if pcArr(4,i)="1" then%>Yes<%else%>No<%end if%></td>
			<td nowrap align="right"><a href="addmodCustField.asp?idcustfield=<%=pcArr(0,i)%>"><img src="images/pcIconGo.jpg" width="12" height="12" alt="Edit" title="Edit"></a>
			<%query="SELECT pcCField_ID FROM pcCustomerFieldsValues WHERE pcCField_ID=" & pcArr(0,i) & ";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
			if rs.eof then%>
			&nbsp;<a href="javascript:if (confirm('You are about to remove this Special Customer Field from your database. Are you sure you want to complete this action?')) location='delCustField.asp?idcustfield=<%=pcArr(0,i)%>'"><img src="images/pcIconDelete.jpg" width="12" height="12" alt="Delete" title="Delete"></a>
			<%else%>
			&nbsp;<a href="javascript:if (confirm('You are about to remove this Special Customer Field and all of the values that customers have saved in it when registering or checking out on your store. Instead of removing the field, you can simply hide it from the storefront by unchecking the Show On Checkout Page and Show on Registration Page checkboxes when modifying the field. This action cannot be undone. Are you sure you want to continue?')) location='delCustField.asp?idcustfield=<%=pcArr(0,i)%>'"><img src="images/pcIconDelete.jpg" width="12" height="12" alt="Delete" title="Delete"></a>
			<%end if
			set rs=nothing%>
			</td>
		</tr>
		<%Next%>
	<%END IF%>
		<tr>
			<td colspan="6" class="pcSpacer">&nbsp;</td>
		</tr>
		<tr> 
			<td colspan="6" align="center">
            <% if pcv_HaveRecord>0 then %>
            <input type="submit" name="submit" value="Update Order" class="submit2">
            <input name="Count" type="hidden" value="<%=Count%>">
            &nbsp;
            <% end if %>
			<input type="button" name="add" value="Add New Special Customer Field" onclick="javascript:location='addmodCustField.asp';">
			&nbsp;
			<input type="button" name="back" value="Back" onClick="javascript:history.back()">
			</td>
		</tr>
	</table>
</form>
<%call closedb()%>
<!--#include file="AdminFooter.asp"-->
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%Dim connTemp
Dim AList(9999)
%>
<% nav=request("nav")
pageTitle="Manage Custom Fields" %>
<%
if nav="bto" then
	section="services"
	else
	section="products"
end if 

call opendb()

' Remove a custom field from the database:

pcv_strAction = request.QueryString("action")
pcv_intIdCustom = request.QueryString("idcustom")
pcv_strCustomType = request.QueryString("type")
if pcv_strAction = "del" then
	if pcv_strCustomType = "S" then
		query="DELETE FROM customfields WHERE idcustom="&pcv_intIdCustom
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
		msg="Custom search field deleted successfully!"
		msgtype=1
		set rs = nothing
	else
		query="DELETE FROM xfields WHERE idxfield="&pcv_intIdCustom
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
		msg="Custom inpout field deleted successfully!"
		msgtype=1
		set rs = nothing
	end if
end if

' End remove custom field

if request("action")="updfield" then
	idcustom=mid(request("idcustom"),2,len(request("idcustom")))
	newvalue=request("newvalue")

	if Left(request("idcustom"),1)="S" then
		query="UPDATE pcSearchFields SET pcSearchFieldName='" & newvalue & "' WHERE idSearchField=" & idcustom
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
	else
		query="UPDATE xfields SET xfield='" & newvalue & "' WHERE idxfield=" & idcustom
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
	end if

	msg="Custom field was renamed successfully!"
	msgtype=1

end if

set rs=nothing
call closedb()

%>
<!--#include file="AdminHeader.asp"-->
<script>
function newWindow(file,window)
	{
		msgWindow=open(file,window,'resizable=yes,scrollbars=yes,width=400,height=500');
		if (msgWindow.opener == null) msgWindow.opener = self;
	}
</script>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr> 
		<th>Custom Field Overview</th>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
		<p>Add custom fields to products to collect, display, or search upon additional product information. There are two types of custom fields:</p>
		<ul>
			<li><u>Input fields</u> allow you to collect information from the customer (e.g. name to be embroidered on the front of a polo shirt)&nbsp;<a href="http://wiki.earlyimpact.com/productcart/input_fields_manage" target="_blank"><img src="images/pcv3_infoIcon.gif" alt="More information on this topic" width="16" height="16" border="0"></a></li>
			<li><u>Search fields</u> allow you to add searchable properties to products (e.g. wine store: year, wine region, wine type, etc.)&nbsp;<a href="http://wiki.earlyimpact.com/productcart/managing_search_fields" target="_blank"><img src="images/pcv3_infoIcon.gif" alt="More information on this topic" width="16" height="16" border="0"></a></li>
		</ul>
     </td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr> 
		<th>Manage Custom Search Fields</th>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
    	<p>You can easily <b>add/edit/delete</b> custom search fields and their values.</p>
      <ul>
        <li><a href="ManageSearchFields.asp">Manage Custom Search Fields</a></li>
        <li><a href="addSFtoPrds.asp?nav=<%=nav%>">Add a custom search field to selected <strong>products</strong></a></li>
        <li><a href="addSFtoCats.asp">Add a custom search field to selected <strong>categories</strong></a> - This allows customers to filter products when they <em>browse by category</em>. You must also turn on the Customer Search Widget on the <a href="SearchOptions.asp" target="_blank">Search Options</a> page.</li>
      </ul>
		</td>
	</tr>      
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr> 
		<th>Manage Custom Input Fields</th>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
  <tr>
  	<td>Custom input fields allow you to let customers provide information on specific products (e.g. text to be embroidered on a shirt).
      <ul>
        <li><a href="addCFtoPrds.asp?nav=<%=nav%>">Add a custom input field to selected <strong>products</strong></a></li>
        <li>Use the features below to remove/rename input fields on multiple products at once.</li>
      </ul>
		<%
		call opendb()
		mytest1=false
		mytest2=false
		query="SELECT idSearchField,pcSearchFieldName FROM pcSearchFields order by pcSearchFieldOrder asc,pcSearchFieldName asc;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)
		if not rs.eof then
			mytest1=true
		end if
		query="SELECT idxfield,xfield FROM xfields order by xfield asc;"
		set rs1=Server.CreateObject("ADODB.Recordset")
		set rs1=connTemp.execute(query)
		if not rs1.eof then
			mytest2=true
		end if
		
		tmpStr=""
		tmpStr1=""
					
		if not rs.eof then
			pcArray=rs.getRows()
			intCount=ubound(pcArray,2)
			For i=0 to intCount
				tmpStr1=tmpStr1 & "<option value=""S" & pcArray(0,i) & """>" & pcArray(1,i) & "</option>" & vbcrlf
			Next
		end if
					
		if not rs1.eof then
			pcArray=rs1.getRows()
			intCount=ubound(pcArray,2)
			For i=0 to intCount
				tmpStr1=tmpStr1 & "<option value=""C" & pcArray(0,i) & """>" & pcArray(1,i) & "</option>" & vbcrlf
				tmpStr=tmpStr & "<option value=""C" & pcArray(0,i) & """>" & pcArray(1,i) & "</option>" & vbcrlf
			Next
		end if
		
		set rs=nothing
		set rs1=nothing
		%>

		<%if (mytest2=true) then%>      
			<Form action="ManageCFields.asp?action=updfield" method="post" name="Form1" class="pcForms">
				<table class="pcCPcontent" style="width:auto">
					<tr>
						<td colspan="2"><a name="1"></a><b>Rename an existing custom input field</b></td>
					</tr>
					<tr>
						<td width="20%" nowrap>Custom Input Field:</td>
		    		<td width="80%">
							<select name="idcustom">
							<%=tmpStr%>
							</select>
						</td>
					</tr>
					<tr>
						<td>New Name:</td>
						<td><input type="text" size="20" value="" name="newvalue"></td>
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td> 
						<input type="submit" name="submit" value="Update" class="submit2" onclick="javascript: if (document.Form1.newvalue.value=='') {alert('Please enter a value for New Name'); document.Form1.newvalue.focus();return(false)} else {return(true)};">
						&nbsp;
						<input type="button" name="show" value="Show Products" onclick="javascript: newWindow('showCFProducts.asp?idcustom='+document.Form1.idcustom.value,'products');">
						</td>
					</tr>
          <tr>
            <td class="pcCPspacer" colspan="2"></td>
          </tr>
				</table>
			</Form>
		<%end if%>
	
			
	<%if (mytest1=true) or (mytest2=true) then%>
	<Form action="delCFfromPrds1.asp?action=delfield" method="post" name="Form3" class="pcForms">
		<table class="pcCPcontent" style="width:auto;">
		<tr>
			<td colspan="2"><b><a name="3"></a>Remove an existing custom field from products</b></td>
		</tr>
		<tr>
			<td width="20%" nowrap>Custom Field:</td>
			<td width="80%">
			<select name="idcustom">
			<%=tmpStr1%>
			</select>
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td> 
			<input type="submit" name="submit" value="Select Products" class="submit2">
			&nbsp;
			<input type="button" name="show" value="Show Products" onclick="javascript: newWindow('showCFProducts.asp?idcustom='+document.Form3.idcustom.value,'products');"></td>
		</tr>
		</table>
	</Form>
	<%end if%>
	</td>
	</tr>
	<tr>
		<td><hr></td>
	</tr>
	<tr>
		<td align="center">
			<form class="pcForms">
			<input type="button" onClick="JavaScript: location.href='LocateProducts.asp?cptype=0'" value="Locate a Product">&nbsp;
			<input type="button" onClick="JavaScript: location.href='menu.asp'" value="Start Page">
			</form>
		</td>
	</tr>
	</table>
<%call closedb()%><!--#include file="AdminFooter.asp"-->
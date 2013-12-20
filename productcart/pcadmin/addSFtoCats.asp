<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add a Custom Search Field to Selected Categories" %>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/utilities.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
Dim query, conntemp, rs
Dim AList(9999)

If request("action")="new" then
	
	SF_name=request("SF_name")
	if SF_name="" then
		response.redirect "addSFtoCats.asp?x="&x&"&message="&server.URLEncode("You must specify a field name.")
	else
		SF_name=replace(SF_name,"'","''")
	end if

	call openDb()

	SF_name=replace(SF_name,"'","''")
	SF_show=request("SF_show"&k)
	CP_show=request("CP_show"&k)
	SF_showSEARCH=request("SF_showSEARCH"&k)
	CP_showSEARCH=request("CP_showSEARCH"&k)
	if SF_show="" then
		SF_show=0
	end if
	if CP_show="" then
		CP_show=0
	end if
	if SF_showSEARCH="" then
		SF_showSEARCH=0
	end if
	if CP_showSEARCH="" then
		CP_showSEARCH=0
	end if		
	SF_order=request("SF_order")
	if SF_order="" then
		SF_order=0
	end if
	query="SELECT idSearchField FROM pcSearchFields WHERE pcSearchFieldName like '" & SF_name & "';"
	set rsQ=connTemp.execute(query)
	if not rsQ.eof then
		msg="The field name was already added!"
		set rsQ=nothing
	else
		query="INSERT INTO pcSearchFields (pcSearchFieldName,pcSearchFieldOrder,pcSearchFieldShow,pcSearchFieldCPShow,pcSearchFieldSearch,pcSearchFieldCPSearch) VALUES ('" & SF_name & "'," & SF_order & "," & SF_show & "," & CP_show & "," & SF_showSEARCH & "," & CP_showSEARCH & ");"
		set rsQ=connTemp.execute(query)
		set rsQ=nothing
		msg="The new custom search field was added successfully!"
		msgtype=1
	end if
	set rsQ=nothing
	
	'get idcustom
	query="SELECT idSearchField FROM pcSearchFields WHERE pcSearchFieldName like '"&SF_name&"';"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=conntemp.execute(query)
	idcustom=rs("idSearchField")
	set rs=nothing
	
	call closeDb()

	
	session("admin_idcustom")=idcustom
	session("admin_skeyword")=idkeyword
	session("admin_customtype")=1
	session("admin_useExist")=0

	response.redirect "addCFtoCats1.asp"
	response.end
		
Else

	if request("action")="exist" then
		
		idcustom=request.Form("customfield")

		session("admin_idcustom")=idcustom
		session("admin_skeyword")=idkeyword
		session("admin_customtype")=1
		session("admin_useExist")=1
		
		response.redirect "addCFtoCats1.asp"
		response.end
	end if
	
End if
%>

	<% ' START show message, if any %>
        <!--#include file="pcv4_showMessage.asp"-->
    <% 	' END show message

	call opendb()
	query="SELECT idSearchField,pcSearchFieldName,pcSearchFieldShow,pcSearchFieldOrder FROM pcSearchFields ORDER BY pcSearchFieldOrder ASC,pcSearchFieldName ASC;"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	if not rs.eof then
					
		set pcv_tempList = new StringBuilder
		pcv_tempList.append "<select name=""customfield"">" & vbcrlf
					
		pcArray=rs.getRows()
		intCount=ubound(pcArray,2)
		set rs=nothing
					
		For i=0 to intCount
			pcv_tempList.append "<option value=""" & pcArray(0,i) & """>" & replace(pcArray(1,i),"""","&quot;") & "</option>" & vbcrlf
		Next
			
		pcv_tempList.append "</select>" & vbcrlf
		
		pcv_tempList=pcv_tempList.toString
		%>
					
					
		<form name="ajaxSearch" method="post" action="addSFtoCats.asp?action=exist" class="pcForms">
		<input name="nav" type="hidden" value="<%=nav%>">
		<table class="pcCPcontent">
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<th colspan="2">Using an existing custom search field:</th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td width="20%"><div align="right">Text to display:</div></td>
			<td width="80%"><%=pcv_tempList%></td>
		</tr>
		<tr> 
			<td height="10"></td>
			<td> 
			<input type="submit" name="submit0" value="Continue" class="submit2">
			&nbsp;
			<input type="button" name="back" value="Back" onClick="javascript:history.back()">
			</td>
		</tr>
		</table>
		</form>
	<%
	end if
	set rs=nothing
	call closeDb()
	%>

<form name="form1" method="post" action="addSFtoCats.asp?action=new" class="pcForms">
<input name="nav" type="hidden" value="<%=nav%>">
	<table class="pcCPcontent">
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>     
		<tr> 
			<th colspan="2">Adding a new custom search field:</th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>   
        <tr>
            <td width="91" nowrap><div align="right">Field Name:</div></td>
      		<td width="844" nowrap><input name="SF_name" type="text" id="SF_name" size="20" maxlength="150"></td>
        </tr>
        <tr> 
            <td><div align="right">Field Order:</div></td>
            <td><input name="SF_order" type="text" id="SF_order" value="0" size="4" maxlength="150"></td>
        </tr>
        <tr> 
            <td nowrap><div align="right">Display Options:</div></td>
            <td>&nbsp;</td>
        </tr>
        <tr> 
            <td align="right"><input type="checkbox" name="CP_showSEARCH" value="1" class="clearBorder"></td>
            <td>Display on search feature in Control Panel</td>
        </tr>
        <tr> 
            <td align="right"><input type="checkbox" name="SF_showSEARCH" value="1" class="clearBorder"></td>
            <td>Display on Advanced Search in Storefront </td>
        </tr>
        <tr> 
            <td align="right"><input type="checkbox" name="CP_show" value="1" class="clearBorder"></td>
            <td>Display on Add/Modify Product Details in the Control Panel</td>
        </tr>
        <tr> 
            <td align="right"><input type="checkbox" name="SF_show" value="1" class="clearBorder"></td>
            <td>Display on Product Details in Storefront</td>
        </tr>
        <tr>       
            <td>&nbsp;</td>
            <td> 
                <input type="submit" name="submit2" value="Add New" class="submit2" onclick="javascript: if (document.form1.new_fieldname.value=='') {alert('Please enter a value for Field Name'); document.form1.new_fieldname.focus();return(false)} else {return(true)};"></td>
        </tr>
  </table>
</form>
<!--#include file="AdminFooter.asp"-->
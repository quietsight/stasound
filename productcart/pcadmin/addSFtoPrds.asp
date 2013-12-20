<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add a Custom Search Field to Selected Products" %>
<% nav=request("nav")
if nav="bto" then
	Section="services"
else
	Section="products"
end if %>
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
	customname=request.Form("name")
	if customname="" then
		response.redirect "addSFtoPrds.asp?nav="&nav&"&idproduct="&idproduct&"&x="&x&"&message="&server.URLEncode("You must specify a field name.")
	else
		customname=replace(customname,"'","''")
	end if
	keyword=request.Form("keyword")
	if keyword<>"" then
		keyword=replace(keyword,"'","''")
	end if
	tmpFieldShow=request.Form("fieldshow")
	if tmpFieldShow="" then
		tmpFieldShow=0
	end if
	tmpFieldOrder=request.Form("fieldorder")
	if tmpFieldOrder="" then
		tmpFieldOrder=0
	end if
	tmpValueOrder=request.Form("valueorder")
	if tmpValueOrder="" then
		tmpValueOrder=0
	end if
	CP_show=request("CP_show")
	if CP_show="" then
		CP_show=0
	end if
	SF_showSEARCH=request("SF_showSEARCH")
	if SF_showSEARCH="" then
		SF_showSEARCH=0
	end if
	CP_showSEARCH=request("CP_showSEARCH")
	if CP_showSEARCH="" then
		CP_showSEARCH=0
	end if	
		
		call openDb()
	
		query="SELECT idSearchField FROM pcSearchFields WHERE pcSearchFieldName like '"&customname&"';"
		set rs=Server.CreateObject("ADODB.Recordset") 
		set rs=conntemp.execute(query)
		if rs.eof then
			query="INSERT INTO pcSearchFields (pcSearchFieldName,pcSearchFieldShow,pcSearchFieldOrder,pcSearchFieldCPShow,pcSearchFieldSearch,pcSearchFieldCPSearch) VALUES ('"&customname&"'," & tmpFieldShow & "," & tmpFieldOrder & "," & CP_show & "," & SF_showSEARCH & "," & CP_showSEARCH & ");"
			set rs=Server.CreateObject("ADODB.Recordset") 
			set rs=conntemp.execute(query)
		end if
		set rs=nothing
		
		'get idcustom
		query="SELECT idSearchField FROM pcSearchFields WHERE pcSearchFieldName like '"&customname&"';"
		set rs=Server.CreateObject("ADODB.Recordset") 
		set rs=conntemp.execute(query)
		idcustom=rs("idSearchField")
		
		query="SELECT idSearchData FROM pcSearchData WHERE idSearchField=" & idcustom & " AND pcSearchDataName like '" & keyword & "';"
		set rs=connTemp.execute(query)
		if not rs.eof then
			idkeyword=rs("idSearchData")
		else
			if keyword<>"" then
				query="INSERT INTO pcSearchData (idSearchField,pcSearchDataName,pcSearchDataOrder) VALUES (" & idcustom & ",'" & keyword & "'," & tmpValueOrder & ");"
				set rs=connTemp.execute(query)
				query="SELECT idSearchData FROM pcSearchData WHERE idSearchField=" & idcustom & " AND pcSearchDataName like '" & keyword & "';"
				set rs=connTemp.execute(query)
				idkeyword=rs("idSearchData")
			else
				idkeyword=0
			end if
		end if
		
		set rs=nothing
		call closeDb()
		
		if idkeyword=0 then
			response.redirect "ManageCFields.asp"
			response.end
		end if
		
		session("admin_idcustom")=idcustom
		session("admin_skeyword")=idkeyword
		session("admin_customtype")=1
		session("admin_useExist")=0
		
		response.redirect "addCFtoPrds1.asp"
		response.end
		
Else

	if request("action")="exist" then
		idcustom=request.Form("customfield")		
		keyword=request.Form("keyword")
		if not keyword="" then
			call openDb()
			keyword=replace(keyword,"'","''")
			tmpValueOrder=request.Form("valueorder")
			query="INSERT INTO pcSearchData (idSearchField,pcSearchDataName,pcSearchDataOrder) VALUES (" & idcustom & ",'" & keyword & "'," & tmpValueOrder & ");"
			set rs=connTemp.execute(query)
			query="SELECT idSearchData FROM pcSearchData WHERE idSearchField=" & idcustom & " AND pcSearchDataName like '" & keyword & "';"
			set rs=connTemp.execute(query)
			idkeyword=rs("idSearchData")
			set rs=nothing
			call closedb()
		else
			idkeyword=request.Form("SearchValues")
		end if
		
		if idkeyword="" or idkeyword="0" then
			response.redirect "ManageCFields.asp"
			response.end
		end if
		
		session("admin_idcustom")=idcustom
		session("admin_skeyword")=idkeyword
		session("admin_customtype")=1
		session("admin_useExist")=1
		
		response.redirect "addCFtoPrds1.asp"
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
		set pcv_tempFunc = new StringBuilder
		pcv_tempFunc.append "<script>" & vbcrlf
		pcv_tempFunc.append "function CheckList(cvalue) {" & vbcrlf
		pcv_tempFunc.append "if (cvalue==0) {" & vbcrlf
		pcv_tempFunc.append "var SelectA = document.ajaxSearch.SearchValues;" & vbcrlf
		pcv_tempFunc.append "SelectA.options.length = 0; }" & vbcrlf
					
		set pcv_tempList = new StringBuilder
		pcv_tempList.append "<select name=""customfield"" onchange=""javascript:CheckList(document.ajaxSearch.customfield.value);"">" & vbcrlf
					
		pcArray=rs.getRows()
		intCount=ubound(pcArray,2)
		set rs=nothing
					
		For i=0 to intCount
			pcv_tempList.append "<option value=""" & pcArray(0,i) & """>" & replace(pcArray(1,i),"""","&quot;") & "</option>" & vbcrlf
			query="SELECT idSearchData,pcSearchDataName FROM pcSearchData WHERE idSearchField=" & pcArray(0,i) & " ORDER BY pcSearchDataOrder ASC,pcSearchDataName ASC;"
			set rs=connTemp.execute(query)
			if not rs.eof then
				tmpArr=rs.getRows()
				LCount=ubound(tmpArr,2)
				pcv_tempFunc.append "if (cvalue==" & pcArray(0,i) & ") {" & vbcrlf
				pcv_tempFunc.append "var SelectA = document.ajaxSearch.SearchValues;" & vbcrlf
				pcv_tempFunc.append "SelectA.options.length = 0;" & vbcrlf
				For j=0 to LCount
					pcv_tempFunc.append "SelectA.options[" & j & "]=new Option(""" & replace(tmpArr(1,j),"""","\""") & """,""" & tmpArr(0,j) & """);" & vbcrlf
				Next
				pcv_tempFunc.append "}" & vbcrlf
			else
				pcv_tempFunc.append "if (cvalue==" & pcArray(0,i) & ") {" & vbcrlf
				pcv_tempFunc.append "var SelectA = document.ajaxSearch.SearchValues;" & vbcrlf
				pcv_tempFunc.append "SelectA.options.length = 0; }" & vbcrlf
			end if
		Next
			
		pcv_tempList.append "</select>" & vbcrlf
		pcv_tempFunc.append "}" & vbcrlf
		pcv_tempFunc.append "</script>" & vbcrlf
		
		pcv_tempList=pcv_tempList.toString
		pcv_tempFunc=pcv_tempFunc.toString
		%>
					
					
		<form name="ajaxSearch" method="post" action="addSFtoPrds.asp?action=exist" class="pcForms">
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
			<td><div align="right">Existing Value:</div></td>
			<td>
				<select name="SearchValues"></select>
	            <%=pcv_tempFunc%>     
				<script>
					CheckList(document.ajaxSearch.customfield.value);
				</script>
			</td>
		</tr>
		<tr> 
			<td><div align="right">New Value:</div></td>
			<td>
				<input name="keyword" type="text" id="keyword" size="30" maxlength="150">
			</td>
		</tr>
		<tr> 
			<td><div align="right">New Value Order:</div></td>
			<td>
				<input name="valueorder" type="text" id="valueorder" size="4" value="0">
			</td>
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
	<%end if
	set rs=nothing
	call closeDb()%>

<form name="form1" method="post" action="addSFtoPrds.asp?action=new" class="pcForms">
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
			<td width="20%"><div align="right">Field Name:</div></td>
			<td width="80%" valign="top"> 
			<input name="name" type="text" id="name" size="20" maxlength="150">
			</td>
		</tr>
		<tr> 
			<td><div align="right">Field Order:</div></td>
			<td><input name="fieldorder" type="text" id="fieldorder" size="4" maxlength="150" value="0"></td>
		</tr>
        <tr> 
            <td nowrap><div align="right">Display Options:</div></td>
            <td>&nbsp;</td>
        </tr>
        <tr> 
            <td align="right"><input type="checkbox" name="CP_showSEARCH" value="1" class="clearBorder">					    </td>
            <td>Display on search feature in Control Panel</td>
        </tr>
        <tr> 
            <td align="right"><input type="checkbox" name="SF_showSEARCH" value="1" class="clearBorder"></td>
            <td>Display on Advanced Search in Storefront </td>
        </tr>
        <tr> 
            <td align="right"><input type="checkbox" name="CP_show" value="1" class="clearBorder">					    </td>
            <td>Display on Add/Modify Product Details in the Control Panel</td>
        </tr>
        <tr> 
            <td align="right"><input type="checkbox" name="fieldshow" value="1" class="clearBorder"></td>
            <td>Display on Product Details in Storefront </td>
        </tr>

		<tr> 
			<td><div align="right">Value:</div></td>
			<td><input name="keyword" type="text" id="keyword" size="30" maxlength="150"></td>
		</tr>
		<tr> 
			<td><div align="right">Value Order:</div></td>
			<td><input name="valueorder" type="text" id="valueorder" size="4" maxlength="150" value="0"></td>
		</tr>
		<tr>       
			<td>&nbsp;</td>
			<td> 
			<input type="submit" name="submit" value="Continue" class="submit2">
			&nbsp;
			<input type="button" name="back" value="Back" onClick="javascript:history.back()">
			</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->
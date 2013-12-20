<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Option Group" %>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp" -->
<%
 
Dim rs, rstemp, strSQL, pid, query, conntemp, AssignID, pidProduct, pidOptionGroup

'START code used when loading the page after adding a new attribute
If request.form("Submit")<>"" then
	'add new attribute
	pidOptionGroup=request.form("idOptionGroup")
	AssignID=request.Form("AssignID")
	pidProduct=request.Form("idProduct")
	refpage=request.Form("refpage")
	attribute=replace(trim(request.form("attrib")),"'","''")
	if attribute="" then
		response.redirect "modOptGrpa.asp?msg="&Server.Urlencode("You need to specify an attribute to add.")&"&idOptionGroup="&pidOptionGroup
		response.end
	end if
	
	call OpenDb()
	Dim pcv_strResults	
	pcv_strResults=""
		query="INSERT INTO options (optionDescrip) VALUES ('"&attribute&"')"
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=connTemp.execute(query)
		query="SELECT idOption FROM options WHERE optionDescrip='"&attribute&"' ORDER BY idOption Desc;"
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=connTemp.execute(query)
		pcv_strResults="Successfully created new Attribute. "
	pidOption=rs("idOption")
	
	'CHECK IF THIS OPTION ALREADY IS ASSIGNED TO THE GROUP
	query="SELECT idOption FROM optGrps WHERE idOptionGroup="&pidOptionGroup&" AND idOption="&pidOption&";"
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=connTemp.execute(query)
	if rs.eof then
		query="INSERT INTO optGrps (idOption, idOptionGroup) VALUES ("&pidOption&", "&pidOptionGroup&")"
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=connTemp.execute(query)
		pcv_strResults = pcv_strResults & "Successfully added Attribute to Option Group."
	else
		pcv_strResults = pcv_strResults & "Attribute already exists in Option Group."
	End if
	
	'if the admin had originally come from a product options page (modPrdOpta2.asp or modPrdOpta3.asp), go back to that page, otherwise stay on the same page	
	if pidProduct = 0 then
		set rs=nothing
		call closeDb()
	 	response.redirect "modOptGrpa.asp?s=1&msg="&Server.Urlencode(pcv_strResults)&"&idOptionGroup="&pidOptionGroup
	else
	 if refpage = "modPrdOpta3" then
		set rs=nothing
		call closeDb()
	  	response.redirect "modPrdOpta3.asp?AssignID="&AssignID&"&idProduct="&pidProduct&"&idOptionGroup="&pidOptionGroup
	  else
		set rs=nothing
		call closeDb()
	  	response.redirect "modPrdOpta2.asp?AssignID="&AssignID&"&idProduct="&pidProduct&"&idOptionGroup="&pidOptionGroup
	  end if
	response.end
	end if
End if
'END code used when loading the page after adding a new attribute

'if the request is coming from modPrdOpta2.asp (specific product), get that information so that the admin can return to that page
pidOptionGroup=request.Querystring("idOptionGroup")
AssignID=request.QueryString("AssignID")
 if AssignID = "" then
 	AssignID = 0
 end if
pidProduct=request.QueryString("idProduct")
 if pidProduct = "" then
 	pidProduct = 0
 end if
refpage=request.QueryString("page")

if trim(pidOptionGroup)="" then
   response.redirect "msg.asp?message=22"
end if

	
Function GetAttributes()

	call openDb()

	' gets group assignments
	query="SELECT options.optionDescrip, options.idOption FROM options INNER JOIN optGrps ON options.idOption=optGrps.idoption WHERE (((optGrps.idOptionGroup)="&pidOptionGroup&")) ORDER BY options.optionDescrip;"
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=connTemp.execute(query)
	
	If Err Then
		CleanUp
		GetAttributes=False
		Exit Function
	ElseIf rs.EOF OR rs.BOF Then
		CleanUp
		GetAttributes=False
		Exit Function
	Else
		GetAttributes=True
	End If
End Function
Sub CleanUp
	Set rs=Nothing
	Set connTemp=Nothing
	call closeDb()
End Sub

call openDb()
query="SELECT * FROM optionsGroups WHERE idOptionGroup=" &pidOptionGroup
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(query)

if err.number <> 0 then
	set rstemp=nothing
	call closeDb()
 	response.redirect "techErr.asp?error="& Server.Urlencode("Error loading option group information on modOptGrpa.asp") 
end If

%>
<!--#include file="AdminHeader.asp"-->
<form method="post" name="modOpGr" action="modOptGrpb.asp" class="pcForms">
<input type="hidden" name="idOptionGroup" size="60" value="<%=pidOptionGroup%>">
<table class="pcCPcontent">
    <tr>
        <td colspan="2" class="pcCPspacer">
            <% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
        </td>
    </tr>
	<tr>                       
		<th colspan="2">Rename Option Group</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>              
	<tr> 
		<td width="20%">Option Group:</td>
		<td width="80%"><input type="text" name="optionGroupDesc" value="<%=rstemp("optionGroupDesc")%>" size="35"></td>
	</tr>
	<tr> 
		<td colspan="2">     
		<input type="submit" name="modify" value="Rename" class="submit2">&nbsp;
		<input type="button" name="Button" value="Back" onClick="document.location.href='ManageOptions.asp';">
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>          
</table>
</form>

<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<th colspan="2">Manage Attributes in this Group</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>               
	<%
		If NOT GetAttributes() Then
			noattrb=1
	%>                      
	<tr>
		<td colspan="2">No attributes found</td>
	</tr> 
	<tr>
		<td colspan="2">
        <div style="float: right;">
        <form name="form1" action="modOptGrpa.asp" method="post" class="pcForms">                
			Add New Attribute:&nbsp;<input type="text" name="attrib">
			<input type="hidden" name="idOptionGroup" value="<%=pidOptionGroup%>">
			<input type="hidden" name="AssignID" value="<%=AssignID%>">
			<input type="hidden" name="idProduct" value="<%=pidProduct%>">
			<input type="hidden" name="refpage" value="<%=refpage%>">
			&nbsp;<input type="submit" name="Submit" value="Add" class="submit2">
		</form>
        </div>
        <h2>Existing attributes:</h2>
        </td>
	</tr> 
	<%
		Else
	%>
	<tr>
		<td colspan="2">
        <div style="float: right;">
        <form name="form1" action="modOptGrpa.asp" method="post" class="pcForms">                
			Add New Attribute:&nbsp;<input type="text" name="attrib">
			<input type="hidden" name="idOptionGroup" value="<%=pidOptionGroup%>">
			<input type="hidden" name="AssignID" value="<%=AssignID%>">
			<input type="hidden" name="idProduct" value="<%=pidProduct%>">
			<input type="hidden" name="refpage" value="<%=refpage%>">
			&nbsp;<input type="submit" name="Submit" value="Add" class="submit2">
		</form>
        </div>
        <h2>Existing attributes:</h2>
        </td>
	</tr> 
	<%
	noattrb=0
	do while not rs.eof
	%>
                      
	<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
        <td width="50%"><%= rs("optionDescrip") %></td>
        <td width="50%" nowrap align="right" class="cpLinksList"> 
            <a href="modOpta.asp?idOption=<%=rs("idOption")%>&idOptionGroup=<%=pidOptionGroup%>">Rename</a> | <a href="javascript:if (confirm('This attribute may have been assigned to one or more products: are you sure you want to delete it?')) location='actionOptions.asp?delete=<%=rs("idOption")%>&idOptionGroup=<%=pidOptionGroup%>'">Delete</a>
        </td>
	</tr>
                      
	<%
    rs.MoveNext
    Loop
    CleanUp
    End If
    %>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2" align="center">
		<form class="pcForms">
		<input type="button" value="Manage Options" onClick="location.href='ManageOptions.asp'">
		<% If noattrb=0 Then %>
        &nbsp;<input type="button" value="Add to Multiple Products" onClick="location.href='AssignMultiOptions.asp?idOptionGroup=<%=pidOptionGroup%>'">
        &nbsp;<input type="button" value="Remove from Multiple Products" onClick="location.href='RevMultiOptions.asp?idOptionGroup=<%=pidOptionGroup%>'">
		<% end if %>
		</form>
		</td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% section = "products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<%
if ucase(right(session("importfile"),4))=".XLS" then
response.redirect "catstep2-xls.asp?append=" & request("append") & "&movecat=" & request("movecat")
end if%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="catcheckfields.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/productcartFolder.asp"-->

<%
append=request("append")
if append<>"" then
session("append")=append
else
append=session("append")
end if
movecat=request("movecat")
if movecat<>"" then
else
movecat="1"
end if
session("movecat")=movecat
if append="1" then
requiredfields = 2
else
requiredfields = 1
end if

if session("append")="1" then 
	pageTitle = "UPDATE"
else
	pageTitle = "IMPORT" 
end if 
pageTitle = pageTitle & " CATEGORY DATA WIZARD - Map fields"
%>

<!--#include file="AdminHeader.asp"-->

<%
sub displayerror(msg)
%>
<!--#include file="pcv4_showMessage.asp"-->
<%
end sub%>

	<table class="pcCPcontent">
	<tr>
        <td colspan="2"><h2>Steps:</h2></td>
    </tr>
    <tr>
        <td width="5%" align="right"><img border="0" src="images/step1.gif"></td>
        <td width="95%"><font color="#A8A8A8">Select category data file</font></td>
    </tr>
    <tr>
        <td align="right"><img border="0" src="images/step2a.gif"></td>
        <td><strong>Map fields</strong></td>
    </tr>
    <tr>
        <td align="right"><img border="0" src="images/step3.gif"></td>
        <td><font color="#A8A8A8">Confirm mapping</font></td>
    </tr>
    <tr>
        <td align="right"><img border="0" src="images/step4.gif"></td>
        <td><font color="#A8A8A8"><%if session("append")="1" then%>Update<%else%>Import<%end if%> results</font></td>
    </tr>
    </table>
    
		<br>
	<!--#include file="../includes/ppdstatus.inc"-->
	<%
	FileCSV = "../pc/catalog/" & session("importfile")
	if PPD="1" then
		FileCSV="/"&scPcFolder&"/pc/catalog/"& session("importfile")
	end if
	findit = Server.MapPath(FileCSV)
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	Err.number=0
	Set f = fso.OpenTextFile(findit, 1)
	if Err.number>0 then
		session("importfilename")=""%>
		<script>
		location="msg.asp?message=38";
		</script><%
	end if
	Topline = f.Readline
	A=split(Topline,",")
	if ubound(a)-lbound(a)+1<requiredfields then
	 session("importfilename")=""%>
	 <script>
	 location="msg.asp?message=35";
	 </script><%
	end if
	validfields=0
	for i=lbound(a) to ubound(a)
		if trim(a(i))<>"" then
			validfields=validfields+1
		end if
	next
	if validfields<requiredfields then
		session("importfilename")=""%>
		<script>
			location="msg.asp?message=35";
		</script>
	<%end if
	session("totalfields")=ubound(a)-lbound(a)+1
	if a(ubound(a))="" then
	session("totalfields")=session("totalfields")-1
	end if
	f.Close
	Set fso = nothing
	Set f = nothing
	msg=request.querystring("msg")
	if msg<>"" then 
		displayerror(msg)
	end if 
	%>
    <form method="post" action="catstep3.asp" class="pcForms">
    <table class="pcCPcontent">
    	<tr>
        	<td colspan="2">
	Use the drop-down menus below to map existing fields in your product database, located on the left side of the page under 'From' to ProductCart database fields, which are located on the right side of the page under 'To'.
	</td>
</tr>
		<tr>
			<td colspan="2" class="pcCPSpacer"></td>
		</tr>
<tr>
	<th width="40%">From:</th>
	<th width="60%">To:</th>
</tr>
	<tr>
			<td colspan="2" class="pcCPSpacer"></td>
	</tr>
	<% validfields=0
	for i=lbound(a) to ubound(a)
	FiName=a(i)
    if trim(FiName)<>"" then
    	if left(FiName,1)=chr(34) then
    	FiName=mid(FiName,2,len(FiName))
    	end if
    	if right(FiName,1)=chr(34) then
    	FiName=mid(FiName,1,len(FiName)-1)
    	end if    	
      	validfields=validfields+1%>
		<tr>
			<td width="40%"><%=FiName%>
				<input type=hidden name="F<%=validfields%>" value="<%=FiName%>">
				<input type=hidden name="P<%=validfields%>" value="<%=i%>">
			</td>
			<td width="60%">
				<select size="1" name="T<%=validfields%>">
					<option value="   ">   </option>
					<option value="Category Name">Category Name</option>
					<option value="Category ID">Category ID</option>
					<option value="Small Image">Small Image</option>
					<option value="Large Image">Large Image</option>
					<option value="Parent Category Name">Parent Category Name</option>
					<option value="Parent Category ID">Parent Category ID</option>
					<option value="Category Short Description">Category Short Description</option>
					<option value="Category Long Description">Category Long Description</option>
					<option value="Hide Category Description">Hide Category Description</option>
					<option value="Display Sub-Categories">Display Sub-Categories</option>
					<option value="Sub-Categories per Row">Sub-Categories per Row</option>
					<option value="Sub-Category Rows per Page">Sub-Category Rows per Page</option>
					<option value="Display Products">Display Products</option>
					<option value="Products per Row">Products per Row</option>
					<option value="Product Rows per Page">Product Rows per Page</option>
					<option value="Hide category">Hide category</option>
					<option value="Hide category from retail customers">Hide category from retail customers</option>
					<option value="Product Details Page Display Option">Product Details Page Display Option</option>
					<option value="Category Meta Tags - Title">Category Meta Tags - Title</option>
					<option value="Category Meta Tags - Description">Category Meta Tags - Description</option>
					<option value="Category Meta Tags - Keywords">Category Meta Tags - Keywords</option>
					<option value="Featured Sub-Category Name">Featured Sub-Category Name</option>
					<option value="Featured Sub-Category ID">Featured Sub-Category ID</option>
					<option value="Use Featured Sub-Category Image">Use Featured Sub-Category Image</option>
					<option value="Category Order">Category Order</option>
                           
					<%if request("T" & validfields)<>"" then%>
						<option value="<%=request("T" & validfields)%>" selected><%=request("T" & validfields)%></option>
					<%else
						FiName1=""
						FiName1=CheckField(FiName)
						if FiName1<>"" then%>
							<option value="<%=FiName1%>" selected><%=FiName1%></option>
						<%end if
					end if%>
				</select>
			</td>
		</tr>
    <%end if
Next%>  
		<tr>
			<td colspan="2" class="pcCPSpacer"><hr></td>
		</tr>                 
		<tr>
			<td colspan="2">
		<input type="hidden" name="validfields" value="<%=validfields%>">         
		<input type="submit" name="submit" value="Map Fields" class="submit2">
        &nbsp; <input type="reset" name="reset" value="Reset">
			</td>
		</tr>
		</table>
        </form>
<!--#include file="AdminFooter.asp"-->
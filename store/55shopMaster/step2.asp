<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% section = "products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<%
if ucase(right(session("importfile"),4))=".XLS" then
response.redirect "step2-xls.asp?append=" & request("append") & "&movecat=" & request("movecat")
end if%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="checkfields.asp"-->
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
requiredfields = 1
else
requiredfields = 4
end if

if session("append")="1" then 
	pageTitle = "UPDATE"
else
	pageTitle = "IMPORT" 
end if 
pageTitle = pageTitle & " PRODUCT DATA WIZARD - Map fields"
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
        <td width="95%"><font color="#A8A8A8">Select product data file</font></td>
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
		session("importfilename")=""
		response.redirect "msg.asp?message=31"
	end if
	Topline = f.Readline
	A=split(Topline,",")
	if ubound(a)-lbound(a)+1<requiredfields then
		session("importfilename")=""
		response.Redirect "msg.asp?message=28"
	end if
	validfields=0
	for i=lbound(a) to ubound(a)
		if trim(a(i))<>"" then
			validfields=validfields+1
		end if
	next
	if validfields<requiredfields then
		session("importfilename")=""
		response.Redirect "msg.asp?message=28"
	end if
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
    <form method="post" action="step3.asp" class="pcForms">
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
For i=lbound(a) to ubound(a)
if trim(a(i))<>"" then
	if left(a(i),1)=chr(34) then
		a(i)=mid(a(i),2,len(a(i)))
	end if
	if right(a(i),1)=chr(34) then
		a(i)=mid(a(i),1,len(a(i))-1)
   	end if    	
	validfields=validfields+1%>
<tr>
	<td width="40%"><%=a(i)%>
		<input type=hidden name="F<%=validfields%>" value="<%=a(i)%>" >
		<input type=hidden name="P<%=validfields%>" value="<%=i%>" >
	</td>
	<td width="60%">
	<select size="1" name="T<%=validfields%>">
		<option value="   ">   </option>
		<option value="SKU">SKU</option>
		<option value="Name">Name</option>
		<option value="Description">Description</option>
		<option value="Short Description">Short Description</option>
		<option value="Product Type">Product Type</option>
		<option value="Online Price">Online Price</option>
		<option value="List Price">List Price</option>
		<option value="Wholesale Price">Wholesale Price</option>
		<option value="Weight">Weight</option>
		<option value="Units to make 1 lb">Units to make 1 lb</option>
		<option value="Stock">Stock</option>
		<option value="Category Name">Category Name</option>
		<option value="Short Category Description">Short Category Description</option>
		<option value="Long Category Description">Long Category Description</option>
		<option value="Category Small Image">Category Small Image</option>
		<option value="Category Large Image">Category Large Image</option>
		<option value="Parent Category">Parent Category</option>
		<option value="Additional Category 1">Additional Category 1</option>
		
		<option value="Short Category Description 1">Short Category Description 1</option>
		<option value="Long Category Description 1">Long Category Description 1</option>
		<option value="Category Small Image 1">Category Small Image 1</option>
		<option value="Category Large Image 1">Category Large Image 1</option>
		
		<option value="Parent Category 1">Parent Category 1</option>
		<option value="Additional Category 2">Additional Category 2</option>
		
		<option value="Short Category Description 2">Short Category Description 2</option>
		<option value="Long Category Description 2">Long Category Description 2</option>
		<option value="Category Small Image 2">Category Small Image 2</option>
		<option value="Category Large Image 2">Category Large Image 2</option>
		
		<option value="Parent Category 2">Parent Category 2</option>
		<option value="Brand Name">Brand Name</option>
		<option value="Brand Logo">Brand Logo</option>
		<option value="Thumbnail Image">Thumbnail Image</option>
		<option value="General Image">General Image</option>
		<option value="Detail view Image">Detail view Image</option>
		<option value="Active">Active</option>
		<option value="Show savings">Show savings</option>
		<option value="Special">Special</option>
		<option value="Featured">Featured</option>
		
		<option value="Product Notes">Product Notes</option>
		<option value="Enable Image Magnifier">Enable Image Magnifier</option>
		<option value="Page Layout">Page Layout</option>
		<option value="Hide SKU on the product details page">Hide SKU on the product details page</option>
		
		<option value="Option 1">Option 1</option>
		<option value="Attributes 1">Attributes 1</option>
		<option value="Option 1 Required">Option 1 Required</option>
		<option value="Option 1 Order">Option 1 Order</option>
		<option value="Option 2">Option 2</option>
		<option value="Attributes 2">Attributes 2</option>
		<option value="Option 2 Required">Option 2 Required</option>
		<option value="Option 2 Order">Option 2 Order</option>
		<option value="Option 3">Option 3</option>
		<option value="Attributes 3">Attributes 3</option>
		<option value="Option 3 Required">Option 3 Required</option>
		<option value="Option 3 Order">Option 3 Order</option>
		<option value="Option 4">Option 4</option>
		<option value="Attributes 4">Attributes 4</option>
		<option value="Option 4 Required">Option 4 Required</option>
		<option value="Option 4 Order">Option 4 Order</option>
		<option value="Option 5">Option 5</option>
		<option value="Attributes 5">Attributes 5</option>
		<option value="Option 5 Required">Option 5 Required</option>
		<option value="Option 5 Order">Option 5 Order</option>
		<option value="Reward Points">Reward Points</option>
		<option value="Non-taxable">Non-taxable</option>
		<option value="No shipping charge">No shipping charge</option>
		<option value="Not for sale">Not for sale</option>
		<option value="Not for sale copy">Not for sale copy</option>
		<option value="Disregard stock">Disregard stock</option>
		<option value="Display No Shipping Text">Display No Shipping Text</option>
		<option value="Minimum Quantity customers can buy">Minimum Quantity customers can buy</option>
		<option value="Force purchase of multiples of minimum">Force purchase of multiples of minimum</option>
		<option value="Oversized Product Details">Oversized Product Details</option>
		<option value="First Unit Surcharge">First Unit Surcharge</option>
		<option value="Additional Unit(s) Surcharge">Additional Unit(s) Surcharge</option>
		
		
		<option value="Mega Tags - Title">Mega Tags - Title</option>
		<option value="Mega Tags - Description">Mega Tags - Description</option>
		<option value="Mega Tags - Keywords">Mega Tags - Keywords</option>
				
		<%if scBTO=1 then%>
		<option value="Hide BTO Price">Hide BTO Price</option>
		<option value="Hide Default Configuration">Hide Default Configuration</option>
		<option value="Disallow purchasing">Disallow purchasing</option>
		<option value="Skip Product Details Page">Skip Product Details Page</option>
		<%end if%>
		
		<%'Start SDBA%>
		<option value="Product Cost">Product Cost</option>
		<option value="Back Order">Back-Order</option>
		<option value="Ship within N Days">Ship within N Days</option>
		<option value="Low inventory notification">Low inventory notification</option>
		<option value="Reorder Level">Reorder Level</option>
		<option value="Is Drop-shipped">Is Drop-shipped</option>
		<option value="Supplier ID">Supplier ID</option>
		<option value="Drop-Shipper ID">Drop-Shipper ID</option>
		<%'End SDBA%>
		
		<option value="Custom Search Field (1)">Custom Search Field (1)</option>
		<option value="Custom Search Field (2)">Custom Search Field (2)</option>
		<option value="Custom Search Field (3)">Custom Search Field (3)</option>
		<option value="Downloadable Product">Downloadable Product</option>
		<option value="Downloadable File Location">Downloadable File Location</option>
		<option value="Make Download URL expire">Make Download URL expire</option>
		<option value="URL Expiration in Days">URL Expiration in Days</option>
		<option value="Use License Generator">Use License Generator</option>
		<option value="Local Generator">Local Generator</option>
		<option value="Remote Generator">Remote Generator</option>
		<option value="License Field Label (1)">License Field Label (1)</option>
		<option value="License Field Label (2)">License Field Label (2)</option>
		<option value="License Field Label (3)">License Field Label (3)</option>
		<option value="License Field Label (4)">License Field Label (4)</option>
		<option value="License Field Label (5)">License Field Label (5)</option>
		<option value="Additional copy">Additional copy</option>
		<option value="Gift Certificate">Gift Certificate</option>
		<option value="Gift Certificate Expiration">Gift Certificate Expiration</option>
		<option value="Electronic Only (Gift Certificate)">Electronic Only (Gift Certificate)</option>
		<option value="Use Generator (Gift Certificate)">Use Generator (Gift Certificate)</option>
		<option value="Expiration Date (Gift Certificate)">Expiration Date (Gift Certificate)</option>
		<option value="Expire N days (Gift Certificate)">Expire N days (Gift Certificate)</option>
		<option value="Custom Generator Filename (Gift Certificate)">Custom Generator Filename (Gift Certificate)</option>
		
		<option value="Google Product Category">Google Product Category</option>
		<option value="Google Shopping - Gender">Google Shopping - Gender</option>
		<option value="Google Shopping - Age">Google Shopping - Age</option>
		<option value="Google Shopping - Color">Google Shopping - Color</option>
		<option value="Google Shopping - Size">Google Shopping - Size</option>
		<option value="Google Shopping - Pattern">Google Shopping - Pattern</option>
		<option value="Google Shopping - Material">Google Shopping - Material</option>
		
	<%if request("T" & validfields)<>"" then%>
		<option value="<%=request("T" & validfields)%>" selected><%=request("T" & validfields)%></option>
    <%else
		FiName=""
		FiName=CheckField(a(i))
		if FiName<>"" then%>
		<option value="<%=FiName%>" selected><%=FiName%></option>
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
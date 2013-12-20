<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true
pageTitle="Control Panel - Message" %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<div style="margin: 30px;">
<div class="pcCPmessage">
	<% varmsg=request.querystring ("message")
    varmsg=replace(varmsg,"<","&lt;")
    varmsg=replace(varmsg,">","&gt;")
	select case cint(varmsg)
		case 1
			response.write "Your User ID and/or Password are incorrect. <a href=""login_1.asp"">Try again</a>."
		case 2
			response.write "You must specify the item ID."
		case 3
			response.write "You must specify a category."
		case 4
			response.write "Your search did not return any orders."
		case 5
			response.write "No products found."
		case 6
			response.write "You need to add products to the store database before adding Option Groups."
		case 7
			response.write "You must specify the affiliate ID."
		case 8
			response.write "This Option Group has been assigned to one or more products and therefore it cannot be deleted."
		case 9
			response.write "This attribute is currently assigned to one or more products in your store, and therefore it was not deleted.<br><br><A href=AdminOptions.asp>Delete anyway</a>&nbsp;&nbsp;<A href=AdminOptions.asp>Back</a>"
		case 10
			response.write "You must first delete the products."
		case 11
			response.write "Please specify required data."
		case 12
			response.write "You must fill out all required fields."
		case 13 
			response.write "Enter valid prices and/or numbers."
		case 14
			response.write "Enter a valid number."
		case 15
			response.write "You must select either a State or Province. You cannot select both."
		case 16
			response.write "You must enter a description."
		case 17
			response.write "Your search did not return any results."
		case 18
			response.write "The category that you selected is the <u>parent category</u> of another category, and therefore it <u>cannot be deleted</u>.<br><br>In order to delete this category you need to either move the subcategories to another parent category, or delete the subcategories.<br><br><a href=""manageCategories.asp?prdType=1"">Manage categories</a>."
		case 19
			response.write "The selected category has been deleted.<br><br><a href=""manageCategories.asp?prdType=1"">Back to manage categories.</a>"
		case 20
			response.write "You must select either a State or s Province. You cannot select both."
		case 21
			response.write "You must specify the option ID."
		case 22
			response.write "You must specify the Option Group ID."
		case 23
			response.write "Currently there are no Option Groups.<br><br>Before you can add Options to your products, you must first create at least one Option Group.<br><br><li><a href=manageOptions.asp>Create Option Groups</a></li>"
		case 24
			response.write "The store database does not contain enough data to generate reports."
		case 25
			response.write "You must enter at least one keyword in the search box."
		case 26
			response.write "Your search did not return any customers. <a href=""Javascript:history.go(-1);"">New search</a>."
		case 27
			response.write "These are no option groups in the store database."
		case 28
		   response.write "The CSV data file that you have uploaded does not contain enough fields. Please consult the ProductCart User Guide for more information on the requirements for the import data file, then upload a new file through the Import Wizard."	
		case 29
		   response.write "The XLS data file that you have uploaded does not contain enough fields. Please consult the ProductCart User Guide for more information on the requirements for the import data file, then upload a new file through the Import Wizard."	
		case 30
			response.write "The XLS data file that you have uploaded is not properly formatted. Please consult the ProductCart User Guide for more information on the requirements for importing an Excel spreadsheet into ProductCart. For instance, you may have not given the name 'IMPORT' to the data range. Please edit the file and upload it again through the Import Wizard."   
		case 31
			response.write "The CSV data file that you have uploaded is not properly formatted. Please consult the ProductCart User Guide for more information on the requirements for importing a Comma Separated Value file into ProductCart. Update the file and then upload it through the Import Wizard."
		case 32
			response.write "It appears that there may be a problem with the way you have defined the import area in your Excel worksheet. Please consult the ProductCart User Guide for detailed information on how to properly format an Excel worksheet for import into your store database. Once you have reformatted the worksheet, please start the Import Wizard again and reupload the file."
		case 33
			response.write "Please add at least one product to one of the categories in your store.<br><br><a href=""menu.asp"">Back</a>"
		case 34
			response.write "The product you selected is inactive."
		'START - Category Import Wizard Messages
		case 35
		   response.write "The CSV data file that you have uploaded does not contain enough fields. Please consult the ProductCart User Guide for more information on the requirements for the import data file, then upload a new file.<br><br><a href='catindex_import.asp'>Restart Import Wizard</a>"	
		case 36
		   response.write "The XLS data file that you have uploaded does not contain enough fields. Please consult the ProductCart User Guide for more information on the requirements for the import data file, then upload a new file.<br><br><a href='catindex_import.asp'>Restart Import Wizard</a>"	
		case 37
			response.write "The XLS data file that you have uploaded is not properly formatted. Please consult the ProductCart User Guide for more information on the requirements for importing an Excel spreadsheet into ProductCart. For instance, you may have not given the name 'IMPORT' to the data range. Please edit the file and upload it again.<br><br><a href='catindex_import.asp'>Restart Import Wizard</a>"   
		case 38
			response.write "The CSV data file that you have uploaded is not properly formatted. Please consult the ProductCart User Guide for more information on the requirements for importing a Comma Separated Value file into ProductCart.<br><br><a href='catindex_import.asp'>Restart Import Wizard</a>"
		case 39
			response.write "It appears that there may be a problem with the way you have defined the import area in your Excel worksheet. Please consult the ProductCart User Guide for detailed information on how to properly format an Excel worksheet for import into your store database. Once you have reformatted the worksheet, please start the Import Wizard again and reupload the file.<br><br><a href='catindex_import.asp'>Restart Import Wizard</a>"
		'END - Category Import Wizard Messages
		' Start Customer Import Wizard
		case 40
			response.write "The CSV data file that you have uploaded is not properly formatted. Please consult the ProductCart User Guide for more information on the requirements for importing a Comma Separated Value file into ProductCart.<br><br><a href='custindex_import_help.asp'>Restart Import Wizard</a>"
		case 41
		   response.write "The CSV data file that you have uploaded does not contain enough fields. Please consult the ProductCart User Guide for more information on the requirements for the import data file, then upload a new file.<br><br><a href='custindex_import_help.asp'>Restart Import Wizard</a>"	
		case 42
		   response.write "You cannot clone a product that has not been assigned to any categories. First assign the product to at least one category, then clone it. <a href=""javascript:history.back()"">Back</a>."	
		case 43
			response.write "No addresses found. <a href='exportData.asp#ups'>Back</a>."
		case 44
			response.write "Google Analytics is not active. Before you can use this feature, activate Google Analytics by entering your Google Analytics Web site Profile ID under '<a href=AdminSettings.asp?tab=5>Store Settings > Miscellaneous</a>'."
		case 45
			response.write "Not a valid Order ID."
		
		' End Customer Import Wizard
		
	end select

	if request.QueryString("bk")=1 AND isNumeric(request.QueryString("bk")) then
		response.write "<br><br><input type=""button"" name=""back"" value=""Back"" onClick=""javascript:history.back()"" class=""ibtnGrey"">"
	end if %>
</div>
</div>
<!--#include file="AdminFooter.asp"-->
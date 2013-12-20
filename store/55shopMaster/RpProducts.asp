<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% pageTitle = "Reward Points - Assign Points to Multiple Products" %>
<% Section = "specials" %>
<%
If request.form("Submit")<>"" then
	ieqtype=request.form("eqtype")
	if ieqtype="" then
		ieqtype="1"
	end if
	
	if ieqtype="2" then
		imultiplier=request.form("multiplier")
		
		if scDecSign="," then
			imultiplier=replace(imultiplier,".","")
			imultiplier=replace(imultiplier,",",".")
		else
			imultiplier=replace(imultiplier,",","")
		end if

		if imultiplier="" then
			response.redirect "RpProducts.asp?msg="&server.URLEncode("You must insert a multiplier into the form to use the multiplier option.")
		end if
	end if

	'// Get Category List
	pcv_IdCategory=request("idcategory")	
	if pcv_IdCategory="" then
		pcv_IdCategory=0
	end if
	pcv_IdCategory=trim(pcv_IdCategory)	
	pcList1=split(pcv_IdCategory,",")	

	
	dim query, conntemp, rs, pcArray, pcv_IntNumrows, pcv_intRowCnt, i
	call openDb()
	
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start:  Do For Each Category
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		For lk=lbound(pcList1) to ubound(pcList1)
		
			'// Filter By Category
			If trim(pcList1(lk))<>"0" then		
				query1=" AND categories_products.idcategory=" & trim(pcList1(lk)) & " "		
			Else		
				query1=""		
			End if
			
			'// Select Products
			query="SELECT categories_products.idcategory, products.idProduct FROM categories_products,products WHERE products.removed=0 AND  products.idproduct=categories_products.idproduct " & query1 & ";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)		
			Count1=-1		
			if not rs.eof then
				pcArray=rs.getRows()
				Count1=ubound(pcArray,2)
			end if		
			set rs=nothing
			
			Count2=0
			
			if Count1>-1 then
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// Start:  Do For Each Product In Category
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				For k=0 to Count1		
					intIdProduct=pcArray(1,k)
					
					if ieqtype="1" then
							query="UPDATE products SET products.iRewardPoints=Round(products.price,0) WHERE idproduct=" & intIdProduct
							set rs = Server.CreateObject("ADODB.Recordset")     
							set rs=conntemp.execute(query)
							set rs=nothing
							if err.number <> 0 then
									response.write "Error: "&Err.Description
									err.number=0
									err.description=""
							end if
					else
							query="UPDATE products SET products.iRewardPoints=Round(products.price*" & imultiplier & ",0)  WHERE idproduct=" & intIdProduct
							set rs = Server.CreateObject("ADODB.Recordset")     
							set rs=conntemp.execute(query)
							set rs=nothing
							if err.number <> 0 then
									response.write "Error: "&Err.Description
									err.number=0
									err.description=""
							end if
					end if	
					
				Next
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'// End:  Do For Each Product In Category
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			end if '// if Count1>-1 then
			
			If trim(pcList1(lk))="0" then
				exit for
			End if
		
		Next
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End:  Do For Each Category
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	

	call closeDb()
	response.redirect "RpProducts.asp?mode=3"
End If %>
<!--#include file="Adminheader.asp"-->
<% if request.QueryString("mode")="3" then %>
	<table class="pcCPcontent">
		<tr> 
			<td align="center"><div class="pcCPmessageSuccess">You have successfully assigned points to the selected products. <a href="rpProducts.asp">Back</a>.</div></td>
		</tr>
	</table>
<% else %>
	<script language="JavaScript">
	<!--
		
	function Form1_Validator(theForm)
	{
	
	if (theForm.buttoncheck.value == "0")
			{
				alert("Please select an option before submitting!");
					theForm.multiplier.focus();
					return (false);
		}
		
	return (true);
	}
	//-->
	</script>

	<form name="form1" method="post" action="RpProducts.asp" onSubmit="return Form1_Validator(this)" class="pcForms">
		<table class="pcCPcontent">     
			<tr>       
				<td>Use the settings below to assign Reward Points to multiple products. Note that this action cannot be undone, but you can still edit the number of points on a product by product basis (or import that information using the Import Wizard).</td>
			</tr>
			<tr> 
				<td class="pcCPspacer"></td>
			</tr>
            <tr>
                <th>Filter products by category</th>
            </tr>
			<tr> 
				<td class="pcCPspacer"></td>
			</tr>
            <tr>
            <td>
                Please <strong>select one or more categories</strong>: all products assigned to those categories will be affected, regardless of whether they are also assigned to other categories that you do not select. Press down the CTRL key on your keyboard to select multiple categories.<br>
                <br>
                <%
                cat_DropDownName="idcategory"
                cat_Type="1"
                cat_DropDownSize="10"
                cat_MultiSelect="1"
                cat_ExcBTOHide="1"
                cat_StoreFront="0"
                cat_ShowParent="1"
                cat_DefaultItem="All categories"
                cat_SelectedItems="0,"
                cat_ExcItems=""
                cat_ExcSubs="0"
                cat_ExcBTOItems="1"
                cat_EventAction=""
                %>
                <!--#include file="../includes/pcCategoriesList.asp"-->
                <%call pcs_CatList()%>							
                </td>
            </tr>
            <tr>
                <td class="pcCPspacer"></td>
            </tr>
			<tr> 
				<th>Indicate how points will be calculated</th>
			</tr>
            <tr>
                <td class="pcCPspacer"></td>
            </tr>
			<tr> 
				<td><input type="radio" name="eqtype" value="1" onclick="document.form1.buttoncheck.value='1';" class="clearBorder"> Points = Price</td>
			</tr>
			<tr> 
				<td>For example, if the price for a product is $28, that product will be assigned 28 points.</td>
			</tr>
			<tr> 
				<td class="pcCPspacer"></td>
			</tr>
			<tr> 
				<td><input type="radio" name="eqtype" value="2" onclick="document.form1.buttoncheck.value='1';" class="clearBorder"> Points = Price * Multiplier <span style="padding-left: 25px;">Multiplier: <input name="multiplier" type="text" size="10" value="10"> (e.g. 10)</span></td>
			</tr>
			<tr> 									
				<td>For example, if the price for a product is $28 and the multiplier is 10, that product will be assigned 280 points.</td>
			</tr>	
			<tr> 
				<td><hr></td>
			</tr>
			<tr> 
				<td align="center">
				<input type="hidden" name="buttoncheck" value="0">
				<input type="submit" name="Submit" value="Assign Reward Points to Products" class="submit2">&nbsp;                
				<input type="button" name="back" value="Back" onClick="document.location.href='rpstart.asp';">
				</td>
			</tr>
		</table>
	</form>
<% end if %>
<!--#include file="Adminfooter.asp"-->
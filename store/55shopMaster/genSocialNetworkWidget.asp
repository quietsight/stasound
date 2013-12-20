<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3%>
<!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/SocialNetworkWidgetConstants.asp"-->
<% pageTitle="Generate E-Commerce Widget for Blogs" %>
<% section="layout" %>
<% dim conntemp, rs, query

Dim pcv_strPageName
pcv_strPageName="genSocialNetworkWidget.asp"

msg=Request("msg")

call opendb()

If request("action")="add" Then

	Session("SNW_TYPE")=Request("exportType")
	Session("SNW_CATEGORY")=Request("catlist")
	Session("SNW_MAX")=Request("prdcount")
	Session("SNW_AFFILIATE")=Request("affiliate")
	
	If Session("SNW_CATEGORY")="" Then
	'	>> Fail
		response.Redirect(pcv_strPageName&"?msg=You must select a category.")  
	End If	
	
	If Session("SNW_TYPE")="0" Then
		
		Dim pcv_strWidgetFlag, pcv_strWidgetError
		pcv_strWidgetFlag=""
		pcv_strWidgetError=""
		
		'// Generate Static XML File
		call pcs_SaveSocialWidgetXML()  
		
		'// Generate Widget
		call pcs_SaveSocialWidgetJS()  

		'// Close Db
		call closedb()

		'// Redirect with Message
		If pcv_strWidgetFlag<>"Success" Then
		'	>> Fail
			response.Redirect(pcv_strPageName&"?msg=There was an error processing your request:" & pcv_strWidgetError)  
		Else	
		'	>> Success
			response.Redirect("../includes/PageCreateSocialNetworkWidget.asp")
		End If
		
	Else

		'// Dynamic XML >>> Generate Constants
		call closedb()
		response.Redirect("../includes/PageCreateSocialNetworkWidget.asp")
	
	End If

End If


Sub pcs_SaveSocialWidgetXML()
	
	Dim pcv_strWidgetXML, pcv_intProductCount
	
	pcv_strWidgetXML = ""
	pcv_strWidgetXML=pcv_strWidgetXML&""
	pcv_intProductCount=0
	
	'// Category ID
	pIdCategory=Session("SNW_CATEGORY")
	
	'// Sort
	if ProdSort="" then
		ProdSort="19"
	end if
	
	select case ProdSort
		Case "19": query1 = " ORDER BY categories_products.POrder Asc"
		Case "0": query1 = " ORDER BY products.SKU Asc"
		Case "1": query1 = " ORDER BY products.description Asc" 	
		Case "2": 
		If Session("customerType")=1 then
			if Ucase(scDB)="SQL" then
				query1 = " ORDER BY (CASE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE Products.bToBPrice WHEN 0 THEN Products.Price ELSE Products.bToBPrice END) ELSE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN Products.pcProd_BTODefaultPrice ELSE Products.pcProd_BTODefaultWPrice END) END) DESC"
			else
				query1 = " ORDER BY (iif(iif((Products.pcProd_BTODefaultWPrice=0) OR (IsNull(Products.pcProd_BTODefaultWPrice)),iif(IsNull(Products.pcProd_BTODefaultPrice),0,Products.pcProd_BTODefaultPrice),Products.pcProd_BTODefaultWPrice)=0,iif(Products.btoBPrice=0,Products.Price,Products.btoBPrice),iif((Products.pcProd_BTODefaultWPrice=0) OR (IsNull(Products.pcProd_BTODefaultWPrice)),Products.pcProd_BTODefaultPrice,Products.pcProd_BTODefaultWPrice))) DESC"
			end if
		else
			if Ucase(scDB)="SQL" then
				query1 = " ORDER BY (CASE (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) WHEN 0 THEN Products.Price ELSE Products.pcProd_BTODefaultPrice END) DESC"
			else
				query1 = " ORDER BY (iif((Products.pcProd_BTODefaultPrice=0) OR (IsNull(Products.pcProd_BTODefaultPrice)),Products.Price,Products.pcProd_BTODefaultPrice)) DESC"
			end if
		End if
	Case "3":
		If Session("customerType")=1 then
			if Ucase(scDB)="SQL" then
				query1 = " ORDER BY (CASE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE Products.bToBPrice WHEN 0 THEN Products.Price ELSE Products.bToBPrice END) ELSE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN Products.pcProd_BTODefaultPrice ELSE Products.pcProd_BTODefaultWPrice END) END) ASC"
			else
				query1 = " ORDER BY (iif(iif((Products.pcProd_BTODefaultWPrice=0) OR (IsNull(Products.pcProd_BTODefaultWPrice)),iif(IsNull(Products.pcProd_BTODefaultPrice),0,Products.pcProd_BTODefaultPrice),Products.pcProd_BTODefaultWPrice)=0,iif(Products.btoBPrice=0,Products.Price,Products.btoBPrice),iif((Products.pcProd_BTODefaultWPrice=0) OR (IsNull(Products.pcProd_BTODefaultWPrice)),Products.pcProd_BTODefaultPrice,Products.pcProd_BTODefaultWPrice))) ASC"
			end if
		else
			if Ucase(scDB)="SQL" then
				query1 = " ORDER BY (CASE (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) WHEN 0 THEN Products.Price ELSE Products.pcProd_BTODefaultPrice END) ASC"
			else
				query1 = " ORDER BY (iif((Products.pcProd_BTODefaultPrice=0) OR (IsNull(Products.pcProd_BTODefaultPrice)),Products.Price,Products.pcProd_BTODefaultPrice)) ASC"
			end if
		End if 	 	
	end select
	
	'// Query Products
	query="SELECT products.idProduct, products.sku, products.description, products.price, products.listhidden, products.listprice, products.serviceSpec, products.bToBPrice, products.smallImageUrl,products.noprices,products.stock, products.noStock,products.pcprod_HideBTOPrice, POrder,products.FormQuantity,products.pcProd_BackOrder FROM products, categories_products WHERE products.idProduct=categories_products.idProduct AND categories_products.idCategory="& pIdCategory &" AND active=-1 AND configOnly=0 and removed=0 " & query1
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)		
	if NOT rs.EOF then
		pcArray_Products = rs.getRows()
		pcv_intProductCount = UBound(pcArray_Products,2)
	end If	
	set rs=nothing
	
	'// Set XML
	pcv_strWidgetXML=pcv_strWidgetXML&"<?xml version='1.0' encoding='iso-8859-1'?>"
	pcv_strWidgetXML=pcv_strWidgetXML&"<Products>"

	pCnt=0
	if pcv_intProductCount=0 then
	'	>> Fail
		response.Redirect(pcv_strPageName&"?msg=The category you selected is empty.")
	end if
	
	do while (pCnt <= pcv_intProductCount) AND (pCnt<cint(Session("SNW_MAX")))
	
		pidProduct=""
		pSku=""
		pDescription=""   
		pPrice=""
		pListHidden=""
		pListPrice=""						   
		pserviceSpec=""
		pBtoBPrice=""   
		pSmallImageUrl="" 
		pnoprices=0
		pStock=""
		pNoStock=""
		pcv_intHideBTOPrice=0
		pFormQuantity=""
		pcv_intBackOrder=""
		pidrelation=""						
		psDesc=""
		pSmallImageUrl=""
		pcv_URL=""
	
		pidProduct=pcArray_Products(0,pCnt) '// rs("idProduct")
		pSku=pcArray_Products(1,pCnt) '// rs("sku")
		pDescription=pcArray_Products(2,pCnt) '// rs("description")   
		pPrice=pcArray_Products(3,pCnt) '// rs("price")
		pListHidden=pcArray_Products(4,pCnt) '// rs("listhidden")
		pListPrice=pcArray_Products(5,pCnt) '// rs("listprice")						   
		pserviceSpec=pcArray_Products(6,pCnt) '// rs("serviceSpec")
		pBtoBPrice=pcArray_Products(7,pCnt) '// rs("bToBPrice")   
		pSmallImageUrl=pcArray_Products(8,pCnt) '// rs("smallImageUrl")   
		pnoprices=pcArray_Products(9,pCnt) '// rs("noprices")
		pStock=pcArray_Products(10,pCnt) '// rs("stock")
		pNoStock=pcArray_Products(11,pCnt) '// rs("noStock")
		pcv_intHideBTOPrice=pcArray_Products(12,pCnt) '// rs("pcprod_HideBTOPrice")
		pFormQuantity=pcArray_Products(14,pCnt) '// rs("FormQuantity")
		pcv_intBackOrder=pcArray_Products(15,pCnt) '// rs("pcProd_BackOrder")
		pidrelation=pcArray_Products(0,pCnt) '// rs("idProduct")
		
		if isNULL(pnoprices) OR pnoprices="" then
			pnoprices=0
		end if
		if isNULL(pcv_intHideBTOPrice) OR pcv_intHideBTOPrice="" then
			pcv_intHideBTOPrice="0"
		end if
		if pnoprices=2 then
			pcv_intHideBTOPrice=1
		end if						
								
		'// Get sDesc
		query="SELECT sDesc FROM products WHERE idProduct="&pidrelation&";"
		set rsDescObj=server.CreateObject("ADODB.RecordSet")
		set rsDescObj=conntemp.execute(query)
		psDesc=rsDescObj("sDesc")
		set rsDescObj=nothing
		
		if pSmallImageUrl="" OR isNULL(pSmallImageUrl) then
			pSmallImageUrl="no_image.gif"
		end if		

		pcv_URL=replace((scStoreURL&"/"&scPcFolder&"/pc/"), "//", "/")
		pcv_URL=replace(pcv_URL,"http:/","http://")
		
		pCnt=pCnt+1
		
		pDescription=ClearHTMLTags2(pDescription,2)
		pDescription=replace(pDescription,"&quot;","""")
		If 44<len(pDescription) then
			pDescription=trim(left(pDescription,44)) & "..."
		End If
	
		pcv_strWidgetXML=pcv_strWidgetXML&"<Product>"
		pcv_strWidgetXML=pcv_strWidgetXML&"		<Description><![CDATA["& pDescription &"]]></Description>"
		pcv_strWidgetXML=pcv_strWidgetXML&"		<Price><![CDATA["&money(pPrice)&"]]></Price>"
		pcv_strWidgetXML=pcv_strWidgetXML&"		<SmallImage><![CDATA["&pcv_URL&"catalog/"&pSmallImageUrl&"]]></SmallImage>"
		pcv_strWidgetXML=pcv_strWidgetXML&"		<URL><![CDATA["&pcv_URL&"viewPrd.asp?idproduct="&pidProduct&"]]></URL>"
		pcv_strWidgetXML=pcv_strWidgetXML&"</Product>"

	loop
	
	pcv_strWidgetXML=pcv_strWidgetXML&"</Products>"
	
	'// Set Path
	if PPD="1" then
		pcStrFolder=Server.Mappath ("/"&scPcFolder&"/pc")
	else
		pcStrFolder=server.MapPath("../pc")
	end if
	
	'response.Write(pcStrFolder & "\pcSyndication.xml")
	'response.End()

	'// Write XML File
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	Set a=fs.CreateTextFile(pcStrFolder & "\pcSyndication.xml",True)
	a.Write(pcv_strWidgetXML)
	a.Close
	Set a=Nothing
	Set fs=Nothing
	
	If err.description="" Then
		pcv_strWidgetFlag="Success"
		pcv_strWidgetError=""
	Else
		pcv_strWidgetFlag=""
		pcv_strWidgetError=err.description
	End If 

End Sub


Sub pcs_SaveSocialWidgetJS()
	On Error Resume Next
	
	Dim pcv_strWidgetXML, pcv_NewLine, pcv_URL, pcv_WidgetExists
	pcv_WidgetExists=""
	pcv_URL=""
	pcv_NewLine=CHR(10)	
	pcv_strWidgetXML = ""
	pcv_strWidgetXML=pcv_strWidgetXML&""
	
	'// Set URL
	pcv_URL=replace((scStoreURL&"/"&scPcFolder&"/pc/"), "//", "/")
	pcv_URL=replace(pcv_URL,"http:/","http://")
	pcv_URL=pcv_URL&"pcSyndication_ShowItems.asp"
	
	'// Set JS File
	pcv_strWidgetXML=pcv_strWidgetXML&"var idaffiliate"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"var pcv_ad_width = 198;"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"var pcv_ad_height = 438;"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"var pcv_style_border = ""0px solid #FFFFFF"";"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"var pcv_style_background = ""#FFFFFF"";"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"var pcv_ad_frame = ""0"";"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"var pcv_ad_src = """& pcv_URL &"?idaffiliate=""+idaffiliate;"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"var pcv_window=window;"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"var pcv_doc=document;"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"pcf_ShowItems(pcv_doc,pcv_window);"& pcv_NewLine

	pcv_strWidgetXML=pcv_strWidgetXML&"function pcf_ShowItems(pcv_window,pcv_doc) {"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"	pcv_window.write('<style type=""text/css"">');"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"	pcv_window.write('.pcProductsFrame {');"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"	pcv_window.write('border: '+pcv_style_border+';');"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"	pcv_window.write('background-color: '+pcv_style_background+';');"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"	pcv_window.write('}');"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"	pcv_window.write('</style>');"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"	pcv_window.write('<iframe name=""pcv_ProductsFrame"" width=""'+pcv_ad_width+'"" class=""pcProductsFrame"" height=""'+pcv_ad_height+'"" frameborder=""'+pcv_ad_frame+'"" src=""'+pcv_ad_src+'"" marginwidth=""0"" marginheight=""0"" vspace=""0"" hspace=""0"" allowtransparency=""false"" scrolling=""no"">');"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"	pcv_window.write('</iframe>');"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"}"& pcv_NewLine
	
	'response.Write(pcv_strWidgetXML)
	'response.End()

	'// Set Path
	if PPD="1" then
		pcStrFolder=Server.Mappath ("/"&scPcFolder&"/pc")
	else
		pcStrFolder=server.MapPath("../pc")
	end if	

	'// Load JS File
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	if fs.FileExists(pcStrFolder & "\pcSyndication.js")=true then
		pcv_WidgetExists="1"
	else
		pcv_WidgetExists="0"
	end if
	set fs=nothing
	
	if pcv_WidgetExists="0" then	
	
		'// Write JS File
		Set fs=Server.CreateObject("Scripting.FileSystemObject")
		Set a=fs.CreateTextFile(pcStrFolder & "\pcSyndication.js",True)
		a.Write(pcv_strWidgetXML)
		a.Close
		Set a=Nothing
		Set fs=Nothing
	
	end if

	If err.description="" Then
		pcv_strWidgetFlag="Success"
		pcv_strWidgetError=""
	Else
		pcv_strWidgetFlag=""
		pcv_strWidgetError=err.description
	End If 

End Sub
%>
<!--#include file="AdminHeader.asp"-->
<style>
.pcCPOverview {
	background-color: #F5F5F5;
	border: 1px solid #FF9900;
	margin: 5px;
	padding: 5px;
	color: #666666;
	font-size:11px;
	text-align: left;
}
.pcCodeStyle {
	font-family: "Courier New", Courier, monospace;
	color: #FF0000;
	font-size: 9;
}
</style>
<% If msg="success" Then %>

<%
Session("SNW_TYPE")=""
Session("SNW_CATEGORY")=""
Session("SNW_MAX")=""
Session("SNW_AFFILIATE")=""
%>
<table class="pcCPcontent">
<tr>
	<td align="center">
		<div align="center">
			<strong>ProductCart E-Commerce Widget created successfully!</strong>
			<br />
      <br />
         	<div class="pcCPOverview">
            	Add the following snippet of JavaScript code to your blog or other page that supports JavaScript to display your widget (<a href="http://www.earlyimpact.com/support/userGuides/socialNetworking.asp" target="_blank">User Guide</a>):
                <br />
                <br />
              	<span class="pcCodeStyle">
                	<%
					pcv_URL=replace((scStoreURL&"/"&scPcFolder&"/pc/"), "//", "/")
					pcv_URL=replace(pcv_URL,"http:/","http://")
					pcv_URL=pcv_URL&"pcSyndication.js"
					%>
					&lt;script type=&quot;text/javascript&quot; src=&quot;<%=pcv_URL%>&quot;&gt; &lt;/script&gt;            
              	</span>
              	<br />
       	</div>
			<br />
            <br />
			<a href="../pc/pcSyndication_Preview.asp?path=<%=pcv_URL%>" target="_blank">Preview Widget</a> | <a href="genSocialNetworkWidget.asp">Generate New E-Commerce Widget</a>
		</div>
	</td>
</tr>
<tr>
	<td class="pcCPspacer"></td>
</tr>
</table>

<% Else %>

<form method="post" name="form1" action="<%=pcv_strPageName%>?action=add" class="pcForms">
	<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2"> How it works</th>
	</tr>
	<%if msg<>"" then%>
	<tr>
		<td colspan="2">
			<div class="pcCPmessage">
				<%=msg%>
			</div>
		</td>
	</tr>
	<%end if%>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2">
			<p>The <strong>ProductCart E-Commerce Widget for Blogs</strong> allows you to take products you have for sale on your store, and show them on a blog or other Web page including popular social networks. The Widget can be styled to match the look and feel of any web site. Read the <a href="http://www.earlyimpact.com/support/userGuides/socialNetworking.asp" target="_blank">ProductCart E-Commerce Widget User Guide</a> for more information. You can choose one of <strong>two publishing methods</strong>:</p>
			<ul>
		    <li><strong>Static Widget:</strong>&nbsp;&nbsp;For better performance. Product information not updated in real time. See the User Guide for details.</li>
	      <li><strong>Dynamic Widget:</strong>&nbsp;&nbsp;Product information always up-to-date, but could cause performance issues.</li>
		  </ul>		  </td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">Creating or Updating your Widget</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2"><p>Show Affiliate Links: 
		    <input type="radio" name="affiliate" value="1" class="clearBorder" <% if SNW_AFFILIATE="1" then response.Write("checked") %>> 
		    Yes&nbsp;&nbsp;
		    <input type="radio" name="affiliate" value="0" class="clearBorder" <% if SNW_AFFILIATE="0" then response.Write("checked") %>> 
		    No</p></td>
	</tr>
	<tr>
		<td colspan="2"><p>Choose the publishing method: <input type="radio" name="exportType" value="0" class="clearBorder" <% if SNW_TYPE="0" then response.Write("checked") %>> Static Widget &nbsp;&nbsp;<input type="radio" name="exportType" value="1" class="clearBorder" <% if SNW_TYPE="1" then response.Write("checked") %>> Dynamic Widget</p></td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"><hr></td>
	</tr>
	<tr>
		<td align="right"><input type="text" name="prdcount" size="3" value="<%=SNW_MAX%>"></td>
		<td nowrap="nowrap">Maximum of products per Widget (e.g. 25)<br />
		  <em>The more products you publish with your Widget the longer it will take to load.</em></td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"><hr></td>
	</tr>
    <tr> 
      <td align="right" valign="top" nowrap="nowrap">
      	Choose a Category:</td>
      <td width="90%" valign="top">
            <select size="8" name="catlist">
            <% query="SELECT idcategory, idParentCategory, categorydesc FROM categories ORDER BY categoryDesc ASC"
            set rstemp=conntemp.execute(query)
            if err.number <> 0 then
                call closeDb()
                response.redirect "techErr.asp?error="& Server.Urlencode("Error in retreiving categories from database: "&Err.Description) 
            end If
            if rstemp.eof then
                catnum="0"
            end If
            if catnum<>"0" then
                
				do until rstemp.eof
                    idparentcategory=""
					idparentcategory=rstemp("idParentCategory")
                   	if idparentcategory<>"1" then 
						
						set rs2=server.createobject("adodb.recordset")
						query="SELECT categoryDesc FROM categories where idcategory="& idparentcategory & " AND iBTOHide<>1"						
						set rs2=conntemp.execute(query)
						if NOT rs2.EOF then
							parentDesc=rs2("categoryDesc")							
							%>                            
							<option value='<%response.write rstemp("idcategory")%>'><%response.write rstemp("categorydesc")&" - Parent: "&parentDesc %></option>
							<% 
						end if
                	
					else
						%>
                        <option value='<%response.write rstemp("idcategory")%>' <% if cint(SNW_CATEGORY)=cint(rstemp("idcategory")) then response.Write("selected") %>><%=rstemp("categorydesc")%></option>
                        <%
					end if
					rstemp.movenext
                loop
            End if %>
        </select>
        <br />
        <br />
        <span class="pcCPnotes"><strong>Tip:</strong> We recommend creating a hidden category called something like &quot;E-Commerce Widget&quot;. The hidden category should contain all the products you want displayed in your widget. Then, select the &quot;E-Commerce Widget&quot; category in the category list shown menu above.</span>
        <br />
      </td>
    </tr>
	<tr>
		<td colspan="2" class="pcCPspacer"><hr></td>
	</tr>
	<tr> 
		<td colspan="2" style="text-align: center;">
			<input name="submit" type="submit" class="submit2" value="Generate ProductCart E-Commerce Widget">
		</td>
	</tr>
	</table>
</form>
<%
End If
call closedb()%><!--#include file="AdminFooter.asp"-->
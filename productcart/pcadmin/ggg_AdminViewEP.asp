<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Dim pageTitle, Section
pageTitle="Product Details"
Section="orders" %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<% 

dim query, conntemp, rstemp

call openDb()

gIDEvent=getUserInput(request("IDEvent"),0)
geID=getUserInput(request("geID"),0)

query="select pcEv_IDEvent,pcEv_Name,pcEv_Date,pcEv_Type,pcEv_IncGcs from pcEvents where pcEv_IDEvent=" & gIDEvent & " and pcEv_Active=1"
set rstemp=connTemp.execute(query)

if rstemp.eof then
else
	gIDEvent=rstemp("pcEv_IDEvent")
	geName=rstemp("pcEv_Name")
	geDate=rstemp("pcEv_Date")
	if gedate<>"" then
		if scDateFrmt="DD/MM/YY" then
			gedate=(day(gedate)&"/"&month(gedate)&"/"&year(gedate))
		else
			gedate=(month(gedate)&"/"&day(gedate)&"/"&year(gedate))
		end if
	end if
	gType=rstemp("pcEv_Type")
	if gType<>"" then
	else
		gType=""
	end if
	geincGc=rstemp("pcEv_IncGcs")
	if geincGc<>"" then
	else
		geincGc="0"
	end if
end if
set rstemp=nothing

query="select products.idproduct,products.sku,products.description,products.imageUrl,products.largeImageURL,products.details,products.sdesc,pcEvProducts.pcEP_Price,pcEvProducts.pcEP_IDConfig, pcEvProducts.pcEP_OptionsArray, pcEvProducts.pcEP_xdetails from products,pcEvProducts where pcEvProducts.pcEP_IDEvent=" & gIDEvent & " and products.idproduct=pcEvProducts.pcEP_IDProduct and pcEvProducts.pcEP_ID=" & geID
set rstemp=connTemp.execute(query)

if rstemp.eof then
else
	pIDProduct=rstemp("idproduct")
	pSku=rstemp("sku")
	pname=rstemp("description")
	pImageURL=rstemp("imageUrl")
	pLgimageURL=rstemp("largeImageURL")
	pdetails=rstemp("details")
	psdesc=rstemp("sdesc")
	pPrice=rstemp("pcEP_Price")
	pIDConfig=rstemp("pcEP_IDConfig")
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Start: Product Options
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcv_strSelectedOptions=""
	pcv_strSelectedOptions = rstemp("pcEP_OptionsArray")
	pcv_strSelectedOptions=pcv_strSelectedOptions&""		
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' End: Product Options
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	pxdetails=rstemp("pcEP_xdetails")
end if
set rstemp=nothing

%> 
<!--#include file="AdminHeader.asp"-->
<script language="JavaScript">
<!--

imagename='';
function enlrge(imgnme)
{
	lrgewin=window.open("about:blank","","height=200,width=200,status=no")
	imagename=imgnme;
	setTimeout('update()',500)
}

function update()
{
	<%
	'**** Check MAC IE Browser ******************

	UserBrowser=Request.ServerVariables("HTTP_USER_AGENT")
	MACBrowser=instr(ucase(UserBrowser),"MAC")

	'**** End of Check MAC IE Browser ******************
	%>

	doc=lrgewin.document;
	doc.open('text/html');
	doc.write('<HTML><HEAD><TITLE>Enlarged Image<\/TITLE><\/HEAD><BODY bgcolor="white" onLoad="if  (self.resizeTo)self.resizeTo((document.images[0].width+10),(document.images[0].height+100))" topmargin="4" leftmargin="0" rightmargin="0" bottommargin="0"><table width=""' + document.images[0].width + '" border="0" cellspacing="0" cellpadding="0"><tr><td>');
	doc.write('<IMG SRC="' + imagename + '"><\/td><\/tr><tr><td><%if MACBrowser=0 then%><form name="viewn"><input type="image" src="images/close.gif" align="right" value="Close Window" onClick="self.close()"><%end if%><\/td><\/tr><\/table>');
	doc.write('<\/form><\/BODY><\/HTML>');
	doc.close();
}

//-->
</script>
<table class="pcCPcontent">
<tr>
	<td width="100%">
		<h2><%=pname%></h2>
		<%response.write dictLanguage.Item(Session("language")&"_viewEP_1")%><%=pSKU%><br>
		<br>
		<%=psdesc%><br>
		<a href="#1"><%response.write dictLanguage.Item(Session("language")&"_viewEP_7")%></a>
		<br>
		<br>
		<b><%response.write dictLanguage.Item(Session("language")&"_viewEP_2")%></b><%=scCurSign & money(pPrice)%>
		<br>
		<br>
		<%IF pIDConfig<>"0" then
			query="SELECT * FROM configSessions WHERE idconfigSession=" & pIDConfig
			set rs=conntemp.execute(query)

			stringProducts=rs("stringProducts")
			stringValues=rs("stringValues")
			stringCategories=rs("stringCategories")
			ArrProduct=Split(stringProducts, ",")
			ArrValue=Split(stringValues, ",")
			ArrCategory=Split(stringCategories, ",")
			Qstring=rs("stringQuantity")
			ArrQuantity=Split(Qstring,",")
			Pstring=rs("stringPrice")
			ArrPrice=split(Pstring,",")
			stringCProducts=rs("stringCProducts")
			stringCValues=rs("stringCValues")
			stringCCategories=rs("stringCCategories")
			ArrCProduct=Split(stringCProducts, ",")
			ArrCValue=Split(stringCValues, ",")
			ArrCCategory=Split(stringCCategories, ",")
			set rs=nothing

			if ArrProduct(0)="na" then
			else%>
				<b><%response.write dictLanguage.Item(Session("language")&"_viewEP_3")%></b><br>
				<%tempCat=""%>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<%
				for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
				query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
				set rsObj=conntemp.execute(query)
				query="SELECT displayQF FROM configSpec_Products WHERE configProduct="&ArrProduct(i) & " and specProduct=" & pIDProduct 
				set rsObj1=conntemp.execute(query)%>
				<tr>
					<td width="10%">&nbsp;</td>
					<td>
						<%if tempCat<>rsObj("categoryDesc") then
						tempCat=rsObj("categoryDesc")%>
						<b><%=tempCat%></b><br>
						<%end if%>				
						<%=rsObj("description")%>
						<%if rsObj1("displayQF")=True then%>
						- <%response.write dictLanguage.Item(Session("language")&"_viewEP_9")%><%=ArrQuantity(i)%>
						<%end if%>
					</td>
				</tr>
				<% set rsObj=nothing
				set rsObj1=nothing
				next%>
				</table><br>
			<%end if 'End of Configuration
	
			if ArrCProduct(0)="na" then 'Additional Charges
			else%>
				<b><%response.write dictLanguage.Item(Session("language")&"_viewEP_4")%></b><br>
				<%tempCat=""%>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<%
				for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
				query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))" 
				set rsObj=conntemp.execute(query)
				%>
				<tr>
					<td width="10%">&nbsp;</td>
					<td>
						<%if tempCat<>rsObj("categoryDesc") then
							tempCat=rsObj("categoryDesc")%>
							<b><%=tempCat%></b><br>
						<%end if%>				
						<%=rsObj("description")%>
					</td>
				</tr>
				<% set rsObj=nothing
				next%>
				</table>					
			<%end if 'End of Additional Charges%>
			<br>
		<%END IF 'Have BTO%>


	<%
	'*************************************************************************************************
	' START: GET OPTIONS
	'*************************************************************************************************
	Dim pPriceToAdd, pOptionDescrip, pOptionGroupDesc, pcv_strSelectedOptions
	Dim pcArray_SelectedOptions, pcv_strOptionsArray, cCounter, xOptionsArrayCount
	Dim pcv_strOptionsPriceArray, pcv_strOptionsPriceArrayCur, pcv_strOptionsPriceTotal
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Get the Options for the item
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	IF len(pcv_strSelectedOptions)>0 AND pcv_strSelectedOptions<>"NULL" THEN
	 
		pcArray_SelectedOptions = Split(pcv_strSelectedOptions,chr(124))
		
		pcv_strOptionsArray = ""
		pcv_strOptionsPriceArray = ""
		pcv_strOptionsPriceArrayCur = ""
		pcv_strOptionsPriceTotal = 0
		xOptionsArrayCount = 0
		
		For cCounter = LBound(pcArray_SelectedOptions) TO UBound(pcArray_SelectedOptions)
			
			' SELECT DATA SET
			' TABLES: optionsGroups, options, options_optionsGroups
			query = 		"SELECT optionsGroups.optionGroupDesc, options.optionDescrip, options_optionsGroups.price, options_optionsGroups.Wprice "
			query = query & "FROM optionsGroups, options, options_optionsGroups "
			query = query & "WHERE idoptoptgrp=" & pcArray_SelectedOptions(cCounter) & " "
			query = query & "AND options_optionsGroups.idOption=options.idoption "
			query = query & "AND options_optionsGroups.idOptionGroup=optionsGroups.idoptiongroup "	
			
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				'//Logs error to the database
				call LogErrorToDatabase()
				'//clear any objects
				set rs=nothing
				'//close any connections
				call closedb()
				'//redirect to error page
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if					
			
			if Not rs.eof then 
				
				xOptionsArrayCount = xOptionsArrayCount + 1
				
				pOptionDescrip=""
				pOptionGroupDesc=""
				pPriceToAdd=""
				pOptionDescrip=rs("optiondescrip")
				pOptionGroupDesc=rs("optionGroupDesc")
				
				If Session("customerType")=1 Then
					pPriceToAdd=rs("Wprice")
					If rs("Wprice")=0 then
						pPriceToAdd=rs("price")
					End If
				Else
					pPriceToAdd=rs("price")
				End If	
				
				'// Generate Our Strings
				if xOptionsArrayCount > 1 then
					pcv_strOptionsArray = pcv_strOptionsArray & chr(124)
					pcv_strOptionsPriceArray = pcv_strOptionsPriceArray & chr(124)
					pcv_strOptionsPriceArrayCur = pcv_strOptionsPriceArrayCur & chr(124)
				end if
				'// Column 4) This is the Array of Product "option groups: options"
				pcv_strOptionsArray = pcv_strOptionsArray & pOptionGroupDesc & ": " & pOptionDescrip
				'// Column 25) This is the Array of Individual Options Prices
				pcv_strOptionsPriceArray = pcv_strOptionsPriceArray & pPriceToAdd
				'// Column 26) This is the Array of Individual Options Prices, but stored as currency "scCurSign & money(pcv_strOptionsPriceTotal) "
				pcv_strOptionsPriceArrayCur = pcv_strOptionsPriceArrayCur & scCurSign & money(pPriceToAdd)
				'// Column 5) This is the total of all option prices
				pcv_strOptionsPriceTotal = pcv_strOptionsPriceTotal + pPriceToAdd
				
			end if
			
			set rs=nothing
		Next
	
	END IF	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Get the Options for the item
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~			

				
	'*************************************************************************************************
	' END: GET OPTIONS
	'*************************************************************************************************
	%>
	
	<%
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START: SHOW PRODUCT OPTIONS
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Dim pcArray_strOptionsPrice, pcArray_strOptions, pcv_intOptionLoopSize, pcv_intOptionLoopCounter, tempPrice, tAprice
	
	if len(pcv_strOptionsArray)>0 then 
	%>
	
	
	<%'response.write dictLanguage.Item(Session("language")&"_Custwlview_15")%>
	
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<%
		'#####################
		' START LOOP
		'#####################	
		'// Generate Our Local Arrays from our Stored Arrays
		
		' Column 11) pcv_strSelectedOptions '// Array of Individual Selected Options Id Numbers	
		pcArray_strSelectedOptions = ""					
		pcArray_strSelectedOptions = Split(trim(pcv_strSelectedOptions),chr(124))
		
		' Column 25) pcv_strOptionsPriceArray '// Array of Individual Options Prices
		pcArray_strOptionsPrice = ""
		pcArray_strOptionsPrice = Split(trim(pcv_strOptionsPriceArray),chr(124))
		
		' Column 4) pcv_strOptionsArray '// Array of Product "option groups: options"
		pcArray_strOptions = ""
		pcArray_strOptions = Split(trim(pcv_strOptionsArray),chr(124))
		
		' Get Our Loop Size
		pcv_intOptionLoopSize = 0
		pcv_intOptionLoopSize = Ubound(pcArray_strSelectedOptions)
		
		' Start in Position One
		pcv_intOptionLoopCounter = 0
		
		' Display Our Options
		For pcv_intOptionLoopCounter = 0 to pcv_intOptionLoopSize
		%>
		<tr>
			<td>
			
				<%= pcArray_strOptions(pcv_intOptionLoopCounter)%>
			
													
				<% 
				tempPrice = pcArray_strOptionsPrice(pcv_intOptionLoopCounter)
				
				if tempPrice="" or tempPrice=0 then
					response.write "&nbsp;"
				else 
					response.write " (" & scCurSign&money(tempPrice) & ")"
				end if 
				%>			
			
			</td>
		</tr>
		<%
		Next
		'#####################
		' END LOOP
		'#####################
		%>
		<tr><td>&nbsp;</td></tr>
	</table>
	
	<% 
	End if
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END: SHOW PRODUCT OPTIONS
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	%>							
		
		
		<%if pxdetails<>"" then%>
			<%=pxdetails%><br>
		<%end if%>
	</td>
	<!-- Show Product Images -->
	<td valign="top">
	<%if pImageUrl<>"" then%>
		<img src='../pc/catalog/<%=pImageUrl%>' alt="" hspace="10"> 
		<% ' show link to detail view image if it exists
		if pLgimageURL<>"" then%>
			<br>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr> 
				<td valign="bottom" align="right">
					<% If InStr(Request.ServerVariables("HTTP_USER_AGENT"),"MSIE") Then %>
						<a href="javascript:enlrge('../pc/catalog/<%=pLgimageURL%>')">
					<% Else If InStr(Request.ServerVariables("HTTP_USER_AGENT"),"Safari") Then 
						Response.Write("<a href=""../pc/catalog//"&pLgimageURL&""" target=""_blank"">")
					Else
						 Response.Write("<a href=""../pc/catalog//"&pLgimageURL&""" target=""_blank"">")
					End If
					End If %>Zoom +</a>
			   </td>
		  </tr>
		 </table>
		<%end if
	'if no image, show no_image.gif
	else%>
		<img src='../pc/catalog/no_image.gif' alt="Product image not available" width="100" height="100">
	<% end if%>
	</td>
	<!-- End of Show Product Images -->
</tr>
<tr>
	<td>
		<hr size="1">
		<a name="1"></a>
		<%response.write dictLanguage.Item(Session("language")&"_viewEP_8")%><br>
		<br>
		<%=pDetails%>
		<br>
		<br>
		<input type="button" name="back" value=" Back " onclick="javascript:history.go(-1);" class="ibtnGrey">
	</td>
</tr>
</table>
<% call closedb() %>
<!--#include file="adminFooter.asp"-->
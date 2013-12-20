<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Modify Product Options" %>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/currencyformatinc.asp" -->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="inc_UpdateDates.asp" -->
<!--#include file="AdminHeader.asp"-->
<% dim f, query, conntemp, rstemp, pIdProduct

call openDB()

'***********************************************************************
' START: ON POST BACK
'***********************************************************************
If request.Form("Submit2")<>"" then
	
	'// Get Our Data
	ROption=replacecomma(request.form("ROption"))
	If trim(ROption)="" then
		ROption="0"
	End If	
			
	pIdProductArray=request.form("idProduct")
	if pIdProductArray<>"" then
		pIdProductArray=replace(pIdProductArray," ","")
	end if
	
	pIdOptionGroup=request.form("idOptionGroup")
	
	'// Split my product array
	pIdProduct=split(pIdProductArray,",")
	errarray=""
	cnt=0
	repeatcnt=0
	
	'////////////////////////////////////////////////////////////////
	'// START: PRODUCT LOOP
	'////////////////////////////////////////////////////////////////
	For i=lBound(pIdProduct) to UBound(pIdProduct)
		if pIdProduct(i)<>"" then
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'  START: Product Option Level Loop
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			For Each intOptionID in Request.Form("idoption")
				
				'// Get Our Data
				pPrice=replacecomma(request.form("price"&intOptionID))
				If pPrice="" then
					pPrice="0"
				End If
				
				pWPrice=replacecomma(request.form("Wprice"&intOptionID))
				If pWPrice="" then
					pWPrice="0"
				End If
				
				pOrd=replacecomma(request.form("ORD"&intOptionID))
				If pOrd="" then
					pOrd="0"
				End If	
						
				'// Check if it exists in database before adding
				contgo=0			
				strSQL="SELECT * FROM options_optionsGroups WHERE idproduct="&pIdProduct(i)&" AND idoptionGroup="&pIdOptionGroup&" AND idOption="&intOptionID&";"
				set rstemp=conntemp.execute(strSQL)
				if rstemp.eof then
					'// Set the update flag
					contgo=1
					'// Add the option
					strSQL="INSERT INTO options_optionsGroups (idproduct, idoptionGroup, idOption, Price, Wprice,sortOrder) VALUES (" & pIdProduct(i) &", " & pIdOptionGroup & ", " & intOptionID & ","& pPrice &","& pWPrice &"," & pOrd & ")"
					set rstemp=conntemp.execute(strSQL)					
				end if
			Next
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'  END: Product Option Level Loop
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'  START: Product Level Tasks
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// If the Option was added for this product check the relationships
			if contgo=1 then
				
				'// If this is a new option group, then we need to add the relation
				strSQL="SELECT idOptionGroup, idproduct FROM pcProductsOptions WHERE idproduct="& pIdProduct(i) &" AND idOptionGroup="& pIdOptionGroup &" "
				'response.Write(strSQL)
				'response.end
				set rsOptionCheck=conntemp.execute(strSQL)	
				if rsOptionCheck.eof then
					strSQL="INSERT INTO pcProductsOptions (idproduct, idOptionGroup, pcProdOpt_Required, pcProdOpt_Order) VALUES (" & pIdProduct(i) &", " & pIdOptionGroup & ", " & ROption & ", 0)"
					set rstemp=conntemp.execute(strSQL)
					'// if the option group is new keep count
					cnt=cnt+1
				end if
				set rsOptionCheck = nothing
				
			end if
			
			'// If the Option was NOT added for this product keep count
			if contgo=0 then
				repeatcnt=repeatcnt+1
			else
				call updPrdEditedDate(pIdProduct(i))
			end if
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'  END: Product Level Tasks
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
		end if
	Next
	'////////////////////////////////////////////////////////////////
	'// END: PRODUCT LOOP
	'////////////////////////////////////////////////////////////////
	
	query="SELECT * FROM OptionsGroups WHERE idOptionGroup="&pIdoptionGroup
	set rstemp=conntemp.execute(query)
	OptionGroupDesc=rstemp("OptionGroupDesc")
	set rstemp=nothing
	call closeDb()
	
	pcv_strMsg = " The option attributes from the group "& OptionGroupDesc &" were assigned to " & cnt & " product(s). "
	
	If repeatcnt>0 then 
		pcv_strMsg = pcv_strMsg & repeatcnt &" product(s) were not updated because these option attributes had already been assigned to them."
	end if

	response.redirect "manageOptions.asp?s=1&msg="&server.urlencode(pcv_strMsg)
	response.end
End If
'***********************************************************************
' END: ON POST BACK
'***********************************************************************

' form parameter
pIdOptionsGroups=request.Querystring("idOptionGroup")
pIdProduct=request("prdlist")

if trim(pIdProduct)="" then
   response.redirect "msg.asp?message=2"
end if
%>
<form method="post" name="modifyProduct" action="modPrdOpta4.asp" class="pcForms">
<table class="pcCPcontent">
	<tr> 
	<td colspan="5">
		<input type="hidden" name="idproduct" value="<%=pIdProduct%>">
		<table width="100%">
		<tr>
			<td width="50%" valign="top"> 
					<input type="hidden" name="idOptionGroup" value="<%=pIdOptionsGroups%>">
					<% query="SELECT * FROM optionsGroups WHERE idoptionGroup=" &pIdOptionsGroups
					set rstemp=conntemp.execute(query) %>
					Assigning product option: <b><%=rstemp("optionGroupDesc")%></b><br>
			</td>
			<td valign="top" align="right"><input type="checkbox" name="ROption" value="1" class="clearBorder"> Required Option</td>
		</tr>
		</table>
	</td>
</tr>   
<tr>
	<td colspan="5"></td>
</tr>               
<tr>          
	<th colspan="2">Attributes</th>
	<th nowrap>Additional Price</th>
	<th nowrap>Wholesale Price</th>
	<th nowrap>Order</th>
</tr>
<tr>
	<td colspan="5"></td>
</tr>
<% 
query="SELECT options.optionDescrip, options.idoption FROM options, optGrps WHERE options.idoption=optGrps.idoption AND  optGrps.idoptionGroup="& rstemp("idoptionGroup") &" ORDER BY optionDescrip"
set rstemp=conntemp.execute(query)
noAttribute="0"
If rstemp.eof then 
	noAttribute="1"%>
<tr> 
	<td colspan="5">
		<div class="pcCPmessage">
			There are currently no attributes assigned to this Option Group.
		</div>
	</td>
</tr>
<% else
	Do until rstemp.eof %>
<tr> 
	<td width="1%"> 
		<input type="checkbox" name="idoption" value="<%=rstemp("idoption")%>" class="clearBorder">
	</td>
	<td width="60%"><%=rstemp("optionDescrip")%></td>
	<td> 
		<%=scCurSign%> <input type="text" name="price<%=rstemp("idoption")%>" size="6" maxlength="10" value="0">
	</td>
	<td> 
		<%=scCurSign%> <input type="text" name="Wprice<%=rstemp("idoption")%>" size="6" maxlength="10" value="0">
	</td>
	<td align="center"> 
		<input type="text" name="ORD<%=rstemp("idoption")%>" size="3" maxlength="10" value="">
	</td>
</tr>
	<% 
	rstemp.movenext
	loop 
end if%>
<tr>
	<td colspan="5" class="pcSpacer">&nbsp;</td>
</tr>
<tr>
	<td colspan="5" align="center"> 
		<input type="submit" name="Submit2" value="Continue" class="submit2">
		&nbsp;
		<input type="button" name="Button" value="Back" onClick="javascript:history.back()">
	</td>
</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->
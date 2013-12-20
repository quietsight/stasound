<%Server.ScriptTimeout = 5400
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin="1*3*"%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/utilities.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../pc/pcSeoFunctions.asp"-->
<%
dim connTemp, rsTemp, query, rs
Dim pcCatArr,CatRecords
Dim CatLevelLimit
Dim pcv_intUseFontTags,  pcv_strStyle, pcv_strStyleHeading, pcv_strStyleLink, pcv_strSMFont, pcv_strSMFSize, pcv_strSMFColor, pcv_strSMLColor, pcv_listType

call opendb()

'/////////////////////////////////////////////////////////////////////////////////////////
'// START: SETTINGS
'/////////////////////////////////////////////////////////////////////////////////////////
CatLevelLimit=4
pcv_strSMBody = ""
pcv_strSMFont=request("fontname")
pcv_strSMFSize=request("fontsize")
pcv_strSMFColor=request("fontcolor")
pcv_strSMLColor=request("linkcolor")
pcv_strHtags=request("htags")
pcv_strUseProDesc=request("prodesc")
pcv_strUseCatDesc=request("catdesc")
SPathInfo=""
pcv_strDelCatList=request("catlist")
pcv_intUseFontTags=request("storefont")
pcv_CatOnly=request("catonly")
if pcv_CatOnly="" then
	pcv_CatOnly="0"
end if
if (pcv_strDelCatList<>"") then
	pcv_strDelCatList=" " & pcv_strDelCatList & ","
end if

if pcv_intUseFontTags="1" then
	pcv_strStyle = " style=""font-family: " & pcv_strSMFont & "; font-size: " & pcv_strSMFSize & "px; color: " & pcv_strSMFColor &";"""
	pcv_strStyleHeading = " style=""font-family: " & pcv_strSMFont & "; font-size: " & pcv_strSMFSize & "; color: " & pcv_strSMFColor &";"""	
	pcv_strStyleLink = " style=""font-family: " & pcv_strSMFont & "; font-size: " & pcv_strSMFSize & "px; color: " & pcv_strSMLColor &";"""
	else
	pcv_strStyle = ""
	pcv_strStyleHeading = ""
	pcv_strStyleLink = ""
end if 
'/////////////////////////////////////////////////////////////////////////////////////////
'// END: SETTINGS
'/////////////////////////////////////////////////////////////////////////////////////////



'/////////////////////////////////////////////////////////////////////////////////////////
'// START: GENERATE STORE MAP
'/////////////////////////////////////////////////////////////////////////////////////////

'// New String Builder
set StringBuilderObj = new StringBuilder

GenSMBody()
strStoreMap = StringBuilderObj.toString 

'// Clean Up
set StringBuilderObj = nothing

'// Get Paths
SPath1=Request.ServerVariables("PATH_INFO")
mycount1=0
do while mycount1<2
	if mid(SPath1,len(SPath1),1)="/" then
		mycount1=mycount1+1
	end if
	if mycount1<2 then
		SPath1=mid(SPath1,1,len(SPath1)-1)
	end if
loop
SPathInfo1 = SPath1	
if Right(SPathInfo1,1)="/" then
	SPathInfo1=SPathInfo1 & "pc/"					
else
	SPathInfo1=SPathInfo1 & "/pc/"
end if

'// Get Map Path
SavedFile = SPathInfo1 & "catalog/inc_StoreMap.asp"
findit = Server.MapPath(Savedfile)

'// Save Store Map
Set fso = server.CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateTextFile(findit,True,False)
on error resume next
dim intShowErrMsg
intShowErrMsg=0
if err.number<>0 then
	intShowErrMsg=1
	err.number=0
end if
f.Write(strStoreMap)
if err.number<>0 then
	intShowErrMsg=1
	err.number=0
end if
f.close
'/////////////////////////////////////////////////////////////////////////////////////////
'// END: GENERATE STORE MAP
'/////////////////////////////////////////////////////////////////////////////////////////
%>
<% pageTitle="Store Map Generation Results" %>
<% section="layout" %>
<!--#include file="AdminHeader.asp"-->
<form method="post" name="form1" action="genStoreMapA.asp" class="pcForms">
<table class="pcCPcontent">
	<tr>
	<td>
    <div class="pcCPmessageSuccess">The store map has been generated successfully! <br /> 
    The file location is: <b>/pc/StoreMap.asp</b>. Click on the link below to view it. <br /><br />
	<a target="_blank" href="../pc/StoreMap.asp">View 
    </a>&nbsp;|&nbsp;<a href="http://wiki.earlyimpact.com/productcart/marketing-generate_store_map" target="_blank">How to use it</a>
    </div>
	</td>
	</tr>
<tr> 
<td class="normal">  
<p align="center">
<input name="back1" type="button" class="ibtnGrey" value="Generate another Store Map" onclick="location='genStoreMap.asp'">&nbsp;
<input name="back" type="button" class="ibtnGrey" value="Return to Start page" onclick="location='menu.asp'">
</td>
</tr>
</table>
</form>
<%
'/////////////////////////////////////////////////////////////////////////////////////////
'// START: SUB - GenProducts
'/////////////////////////////////////////////////////////////////////////////////////////
Sub GenProducts(catID,levelID)

	Dim query,pcPrdArr,PrdCount,i

	query="select products.idproduct,products.description,products.sdesc from categories_products,products where categories_products.idcategory=" & catID & " and products.idproduct=categories_products.idproduct AND products.active<>0 order by products.description asc"
	set rsp=connTemp.execute(query)

	tempStr1=""
	
	IF not rsp.eof then
		if (catID<>1) then
			tempStr1="<ul>"
		end if
		tempStr1=tempStr1 & vbcrlf
		
		pcPrdArr=rsp.GetRows()
		PrdCount=ubound(pcPrdArr,2)
		For i=0 to PrdCount
			IDproduct=pcPrdArr(0,i)
			ProName=pcPrdArr(1,i)
			ShortDesc=pcPrdArr(2,i)

			'// SEO Links
			'// Build Navigation Product Link
			if scSeoURLs=1 then
				pcStrPrdLink=ProName & "-" & catID & "p" & IDproduct & ".htm"
				pcStrPrdLink=removeChars(pcStrPrdLink)
			else
				pcStrPrdLink="viewPrd.asp?idproduct=" & IDProduct
			end if
			'//

			tempStr1=tempStr1 & "<li>" & vbcrlf
			tempStr1=tempStr1 & "<a href=""" & pcStrPrdLink & """" & pcv_strStyleLink & ">" & ProName & "</a>"
			if (pcv_strUseProDesc="1") AND (len(ShortDesc)>0) then
				if NOT pcv_strHtags="1" then
					tempStr1=tempStr1 & "<br />" & vbcrlf
				end if
				tempStr1=tempStr1 & "<span" & pcv_strStyle & ">" & ShortDesc & "</span>" & vbcrlf
			end if
			tempStr1=tempStr1 & "</li>" & vbcrlf 
		Next


		if (catID<>1) then
			tempStr1=tempStr1 & "</ul>" & vbcrlf
		end if
	END IF

	set rsp=nothing
	StringBuilderObj.append tempStr1
	
End Sub
 '/////////////////////////////////////////////////////////////////////////////////////////
'// END: SUB - GenProducts
'/////////////////////////////////////////////////////////////////////////////////////////
 
 
 
'/////////////////////////////////////////////////////////////////////////////////////////
'// START: SUB - GenCat
'/////////////////////////////////////////////////////////////////////////////////////////
Sub GenCat(catID, tmpCatName, tmpCatsDesc, CurrentLevel)

	tempStr1=""
	IDCat=catID
	CATName=tmpCatName
	ShortDesc=tmpCatsDesc

	'// SEO Links
	'// Build Navigation Category Link
	if scSeoURLs=1 then
		pcStrCatLink=CATName & "-c" & IDCat & ".htm"
		pcStrCatLink=removeChars(pcStrCatLink)
	else
		pcStrCatLink="viewCategories.asp?pageStyle=" & bType & "&idCategory=" & IDCat
	end if
  
	if pcv_strHtags="1" then
		StringBuilderObj.append "<h" & CurrentLevel + 1 & ">"
	end if

	StringBuilderObj.append "<a href=""" & SPathInfo & pcStrCatLink & """><span" & pcv_strStyleLink & ">" & CATName & "</span></a>"
	
	if pcv_strHtags="1" then
		StringBuilderObj.append "</h" & CurrentLevel + 1 & ">" & vbcrlf
		end if
 
	if (pcv_strUseCatDesc="1") AND (len(ShortDesc)>0) then
		if NOT pcv_strHtags="1" then
			StringBuilderObj.append "<br />" & vbcrlf 
		end if
		StringBuilderObj.append "<span" & pcv_strStyle & ">" & ShortDesc & "</span>" & vbcrlf 
	end if

End Sub
'/////////////////////////////////////////////////////////////////////////////////////////
'// START: SUB - GenCat
'/////////////////////////////////////////////////////////////////////////////////////////



'/////////////////////////////////////////////////////////////////////////////////////////
'// START: SUB - LoopCats
'/////////////////////////////////////////////////////////////////////////////////////////
Sub LoopCats(IDParent,CurrentLevel)
	Dim HaveCatL,m
	Dim IDCategory,CatName,CatSDesc
	
	if pcv_strHtags="1" then
		pcv_listType = " style=""list-style-type: none; margin-left: 0; padding-left: 1em;"""
		else
		pcv_listType = ""
	end if

	HaveCatL=0
	For m=0 to CatRecords
		if Clng(pcCatArr(1,m))=IDParent then
			if HaveCatL=0 then
				HaveCatL=1
				StringBuilderObj.append "<ul" & pcv_listType & ">" & vbcrlf
			End if
			IDCategory=pcCatArr(0,m)
			CatName=pcCatArr(2,m)
			CatSDesc=pcCatArr(3,m)
			If clng(IDCategory)<>1 then
				If instr(pcv_strDelCatList," " & IDCategory & ",")=0 THEN
					StringBuilderObj.append "<li>" & vbcrlf

					Call GenCat(IDCategory,CatName,CatSDesc,CurrentLevel)

					if CurrentLevel<CatLevelLimit then
						Call LoopCats(IDCategory,Clng(CurrentLevel)+1)
					end if
					
					if pcv_CatOnly="0" then
						Call GenProducts(IDCategory,Clng(CurrentLevel)+1)
					end if

					StringBuilderObj.append "</li>" & vbcrlf
				End if
			End if
		end if
	Next
	
	If CurrentLevel=1 then
		if instr(pcv_strDelCatList," 1,")=0 then
			if pcv_CatOnly="0" then
				Call GenProducts(1,2)
			end if
		end if
	End if
	
	if HaveCatL=1 then
		StringBuilderObj.append "</ul>" & vbcrlf
	end if

End Sub
'/////////////////////////////////////////////////////////////////////////////////////////
'// END: SUB - LoopCats
'/////////////////////////////////////////////////////////////////////////////////////////



'/////////////////////////////////////////////////////////////////////////////////////////
'// START: SUB - GenSMBody
'/////////////////////////////////////////////////////////////////////////////////////////
Sub GenSMBody()

	if pcv_strHtags="1" then
		StringBuilderObj.append "<h1" & pcv_strStyleHeading & ">"
		else
		StringBuilderObj.append "<span" & pcv_strStyle & "><strong>"
	end if
	StringBuilderObj.append "Store Map"
	if pcv_strHtags="1" then
		StringBuilderObj.append "</h1>" & vbcrlf
	else
		StringBuilderObj.append "</strong></span><br><br>" & vbcrlf
	end if
	
	query="SELECT idcategory,idParentCategory,categorydesc,sdesc FROM categories where iBTOHide<>1 order by categorydesc asc"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcCatArr=rs.GetRows()
		set rs=nothing
		CatRecords=ubound(pcCatArr,2)
		Call LoopCats(1,1)
	end if
	set rs=nothing
	
End Sub	
'/////////////////////////////////////////////////////////////////////////////////////////
'// END: SUB - GenSMBody
'/////////////////////////////////////////////////////////////////////////////////////////

%>
<!--#include file="AdminFooter.asp"-->
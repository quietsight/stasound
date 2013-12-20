<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="catcommon.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="inc_UpdateDates.asp" -->

<%
on error resume next
Server.ScriptTimeout = 5400

dim CSVRecord(100),f, query, conntemp, rstemp, rstemp1,TopRecord(100)

Function checkParent(ParentName,ParentID)
Dim rs,query,tmp1,tmpCatName
	tmp1=0
	if ParentName<>"" then
		tmpCatName=ParentName
		tmpCatName=replace(tmpCatName,"&quot;","""")
		tmpCatName=replace(tmpCatName,"&amp;","&")
		tmpCatName=replace(tmpCatName,"&","&amp;")
		tmpCatName=replace(tmpCatName,"""","&quot;")
	end if
	if ParentID<>"" and ParentID<>"0" then
		query="SELECT IDCategory FROM Categories WHERE IDCategory=" & ParentID & ";"
	else
		query="SELECT IDCategory FROM Categories WHERE categoryDesc like '" & ParentName & "' OR categoryDesc like '" & tmpCatName & "';"
	end if
	set rs=connTemp.execute(query)
	if not rs.eof then
		tmp1=rs("IDCategory")
	end if
	set rs=nothing
	checkParent=tmp1
End Function

Function checkFeaturedCat(tmpCatN,tmpCatID,ParentID)
Dim rs,query,tmp1,tmpCatName
	tmp1=0
	if tmpCatN<>"" then
		tmpCatName=tmpCatN
		tmpCatName=replace(tmpCatName,"&quot;","""")
		tmpCatName=replace(tmpCatName,"&amp;","&")
		tmpCatName=replace(tmpCatName,"&","&amp;")
		tmpCatName=replace(tmpCatName,"""","&quot;")
	end if
	if tmpCatID<>"" and tmpCatID<>"0" then
		query="SELECT IDCategory FROM Categories WHERE IDCategory=" & tmpCatID & " AND idParentCategory=" & ParentID & ";"
	else
		query="SELECT IDCategory FROM Categories WHERE (categoryDesc like '" & tmpCatN & "' OR categoryDesc like '" & tmpCatName & "') AND idParentCategory=" & ParentID & ";"
	end if
	set rs=connTemp.execute(query)
	if not rs.eof then
		tmp1=rs("IDCategory")
	end if
	set rs=nothing
	checkFeaturedCat=tmp1
End Function

call openDb()
	Append=session("append")
	FileCSV = "../pc/catalog/" & session("importfile")
	if PPD="1" then
		FileCSV="/"&scPcFolder&"/pc/catalog/"&session("importfile")
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
	TotalCSVlines=0
	ImportedRecords=0
	fields=session("totalfields")
	TopLine=f.Readline
		
	'Get previous information before import/update Categories
	query="Select IDCategory from Categories order by IDCategory desc;"
	set rstemp4=connTemp.execute(query)
	
	if not rstemp4.eof then
	PreIDCategory="" & rstemp4("IDCategory")
	else
	PreIDCategory="0"
	end if
	
	if session("append")="1" then
	UpdateType="UPDATE"
	else
	UpdateType="IMPORT"
	end if
	PreRecords=""
	
	CategoryUpdateError=0
	
	Do While not f.AtEndofStream
	
	TempCSV=f.Readline
	TempCSV=replace(TempCSV,chr(34) & chr(34),"&quot;")
	TempCSV=replace(TempCSV,"'","''")
	A=split(TempCSV,",")
	RecordError=false
	TotalCSVlines=TotalCSVlines+1

		'Maybe Category Text Fields have commas or "End Text Line - VBCrLf" characters, fix these problems
	 	i=0
	 	j=0
	 	Do while j<fields
	 	if i>ubound(a) then
			ErrorsReport=ErrorsReport & "Record " & TotalCSVlines & ": Does not have enough data fields." & vbcrlf
			RecordError=true
			exit do
		end if
	 	if left(A(i),1)= chr(34) then
	 		if len(a(i))>1 then
	 		CSVRecord(j)=mid(A(i),2,len(A(i)))
	 		else
	 		CSVRecord(j)=""
	 		end if
	 		if (right(A(i),1)=chr(34)) and (len(a(i))>1) then
	 			CSVRecord(j) = mid(CSVRecord(j),1,len(CSVRecord(j))-1)
	 		else
	 			Do
	 				i=i+1
	 				if i<=ubound(a) then 
	 					CSVRecord(j)=CSVRecord(j) & "," & A(i)
	 				else
	 				Do
	 					Templine=f.ReadLine
	 					Templine=replace(Templine,chr(34) & chr(34),"&quot;")
	 					Templine=replace(Templine,"'","''")
	 					check=instr(Templine,chr(34) & ",")
	 					if check=0 then
	 						CSVRecord(j)=CSVRecord(j) & vbcrlf & Templine
	 					end if
	 				Loop Until (check>0)
	 				A=split(Templine,",")
	 				i=0
	 				CSVRecord(j)=CSVRecord(j) & vbcrlf & A(i)
	 				end if
	 			Loop Until right(A(i),1)=chr(34)
	 			CSVRecord(j)=mid(CSVRecord(j),1,len(CSVRecord(j))-1)
	 		end if
	 		i=i+1
	 	else
		 	CSVRecord(j)=A(i)	
		 	i=i+1
		end if
		j=j+1
		Loop 	
	 	
if RecordError=False then%>
	<!--#include file="catcommon1.asp"-->
<%end if%>
<%
if RecordError=false then 'STEP 1
	newCatName=""
	if catName<>"" then
		newCatName=catName
		newCatName=replace(newCatName,"&quot;","""")
		newCatName=replace(newCatName,"&amp;","&")
		newCatName=replace(newCatName,"&","&amp;")
		newCatName=replace(newCatName,"""","&quot;")
	end if

	if session("append")="1" then
		if catID<>"0" and catID<>"" then
			query="Select IDCategory from Categories WHERE IDCategory=" & catID & ";"
		else
			query="Select IDCategory from Categories WHERE categoryDesc like '" & catName & "' OR categoryDesc like '" & newCatName & "';"
		end if
	else
		query="Select IDCategory from Categories WHERE categoryDesc like '" & catName & "' OR categoryDesc like '" & newCatName & "';"
	end if
	set rstemp4=connTemp.execute(query)
	testCat=0
	if not rstemp4.eof then
		testCat=1
		IDCategory=rstemp4("idCategory")
		catID=IDCategory
	end if
	
	pAppend=0
	IF session("append")="1" then
		IF testCat=1 then 'EXISTING Category

			temp4=""
			
			tmpIDParent=0
			if catParentNameID>-1 OR catParentIDID>-1 then
				tmpIDParent=checkParent(catParentName,catParentID)
				if tmpIDParent=0 then
							tmpIDParent=1
				end if
				temp4=temp4 & ",idParentCategory=" & tmpIDParent
			end if
			
			if catNameID>-1 then
				temp4=temp4 & ",categoryDesc='" & newCatName & "'"
			end if
			
			if catOrderID>-1 then
				temp4=temp4 & ",priority=" & catOrder
			end if
			
			if catSImgID>-1 then
				temp4=temp4 & ",image='" & catSImg & "'"
			end if
			
			if catLImgID>-1 then
				temp4=temp4 & ",largeimage='" & catLImg & "'"
			end if
			
			if catHideCatID>-1 then
				temp4=temp4 & ",iBTOhide=" & catHideCat
			end if
			
			if catSDescID>-1 then
				temp4=temp4 & ",SDesc='" & catSDesc & "'"
			end if
			
			if catLDescID>-1 then
				temp4=temp4 & ",LDesc='" & catLDesc & "'"
			end if
			
			if catNotShowDescID>-1 then
				temp4=temp4 & ",HideDesc=" & catNotShowDesc
			end if
			
			if catHideCatRetailID>-1 then
				temp4=temp4 & ",pccats_RetailHide=" & catHideCatRetail
			end if
			
			if catDisplayCatsID>-1 then
				temp4=temp4 & ",pcCats_SubCategoryView=" & catDisplayCats
			end if
			
			if catCatsPerRowID>-1 then
				temp4=temp4 & ",pcCats_CategoryColumns=" & catCatsPerRow
			end if
			
			if catCatRowsID>-1 then
				temp4=temp4 & ",pcCats_CategoryRows=" & catCatRows
			end if
			
			if catDisplayPrdsID>-1 then
				temp4=temp4 & ",pcCats_PageStyle='" & catDisplayPrds & "'"
			end if
			
			if catPrdsPerRowID>-1 then
				temp4=temp4 & ",pcCats_ProductColumns=" & catPrdsPerRow
			end if
			
			if catPrdRowsID>-1 then
				temp4=temp4 & ",pcCats_ProductRows=" & catPrdRows
			end if
			
			tmpFeaturedCat=0
			if catFeaturedNameID>-1 OR catFeaturedIDID>-1 then
				tmpFeaturedCat=checkFeaturedCat(catFeaturedName,catFeaturedID,catID)
				temp4=temp4 & ",pcCats_FeaturedCategory=" & tmpFeaturedCat
			end if
						
			if catUseCatImgID>-1 then
				temp4=temp4 & ",pcCats_FeaturedCategoryImage=" & catUseCatImg
			end if
			
			if catPrdDisplayOptID>-1 then
				temp4=temp4 & ",pcCats_DisplayLayout='" & catPrdDisplayOpt & "'"
			end if
			
			if catMTTitleID>-1 then
				temp4=temp4 & ",pcCats_MetaTitle='" & catMTTitle & "'"
			end if
			
			if catMTDescID>-1 then
				temp4=temp4 & ",pcCats_MetaDesc='" & catMTDesc & "'"
			end if
			
			if catMTKeyID>-1 then
				temp4=temp4 & ",pcCats_MetaKeywords='" & catMTKey & "'"
			end if
						
			'Get Category information before update
			query="select * from Categories where IDCategory=" & catID & ";"
			set rstemp=conntemp.execute(query)
			
			IF not rstemp.eof THEN

			PreRecord1=""
			PreRecord1=PreRecord1 & rstemp("idCategory") & "****"
			
			iCols = rstemp.Fields.Count
		    for dd=1 to iCols-1
		    FType="" & Rstemp.Fields.Item(dd).Type
		    if (Rstemp.Fields.Item(dd).Name="dtRewardsStarted") then
		    FType="DYDL"
		    end if
		    if (Ftype="202") or (Ftype="203") or (FType="DYDL") then
		    PTemp=Rstemp.Fields.Item(dd).Value
		    if PTemp<>"" then
		    PTemp=replace(PTemp,"'","''")
		    PTemp=replace(PTemp,vbcrlf,"DuLTVDu")
		    end if
		    if FType="DYDL" then
		    if scDB="Access" then
		    myStr11="#"
		    else
		    myStr11="'"
		    end if
		    else
		    myStr11="'"
		    end if
		    if dd=1 then
		    PreRecord1=PreRecord1 & myStr11 & PTemp & myStr11
		    else
		    PreRecord1=PreRecord1 & "@@@@@" & myStr11 & PTemp & myStr11
		    end if
		    else
		    PTemp="" & Rstemp.Fields.Item(dd).Value
		    if PTemp<>"" then
		    else
		    PTemp="0"
		    end if
		    if dd=1 then
		    PreRecord1=PreRecord1 & PTemp
		    else
		    PreRecord1=PreRecord1 & "@@@@@" & PTemp
		    end if
		    end if
			next
			PreRecords=PreRecords & PreRecord1 & vbcrlf
			END IF

			if temp4<>"" then
				temp4=mid(temp4,2,len(temp4))
				query="update Categories set " & temp4 & " where IDCategory=" & catID & ";"
				'query=replace(query,chr(34),"&quot;")
				'query=replace(query,"**DD**",chr(34))
				set rstemp=conntemp.execute(query)
				pAppend=1
				
				call updCatEditedDate(catID,"")
			end if
				
		ELSE 'Do not have existing Category
			CategoryUpdateError=1
			if catName<>"" then
				ErrorsReport=ErrorsReport & "Record " & TotalCSVlines & ": The Category Name: '" & catName & "' is not in the database." & vbcrlf
			else
				ErrorsReport=ErrorsReport & "Record " & TotalCSVlines & ": The Category ID: " & catID & " is not in the database." & vbcrlf
			end if
			RecordError=true
		END IF

	ELSE 'Append=0
		tmpIDParent=0
		if catParentNameID>-1 OR catParentIDID>-1 then
			tmpIDParent=checkParent(catParentName,catParentID)
		end if
		if tmpIDParent=0 then
			tmpIDParent=1
		end if
		if testCat=1 then
			query="Select IDCategory from Categories WHERE (categoryDesc like '" & catName & "' OR categoryDesc like '" & newCatName & "') AND idParentCategory=" & tmpIDParent & ";"
			set rs=connTemp.execute(query)
			if not rs.eof then
				ErrorsReport=ErrorsReport & "Record " & TotalCSVlines & ": The Category Name: '" & catName & "' could not be imported because it already exists in the same parent category." & vbcrlf
				RecordError=true
			end if
			set rs=nothing
		end if
		if RecordError=false then
			query="INSERT INTO Categories (idParentCategory,categoryDesc,priority,image,largeimage,iBTOhide,SDesc,LDesc,HideDesc,pccats_RetailHide,pcCats_SubCategoryView,pcCats_CategoryColumns,pcCats_CategoryRows,pcCats_PageStyle,pcCats_ProductColumns,pcCats_ProductRows,pcCats_FeaturedCategory,pcCats_FeaturedCategoryImage,pcCats_DisplayLayout,pcCats_MetaTitle,pcCats_MetaDesc,pcCats_MetaKeywords) VALUES "
			query=query & "(" & tmpIDParent & ",'" & newCatName & "'," & catOrder & ",'" & catSImg & "','" & catLImg & "'," & catHideCat & ",'" & catSDesc & "','" & catLDesc & "'," & catNotShowDesc & "," & catHideCatRetail & "," & catDisplayCats & "," & catCatsPerRow & "," & catCatRows & ",'" & catDisplayPrds & "'," & catPrdsPerRow & "," & catPrdRows & "," & catFeaturedID & "," & catUseCatImg & ",'" & catPrdDisplayOpt & "','" & catMTTitle & "','" & catMTDesc & "','" & catMTKey & "');"
			'query=replace(query,chr(34),"&quot;")
			'query=replace(query,"**DD**",chr(34))
			set rstemp=conntemp.execute(query)
			
			query="SELECT IDCategory FROM Categories WHERE categoryDesc='" & newCatName & "' ORDER BY IDCategory Desc;"
			set rstemp=conntemp.execute(query)
			pIdCategory = rstemp("idCategory")
			set rstemp=nothing
			
			call updCatCreatedDate(pIdCategory,"")
			
		end if
		
	END IF 'Update/Import
	
	IF RecordError=false THEN
		if session("append")="1" then
			query="SELECT IDCategory FROM Categories WHERE IDCategory=" & catID & ";"
		else
			query="SELECT IDCategory FROM Categories WHERE categoryDesc='" & newCatName & "' ORDER BY IDCategory Desc;"
		end if
		set rstemp=conntemp.execute(query)
		pIdCategory = rstemp("idCategory")
		set rstemp=nothing
		
		'--------------------------------------------------------------
		' START - Update breadcrumb navigation in case the category was imported/moved
		'--------------------------------------------------------------
		redim arrCategories(999,4)
		indexCategories=0
		pUrlString=Cstr("")
		pIdCategory2=pidCategory

		' load category array with all categories until parent
		do while pIdCategory2>1
			query="SELECT categoryDesc, idCategory, idParentcategory, largeimage, SDesc, LDesc, HideDesc FROM categories WHERE idCategory=" & pIdCategory2 &" ORDER BY priority, categoryDesc ASC"
			SET rs=Server.CreateObject("ADODB.RecordSet")
			SET rs=conntemp.execute(query)

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
 
			if rs.eof then
				set rs=nothing
				call closeDb()
				response.redirect "msg.asp?message=86"           
			end if
			
			'categoryDesc, idCategory, idParentcategory, largeimage, SDesc, LDesc, HideDesc
			if pIdCategory2=pidCategory then
				pCategoryName=rs("categoryDesc")
				intIdCategory=rs("idCategory")
				intIdParentCategory=rs("idParentCategory")
				plargeImage=rs("largeimage")
				if pLargeImage = "no_image.gif" then
					pLargeImage = ""
				end if
				SDesc=rs("SDesc")
				LDesc=rs("LDesc")
				HideDesc=rs("HideDesc")
				if isNULL(HideDesc) OR HideDesc="" then
					HideDesc="0"
				end if
			else
				pCategoryName=rs("categoryDesc")
				intIdCategory=rs("idCategory")
				intIdParentCategory=rs("idParentCategory")
			end if
			
			pIdCategory3=intIdParentCategory 
			arrCategories(indexCategories,0)=pCategoryName
			arrCategories(indexCategories,1)=intIdCategory
			arrCategories(indexCategories,2)=intIdParentCategory
			pIdCategory2=pIdCategory3
			indexCategories=indexCategories + 1   
		loop
		set rs=nothing
		
		'create new breadcrumb and enter it into database
		strDBBreadCrumb=""
		for f1=indexCategories-1 to 0 step -1
			If arrCategories(f1,2)="1" Then
				strDBBreadCrumb=strDBBreadCrumb&arrCategories(f1,1)&"||"&arrCategories(f1,0)
			Else
				strDBBreadCrumb=strDBBreadCrumb&"|,|"&arrCategories(f1,1)&"||"&arrCategories(f1,0)
			End If
		next
		'enter BreadCrumb into database
		query="UPDATE categories SET pccats_BreadCrumbs='"&replace(strDBBreadCrumb,"'","''")&"' WHERE idCategory="&pIdCategory&";"
		SET rs=Server.CreateObject("ADODB.RecordSet")
		SET rs=conntemp.execute(query)
		'--------------------------------------------------------------
		' END - Update breadcrumb
		'--------------------------------------------------------------
	END IF	
	
	if RecordError=false then
		ImportedRecords=ImportedRecords+1
	end if
	
end if 'END STEP 1
	
	Loop

	f.Close
	Set f = nothing
	
	'Delete Import File
	'Set fso = server.CreateObject("Scripting.FileSystemObject")
	'Set f = fso.GetFile(Server.MapPath(FileCSV))
	'f.Delete
	'Set fso = nothing
	'Set f = nothing
	
	'Update Category Tree XML Cache
	%>
    <!--#include file="inc_genCatXML.asp"-->
    <%
	
	call closeDB()
	
	if ImportedRecords>0 then
	
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	Set afi=fs.CreateTextFile(server.MapPath(".") & "\importlogs\categorylogs.txt",True)
		
	afi.Writeline(UpdateType)
	afi.Writeline(PreIDCategory)
	afi.Writeline(PreRecords)
	afi.Close
	
	end if
	
	session("importfile")=""
	session("totalfields")=0
	
	if CategoryUpdateError=1 then
	ErrorsReport="One of the records you are importing does not currently exist in the database. The Update feature is strictly for modifying existing Category information. Please correct the error and try again." &vbcrlf&vbcrlf &ErrorsReport
	end if


if session("append")="1" then 
	pageTitle = "UPDATE"
else
	pageTitle = "IMPORT" 
end if 
pageTitle = pageTitle & " CATEGORY DATA WIZARD - Review Import Results"
section = "products" %>
<!--#include file="AdminHeader.asp"-->
<script type="text/javascript" language="javascript" src="../includes/spry/SpryDOMUtils.js"></script>
<style type="text/css">
<!--
.grayBG {
	background-color: #F5F5F5;
}
-->
</style>
<table class="pcCPcontent">
<tr>
	<td valign="top">
        <table class="pcCPcontent">
        <tr>
            <td colspan="2"><h2>Steps:</h2></td>
        </tr>
        <tr>
            <td width="5%" align="right"><img border="0" src="images/step1.gif"></td>
            <td width="95%"><font color="#A8A8A8">Select category data file</font></td>
        </tr>
        <tr>
            <td align="right"><img border="0" src="images/step2.gif"></td>
            <td><font color="#A8A8A8">Map fields</font></td>
        </tr>
        <tr>
            <td align="right"><img border="0" src="images/step3.gif"></td>
            <td><font color="#A8A8A8">Confirm mapping</font></td>
        </tr>
        <tr>
            <td align="right"><img border="0" src="images/step4a.gif"></td>
            <td><strong><%if session("append")="1" then%>Update<%else%>Import<%end if%> results</strong></td>
        </tr>
        </table>
		
        <div class="pcCPmessageSuccess">
		Total <b><%=ImportedRecords%></b> records were <%if session("append")="1" then%>updated<%else%>imported<%end if%> successfully!
        </div>
        
		<%if TotalCSVlines-ImportedRecords>0 then%>
        	<div class="pcCPmessage">
			Total <b><%=TotalCSVlines-ImportedRecords%></b> records could not be <%if session("append")="1" then%>updated<%else%>imported<%end if%> successfully!
            </div>
		<%end if%>

		<% if ErrorsReport<>"" then%>
		<table class="pcCPcontent">
            <tr>
                <th>
                    Error(s) Report
                </th>
            </tr>
	        <tr>
	            <td align="center">
	                <div style="width: 98%; height: 150px; overflow: scroll; border: 1px dotted #E1E1E1; margin-top: 10px;">
	                    <table id="noheaderodd" style="font-family: Arial; font-size: 9px; width: 100%; text-align: left">
	                        <%=ErrorsReport%>
	                    </table>
	                </div>
	            </td>
	        </tr>
		</table>
		<% end if %>
		<br /><br />
		<input type="button" name="mainmenu" value="Manage Categories" onClick="location='manageCategories.asp';"  class="ibtnGrey">
	</td>
</tr>
</table>
<%session("append")=0%>
<!--#include file="AdminFooter.asp"-->
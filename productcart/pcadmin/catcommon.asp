<%
catNameID=-1
catIDID=-1
catSImgID=-1
catLImgID=-1
catParentNameID=-1
catParentIDID=-1
catSDescID=-1
catLDescID=-1
catNotShowDescID=-1
catDisplayCatsID=-1
catCatsPerRowID=-1
catCatRowsID=-1
catUseCatImgID=-1
catDisplayPrdsID=-1
catPrdsPerRowID=-1
catPrdRowsID=-1
catHideCatID=-1
catHideCatRetailID=-1
catPrdDisplayOptID=-1
catMTTitleID=-1
catMTDescID=-1
catMTKeyID=-1
catFeaturedNameID=-1
catFeaturedIDID=-1
catOrderID=-1

catName=""
catID=0
catSImg=""
catLImg=""
catParentName=""
catParentID=0
catSDesc=""
catLDesc=""
catNotShowDesc=0
catDisplayCats=0
catCatsPerRow=0
catCatRows=0
catUseCatImg=0
catDisplayPrds=""
catPrdsPerRow=0
catPrdRows=0
catHideCat=0
catHideCatRetail=0
catPrdDisplayOpt=""
catMTTitle=""
catMTDesc=""
catMTKey=""
catFeaturedName=""
catFeaturedID=0
catOrder=0

TempCategories=""
ErrorsReport=""%>

<!--#include file="catcheckfields.asp"-->

<%

For i=1 to request("validfields")
	
	BLine=0
	Select Case request("T" & i)
	Case "Category Name": catNameID=request("P" & i)
	BLine=1
	Case "Category ID": catIDID=request("P" & i)
	BLine=2
	Case "Small Image": catSImgID=request("P" & i)
	BLine=3
	Case "Large Image": catLImgID=request("P" & i)
	BLine=4
	Case "Parent Category Name": catParentNameID=request("P" & i)
	BLine=5
	Case "Parent Category ID": catParentIDID=request("P" & i)
	BLine=6
	Case "Category Short Description": catSDescID=request("P" & i)
	BLine=7
	Case "Category Long Description": catLDescID=request("P" & i)
	BLine=8
	Case "Hide Category Description": catNotShowDescID=request("P" & i)
	BLine=9
	Case "Display Sub-Categories": catDisplayCatsID=request("P" & i)
	BLine=10
	Case "Sub-Categories per Row": catCatsPerRowID=request("P" & i)
	BLine=11
	Case "Sub-Category Rows per Page": catCatRowsID=request("P" & i)
	BLine=12
	Case "Use Featured Sub-Category Image": catUseCatImgID=request("P" & i)
	BLine=13
	Case "Display Products": catDisplayPrdsID=request("P" & i)
	BLine=14
	Case "Products per Row": catPrdsPerRowID=request("P" & i)
	BLine=15
	Case "Product Rows per Page": catPrdRowsID=request("P" & i)
	BLine=16
	Case "Hide category": catHideCatID=request("P" & i)
	BLine=17
	Case "Hide category from retail customers": catHideCatRetailID=request("P" & i)
	BLine=18
	Case "Product Details Page Display Option": catPrdDisplayOptID=request("P" & i)
	BLine=19
	Case "Category Meta Tags - Title": catMTTitleID=request("P" & i)
	BLine=20
	Case "Category Meta Tags - Description": catMTDescID=request("P" & i)
	BLine=21
	Case "Category Meta Tags - Keywords": catMTKeyID=request("P" & i)
	BLine=22
	Case "Featured Sub-Category Name": catFeaturedNameID=request("P" & i)
	BLine=23
	Case "Featured Sub-Category ID": catFeaturedIDID=request("P" & i)
	BLine=24
	Case "Category Order": catOrderID=request("P" & i)
	BLine=25
	End Select
	
	if BLine>0 then
	TempStr=request("F" & i) & "*****"
	if instr(ALines(BLine-1),TempStr)=0 then
	ALines(BLine-1)=ALines(BLine-1) & TempStr
	end if
	BLine=0
	end if
Next

	SavedFile = "importlogs/catsave.txt"
	findit = Server.MapPath(Savedfile)
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	Err.number=0
	Set f = fso.OpenTextFile(findit, 2)
	For dd=lbound(ALines) to ubound(ALines)
	f.WriteLine ALines(dd)
	Next
	f.close
%>
<%
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

if catNameID<>-1 then
catName=trim(CSVRecord(catNameID))
end if

if catIDID<>-1 then
catID=trim(CSVRecord(catIDID))
end if

if catID="" then
	catID=0
end if

if Not IsNumeric(catID) then
	catID=0
end if

if catSImgID<>-1 then
catSImg=trim(CSVRecord(catSImgID))
end if

if catSImg="" then
	catSImg="no_image.gif"
end if

if catLImgID<>-1 then
catLImg=trim(CSVRecord(catLImgID))
end if

if catLImg="" then
	catLImg="no_image.gif"
end if

if catParentNameID<>-1 then
catParentName=trim(CSVRecord(catParentNameID))
end if

if catParentIDID<>-1 then
catParentID=trim(CSVRecord(catParentIDID))
end if

if catParentID="" then
	catParentID=0
end if

if Not IsNumeric(catParentID) then
	catParentID=0
end if

if catSDescID<>-1 then
catSDesc=trim(CSVRecord(catSDescID))
end if

if catLDescID<>-1 then
catLDesc=trim(CSVRecord(catLDescID))
end if

if catNotShowDescID<>-1 then
catNotShowDesc=trim(CSVRecord(catNotShowDescID))
end if

if catNotShowDesc="" then
	catNotShowDesc=0
end if

if Not IsNumeric(catNotShowDesc) then
	catNotShowDesc=0
end if

if catDisplayCatsID<>-1 then
catDisplayCats=trim(CSVRecord(catDisplayCatsID))
end if

if catDisplayCats="" then
	catDisplayCats=0
end if

if Not IsNumeric(catDisplayCats) then
	catDisplayCats=0
end if

if catCatsPerRowID<>-1 then
catCatsPerRow=trim(CSVRecord(catCatsPerRowID))
end if

if catCatsPerRow="" then
	catCatsPerRow=0
end if

if Not IsNumeric(catCatsPerRow) then
	catCatsPerRow=0
end if

if catCatRowsID<>-1 then
catCatRows=trim(CSVRecord(catCatRowsID))
end if

if catCatRows="" then
	catCatRows=0
end if

if Not IsNumeric(catCatRows) then
	catCatRows=0
end if

if catUseCatImgID<>-1 then
catUseCatImg=trim(CSVRecord(catUseCatImgID))
end if

if catUseCatImg="" then
	catUseCatImg=0
end if

if Not IsNumeric(catUseCatImg) then
	catUseCatImg=0
end if

if catDisplayPrdsID<>-1 then
catDisplayPrds=trim(CSVRecord(catDisplayPrdsID))
end if

if catPrdsPerRowID<>-1 then
catPrdsPerRow=trim(CSVRecord(catPrdsPerRowID))
end if

if catPrdsPerRow="" then
	catPrdsPerRow=0
end if

if Not IsNumeric(catPrdsPerRow) then
	catPrdsPerRow=0
end if

if catPrdRowsID<>-1 then
catPrdRows=trim(CSVRecord(catPrdRowsID))
end if

if catPrdRows="" then
	catPrdRows=0
end if

if Not IsNumeric(catPrdRows) then
	catPrdRows=0
end if

if catHideCatID<>-1 then
catHideCat=trim(CSVRecord(catHideCatID))
end if

if catHideCat="" then
	catHideCat=0
end if

if Not IsNumeric(catHideCat) then
	catHideCat=0
end if

if catHideCatRetailID<>-1 then
catHideCatRetail=trim(CSVRecord(catHideCatRetailID))
end if

if catHideCatRetail="" then
	catHideCatRetail=0
end if

if Not IsNumeric(catHideCatRetail) then
	catHideCatRetail=0
end if

if catPrdDisplayOptID<>-1 then
catPrdDisplayOpt=trim(CSVRecord(catPrdDisplayOptID))
end if

if catMTTitleID<>-1 then
catMTTitle=trim(CSVRecord(catMTTitleID))
end if

if catMTDescID<>-1 then
catMTDesc=trim(CSVRecord(catMTDescID))
end if

if catMTKeyID<>-1 then
catMTKey=trim(CSVRecord(catMTKeyID))
end if

if catFeaturedNameID<>-1 then
catFeaturedName=trim(CSVRecord(catFeaturedNameID))
end if

if catFeaturedIDID<>-1 then
catFeaturedID=trim(CSVRecord(catFeaturedIDID))
end if

if catFeaturedID="" then
	catFeaturedID=0
end if

if Not IsNumeric(catFeaturedID) then
	catFeaturedID=0
end if

if catOrderID<>-1 then
catOrder=trim(CSVRecord(catOrderID))
end if

if catOrder<>"" then
else
	catOrder=0
end if

if Not IsNumeric(catOrder) then
	catOrder=0
end if
			
if session("append")="1" then
	if (Clng(catID)=0) AND (catName="") then
		ErrorsReport=ErrorsReport & "Record " & TotalCSVlines & ": does not contain a Category ID or Category Name." & vbcrlf
		RecordError=true
	end if
else		
	if catName="" then
		ErrorsReport=ErrorsReport & "Record " & TotalCSVlines & ": does not contain a Category Name." & vbcrlf
		RecordError=true
	end if
end if

if catName<>"" then
	catName=replace(catName,"'","''")
end if
if catParentName<>"" then
	catParentName=replace(catParentName,"'","''")
end if
if catSDesc<>"" then
	catSDesc=replace(catSDesc,"'","''")
end if
if catLDesc<>"" then
	catLDesc=replace(catLDesc,"'","''")
end if
if catMTTitle<>"" then
	catMTTitle=replace(catMTTitle,"'","''")
end if
if catMTDesc<>"" then
	catMTDesc=replace(catMTDesc,"'","''")
end if
if catMTKey<>"" then
	catMTKey=replace(catMTKey,"'","''")
end if
if catFeaturedName<>"" then
	catFeaturedName=replace(catFeaturedName,"'","''")
end if
%>
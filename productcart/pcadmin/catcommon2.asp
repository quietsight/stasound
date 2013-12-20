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
catName=trim(rsExcel.Fields.Item(int(catNameID)).Value)
end if

if catIDID<>-1 then
catID=trim(rsExcel.Fields.Item(int(catIDID)).Value)
end if

if catID<>"" then
else
	catID=0
end if

if Not IsNumeric(catID) then
	catID=0
end if

if catSImgID<>-1 then
catSImg=trim(rsExcel.Fields.Item(int(catSImgID)).Value)
end if

if catSImg<>"" then
else
	catSImg="no_image.gif"
end if

if catLImgID<>-1 then
catLImg=trim(rsExcel.Fields.Item(int(catLImgID)).Value)
end if

if catLImg<>"" then
else
	catLImg="no_image.gif"
end if

if catParentNameID<>-1 then
catParentName=trim(rsExcel.Fields.Item(int(catParentNameID)).Value)
end if

if catParentIDID<>-1 then
catParentID=trim(rsExcel.Fields.Item(int(catParentIDID)).Value)
end if

if catParentID<>"" then
else
	catParentID=0
end if

if Not IsNumeric(catParentID) then
	catParentID=0
end if

if catSDescID<>-1 then
catSDesc=trim(rsExcel.Fields.Item(int(catSDescID)).Value)
end if

if catLDescID<>-1 then
catLDesc=trim(rsExcel.Fields.Item(int(catLDescID)).Value)
end if

if catNotShowDescID<>-1 then
catNotShowDesc=trim(rsExcel.Fields.Item(int(catNotShowDescID)).Value)
end if

if catNotShowDesc<>"" then
else
	catNotShowDesc=0
end if

if Not IsNumeric(catNotShowDesc) then
	catNotShowDesc=0
end if

if catDisplayCatsID<>-1 then
catDisplayCats=trim(rsExcel.Fields.Item(int(catDisplayCatsID)).Value)
end if

if catDisplayCats<>"" then
else
	catDisplayCats=0
end if

if Not IsNumeric(catDisplayCats) then
	catDisplayCats=0
end if

if catCatsPerRowID<>-1 then
catCatsPerRow=trim(rsExcel.Fields.Item(int(catCatsPerRowID)).Value)
end if

if catCatsPerRow<>"" then
else
	catCatsPerRow=0
end if

if Not IsNumeric(catCatsPerRow) then
	catCatsPerRow=0
end if

if catCatRowsID<>-1 then
catCatRows=trim(rsExcel.Fields.Item(int(catCatRowsID)).Value)
end if

if catCatRows<>"" then
else
	catCatRows=0
end if

if Not IsNumeric(catCatRows) then
	catCatRows=0
end if

if catUseCatImgID<>-1 then
catUseCatImg=trim(rsExcel.Fields.Item(int(catUseCatImgID)).Value)
end if

if catUseCatImg<>"" then
else
	catUseCatImg=0
end if

if Not IsNumeric(catUseCatImg) then
	catUseCatImg=0
end if

if catDisplayPrdsID<>-1 then
catDisplayPrds=trim(rsExcel.Fields.Item(int(catDisplayPrdsID)).Value)
end if

if catPrdsPerRowID<>-1 then
catPrdsPerRow=trim(rsExcel.Fields.Item(int(catPrdsPerRowID)).Value)
end if

if catPrdsPerRow<>"" then
else
	catPrdsPerRow=0
end if

if Not IsNumeric(catPrdsPerRow) then
	catPrdsPerRow=0
end if

if catPrdRowsID<>-1 then
catPrdRows=trim(rsExcel.Fields.Item(int(catPrdRowsID)).Value)
end if

if catPrdRows<>"" then
else
	catPrdRows=0
end if

if Not IsNumeric(catPrdRows) then
	catPrdRows=0
end if

if catHideCatID<>-1 then
catHideCat=trim(rsExcel.Fields.Item(int(catHideCatID)).Value)
end if

if catHideCat<>"" then
else
	catHideCat=0
end if

if Not IsNumeric(catHideCat) then
	catHideCat=0
end if

if catHideCatRetailID<>-1 then
catHideCatRetail=trim(rsExcel.Fields.Item(int(catHideCatRetailID)).Value)
end if

if catHideCatRetail<>"" then
else
	catHideCatRetail=0
end if

if Not IsNumeric(catHideCatRetail) then
	catHideCatRetail=0
end if

if catPrdDisplayOptID<>-1 then
catPrdDisplayOpt=trim(rsExcel.Fields.Item(int(catPrdDisplayOptID)).Value)
end if

if catMTTitleID<>-1 then
catMTTitle=trim(rsExcel.Fields.Item(int(catMTTitleID)).Value)
end if

if catMTDescID<>-1 then
catMTDesc=trim(rsExcel.Fields.Item(int(catMTDescID)).Value)
end if

if catMTKeyID<>-1 then
catMTKey=trim(rsExcel.Fields.Item(int(catMTKeyID)).Value)
end if

if catFeaturedNameID<>-1 then
catFeaturedName=trim(rsExcel.Fields.Item(int(catFeaturedNameID)).Value)
end if

if catFeaturedIDID<>-1 then
catFeaturedID=trim(rsExcel.Fields.Item(int(catFeaturedIDID)).Value)
end if

if catFeaturedID<>"" then
else
	catFeaturedID=0
end if

if Not IsNumeric(catFeaturedID) then
	catFeaturedID=0
end if

if catOrderID<>-1 then
catOrder=trim(rsExcel.Fields.Item(int(catOrderID)).Value)
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
		ErrorsReport=ErrorsReport & "Record " & TotalXLSlines & ": does not contain a Category ID or Category Name." & vbcrlf
		RecordError=true
	end if
else		
	if catName="" then
		ErrorsReport=ErrorsReport & "Record " & TotalXLSlines & ": does not contain a Category Name." & vbcrlf
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
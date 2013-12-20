<%
	'// If category ID doesn't exist, get the first category that the product has been assigned to, filtering out hidden categories
	if pIdCategory=0 or trim(pIdCategory)="" then
		query="SELECT categories_products.idCategory FROM categories_products INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE categories_products.idProduct="& pIdProduct &" AND categories.iBTOhide<>1 AND categories.pccats_RetailHide<>1"
		set rsCat=server.CreateObject("ADODB.RecordSet")
		set rsCat=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rsCat=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
		if not rsCat.EOF then
			pIdCategoryTemp=rsCat("idCategory")
		else
			pIdCategoryTemp=1
		end if
		set rsCat=nothing
	end if
%>
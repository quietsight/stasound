<%
Dim pcv_cats1,pcv_cats2,pcv_cats3
Dim pcv_tmpParent,pcv_strTmpCat
Dim pcv_havecats1,pcv_havecats2

pcv_havecats1=0
pcv_havecats2=0

pcv_strTmpCat=""

Sub pcs_CatList()
Dim tmp_A, tmp_B, k, l , m,tmp1,tmp2

IF cat_DropDownName="" then
	cat_DropDownName="idcategory"
END IF
IF cat_Type="" then
	cat_Type="0"
END IF
IF cat_DropDownSize="" then
	cat_DropDownSize="1"
END IF
IF cat_MultiSelect="" then
	cat_MultiSelect="0"
END IF
IF cat_ExcBTOHide="" then
	cat_ExcBTOHide="0"
END IF
IF cat_StoreFront="" then
	cat_StoreFront="0"
END IF
IF cat_ShowParent="" then
	cat_ShowParent="0"
END IF
IF cat_ExcBTOItems="" then
	cat_ExcBTOItems="0"
END IF

IF NOT validNum(cat_CurrentCategory) then
	cat_CurrentCategory=Cint(1)
END IF

'cat_DefaultItem
'cat_EventAction

tmp_A=split(cat_SelectedItems,",")
tmp_B=split(cat_ExcItems,",")

IF cat_ExcSubs="" then
	cat_ExcSubs="0"
END IF

call opendb()

query="SELECT idCategory,categoryDesc,idParentCategory,pccats_BreadCrumbs FROM categories "
query1=""
if cat_ExcBTOHide="1" THEN
	query1=query1 & "categories.iBTOhide=0"
end if
if cat_StoreFront="1" then
	if session("customerType")="1" then
	else
		if query1<>"" then
			query1=	query1 & " AND "
		end if
		query1= query1 & "categories.pccats_RetailHide=0"
	end if
end if
if query1<>"" then
	query=query & " WHERE " & query1
end if
if cat_ExcSubs="1" then
	if query1<>"" then
		query=query & " AND idParentCategory=" & cat_CurrentCategory & " ORDER BY categories.categoryDesc ASC;"
	else
		query=query & " WHERE idParentCategory=" & cat_CurrentCategory & " ORDER BY categories.categoryDesc ASC;"
	end if
else
	query=query & " ORDER BY categories.categoryDesc ASC;"
end if

if pcv_havecats1=0 then
	set rsC=connTemp.execute(query)
	if not rsC.eof then
		pcv_cats1=rsC.GetRows()
		pcv_havecats1=1
	end if
	set rsC=nothing
end if

IF cat_Type="1" THEN
	if cat_ExcBTOItems="1" then
		tmpStrEx=" categories.idcategory IN (SELECT DISTINCT categories_products.idcategory FROM products INNER JOIN categories_products ON (products.idproduct=categories_products.idproduct AND products.configOnly=0) GROUP by (categories_products.idcategory)) "
	else
		tmpStrEx=" categories.idcategory IN (SELECT DISTINCT categories_products.idcategory FROM categories_products GROUP by (categories_products.idcategory)) "
	end if
	
	if scDB="Access" then
	query="SELECT DISTINCT categories.idCategory,categories.categoryDesc,categories.idParentCategory,categories.pccats_BreadCrumbs FROM categories WHERE " & tmpStrEx
	else
	query="SELECT DISTINCT categories.idCategory,categories.categoryDesc,categories.idParentCategory,cast(categories.pccats_BreadCrumbs as varchar(8000)) pccats_BreadCrumbs FROM categories WHERE " & tmpStrEx
	end if 
	
	if query1<>"" then
		if Left(query1,5)<>" AND " then
			query=query & " AND " & query1
		else
			query=query & query1
		end if
	end if
	if cat_ExcSubs="1" then
		query=query & " AND idParentCategory=" & cat_CurrentCategory & " ORDER BY categories.categoryDesc ASC;"
	else
		query=query & " ORDER BY categories.categoryDesc ASC;"
	end if
	if pcv_havecats2=0 then
		set rsC=connTemp.execute(query)
		if not rsC.eof then
			pcv_cats2=rsC.GetRows()
			pcv_havecats2=1
		end if
		set rsC=nothing
	end if
	pcv_cats3=pcv_cats2
ELSE
	pcv_cats3=pcv_cats1
END IF

IF IsNull(pcv_cats3) then
	response.write "Categories not found."
ELSE%>
	<select name="<%=cat_DropDownName%>" size="<%=cat_DropDownSize%>" <%if cat_MultiSelect="1" then%>multiple<%end if%> <%if cat_EventAction<>"" then%><%=cat_EventAction%><%end if%>>
		<%if cat_DefaultItem<>"" then%>
			<option value="0"><%=cat_DefaultItem%></option>
		<%end if%>
		<%if cat_DisplayRoot<>"" then%>
			<option value="1">&lt; No Parent - Top Level Category &gt;</option>
		<%end if%>
		<%For k=lbound(pcv_cats3,2) to ubound(pcv_cats3,2)
			if trim(pcv_cats3(0,k))<>"" then
			tmp1=0
			For l=lbound(tmp_B) to ubound(tmp_B)
				if trim(tmp_B(l))<>"" then
					if clng(tmp_B(l))=clng(pcv_cats3(0,k)) then
						tmp1=1
						exit for
					end if
					if tmp1=0 then
						if (clng(tmp_B(l))=clng(pcv_cats3(2,k))) and (cat_ExcSubs="1") then
							if pcv_strTmpCat="" then
								pcv_strTmpCat="****" & pcv_cats3(0,k) & "****"
							else
								pcv_strTmpCat=pcv_strTmpCat & pcv_cats3(0,k) & "****"
							end if
							tmp1=1
							exit for
						end if
					end if
					if tmp1=0 then
						if (Instr(pcv_strTmpCat,"****" & pcv_cats3(2,k) & "****")>0) and (cat_ExcSubs="1") then
							if pcv_strTmpCat="" then
								pcv_strTmpCat="****" & pcv_cats3(0,k) & "****"
							else
								pcv_strTmpCat=pcv_strTmpCat & pcv_cats3(0,k) & "****"
							end if
							tmp1=1
							exit for
						end if
					end if
				end if
			Next
			if tmp1=0 then
				tmp2=0
				For l=lbound(tmp_A) to ubound(tmp_A)
					if trim(tmp_A(l))<>"" then
					if clng(tmp_A(l))=clng(pcv_cats3(0,k)) then
						tmp2=1
						exit for
					end if
					end if
				Next%>
				<option value="<%=pcv_cats3(0,k)%>" <%if tmp2=1 then%>selected<%end if%>><%=ReplaceChars(pcv_cats3(1,k))%><%if cat_ShowParent="1" then%>&nbsp;<%=ReplaceChars(pcf_catGetParent(pcv_cats3(0,k),pcv_cats3(2,k),pcv_cats3(3,k)))%><%end if%></option>
			<%end if
			end if
		Next%>
	</select>
<%END IF

set pcv_cats1=nothing
set pcv_cats2=nothing
set pcv_cats3=nothing

End Sub

Sub pcf_genParent(m)
	if pcv_tmpParent<>"" then
		pcv_tmpParent="/" & pcv_tmpParent
	end if
	pcv_tmpParent=pcv_cats1(1,m) & pcv_tmpParent
	if pcv_cats1(2,m)<>"1" then
		pcf_FindParent(pcv_cats1(2,m))
	end if
End Sub

Sub pcf_FindParent(pcv_parentCategory)
Dim k,pcv_end,pcv_back
	pcv_end=ubound(pcv_cats1,2)
	For k=lbound(pcv_cats1,2) to ubound(pcv_cats1,2)
		pcv_back=pcv_end-k
		if pcv_parentCategory=pcv_cats1(0,k) then
			call pcf_genParent(k)
			exit for
		end if
		if pcv_parentCategory=pcv_cats1(0,pcv_back) then
			call pcf_genParent(pcv_back)
			exit for
		end if
	Next
End Sub

Function pcf_catGetParent(pcv_idcategory,pcv_parentCategory,pcv_BreadCrumbs)
Dim tmp_ParentText,tmp_C,tmp_D,p
	tmp_ParentText=""
	IF trim(pcv_BreadCrumbs)<>"" THEN
		tmp_C=split(pcv_BreadCrumbs,"|,|")
		For p=lbound(tmp_C) to ubound(tmp_C)
			if trim(tmp_C(p))<>"" then
			tmp_D=split(tmp_C(p),"||")
			if (clng(tmp_D(0))<>clng(pcv_idcategory)) AND (clng(tmp_D(0))<>"1") then
				if tmp_ParentText="" then
					tmp_ParentText="["
				else
					tmp_ParentText=tmp_ParentText & "/"
				end if
				tmp_ParentText=tmp_ParentText & tmp_D(1)
			end if
			end if
		Next
		if tmp_ParentText<>"" then
			tmp_ParentText=tmp_ParentText & "]"
		end if
	ELSE
		pcv_tmpParent=""
		if pcv_parentCategory<>"1" then
		pcf_FindParent(pcv_parentCategory)
		end if
		if pcv_tmpParent<>"" then
			pcv_tmpParent="[" & pcv_tmpParent & "]"
		end if
		tmp_ParentText=pcv_tmpParent
	END IF
	pcf_catGetParent=tmp_ParentText
End Function

function ReplaceChars(tmpText)
Dim tmp1
	tmp1=tmpText
	if tmp1<>"" then
		tmp1=replace(tmp1,"<","&lt;")
		tmp1=replace(tmp1,">","&gt;")
	end if
	ReplaceChars=tmp1
end function
		
%>
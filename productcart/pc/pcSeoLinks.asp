<%
dim pcStrCategoryDesc, pcStrCatLink, pcStrFeaturedCatLink, pcStrPrdLink, pcStrBCLink, intBCId, strBCDesc, pcv_strPreviousPageDesc, pcv_strNextPageDesc, pcv_strPreviousPage, pcv_strNextPage, pcStrPrdPreLink, pcStrPrdNextLink, pIdSeoCat, pIdCategoryTemp, pcStrPrdCSLink, pidrelation, pcStrCntPageLink, pcvContentPageName, pcIntContentPageID, pcStrPrdLinkCan

Public Sub pcGenerateSeoLinks

	'//=====================================
	'// CATEGORY LINKS - START
	'//=====================================

		if intIdCategory="1" then
			pcStrCatLink="viewcategories.asp"
		else
			if pcStrCategoryDesc = "" then
				pcStrCategoryDesc = strCategoryDesc
			end if
			pcStrCatLink=pcStrCategoryDesc & "-c" & intIdCategory & ".htm"
			pcStrCatLink=removeChars(pcStrCatLink)
		end if
		if scSeoURLs<>1 then
			pcStrCatLink="viewCategories.asp?idCategory="&intIdCategory
			pcStrCatLink2="viewCategories.asp"
		end if
		
			
		'Build Featured Category Link
		pcStrFeaturedCatLink=pcStrCategoryDesc & "-c" & pFeaturedCategory & ".htm"
		pcStrFeaturedCatLink=removeChars(pcStrFeaturedCatLink)
		if scSeoURLs<>1 then
			pcStrFeaturedCatLink="viewCategories.asp?idCategory=" & pFeaturedCategory
		end if
	'//=====================================
	'// CATEGORY LINKS - END
	'//=====================================

	
	'//=====================================
	'// PRODUCT LINKS - START
	'//=====================================

		if pIdCategory<>"" and pIdCategory<>0 then
			pIdSeoCat=pIdCategory
			else
			pIdSeoCat=pIdCategoryTemp		
		end if
		pcStrPrdLink=pDescription & "-" & pIdSeoCat & "p" & pIdProduct & ".htm"
		pcStrPrdLink=removeChars(pcStrPrdLink)
		if scSeoURLs<>1 then
			pcStrPrdLink="viewPrd.asp?idproduct="&pIdProduct&"&idcategory="&pIdCategory
		end if
		
		'Build Canonical URL Link
		'Since the same product could be assigned to multiple categories, it makes sense not to include the category in the Canonical URL
		pcStrPrdLinkCan=pDescription & "-p" & pIdProduct & ".htm"
		pcStrPrdLinkCan=removeChars(pcStrPrdLinkCan)
		if scSeoURLs<>1 then
			pcStrPrdLinkCan="viewPrd.asp?idproduct="&pIdProduct
		end if
	
		'Build BreadCrumbs Link
		pcStrBCLink=strBCDesc & "-c" & intBCId & ".htm"
		pcStrBCLink=removeChars(pcStrBCLink)
		if scSeoURLs<>1 then
			pcStrBCLink="viewCategories.asp?idCategory="&intBCId
		end if		
		
		'Build Previous Product Link
		pcStrPrdPreLink=pcv_strPreviousPageDesc & "-" & pIdCategory & "p" & pcv_strPreviousPage & ".htm"
		pcStrPrdPreLink=removeChars(pcStrPrdPreLink)
		if scSeoURLs<>1 then
			pcStrPrdPreLink="viewPrd.asp?idproduct="&pcv_strPreviousPage&"&idcategory="&pIdCategory
		end if	
		
		'Build Next Product Link
		pcStrPrdNextLink=pcv_strNextPageDesc & "-" & pIdCategory & "p" & pcv_strNextPage & ".htm"
		pcStrPrdNextLink=removeChars(pcStrPrdNextLink)
		if scSeoURLs<>1 then
			pcStrPrdNextLink="viewPrd.asp?idproduct="&pcv_strNextPage&"&idcategory="&pIdCategory
		end if
		
		'Build Cross Selling Product Link
		pcStrPrdCSLink=pDescription & "-" & pIdCategoryTemp & "p" & pidrelation & ".htm"
		pcStrPrdCSLink=removeChars(pcStrPrdCSLink)
		if scSeoURLs<>1 then
			pcStrPrdCSLink="viewPrd.asp?idproduct=" & pidrelation
		end if	
	
	'//=====================================
	'// PRODUCT LINKS - END
	'//=====================================

	
	'//=====================================
	'// CONTENT LINKS - START
	'//=====================================
		if pcvContentPageName="" then pcvContentPageName=pcv_PageNameH
		if pcvPageType = "parent" then
			pcStrCntPageLink=pcvContentPageName & "-e" & pcIntContentPageID & ".htm"
			pcStrCntPageLink=removeChars(pcStrCntPageLink)
			if scSeoURLs<>1 then
				pcStrCntPageLink="viewPages.asp?idpage=" & pcIntContentPageID
			end if
		else
			pcStrCntPageLink=pcvContentPageName & "-d" & pcIntContentPageID & ".htm"
			pcStrCntPageLink=removeChars(pcStrCntPageLink)
			if scSeoURLs<>1 then
				pcStrCntPageLink="viewcontent.asp?idpage=" & pcIntContentPageID
			end if
		end if
	'//=====================================
	'// PRODUCT LINKS - END
	'//=====================================

End Sub

%>
<!--#include file="pcSeoFunctions.asp"-->
<%'This script is Copyright(c) NetSource Commerce, http://www.productcart.com.

' ******************************************************************************
' This file dynamically generates meta tags to make ProductCart pages more
' search engine friendly. You will need to invoke this file from pc/header.asp
' as described below. Make sure that header.asp is correctly setup to
' use this file. See: http://wiki.earlyimpact.com/productcart/seo-meta-tags
'
' ******************************************************************************

Sub GenerateMetaTags()

Title = ""
Keywords = ""
mtDescription = ""

' ******************************************************************************
' Clear canonical URL
' ******************************************************************************
Dim pcStrCanonicalURL, tempCanonicalURL
tempCanonicalURL=""
pcStrCanonicalURL=""

' ******************************************************************************
' Clear image URL
' ******************************************************************************
Dim pcStrImageURL, tempImageURL
tempImageURL=""
pcStrImageURL=""

' ******************************************************************************
' Get Product and Category ID
' ******************************************************************************
GMidproduct=session("idProductRedirect")
	if GMidproduct="" then
		GMidproduct=request("idproduct")
	end if
	session("idProductRedirect")=""

GMidcategory=session("idCategoryRedirect")
	if GMidcategory = "" then
		GMidcategory=request("idcategory")
	end if
	if GMidcategory = "" then
		GMidcategory = mIdCategory
	end if
	session("idCategoryRedirect")=""
GMpcCartIndex=request("pcCartIndex")
GMidcontent=session("idContentPageRedirect")
	if GMidcontent = "" then
		GMidcontent=request("idpage")
	end if
	if GMidcontent = "" then
		GMidcontent=pcv_IDPage
	end if


' ******************************************************************************
' PRODUCT-specific Meta Tags
' ******************************************************************************

if (GMidproduct="") and (GMpcCartIndex<>"") then
	pcCartArray = Session("pcCartSession")
	GMidproduct=pcCartArray(GMpcCartIndex,0)
end if

GMTags=False

if validNum2(GMidproduct) then
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.Open scDSN

	'// Get information from "products" table
	query="select description,details,sDesc,pcprod_MetaTitle,pcprod_MetaDesc,pcprod_MetaKeywords from Products where idProduct=" & GMidproduct
	set rsTagObj=server.CreateObject("ADODB.RecordSet")
	set rsTagObj=conn.execute(query)

	if not rsTagObj.eof then
		GMTags=True
		mtPName=rsTagObj("description")
		mtPName=ClearHTMLTags2(mtPName,0)
		mtPDesc=rsTagObj("details")
		mtPDesc=ClearHTMLTags2(mtPDesc,0)
		mtPsDesc=rsTagObj("sDesc")
		if mtPsDesc<>"" then
			mtPsDesc=ClearHTMLTags2(mtPsDesc,0)
			mtPsDesc=Left(mtPsDesc,200)
		else
			mtPsDesc=Left(mtPDesc,200)
		end if
		mtPMetaTitle=rsTagObj("pcprod_MetaTitle")
			mtPMetaTitle=ClearHTMLTags2(mtPMetaTitle,0)
		mtPMetaDesc=rsTagObj("pcprod_MetaDesc")
			mtPMetaDesc=ClearHTMLTags2(mtPMetaDesc,0)
		mtPMetaKeywords=rsTagObj("pcprod_MetaKeywords")
		set rsTagObj=nothing

		' Get information from "Categories" table
		myTest=0
		If validNum2(GMidcategory) then
			query="select categoryDesc from Categories where idcategory=" & GMidcategory
			myTest=1
		else
			query="select categories.categoryDesc from Categories,Categories_Products where Categories_Products.idProduct=" & GMidproduct & " and Categories.idcategory=Categories_Products.idcategory"
		end if
		set rsTagObj=server.CreateObject("ADODB.RecordSet")
		set rsTagObj=conn.execute(query)

		mtCDesc=""
		if not rsTagObj.eof then
			mtCDesc=rsTagObj("categoryDesc")
			mtCDesc=ClearHTMLTags2(mtCDesc,0)
			if mtCDesc<>"" then
				mtCDesc=Left(mtCDesc,200)
			end if
		end if
		set rsTagObj=nothing

		'// Product Details Page: TITLE
		if not isNull(mtPMetaTitle) and mtPMetaTitle<>"" then
				Title=mtPMetaTitle
			else
				if (myTest=1) and (mtCDesc<>"") then
					Title=mtPName & " - " & mtCDesc
				else
					Title=mtPName
				end if
		end if

		'// Add store name to product page title by uncommenting the following 3 lines of code
		'if scCompanyName<>"" then
		'	Title=Title & " - " & scCompanyName
		'end if

		'// Product Details Page: KEYWORDS
		if not isNull(mtPMetaKeywords) and mtPMetaKeywords<>"" then
				Keywords=mtPMetaKeywords
			else
				Keywords=mtPName & "," & mtCDesc & "," & DefaultKeywords & "," & scCompanyName
		end if

		'// Product Details Page: DESCRIPTION
		if not isNull(mtPMetaDesc) and mtPMetaDesc<>"" then
				mtDescription=mtPMetaDesc
			else
				if not isNull(mtCDesc) and mtCDesc<>"" then
					mtDescription=mtPName & " (" & mtCDesc & "). " & mtPsDesc
				else
					mtDescription=mtPName & ". " & mtPsDesc
				end if
		end if

		'// Create canonical URL
			pIdProduct=GMidproduct
			'// Call SEO Routine
			pcGenerateSeoLinks
			'//
			tempCanonicalURL=replace((scStoreURL&"/"&scPcFolder&"/pc/"&pcStrPrdLinkCan),"//","/")
			tempCanonicalURL=replace(tempCanonicalURL,"http:/","http://")
			pcStrCanonicalURL=tempCanonicalURL

		'// Create image URL
			tempImageURL=replace((scStoreURL&"/"&scPcFolder&"/pc/catalog/"&pimageUrl),"//","/")
			tempImageURL=replace(tempImageURL,"http:/","http://")
			pcStrImageURL=tempImageURL

		'// Edit for "Tell a Friend" page
		if pcStrPageName = "tellafriend.asp" then
			Title = dictLanguage.Item(Session("language")&"_tellafriend_1") & " - " & Title
			mtDescription = ""
			Keywords = ""
			pcStrCanonicalURL = ""
			pcStrImageURL = ""
		end if

	end if
	conn.Close
	set conn=nothing
end if
' ******************************************************************************
' END PRODUCT-specific Meta Tags
' ******************************************************************************

' ******************************************************************************
' CATEGORY-specific Meta Tags
' ******************************************************************************

if (GMTags=False) and (validNum2(GMidcategory)) then
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.Open scDSN

	query="select categoryDesc, SDesc, LDesc, pcCats_MetaTitle, pcCats_MetaDesc, pcCats_MetaKeywords from categories where idCategory=" & GMidcategory
	set rsTagObj=server.CreateObject("ADODB.RecordSet")
	set rsTagObj=conn.execute(query)

	if not rsTagObj.eof then
		GMTags=True
		mtCName=rsTagObj("categoryDesc")
			if lcase(trim(mtCName))="root" then
				mtCName=dictLanguage.Item(Session("language")&"_titles_9")
			end if
		mtCName=ClearHTMLTags2(mtCName,0)
		pcStrCategoryDesc=mtCName '// Used for the canonical URL

		mtCsDesc=rsTagObj("SDesc")
		mtCsDesc=ClearHTMLTags2(mtCsDesc,0)
		mtCDesc=rsTagObj("LDesc")
		mtCDesc=ClearHTMLTags2(mtCDesc,0)

		'Do not use the information stored in the database if you are in the root category
		if GMidcategory="1" then
			GMTags=False
		end if

		mtCMetaTitle=rsTagObj("pcCats_MetaTitle")
			mtCMetaTitle=ClearHTMLTags2(mtCMetaTitle,0)
		mtCMetaDesc=rsTagObj("pcCats_MetaDesc")
			mtCMetaDesc=ClearHTMLTags2(mtCMetaDesc,0)
		mtCMetaKeywords=rsTagObj("pcCats_MetaKeywords")
		set rsTagObj=nothing

		if mtCsDesc<>"" then
			mtCsDesc=Left(mtCsDesc,200)
		end if

		if mtCDesc<>"" then
			mtCDesc=Left(mtCDesc,200)
		end if

		if trim(mtCsDesc) <> trim(mtCDesc) then
			mtCDesc=mtCsDesc & " " & mtCDesc
			mtCDesc=trim(mtCDesc)
		end if

		'// Category Page: TITLE
		if not isNull(mtCMetaTitle) and mtCMetaTitle<>"" then
				Title=mtCMetaTitle
			else
				if scCompanyName<>"" then
					Title=mtCName & " - " & scCompanyName
				else
					Title=mtCName
				end if
		end if

		'// Category Page: KEYWORDS
		if not isNull(mtCMetaKeywords) and mtCMetaKeywords<>"" then
				Keywords=mtCMetaKeywords
			else
				Keywords=mtCName & "," & DefaultKeywords & "," & scCompanyName
		end if

		'// Category Page: DESCRIPTION
		if not isNull(mtCMetaDesc) and mtCMetaDesc<>"" then
				mtDescription=mtCMetaDesc
			else
				mtDescription=mtCName & ". " & mtCDesc
		end if

		'// Create canonical URL
			intIdCategory=GMidcategory
			'// Call SEO Routine
			pcGenerateSeoLinks
			'//
			tempCanonicalURL=replace((scStoreURL&"/"&scPcFolder&"/pc/"&pcStrCatLink),"//","/")
			tempCanonicalURL=replace(tempCanonicalURL,"http:/","http://")
			pcStrCanonicalURL=tempCanonicalURL

		'// Create image URL
			Dim pcStrCatImgTemp
			query = "SELECT image FROM Categories WHERE idCategory = " & intIdCategory
			SET rsCatImg=Server.CreateObject("ADODB.RecordSet")
			SET rsCatImg=conn.execute(query)
			pcStrCatImgTemp=rsCatImg("image")
			SET rsCatImg=nothing
			if trim(pcStrCatImgTemp)<>"" and intIdCategory<>"1" then
				tempImageURL=replace((scStoreURL&"/"&scPcFolder&"/pc/catalog/"&pcStrCatImgTemp),"//","/")
			elseif scCompanyLogo<>"" then
				tempImageURL=replace((scStoreURL&"/"&scPcFolder&"/pc/catalog/"&scCompanyLogo),"//","/")
			else
				tempImageURL=""
			end if
			tempImageURL=replace(tempImageURL,"http:/","http://")
			pcStrImageURL=tempImageURL

	end if
	conn.Close
	set conn=nothing
end if

' ******************************************************************************
' END CATEGORY-specific Meta Tags
' ******************************************************************************
' ******************************************************************************
' CONTENT PAGE-specific Meta Tags
' ******************************************************************************
if (GMTags=False) and (validNum2(GMidcontent)) then
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.Open scDSN

		'// Title, description, and keywords are set on viewcontent.asp

		'// Create canonical URL
			pcIntContentPageID=GMidcontent
			'// Call SEO Routine
			pcGenerateSeoLinks
			'//
			tempCanonicalURL=replace((scStoreURL&"/"&scPcFolder&"/pc/"&pcStrCntPageLink),"//","/")
			tempCanonicalURL=replace(tempCanonicalURL,"http:/","http://")
			pcStrCanonicalURL=tempCanonicalURL

		'// Create image URL
			Dim pcStrPageImgTemp
			query = "SELECT pcCont_Thumbnail FROM pcContents WHERE pcCont_IDPage = " & GMidcontent
			SET rsPageImg=Server.CreateObject("ADODB.RecordSet")
			SET rsPageImg=conn.execute(query)
			pcStrPageImgTemp=rsPageImg("pcCont_Thumbnail")
			SET rsPageImg=nothing
			if trim(pcStrPageImgTemp)<>"" then
				tempImageURL=replace((scStoreURL&"/"&scPcFolder&"/pc/catalog/"&pcStrPageImgTemp),"//","/")
			elseif scCompanyLogo<>"" then
				tempImageURL=replace((scStoreURL&"/"&scPcFolder&"/pc/catalog/"&scCompanyLogo),"//","/")
			else
				tempImageURL=""
			end if
			tempImageURL=replace(tempImageURL,"http:/","http://")
			pcStrImageURL=tempImageURL

	conn.Close
	set conn=nothing
end if

' ******************************************************************************
' END CONTENT PAGE-specific Meta Tags
' ******************************************************************************

'// Build the meta tags

if (GMTags=False) and ((pcv_DefaultTitle<>"") or (pcv_DefaultKeywords<>"") or (pcv_DefaultDescription<>"")) then
	GMTags=True
	Title= pcv_DefaultTitle
	Keywords = pcv_DefaultKeywords
	mtDescription = pcv_DefaultDescription
end if

if (GMTags=False) and (scCompanyName<>"") then
	GMTags=True
	Title = scMetaTitle & " - " & scCompanyName
	Keywords = scMetaKeywords
	mtDescription = scMetaDescription
end if

if (GMTags=False) then
	Title = scMetaTitle
	Keywords = scMetaKeywords
	mtDescription = scMetaDescription
end if

'// START - Adjust Meta Tags for certain storefront pages

	pcStrPageName = lcase(pcStrPageName)
	pcStrPageNameOR = lcase(pcStrPageNameOR)
	tempPageURL=replace((scStoreURL&"/"&scPcFolder&"/pc/"),"//","/")
	tempPageURL=replace(tempPageURL,"http:/","http://")
	tempPageURL=replace(tempPageURL,"https:/","https://")

	'// Canonical URL
	if pcStrPageName = "home.asp" or pcStrPageName = "showbestsellers.asp" or pcStrPageName = "showfeatured.asp" or pcStrPageName = "showspecials.asp" or pcStrPageName = "shownewarrivals.asp" or (pcStrPageName = "viewcategories.asp" and GMidcategory=1) then
		pcStrCanonicalURL=tempPageURL & pcStrPageName
	end if
	if pcStrPageNameOR = "showrecentlyreviewed.asp" then
		pcStrCanonicalURL=tempPageURL & pcStrPageNameOR
	end if
	if (validNum2(GMidcategory)) and (lcase(pcStrPageName)="showsearchresults.asp") then
		pcStrCanonicalURL=""
	end if

	'// Title Meta Tag
	if pcStrPageName = "showbestsellers.asp" then
		Title = dictLanguage.Item(Session("language")&"_viewBestSellers_2") & " - " & scCompanyName
	end if
	if pcStrPageName = "shownewarrivals.asp" then
		Title = dictLanguage.Item(Session("language")&"_viewNewArrivals_2") & " - " & scCompanyName
	end if
	if pcStrPageName = "showspecials.asp" then
		Title = dictLanguage.Item(Session("language")&"_viewSpc_2") & " - " & scCompanyName
	end if
	if pcStrPageName = "showfeatured.asp" then
		Title = dictLanguage.Item(Session("language")&"_mainIndex_7") & " - " & scCompanyName
	end if
	if pcStrPageName = "showsearchresults.asp" then
		Title = dictLanguage.Item(Session("language")&"_ShowSearch_50") & " - " & scCompanyName
	end if
	if pcStrPageName = "viewcategories.asp" and GMidcategory=1 then
		Title = dictLanguage.Item(Session("language")&"_titles_9") & " - " & scCompanyName
	end if
	if pcStrPageNameOR = "showrecentlyreviewed.asp" then
		Title = dictLanguage.Item(Session("language")&"_ShowRecentRev_1") & " - " & scCompanyName
	end if

'// END - Adjust Meta Tags

'// START - Clean up

	if Title<>"" and not isNull(Title) then
		Title=replace(Title,"""","&quot;")
		Title=replace(Title," - ,",",")
	end if
	if Keywords<>"" and not isNull(Keywords) then
		Keywords=replace(Keywords,"""","")
		Keywords=replace(Keywords,"&quot;","")
		'Keywords=replace(Keywords," - ,",",")
	end if
	if mtDescription<>"" and not isNull(mtDescription) then
		mtDescription=replace(mtDescription,"""","")
		mtDescription=replace(mtDescription,"&quot;","")
		mtDescription=replace(mtDescription," - ,",",")
	end if

'// END - Clean up

'// START - Write Meta Tags
	if trim(Title)<>"" then
		Response.Write "<TITLE>" & Title & "</TITLE>" & vbcrlf
	end if
	if trim(mtDescription)<>"" then
		Response.Write "<META NAME=""Description"" CONTENT=""" & mtDescription & """ />" & vbcrlf
	end if
	if trim(Keywords)<>"" then
		Response.Write "<META NAME=""Keywords"" CONTENT=""" & Keywords & """ />" & vbcrlf
	end if
	Response.Write "<META NAME=""Robots"" CONTENT=""index,follow"" />" & vbcrlf & _
	"<META NAME=""Revisit-after"" CONTENT=""30"" />" & vbcrlf

	'// Add Canonical URL
	if pcStrCanonicalURL<>"" then
		Response.Write "<link rel=""canonical"" href=""" & pcStrCanonicalURL & """ />" & vbcrlf
	end if

	'// Add image URL
	Dim tmpHasImg
	tmpHasImg=0
	if pcStrImageURL<>"" then
		if len(pcStrImageURL)-InstrRev(pcStrImageURL,"/")>0 then
		tmpStr=Right(pcStrImageURL,len(pcStrImageURL)-InstrRev(pcStrImageURL,"/"))
		if tmpStr<>"" then
			if Instr(tmpStr,".")>0 then
				tmpHasImg=1
			end if
		end if
		end if
	end if
	if (pcStrImageURL<>"") AND (tmpHasImg=1) then
		Response.Write "<link rel=""image_src"" href=""" & pcStrImageURL & """ />" & vbcrlf
	end if
'// END - Write Meta Tags

End Sub


	'[ClearHTMLTags2]

	'Coded by Jóhann Haukur Gunnarsson
	'joi@innn.is

	'	Purpose: This function clears all HTML tags from a string using Regular Expressions.
	'	 Inputs: strHTML2;	A string to be cleared of HTML TAGS
	'		 intWorkFlow2;	An integer that if equals to 0 runs only the regEx2p filter
	'							  .. 1 runs only the HTML source render filter
	'							  .. 2 runs both the regEx2p and the HTML source render
	'							  .. >2 defaults to 0
	'	Returns: A string that has been filtered by the function


	function ClearHTMLTags2(strHTML2, intWorkFlow2)

		'Variables used in the function

		dim regEx2, strTagLess2

		'---------------------------------------
		strTagLess2 = strHTML2
		'Move the string into a private variable
		'within the function
		'---------------------------------------

		'---------------------------------------
		'NetSource Commerce codes
		IF strTagLess2<>"" THEN
		strTagLess2=replace(strTagLess2,"<br>"," ")
		strTagLess2=replace(strTagLess2,"<BR>"," ")
		strTagLess2=replace(strTagLess2,"<p>"," ")
		strTagLess2=replace(strTagLess2,"<P>"," ")
		strTagLess2=replace(strTagLess2,"</p>"," ")
		strTagLess2=replace(strTagLess2,"</P>"," ")
		strTagLess2=replace(strTagLess2,vbcrlf," ")
		strTagLess2=trim(strTagLess2)
		do while instr(strTagLess2,"  ")>0
		strTagLess2=replace(strTagLess2,"  "," ")
		loop
		END IF
		'Modify the string to a friendly ONLY 1 LINE string
		'---------------------------------------

		IF strTagLess2<>"" THEN

		'regEx2 initialization

		'---------------------------------------
		set regEx2 = New regExp
		'Creates a regEx2p object
		regEx2.IgnoreCase = True
		'Don't give frat about case sensitivity
		regEx2.Global = True
		'Global applicability
		'---------------------------------------


		'Phase I
		'	"bye bye html tags"


		if intWorkFlow2 <> 1 then

			'---------------------------------------
			regEx2.Pattern = "<[^>]*>"
			'this pattern mathces any html tag
			strTagLess2 = regEx2.Replace(strTagLess2, "")
			'all html tags are stripped
			'---------------------------------------

		end if


		'Phase II
		'	"bye bye rouge leftovers"
		'	"or, I want to render the source"
		'	"as html."

		'---------------------------------------
		'We *might* still have rouge < and >
		'let's be positive that those that remain
		'are changed into html characters
		'---------------------------------------


		if intWorkFlow2 > 0 and intWorkFlow2 < 3 then


			regEx2.Pattern = "[<]"
			'matches a single <
			strTagLess2 = regEx2.Replace(strTagLess2, "&lt;")

			regEx2.Pattern = "[>]"
			'matches a single >
			strTagLess2 = regEx2.Replace(strTagLess2, "&gt;")
			'---------------------------------------

		end if


		'Clean up

		'---------------------------------------
		set regEx2 = nothing
		'Destroys the regEx2p object
		'---------------------------------------

		END IF 'vefiry strTagLess2 (null strings)

		'---------------------------------------
		ClearHTMLTags2 = strTagLess2
		'The results are passed back
		'---------------------------------------

	end function

	'check for real integers
	Function validNum2(strInput)
		DIM iposition		' Current position of the character or cursor
		validNum2 =  true
		if isNULL(strInput) OR trim(strInput)="" then
			validNum2 = false
		else
			'loop through each character in the string and validate that it is a number or integer
			For iposition=1 To Len(trim(strInput))
				if InStr(1, "12345676890", mid(strInput,iposition,1), 1) = 0 then
					validNum2 =  false
					Exit For
				end if
			Next
		end if
	end Function

%>
<!--#include file="pcSeoLinks.asp"-->
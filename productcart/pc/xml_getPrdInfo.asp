<%Response.ContentType = "text/xml"%><?xml version="1.0" ?>
<%

'*****************************************************
'* BEGIN: Check HTTP Referer
'*****************************************************

strPath=Request.ServerVariables("PATH_INFO")
dim iCnt, strPath,strPathInfo
iCnt=0
do while iCnt<1
	if mid(strPath,len(strPath),1)="/" then
		iCnt=iCnt+1
	end if
	if iCnt<1 then
		strPath=mid(strPath,1,len(strPath)-1)
	end if
loop
if Ucase(Request.ServerVariables("HTTPS"))="OFF" then
	strPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & strPath
else
	strPathInfo="https://" & Request.ServerVariables("HTTP_HOST") & strPath
end if
				
if Right(strPathInfo,1)="/" then
else
	strPathInfo=strPathInfo & "/"
end if

strRefferer=Request.ServerVariables("HTTP_REFERER")

'*****************************************************
'* END: Check HTTP Referer
'*****************************************************

'*****************************************************
'* BEGIN: Check Query
'*****************************************************

pIdProduct=getUserInput(request("idProduct"),0)

tmpTest=0

if pIdProduct="" then
	tmpTest=1
else
	if pIdProduct="0" then
		tmpTest=1
	else
		if not IsNumeric(pIdProduct) then
			tmpTest=1
		end if
	end if
end if

'*****************************************************
'* END: Check Query
'*****************************************************
			
if ((session("store_useAjax")<>"") AND (Instr(ucase(strRefferer),ucase(strPathInfo))=0)) OR (tmpTest=1) then%>
<bcontent>nothing</bcontent>
<%response.end
end if%><!--#include file="pcStartSession.asp" --><!--#include file="../includes/settings.asp"--><!--#include file="../includes/storeconstants.asp"--><!--#include file="../includes/opendb.asp"--><!--#include file="../includes/languages.asp"--><!--#include file="../includes/currencyformatinc.asp"--><!--#include file="../includes/shipFromSettings.asp"--><!--#include file="../includes/taxsettings.asp"--><!--#include file="../includes/languages_ship.asp"--><!--#include file="../includes/adovbs.inc"--><!--#include file="../includes/stringfunctions.asp"--><% 
Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN  

Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")

dim query, conntemp, rs, pIdProduct, pDescription, pPrice, pDetails, pListPrice, pLgimageURL, pImageUrl, pWeight, pSku ,pconfigOnly,pserviceSpec, pBtoBPrice, plistHidden, pArequired,pBrequired, pStock, pEmailText, pformQuantity, pnoshipping, pcustom1, pcustom2, pcustom3, pcontent1, pcontent2, pcontent3, pxfield1, pxfield2, pxfield3, px1req, px2req, px3req, pNoStock, psDesc, pnoshippingtext

pIdProduct=getUserInput(request("idProduct"),0)

Dim iAddDefaultPrice,	iAddDefaultWPrice%><!--#include file="pcCheckPricingCats.asp"--><%

'--> open database connection
call opendb()


' --> gets product details from db
query="SELECT iRewardPoints, description, sku, configOnly, serviceSpec, price, btobprice, details, listprice, listHidden, imageurl, smallImageUrl, largeImageURL, Arequired, Brequired, stock, weight, emailText, formQuantity, noshipping, noprices, IDBrand, sDesc, noStock, noshippingtext,pcprod_HideBTOPrice,pcprod_QtyValidate,pcprod_MinimumQty,pcprod_HideDefConfig, notax, pcProd_BackOrder, pcProd_ShipNDays, pcProd_HideSKU FROM products WHERE idProduct=" &pidProduct& " AND active=-1"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

' --> set product variables <---
iRewardPoints=rs("iRewardPoints")
pDescription=ClearHTMLTags2(replace(rs("description"),"&quot;",chr(34)),0)
pSku= rs("sku")
pconfigOnly=rs("configOnly")
pserviceSpec=rs("serviceSpec")
pPrice=rs("price")
pBtoBPrice=rs("bToBPrice")
pDetails=ClearHTMLTags2(replace(rs("details"),"&quot;",chr(34)),0)
pListPrice=rs("listPrice")
plistHidden=rs("listHidden")
pimageUrl=rs("imageUrl")
pSmimageURL=rs("smallImageUrl")
pLgimageURL=rs("largeImageURL")
pArequired=rs("Arequired")
pBrequired=rs("Brequired")
pStock=rs("stock")
pWeight=rs("weight")
pEmailText=rs("emailText")
pFormQuantity=rs("formQuantity")
pnoshipping=rs("noshipping")
pnoprices=rs("noprices")
if isNull(pnoprices) OR pnoprices="" then
	pnoprices=0
end if
pIDBrand=rs("IDBrand")
psDesc=ClearHTMLTags2(rs("sDesc"),0)
pNoStock=rs("noStock")
pnoshippingtext=rs("noshippingtext")
pcv_intHideBTOPrice=rs("pcprod_HideBTOPrice")
if isNull(pcv_intHideBTOPrice) OR pcv_intHideBTOPrice="" then
	pcv_intHideBTOPrice="0"
end if
if pnoprices=2 then
	pcv_intHideBTOPrice=1
end if
pcv_intQtyValidate=rs("pcprod_QtyValidate")
if isNull( pcv_intQtyValidate) OR  pcv_intQtyValidate="" then
	pcv_intQtyValidate="0"
end if				
pcv_lngMinimumQty=rs("pcprod_MinimumQty")
if isNull(pcv_lngMinimumQty) OR pcv_lngMinimumQty="" then
	pcv_lngMinimumQty="0"
end if
intpHideDefConfig=rs("pcprod_HideDefConfig")
if isNull(intpHideDefConfig) OR intpHideDefConfig="" then
	intpHideDefConfig="0"
end if
pnotax=rs("notax")
pcv_intBackOrder = rs("pcProd_BackOrder")
if isNull(pcv_intBackOrder) OR pcv_intBackOrder="" then
	pcv_intBackOrder = 0
end if
pcv_intShipNDays = rs("pcProd_ShipNDays")
if isNull(pcv_intShipNDays) OR pcv_intShipNDays="" then
	pcv_intShipNDays = 0
end if
pcv_intHideSKU = rs("pcProd_HideSKU")
if isNull(pcv_intHideSKU) OR pcv_intHideSKU="" then
	pcv_intHideSKU = 0
end if
set rs=nothing

' Check to see if the product has been assigned to a brand. If so, get the brand name
if (pIDBrand&""<>"") and (pIDBrand&""<>"0") then
 	query="select BrandName from Brands where IDBrand=" & pIDBrand
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if not rs.eof then
		BrandName=rs("BrandName")
	end if
	set rs=nothing
end if

' Check to see if this is a BTO product. If so, get additional product information
if pserviceSpec<>0 then
	query="SELECT categories.categoryDesc, products.description, configSpec_products.configProductCategory, configSpec_products.price, categories_products.idCategory, categories_products.idProduct, products.weight FROM categories, products, categories_products INNER JOIN configSpec_products ON categories_products.idCategory=configSpec_products.configProductCategory WHERE (((configSpec_products.specProduct)="&pIdProduct&") AND ((configSpec_products.configProduct)=[categories_products].[idproduct]) AND ((categories_products.idCategory)=[categories].[idcategory]) AND ((categories_products.idProduct)=[products].[idproduct]) AND ((configSpec_products.cdefault)<>0)) ORDER BY configSpec_products.catSort, categories.idCategory, configSpec_products.prdSort;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	set rs=nothing
	query="SELECT * FROM configSpec_Charges WHERE specProduct="&pIdProduct
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=conntemp.execute(query)
	BTOCharges=0
	if not rstemp.eof then
		BTOCharges=1
	end if
	set rstemp=nothing
end if

Dim tmpList
tmpList=""

if len(pDescription )>50 then
	pDescription =Left(pDescription ,50) & "..."
end if

tmpList=pDescription & "|||" & "<table class=""mainbox"">"

	if pSmimageURL<>"" and pSmimageURL<>"no_image.gif" then
		tmpList=tmpList & "<tr><td class=""mainboxImg"">"
		tmpList=tmpList & "<img src=""catalog/" & pSmimageURL & """ >" & vbcrlf
		tmpList=tmpList & "</td></tr>"
	end if

	tmpList=tmpList & "<tr><td>"
	if pcv_intHideSKU="0" then
		tmpList=tmpList & dictLanguage.Item(Session("language")&"_CustviewPastD_5") & ": " & pSku & "<br>" & vbcrlf
	end if

'Begin WEIGHT
if scShowProductWeight="-1" then
		if int(pWeight)>0 then
			if scShipFromWeightUnit="KGS" then
				pKilos=Int(pWeight/1000)
				pWeight_g=pWeight-(pKilos*1000)
				pWeight=pKilos
				if pWeight_g>0 then
					tmpList=tmpList & ship_dictLanguage.Item(Session("language")&"_viewCart_c") & pWeight&" "& ship_dictLanguage.Item(Session("language")&"_xmlPrdInfo_kg") &" "&pWeight_g&" "& ship_dictLanguage.Item(Session("language")&"_xmlPrdInfo_g") & "<br>" & vbcrlf
				else
					tmpList=tmpList & ship_dictLanguage.Item(Session("language")&"_viewCart_c") & pWeight&" "& ship_dictLanguage.Item(Session("language")&"_xmlPrdInfo_kg") & "<br>" & vbcrlf
				end if
			else
				pPounds=Int(pWeight/16)
				pWeight_oz=pWeight-(pPounds*16)
				pWeight=pPounds
				if pWeight_oz>0 then
					tmpList=tmpList & ship_dictLanguage.Item(Session("language")&"_viewCart_c") & pWeight&" "& ship_dictLanguage.Item(Session("language")&"_xmlPrdInfo_lbs") &" "&pWeight_oz&" "& ship_dictLanguage.Item(Session("language")&"_xmlPrdInfo_ozs") & "<br>" & vbcrlf
				else
					tmpList=tmpList & ship_dictLanguage.Item(Session("language")&"_viewCart_c") & pWeight&" "& ship_dictLanguage.Item(Session("language")&"_xmlPrdInfo_lbs") & "<br>" & vbcrlf
				end if
			end if
		end if
end if
' End WEIGHT
	
' If the product has been assigned to a BRAND, display it here
if sBrandPro="1" then
	if (pIDBrand&""<>"") and (pIDBrand&""<>"0") then
			tmpList=tmpList & dictLanguage.Item(Session("language")&"_viewPrd_brand") & BrandName & "<br>" & vbcrlf
	end if
end if
	 
' If Show Units in stock is on, show the stock level here
if scdisplayStock=-1 AND pNoStock=0 then
	if pstock > 0 then
		tmpList=tmpList  & dictLanguage.Item(Session("language")&"_advSrca_8") & " " & pStock & " items<br>" & vbcrlf
	end if
end if

'// Start Reward Points	
pcv_BTORP=Clng(0)

if pserviceSpec=true then
'// Product is BTO
	' Get data
	query="SELECT sum(products.iRewardPoints) As RewardToTal FROM categories, products, categories_products INNER JOIN configSpec_products ON categories_products.idCategory=configSpec_products.configProductCategory WHERE (((configSpec_products.specProduct)="&pIdProduct&") AND ((configSpec_products.configProduct)=[categories_products].[idproduct]) AND ((categories_products.idCategory)=[categories].[idcategory]) AND ((categories_products.idProduct)=[products].[idproduct]) AND ((configSpec_products.cdefault)<>0));"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
		
	if not rs.eof then
		pcv_BTORP=rs("RewardToTal")
		if IsNull(pcv_BTORP) or pcv_BTORP="" then
			pcv_BTORP=0
		end if
	end if
	set rs=nothing
end if
	

If RewardsActive=1 then
	' Show Reward Points associated with this product, if any
	' By default, Reward Points are not shown to Wholesale Customers
	if Clng(iRewardPoints+clng(pcv_BTORP))>"0" and session("customerType")<>"1" then
		tmpList=tmpList & RewardsLabel & ": " & Clng(iRewardPoints+clng(pcv_BTORP)) & "<br>" & vbcrlf
	else
		' If the system is setup to include Wholesale Customers, then show Reward Points to them too
		if Clng(iRewardPoints+clng(pcv_BTORP))>"0" and session("customerType")="1" and RewardsIncludeWholesale=1 then
			tmpList=tmpList & RewardsLabel & ": " & Clng(iRewardPoints+clng(pcv_BTORP)) & "<br>" & vbcrlf
		end if 
	end If
End If
'// End Reward Points

%><!--#include file="pcGetPrdPrices.asp"--><%

				
' Show product prices
' Don't show prices if the BTO product has been set up to hide prices (pnoprices)
If pnoprices<2 Then
	 
		' Display the online price if it's not zero
		if (pPrice>0) and (pcv_intHideBTOPrice<>"1") then
			tmpList=tmpList & dictLanguage.Item(Session("language")&"_prdD1_1") & ": " & scCurSign & money(pPrice) & "<br>" & vbcrlf
	
			' If the List Price is not zero and higher than the online price, display striken through
			if ((pListPrice-pPrice)>0) and (pcv_intHideBTOPrice<>"1") then
				tmpList=tmpList & dictLanguage.Item(Session("language")&"_viewPrd_20") & scCurSign & money(pListPrice) & "<br>" & vbcrlf
			end if
			
			' If the product is setup to use the Show Savings feature, show the savings if they exist and the customer is retail
			if ((pListPrice-pPrice)>0) AND (plistHidden<0) AND (session("customerType")<>1) and (pcv_intHideBTOPrice<>"1") then
				tmpList=tmpList & dictLanguage.Item(Session("language")&"_prdD1_2") & scCurSign & money((pListPrice-pPrice)) & "<br>" & vbcrlf
			end if
		
		end if 'this is the IF statement regarding the online price being > zero
		
		' If this is a wholesale customer and the wholesale price is > zero, display it here
		if pcv_intHideBTOPrice<>"1" then
			if session("customerCategory")<>0 and (dblpcCC_Price>"0") and (pcv_intHideBTOPrice<>"1") then
				tmpList=tmpList & session("customerCategoryDesc")&": " & scCurSign & money(dblpcCC_Price) & "<br>" & vbcrlf
			else
				if (dblpcCC_Price>"0") and (session("customerType")="1") and (pcv_intHideBTOPrice<>"1") then 
					tmpList=tmpList & dictLanguage.Item(Session("language")&"_prdD1_4") & scCurSign & money(dblpcCC_Price) & "<br>" & vbcrlf
				end if
			end if
		end if		
	
end if 'this is the IF statement regarding the BTO product being setup not to show prices

'Show "Back-Order" message
If (scOutofStockPurchase=-1 AND CLng(pStock)<1 AND pserviceSpec=0 AND pNoStock=0 AND pcv_intBackOrder=1) OR (pserviceSpec<>0 AND scOutofStockPurchase=-1 AND iBTOOutofstockpurchase=-1 AND CLng(pStock)<1 AND pNoStock=0 AND pcv_intBackOrder=1) Then
	If clng(pcv_intShipNDays)>0 then
		tmpList=tmpList & dictLanguage.Item(Session("language")&"_sds_viewprd_1") & pcv_intShipNDays & dictLanguage.Item(Session("language")&"_sds_viewprd_1b") & "<br>" & vbcrlf
	End if
End If

'Show "Out of Stock" message
if (scShowStockLmt=-1 AND CLng(pStock)<1 AND pserviceSpec=0 AND pNoStock=0 AND pcv_intBackOrder=0) OR (pserviceSpec<>0 AND scShowStockLmt=-1 AND iBTOOutofstockpurchase=-1 AND CLng(pStock)<1 AND pNoStock=0 AND pcv_intBackOrder=0) then
	tmpList=tmpList & dictLanguage.Item(Session("language")&"_viewPrd_7") & "<br>" & vbcrlf
end if

' Show description
	if len(psDesc)>150 then
		psDesc=Left(psDesc,150) & "..."
	end if
	if psDesc<>"" then
		tmpList=tmpList & "<br>" & psDesc & "<br>" & vbcrlf
		else
			if len(pDetails)>150 then
				pDetails=Left(pDetails,150) & "..."
			end if
		tmpList=tmpList & "<br>" & pDetails & "<br>" & vbcrlf
	end if
' End show description

tmpList=tmpList & "</td></tr></table>"

set rs=nothing
set rstemp=nothing
Set RSlayout = nothing
Set rsIconObj = nothing
Set conlayout=nothing
call closedb()%><bcontent><%=Server.HTMLEncode(tmpList)%></bcontent>
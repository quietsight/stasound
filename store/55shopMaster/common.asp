<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

InvalidGrp1=0
InvalidGrp2=0
InvalidGrp3=0
InvalidGrp4=0
InvalidGrp5=0

skuid=-1
nameid=-1
descid=-1
sdescid=-1
ptypeid=-1
opriceid=-1
lpriceid=-1
wpriceid=-1
weightid=-1
unitslbID=-1
stockid=-1

surcharge1id=-1
surcharge2id=-1

surcharge1=0
surcharge2=0

prdnoteid=-1
eimagid=-1
playoutid=-1
hideskuid=-1

goCatid=-1
goGenid=-1
goAgeid=-1
goColorid=-1
goSizeid=-1
goPatid=-1
goMatid=-1

goCat=""
goGen=""
goAge=""
goColor=""
goSize=""
goPat=""
goMat=""

prdnote=""
eimag=0
playout=""
hidesku=0


mt_titleID=-1
mt_descID=-1
mt_keyID=-1

categoryid=-1
SCatDescid=-1
LCatDescid=-1
csimageid=-1
climageid=-1
pcategoryid=-1

category1id=-1
SCatDesc1id=-1
LCatDesc1id=-1
csimage1id=-1
climage1id=-1
pcategory1id=-1

category2id=-1
SCatDesc2id=-1
LCatDesc2id=-1
csimage2id=-1
climage2id=-1
pcategory2id=-1

timageid=-1
gimageid=-1
dimageid=-1
activeid=-1
savingid=-1
specialid=-1
featuredid=-1
rwpid=-1
ntaxid=-1
nshipid=-1
nforsaleid=-1
nforsalecopyid=-1
distockid=-1
dishiptextid=-1
OverSizeID=-1
dim customfieldsid(2)
customfieldsid(0) = -1
customfieldsid(1) = -1
customfieldsid(2) = -1
brandnameid=-1
brandlogoid=-1

fileurlid=-1
urlexpireid=-1
expiredaysid=-1
licenseid=-1
localLGid=-1
remoteLGid=-1
LFN1id=-1
LFN2id=-1
LFN3id=-1
LFN4id=-1
LFN5id=-1
Addcopyid=-1


MQtyID=-1
VQtyID=-1

Opt1ID=-1
Attr1ID=-1
Opt1ReqID=-1
Opt1OrderID=-1
Opt2ID=-1
Attr2ID=-1
Opt2ReqID=-1
Opt2OrderID=-1
Opt3ID=-1
Attr3ID=-1
Opt3ReqID=-1
Opt3OrderID=-1
Opt4ID=-1
Attr4ID=-1
Opt4ReqID=-1
Opt4OrderID=-1
Opt5ID=-1
Attr5ID=-1
Opt5ReqID=-1
Opt5OrderID=-1

downprdID=-1
giftcertID=-1
giftexpID=-1
giftelectID=-1
giftgenID=-1
giftexpdateID=-1
giftexpdaysID=-1
giftcustgenfileID=-1

'BTO
hidebtopriceid=-1
hideconfid=-1
dispurchaseid=-1
skipdetailsid=-1

'Start SDBA
prd_CostID=-1
prd_BackOrderID=-1
prd_ShipNDaysID=-1
prd_NotifyStockID=-1
prd_ReorderLevelID=-1
prd_IsDropShippedID=-1
prd_SupplierID=-1
prd_DropShipperID=-1
prd_IsDropShipperID=-1

prd_Cost=0
prd_BackOrder=0
prd_ShipNDays=0
prd_NotifyStock=0
prd_ReorderLevel=0
prd_IsDropShipped=0
prd_Supplier=0
prd_DropShipper=0
prd_IsDropShipper=0
'End SDBA

'BTO
prd_hidebtoprice=0
prd_hideconf=0
prd_dispurchase=0
prd_skipdetails=0

prd_Opt1=""
prd_Attr1=""
prd_Opt1Req=0
prd_Opt1Order=1
prd_Opt2=""
prd_Attr2=""
prd_Opt2Req=0
prd_Opt2Order=2
prd_Opt3=""
prd_Attr3=""
prd_Opt3Req=0
prd_Opt3Order=3
prd_Opt4=""
prd_Attr4=""
prd_Opt4Req=0
prd_Opt4Order=4
prd_Opt5=""
prd_Attr5=""
prd_Opt5Req=0
prd_Opt5Order=5

prd_downprd=0
prd_giftcert=0
prd_giftexp=0
prd_giftelect=0
prd_giftgen=0
prd_giftexpdate=""
prd_giftexpdays=0
prd_giftcustgenfile=""


		psku=""
		pname=""
		pdesc=""
		sdesc=""
		pptype=""
		poprice=""
		plprice=""
		pwprice=""
		pweight=""
		unitslb=0
		pstock=""
		
		mt_title=""
		mt_desc=""
		mt_key=""
		
		pcategory=""
		SCATDesc=""
		LCATDesc=""
		pcsimage=""
		pclimage=""
		ppcategory=""
		
		pcategory1=""
		SCATDesc1=""
		LCATDesc1=""
		pcsimage1=""
		pclimage1=""
		ppcategory1=""
		
		pcategory2=""
		SCATDesc2=""
		LCATDesc2=""
		pcsimage2=""
		pclimage2=""
		ppcategory2=""		
		
		ptimage=""
		pgimage=""
		pdimage=""
		pactive=""
		psaving=""
		pspecial=""
		pfeatured=0
		prwp=""
		pntax=""
		pnship=""
		pnforsale=""
		pnforsalecopy=""
		distock=0
		dishiptext=0
		oversize=""
		dim customfields(2)
		customfields(0) = ""
		customfields(1) = ""
		customfields(2) = ""
		dim customfieldsname(2)
		customfieldsname(0) = ""
		customfieldsname(1) = ""
		customfieldsname(2) = ""
		brandname=""
		brandlogo=""
		
		fileurl=""
		urlexpire=0
		expiredays=0
		license=0
		localLG=""
		remoteLG=""
		LFN1=""
		LFN2=""
		LFN3=""
		LFN4=""
		LFN5=""
		Addcopy=""
		
		MQty=0
		VQty=0
		

TempProducts=""
ErrorsReport=""%>

<!--#include file="checkfields.asp"-->

<%

For i=1 to request("validfields")
	
	BLine=0
	Select Case request("T" & i)
	Case "SKU": skuid=request("P" & i)
	BLine=1
	Case "Name": nameid=request("P" & i)
	BLine=2
	Case "Description": descid=request("P" & i)
	BLine=3
	Case "Short Description": sdescid=request("P" & i)
	BLine=4
	Case "Product Type": ptypeid=request("P" & i)	
	BLine=5
	Case "Online Price": opriceid=request("P" & i)
	BLine=6
	Case "List Price": lpriceid=request("P" & i)
	BLine=7
	Case "Wholesale Price": wpriceid=request("P" & i)
	BLine=8
	Case "Weight": weightid=request("P" & i)
	BLine=9
	Case "Stock": stockid=request("P" & i)
	BLine=10
	Case "Category Name": categoryid=request("P" & i)
	BLine=11
	Case "Category Small Image": csimageid=request("P" & i)
	BLine=12
	Case "Category Large Image": climageid=request("P" & i)
	BLine=13
	Case "Parent Category": pcategoryid=request("P" & i)
	BLine=14
	Case "Additional Category 1": category1id=request("P" & i)
	BLine=15
	Case "Parent Category 1": pcategory1id=request("P" & i)
	BLine=16
	Case "Additional Category 2": category2id=request("P" & i)			
	BLine=17
	Case "Parent Category 2": pcategory2id=request("P" & i)
	BLine=18
	Case "Brand Name": brandnameid=request("P" & i)
	BLine=19
	Case "Brand Logo": brandlogoid=request("P" & i)		
	BLine=20
	Case "Thumbnail Image": timageid=request("P" & i)
	BLine=21
	Case "General Image": gimageid=request("P" & i)
	BLine=22
	Case "Detail view Image": dimageid=request("P" & i)
	BLine=23
	Case "Active": activeid=request("P" & i)
	BLine=24
	Case "Show savings": savingid=request("P" & i)
	BLine=25
	Case "Special": specialid=request("P" & i)
	BLine=26
	Case "Reward Points": rwpid=request("P" & i)
	BLine=51
	Case "Non-taxable": ntaxid=request("P" & i)
	BLine=27
	Case "No shipping charge": nshipid=request("P" & i)
	BLine=28
	Case "Not for sale": nforsaleid=request("P" & i)
	BLine=29
	Case "Not for sale copy": nforsalecopyid=request("P" & i)
	BLine=30
	Case "Disregard stock": distockid=request("P" & i)
	BLine=31
	Case "Display No Shipping Text": dishiptextid=request("P" & i)
	BLine=32
	Case "Custom Search Field (1)": customfieldsid(0)=request("P" & i)
	BLine=33
	Case "Custom Search Field (2)": customfieldsid(1)=request("P" & i)
	BLine=34
	Case "Custom Search Field (3)": customfieldsid(2)=request("P" & i)
	BLine=35
	Case "Downloadable File Location": fileurlid=request("P" & i)
	BLine=36
	Case "Make Download URL expire": urlexpireid=request("P" & i)
	BLine=37
	Case "URL Expiration in Days": expiredaysid=request("P" & i)
	BLine=38
	Case "Use License Generator": licenseid=request("P" & i)
	BLine=39
	Case "Local Generator": LocalLGid=request("P" & i)
	BLine=40
	Case "Remote Generator": RemoteLGid=request("P" & i)
	BLine=41
	Case "License Field Label (1)": LFN1id=request("P" & i)
	BLine=42
	Case "License Field Label (2)": LFN2id=request("P" & i)
	BLine=43
	Case "License Field Label (3)": LFN3id=request("P" & i)
	BLine=44
	Case "License Field Label (4)": LFN4id=request("P" & i)
	BLine=45
	Case "License Field Label (5)": LFN5id=request("P" & i)
	BLine=46
	Case "Additional copy": Addcopyid=request("P" & i)
	BLine=47
	Case "Short Category Description": SCATDescid=request("P" & i)
	BLine=48
	Case "Long Category Description": LCATDescid=request("P" & i)
	BLine=49
	Case "Minimum Quantity customers can buy": MQtyID=request("P" & i)
	BLine=50
	Case "Oversized Product Details": OverSizeID=request("P" & i)
	BLine=52
	Case "Option 1": Opt1ID=request("P" & i)
	BLine=53
	Case "Attributes 1": Attr1ID=request("P" & i)
	BLine=54
	Case "Option 1 Required": Opt1ReqID=request("P" & i)
	BLine=55
	Case "Option 2": Opt2ID=request("P" & i)
	BLine=56
	Case "Attributes 2": Attr2ID=request("P" & i)
	BLine=57
	Case "Option 2 Required": Opt2ReqID=request("P" & i)
	BLine=58
	
	'**** Apparel Product Fields: 59 - 67
	
	Case "Force purchase of multiples of minimum": VQtyID=request("P" & i)
	BLine=68
	
	'Start SDBA
	Case "Product Cost": prd_CostID=request("P" & i)
	BLine=69
	Case "Back Order": prd_BackOrderID=request("P" & i)
	BLine=72
	Case "Ship within N Days": prd_ShipNDaysID=request("P" & i)
	BLine=73
	Case "Low inventory notification": prd_NotifyStockID=request("P" & i)
	BLine=74
	Case "Reorder Level": prd_ReorderLevelID=request("P" & i)
	BLine=75
	Case "Is Drop-shipped": prd_IsDropShippedID=request("P" & i)
	BLine=76
	Case "Supplier ID": prd_SupplierID=request("P" & i)
	BLine=77
	Case "Drop-Shipper ID": prd_DropShipperID=request("P" & i)
	BLine=78
	Case "Drop-Shipper is also a Supplier": prd_IsDropShipperID=request("P" & i)
	BLine=79
	'End SDBA
	
	Case "Option 1 Order": Opt1OrderID=request("P" & i)
	BLine=80
	Case "Option 2 Order": Opt2OrderID=request("P" & i)
	BLine=81
	Case "Option 3": Opt3ID=request("P" & i)
	BLine=82
	Case "Attributes 3": Attr3ID=request("P" & i)
	BLine=83
	Case "Option 3 Required": Opt3ReqID=request("P" & i)
	BLine=84
	Case "Option 3 Order": Opt3OrderID=request("P" & i)
	BLine=85
	Case "Option 4": Opt4ID=request("P" & i)
	BLine=86
	Case "Attributes 4": Attr4ID=request("P" & i)
	BLine=87
	Case "Option 4 Required": Opt4ReqID=request("P" & i)
	BLine=88
	Case "Option 4 Order": Opt4OrderID=request("P" & i)
	BLine=89
	Case "Option 5": Opt5ID=request("P" & i)
	BLine=90
	Case "Attributes 5": Attr5ID=request("P" & i)
	BLine=91
	Case "Option 5 Required": Opt5ReqID=request("P" & i)
	BLine=92
	Case "Option 5 Order": Opt5OrderID=request("P" & i)
	BLine=93
	Case "Downloadable Product": downprdID=request("P" & i)
	BLine=94
	Case "Gift Certificate": giftcertID=request("P" & i)
	BLine=95
	Case "Gift Certificate Expiration": giftexpID=request("P" & i)
	BLine=96
	Case "Electronic Only (Gift Certificate)": giftelectID=request("P" & i)
	BLine=97
	Case "Use Generator (Gift Certificate)": giftgenID=request("P" & i)
	BLine=98
	Case "Expiration Date (Gift Certificate)": giftexpdateID=request("P" & i)
	BLine=99
	Case "Expire N days (Gift Certificate)": giftexpdaysID=request("P" & i)
	BLine=100
	Case "Custom Generator Filename (Gift Certificate)": giftcustgenfileID=request("P" & i)
	BLine=101
	Case "Hide BTO Price": hidebtopriceid=request("P" & i)
	BLine=102
	Case "Hide Default Configuration": hideconfid=request("P" & i)
	BLine=103
	Case "Disallow purchasing": dispurchaseid=request("P" & i)
	BLine=104
	Case "Skip Product Details Page": skipdetailsid=request("P" & i)
	BLine=105
	Case "Short Category Description 1": SCatDesc1id=request("P" & i)
	BLine=106
	Case "Long Category Description 1": LCatDesc1id=request("P" & i)
	BLine=107
	Case "Category Small Image 1": csimage1id=request("P" & i)
	BLine=108
	Case "Category Large Image 1": climage1id=request("P" & i)
	BLine=109
	Case "Short Category Description 2": SCatDesc2id=request("P" & i)
	BLine=110
	Case "Long Category Description 2": LCatDesc2id=request("P" & i)
	BLine=111
	Case "Category Small Image 2": csimage2id=request("P" & i)
	BLine=112
	Case "Category Large Image 2": climage2id=request("P" & i)
	BLine=113
	Case "Mega Tags - Title": mt_titleID=request("P" & i)
	BLine=114
	Case "Mega Tags - Description": mt_descID=request("P" & i)
	BLine=115
	Case "Mega Tags - Keywords": mt_keyID=request("P" & i)
	BLine=116
	Case "Featured": featuredID=request("P" & i)
	BLine=117
	Case "Units to make 1 lb": unitslbID=request("P" & i)
	BLine=118
	
	Case "First Unit Surcharge": surcharge1ID=request("P" & i)
	BLine=119
	
	Case "Additional Unit(s) Surcharge": surcharge2ID=request("P" & i)
	BLine=120
	
	Case "Product Notes": prdnoteID=request("P" & i)
	BLine=121
	
	Case "Enable Image Magnifier": eimagID=request("P" & i)
	BLine=122
	
	Case "Page Layout": playoutID=request("P" & i)
	BLine=123
	
	Case "Hide SKU on the product details page": hideskuID=request("P" & i)
	BLine=124
	
	Case "Google Product Category": goCatid=request("P" & i)
	BLine=125
	
	Case "Google Shopping - Gender": goGenid=request("P" & i)
	BLine=126
	
	Case "Google Shopping - Age": goAgeid=request("P" & i)
	BLine=127
	
	Case "Google Shopping - Color": goColorid=request("P" & i)
	BLine=128
	
	Case "Google Shopping - Size": goSizeid=request("P" & i)
	BLine=129
	
	Case "Google Shopping - Pattern": goPatid=request("P" & i)
	BLine=130
	
	Case "Google Shopping - Material": goMatid=request("P" & i)
	BLine=131
	
	End Select
	if BLine>0 then
	TempStr=request("F" & i) & "*****"
	if instr(ALines(BLine-1),TempStr)=0 then
	ALines(BLine-1)=ALines(BLine-1) & TempStr
	end if
	BLine=0
	end if
Next

	SavedFile = "importlogs/save.txt"
	findit = Server.MapPath(Savedfile)
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	Err.number=0
	Set f = fso.OpenTextFile(findit, 2)
	For dd=lbound(ALines) to ubound(ALines)
	f.WriteLine ALines(dd)
	Next
	f.close
%>
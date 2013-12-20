<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=0%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<html>
<head>
<title>Test License Generator</title>
</head>
<body>

<% 
on error resume next

pIdProduct=request.Querystring("idProduct")
if not validNum(pIdProduct) then response.redirect "menu.asp"

dim query, conntemp, rs
call openDB()

query="select * from DProducts where Idproduct=" & pIdproduct
set rs=Server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query)

pLicense=rs("License")
pLocalLG=rs("LocalLG")
pRemoteLG=rs("RemoteLG")
pLL1=rs("LicenseLabel1")
pLL2=rs("LicenseLabel2")
pLL3=rs("LicenseLabel3")
pLL4=rs("LicenseLabel4")
pLL5=rs("LicenseLabel5")

if (pLicense="0") or (pLicense="") then%>
	<font face="Arial" size=2 color=#FF0000>You did not specify a License Generator for this product. If you have entered a value in one of the License Generator fields and received this message, save the product information, then click on 'Modify this product again' on the confirmation page. At that time you will be able to use the 'Test License Generator' feature.</font><br>
<%
else
	if pLocalLG<>"" then
	SPath1=Request.ServerVariables("PATH_INFO")
	mycount1=0
	do while mycount1<1
	if mid(SPath1,len(SPath1),1)="/" then
	mycount1=mycount1+1
	end if
	if mycount1<1 then
	SPath1=mid(SPath1,1,len(SPath1)-1)
	end if
	loop
	SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & SPath1
	
	if Right(SPathInfo,1)="/" then
	pLocalLG=SPathInfo & "licenses/" & pLocalLG					
	else
	pLocalLG=SPathInfo & "/licenses/" & pLocalLG
	end if
	L_Action=pLocalLG
	else
	L_Action=pRemoteLG
	end if
	
	'Demo Date
	pIdOrder=5236
	pOrderDate=date()
	pProcessDate=date()
	pIdCustomer=125
	pQuantity=1
		
	query="select sku from products where Idproduct=" & pIdproduct
	set rs=connTemp.execute(query)
	pSKU=rs("sku")
	set rs=nothing
	
	L_postdata=""
	L_postdata=L_postdata&"idorder=" & pIdOrder
	L_postdata=L_postdata&"&orderDate=" & pOrderDate
	L_postdata=L_postdata&"&ProcessDate=" & pProcessDate
	L_postdata=L_postdata&"&idcustomer=" & pIdCustomer
	L_postdata=L_postdata&"&idproduct=" & pIdproduct
	L_postdata=L_postdata&"&quantity=" & pQuantity
	L_postdata=L_postdata&"&sku=" & pSKU

	Set srvXmlHttp = server.createobject("Msxml2.serverXmlHttp")
	srvXmlHttp.open "POST", L_Action, False
	srvXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	srvXmlHttp.send L_postdata
	result1 = srvXmlHttp.responseText
	myErr=""
	if err.number<>0 then
		myErr=err.Description
		err.number=0
		else
		if instr(ucase(result1),"HTTP ")>0 then
			myErr="The License Generator cannot be found. Please check its location and try again."
		else
			AR=split(result1,"<br>")
			rIdOrder=AR(0)
			rIdProduct=AR(1)
			Lic1=split(AR(2),"***")
			Lic2=split(AR(3),"***")
			Lic3=split(AR(4),"***")
			Lic4=split(AR(5),"***")
			Lic5=split(AR(6),"***")
		end if
	end if%>
	<font face="Arial" size="4">
	<b>Testing License Generator</b></font><font face=Arial size=2><br><br>
	<b>Demo Data:</b><br>
	<br>
	Order ID: #<%=pIdOrder%><br>
	Order Date: <%=pOrderDate%><br>
	Process Date: <%=pProcessDate%><br>
	Customer ID: <%=pIdCustomer%><br>
	Product ID: <%=pIdProduct%><br>
	Quantity: <%=pQuantity%><br>
	Product SKU: <%=psku%><br>
	<br>
	<b>Returns:</b><br><br />
	<%if myErr<>"" then%>
	<font color=#FF0000><%=myErr%></font>
	<%else%>
	Order ID: #<%=rIdOrder%><br>
	Product ID: <%=rIdProduct%><br>
	<%if Lic1(0)<>"" then%>
	<%=pLL1%>: <%=Lic1(0)%><br>
	<%end if%>
	<%if Lic2(0)<>"" then%>
	<%=pLL2%>: <%=Lic2(0)%><br>
	<%end if%>
	<%if Lic3(0)<>"" then%>
	<%=pLL3%>: <%=Lic3(0)%><br>
	<%end if%>
	<%if Lic4(0)<>"" then%>
	<%=pLL4%>: <%=Lic4(0)%><br>
	<%end if%>
	<%if Lic5(0)<>"" then%>
	<%=pLL5%>: <%=Lic5(0)%><br>
	<%end if%>		
	<%end if%>
	</font>
	<%
	end if
set rs=nothing
call closeDb()
%>
</body>
</html>
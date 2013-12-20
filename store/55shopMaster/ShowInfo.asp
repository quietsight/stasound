<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include FILE="../includes/languages.asp"-->
<!--#include FILE="../includes/ErrorHandler.asp"--> 
<%  
Dim conntemp

call openDB()

Dim ProductArray, i, mySQL, rs
categoryDescName = getUserInput(Request.QueryString("cd"),200)
ProductArray = getUserInput(Request.QueryString("SIArray"),0)
ProductArray = Split(ProductArray,",")
%>
<html>
<head>
<title>More details for <%=categoryDescName%></title>

<script language="JavaScript">
<!--
imagename = '';
function enlrge(imgnme) {
	lrgewin = window.open("about:blank","","height=200,width=200")
	imagename = imgnme;
	setTimeout('update()',500)
}
function win(fileName)
	{
	myFloater = window.open('','myWindow','scrollbars=auto,status=no,width=400,height=300')
	myFloater.location.href = fileName;
	}
	function viewWin(file)
	{
	myFloater = window.open('','myWindow','scrollbars=yes,status=no,width=400,height=400')
	myFloater.location.href = file;
	}
function update() {
doc = lrgewin.document;
doc.open('text/html');
doc.write('<HTML><HEAD><TITLE>Enlarged Image<\/TITLE><\/HEAD><BODY bgcolor="white" onLoad="if (document.all || document.layers) window.resizeTo((document.images[0].width + 10),(document.images[0].height + 80))" topmargin="4" leftmargin="0" rightmargin="0" bottommargin="0"><table width=""' + document.images[0].width + '" height="' + document.images[0].height +'"border="0" cellspacing="0" cellpadding="0"><tr><td>');
doc.write('<IMG SRC="' + imagename + '"><\/td><\/tr><tr><td><form name="viewn"><A HREF="javascript:window.close()"><img  src="../pc/images/close.gif" align="right" border=0><\/a><\/td><\/tr><\/table>');
doc.write('<\/form><\/BODY><\/HTML>');
doc.close();
}

//-->
</script>
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
</head>
<body>
<div id="pcMain">
<!--#include file="../includes/javascripts/pcWindowsViewPrd.asp"-->
<%
	for i = lbound(ProductArray) to (UBound(ProductArray)-1)
		if validNum(ProductArray(i)) then	
			query="SELECT products.description, products.smallImageUrl, products.largeImageURL, products.details FROM products WHERE products.idProduct="& ProductArray(i) &";"
			set rs=server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)
			If NOT rs.eof then
				'// Assign variables	
				'// Make the popup link, but dont set large image preference if the large image doesnt exist
				pcv_productName=rs("description")
				pcv_strShowImage_Url = rs("smallImageUrl")
				pcv_strShowImage_LargeUrl = rs("largeImageURL")
				If len(pcv_strShowImage_LargeUrl)>0 Then		
					pcv_strLargeUrlPopUp= "javascript:pcAdditionalImages('../pc/catalog/"&pcv_strShowImage_LargeUrl&"','"&ProductArray(i)&"')" 
				Else
					pcv_strShowImage_LargeUrl = pcv_strShowImage_Url '// we dont have one, show the regular size
					pcv_strLargeUrlPopUp= "javascript:pcAdditionalImages('../pc/catalog/"&pcv_strShowImage_Url&"','"&ProductArray(i)&"')" 
				End If
				pcv_productDetails=rs("details")
				%>
					<table class="pcBTOpopup">
							<tr> 
								<td colspan="2">
									<div class="pcSectionTitle"><%=pcv_productName%></div>
    </td>
  </tr>
  <tr> 
					<% if iBTOPopImage=1 then %>
								<td valign="top" colspan="2" style="padding: 10px;">
								<p><%=pcv_productDetails%></p>
							</td>
							<% else %>
							<td valign="top" width="70%" style="padding: 10px;">
								<p><%=pcv_productDetails%></p>
							</td>
								<td valign="top" width="30%"> 
								<% if pcv_strShowImage_Url<>"" then %>
									<% if pcv_strShowImage_LargeUrl<>"" then %>
										<a href="javascript:enlrge('../pc/catalog/<%=pcv_strShowImage_LargeUrl%>')">
										<img class="ProductThumbnail" alt="<%=pcv_productName%>" src="../pc/catalog/<%=pcv_strShowImage_Url%>" align="right">
										</a> 
					<% else %>
										<img src="../pc/catalog/<%=pcv_strShowImage_Url%>" alt="<%=pcv_productName%>">
									<% end if %>	
									<% 
									'// If there are additional images.
									query = "SELECT pcProductsImages.idProduct FROM pcProductsImages WHERE pcProductsImages.idProduct=" & ProductArray(i) &";"
									set rs=server.createobject("adodb.recordset")
									set rs=conntemp.execute(query)	
									if err.number<>0 then
										call LogErrorToDatabase()
										set rs=nothing
										call closedb()
										response.redirect "techErr.asp?err="&pcStrCustRefID
									end if				
									If NOT rs.EOF Then					 
									%>
										<div align="right"><a href="<%=pcv_strLargeUrlPopUp%>"><%=dictLanguage.Item(Session("language")&"_ShowInfo_1")%></a></div>
									<% end if %>
					<% end if %>
    </td>
							<% end if %>
  </tr>
</table>
<%
end if
			end if ' ValidNum
		next
set rs = nothing
call closeDB()
%>
	<div align="right">
		<A HREF="javascript:window.close()"><img  src="../pc/images/close.gif" border="0"></a>
	</div>
</div>
</body>
</html>
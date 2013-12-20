<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/adovbs.inc"-->
<%
on error resume next
Dim rsAddImgObj, pcv_recordCount, iPageCount

if request("iPageCurrent")="" then
    iPageCurrent=1 
else
    iPageCurrent=Request("iPageCurrent")
end If

call openDb()

' // SELECT DATA SET
' TABLES: pcProductsImages
' COLUMNS ORDER: pcProductsImages.pcProdImage_Url, pcProductsImages.pcProdImage_LargeUrl, pcProductsImages.pcProdImage_Order

query = 		"SELECT pcProdImage_ID, pcProdImage_Url, pcProdImage_LargeUrl, pcProdImage_Order "
query = query & "FROM pcProductsImages "
query = query & "WHERE idProduct=" & pIdProduct &" "
query = query & "ORDER BY pcProdImage_Order;"	
set rsAddImgObj=server.createobject("adodb.recordset")

pcv_recordCount=0
iPageCount=1

rsAddImgObj.CacheSize=10
rsAddImgObj.PageSize=10
rsAddImgObj.Open query, conntemp, adOpenStatic, adLockReadOnly
pcv_recordCount = rsAddImgObj.RecordCount

if pcv_recordCount>0 then

    rsAddImgObj.MoveFirst

    ' get the max number of pages
    iPageCount=rsAddImgObj.PageCount
    If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
    If iPageCurrent < 1 Then iPageCurrent=1

    ' set the absolute page
    rsAddImgObj.AbsolutePage=iPageCurrent
end if

if fromPage="" then
	fromPage = "modifyProduct.asp"
end If

%>
<script language="JavaScript">
<!--
	imagename = '';
	function enlrge(imgnme) {
		lrgewin = window.open("about:blank","","height=200,width=200")
		imagename = imgnme;
		setTimeout('update()',500)
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

function newAddWindow(file,window) {
	addWindow=open(file,window,'resizable=no,width=500,height=400,scrollbars=1');
	if (addWindow.opener == null) addWindow.opener = self;
}

//-->
</script>

<table class="pcCPcontent" style="width:100%; border: 1px solid #CCCCCC">
    <tr>
	    <td colspan="3">There are <%=pcv_recordCount%> additional product images assigned to this product.</td>
    </tr>
    <tr>
        <td width="40%" nowrap><b>General Image</b></td>
        <td width="40%" nowrap><b>Detail View Image</b></td>
        <td nowrap align="right">&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=425')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
    </tr>                  
		<tr>
			<td colspan="3" class="pcCPspacer" style="border-top: 1px solid #CCCCCC"></td>
		</tr>
                      
<% If rsAddImgObj.eof Then %>
                      
    <tr> 
       <td colspan="3">No Additional Product Images Found</td>
    </tr>
                      
<%
	else 
	    Dim Count
	    Count=0
	    Do While NOT rsAddImgObj.EOF And Count < rsAddImgObj.PageSize
%>
                      
	<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
        <td>
            <% if rsAddImgObj("pcProdImage_Url") <> "" then%>
	            <a href="javascript:enlrge('../pc/catalog/<%=rsAddImgObj("pcProdImage_Url")%>')">
	                <img src="../pc/catalog/<%=rsAddImgObj("pcProdImage_Url")%>" align=absbottom class="pcShowProductImageM">
	            </a>
                <a href="javascript:enlrge('../pc/catalog/<%=rsAddImgObj("pcProdImage_Url") %>')"><%= rsAddImgObj("pcProdImage_Url") %></a>
            <% end if %>
        </td>
        <td>
            <% if rsAddImgObj("pcProdImage_LargeUrl") <> "" then%>        
	            <a href="javascript:enlrge('../pc/catalog/<%=rsAddImgObj("pcProdImage_LargeUrl")%>')">
	                <img src="../pc/catalog/<%=rsAddImgObj("pcProdImage_LargeUrl")%>" align=absbottom class="pcShowProductImageM">
	            </a>
                <a href="javascript:enlrge('../pc/catalog/<%=rsAddImgObj("pcProdImage_LargeUrl") %>')"><%= rsAddImgObj("pcProdImage_LargeUrl") %></a>
            <% end if %>
        </td>
        <td align="right" nowrap class="cpLinkslist">
            <a href="javascript:newAddWindow('addImg_popup.asp?idproduct=<%= pIdProduct %>&imgid=<%= rsAddImgObj("pcProdImage_ID") %>','addwindow')">Edit</a>&nbsp;|&nbsp;<a href="javascript:if (confirm('You are about to remove these images from this product. The actual files will not be deleted from the Web server. Do you want to continue?')) location='delPrdAddImg.asp?idproduct=<%= pIdProduct %>&pid=<%= rsAddImgObj("pcProdImage_ID") %>&timg=<%= rsAddImgObj("pcProdImage_Url") %>&dimg=<%= rsAddImgObj("pcProdImage_LargeUrl") %>&redir=<%=fromPage %>'">Delete</a>
        </td>
    </tr>
<% 
            count=count + 1
		    rsAddImgObj.MoveNext
		Loop
	end If
	set rsAddImgObj = nothing
%>                
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
    <tr>
			<td colspan="3">
			<% If iPageCount>1 Then %>
					<p><%response.Write("Currently viewing page "& iPageCurrent & " of "& iPageCount )%></p>
					<p>
			<%' display Next / Prev buttons
			if iPageCurrent > 1 then %>
					<a href="<%=fromPage %>?idproduct=<%= pIdProduct %>&iPageCurrent=<%=iPageCurrent-1%>&prdType=<%=pcv_ProductType%>"><img src="../pc/images/prev.gif"></a> 
			<% end If
			For I=1 To iPageCount
					If Cint(I)=Cint(iPageCurrent) Then %>
					<b><%=I%></b> 
				<% Else %>
					<a href="<%=fromPage %>?idproduct=<%= pIdProduct %>&iPageCurrent=<%=I%>&prdType=<%=pcv_ProductType%>"> 
					<%=I%></a> 
				<% End If %>
			<% Next %>
			<% if CInt(iPageCurrent) < CInt(iPageCount) then %>
				<a href="<%=fromPage %>?idproduct=<%= pIdProduct %>&iPageCurrent=<%=iPageCurrent+1%>&prdType=<%=pcv_ProductType%>"> <img src="../pc/images/next.gif"></a> 
			<% end If %>
			</p>
			<% End If %>
			</td>
    </tr>
</table>
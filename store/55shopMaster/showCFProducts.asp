<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>ProductCart shopping cart software - Control Panel - Manage Custom Fields</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="description" content="ProductCart asp shopping cart software is published by NetSource Commerce. ProductCart's Control Panel allows you to manage every aspect of your ecommerce store. For more information and for technical support, please visit NetSource Commerce at http://www.earlyimpact.com">
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="background-image: none;">
<%
dim query, conntemp, rs

if request("iPageCurrent")="" then
    iPageCurrent=1 
else
    iPageCurrent=Request("iPageCurrent")
end If

'sorting order
Dim strORD

strORD=request("order")
if strORD="" then
	strORD="products.description"
End If

strSort=request("sort")
if strSort="" Then
	strSort="ASC"
End If

idcustom=mid(request("idcustom"),2,len(request("idcustom")))
idcustom1=request("idcustom")
idcustomType=Left(request("idcustom"),1)

call openDb()

' Remove a custom field from the database:

pcv_strAction = request.QueryString("action")
pcv_intIdCustom = request.QueryString("idcustom")
pcv_strCustomType = request.QueryString("type")
if pcv_strAction = "del" then
	if pcv_strCustomType = "S" then
		query="DELETE FROM pcSearchFields WHERE idSearchField="&pcv_intIdCustom
		set rs=connTemp.execute(query)
		msg="Custom search field deleted successfully!"
		set rs = nothing
	else
		query="DELETE FROM xfields WHERE idxfield="&pcv_intIdCustom
		set rs=connTemp.execute(query)
		msg="Custom inpout field deleted successfully!"
		set rs = nothing
	end if %>

<table class="pcCPcontent" style="width: 100%;">
<tr> 
	<td>
	<div class="pcCPmessageSuccess">The field was deleted successfully</div>
	</td>
</tr>
<tr>
<td align="right">&nbsp;</td>
</tr>
<tr>
<td align="right"><input type="button" name="back" value="Close window" onClick="opener.location.href='ManageCFields.asp#1'; window.close();" class="ibtnGrey">
</td>
</tr>
</table>

<% END IF ' End remove custom field

IF pcv_strAction <> "del" THEN ' If not deleting the field, load the main page

	query=""
	if Left(request("idcustom"),1)="C" then
		query=" AND ((xfield1=" & idcustom & ") OR (xfield2=" & idcustom & ") OR (xfield3=" & idcustom & "))"
		query="SELECT * FROM products WHERE removed=0 AND configOnly=0 " & query & " ORDER BY "& strORD &" "& strSort
	else
		if request("SearchValues")<>"" then
			idvalue=request("SearchValues")
			query="SELECT products.* FROM Products,pcSearchFields_Products WHERE products.removed=0 AND products.configOnly=0 AND products.idproduct=pcSearchFields_Products.idproduct AND pcSearchFields_Products.idSearchData=" & idvalue & " ORDER BY "& strORD &" "& strSort
		else
			if trim(idcustom)="" then
				idcustom=request.QueryString("idcustom")
			end if
			query="SELECT * FROM Products WHERE products.removed=0 AND products.configOnly=0 AND (products.idproduct IN (SELECT DISTINCT pcSearchFields_Products.idProduct FROM pcSearchFields_Products INNER JOIN pcSearchData ON pcSearchFields_Products.idSearchData=pcSearchData.idSearchData WHERE pcSearchData.idSearchField=" & idcustom & "))" & " ORDER BY "& strORD &" "& strSort
		end if
	end if
	
	Set rsInv=Server.CreateObject("ADODB.Recordset")
	rsInv.CacheSize=25
	rsInv.PageSize=25

	rsInv.Open query, connTemp, adOpenStatic, adLockReadOnly
	pcRecordCount = rsInv.RecordCount
	%>
										
    
    
    <% If rsInv.eof Then %>
       
        <table class="pcCPcontent" style="width: 100%;">
            <tr> 
                <td>
                    <p>None of the products in the store catalog are using this custom field.</p>
                    <p>Both active and inactive products were included in the search.</p>
                    <p><a href="showCFProducts.asp?action=del&idcustom=<%=idcustom%>&type=<%=idcustomType%>">Remove this custom field</a> from the database.</p></td>
            </tr>
            <tr>
                <td align="right">&nbsp;</td>
            </tr>
            <tr>
                <td align="right"><input type="button" name="back" value="Close window" onClick="window.close();" class="ibtnGrey"></td>
            </tr>
        </table>

	<% Else %>
       
        <table class="pcCPcontent" style="width: 100%;">                                   
            <tr>
            	<td colspan="2">There are <%=pcRecordCount%> product(s) using the selected search field:</font></td>
            </tr>
            <tr> 
            	<td colspan="2" class="pcCPspacer"></td>
            </tr>
            <tr>
            	<th nowrap><a href="showCFProducts.asp?idcustom=<%=request("idcustom")%>&SearchValues=<%=request("SearchValues")%>&iPageCurrent=<%=iPageCurrent%>&order=products.sku&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="showCFProducts.asp?idcustom=<%=request("idcustom")%>&SearchValues=<%=request("SearchValues")%>&iPageCurrent=<%=iPageCurrent%>&order=products.sku&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a> SKU</th>
           		<th nowrap><a href="showCFProducts.asp?idcustom=<%=request("idcustom")%>&SearchValues=<%=request("SearchValues")%>&iPageCurrent=<%=iPageCurrent%>&order=products.description&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="showCFProducts.asp?idcustom=<%=request("idcustom")%>&SearchValues=<%=request("SearchValues")%>&iPageCurrent=<%=iPageCurrent%>&order=products.description&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a> Product</th>
            </tr>
            <tr> 
            	<td colspan="2" class="pcCPspacer"></td>
            </tr>
                                    
			<%	rsInv.MoveFirst
                        
				' get the max number of pages
				Dim iPageCount
				iPageCount=rsInv.PageCount
				If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
				If iPageCurrent < 1 Then iPageCurrent=1
					
				' set the absolute page
				rsInv.AbsolutePage=iPageCurrent  
				Count=0
				Do While NOT rsInv.EOF And Count < rsInv.PageSize 
					count=count + 1
					%>
					<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
						<td width="30%"><%=rsInv("sku")%></td>
						<td width="70%"><a href="JavaScript:window.opener.location.href='AdminCustom.asp?idproduct=<%=rsInv("idproduct")%>'; self.close();"><%=rsInv("description")%></a></td>
					</tr>
					<% 
					rsInv.MoveNext
				Loop
				%>
            <tr> 
            	<td colspan="4" class="pcCPspacer"></td>
            </tr>
            <tr>
                <td colspan="4" align="right"><input type="button" name="back" value="Close Window" onClick="window.close();" class="ibtnGrey"></td>
            </tr>
        </table>
    
	<% End If %> 
                
	<% If iPageCount>1 Then %>           
      <table class="pcCPcontent" style="width: 100%;">
          <tr>
              <td><%response.Write("Currently viewing page "& iPageCurrent & " of "& iPageCount)%></td></tr>
          <tr> 
              <td> 
                <%' display Next / Prev buttons
                if iPageCurrent > 1 then %>
                  <a href="showCFProducts.asp?idcustom=<%=request("idcustom")%>&SearchValues=<%=request("SearchValues")%>&iPageCurrent=<%=iPageCurrent-1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/prev.gif" width="10" height="10" border="0"></a> 
                <% 
                end If
                
                For I=1 To iPageCount
                  If Cint(I)=Cint(iPageCurrent) Then %>
                      <b><%=I%></b>
                   <% Else %>
                      <a href="showCFProducts.asp?idcustom=<%=request("idcustom")%>&SearchValues=<%=request("SearchValues")%>&iPageCurrent=<%=I%>&order=<%=strORD%>&sort=<%=strSort%>"><%=I%></a> 
                  <% End If %>
                <% Next %>
                <% if CInt(iPageCurrent) < CInt(iPageCount) then %>
                      <a href="showCFProducts.asp?idcustom=<%=request("idcustom")%>&SearchValues=<%=request("SearchValues")%>&iPageCurrent=<%=iPageCurrent+1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
                <% end If %>
              </td>
          </tr>
      </table>
    <% End If

END IF '// if pcv_strAction <> "del" then 
%>
</body>
</html>
<%call closedb()%>
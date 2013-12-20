<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/validation.asp" --> 
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/currencyformatinc.asp" -->
<!--#include file="inc_UpdateDates.asp" -->
<%

Dim rsOrd, connTemp, strSQL, pid, rstemp, query
call openDb()

if request("iPageCurrent")="" then
    iPageCurrent=1 
else
    iPageCurrent=Request("iPageCurrent")
end If

if request("pagesize")<>"" then
	session("pagesize")=request("pagesize")
end if
pcv_pagesize=session("pagesize")
if not validNum(pcv_pagesize) then
	pcv_pagesize=25
end if

'// Filter by category
Dim pcIntCategoryID, queryCat
pcIntCategoryID=request("idcategory")
	if not validNum(pcIntCategoryID) then
		pcIntCategoryID=request("idcat")
	end if
if validNum(pcIntCategoryID) then
	queryCat=" AND products.idproduct IN (SELECT DISTINCT categories_products.idproduct FROM categories_products WHERE categories_products.idcategory=" & pcIntCategoryID & ")"
	' Get Category Name:
	query="SELECT categoryDesc FROM categories WHERE idCategory="&pcIntCategoryID
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)
	pcStrCategoryName=rstemp("categoryDesc")
	set rstemp=nothing
end if


'// Sorting Order
Dim strORD

strORD=request("order")
if strORD="" then
	strORD="description"
End If

strSort=request("sort")
if strSort="" Then
	strSort="ASC"
End If 

if request("action")="update" then
 count=request("count")

 for i=1 to count
 	if request("C" & i)="1" then
		dblPrice=replacecomma(request("price" & i))
		if dblPrice="" then
		dblPrice=0
		end if
		if IsNumeric(dblPrice)=false then
		dblPrice=0
		end if
		
		dblLPrice=replacecomma(request("lprice" & i))
		if dblLPrice="" then
		dblLPrice=0
		end if
		if IsNumeric(dblLPrice)=false then
		dblLPrice=0
		end if
		
		dblWPrice=replacecomma(request("wprice" & i))
		if dblWPrice="" then
		dblWPrice=0
		end if
		if IsNumeric(dblWPrice)=false then
		dblWPrice=0
		end if
		
		intListHidden=request("shows" & i)
		if intListHidden="" then
		intListHidden=0
		end if
		if IsNumeric(intListHidden)=false then
		intListHidden=0
		end if
		
		query="UPDATE products SET price=" & dblPrice & ",listPrice=" & dblLPrice & ",bToBPrice=" & dblWPrice & ",listHidden=" & intListHidden & " WHERE idproduct="& request("ID" & i)
		set rstemp=Server.CreateObject("ADODB.Recordset")
		Set rstemp=conntemp.execute(query)
		
		call updPrdEditedDate(request("ID" & i))
  end if
 next 
end if

if validNum(pcIntCategoryID) then
	pageTitle="Update Product Prices in <strong>" & pcStrCategoryName & "</strong>"
	else
	pageTitle="Update Product Prices"
end if	
%>
<!--#include file="AdminHeader.asp"-->
<!--#include file="pcv4_showMessage.asp"-->
<form method="POST" name="checkboxform" action="updPrdPrices.asp?action=update&iPageCurrent=<%=request("iPageCurrent")%>&order=<%=request("order")%>&sort=<%=request("sort")%>" onSubmit="return Form1_Validator(this)" class="pcForms">
    <div style="padding: 10px;">
        Products per page: <select name="pagesize" onchange="location='updPrdPrices.asp?pagesize=' + document.checkboxform.pagesize.value + '&order=<%=request("order")%>&sort=<%=request("sort")%>';">
        <option value="25" selected>25</option>
        <option value="50" <%if pcv_pagesize="50" then%>selected<%end if%>>50</option>
        <option value="100" <%if pcv_pagesize="100" then%>selected<%end if%>>100</option>
        </select>

        &nbsp;
        Only show products from:
        <%
		
		cat_DropDownName="idcat"
		cat_Type="1"
		cat_DropDownSize="1"
		cat_MultiSelect="0"
		cat_ExcBTOHide="0"
		cat_StoreFront="0"
		cat_ShowParent="1"
		cat_DefaultItem=""
		cat_SelectedItems="" & pcIntCategoryID & ","
		cat_ExcItems=""
		cat_ExcSubs="0"
		cat_EventAction="onchange=""location='updPrdPrices.asp?idcat=' + document.checkboxform.idcat.value + ''"""
		%>
		<!--#include file="../includes/pcCategoriesList.asp"-->
		<%call pcs_CatList()%>
        
    </div>
    <table class="pcCPcontent">
        <tr> 
            <th nowrap><a href="updPrdPrices.asp?iPageCurrent=<%=iPageCurrent%>&order=sku&sort=ASC&idcat=<%=pcIntCategoryID%>"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="updPrdPrices.asp?iPageCurrent=<%=iPageCurrent%>&order=sku&sort=Desc&idcat=<%=pcIntCategoryID%>"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;SKU</th>
            <th nowrap><a href="updPrdPrices.asp?iPageCurrent=<%=iPageCurrent%>&order=description&sort=ASC&idcat=<%=pcIntCategoryID%>"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="updPrdPrices.asp?iPageCurrent=<%=iPageCurrent%>&order=description&sort=Desc&idcat=<%=pcIntCategoryID%>"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Product</th>
            <th nowrap><a href="updPrdPrices.asp?iPageCurrent=<%=iPageCurrent%>&order=price&sort=ASC&idcat=<%=pcIntCategoryID%>"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="updPrdPrices.asp?iPageCurrent=<%=iPageCurrent%>&order=price&sort=Desc&idcat=<%=pcIntCategoryID%>"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Online Price</th>
            <th nowrap><a href="updPrdPrices.asp?iPageCurrent=<%=iPageCurrent%>&order=listPrice&sort=ASC&idcat=<%=pcIntCategoryID%>"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="updPrdPrices.asp?iPageCurrent=<%=iPageCurrent%>&order=listPrice&sort=Desc&idcat=<%=pcIntCategoryID%>"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;List Price</th>
            <th nowrap><a href="updPrdPrices.asp?iPageCurrent=<%=iPageCurrent%>&order=bToBPrice&sort=ASC&idcat=<%=pcIntCategoryID%>"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="updPrdPrices.asp?iPageCurrent=<%=iPageCurrent%>&order=bToBPrice&sort=Desc&idcat=<%=pcIntCategoryID%>"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;W. Price</th>
            <th>Savings</th>
            <th>Update</th>
        </tr>
        <tr>                     
            <td colspan="7" class="pcCPspacer"></td>
        </tr>
                      
<% 
query="SELECT idproduct,sku,description,price,listPrice,bToBPrice,listHidden FROM products WHERE active=-1 AND removed=0 AND configOnly=0 AND serviceSpec=0 " & queryCat & " ORDER BY "& strORD &" "& strSort
Set rsInv=Server.CreateObject("ADODB.Recordset")
rsInv.CacheSize=pcv_pagesize
rsInv.PageSize=pcv_pagesize

rsInv.Open query, connTemp, adOpenStatic, adLockReadOnly

If rsInv.eof Then %>
                      
    <tr>                     
    	<td colspan="7">No products found.</td>
    </tr>

                      
<% Else 
											
	rsInv.MoveFirst
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
		
		lngIDProduct=rsInv("idproduct")
		strSKU=rsInv("sku")
		strPrdName=rsInv("description")
		dblPrdPrice=rsInv("price")
		if dblPrdPrice<>"" then
		else
		dblPrdPrice=0
		end if
		dblPrdLPrice=rsInv("listprice")
		if dblPrdLPrice<>"" then
		else
		dblPrdLPrice=0
		end if
		dblPrdWPrice=rsInv("bToBprice")
		if dblPrdWPrice<>"" then
		else
		dblPrdWPrice=0
		end if
		intPrdLHidden=rsInv("listHidden")
		if intPrdLHidden<>"" then
		else
		intPrdLHidden=0
		end if
%>
                      
		<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
            <td align="center"><div align="left"><%=strSKU%></div></td>
            <td width="80%"><a href="FindProductType.asp?idproduct=<%=lngIDProduct%>" target="_blank"><%=strPrdName%></a></td>
            <td align="center"><input type="text" name="price<%=count%>" size="7" style="text-align: right" value="<%=money(dblPrdPrice)%>"></td>
            <td align="center"><input type="text" name="lprice<%=count%>" size="7" style="text-align: right" value="<%=money(dblPrdLPrice)%>"></td>
            <td align="center"><input type="text" name="wprice<%=count%>" size="7" style="text-align: right" value="<%=money(dblPrdWPrice)%>"></td>
            <td align="center"><input type="checkbox" name="shows<%=count%>" value="-1" <%if intPrdLHidden="-1" then%>checked<%end if%> class="clearBorder"></td>
            <td align="center">
                <input type="checkbox" name="C<%=count%>" value="1" class="clearBorder">
                <input type="hidden" name="ID<%=count%>" value="<%=lngIDProduct%>">
            </td>
            </tr>
                      
<% 
	rsInv.MoveNext
Loop
set rsInv=nothing
call closeDb()
%>
						
<tr>
<td colspan="7" align="right" class="pcSmallText">
	<input type="hidden" name="count" value=<%=count%>>
	<a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a>
</td>
</tr>
<tr>
<td colspan="7" align="left">
  <input type="submit" name="submit" value="Update Prices" class="submit2">&nbsp;
  <% if validNum(pcIntCategoryID) then %>
  <input type="button" name="back" value="Edit Category" onClick="document.location.href='modcata.asp?idcategory=<%=pcIntCategoryID%>'">&nbsp;
  <input type="button" name="back" value="Back" onClick="javascript:history.back()">
  <input type="hidden" name="idcategory" value="<%=pcIntCategoryID%>">
  <% else %>
  <input type="button" name="back" value="Back" onClick="javascript:history.back()">
  <% end if %>
</td>
</tr>             
<%End If%>     
</table>             
<% If iPageCount>1 Then %>
<hr>
<table class="pcCPcontent">              
<tr> 
<td><%response.Write("Currently viewing page "& iPageCurrent & " of "& iPageCount & " <P>")%></td>
</tr>
<tr>                   
<td> 
<%' display Next / Prev buttons
if iPageCurrent > 1 then %>
<a href="updPrdPrices.asp?idcat=<%=pcIntCategoryID%>&iPageCurrent=<%=iPageCurrent-1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/prev.gif" width="10" height="10" border="0"></a> 
<% end If

For I=1 To iPageCount
	If Cint(I)=Cint(iPageCurrent) Then %>
		<b><%=I%></b> 
	<% Else %>
	<a href="updPrdPrices.asp?idcat=<%=pcIntCategoryID%>&iPageCurrent=<%=I%>&order=<%=strORD%>&sort=<%=strSort%>"><%=I%></a> 
	<% End If %>

<% Next %>
<% if CInt(iPageCurrent) < CInt(iPageCount) then %>
	<a href="updPrdPrices.asp?idcat=<%=pcIntCategoryID%>&iPageCurrent=<%=iPageCurrent+1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
<% end If %>
</td>
</tr>
</table>
<% End If %>
</form>
<script language="JavaScript">
<!--
function checkAll() {
for (var j = 1; j <= <%=count%>; j++) {
box = eval("document.checkboxform.C" + j); 
if (box.checked == false) box.checked = true;
   }
}

function uncheckAll() {
for (var j = 1; j <= <%=count%>; j++) {
box = eval("document.checkboxform.C" + j); 
if (box.checked == true) box.checked = false;
   }
}
	
function isDigit(s)
{
var test=""+s;
if(test=="+"||test=="-"||test==","||test=="."||test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
		{
		return(true) ;
		}
		return(false);
	}
	
function allDigit(s)
	{
		var test=""+s ;
		for (var k=0; k <test.length; k++)
		{
			var c=test.substring(k,k+1);
			if (isDigit(c)==false)
			{
				return (false);
			}
		}
		return (true);
	}

function Form1_Validator(theForm)
{
	for (var j = 1; j <= <%=count%>; j++) 
	{
	box = eval("document.checkboxform.C" + j); 
	if (box.checked == true)
	{
	qtt= eval("document.checkboxform.price" + j);
		if (qtt.value == "")
	  	{
	    alert("Please enter a value for this field!");
	    qtt.focus();
	    return (false);
		}
		else
		{
			if (allDigit(qtt.value) == false)
			{
		    alert("Please enter a numeric value for this Field.");
		    qtt.focus();
		    return (false);
		    }
	    }
	qtt= eval("document.checkboxform.lprice" + j);
		if (qtt.value == "")
	  	{
	    alert("Please enter a value for this field!");
	    qtt.focus();
	    return (false);
		}
		else
		{
			if (allDigit(qtt.value) == false)
			{
		    alert("Please enter a numeric value for this Field.");
		    qtt.focus();
		    return (false);
		    }
	    }
	qtt= eval("document.checkboxform.wprice" + j);
		if (qtt.value == "")
	  	{
	    alert("Please enter a value for this field!");
	    qtt.focus();
	    return (false);
		}
		else
		{
			if (allDigit(qtt.value) == false)
			{
		    alert("Please enter a numeric value for this Field.");
		    qtt.focus();
		    return (false);
		    }
	    }
	}
	}

return (true);
}
//-->
</script>   
<!--#include file="AdminFooter.asp"-->
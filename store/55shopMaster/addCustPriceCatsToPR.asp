<%
'This file is part of CustCatductCart, an ecommerce application developed and sold by NetSource Commerce. CustCatductCart, its source code, the CustCatductCart name and logo are CustCatperty of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of CustCatductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add New Categories to the Promotion" %>
<% Section="specials" %>
<%PmAdmin=3%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
Dim connTemp,mySQL,rstemp

pIDPromo=request("idcode")
if not validNum(pIDPromo) then
	pIDPromo=0
end if

pIDProduct=request("idproduct")

if pIDProduct="" or pIDProduct="0" then
	response.redirect "menu.asp"
end if

if request("action")="add" then

	Count=request("Count")
	if Count="" then
	Count="0"
	end if
	
	call opendb()
	For i=1 to Count
		if request("CustCat" & i)="1" then
		IDCustCat=request("IDCustCat" & i)
			if pIDPromo<>"0" then
				mySQL="insert into pcPPFCustPriceCats (pcPrdPro_id,idCustomerCategory) values (" & pIDPromo & "," & IDCustCat & ")"
				set rstemp=Server.CreateObject("ADODB.Recordset")
				set rstemp=connTemp.execute(mySQL)
			else
				session("admin_PRFCustPriceCATs")=session("admin_PRFCustPriceCATs") & IDCustCat & ","
			end if
		end if
	Next
	set rstemp=nothing
	call closedb()
	
	if pIDPromo=0 then
		response.redirect "AddPromotionPrd.asp?idproduct=" & pIDProduct
	else
		response.redirect "ModPromotionPrd.asp?idproduct=" & pIDProduct
	end if
end if

%>

<form name="hForm" method="post" action="addCustPriceCatsToPR.asp?action=add" class="pcForms">
	<table class="pcCPcontent">
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
	    <th colspan="2">Select one or more pricing category</th>
    </tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
<%
Count=0
call opendb()
query="SELECT idcustomerCategory, pcCC_Name, pcCC_CategoryType FROM pcCustomerCategories WHERE idcustomerCategory NOT IN (SELECT idcustomerCategory FROM pcPPFCustPriceCats WHERE pcPrdPro_id = "& pIDPromo &");"
SET rs=Server.CreateObject("ADODB.RecordSet")
SET rs=conntemp.execute(query)
if NOT rs.eof then 
	do while not rs.eof 
		Count=Count+1	
		intIdcustomerCategory=rs("idcustomerCategory")
		strpcCC_Name=rs("pcCC_Name")
		strpcCC_CategoryType = rs("pcCC_CategoryType")
		%>
		<tr>
		<td width="5%" align="right"><input type="checkbox" name="CustCat<%=Count%>" value="1" class="clearBorder">
		<input type=hidden name="IDCustCat<%=Count%>" value="<%=intIdcustomerCategory%>"></td>
		<td width="95%" align="left"><a href="AdminCustomerCategory.asp?mode=2&id=<%=intIdcustomerCategory%>" target="_blank"><%=strpcCC_Name%></a></td>
    </tr>
		<%
		rs.moveNext
	loop
end if
SET rs=nothing
Call closeDB()
if Count=0 then
%>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
    <tr>
    <td colspan="2"><div class="pcCPmessage">No pricing category available</div></td>
    </tr>
<%end if%>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
    <tr>
    <td colspan="2"><%if Count>0 then%><a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a>
    <script language="JavaScript">
    <!--
    function checkAll() {
    for (var j = 1; j <= <%=count%>; j++) {
    box = eval("document.hForm.CustCat" + j); 
    if (box.checked == false) box.checked = true;
       }
    }
    
    function uncheckAll() {
    for (var j = 1; j <= <%=count%>; j++) {
    box = eval("document.hForm.CustCat" + j); 
    if (box.checked == true) box.checked = false;
       }
    }
    
    //-->
    </script>
    <%end if%>
    </td>
    </tr>
 		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
    <tr>
    <td colspan="2">
    <input name="submit" type="submit" class="submit2" value="Add to the Promotion">
    &nbsp;
    <input name="back" type="button" onClick="javascript:history.back();" value="Back">
    <input type=hidden name="Count" value="<%=Count%>">
    <input type=hidden name="idcode" value="<%=pIDPromo%>">
	<input type=hidden name="idproduct" value="<%=pIDProduct%>">
    </td>
    </tr>
</table>
</form>

<!--#include file="AdminFooter.asp"-->
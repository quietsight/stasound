<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%pageTitle="Export Products"%>
<% Section="products" %>
<%PmAdmin=2%>
<%PmAdmin=2%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="inc_srcPrdQuery.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<%
totalrecords=0
Dim connTemp
call opendb()
Set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=connTemp.execute(query)
if rstemp.eof then%>
	<!--#include file="AdminHeader.asp"-->
    <div class="pcCPmessage">
        Your search did not return any results. <a href="javascript:history.back()">Back</a>.
    </div>
    <!--#include file="AdminFooter.asp"-->
    <%
    response.end
end if

Dim strCol, Count, HTMLResult
HTMLResult=""
CSVResult=""
Count = 0

Function GenFileName()
	dim fname
	fname="File-"
	systime=now()
	fname= fname & cstr(year(systime)) & cstr(month(systime)) & cstr(day(systime)) & "-"
	fname= fname  & cstr(hour(systime)) & cstr(minute(systime)) & cstr(second(systime))
	GenFileName=fname
End Function

IF request("src_expFormat")<>"1" THEN
HTMLResult="<html><head>"& vbcrlf
HTMLResult=HTMLResult & "<link href=""pcv4_ControlPanel.css"" rel=""stylesheet"" type=""text/css"">" & vbcrlf
HTMLResult=HTMLResult & "</head><body style='padding: 20px; background-image: none;'>" & vbcrlf
HTMLResult=HTMLResult & "<h2>"
if src_PageType="0" then
	HTMLResult=HTMLResult & "Supplier: "
	query="SELECT pcSupplier_Company,pcSupplier_FirstName,pcSupplier_LastName FROM pcSuppliers WHERE pcSupplier_ID=" & src_IDSDS
	set rs1=ConnTemp.execute(query)
	HTMLResult=HTMLResult & rs1("pcSupplier_Company") & " (" & rs1("pcSupplier_FirstName") & " " & rs1("pcSupplier_LastName") & ")"
	set rs1=nothing
else
	HTMLResult=HTMLResult & "Drop-Shipper: "
	if src_IsDropShipper="1" then
		query="SELECT pcSupplier_Company,pcSupplier_FirstName,pcSupplier_LastName FROM pcSuppliers WHERE pcSupplier_ID=" & src_IDSDS
		set rs1=ConnTemp.execute(query)
		HTMLResult=HTMLResult & rs1("pcSupplier_Company") & " (" & rs1("pcSupplier_FirstName") & " " & rs1("pcSupplier_LastName") & ")"
		set rs1=nothing
	else
		query="SELECT pcDropShipper_Company,pcDropShipper_FirstName,pcDropShipper_LastName FROM pcDropShippers WHERE pcDropShipper_ID=" & src_IDSDS
		set rs1=ConnTemp.execute(query)
		HTMLResult=HTMLResult & rs1("pcDropShipper_Company") & " (" & rs1("pcDropShipper_FirstName") & " " & rs1("pcDropShipper_LastName") & ")"
		set rs1=nothing
	end if
end if
HTMLResult=HTMLResult &"</h2>" & vbcrlf
if request("src_sdsStockAlarm")<>"1" then
	HTMLResult=HTMLResult &"<p>Filter: <b>All products that belong to the selected "
	if src_PageType="0" then
		HTMLResult=HTMLResult & "Supplier</b></p>"
	else
		HTMLResult=HTMLResult & "Drop-Shipper</b></p>"
	end if
else
	HTMLResult=HTMLResult &"<p>Filter: <b>All products whose current inventory count is lower than the Reorder Level</b></p>"
end if

HTMLResult=HTMLResult & "<table class=""pcCPcontent"">" & vbcrlf
HTMLResult=HTMLResult & "<tr>" & vbcrlf
HTMLResult=HTMLResult & "<th width=""20%"">SKU</th>" & vbcrlf
HTMLResult=HTMLResult & "<th width=""60%"">Product</th>" & vbcrlf
if (src_IDSDS<>"") and (src_IDSDS<>"0") then
	HTMLResult=HTMLResult & "<th>&nbsp;</th>" & vbcrlf
	HTMLResult=HTMLResult & "<th nowrap>Stock Level</th>" & vbcrlf
	HTMLResult=HTMLResult & "<th nowrap>Reorder Level</th>" & vbcrlf
	HTMLResult=HTMLResult & "<th>Price</th>" & vbcrlf
	HTMLResult=HTMLResult & "<th>Cost</th>" & vbcrlf
end if
HTMLResult=HTMLResult & "</tr>" & vbcrlf

do while not rsTemp.eof
				
	count=count + 1
	pidProduct=trim(rstemp("idProduct"))
	pDescription=rstemp("description")
	pactive=rstemp("active")
	psku=rstemp("sku")
	pBTO=rstemp("serviceSpec")
	pItem=rstemp("configOnly")
	
	pcv_stock=rstemp("stock")
	if pcv_stock<>"" then
	else
		pcv_stock=0
	end if
	pcv_ReorderLevel=rstemp("pcProd_ReorderLevel")
	if pcv_ReorderLevel<>"" then
	else
		pcv_ReorderLevel=0
	end if
	pcv_Price=rstemp("price")
	if pcv_Price<>"" then
	else
		pcv_Price=0
	end if
	pcv_Cost=rstemp("cost")
	if pcv_Cost<>"" then
	else
		pcv_Cost=0
	end if

	HTMLResult=HTMLResult & "<tr>" & vbcrlf
	HTMLResult=HTMLResult & "<td nowrap>" & psku & "</td>" & vbcrlf
	HTMLResult=HTMLResult & "<td nowrap>" & pdescription & "</td>" & vbcrlf
	if (src_IDSDS<>"") and (src_IDSDS<>"0") then
		HTMLResult=HTMLResult & "<td nowrap>"
		if cint(pactive)=0 then
		HTMLResult=HTMLResult & "<img src=""images/notactive.gif"" width=""32"" height=""16"">"
		else
		HTMLResult=HTMLResult & "&nbsp;"
		end if
		HTMLResult=HTMLResult & "</td>" & vbcrlf
		HTMLResult=HTMLResult & "<td nowrap>" & pcv_stock & "</td>" & vbcrlf
		HTMLResult=HTMLResult & "<td nowrap>" & pcv_ReorderLevel & "</td>" & vbcrlf
		HTMLResult=HTMLResult & "<td nowrap align=""right"">" & scCurSign & money(pcv_Price) & "</td>" & vbcrlf
		HTMLResult=HTMLResult & "<td nowrap align=""right"">"
		if cdbl(pcv_Cost)=0 then
			HTMLResult=HTMLResult & "N/A"
		else
			HTMLResult=HTMLResult & scCurSign & money(pcv_Cost)
		end if
		HTMLResult=HTMLResult & "</td>" & vbcrlf
	end if
	HTMLResult=HTMLResult & "</tr>" & vbcrlf

rsTemp.MoveNext
loop
HTMLResult=HTMLResult & "</table>" & vbcrlf
HTMLResult=HTMLResult & "</body></html>"
set rstemp=nothing

ELSE 'CSV File

CSVResult=CSVResult & chr(34) & "SKU" & chr(34) & "," & chr(34) & "Product" & chr(34) & ","
if (src_IDSDS<>"") and (src_IDSDS<>"0") then
	CSVResult=CSVResult & chr(34) & "Status" & chr(34) & "," & chr(34) & "Stock Level" & chr(34) & "," & chr(34) & "Reorder Level" & chr(34) & "," & chr(34) & "Price" & chr(34) & "," & chr(34) & "Cost" & chr(34) & ","
end if
CSVResult=CSVResult & vbcrlf

do while not rsTemp.eof
				
	count=count + 1
	pidProduct=trim(rstemp("idProduct"))
	pDescription=rstemp("description")
	pactive=rstemp("active")
	psku=rstemp("sku")
	pBTO=rstemp("serviceSpec")
	pItem=rstemp("configOnly")
	
	pcv_stock=rstemp("stock")
	if pcv_stock<>"" then
	else
		pcv_stock=0
	end if
	pcv_ReorderLevel=rstemp("pcProd_ReorderLevel")
	if pcv_ReorderLevel<>"" then
	else
		pcv_ReorderLevel=0
	end if
	pcv_Price=rstemp("price")
	if pcv_Price<>"" then
	else
		pcv_Price=0
	end if
	pcv_Cost=rstemp("cost")
	if pcv_Cost<>"" then
	else
		pcv_Cost=0
	end if

	CSVResult=CSVResult & chr(34) & psku & chr(34) & "," & chr(34) & pdescription & chr(34) & ","
	if (src_IDSDS<>"") and (src_IDSDS<>"0") then
		if cint(pactive)=0 then
			CSVResult=CSVResult & chr(34) & "Inactive" & chr(34) & ","
		else
			CSVResult=CSVResult & chr(34) & "Active" & chr(34) & ","
		end if
		CSVResult=CSVResult & chr(34) & pcv_stock & chr(34) & "," & chr(34) & pcv_ReorderLevel & chr(34) & "," & chr(34) & scCurSign & money(pcv_Price) & chr(34) & ","
		if cdbl(pcv_Cost)=0 then
			CSVResult=CSVResult & chr(34) & "N/A" & chr(34) & ","
		else
			CSVResult=CSVResult & chr(34) & scCurSign & money(pcv_Cost) & chr(34) & ","
		end if
	end if
	CSVResult=CSVResult & vbcrlf

rsTemp.MoveNext
loop
set rstemp=nothing

END IF
call closedb()
IF request("src_expFormat")<>"1" THEN%>
<%=HTMLResult%>
<%ELSE
	strFile=GenFileName()
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	Set a=fs.CreateTextFile(server.MapPath(".") & "\" & strFile & ".csv",True)
	a.Write(CSVResult)
	a.Close
	Set a=Nothing
	Set fs=Nothing
	response.redirect "getFile.asp?frompage=1&file="& strFile &"&Type=csv"
END IF%>
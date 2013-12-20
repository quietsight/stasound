<%PmAdmin=0%><!--#include file="adminv.asp"--><%Response.ContentType = "text/xml"%><?xml version="1.0" ?>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="inc_srcSDSQuery.asp"-->
<%
totalrecords=0
Dim connTemp
call opendb()
Set rstemp=Server.CreateObject("ADODB.Recordset")
rstemp.CacheSize=iPageSize
rstemp.PageSize=iPageSize
rstemp.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText
rstemp.AbsolutePage=iPageCurrent

Dim strCol, Count, HTMLResult
HTMLResult=""
Count = 0
strCol = "#E1E1E1"

HTMLResult=HTMLResult & "<form name=""srcresult"" class=""pcForms"">" & vbcrlf
HTMLResult=HTMLResult & "<table class=""pcCPcontent"">" & vbcrlf
HTMLResult=HTMLResult & "<tr>" & vbcrlf
HTMLResult=HTMLResult & "<th>&nbsp;</th>" & vbcrlf
HTMLResult=HTMLResult & "<th nowrap>Company</th>" & vbcrlf
HTMLResult=HTMLResult & "<th nowrap>Name</th>" & vbcrlf
HTMLResult=HTMLResult & "<th nowrap>Phone</th>" & vbcrlf
HTMLResult=HTMLResult & "<th nowrap>Email</th>" & vbcrlf
if src_PageType="0" then ' The page is showing suppliers
	HTMLResult=HTMLResult & "<th nowrap>Drop-Shipping</th>" & vbcrlf
end if
HTMLResult=HTMLResult & "<th align=""right"">&nbsp;</th>" & vbcrlf
HTMLResult=HTMLResult & "</tr>" & vbcrlf
HTMLResult=HTMLResult & "<tr><td colspan=""7"" class=""pcCPspacer""></td></tr>" & vbcrlf

src_DisplayType=getUserInput(request("src_DisplayType"),0)
src_ShowLinks=getUserInput(request("src_ShowLinks"),0)
if src_PageType="0" then
	tmp_Title="Supplier"
else
	tmp_Title="Drop-Shipper"
end if

do while (not rsTemp.eof) and (count < rsTemp.pageSize)

	count=count + 1
	pidsds=rstemp(pcv_Table & "_ID")
	pname=rstemp(pcv_Table & "_FirstName")
	pLastName=rstemp(pcv_Table & "_LastName")
	pCompany=rstemp(pcv_Table & "_Company")
	pemail=rstemp(pcv_Table & "_Email")
	pphone=rstemp(pcv_Table & "_Phone")
	pIsDropShipper=rstemp("IsDropShipper")
	if isNull(pIsDropShipper) or pIsDropShipper="" or pIsDropShipper="0" then
		pIsDropShipper=0
	end if
	
	if src_PageType="0" then ' The page is showing suppliers
		tmp_pagetype=0
	else
		tmp_pagetype=1 ' The page is showing drop-shippers -> set variable to 1
	end if
	
	' Debugging
	' HTMLResult=HTMLResult & "<tr colspan=""5"">pIsDropShipper=" & pIsDropShipper & " and " & "tmp_pagetype=" & tmp_pagetype & "</td></tr>"
	
	
	HTMLResult=HTMLResult & "<tr onMouseOver=""this.className='activeRow'"" onMouseOut=""this.className='cpItemlist'"" class=""cpItemlist"">" & vbcrlf
	HTMLResult=HTMLResult & "<td>" & vbcrlf

	if src_PageType="1" then ' The page is showing drop-shippers
		pidsds1=pidsds & "_" & pIsDropShipper
	else
		pidsds1=pidsds
	end if
	if src_DisplayType="1" then
		HTMLResult=HTMLResult & "<input type=checkbox name=""C" & count & """ value=""" & pidsds1 & """ onclick=""javascript:updvalue(this);"">" & "&nbsp;"
	else
		if src_DisplayType="2" then
			HTMLResult=HTMLResult & "<input type=radio name=""R1"" value=""" & pidsds1 & """ onclick=""javascript:updvalue(this);"">" & "&nbsp;"
		else
			HTMLResult=HTMLResult & pidsds
		end if
	end if
	
	HTMLResult=HTMLResult & "</td>" & vbcrlf
	HTMLResult=HTMLResult & "<td><a href=""sds_modify.asp?idsds=" & pidsds & "&pagetype=" & tmp_pagetype & "&clone=0"">" & pCompany & "</a></td>" & vbcrlf
	HTMLResult=HTMLResult & "<td>" & pLastName&", "&pname
	HTMLResult=HTMLResult & "</td>" & vbcrlf
	HTMLResult=HTMLResult & "<td nowrap>" & pphone & "</td>" & vbcrlf
	HTMLResult=HTMLResult & "<td><a target=""_blank"" href=""mailto:" & pemail & """>" & pemail &"</a></td>" & vbcrlf
	if src_PageType="0" then ' The page is showing suppliers
		HTMLResult=HTMLResult & "<td>" & vbcrlf
		if pIsDropShipper="1" then ' The supplier is also a drop-shipper -> show it
			pcv_DSType="Yes"
		else
			pcv_DSType="&nbsp;"
		end if
		HTMLResult=HTMLResult & pcv_DSType & "</td>" & vbcrlf
	end if
	HTMLResult=HTMLResult & "<td align=""right""  class=""cpLinksList"">" & vbcrlf
	if src_ShowLinks="1" then
		HTMLResult=HTMLResult & "<a href=""sds_modify.asp?idsds=" & pidsds & "&pagetype=" & tmp_pagetype & "&clone=0"">Edit</a> | <a href=""sds_modify.asp?idsds=" & pidsds & "&pagetype=" & tmp_pagetype & "&clone=1"">Clone</a> | <a href=""javascript:getPrdReport(" & pidsds & "," & pIsDropShipper & ");"">Products</a>"
		
		query="SELECT products.idproduct FROM products,pcDropShippersSuppliers WHERE products.removed=0 AND products." & pcv_Table & "_ID=" & pidsds & " AND pcDropShippersSuppliers.idproduct=products.idproduct AND pcDropShippersSuppliers.pcDS_IsDropShipper=" & pIsDropShipper
		set rs1=connTemp.execute(query)
			
		if rs1.eof then
			if src_PageType<>"0" then ' This is a drop-shipper
				HTMLResult=HTMLResult & " | <a href=""javascript:if (confirm('You are about to remove this " & tmp_Title & " from your database. Are you sure you want to complete this action?')) location='sds_delete.asp?action=del&idsds=" & pidsds & "&pagetype="  & tmp_pagetype & "';"">Remove</a>" & vbcrlf
			else ' This is a supplier
				query="SELECT products.idproduct FROM products WHERE pcSupplier_ID=" & pidsds & " AND removed=0;"
				set rs1=connTemp.execute(query)
				if rs1.eof then
					HTMLResult=HTMLResult & " | <a href=""javascript:if (confirm('You are about to remove this " & tmp_Title & " from your database. Are you sure you want to complete this action?')) location='sds_delete.asp?action=del&idsds=" & pidsds & "&pagetype="  & tmp_pagetype & "';"">Remove</a>" & vbcrlf
				end if
				set rs1=nothing
			end if
		end if
		set rs1=nothing
	else
		HTMLResult=HTMLResult & "&nbsp;"
	end if
	HTMLResult=HTMLResult & "</td>" & vbcrlf
	HTMLResult=HTMLResult & "</tr>" & vbcrlf
rsTemp.MoveNext
loop

HTMLResult=HTMLResult & "</table>" & vbcrlf
HTMLResult=HTMLResult & "<input type=hidden name=count value=""" & count & """>" & vbcrlf
HTMLResult=HTMLResult & "</form>" & vbcrlf

set rstemp=nothing
call closedb()

'*** Fixed FireFox issues
Dim tmpData,tmpData1
Dim tmp1,tmp2,i,Count1
tmpData=Server.HTMLEncode(HTMLResult)
Count1=0
tmpData1=""
tmp1=split(tmpData,"&lt;/tr&gt;")
For i=lbound(tmp1) to ubound(tmp1)
	if i>lbound(tmp1) then
		tmp2="&lt;/tr&gt;" & tmp1(i)
	else
		tmp2=tmp1(i)
	end if
	Count1=Count1+1
	tmpData1=tmpData1 & "<data" & Count1 & ">" & tmp2 & "</data" & Count1 & ">" & vbcrlf
Next
%><note>
<data0><%=Count1%></data0>
<%=tmpData1%>
</note>
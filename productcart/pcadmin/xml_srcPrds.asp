<%PmAdmin=0%><!--#include file="adminv.asp"--><%Response.ContentType = "text/xml"%><?xml version="1.0" ?>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="inc_srcPrdQuery.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<%totalrecords=0
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
HTMLResult=HTMLResult & "<th width=""10%"">SKU</th>" & vbcrlf
HTMLResult=HTMLResult & "<th width=""60%"">Product</th>" & vbcrlf
if session("srcprd_DiscArea")="1" then
	HTMLResult=HTMLResult & "<th>&nbsp;</th>" & vbcrlf
	HTMLResult=HTMLResult & "<th>&nbsp;</th>" & vbcrlf
	HTMLResult=HTMLResult & "<th>&nbsp;</th>" & vbcrlf
end if
if (src_IDSDS<>"") and (src_IDSDS<>"0") then
	HTMLResult=HTMLResult & "<th>&nbsp;</th>" & vbcrlf
	HTMLResult=HTMLResult & "<th nowrap>Stock Level</th>" & vbcrlf
	HTMLResult=HTMLResult & "<th nowrap>Reorder Level</th>" & vbcrlf
	HTMLResult=HTMLResult & "<th>Price</th>" & vbcrlf
	HTMLResult=HTMLResult & "<th>Cost</th>" & vbcrlf
end if
HTMLResult=HTMLResult & "<th width=""20%"" colspan=""2"">&nbsp;</th>" & vbcrlf
HTMLResult=HTMLResult & "</tr><tr><td colspan='9' class='pcCPSpacer'></td></tr>" & vbcrlf

src_DisplayType=getUserInput(request("src_DisplayType"),0)
src_ShowLinks=getUserInput(request("src_ShowLinks"),0)

do while (not rsTemp.eof) and (count < rsTemp.pageSize)
				
	If strCol <> "#FFFFFF" Then
		strCol = "#FFFFFF"
	Else 
		strCol = "#E1E1E1"
	End If
	count=count + 1
	pidProduct=trim(rstemp("idProduct"))
	pDescription=rstemp("description")
	pactive=rstemp("active")
	pSmallImageUrl=rstemp("smallImageUrl")
	psku=rstemp("sku")
	pBTO=rstemp("serviceSpec")
	pItem=rstemp("configOnly")
	
	'Start SDBA
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
	'End SDBA

	'Start SDBA
	if (src_IDSDS<>"") and (src_IDSDS<>"0") and (clng(pcv_stock)<clng(pcv_ReorderLevel)) and (src_sdsStockAlarm<>"1") then
		HTMLResult=HTMLResult & "<tr bgcolor=""#FFFF99"">" & vbcrlf
	else
		HTMLResult=HTMLResult & "<tr onMouseOver=""this.className='activeRow'"" onMouseOut=""this.className='cpItemlist'"" class=""cpItemlist"">" & vbcrlf
	end if
	'End SDBA
	
	if src_DisplayType="1" then
		HTMLResult=HTMLResult & "<td><input type=checkbox name=""C" & count & """ value=""" & pidProduct & """ onclick=""javascript:updvalue(this);"" class=""clearBorder""></td>" & vbcrlf
	else
		if src_DisplayType="2" then
			HTMLResult=HTMLResult & "<td><input type=radio name=""R1"" value=""" & pidProduct & """ onclick=""javascript:updvalue(this);"" class=""clearBorder""></td>" & vbcrlf
		else
			HTMLResult=HTMLResult & "<td>&nbsp;</td>" & vbcrlf
		end if
	end if
	HTMLResult=HTMLResult & "<td>" & psku & "</td>" & vbcrlf
	HTMLResult=HTMLResult & "<td><a href='FindProductType.asp?id=" & pidProduct & "'>"
	if pSmallImageUrl <> "" then
		HTMLResult=HTMLResult & "<img src='../pc/catalog/" & pSmallImageUrl & "' align='absbottom' class='pcShowProductImageM'>"
	end if
	HTMLResult=HTMLResult & pdescription & "</a></td>" & vbcrlf
	
	if session("srcprd_DiscArea")="1" AND (session("cp_lct_src_PromoType")="") then
		pcv_HaveQtyDisc=0
		query= "SELECT idproduct FROM discountsPerQuantity WHERE discountDesc='PD' AND idproduct="&pidProduct
		set rsHQ=server.createobject("adodb.recordset")
		set rsHQ=conntemp.execute(query)
		if NOT rsHQ.eof then
			pcv_HaveQtyDisc=1
		end if
		set rsHQ=nothing
		if pcv_HaveQtyDisc=1 then
			HTMLResult=HTMLResult & "<td nowrap>Discounts Applied</td>" & vbcrlf
			HTMLResult=HTMLResult & "<td><div align=""center""><a href=""ModDctQtyPrd.asp?idproduct=" & pidProduct & """><img src=""images/pcIconGo.jpg"" border=""0""></a></div></td>" & vbcrlf
			HTMLResult=HTMLResult & "<td>&nbsp;</td>" & vbcrlf
			
		else
			HTMLResult=HTMLResult & "<td>&nbsp;</td>" & vbcrlf
			HTMLResult=HTMLResult & "<td>&nbsp;</td>" & vbcrlf
			HTMLResult=HTMLResult & "<td><div align=""center""><a href=""AdminDctQtyPrd.asp?idproduct=" & pidProduct & "&mode=f""><img src=""images/pcIconPlus.jpg"" border=""0""></a></div></td>" & vbcrlf
		end if
	end if
	
	if session("srcprd_DiscArea")="1" AND (session("cp_lct_src_PromoType")<>"") then
		pcv_HaveQtyDisc=0
		query= "SELECT idproduct, pcPrdPro_PromoMsg FROM pcPrdPromotions WHERE idproduct="&pidProduct
		set rsHQ=server.createobject("adodb.recordset")
		set rsHQ=conntemp.execute(query)
		if NOT rsHQ.eof then
			pcv_HaveQtyDisc=1
			pcv_PromoDesc=rsHQ("pcPrdPro_PromoMsg")
		end if
		set rsHQ=nothing
		if pcv_HaveQtyDisc=1 then
			HTMLResult=HTMLResult & "<td align=""right"" colspan=""2"" nowrap>Promotion Applied<br><span class='pcSmallText'>" & pcv_PromoDesc & "</span></td>" & vbcrlf
			HTMLResult=HTMLResult & "<td><div align=""center""><a href=""ModPromotionPrd.asp?idproduct=" & pidProduct & "&iMode=start" & """><img src=""images/pcIconGo.jpg"" border=""0""></a></div></td>" & vbcrlf
			
		else
			HTMLResult=HTMLResult & "<td>&nbsp;</td>" & vbcrlf
			HTMLResult=HTMLResult & "<td>&nbsp;</td>" & vbcrlf
			HTMLResult=HTMLResult & "<td><div align=""center""><a href=""AddPromotionPrd.asp?idproduct=" & pidProduct & """><img src=""images/pcIconPlus.jpg"" border=""0""></a></div></td>" & vbcrlf
		end if
	end if
	
	'Start SDBA
	if (src_IDSDS<>"") and (src_IDSDS<>"0") then
		HTMLResult=HTMLResult & "<td>"
		if cint(pactive)=0 then
		HTMLResult=HTMLResult & "<img src=""images/notactive.gif"" width=""32"" height=""16"">"
		else
		HTMLResult=HTMLResult & "&nbsp;"
		end if
		HTMLResult=HTMLResult & "</td>" & vbcrlf
		HTMLResult=HTMLResult & "<td>" & pcv_stock & "</td>" & vbcrlf
		HTMLResult=HTMLResult & "<td>" & pcv_ReorderLevel & "</td>" & vbcrlf
		HTMLResult=HTMLResult & "<td align=""right"">" & scCurSign & money(pcv_Price) & "</td>" & vbcrlf
		HTMLResult=HTMLResult & "<td align=""right"">"
		if cdbl(pcv_Cost)=0 then
			HTMLResult=HTMLResult & "N/A"
		else
			HTMLResult=HTMLResult & scCurSign & money(pcv_Cost)
		end if
		HTMLResult=HTMLResult & "</td>" & vbcrlf
	end if
	'End SDBA
	
	if src_ShowLinks="1" then
	
	if (cint(pBTO)=0) and (cint(pItem)=0) then
		HTMLResult=HTMLResult & "<td align=""right"" class=""cpLinksList"" nowrap>" & vbcrlf
		HTMLResult=HTMLResult & "<a href=""FindProductType.asp?id=" & pidproduct & """>Details</a> | <a href=""modPrdOpta.asp?idproduct=" & pidproduct & """>Options</a> | <a href=""AdminCustom.asp?idproduct=" & pidproduct &  """>Custom Fields</a> | <a href=""FindDupProductType.asp?idproduct=" & pidproduct & """>Clone</a> | <a href=""javascript:if (confirm('You are about to permanently delete this product from your database. Are you sure you want to complete this action?')) location='delPrdb.asp?idproduct=" & pidproduct & "&redir=srcPrds.asp'"">Delete</a></td>" & vbcrlf
		HTMLResult=HTMLResult & "<td width=""1%"">" & vbcrlf
		if cint(pactive)=0 then
		HTMLResult=HTMLResult & "<img src=""images/notactive.gif"" width=""32"" height=""16"">"
		else
		HTMLResult=HTMLResult & "&nbsp;"
		end if
		HTMLResult=HTMLResult & "</td>" & vbcrlf
	end if
	
	if (cint(pBTO)<>0) then
		HTMLResult=HTMLResult & "<td align=""right"" class=""cpLinksList"" nowrap>" & vbcrlf
		HTMLResult=HTMLResult & "<a href=""FindProductType.asp?id=" & pidproduct & """>Details</a> | <a href=""modBTOconfiga.asp?idProduct=" & pidProduct & """>Configure</a> | <a href=""AdminCustom.asp?idproduct=" & pidproduct & """>Custom Fields</a> | <a href=""FindDupProductType.asp?idproduct=" & pidproduct & """>Clone</a> | <a href=""javascript:if (confirm('You are about to permanently delete this product from your database. Are you sure you want to complete this action?')) location = 'delPrdb.asp?redirect=BTO&idproduct=" & pidproduct & "&redir=srcPrds.asp'"">Delete</a></td>" & vbcrlf
		HTMLResult=HTMLResult & "<td width=""1%"">" & vbcrlf
		if cint(pactive)=0 then
		HTMLResult=HTMLResult & "<img src=""images/notactive.gif"" width=""32"" height=""16"">"
		else
		HTMLResult=HTMLResult & "&nbsp;"
		end if
		HTMLResult=HTMLResult & "</td>" & vbcrlf
	end if
	
	if (cint(pItem)<>0) then
		HTMLResult=HTMLResult & "<td align=""right"" class=""cpLinksList"" nowrap>" & vbcrlf
		HTMLResult=HTMLResult & "<a  href=""FindProductType.asp?id=" & pidproduct & """>Edit</a> | <a  href=""FindDupProductType.asp?idproduct=" & pidproduct & """>Clone</a> | <a  href=""javascript:if (confirm('You are about to permanently delete this product from your database. Are you sure you want to complete this action?')) location = 'delPrdb.asp?idproduct=" & pidproduct & "&redir=srcPrds.asp'"">Delete</a></td>" & vbcrlf
		HTMLResult=HTMLResult & "<td width=""1%"">" & vbcrlf
		if cint(pactive)=0 then
		HTMLResult=HTMLResult & "<img src=""images/notactive.gif"" width=""32"" height=""16"">"
		else
		HTMLResult=HTMLResult & "&nbsp;"
		end if
		HTMLResult=HTMLResult & "</td>" & vbcrlf
	end if
	
	else
		HTMLResult=HTMLResult & "<td>&nbsp;</td><td>&nbsp;</td>"
	end if
	
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
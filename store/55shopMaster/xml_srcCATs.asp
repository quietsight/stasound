<%PmAdmin=0%><!--#include file="adminv.asp"--><%Response.ContentType = "text/xml"%><?xml version="1.0" ?>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="inc_srcCATQuery.asp"-->
<%totalrecords=0
Dim connTemp
call opendb()
Set rstemp=Server.CreateObject("ADODB.Recordset")

rstemp.CacheSize=iPageSize
rstemp.PageSize=iPageSize
rstemp.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

rstemp.AbsolutePage=iPageCurrent


'--- Get Parent Categories ---

Dim tmpParent
Dim intCount
Dim pcArrP
tmpParent=""

query="SELECT categories.idcategory,categories.categorydesc,categories.idParentCategory FROM categories ORDER BY categories.idcategory asc;"
set rstemp1=conntemp.execute(query)

if not rstemp1.eof then
	pcArrP=rstemp1.getRows()
	intCount=ubound(pcArrP,2)
end if

set rstemp1=nothing

Function FindParent(idCat)
Dim k
	if clng(idCat)<>1 then
	For k=0 to intCount
		if (clng(pcArrP(0,k))=clng(idCat)) and (clng(pcArrP(0,k))<>1)	then
			if tmpParent<>"" then
			tmpParent="/" & tmpParent
			end if
			tmpParent=pcArrP(1,k) & tmpParent
			FindParent(pcArrP(2,k))
			exit for
		end if
	Next
	end if
End function

'--- End of Parent Categories ---

Dim strCol, Count, HTMLResult
HTMLResult=""
Count = 0
strCol = "#E1E1E1"

HTMLResult=HTMLResult & "<form name=""srcresult"" class=""pcForms"">" & vbcrlf
HTMLResult=HTMLResult & "<table class=""pcCPcontent"">" & vbcrlf
HTMLResult=HTMLResult & "<tr>" & vbcrlf
HTMLResult=HTMLResult & "<th>&nbsp;</th>" & vbcrlf
HTMLResult=HTMLResult & "<th nowrap>Category Name</th>" & vbcrlf
HTMLResult=HTMLResult & "<th nowrap>Parent</th>" & vbcrlf
if session("srcCat_DiscArea")="1" then
	HTMLResult=HTMLResult & "<th>&nbsp;</th>" & vbcrlf
	HTMLResult=HTMLResult & "<th>&nbsp;</th>" & vbcrlf
	HTMLResult=HTMLResult & "<th>&nbsp;</th>" & vbcrlf
end if
HTMLResult=HTMLResult & "<th align=""right"">&nbsp;</th>" & vbcrlf
HTMLResult=HTMLResult & "</tr><tr><td colspan='6' class='pcCPSpacer'></td></tr>" & vbcrlf

src_DisplayType=getUserInput(request("src_DisplayType"),0)
src_ShowLinks=getUserInput(request("src_ShowLinks"),0)

do while (not rsTemp.eof) and (count < rsTemp.pageSize)
				
	If strCol <> "#FFFFFF" Then
		strCol = "#FFFFFF"
	Else 
		strCol = "#E1E1E1"
	End If
	count=count + 1
	pcv_idcategory=rstemp("idcategory")
	pcv_catname=rstemp("CategoryDesc")
	pcv_idparent=rstemp("idParentCategory")
	pcv_cathidden=rstemp("iBTOhide")
	
	tmpParent=""
	if pcv_idparent="1" then
	else 
		FindParent(pcv_idparent)
		if tmpParent<>"" then
		tmpParent=" [" & tmpParent & "]"
		end if
	end if
	
	if tmpParent="" then
	tmpParent="[Top Level Category]"
	end if
	
	strSQL="SELECT idCategory,categoryDesc FROM categories WHERE idParentCategory="& pcv_idcategory
	dim rs2
	set rs2=server.CreateObject("ADODB.Recordset")
	rs2.Open strSQL, conntemp, adOpenStatic, adLockReadOnly
	subVar="0"
	if rs2.eof then
		subVar="1"
	End If
	Set rs2=nothing
	
	mySQL="SELECT idParentCategory FROM categories WHERE idcategory=" & pcv_idparent
   	set rstemp1=connTemp.execute(mySQL)
   	if not rstemp1.eof then
   	pcv_top=rstemp1("idParentCategory")
   	else
   	pcv_top=""
   	end if
   	set rstemp1=nothing
   	

	HTMLResult=HTMLResult & "<tr onMouseOver=""this.className='activeRow'"" onMouseOut=""this.className='cpItemlist'"" class=""cpItemlist"">" & vbcrlf
	HTMLResult=HTMLResult & "<td>" & vbcrlf

	if src_DisplayType="1" then
		HTMLResult=HTMLResult & "<input type=checkbox name=""C" & count & """ value=""" & pcv_idcategory & """ onclick=""javascript:updvalue(this);"" class=""clearBorder"">" & "&nbsp;"
	else
		if src_DisplayType="2" then
			HTMLResult=HTMLResult & "<input type=radio name=""R1"" value=""" & pcv_idcategory & """ onclick=""javascript:updvalue(this);"" class=""clearBorder"">" & "&nbsp;"
		else
			HTMLResult=HTMLResult & "&nbsp;"
		end if
	end if

	HTMLResult=HTMLResult & "</td>" & vbcrlf
	HTMLResult=HTMLResult & "<td>" & pcv_catname & "</td>" & vbcrlf
	HTMLResult=HTMLResult & "<td nowrap>" & tmpParent & "</td>" & vbcrlf
	
	if session("srcCat_DiscArea")="1" and request("CatDiscType")<>"" then
		pcv_HaveQtyDisc=0
		query="SELECT pcCD_idDiscount FROM pcCatDiscounts WHERE pcCD_idcategory="&pcv_idcategory
		set rsHQ=server.createobject("adodb.recordset")
		set rsHQ=conntemp.execute(query)
		if NOT rsHQ.eof then
			pcv_HaveQtyDisc=1
		end if
		set rsHQ=nothing
		if pcv_HaveQtyDisc=1 then
			HTMLResult=HTMLResult & "<td nowrap>Discounts Applied</td>" & vbcrlf
			HTMLResult=HTMLResult & "<td><div align=""center""><a href=""ModDctQtyCat.asp?idcategory=" & pcv_idcategory & """><img src=""images/pcIconGo.jpg"" border=""0""></a></div></td>" & vbcrlf
			HTMLResult=HTMLResult & "<td>&nbsp;</td>" & vbcrlf
			
		else
			HTMLResult=HTMLResult & "<td>&nbsp;</td>" & vbcrlf
			HTMLResult=HTMLResult & "<td>&nbsp;</td>" & vbcrlf
			HTMLResult=HTMLResult & "<td><div align=""center""><a href=""AdminDctQtyCat.asp?idcategory=" & pcv_idcategory & "&mode=f""><img src=""images/pcIconPlus.jpg"" border=""0""></a></div></td>" & vbcrlf
		end if
	end if
	
	if session("srcCat_DiscArea")="1" and request("CatPromoType")<>"" then
		pcv_HaveQtyDisc=0
		query="SELECT pcCatPro_id, pcCatPro_PromoMsg FROM pcCatPromotions WHERE idcategory="&pcv_idcategory
		set rsHQ=server.createobject("adodb.recordset")
		set rsHQ=conntemp.execute(query)
		if NOT rsHQ.eof then
			pcv_HaveQtyDisc=1
			pcv_PromoDesc=rsHQ("pcCatPro_PromoMsg")
		end if
		set rsHQ=nothing
		if pcv_HaveQtyDisc=1 then
			HTMLResult=HTMLResult & "<td align=""right"" colspan=""2"" nowrap>Promotion Applied<br><span class='pcSmallText'>" & pcv_PromoDesc & "</span></td>" & vbcrlf
			HTMLResult=HTMLResult & "<td><div align=""center""><a href=""ModPromotionCat.asp?idcategory=" & pcv_idcategory & "&iMode=start" & """><img src=""images/pcIconGo.jpg"" border=""0""></a></div></td>" & vbcrlf
			
		else
			HTMLResult=HTMLResult & "<td>&nbsp;</td>" & vbcrlf
			HTMLResult=HTMLResult & "<td>&nbsp;</td>" & vbcrlf
			HTMLResult=HTMLResult & "<td><div align=""center""><a href=""AddPromotionCat.asp?idcategory=" & pcv_idcategory & """><img src=""images/pcIconPlus.jpg"" border=""0""></a></div></td>" & vbcrlf
		end if
	end if
	
	HTMLResult=HTMLResult & "<td align=""right"" nowrap class=""cpItemlist"">" & vbcrlf
	if src_ShowLinks="1" then
		pcIntHidden=0
		if (pcv_cathidden<>"") and (pcv_cathidden="1") then
			pcIntHidden=1
			HTMLResult=HTMLResult & "<img src=""images/hidden.gif"" width=""33"" height=""9"" border=0 align=""baseline"">&nbsp;"
		end if

		If subVar<>"1" then
			HTMLResult=HTMLResult & "<a target=""_blank"" href=""viewCat.asp?hidden=" & pcIntHidden & "&parent=" & pcv_idcategory & """>View Subcategories</a> | "
		Else
			HTMLResult=HTMLResult & "<a target=""_blank"" href=""editCategories.asp?lid=" & pcv_idcategory & """>View/Add Products</a> | "
		End If
		HTMLResult=HTMLResult & "<a target=""_blank"" href=""modCata.asp?idcategory=" & pcv_idcategory & "&top=" & pcv_top & "&parent=" & pcv_idparent & """>Edit</a>"
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
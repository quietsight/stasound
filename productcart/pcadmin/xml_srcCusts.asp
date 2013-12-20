<%PmAdmin=0%><!--#include file="adminv.asp"--><%Response.ContentType = "text/xml"%><?xml version="1.0" ?>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="inc_srcCustQuery.asp"-->
<%totalrecords=0
Dim connTemp
call opendb()
Set rstemp=Server.CreateObject("ADODB.Recordset")

rstemp.CacheSize=iPageSize
rstemp.PageSize=iPageSize
rstemp.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

rstemp.AbsolutePage=iPageCurrent

Dim Count, HTMLResult
HTMLResult=""
Count = 0

HTMLResult=HTMLResult & "<form name=""srcresult"" class=""pcForms"">" & vbcrlf
HTMLResult=HTMLResult & "<table class=""pcCPcontent"">" & vbcrlf
HTMLResult=HTMLResult & "<tr>" & vbcrlf
HTMLResult=HTMLResult & "<th width=""5%"">&nbsp;</th>" & vbcrlf
HTMLResult=HTMLResult & "<th width=""30%"">Name</th>" & vbcrlf
HTMLResult=HTMLResult & "<th width=""15%"">Phone</th>" & vbcrlf
HTMLResult=HTMLResult & "<th width=""30%"">Company</th>" & vbcrlf
HTMLResult=HTMLResult & "<th width=""15%"">Type</th>" & vbcrlf
HTMLResult=HTMLResult & "<th width=""5%"" align=""right"">&nbsp;</th>" & vbcrlf
HTMLResult=HTMLResult & "</tr><tr><td colspan='6' class='pcCPSpacer'></td></tr>" & vbcrlf

src_DisplayType=getUserInput(request("src_DisplayType"),0)
src_ShowLinks=getUserInput(request("src_ShowLinks"),0)

do while (not rsTemp.eof) and (count < rsTemp.pageSize)
				
	count=count + 1
	pLastName=rstemp("LastName")
	pname=rstemp("name")
	pcustomerCompany=rstemp("customerCompany")
	pphone=rstemp("phone")
	pcustomerType=rstemp("customerType")
	pidcustomer=rstemp("idcustomer")
	pemail=rstemp("email")

	HTMLResult=HTMLResult & "<tr onMouseOver=""this.className='activeRow'"" onMouseOut=""this.className='cpItemlist'"" class=""cpItemlist"">" & vbcrlf
	HTMLResult=HTMLResult & "<td>" & vbcrlf

	if src_DisplayType="1" then
		HTMLResult=HTMLResult & "<input type=checkbox name=""C" & count & """ value=""" & pidcustomer & """ onclick=""javascript:updvalue(this);"">" & "&nbsp;"
	else
		if src_DisplayType="2" then
			HTMLResult=HTMLResult & "<input type=radio name=""R1"" value=""" & pidcustomer & """ onclick=""javascript:updvalue(this);"">" & "&nbsp;"
		else
			HTMLResult=HTMLResult & "&nbsp;"
		end if
	end if
	
	HTMLResult=HTMLResult & "</td>" & vbcrlf
	HTMLResult=HTMLResult & "<td><a target=""_blank"" href=""modCusta.asp?idcustomer=" & pidcustomer & """>" & pLastName&", "&pname
	if pcustomerType="3" then
		HTMLResult=HTMLResult & "<img src=""images/pcadmin_lockedaccount.jpg"">"
	end if
	HTMLResult=HTMLResult & "</a></td>" & vbcrlf
	HTMLResult=HTMLResult & "<td nowrap>" & pphone & "</td>" & vbcrlf
	HTMLResult=HTMLResult & "<td><a target=""_blank"" href=""modCusta.asp?idcustomer=" & pidcustomer & """>" & pcustomerCompany & "</a></td>" & vbcrlf
	HTMLResult=HTMLResult & "<td>" & vbcrlf
	if pcustomerType="1" then
		custType="Wholesale"
	else
		custType="Retail"
	end if
	HTMLResult=HTMLResult & custType & "</td>" & vbcrlf
	HTMLResult=HTMLResult & "<td align=""right"" nowrap class=""cpLinksList"">" & vbcrlf
	if src_ShowLinks="1" then
		HTMLResult=HTMLResult & "<a target=""_blank"" href=""mailto:" & pemail & """>E-Mail</a> - <a target=""_blank"" href=""modCusta.asp?idcustomer=" & pidcustomer & """>Edit</a> - <a target=""_blank"" href=""viewCustOrders.asp?idcustomer=" & pidcustomer & """>View Orders</a>"
		if pcf_GetCustType(pidcustomer)=0 then
			HTMLResult=HTMLResult & " - <a href=""adminPlaceOrder.asp?idcustomer=" & pidcustomer & " target=""_blank"">Place Order</a>" & vbcrlf
		end if
	else
		HTMLResult=HTMLResult & "&nbsp;"
	end if
	HTMLResult=HTMLResult & "</td>" & vbcrlf
	HTMLResult=HTMLResult & "</tr>" & vbcrlf

rsTemp.MoveNext
loop
HTMLResult=HTMLResult & "</table>" & vbcrlf
HTMLResult=HTMLResult & "<input type=""hidden"" name=""count"" value=""" & count & """>" & vbcrlf
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
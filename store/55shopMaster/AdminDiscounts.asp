<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
tmpQuery=""
if request("list")="archive" then
	tmpQuery="active=2"
	pageTitle="Discounts by Code - Archived"
else
	tmpQuery="active=-1 OR active=0"
	pageTitle="Discounts by Code - Summary"
end if %>
<% Section="specials" %>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp"-->

<!--#include file="AdminHeader.asp"-->

<% Dim rs, connTemp, query, pcvDiscounts, i

call openDb()


Dim iStart, iOffset
iStart = Request("Start")
iOffset = Request("Offset")

if Not IsNumeric(iStart) or Len(iStart) = 0 then
	iStart = 0
else
	iStart = CInt(iStart)
end if

if Not IsNumeric(iOffset) or Len(iOffset) = 0 then
	iOffset = 50 'Change to alter how many are shown per page
else
	iOffset = Cint(iOffset)
end if

tmpOrder=request("order")
if tmpOrder="" then
	tmpOrder="discountdesc"
else
	if tmpOrder="Name" then
		tmpOrder="discountdesc"
	end if
	if tmpOrder="Code" then
		tmpOrder="discountcode"
	end if
end if
tmpSort=request("sort")
if tmpSort="" then
	tmpSort="ASC"
end if
tmpQuery=tmpQuery & " ORDER BY " & tmpOrder & " " & tmpSort

query="SELECT iddiscount, discountdesc, pricetodiscount, percentagetodiscount, discountcode, active, expDate, onetime, pcSeparate, pcDisc_Auto, pcDisc_StartDate FROM discounts WHERE " & tmpQuery & ";"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if rs.EOF then
	pcvDiscounts = "No"
	set rs = nothing
	call closeDb()
else
	pcArray=rs.getRows()
	Dim iRows, iCols, iRowLoop, iColLoop, iStop
	iRows = UBound(pcArray, 2)
	iCols = UBound(pcArray, 1)
	
	If iRows > (iOffset + iStart) Then
		iStop = iOffset + iStart - 1
	Else
		iStop = iRows
	End If
	
	set rs = nothing
	call closeDb()	
end if
%>

<table class="pcCPcontent">
	<tr>
		<td colspan="10" class="pcCPspacer">The following is a list of current Discount Codes. The &quot;Delete&quot; icon is not shown if the discount was used in an order. In that case, archive the discount to hide it. You can view archived discount codes (if any) by clicking on the corresponding button at the botto of the page.</td>
	</tr>
    <tr>
        <td colspan="10" class="pcCPspacer">
            <% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
        </td>
    </tr>
	<tr> 
		<th width="40%" nowrap><a href="AdminDiscounts.asp?Start=<%=iStart-iOffset %>&Offset=<%=iOffset%>&list=<%=request("list")%>&order=Name&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a><a href="AdminDiscounts.asp?Start=<%=iStart-iOffset %>&Offset=<%=iOffset%>&list=<%=request("list")%>&order=Name&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a>&nbsp;Name<span class="pcSmallText">&nbsp;|&nbsp;<a href="AddDiscounts.asp">Add New</a></span></th>
		<th nowrap><a href="AdminDiscounts.asp?Start=<%=iStart-iOffset %>&Offset=<%=iOffset%>&list=<%=request("list")%>&order=Code&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a><a href="AdminDiscounts.asp?Start=<%=iStart-iOffset %>&Offset=<%=iOffset%>&list=<%=request("list")%>&order=Code&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a>&nbsp;Code</th>
		<th>Active</th>
		<th>Type</th>
        <th>Starts</th>
		<th>Expires</th>
		<th nowrap="nowrap">One-time</th>
		<th>Multiple</th>
		<th>Auto</th>
		<th nowrap="nowrap"></th>
	</tr>
	<tr>
		<td colspan="10" class="pcCPspacer"></td>
	</tr>
	<% If pcvDiscounts = "No" Then %>
        <tr> 
            <td colspan="10"><div class="pcCPmessage">No discounts found</div></td>
        </tr>
	<% Else 
	
		'For i=0 to iRows
		For i = iStart to iStop
			iddiscount=pcArray(0,i)
			discountdesc=pcArray(1,i)
			if discountdesc<>"" AND isNULL(discountdesc)=False then		
				discountdesc=replace(discountdesc,"""","&quot;")
			end if
			
			pricetodiscount=pcArray(2,i)
			if pricetodiscount<>"" then
			else
				pricetodiscount="0"
			end if
			if pricetodiscount>"0" then
				discounttype="1"
			end if

			percentagetodiscount=pcArray(3,i)
			if percentagetodiscount<>"" then
			else
				percentagetodiscount="0"
			end if
			if percentagetodiscount>"0" then
				discounttype="2"
			end if
	
			discountcode=pcArray(4,i)
		
			active=pcArray(5,i)
			
			expDate=pcArray(6,i)
			if not isNull(expDate) and trim(expDate)<>"" then
				expDate=ShowDateFrmt(expDate)
			end if
			if trim(expDate) =  "//" then
				expDate=""
			end if
			if expDate="" or isNull(expDate) then
				expDate="No"
			end if
			
			onetime=pcArray(7,i)
			pcSeparate=pcArray(8,i)
			pcAuto=pcArray(9,i)
	
			startDate=pcArray(10,i)
			if not isNull(startDate) and trim(startDate)<>"" then
				startDate=ShowDateFrmt(startDate)
			end if
			if trim(startDate) =  "//" then
				startDate=""
			end if
			if startDate="" or isNull(startDate) then
				startDate="No"
			end if

			%>
			<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
				<td><a href="modDiscounts.asp?mode=Edit&iddiscount=<%=iddiscount%>"><%=discountdesc%></a></td>
				<td><%=discountcode%></td>
				<td>
					<%
					if active="-1" then
						response.write "Yes"
					else
						if active="0" then
							response.write "No"
						else
							response.write "-"
						end if
					end if
					%>
				</td>
				<td>
					<% 
					if discounttype = 1 then
						response.write scCurSign
					elseif discounttype then
						response.write "%"
					else
						response.write "Other"
					end if
					%>
				</td>
				<td><%=startDate%></td>
				<td><%=expDate%></td>
				<td>
					<%
					if onetime <> 0 then
						response.write "Yes"
					else
						response.write "No"
					end if
					%>
				</td>
				<td>
					<%
					if pcSeparate <> 0 then
						response.write "Yes"
					else
						response.write "No"
					end if
					%>
				</td>
				<td>
					<%
					if pcAuto <> 0 then
						response.write "Yes"
					else
						response.write "No"
					end if
					%>
				</td>
				<td nowrap="nowrap" align="right" class="cpLinksList">
				<%
				call opendb()
				query="SELECT TOP 1 orders.idorder FROM Orders WHERE orders.DiscountDetails<>'No discounts applied.' AND orders.DiscountDetails like '%" & replace(discountdesc, "'", "''") & "%';"
				tmpHaveOrders=0
				set rsQ=connTemp.execute(query)
				if not rsQ.eof then
					tmpHaveOrders=1
				end if
				set rsQ=nothing
				call closedb()
				%>
				
			<a href="modDiscounts.asp?mode=Edit&iddiscount=<%=iddiscount%>" title="Edit this discount code."><img src="images/pcIconGo.jpg" alt="Edit Discount Code" border="0"></a>&nbsp;
			  <%if tmpHaveOrders=0 then%><a href="javascript:if (confirm('You are about to permanently delete this discount code from the database. Instead, you could make it inactive and archive it. Are you sure you want to complete this action?')) location='modDiscounts.asp?mode=Del&iddiscount=<%=iddiscount%>'" title="Delete this discount code."><img src="images/pcIconDelete.jpg" alt="Delete" border="0"></a>
			  <%else%>
              <a href="DiscsalesReport.asp?IDDiscount=<%=iddiscount%>" target="_blank" title="List orders that included this discount code."><img src="images/pcIconList.jpg" width="12" height="12" alt="List orders"></a>
              <%end if%>
              </td>
			</tr>             
		<% Next
	End if %>
	<tr>
		<td colspan="10" class="pcCPspacer"></td>
	</tr>
    <tr>
    	<td colspan="10">
        <%
		dim showPipe
		showPipe = 0
		if iStart > 0 then
			'Show Prev link
			Response.Write "<A HREF=""AdminDiscounts.asp?order="&request("order")&"&sort="&request("sort")&"&Start=" & iStart-iOffset &"&Offset=" & iOffset & """>Previous " & iOffset & "</A>"
			showPipe = 1
		  end if
			
		  if iStop < iRows then
		  	if showPipe = 1 Then
				response.write " | "
			end if
			'Show Next link
			Response.Write " <A HREF=""AdminDiscounts.asp?order="&request("order")&"&sort="&request("sort")&"&Start=" & iStart+iOffset &"&Offset=" & iOffset & """>Next " & iOffset & "</A>"
		  end if
		  %>
	</td>
    </tr>
    
	<tr>
		<td colspan="10">
    	<form class="pcForms">
			<input type="button" value="Add New Discount" onClick="location.href='AddDiscounts.asp'" class="submit2">
			<%
			call opendb()
			if request("list")="archive" then
				tmpQuery="active=-1 OR active=0"
			else
				tmpQuery="active=2"
			end if
			query="SELECT TOP 1 iddiscount FROM discounts WHERE " & tmpQuery & ";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if not rs.eof then
				if request("list")="archive" then%>
					&nbsp;<input type="button" name="archive" value="View Current Discount Codes" onClick="location.href='AdminDiscounts.asp'">
				<%else%>
					&nbsp;<input type="button" name="archive" value="View Archived Discount Codes" onClick="location.href='AdminDiscounts.asp?list=archive'">
				<%end if
			end if
			set rs=nothing
			call closedb()
			%>
            &nbsp;<input type="button" name="back" value="Back" onClick="javascript:history.back()">
        </form>
		</td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->
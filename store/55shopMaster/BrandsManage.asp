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
<%
dim query, conntemp, rs, pcIntParent, pcvParentBrandName, pcv_maxitems
pcIntParent=request("parent")
if not validNum(pcIntParent) then pcIntParent = 0
	if pcIntParent>0 then
		' Load Parent Brand Name
		call openDb()
		query="SELECT BrandName FROM Brands WHERE idBrand="&pcIntParent
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		pcvParentBrandName=rs("BrandName")
		pageTitle="Manage Brands under " & pcvParentBrandName
		set rs=nothing
		call closeDb()
	else
		pcvParentBrandName=""
		pageTitle="Manage Brands"
	end if

pcv_maxitems=50
	
if request("iPageCurrent")="" then
	iPageCurrent=1 
else
	iPageCurrent=Request("iPageCurrent")
end If

dim i, idBrand, priority
sMode=request.form("submitForm")
if sMode<>"" then
	call openDb()
	pcIntParent=request.form("parent")
	iCnt=request.form("iCnt")
	set rs=server.CreateObject("ADODB.RecordSet")
	for i=1 to iCnt
		idBrand=request.form("idBrand"&i)
		priority=request.form("priority"&i)
		query="UPDATE Brands SET "
		query=query & "pcBrands_Order=" &priority
		query=query & " WHERE idBrand="&idBrand&";"
		on error resume next
		rs.open query,conntemp
	next
	set rs=nothing
	call closeDb()	
	response.redirect "BrandsManage.asp?parent="&pcIntParent&"&iPageCurrent=" & Request("iPageCurrent")
	response.end
end if

' Load Brands
call openDb()
query="SELECT idBrand, BrandName, pcBrands_Order FROM Brands WHERE pcBrands_Parent="&pcIntParent&" ORDER BY pcBrands_Order, BrandName ASC"
Set rs=Server.CreateObject("ADODB.Recordset")

rs.CacheSize=pcv_maxitems
rs.PageSize=pcv_maxitems

rs.Open query, connTemp, adOpenStatic, adLockReadOnly

if err.number <> 0 then
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?error="& Server.Urlencode("Error loading brands") 
end If
Dim iCnt, pcIntNoResults
iCnt=0
pcIntNoResults=0

If not rs.eof Then
	rs.MoveFirst

	' get the max number of pages
	Dim iPageCount
	iPageCount=rs.PageCount
	If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
	If iPageCurrent < 1 Then iPageCurrent=1
		
	' set the absolute page
	rs.AbsolutePage=iPageCurrent
Else
	pcIntNoResults=1
End if
%>
<!--#include file="AdminHeader.asp"-->

	<form name="form1" method="post" action="BrandsManage.asp" class="pcForms">
		<input type="hidden" name="parent" value="<%=pcIntParent%>">
        <table class="pcCPcontent">
            <tr>
                <td colspan="3" class="pcCPspacer">
                    <% ' START show message, if any %>
                        <!--#include file="pcv4_showMessage.asp"-->
                    <% 	' END show message %>
                </td>
            </tr>
            <tr>
                <td colspan="3">
                	<%
					if pcvParentBrandName="" then
					%>
                	These are the top level brands. They are displayed in the storefront on the <a href="../pc/viewbrands.asp" target="_blank">View Brands</a> page, based on the settings entered on the <a href="AdminSettings.asp?tab=3#brandSettings">Display Settings</a> page. <u>NOTE</u>: if you assign sub-brands to a brand, move the products assigned to the brand (if any) to the sub-brands you have created. That's because the page that displays sub-brands in the storefront does not also show products.
<%
					else
					%>
                    These are the brands under <strong><a href="BrandsEdit.asp?idbrand=<%=pcIntParent%>"><%=pcvParentBrandName%></a></strong>. Back to the <a href="BrandsManage.asp">Top Level Brands</a>.
<%
					end if
					%>
                </td>
            </tr>
            <tr>
                <td colspan="3" class="pcCPspacer"></td>
            </tr>
			<tr> 					
				<th width="5%">Order</th>
				<th colspan="2">Brand</th>
			</tr>
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr>
			<% 
			if pcIntNoResults=1 then
			%>
            <tr><td colspan="3">No brands found. <a href="BrandsAdd.asp">Add new</a></td></tr>
            <%
			else
				pcArr=rs.getRows(pcv_maxitems)
				set rs=nothing
				intCount=ubound(pcArr,2)
				For i=0 to intCount
				pcvBrandName=pcArr(1,i)
				tidBrand=pcArr(0,i)
				tpriority=pcArr(2,i)
				iCnt=iCnt+1
				
				'// Calculate number of subbrands
					query="SELECT Count(*) As tmpCount FROM Brands WHERE pcBrands_Parent="&tidBrand
					set rsQ=connTemp.execute(query)
					if not rsQ.eof then
						iBrandsCount=rsQ("tmpCount")
					end if
					set rsQ = nothing
					
				'// Calculate number of subbrands
					query="SELECT Count(*) As tmpCount FROM products WHERE active=-1 AND configOnly=0 AND removed=0 AND IDBrand="&tidBrand
					set rsQ=connTemp.execute(query)
					if not rsQ.eof then
						iProductCount=rsQ("tmpCount")
					end if
					set rsQ = nothing				
				 %>
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
					<td> 
						<input type="text" name="priority<%=iCnt%>" size="2" maxlength="4" value="<%=tpriority%>">
					</td>
					<td> 
						<a href="BrandsEdit.asp?idbrand=<%=tidBrand%>"><%=pcvBrandName%></a>
						<input type="hidden" name="idBrand<%=iCnt%>" value="<%=tidBrand%>">
					</td>
                    <td align="right" class="cpLinksList">
                    	<a href="BrandsEdit.asp?idbrand=<%=tidBrand%>">Edit</a> | <% if pcvParentBrandName="" then %><a href="BrandsManage.asp?parent=<%=tidBrand%>">SubBrands (<%=iBrandsCount%>)</a> | <%end if%><a href="BrandsProducts.asp?idbrand=<%=tidBrand%>">Products (<%=iProductCount%>)</a> | <a href="javascript:if (confirm('You are about to permanently remove this Brand from the database. This action CANNOT be undone. You might want to consider making the Brand \'Inactive\' instead (this setting is on the Edit Brand page). Click OK to confirm the removal or CANCEL to keep the current data.')) location='BrandsDel.asp?idbrand=<%=tidBrand%>';">Delete</a>
                    </td>
				</tr>
				<%
				Next
				%>
				<tr>
				<td colspan="7" class="pcCPspacer"></td>
				</tr>
				<%If iPageCount>1 Then%>
				<tr>
					<td colspan="3" class="pcCPspacer"></td>
				</tr>                            
				<tr> 
					<td colspan="3"><%response.Write("Currently viewing page "& iPageCurrent & " of "& iPageCount)%></td>
				</tr>
				<tr>                   
				<td colspan="3"> 
					<%' display Next / Prev buttons
					if iPageCurrent > 1 then %>
					<a href="BrandsManage.asp?parent=<%=pcIntParent%>&iPageCurrent=<%=iPageCurrent-1%>"><img src="../pc/images/prev.gif" width="10" height="10" border="0"></a> 
					<%
					end If
					For I=1 To iPageCount
					If Cint(I)=Cint(iPageCurrent) Then %>
						<b><%=I%></b> 
					<%
					Else
					%>
						<a href="BrandsManage.asp?parent=<%=pcIntParent%>&iPageCurrent=<%=I%>"><%=I%></a> 
					<%
					End If
					Next
					if CInt(iPageCurrent) < CInt(iPageCount) then %>
							<a href="BrandsManage.asp?parent=<%=pcIntParent%>&iPageCurrent=<%=iPageCurrent+1%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
					<%
					end If
					%>
				</td>
				</tr>
				<%End If
			end if
			set rs=nothing
			call closedb()
			%>
			<input type="hidden" name="iCnt" value="<%=iCnt%>">
			<input type="hidden" name="iPageCurrent" value="<%=iPageCurrent%>">
		<tr> 
			<td colspan="3" style="padding-top: 20px;"> 
            <% if pcIntNoResults=0 then %>
			<input name="submitForm" type="submit" class="submit2" value="Update Brands Order">&nbsp;
            <% end if %>
            <input name="addNew" type="button" value="Add New" onClick="document.location.href='BrandsAdd.asp'">
            <input name="back" type="button" value="Back" onClick="javascript:history.go(-1);">
			</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->
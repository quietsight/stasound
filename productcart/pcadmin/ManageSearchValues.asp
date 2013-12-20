<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Custom Search Field Values" %>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="AdminHeader.asp"-->

<%
idSearchField=request("idSearchField")

if idSearchField="" then
	response.redirect "ManageSearchFields.asp"
else
	if (Not IsNumeric(idSearchField)) OR (idSearchField="0") then
		response.redirect "ManageSearchFields.asp"
	end if
end if

Dim rstemp, connTemp, query, rs, rsQ

call opendb()

tmp_pagesize=50

if request("iPageCurrent")="" then
	if request("iPageCurrent1")="" then
		iPageCurrent=1
	else
		iPageCurrent=Request("iPageCurrent1")
	end if
else
    iPageCurrent=Request("iPageCurrent")
end If

'sorting order
Dim strORD

strORD=request("order")
if strORD="" then
	strORD="idSearchData"
End If

strSort=request("sort")
if strSort="" Then
	strSort="ASC"
End If

IF request("action")="update" then
IF request("submit2")<>"" THEN
	SFV_name=request("new_valuename")
	if SFV_name<>"" then
		SFV_name=replace(SFV_name,"'","''")
	end if
	SFV_order=request("new_valueorder")
	if SFV_order="" then
		SFV_order=0
	end if
	query="SELECT idSearchData FROM pcSearchData WHERE pcSearchDataName like '" & SFV_name & "' AND idSearchField=" & idSearchField & ";"
	set rsQ=connTemp.execute(query)
	if not rsQ.eof then
		msg="The field value was already added!"
		set rsQ=nothing
	else
		query="INSERT INTO pcSearchData (idSearchField,pcSearchDataName,pcSearchDataOrder) VALUES (" & idSearchField & ",'" & SFV_name & "'," & SFV_order & ");"
		set rsQ=connTemp.execute(query)
		set rsQ=nothing
		msg="The new custom search field value was added successfully!"
		msgtype=1
	end if
	set rsQ=nothing
ELSE
	Count=request("Count")
	IF request("Count")<>"" then
		For k=1 to clng(Count)
			if request("C"&k)="1" then
				SFV_id=request("SFV_id"&k)
				SFV_name=request("SFV_name"&k)
				SFV_name=replace(SFV_name,"'","''")
				SFV_order=request("SFV_order"&k)
				if SFV_order="" then
					SFV_order=0
				end if
				query="UPDATE pcSearchData SET pcSearchDataName='" & SFV_name & "',pcSearchDataOrder=" & SFV_order & " WHERE idSearchData=" & SFV_id & ";"
				set rsQ=connTemp.execute(query)
				set rsQ=nothing
				msg="Custom Search Fields were updated successfully!"
				msgtype=1
			end if
		Next
	END IF
END IF
END IF

%>

<script>
function newWindow(file,window)
	{
		msgWindow=open(file,window,'resizable=yes,scrollbars=yes,width=400,height=500');
		if (msgWindow.opener == null) msgWindow.opener = self;
	}
</script>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<%	query="SELECT idSearchData,pcSearchDataName,pcSearchDataOrder FROM pcSearchData WHERE idSearchField=" & idSearchField & " ORDER BY "& strORD &" "& strSort
	set rstemp=server.CreateObject("ADODB.RecordSet")
	rstemp.CacheSize=tmp_pagesize
	rstemp.PageSize=tmp_pagesize
	rstemp.Open query, connTemp, adOpenStatic, adLockReadOnly
	%>

	<form name="form1" action="ManageSearchValues.asp?idSearchField=<%=idSearchField%>&action=update" method="post" class="pcForms">
		<table class="pcCPcontent">
			<tr>
				<td colspan="4">
					<%query="SELECT pcSearchFieldName FROM pcSearchFields WHERE idSearchField=" & idSearchField & ";"
					set rsQ=connTemp.execute(query)
					if not rsQ.eof then%>
						<h2>Search Field Name: <b><%=rsQ("pcSearchFieldName")%></b></h2>
					<%end if
					set rsQ=nothing%>
				</td>
			</tr>
			<tr>
				<th nowrap valign="top"><a href="ManageSearchValues.asp?idSearchField=<%=idSearchField%>&order=idSearchData&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="ManageSearchValues.asp?idSearchField=<%=idSearchField%>&order=idSearchData&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;<b>ID</b></th>
				<th nowrap valign="top"><a href="ManageSearchValues.asp?idSearchField=<%=idSearchField%>&order=pcSearchDataName&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="ManageSearchValues.asp?idSearchField=<%=idSearchField%>&order=pcSearchDataName&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;<b>Value Name</b></th>
				<th nowrap valign="top" colspan="2"><a href="ManageSearchValues.asp?idSearchField=<%=idSearchField%>&order=pcSearchDataOrder&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="ManageSearchValues.asp?idSearchField=<%=idSearchField%>&order=pcSearchDataOrder&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;<b>Value Order</b></th>
			</tr>
			<tr> 
				<td colspan="4" class="pcCPspacer"></td>
			</tr>
			<% If rstemp.eof Then %>
				<tr> 
					<td colspan="4"><div class="pcCPmessage">No values have so far been associated with this customer search field. Add the values to the field using the form below.</div></td>
				</tr>
			<% Else
				rstemp.MoveFirst

				' get the max number of pages
				Dim iPageCount
				iPageCount=rstemp.PageCount
				If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
				If iPageCurrent < 1 Then iPageCurrent=1
				
				' set the absolute page
				rstemp.AbsolutePage=iPageCurrent
			
				pcArray=rstemp.GetRows(tmp_pagesize)
				intCount=ubound(pcArray,2)
				
				pcv_HaveSearchFields=1
				
				set rstemp=nothing
				
				Dim Count
				Count=0
				For k=0 to intCount
					Count=Count+1
					SFV_id=pcArray(0,k)
					SFV_name=pcArray(1,k)
					if SFV_name<>"" then
						SFV_name=replace(SFV_name,"""","&quot;")
					end if
					SFV_order=pcArray(2,k)
					%>
					<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
						<td nowrap>
							<input type="checkbox" name="C<%=Count%>" value="1" class="clearBorder">
              &nbsp;<%=SFV_id%>
							<input type="hidden" name="SFV_id<%=Count%>" value="<%=SFV_id%>">
						</td>
						<td nowrap>
							<input type="text" size="50" name="SFV_name<%=Count%>" value="<%=SFV_name%>">
						</td>
						<td nowrap>
							<input type="text" size="4" name="SFV_order<%=Count%>" value="<%=SFV_order%>">
						</td>
						<td nowrap align="right">
							<%
							query="SELECT TOP 1 idproduct FROM Products WHERE products.removed=0 AND products.configOnly=0 AND (products.idproduct IN (SELECT idProduct FROM pcSearchFields_Products WHERE idSearchData=" & SFV_id & "));"
							set rsQ=connTemp.execute(query)
							pcv_HavePrds=0
							if not rsQ.eof then
							pcv_HavePrds=1%>
								<a href="javascript: newWindow('showCFProducts.asp?idcustom=S<%=idSearchField%>&SearchValues=<%=SFV_id%>','products');" title="List products using it"><img src="images/pcIconList.jpg" alt="List products using it"></a>&nbsp<%end if
							set rsQ=nothing%>
              <a href="javascript: if (confirm('Are you sure you want to permanently delete this custom field value from the database? <%if pcv_HavePrds=1 then%>It will also be removed from all products to which it had been assigned.<%end if%>?')) {location='delSearchFieldValue.asp?idSearchField=<%=idSearchField%>&idSearchData=<%=SFV_id%>';}" title="Delete"><img src="images/pcIconDelete.jpg" alt="Delete"></a>
							
						</td>
					</tr>     
				<%Next%>
				<tr>
					<td colspan="4" class="pcCPspacer"></td>
				</tr>
				<tr>
					<td colspan="4" class="cpLinksList"><a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a></td>
				</tr>
				<tr> 
					<td colspan="4" class="pcSpacer"><hr></td>
				</tr>
				<%If iPageCount>1 Then
				%>
				<tr> 
					<td colspan="4"><%response.Write("Currently viewing page "& iPageCurrent & " of "& iPageCount)%></td>
				</tr>
				<tr>                   
					<td colspan="4"> 
					<%' display Next / Prev buttons
					if iPageCurrent > 1 then %>
					<a href="ManageSearchValues.asp?idSearchField=<%=idSearchField%>&iPageCurrent=<%=iPageCurrent-1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/prev.gif" width="10" height="10" border="0"></a> 
					<%
					end If
					For I=1 To iPageCount
					If Cint(I)=Cint(iPageCurrent) Then %>
						<b><%=I%></b> 
					<%
					Else
					%>
						<a href="ManageSearchValues.asp?idSearchField=<%=idSearchField%>&iPageCurrent=<%=I%>&order=<%=strORD%>&sort=<%=strSort%>"> 
						<%=I%></a> 
					<%
					End If
					Next
					if CInt(iPageCurrent) < CInt(iPageCount) then %>
						<a href="ManageSearchValues.asp?idSearchField=<%=idSearchField%>&iPageCurrent=<%=iPageCurrent+1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
					<%end If%>
					</td>
				</tr>
				<tr>
					<td colspan="5" class="pcCPspacer"><hr></td>
				</tr>
				<% End If %>
                <tr>
                <td colspan="5">
                    <input type="hidden" name="idSearchField" value="<%=idSearchField%>">
                    <input name="Count" type="hidden" value="<%=Count%>"><input name="iPageCurrent1" type="hidden" value="<%=iPageCurrent%>">
                    <%if count<>"" then%>
                    <input name="submit1" type="submit" value="Update Selected" size="20" class="submit2">
                    <%end if%> 
                    <input type="button" value="Back" onClick="location='ManageSearchFields.asp';">
                </td>
                </tr>
				<tr>
					<td colspan="5" class="pcCPspacer"><hr></td>
				</tr>
			<%
			End If
			set rstemp=nothing
			call closedb()
			%>
			<tr>
				<td colspan="5">
					<table class="pcCPcontent">
					<tr>
						<td colspan="2">
							<b>Add new custom search field value:</b>
						</td>
					</tr>
					<tr>
						<td width="75" nowrap><div align="right">New Value:</div></td>
						<td nowrap> 
							<input name="new_valuename" type="text" size="20" maxlength="150">
						</td>
					</tr>
					<tr> 
						<td><div align="right">Order:</div></td>
						<td><input name="new_valueorder" type="text" size="4" maxlength="150" value="0"> 
					    <span class="pcSmallText">&nbsp;The order in which the value is shown when all values associated with this search field are presented.</span></td>
					</tr>
					<tr>       
						<td>&nbsp;</td>
						<td>
							<input type="submit" name="submit2" value="Add New" class="submit2" onclick="javascript: if (document.form1.new_valuename.value=='') {alert('Please enter a value in the New Value field'); document.form1.new_valuename.focus();return(false)} else {return(true)};"> 
                            <input type="button" value="Back" onClick="location='ManageSearchFields.asp';">
						</td>
					</tr>
					</table>
				</td>
			</tr>
		</table>
		<input type="hidden" name="order" value="<%=strORD%>">
		<input type="hidden" name="sort" value="<%=strSort%>">
		<%if count<>"" then%>
			<script language="JavaScript">
			<!--
			function checkAll() {
			for (var j = 1; j <= <%=count%>; j++) {
			box = eval("document.form1.C" + j); 
			if (box.checked == false) box.checked = true;
				 }
			}
			
			function uncheckAll() {
			for (var j = 1; j <= <%=count%>; j++) {
			box = eval("document.form1.C" + j); 
			if (box.checked == true) box.checked = false;
				 }
			}
			
			//-->
			</script>
		<%end if%>
	</form>
<!--#include file="AdminFooter.asp"-->
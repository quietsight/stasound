<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Custom Search Fields" %>
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

<% Dim rstemp, connTemp, query, rs, rsQ

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
	strORD="pcSearchFieldOrder"
End If

strSort=request("sort")
if strSort="" Then
	strSort="ASC"
End If

IF request("action")="update" then
	IF request("submit2")<>"" THEN
		SF_name=request("SF_name")
		SF_name=replace(SF_name,"'","''")
		SF_show=request("SF_show"&k)
		CP_show=request("CP_show"&k)
		SF_showSEARCH=request("SF_showSEARCH"&k)
		CP_showSEARCH=request("CP_showSEARCH"&k)
		if SF_show="" then
			SF_show=0
		end if
		if CP_show="" then
			CP_show=0
		end if
		if SF_showSEARCH="" then
			SF_showSEARCH=0
		end if
		if CP_showSEARCH="" then
			CP_showSEARCH=0
		end if		
		SF_order=request("SF_order")
		if SF_order="" then
			SF_order=0
		end if
		query="SELECT idSearchField FROM pcSearchFields WHERE pcSearchFieldName like '" & SF_name & "';"
		set rsQ=connTemp.execute(query)
		if not rsQ.eof then
			msg="The field name was already added!"
			set rsQ=nothing
		else
			query="INSERT INTO pcSearchFields (pcSearchFieldName,pcSearchFieldOrder,pcSearchFieldShow,pcSearchFieldCPShow,pcSearchFieldSearch,pcSearchFieldCPSearch) VALUES ('" & SF_name & "'," & SF_order & "," & SF_show & "," & CP_show & "," & SF_showSEARCH & "," & CP_showSEARCH & ");"
			set rsQ=connTemp.execute(query)
			set rsQ=nothing
			msg="The new custom search field was added successfully!"
		end if
		set rsQ=nothing
	ELSE
		Count=request("Count")
		IF request("Count")<>"" then
			For k=1 to clng(Count)
				if request("C"&k)="1" then
					SF_id=request("SF_id"&k)
					SF_name=request("SF_name"&k)
					SF_name=replace(SF_name,"'","''")
					SF_show=request("SF_show"&k)
					CP_show=request("CP_show"&k)
					SF_showSEARCH=request("SF_showSEARCH"&k)
					CP_showSEARCH=request("CP_showSEARCH"&k)
					if SF_show="" then
						SF_show=0
					end if
					if CP_show="" then
						CP_show=0
					end if
					if SF_showSEARCH="" then
						SF_showSEARCH=0
					end if
					if CP_showSEARCH="" then
						CP_showSEARCH=0
					end if
					SF_order=request("SF_order"&k)
					if SF_order="" then
						SF_order=0
					end if
					query="UPDATE pcSearchFields SET pcSearchFieldName='" & SF_name & "',pcSearchFieldOrder=" & SF_order & ",pcSearchFieldShow=" & SF_show & ",pcSearchFieldCPShow=" & CP_show & ",pcSearchFieldSearch=" & SF_showSEARCH & ",pcSearchFieldCPSearch=" & CP_showSEARCH & " WHERE idSearchField=" & SF_id & ";"
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
		msgWindow=open(file,window,'resizable=yes,scrollbars=yes,width=475,height=500');
		if (msgWindow.opener == null) msgWindow.opener = self;
	}
</script>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<%	query="SELECT idSearchField,pcSearchFieldName,pcSearchFieldOrder,pcSearchFieldShow,pcSearchFieldCPShow,pcSearchFieldSearch,pcSearchFieldCPSearch FROM pcSearchFields ORDER BY "& strORD &" "& strSort
	set rstemp=server.CreateObject("ADODB.RecordSet")
	rstemp.CacheSize=tmp_pagesize
	rstemp.PageSize=tmp_pagesize
	rstemp.Open query, connTemp, adOpenStatic, adLockReadOnly
	%>

	<form name="form1" action="ManageSearchFields.asp?action=update" method="post" class="pcForms">
		<table class="pcCPcontent">
    	<tr>
      	<td colspan="8" class="pcCPspacer"></td>
      </tr>
			<tr>
				<th valign="top" nowrap="nowrap">ID
        <div style="margin: 2px 0 2px 0;"><a href="ManageSearchFields.asp?order=idSearchField&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="ManageSearchFields.asp?order=idSearchField&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>
         </div>
        </th>
			  <th valign="top" nowrap="nowrap">Field Name <a href="#new" title="Add new field"><img src="images/pcIconPlus.jpg" alt="Add new"></a>
   	    <div style="margin: 2px 0 2px 0;"><a href="ManageSearchFields.asp?order=pcSearchFieldName&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="ManageSearchFields.asp?order=pcSearchFieldName&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a></div>
        </th>
        <th valign="top" nowrap="nowrap">Order&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=453')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a><div style="margin: 2px 0 2px 0;"><a href="ManageSearchFields.asp?order=pcSearchFieldOrder&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="ManageSearchFields.asp?order=pcSearchFieldOrder&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a></div>
        </th>
<th valign="top" nowrap="nowrap">CP Search&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=454')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
        <div style="margin: 2px 0 2px 0; font-weight: normal;"><a href="ManageSearchFields.asp?order=pcSearchFieldCPSearch&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="ManageSearchFields.asp?order=pcSearchFieldCPSearch&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a></div>         
        </th>
        <th valign="top" nowrap="nowrap">SF Search&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=454')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
        <div style="margin: 2px 0 2px 0; font-weight: normal;"><a href="ManageSearchFields.asp?order=pcSearchFieldSearch&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="ManageSearchFields.asp?order=pcSearchFieldSearch&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;</div>
      </th>
      <th valign="top" nowrap="nowrap">CP Details&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=455')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
        <div style="margin: 2px 0 2px 0; font-weight: normal;"><a href="ManageSearchFields.asp?order=pcSearchFieldCPShow&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="ManageSearchFields.asp?order=pcSearchFieldCPShow&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a></div>
       </th>
      <th valign="top" nowrap="nowrap" colspan="2">SF Details&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=455')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
        <div style="margin: 2px 0 2px 0; font-weight: normal;"><a href="ManageSearchFields.asp?order=pcSearchFieldShow&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="ManageSearchFields.asp?order=pcSearchFieldShow&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a></div>
      </th>
		  </tr>
			<tr style="background-color: #f5f5f5;">
      	<td valign="top"></td>
        <td valign="top"></td>
        <td valign="top"></td>
        <td valign="top" nowrap><input type="checkbox" name="DSK" value="1" onclick="javascript:return checkCPS(this);" class="clearBorder"> Toggle All</td>
        <td valign="top" nowrap><input type="checkbox" name="DSK" value="1" onclick="javascript:return checkSFS(this);" class="clearBorder"> Toggle All</td>
        <td valign="top" nowrap><input type="checkbox" name="DSK" value="1" onclick="javascript:return checkCPSHW(this);" class="clearBorder"> Toggle All</td>
        <td valign="top" nowrap><input type="checkbox" name="DSK" value="1" onclick="javascript:return checkSHW(this);" class="clearBorder"> Toggle All</td>
        <td valign="top"></td>
			</tr>
			<tr> 
				<td colspan="8" class="pcSpacer"></td>
			</tr>
			<% If rstemp.eof Then %>
				<tr> 
				<td colspan="8"><div class="pcCPmessage">No Custom Search Fields Found</div></tr>
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
				
				Dim strCol, Count
				Count=0
				strCol="#E1E1E1"
				For k=0 to intCount
					Count=Count+1
					SF_id=pcArray(0,k)
					SF_name=pcArray(1,k)
					if SF_name<>"" then
						SF_name=replace(SF_name,"""","&quot;")
					end if
					SF_order=pcArray(2,k)
					SF_show=pcArray(3,k)
					CP_show=pcArray(4,k)
					SF_showSEARCH=pcArray(5,k)
					CP_showSEARCH=pcArray(6,k)
					
					If strCol <> "#FFFFFF" Then
						strCol="#FFFFFF"
					Else 
						strCol="#E1E1E1"
					End If %>
					<tr bgcolor="<%= strCol %>">
						<td>
							<input type="checkbox" name="C<%=Count%>" value="1" class="clearBorder">
              &nbsp;<%=SF_id%>
							<input type="hidden" name="SF_id<%=Count%>" value="<%=SF_id%>"></td>
						<td><input type="text" size="18" name="SF_name<%=Count%>" value="<%=SF_name%>"></td>
						<td><input type="text" size="4" name="SF_order<%=Count%>" value="<%=SF_order%>"></td>
						<td><input type="checkbox" name="CP_showSEARCH<%=Count%>" value="1" <%if CP_showSEARCH="1" then%>checked<%end if%> class="clearBorder"></td>
						<td><input type="checkbox" name="SF_showSEARCH<%=Count%>" value="1" <%if SF_showSEARCH="1" then%>checked<%end if%> class="clearBorder"></td>
						<td><input type="checkbox" name="CP_show<%=Count%>" value="1" <%if CP_show="1" then%>checked<%end if%> class="clearBorder"></td>
            <td><input type="checkbox" name="SF_show<%=Count%>" value="1" <%if SF_show="1" then%>checked<%end if%> class="clearBorder"></td>
				  	<td nowrap><%query="SELECT idSearchData FROM pcSearchData WHERE idSearchField=" & SF_id & ";"
							set rsQ=connTemp.execute(query)
							if not rsQ.eof then%>
								<a href="ManageSearchValues.asp?idSearchField=<%=SF_id%>" title="Display values"><img src="images/pcIconGo.jpg" alt="Display values"></a>
		<%else%>
								<a href="ManageSearchValues.asp?idSearchField=<%=SF_id%>" title="Add values"><img src="images/pcIconPlus.jpg" alt="Add values"></a>
							<%end if
							set rsQ=nothing
							
							query="SELECT TOP 1 idproduct FROM Products WHERE products.removed=0 AND products.configOnly=0 AND (products.idproduct IN (SELECT DISTINCT pcSearchFields_Products.idProduct FROM pcSearchFields_Products INNER JOIN pcSearchData ON pcSearchFields_Products.idSearchData=pcSearchData.idSearchData WHERE pcSearchData.idSearchField=" & SF_id & "));"
							set rsQ=connTemp.execute(query)
							pcv_HavePrds=0
							if not rsQ.eof then
							pcv_HavePrds=1%>
								&nbsp;<a href="javascript: newWindow('showCFProducts.asp?idcustom=S<%=SF_id%>','products');" title="List products using this search field"><img src="images/pcIconList.jpg" align="List products using this search field"></a>
							<%end if
							set rsQ=nothing
							%>
                            <%
							query="SELECT categories.idCategory, categories.categoryDesc, categories.idParentCategory FROM categories INNER JOIN pcSearchFields_Categories ON categories.idCategory = pcSearchFields_Categories.idCategory WHERE pcSearchFields_Categories.idSearchData="&SF_id&" ORDER BY categories.categoryDesc, categories.idParentCategory;"
							set rsQ=connTemp.execute(query)
							if not rsQ.eof then
								intTotCat=1
							else
								intTotCat=0
							end if
							set rsQ=nothing
							%>
              	&nbsp;<a href="javascript: newWindow('showCFCategories.asp?intTotCat=<%=intTotCat%>&idcustom=<%=SF_id%>','products');" title="Show categories associated with this search field"><img src="images/pcIconNext.jpg" alt="Show categories associated with this search field"></a>
								&nbsp;<a href="javascript: if (confirm('Are you sure you want to permanently delete this custom field from database<%if pcv_HavePrds=1 then%>, also permanently delete it from all of products<%end if%>?')) {location='delSearchField.asp?idSearchField=<%=SF_id%>';}" title="Delete"><img src="images/pcIconDelete.jpg" alt="Delete"></a>
              </td>
					</tr>     
				<%Next%>
				<tr>
					<td colspan="8">&nbsp;</td>
				</tr>
				<tr>
					<td colspan="8"><b><a href="javascript:checkAll();">Check All</a></b><b> |&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a></b></td>
				</tr>
				<tr> 
					<td colspan="8" class="pcSpacer"><hr></td>
				</tr>
				<%If iPageCount>1 Then
				%>
				<tr> 
					<td colspan="8"><%response.Write("Currently viewing page "& iPageCurrent & " of "& iPageCount)%></td>
				</tr>
				<tr>                   
					<td colspan="8"> 
					<%' display Next / Prev buttons
					if iPageCurrent > 1 then %>
					<a href="ManageSearchFields.asp?iPageCurrent=<%=iPageCurrent-1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/prev.gif" width="10" height="10" border="0"></a> 
					<%
					end If
					For I=1 To iPageCount
					If Cint(I)=Cint(iPageCurrent) Then %>
						<b><%=I%></b> 
					<%
					Else
					%>
						<a href="ManageSearchFields.asp?iPageCurrent=<%=I%>&order=<%=strORD%>&sort=<%=strSort%>"> 
						<%=I%></a> 
					<%
					End If
					Next
					if CInt(iPageCurrent) < CInt(iPageCount) then %>
						<a href="ManageSearchFields.asp?iPageCurrent=<%=iPageCurrent+1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
					<%end If%>
          </td>
				</tr>
				<tr>
					<td colspan="8" class="pcCPspacer"><hr></td>
				</tr>
				<% End If %>
			<%
			End If
			set rstemp=nothing
			%>
			<tr>
        <td align="center" colspan="8" style="padding-bottom: 20px;">
          <input type="hidden" name="order" value="<%=strORD%>">
          <input type="hidden" name="sort" value="<%=strSort%>">
          <input name="Count" type="hidden" value="<%=Count%>">
          <input name="iPageCurrent1" type="hidden" value="<%=iPageCurrent%>">

          <%if count<>"" then%>
          <input name="submit1" type="submit" value="Update Selected" class="submit2">
          &nbsp;
          <input type="button" name="button" value="Assign to Products" onClick="document.location.href='addSFtoPrds.asp?nav='">
          &nbsp;
          <input type="button" name="button2" value="Map to Export Fields" onClick="document.location.href='SearchFields_Export.asp'">          
          <%end if%>
          &nbsp;
          <input type="button" value="Back" onClick="document.location.href='ManageCFields.asp';">
          &nbsp;
          <input type="button" value="Help" onClick="window.open('http://wiki.earlyimpact.com/productcart/managing_search_fields');">
        </td>
			</tr>
			<tr>
				<td colspan="8" style="background-color: #f5f5f5; padding: 10px;">
        	<a name="new">&nbsp;</a><br />
					<table class="pcCPcontent">
					<tr>
						<td colspan="2">
							<b>Add new custom search field:</b>
             </td>
					</tr>
					<tr>
						<td nowrap><div align="right">Field Name:</div></td>
						<td nowrap><input name="SF_name" type="text" id="SF_name" size="20" maxlength="150"></td>
					</tr>
					<tr> 
						<td><div align="right">Field Order:</div></td>
						<td><input name="SF_order" type="text" id="SF_order" value="0" size="4" maxlength="150"></td>
					</tr>
					<tr> 
						<td nowrap><div align="right">Display Options:</div></td>
						<td>&nbsp;</td>
					</tr>
					<tr> 
						<td align="right"><input type="checkbox" name="CP_showSEARCH" value="1" class="clearBorder"></td>
						<td>Display on the Advanced Search pages in the Control Panel</td>
					</tr>
					<tr> 
						<td align="right"><input type="checkbox" name="SF_showSEARCH" value="1" class="clearBorder"></td>
						<td>Display on the Advanced Search page in the storefront (search.asp) </td>
					</tr>
					<tr> 
						<td align="right"><input type="checkbox" name="CP_show" value="1" class="clearBorder"></td>
						<td>Display on the Modify Product page in the Control Panel</td>
					</tr>
					<tr> 
						<td align="right"><input type="checkbox" name="SF_show" value="1" class="clearBorder"></td>
						<td>Display on the Product Details page in the storefront (viewPrd.asp)</td>
					</tr>
					<tr>       
						<td>&nbsp;</td>
						<td><input type="submit" name="submit2" value="Add New" class="submit2" onclick="javascript: if (document.form1.new_fieldname.value=='') {alert('Please enter a value for Field Name'); document.form1.new_fieldname.focus();return(false)} else {return(true)};"></td>
					</tr>
			  </table>
       </td>
			</tr>
		</table>

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
			
			// Storefront Search
			function checkAllSFS() {
			for (var j = 1; j <= <%=count%>; j++) {
			box = eval("document.form1.SF_showSEARCH" + j); 
			if (box.checked == false) box.checked = true;
				 }
			}
			
			function uncheckAllSFS() {
			for (var j = 1; j <= <%=count%>; j++) {
			box = eval("document.form1.SF_showSEARCH" + j); 
			if (box.checked == true) box.checked = false;
				 }
			}
			
			function checkSFS(tmpItem)
			{
				if (tmpItem.checked==true)
				{
					checkAllSFS();
				}
				else
				{
					uncheckAllSFS();
				}
				return(true);
			}
			
			
			// Control Panel Search
			function checkAllCPS() {
			for (var j = 1; j <= <%=count%>; j++) {
			box = eval("document.form1.CP_showSEARCH" + j); 
			if (box.checked == false) box.checked = true;
				 }
			}
			
			function uncheckAllCPS() {
			for (var j = 1; j <= <%=count%>; j++) {
			box = eval("document.form1.CP_showSEARCH" + j); 
			if (box.checked == true) box.checked = false;
				 }
			}
			
			function checkCPS(tmpItem)
			{
				if (tmpItem.checked==true)
				{
					checkAllCPS();
				}
				else
				{
					uncheckAllCPS();
				}
				return(true);
			}
			

			// Storefront Product Pages
			function checkAllSHW() {
			for (var j = 1; j <= <%=count%>; j++) {
			box = eval("document.form1.SF_show" + j); 
			if (box.checked == false) box.checked = true;
				 }
			}
			
			function uncheckAllSHW() {
			for (var j = 1; j <= <%=count%>; j++) {
			box = eval("document.form1.SF_show" + j); 
			if (box.checked == true) box.checked = false;
				 }
			}
			
			function checkSHW(tmpItem)
			{
				if (tmpItem.checked==true)
				{
					checkAllSHW();
				}
				else
				{
					uncheckAllSHW();
				}
				return(true);
			}
			
			// Control Panel Product Pages
			function checkAllCPSHW() {
			for (var j = 1; j <= <%=count%>; j++) {
			box = eval("document.form1.CP_show" + j); 
			if (box.checked == false) box.checked = true;
				 }
			}
			
			function uncheckAllCPSHW() {
			for (var j = 1; j <= <%=count%>; j++) {
			box = eval("document.form1.CP_show" + j); 
			if (box.checked == true) box.checked = false;
				 }
			}
			
			function checkCPSHW(tmpItem)
			{
				if (tmpItem.checked==true)
				{
					checkAllCPSHW();
				}
				else
				{
					uncheckAllCPSHW();
				}
				return(true);
			}



			//-->
			</script>
		<%end if%>
	</form>
<!--#include file="AdminFooter.asp"-->
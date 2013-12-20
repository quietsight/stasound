<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="Product Reviews - Product Exclusion List"
pageIcon="pcv4_icon_reviews.png"
section="reviews" 
%>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%Dim rs, connTemp, query
private const MaxRecords=15 'Max Records to display per page

if request("iPageCurrent")="" then
    iPageCurrent=1 
else
    iPageCurrent=Request("iPageCurrent")
end If

pcPageName="prv_PrdExc.asp"

call openDb()

IF request("action")="update" then
	Count=request("count")
	if (Count>"0") and (IsNumeric(Count)) then
		For k=1 to Count
			if request("C" & k)="1" then
				pcv_ID=request("ID" & k)
				query="DELETE FROM pcRevExc WHERE pcRE_IDProduct=" & pcv_ID
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				set rs=nothing
			end if
		Next
	end if
END IF

query="SELECT pcRevExc.pcRE_IDProduct,products.description FROM pcRevExc,products WHERE products.idproduct=pcRevExc.pcRE_IDProduct order by products.description asc"
Set rsInv=Server.CreateObject("ADODB.Recordset")
rsInv.CacheSize=MaxRecords
rsInv.PageSize=MaxRecords
rsInv.Open query, connTemp, adOpenStatic, adLockReadOnly
	
if rsInv.eof then
	DataEmpty=1
else
	DataEmpty=0
end if

%>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<form method="POST" action="prv_PrdExc.asp?action=update" name="checkboxform" class="pcForms">
	<table class="pcCPcontent">
		<% If DataEmpty=1 Then %>
			<tr> 
                <td colspan="3">
					<div class="pcCPmessage">
						No products found. Product reviews apply to all products.
					</div>
				</td>
			</tr>
		<% Else %>
			<tr>
	            <td colspan="3">This is a list of products for which customers cannot view and/or post reviews.</td>
			</tr>
			<tr>
	            <td colspan="3" class="pcCPspacer"></td>
			</tr>
			<tr>
	            <th>ID</th>
    	        <th nowrap colspan="2">Product Name</th>
			</tr>
			<tr>
    	        <td colspan="3" class="pcCPspacer"></td>
			</tr>
			<%
			rsInv.MoveFirst
			' get the max number of pages
			Dim iPageCount
			iPageCount=rsInv.PageCount
			If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
			If iPageCurrent < 1 Then iPageCurrent=1
				
			' set the absolute page
			rsInv.AbsolutePage=iPageCurrent  
			
			Count=0
			Do While NOT rsInv.EOF And Count < rsInv.PageSize
				Count=Count+1
				pcv_ID=rsInv("pcRE_IDProduct")
				pcv_Name=rsInv("description")
				%>
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
                    <td nowrap width="5%"><%=pcv_ID%></td>
                    <td width="90%"><a href="FindProductType.asp?id=<%=pcv_ID%>" target="_blank"><%=pcv_Name%></a></td>
                    <td align="right" width="5%">
                    <input type="checkbox" size="3" name="C<%=count%>" value="1" class="clearBorder">
                    <input type="hidden" name="ID<%=count%>" value="<%=pcv_ID%>">
                    </td>
				</tr>
			<% 
			rsInv.MoveNext
			Loop 
		%>
			<%if count>"0" then%>
                <tr>
                <td colspan="3" class="cpLinksList" align="right">
                    <script language="JavaScript">
                    <!--
                    function checkAll() {
                    for (var j = 1; j <= <%=count%>; j++) {
                    box = eval("document.checkboxform.C" + j); 
                    if (box.checked == false) box.checked = true;
                         }
                    }
                    
                    function uncheckAll() {
                    for (var j = 1; j <= <%=count%>; j++) {
                    box = eval("document.checkboxform.C" + j); 
                    if (box.checked == true) box.checked = false;
                         }
                    }
                    
                    //-->
                    </script>
                    <input type="hidden" name="count" value=<%=count%>>
                    <a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a>
                </td>
                </tr>
            <%end if%>
		<% End If %>
        
		<%If iPageCount>1 Then%>                     
		<tr> 
			<td colspan="3" class="cpLinksList">
			<%response.Write("Currently viewing page "& iPageCurrent & " of "& iPageCount)%>
            &nbsp;|&nbsp; 
			<%' display Next / Prev buttons
			if iPageCurrent > 1 then %>
			<a href="<%=pcPageName%>?iPageCurrent=<%=iPageCurrent-1%>"><img src="../pc/images/prev.gif" width="10" height="10" border="0"></a> 
			<%
			end If
			For I=1 To iPageCount
			If Cint(I)=Cint(iPageCurrent) Then %>
				<b><%=I%></b> 
			<%
			Else
			%>
				<a href="<%=pcPageName%>?iPageCurrent=<%=I%>"><%=I%></a> 
			<%
			End If
			Next
			if CInt(iPageCurrent) < CInt(iPageCount) then %>
					<a href="<%=pcPageName%>?iPageCurrent=<%=iPageCurrent+1%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
			<%
			end If
			%>
		</td>
		</tr>
		<%End If%>
	  	<tr>
			<td colspan="3" class="pcCPspacer"><hr></td>
		</tr>  
		<tr>
            <td colspan="3">
			<input type="button" value="Add Products to Exclusion List" onClick="location.href='prv_AddPrdExc.asp'" class="submit2">&nbsp;
			<%if DataEmpty<>1 then%>
                <input type="submit" value="Remove Selected" name="submit" onclick="return(confirm('You are about to remove the selected products from the exclusion list. Are you sure you want to complete this action?'));">  
            <%end if%>
			</td>
		</tr>
		<tr>
		<td align="left" colspan="3">
			<input type="button" value="Manage Fields" onClick="location.href='prv_FieldManager.asp'">&nbsp;
			<input type="button" value="Product Review Settings" onClick="location.href='PrvSettings.asp'">
			</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->
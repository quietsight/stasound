<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="Product Reviews - Product Specific Settings" 
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
<%

Dim rs, connTemp, query

call openDb()
	
IF request("action")="update" then
	Count=request("count")
	if (Count>"0") and (IsNumeric(Count)) then
		For k=1 to Count
			if request("C" & k)="1" then
				pcv_ID=request("ID" & k)
				query="SELECT pcRev_IDReview FROM pcReviews WHERE pcRev_IDProduct=" & pcv_ID
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				if rs.eof then
					query="DELETE FROM pcReviewSpecials WHERE pcRS_IDProduct=" & pcv_ID
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=connTemp.execute(query)
					set rstemp=nothing
				else
					msg="One or more of the selected products could not be edited because product reviews have already been created for them."
				end if
				set rs=nothing
			end if
		Next
	end if
END IF

query="SELECT pcReviewSpecials.pcRS_IDProduct,products.description FROM pcReviewSpecials,products WHERE products.idproduct=pcReviewSpecials.pcRS_IDProduct order by products.description asc"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
	
if rs.eof then
	DataEmpty=1
else
	DataEmpty=0
	pcArray=rs.getRows()
	intCount=ubound(pcArray,2)
end if
set rs=nothing

call closedb()	
%>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>


<form method="POST" action="prv_SpecialPrd.asp?action=update" name="checkboxform" class="pcForms">
	<table class="pcCPcontent">
		<tr>
            <td colspan="4">All the <a href="prv_FieldManager.asp">product review fields</a> that you have defined in your store are used for all products <u>except for the products listed below</u>, for which you can define a smaller set of fields. If you wish not to allow customers to post reviews for selected products, you can do so by using the <a href="prv_PrdExc.asp">product exclusion</a> feature.</td>
		</tr>
		<tr>
            <td colspan="4" class="pcCPspacer"></td>
		</tr>
		<tr>
            <th>ID</td>
            <th nowrap>Product Name</th>
            <th colspan="2">Action</th>
		</tr>
		<tr>
            <td colspan="4" class="pcCPspacer"></td>
		</tr>
		<% If DataEmpty=1 Then %>
			<tr> 
                <td colspan="4">No Products Found</td>
			</tr>
		<% Else 
			Dim Count
			Count=0
			For k=0 to intCount
				Count=Count+1
				pcv_ID=pcArray(0,k)
				pcv_Name=pcArray(1,k)
				%>
													
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
                    <td width="5%"><%=pcv_ID%></td>
                    <td width="80%"><a href="FindProductType.asp?id=<%=pcv_ID%>" target="_blank"><%=pcv_Name%></a></td>
                    <td align="right" width="10%"><a href="prv_EditSpecial.asp?IDProduct=<%=pcv_ID%>">View/Edit</a></td>
                    <td width="5%">
                    <input type="checkbox" size="3" name="C<%=count%>" value="1">
                    <input type="hidden" name="ID<%=count%>" value="<%=pcv_ID%>">
                    </td>
				</tr>
			<% 
			Next
		End If %>
		<%if count>"0" then%>
            <tr>
                <td colspan="4" class="pcCPspacer"></td>
            </tr>
			<tr>
			<td colspan="4" class="cpLinksList" align="right">
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
				<input type=hidden name=count value=<%=count%>>
				<a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a>
			</td>
			</tr>
		<%end if%>
        <tr>
            <td colspan="4" class="pcCPspacer"></td>
        </tr>
		<tr>
		<td colspan="4">
			<input type="button" value="Add New Product" onClick="location.href='prv_AddSpecialPrd.asp'" class="submit2">
			<%if count>0 then%>
            &nbsp;<input type="submit" value="Remove Selected from List" name="submit" onclick="return(confirm('Custom settings will be removed from the selected products. All product review fields will be shown for those products. Are you sure you want to complete this action?'));">  
            <%end if%>
		</td>
		</tr>
		<tr>
		<td colspan="4">
			<input type="button" value="Manage Fields" onClick="location.href='prv_FieldManager.asp'">&nbsp;
			<input type="button" value="Product Exclusions" onClick="location.href='prv_PrdExc.asp'">
		</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->
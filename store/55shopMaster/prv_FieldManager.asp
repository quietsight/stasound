<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="Product Reviews - Manage Fields"
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

call opendb()

IF request("action")="update" then
	Count=request("count")
	if (Count>"0") and (IsNumeric(Count)) then
		For k=1 to Count
			if request("C" & k)="1" then
				pcv_ID=request("ID" & k)
				pcv_Name=request("name" & k)
				pcv_Type=request("type" & k)
				if pcv_Type="" then
					pcv_Type="0"
				end if
				pcv_Active=request("active" & k)
				if pcv_Active="" then
					pcv_Active="0"
				end if
				pcv_Required=request("required" & k)
				if pcv_Required="" then
					pcv_Required="0"
				end if
				pcv_Order=request("order" & k)
				if pcv_Order="" then
					pcv_Order="0"
				end if
				if request("submit1")<>"" then
					query="UPDATE pcRevFields SET pcRF_Name='" & pcv_Name & "',pcRF_Type=" & pcv_Type & ",pcRF_Active=" & pcv_Active & ",pcRF_Required=" & pcv_Required & ",pcRF_Order=" & pcv_Order & " WHERE pcRF_IDField=" & pcv_ID
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=connTemp.execute(query)
					set rs=nothing
				end if
				if (request("submit2")<>"") and (pcv_ID<>"1") and (pcv_ID<>"2") then
					query="SELECT pcRD_IDReview FROM pcReviewsData WHERE pcRD_IDField=" & pcv_ID
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=connTemp.execute(query)
					if rs.eof then	
						query="DELETE FROM pcRevFields WHERE pcRF_IDField=" & pcv_ID
						set rs=connTemp.execute(query)
						query="DELETE FROM pcRevLists WHERE pcRL_IDField=" & pcv_ID
						set rs=connTemp.execute(query)
						set rs=nothing
					else
						msg="One or more of the fields cannot be deleted because they have been used in product reviews. Instead, you can make them 'Inactive'."
					end if
				end if
			end if
		Next
	end if
END IF
	
query="SELECT pcRF_IDField,pcRF_Name,pcRF_Type,pcRF_Active,pcRF_Required,pcRF_Order FROM pcRevFields ORDER BY pcRF_Order asc,pcRF_IDField asc"
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

<form method="POST" action="prv_FieldManager.asp?action=update" name="checkboxform" class="pcForms">
	<table class="pcCPcontent">
		<tr>
			<td colspan="8">
				<p>These are the fields that are shown to your customers when they choose to write a review for a products. By default, <u>all of these fields</u> are shown whenever a review is written for <u>any product</u> in your store catalog, except for:</p>
				<ul>
                    <li>Products that have been <strong>excluded</strong> from the Product Reviews feature. <a href="prv_PrdExc.asp">View product exclusion list</a>.</li>
                    <li>Products for which <strong>specific settings</strong> have been configured. <a href="prv_SpecialPrd.asp">View product-specific settings</a>.</li>
				</ul>
			</td>
		</tr>
		<tr> 
            <td colspan="8" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th>ID</th>
			<th width="30%" align="center" nowrap>Field Name</th>
            <th align="center" nowrap>Type</th> 
            <th align="center" nowrap>Active</th>
            <th align="center" nowrap>Required</th>
            <th align="center" nowrap>Order</th>
            <th align="center" nowrap>Select</th>
            <th nowrap>&nbsp;</th>
		</tr>
		<% If DataEmpty=1 Then %>
            <tr> 
                <td colspan="8" class="pcCPspacer"></td>
            </tr>
			<tr> 
                <td colspan="8">No Fields Found</td>
			</tr>
		<% Else 
			Dim Count
			Count=0
			For k=0 to intCount
				Count=Count+1
				pcv_ID=pcArray(0,k)
				pcv_Name=pcArray(1,k)
				pcv_Type=pcArray(2,k)
				pcv_Active=pcArray(3,k)
				pcv_Required=pcArray(4,k)
				pcv_Order=pcArray(5,k)%>
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist" style="height: 40px; vertical-align: middle;"> 
                    <td align="right" nowrap><%=pcv_ID%></td>
                    <td align="left">
                    <%if (pcv_ID="1") or (pcv_ID="2") then%>
                        <%=pcv_Name%>
                        <input type=hidden name="name<%=count%>" value="<%=pcv_Name%>">
                    <%else%>
                        <input type=text size="40" name="name<%=count%>" value="<%=pcv_Name%>">
                    <%end if%>
                    </td>
                    <td align="center">
                    <%if (pcv_ID="1") or (pcv_ID="2") then%>
                        1-row text field
                        <input type="hidden" name="type<%=count%>" value="0">
                    <%else%>
                        <select name="type<%=count%>">
                        <option value="0" <%if pcv_Type="0" then%>selected<%end if%>>1-row text field</option>
                        <option value="1" <%if pcv_Type="1" then%>selected<%end if%>>Text area</option>
                        <option value="2" <%if pcv_Type="2" then%>selected<%end if%>>Drop-down list</option>
                        <option value="3" <%if pcv_Type="3" then%>selected<%end if%>>'Feeling' Rating</option>
                        <option value="4" <%if pcv_Type="4" then%>selected<%end if%>>'Mark' Rating</option>
                        </select>
                    <%end if%>
                    </td>
                    <td align="center">
						<%if (pcv_ID="1") or (pcv_ID="2") then%>
                            Active
                            <input type="hidden" name="active<%=count%>" value="1">
                        <%else%>
                            <input type="checkbox" name="active<%=count%>" value="1" <%if pcv_Active="1" then%>checked<%end if%>>
                        <%end if%>
                    </td>
                    <td align="center">
                    	<input type="checkbox" name="required<%=count%>" value="1" <%if pcv_Required="1" then%>checked<%end if%>>
                    </td>
                    <td align="center">
                    <%if (pcv_ID<>"1") and (pcv_ID<>"2") then%>
                        <input type="text" size="3" name="order<%=count%>" value="<%=pcv_Order%>">
                    <%else%>
                        <input type="hidden" name="order<%=count%>" value="0">
                    <%end if%>
                    </td>
                    <td align="center">
                    	<% if not (pcv_ID=1 or pcv_ID=2) then %>
                    	<input type="checkbox" name="C<%=count%>" value="1">
                        <% end if %>
                    	<input type="hidden" name="ID<%=count%>" value="<%=pcv_ID%>">
                    </td>
                    <td align="center" nowrap>
                    <%if pcv_Type="2" then%>
                    <a href="prv_EditList.asp?IDField=<%=pcv_ID%>">View/Edit Values</a>
                    <%end if%>
                    &nbsp;
                    </td>
				</tr>
			<% 
			Next
		End If 

		if count>"0" then
		%>
            <tr> 
                <td colspan="8" class="pcCPspacer"></td>
            </tr>
			<tr>
            	<td colspan="6"></td>
                <td colspan="2" class="cpLinksList">
					<script language="JavaScript">
                    <!--
					function checkAll() {
					for (var j = 3; j <= <%=count%>; j++) {
					box = eval("document.checkboxform.C" + j); 
					if (box.checked == false) box.checked = true;
					   }
					}
					
					function uncheckAll() {
					for (var j = 3; j <= <%=count%>; j++) {
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
		<%
		end if
		%>
		<tr>
            <td colspan="8">
                <input type="button" value="Add New" onClick="location.href='prv_AddField.asp'">
                <%if count>0 then%>
                <input type="submit" value="Update" name="submit1" class="submit2">&nbsp;
                <input type="submit" value="Delete selected" name="submit2" onclick="return(confirm('You are about to remove the selected fields. Are you sure you want to complete this action?'));">  
                <%end if%>
                <hr />
                <input type="button" value=" Back " onClick="javascript:history.back()">
            </td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="Product Reviews: Apply Special Settings" 
pageIcon="pcv4_icon_reviews.png"
Section="reviews" 
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
	
pcv_PrdList=request("PrdList")
	
IF request("action")="update" then
	Count=request("count")
	if (Count>"0") and (validNum(Count)) then
		pcv_strFieldList=""
		pcv_strFieldOrder=""
		pcv_strRequired=""
		For k=1 to Count
			if (request("active" & k)="1") or (request("ID" & k)="1") or (request("ID" & k)="2") then
				pcv_ID=request("ID" & k)
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
				if pcv_Active="1" then
					pcv_strFieldList=pcv_strFieldList & pcv_ID & ","
					pcv_strFieldOrder=pcv_strFieldOrder & pcv_Order & ","
					pcv_strRequired=pcv_strRequired & pcv_Required & ","
				end if
			end if
		Next

		if (pcv_strFieldList<>"") and (pcv_PrdList<>"") then
			pcArray=split(pcv_PrdList,",")
			call openDb()
			Dim pcv_PrdListTotal
			pcv_PrdListTotal=ubound(pcArray)-1
			For k=0 to pcv_PrdListTotal
				if pcArray(k)<>"" then
					query="INSERT INTO pcReviewSpecials (pcRS_IDProduct,pcRS_FieldList,pcRS_FieldOrder,pcRS_Required) VALUES (" & pcArray(k) & ",'" & pcv_strFieldList & "','" & pcv_strFieldOrder & "','" & pcv_strRequired & "')"
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=connTemp.execute(query)
					set rs=nothing
				end if
			Next
			call closedb()
		end if
	end if
	response.redirect "prv_SpecialPrd.asp?s=1&msg=" & Server.URLEncode("Product Reviews fields successfully applied to the selected products.")
END IF
	
call opendb()

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
<form method="POST" action="prv_CustomizePrd.asp?action=update" name="checkboxform" class="pcForms">
	<input type="hidden" name="PrdList" value="<%=request("PrdList")%>">
	<table class="pcCPcontent">
		<tr>
			<th>ID</th>
			<th width="50%" nowrap>Field Name</th>
			<th>Type</th> 
			<th>Active</th>
			<th>Required</th>
			<th>Order</th>
		</tr>
		<tr> 
			<td colspan="6" class="pcCPspacer"></td>
		</tr>
		<% If DataEmpty=1 Then %>
			<tr> 
				<td colspan="6">No Fields Found.</td>
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
												
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
					<td><%=pcv_ID%><input type="hidden" name="ID<%=count%>" value="<%=pcv_ID%>"></td>
					<td><a href="FindProductType.asp?id=<%=pcv_ID%>" target="_blank"><%=pcv_Name%></a></td>
					<td>
						<%if pcv_Type="0" then%>1-row text field<%end if%>
                        <%if pcv_Type="1" then%>Text area<%end if%>
                        <%if pcv_Type="2" then%>Drop-down list<%end if%>
                        <%if pcv_Type="3" then%>'Feeling' Rating<%end if%>
                        <%if pcv_Type="4" then%>'Mark' Rating<%end if%>
					</td>
					<td>
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
				</tr>
			<% Next
		End If %>        
        <tr> 
            <td colspan="6" class="pcCPspacer"></td>
        </tr>
		<tr>
			<td colspan="6">
			<%if count>0 then%>
				<input type="submit" value="Assign to Selected Products" name="submit" class="submit2">
                <input type="hidden" name="count" value="<%=count%>">
			<%end if%>
                &nbsp;
                <input type="button" value=" Back " onClick="javascript:history.back()">
			</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->
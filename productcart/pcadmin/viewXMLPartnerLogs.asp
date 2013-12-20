<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="XML Tools - Transactions Log" %>
<% section="layout" %>
<%PmAdmin=19%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%Dim rs, connTemp, query, i, tmpBackup, Count, fso, afi

pidPartner=trim(request("idPartner"))

If Not IsNumeric(pidPartner) then
	response.redirect "techErr.asp?error="&Server.URLEncode("An error occurred when submitting your query.")
End If

if request("iPageCurrent")="" then
    iPageCurrent=1 
else
    iPageCurrent=Request("iPageCurrent")
end If

If request("action")="upd" and request("delTrans")<>"" then
	Count=request("Count")
	if Count<>"" AND Clng(Count)>0 then
		For i=1 to Count
			TransID=request("TransID" & i)
			eCheck=request("C" & i)
			if eCheck="1" then
				call opendb()
					query="select pcXL_BackupFile FROM pcXMLLogs WHERE pcXP_ID=" & pidPartner & " AND pcXL_ID=" & TransID & ";"
					set rstemp=connTemp.execute(query)
					tmpBackup=""
					if not rstemp.eof then
						tmpBackup=rstemp("pcXL_BackupFile")
					end if
					set rstemp=nothing
					
					query="DELETE FROM pcXMLLogs WHERE pcXP_ID=" & pidPartner & " AND pcXL_ID=" & TransID & ";"
					set rstemp=connTemp.execute(query)
					set rstemp=nothing

				call closedb()
				
				on error resume next
				if tmpBackup<>"" then
					Set fso=Server.CreateObject("Scripting.FileSystemObject")
					Set afi = fso.GetFile(server.MapPath("..\xml") & "\logs\" & tmpBackup)
					afi.Delete
					Set afi=nothing
					Set fso=nothing
				end if
				msg="6"
			end if
		Next
	end if
end if

call openDB()
%>
<!--#include file="AdminHeader.asp"-->
<form name="Form1" method="post" action="viewXMLPartnerLogs.asp?action=upd">
<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
		<%query="SELECT pcXP_PartnerID,pcXP_Name FROM pcXMLPartners WHERE pcXP_ID=" & pidPartner & ";"
		set rs=connTemp.execute(query)
		if not rs.eof then
			pa_UserID=rs("pcXP_PartnerID")
			pa_Name=rs("pcXP_Name")%>
		XML Partner: <%=pa_UserID%><%if pa_Name<>"" then%>&nbsp(<%=pa_Name%>)<%end if%>
		&nbsp;-&nbsp;<a href="modXMLPartner.asp?idPartner=<%=pidPartner%>">View/Edit</a>
		<%end if
		set rs=nothing%>
		<%IF pidPartner>"0" THEN
		query="SELECT pcXL_ID FROM pcXMLLogs WHERE pcXP_ID=0;"
		set rs=connTemp.execute(query)
		if not rs.eof then%>
		&nbsp;|&nbsp;<a href="viewXMLPartnerLogs.asp?idPartner=0">Unauthorized Transactions Log</a>
		<%end if
		set rs=nothing
		ELSE%>
			XML Transactions of Unknown Partners
		<%END IF%>
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<% 
	If Request.QueryString("msg")<>"" or msg<>"" Then
	If Request.QueryString("msg")<>"" then
		msg=request("msg")
	End if%>
	<tr> 
		<td align="center">
        <% if msg<5 then %>
			<div class="pcCPmessage">
				<%Select Case msg
				Case "1": response.write "Cannot find the XML Request"
				Case "2": response.write "Cannot undo the XML Request because request type is invalid."
				Case "3": response.write "The XML Request was already undone"
				Case "4": response.write "Cannot undo the XML Request because the backup file was not found"
				End Select%>
			</div>
       <% else %>
			<div class="pcCPmessageSuccess">
				<%Select Case msg
				Case "5": response.write "The XML Request has been undone successfully!"
				Case "6": response.write "Selected XML Requests were removed successfully!"
				End Select%>
			</div>
       <% end if %>
		</td>
	</tr>
	<% End If %>
	
	<tr>
		<td align="center">
			<table class="pcCPcontent">
				<tr>
					<th nowrap>&nbsp;</th>
					<th nowrap>Date</th>
					<th nowrap>Transaction Key</th>
					<th nowrap>Transaction Type</th>
					<th nowrap>Status</th>
					<th nowrap>Description</th>
					<th nowrap>&nbsp;</th>
				</tr>
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
				<%
				query="select pcXL_ID,pcXP_ID,pcXL_RequestKey,pcXL_RequestType,pcXL_UpdatedID,pcXL_BackupFile,pcXL_Undo,pcXL_ResultCount,pcXL_Date,pcXL_LastID,pcXL_UndoID,pcXL_Status,pcXL_RequestXML,pcXL_ResponseXML FROM pcXMLLogs WHERE pcXP_ID=" & pidPartner & " ORDER BY pcXL_ID DESC;"
			
				Set rstemp=Server.CreateObject("ADODB.Recordset")	
				rstemp.CacheSize=25
				rstemp.PageSize=25
				rstemp.Open query, connTemp, adOpenStatic, adLockReadOnly
				
				Count=0
				If rstemp.eof Then
				%>             
					<tr> 
						<td colspan="2"><p>No XML Transactions Found.</p></td>
					</tr>            
				<%
				Else
				
					' get the max number of pages
					Dim iPageCount
					iPageCount=rstemp.PageCount
					iPageSize=rstemp.PageSize
					If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
					If iPageCurrent < 1 Then iPageCurrent=1
					rstemp.AbsolutePage=iPageCurrent
				
					pcArr=rstemp.getRows()
					set rstemp=nothing
					intCount=ubound(pcArr,2)+1					
					i=0
					do while (i < intCount) and (i < iPageSize)	
						Count=Count+1
						xt_ID=pcArr(0,i)
						pa_ID=pcArr(1,i)
						xt_Key=pcArr(2,i)
						xt_Type=pcArr(3,i)
						xt_UpdID=pcArr(4,i)
						xt_BFile=pcArr(5,i)
						xt_Undo=pcArr(6,i)
						xt_RCount=pcArr(7,i)
						xt_Date=pcArr(8,i)
						xt_LastID=pcArr(9,i)
						xt_UID=pcArr(10,i)
						xt_Status=pcArr(11,i)
						if xt_Status<>"" then
						else
							xt_Status=0
						end if
						xt_RequestXML=trim(pcArr(12,i))
						xt_ResponseXML=trim(pcArr(13,i))
						
						if xt_UID<>0 then
							query="SELECT pcXL_RequestKey FROM pcXMLLogs WHERE pcXL_ID=" & xt_UID & ";"
							set rs=connTemp.execute(query)
							if not rs.eof then
								xt_URequest=rs("pcXL_RequestKey")
							end if
							set rs=nothing
						end if
						
						%>
						<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
							<td nowrap>
								<input type="hidden" name="TransID<%=Count%>" value="<%=xt_ID%>">
								<input type="checkbox" name="C<%=Count%>" value="1">
							</td>
							<td nowrap valign="top"><%=xt_Date%></td>
							<td nowrap valign="top"><%=xt_Key%></td>
							<td nowrap valign="top">
								<%Select Case clng(xt_Type)
									Case 0: response.write "<b>SearchProducts</b>"
									Case 1: response.write "<b>SearchCustomers</b>"
									Case 2: response.write "<b>SearchCustomers</b>"
									Case 3: response.write "<b>GetProductDetails</b>"
									Case 4: response.write "<b>GetCustomerDetails</b>"
									Case 5: response.write "<b>GetOrderDetails</b>"
									Case 6: response.write "<b>NewProducts</b>"
									Case 7: response.write "<b>NewCustomers</b>"
									Case 8: response.write "<b>NewOrders</b>"
									Case 9: response.write "<b>AddProduct</b>"
									Case 10: response.write "<b>AddCustomer</b>"
									Case 11: response.write "<b>UpdateProduct</b>"
									Case 12: response.write "<b>UpdateCustomer</b>"
									Case 13: response.write "<b>UndoRequest</b>"
									Case 13: response.write "<b>UndoRequest</b>"
									Case 14: response.write "<b>SetExportFlagRequest</b>"
									Case Else: response.write "<b>Unknown method</b>"
								End Select
								%>
							</td>
							<td nowrap valign="top">
								<%Select Case clng(xt_Status)
									Case 0: response.write "<b>Errors</b>"
									Case 1: response.write "Successful"
									Case 2: response.write "Successful<br><i>with some errors</i>"
									Case Else: response.write "<b>Unknown</b>"
								End Select%>
							</td>
							<td width="50%" valign="top">
								<%Select Case clng(xt_Type)
									Case 0:%>Search Products. Total Results: <%=xt_RCount%> product(s)
									<%Case 1:%>Search Customers. Total Results: <%=xt_RCount%> customer(s)
									<%Case 2:%>Search Orders. Total Results: <%=xt_RCount%> order(s)
									<%Case 3:%>Request details of the Product<%if xt_UpdID>0 then%>&nbsp;ID#: <%=xt_UpdID%><%end if%>
									<%Case 4:%>Request details of the Customer<%if xt_UpdID>0 then%>&nbsp;ID#: <%=xt_UpdID%><%end if%>
									<%Case 5:%>Request details of the Order<%if xt_UpdID>0 then%>&nbsp;ID#: <%=xt_UpdID%><%end if%>
									<%Case 6:%>Request new products list. Total Results: <%=xt_RCount%> product(s)
									<%Case 7:%>Request new customers list. Total Results: <%=xt_RCount%> customer(s)
									<%Case 8:%>Request new orders list. Total Results: <%=xt_RCount%> order(s)
									<%Case 9:%>Add new product<%if xt_UpdID>0 then%>&nbsp;ID#: <%=xt_UpdID%><%if xt_BFile<>"" then%><br><i>Backup File: <%=xt_BFile%></i><%end if%><%end if%>
									<%Case 10:%>Add new customer<%if xt_UpdID>0 then%>&nbsp;ID#: <%=xt_UpdID%><%if xt_BFile<>"" then%><br><i>Backup File: <%=xt_BFile%></i><%end if%><%end if%>
									<%Case 11:%>Update the product<%if xt_UpdID>0 then%>&nbsp;ID#: <%=xt_UpdID%><%if xt_BFile<>"" then%><br><i>Backup File: <%=xt_BFile%></i><%end if%><%end if%>
									<%Case 12:%>Update the customer<%if xt_UpdID>0 then%>&nbsp;ID#: <%=xt_UpdID%><%if xt_BFile<>"" then%><br><i>Backup File: <%=xt_BFile%></i><%end if%><%end if%>
									<%Case 13:%>Undo the Transaction Key<%if xt_URequest>0 then%>: <%=xt_URequest%><%end if%>
									<%Case 14:%>Updated Exported Flag for <%if xt_UpdID>0 then%>&nbsp;ID#: <%=xt_UpdID%><%end if%>
								<%End Select%>
							</td>
							<td nowrap valign="top">
								<%if xt_RequestXML<>"" or xt_ResponseXML<>"" then%>
								<a href="viewXMLLogDetails.asp?idPartner=<%=pidPartner%>&idxml=<%=xt_ID%>">Details</a>
								<%end if%>
								<%if clng(xt_Undo)=1 then%><%if xt_RequestXML<>"" or xt_ResponseXML<>"" then%>&nbsp;|&nbsp;<%end if%>Undone
								<%else
								if clng(xt_Type)>=9 and clng(xt_Type)<=12 then%>
									<%if xt_RequestXML<>"" or xt_ResponseXML<>"" then%>&nbsp;|&nbsp;<%end if%><a href="javascript:if (confirm('You are about to undo the XML Transaction: <%=xt_Key%>. Are you sure you want to complete this action?')) location='undoXMLRequest.asp?action=del&idPartner=<%=pidPartner%>&idxml=<%=xt_ID%>'">Undo</a>
								<%end if
								end if%>
							</td>
						</tr>              
						<%
						i=i+1
					loop
				End If
				%>
			</table>
		</td>
	</tr>
	<tr>
		<td><input type="hidden" name="Count" value="<%=Count%>">
		<input type="hidden" name="idPartner" value="<%=pidPartner%>">
		<%if Count>0 then%>
			<a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a>
			<script language="JavaScript">
				function checkAll()
				{
					for (var j = 1; j <= <%=Count%>; j++)
					{
						box = eval("document.Form1.C" + j); 
						if (box.checked == false) box.checked = true;
					}
				}

				function uncheckAll()
				{
					for (var j = 1; j <= <%=Count%>; j++)
					{
						box = eval("document.Form1.C" + j); 
						if (box.checked == true) box.checked = false;
					}
				}
			</script>
		<%end if%>
		</td>
	</tr>
	<tr>
		<td> 
		<% If iPageCount>1 Then %>
			<%response.Write("Currently viewing page "& iPageCurrent & " of "& iPageCount & "<br>")%>
			<p class="pcPageNav">
				<%if iPageCurrent > 1 then %>
					<a href="viewXMLPartnerLogs.asp?idPartner=<%=request("idPartner")%>&iPageCurrent=<%=iPageCurrent-1%>"><img src="../pc/images/prev.gif" width="10" height="10" border="0"></a>
				<% end If
				For I=1 To iPageCount
					If Cint(I)=Cint(iPageCurrent) Then %>
						<%=I%> 
					<% Else %>
						<a href="viewXMLPartnerLogs.asp?idPartner=<%=request("idPartner")%>&iPageCurrent=<%=I%>"><%=I%></a> 
					<% End If %>
				<% Next %>
				<% if CInt(iPageCurrent) < CInt(iPageCount) then %>
					<a href="viewXMLPartnerLogs.asp?idPartner=<%=request("idPartner")%>&iPageCurrent=<%=iPageCurrent+1%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a>
				<% end If %>
			</p>
		<% End If %>
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td class="pcCPspacer">
			<%if Count>0 then%>
				<input type="submit" name="delTrans" value="Delete Selected" class="submit2" onclick="if (confirm('You are about to remove XML Transaction logs from your database. Are you sure you want to complete this action?')) {return(true)} else {return(false)}">&nbsp;
			<%end if%>
			<input type="button" name="Main" value="Manage Partners" onClick="location.href='AdminManageXMLPartner.asp'" class="ibtnGrey">&nbsp;
					<input type="button" name="Back" value="XML Tools Manager" onclick="location='XMLToolsManager.asp';" class="ibtnGrey"></td>
	</tr>
</table>
</form>
<%call closedb()%><!--#include file="AdminFooter.asp"-->
<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Shipping Center for UPS" %>
<% Section="mngAcc" %>
<%PmAdmin=9%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/UPSconstants.asp"-->
<!--#include file="../includes/pcUPSClass.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcShipTestModes.asp" -->
<% ON ERROR RESUME NEXT
Const iPageSize=5

Dim iPageCurrent, conntemp, rs, varFlagIncomplete, uery, strORD, pcv_intOrderID



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'// SET PAGE NAMES
pcPageName = "UPS_ManageShipmentsResults.asp"
ErrPageName = "UPS_ManageShipmentsTrack.asp"

'// OPEN DATABASE
call openDb()

'// SET THE UPS OBJECT
set objUPSClass = New pcUPSClass

'// GET PAGE NUMBER
if request.querystring("iPageCurrent")="" then
	iPageCurrent=1 
else
	iPageCurrent=Request.QueryString("iPageCurrent")
end if

'// SORT ORDER
strORD=request("order")
if strORD="" then
	strORD="pcPackageInfo_ShippedDate DESC, idOrder"
End If
strSort=request("sort")
if strSort="" Then
	strSort="DESC"
End If 

'// GET ORDER ID
pcv_strOrderID=Request("id")
pcv_strSessionOrderID=Session("pcAdminOrderID")
if pcv_strSessionOrderID="" OR len(pcv_strOrderID)>0 then
	pcv_intOrderID=pcv_strOrderID
	Session("pcAdminOrderID")=pcv_intOrderID
else
	pcv_intOrderID=pcv_strSessionOrderID
end if
'response.write Session("pcAdminOrderID")
'response.end

' GET THE PACKAGES
' >>> Tables: pcPackageInfo
query = 		"SELECT pcPackageInfo.* "
query = query & "FROM pcPackageInfo "
query = query & "WHERE pcPackageInfo.idOrder=" & pcv_intOrderID &" "	

' >>> Conditions
If Request.QueryString("TypeSearch")="idOrder" Then
	tempqryStr=Request.QueryString("advquery")
	if tempqryStr="" then
		tempqryStr=" ORDER BY "& strORD &" "& strSort
	else
		tempqryStr=(int(tempqryStr) - scpre)
		query=query & " WHERE idOrder LIKE '%" & _
		tempqryStr & "%' ORDER BY "& strORD &" "& strSort
	end if
End If	
'If Request.QueryString("TypeSearch")="orderstatus" Then
'	query=query & " WHERE orders.idCustomer=customers.idCustomer AND orderstatus LIKE '%" & _
'	Request.QueryString("advquery") & "%' ORDER BY "& strORD &" "& strSort
'End If

set rs=Server.CreateObject("ADODB.Recordset") 

rs.CursorLocation=adUseClient
rs.CacheSize=iPageSize
rs.PageSize=iPageSize
rs.Open query, conntemp

if err.number <> 0 then
	call rs.Close
	set rs=nothing
	response.redirect "techErr.asp?error="& Server.Urlencode("Error in listorders: "&Err.Description) 
end If

rs.MoveFirst

'// GET MAX PAGES
Dim iPageCount
iPageCount=rs.PageCount
If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
If iPageCurrent < 1 Then iPageCurrent=1

'// SET ABSOLUTE PAGE
rs.AbsolutePage=iPageCurrent

'// DISPLAY ERROR MSG
msg=request.querystring("msg")

if msg<>"" then 
	%>
	<div class="pcCPmessage">
		<img src="images/pcadmin_note.gif" alt="Alert" width="20" height="20"> <%=msg%>
	</div>
	<% 
end if 


'// DISPLAY HEADER
if rs.eof then 
	presults="0"
else 
%>
	<table class="pcCPcontent">
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="3">UPS OnLine&reg; Tools - Manage Shipments for Order Number <%=(scpre+int(pcv_intOrderID))%><a name="top"></a></th>
		</tr>
		<% if UPS_TESTMODE="1" then %>
            <tr>
                <td colspan="2" class="pcSpacer"></td>
            </tr>
            <tr>
                <td colspan="2">
                    <div class="pcCPmessage">
                        UPS Shipping Wizard is currently running in Test Mode 
                    </div>
                </td>
            </tr>
        <% end if %>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td width="25%" align="left" valign="bottom"> 
			<% 
				'// Showing total number of pages found and the current page number
				Response.Write "Displaying Page <b>" & iPageCurrent & "</b> of <b>" & iPageCount & "</b><br>"
				Response.Write "Total Shipments Found: <b>" & rs.RecordCount & "</b>" 
				%></td>
			<td width="50%" align="center">&nbsp;</td>
		    <td width="25%" align="right" valign="bottom">&nbsp;</td>
		</tr>
	</table>
<% end if %>

<table class="pcCPcontent">

	<form name="checkboxform" action="UPS_ManageShipmentsTrack.asp?id=<%=pcv_intOrderID%>&action=batch" method="post" class="pcForms">
		<tr> 
			<th nowrap>Shipped</th>
			<th nowrap>UPS Tracking Number </th>
			<th nowrap>Contents Description</th>
			<th nowrap>Package Details</th>
			<th nowrap>Options</th>
			<th nowrap>Returns</th>
			<th nowrap>Select</th>
		</tr>
		<% Dim mcount
		mcount=0
		If rs.EOF Then %>
			<tr>
			<td colspan="11">
				<div class="pcCPmessage"><img src="images/pcadmin_note.gif" width="20" height="20"> No Results Found</div>			</td>
			</tr>
		<% Else
			' Showing relevant records
			Dim strCol
			strCol="#E1E1E1" 
			Dim rcount, i, x
			
			For i=1 To rs.PageSize
				
				rcount=i
				If currentPage > 1 Then
					For x=1 To (currentPage - 1)
						rcount=10 + rcount
					Next
				End If
                          
				If Not rs.EOF Then 
					If strCol <> "#FFFFFF" Then
						strCol="#FFFFFF"
					Else 
						strCol="#E1E1E1"
					End If
					
					pcv_intPackageInfo_ID=rs("pcPackageInfo_ID")
					pcv_intOrder=rs("idOrder")
					pidPackageNumber=rs("pcPackageInfo_PackageNumber")
					pcv_strTrackingNumber=rs("pcPackageInfo_TrackingNumber")
					pidPackageWeight=rs("pcPackageInfo_PackageWeight")
					pidPackageShippedDate=rs("pcPackageInfo_ShippedDate")
					pidPackageUPSCODAmount=rs("pcPackageInfo_UPSCODAmount")
					pcv_strShipMethod = rs("pcPackageInfo_ShipMethod")
					strLabelImageFormat = rs("pcPackageInfo_UPSLabelFormat")
					mcount=mcount+1 
					%>
												
					<tr> 
						<% if pidPackageShippedDate<>"" then
							pidPackageShippedDate=ShowDateFrmt(pidPackageShippedDate)
						end if %>
						<td bgcolor="<%= strCol %>"><%=pidPackageShippedDate%></td>
						<td bgcolor="<%= strCol %>">
							<a href="UPS_ManageShipmentsTrack.asp?id=<%=pcv_intOrder%>&PackageInfo_ID=<%=pcv_intPackageInfo_ID%>"><%=pcv_strTrackingNumber%></a>						</td>
						<td bgcolor="<%= strCol %>">
						<%
						' GET THE PACKAGE CONTENTS
						' >>> Tables: products, ProductsOrdered
						query = 		"SELECT ProductsOrdered.pcPackageInfo_ID , products.description, products.idProduct  "
						query = query & "FROM ProductsOrdered "
						query = query & "INNER JOIN products "
						query = query & "ON ProductsOrdered.idProduct = products.idProduct "
						query = query & "WHERE ProductsOrdered.pcPackageInfo_ID=" & pcv_intPackageInfo_ID &" "
												
						set rs2=server.CreateObject("ADODB.RecordSet")
						set rs2=conntemp.execute(query)		
						
						if err.number<>0 then
							'// handle admin error
						end if
						
						if NOT rs2.eof then
							Do until rs2.eof	
								pcv_strProductDescription = rs2("description")
								%>
								<%=pcv_strProductDescription%><br />
								<%
							rs2.movenext
							Loop
						end if						
						%>						</td>
						<td bgcolor="<%= strCol %>">
							
						
							<table width="100%" border="0" cellspacing="0" cellpadding="0">
							  <tr>
								<td align="right">Weight:</td>
								<td align="left"><b><%=pidPackageWeight%> lbs.</b></td>
							  </tr>
						  </table>						</td>
						<td bgcolor="<%= strCol %>">
							<% if strLabelImageFormat="EPL" then
								'check if label exists
								'on error resume next
								Set fs=server.CreateObject("Scripting.FileSystemObject")				
								findit=Server.MapPath("UPSLabels/Label"&pcv_strTrackingNumber&".txt")	
								if (fs.FileExists(findit))=true then %>	
								<a href="UPSLabels/Label<%=pcv_strTrackingNumber%>.txt" target="_blank">View/ Print Label</a><br />
								<% else %>
								&nbsp;
								<% end if %>
							<% else
								'check if label exists
								'on error resume next
								Set fs=server.CreateObject("Scripting.FileSystemObject")				
								findit=Server.MapPath("UPSLabels/Label"&pcv_strTrackingNumber&".html")	
								if (fs.FileExists(findit))=true then %>	
									<a href="UPSLabels/Label<%=pcv_strTrackingNumber%>.html" target="_blank">View/ Print Label</a><br />
								<% else %>
									&nbsp;
								<% end if %>
							<% end if %>
						</td>
						<td bgcolor="<%= strCol %>">
							<a href="UPS_ManageShipmentsCancel.asp?id=<%=pcv_intOrder%>&PackageInfo_ID=<%=pcv_intPackageInfo_ID%>">Cancel Shipment</a><br />
						<td bgcolor="<%= strCol %>"><input type=checkbox name="check<%=mcount%>" value="<%=pcv_intPackageInfo_ID%>"></td>
					</tr>
					<% rs.MoveNext
				End If
			Next%>
			<input type=hidden name="count" value="<%=mcount%>">								
		<% End If %>
		<tr> 
			<td colspan="11">
				<%if mcount>0 then%>
					<a href="javascript:checkAll();"><b>Check All</b></a><b>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a></b><br>
					<br><input type=submit name="submit" value="Track All Selected Packages" class="ibtnGrey">
					<script language="JavaScript">
					<!--
					function checkAll() {
					for (var j = 1; j <= <%=mcount%>; j++) {
					box = eval("document.checkboxform.check" + j); 
					if (box.checked == false) box.checked = true;
						 }
					}
						
					function uncheckAll() {
					for (var j = 1; j <= <%=mcount%>; j++) {
					box = eval("document.checkboxform.check" + j); 
					if (box.checked == true) box.checked = false;
						 }
					}
					
					//-->
					</script>
				<%end if%>			</td>
		</tr>
	</form>
              
	<tr>
		<td colspan="11"> 
			<% if pResults<>"0" Then %>
				<table width="100%" border="0" cellspacing="0" cellpadding="4">
					<tr> 
						<td> 
							<form method="post" action="" name="" class="pcForms">
							<b> 
							<% Response.Write("<font size=2 face=arial>Page "& iPageCurrent & " of "& iPageCount & "</font><P>")%>
							<% 'Display Next / Prev buttons
							if iPageCurrent > 1 then
								'We are not at the beginning, show the prev button %>
								<a href="UPS_ManageShipmentsResults.asp?id=<%=pcv_intOrder%>&TypeSearch=<%=request.querystring("TypeSearch")%>&advquery=<%=request.querystring("advquery")%>&FromDate=<%=request.querystring ("FromDate") %>&ToDate=<%=request.querystring ("ToDate") %>&iPageCurrent=<%=iPageCurrent-1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/prev.gif" border="0" width="10" height="10"></a> 
							<% end If
							If iPageCount <> 1 then
								For I=1 To iPageCount
									If I=iPageCurrent Then %>
										<%=I%> 
									<% Else %>
										<a href="UPS_ManageShipmentsResults.asp?id=<%=pcv_intOrder%>&TypeSearch=<%=request.querystring("TypeSearch")%>&advquery=<%=request.querystring("advquery")%>&FromDate=<%=request.querystring ("FromDate") %>&ToDate=<%=request.querystring ("ToDate") %>&iPageCurrent=<%=I%>&order=<%=strORD%>&sort=<%=strSort%>"><%=I%></a> 
									<% End If %>
								<% Next %>
							<% end if %>
							<% if CInt(iPageCurrent) <> CInt(iPageCount) then
								'We are not at the end, show a next link %>
								<a href="UPS_ManageShipmentsResults.asp?id=<%=pcv_intOrder%>&TypeSearch=<%=request.querystring("TypeSearch")%>&advquery=<%=request.querystring("advquery")%>&FromDate=<%=request.querystring ("FromDate") %>&ToDate=<%=request.querystring ("ToDate") %>&iPageCurrent=<%=iPageCurrent+1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
							<% end If 
							call closeDb() %>
							</b> 
							</form>						</td>
					</tr>
				</table>			</td>
		</tr>
		<tr>
			<td colspan="11" align="center">
			<% 
			'pcv_strPreviousPage = Request.ServerVariables("HTTP_REFERER")
				pcv_strPreviousPage = "Orddetails.asp?id=" & pcv_intOrder
			%>
				<input type="button" name="Button" value="Go Back To Order Details" onClick="document.location.href='<%=pcv_strPreviousPage%>'" class="ibtnGrey">
			<% else 
				if request("del")="YES" then
					pcv_strPreviousPage = "Orddetails.asp?id=" & request("id") %>
					<input type="button" name="Button" value="There are no Packages to Display >>> Go To Order Details" onClick="document.location.href='<%=pcv_strPreviousPage%>'" class="ibtnGrey">
				<% else %>
						<input type="button" name="Button" value="There are no Packages to Display >>> Go Back" onClick="javascript:history.back()" class="ibtnGrey">
				<% end if
			end if %>		</td>
	</tr>
		<tr>
		  <td colspan="11" align="center"><br /><%= pcf_UPSWriteLegalDisclaimers %></td>
		</tr>
</table>
<%
'// DESTROY THE UPS OBJECT
set objUPSClass = nothing
%>
<!--#include file="AdminFooter.asp"-->
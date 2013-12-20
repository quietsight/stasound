<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Shipping Center for FedEx" %>
<% Section="mngAcc" %>
<%PmAdmin=9%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/FedExconstants.asp"-->
<!--#include file="../includes/pcFedExClass.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="AdminHeader.asp"-->
<% 
Const iPageSize=5

Dim iPageCurrent, conntemp, rs, varFlagIncomplete, uery, strORD, pcv_intOrderID



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'// SET PAGE NAMES
pcPageName = "FedExWS_ManageShipmentsResults.asp"
ErrPageName = "FedExWS_ManageShipmentsTrack.asp"

'// OPEN DATABASE
call openDb()

'// SET THE FEDEX OBJECT
set objFedExClass = New pcFedExClass

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
			<th colspan="3">Manage FedEx<sup>&reg;</sup> Shipments for Order Number <%=(scpre+int(pcv_intOrderID))%><a name="top"></a></th>
		</tr>
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
			<td width="50%" align="center">
				<img src="images/Clct_Prf_2c_Pos_Plt_150.png">
			</td>
		    <td width="25%" align="right" valign="bottom"><input type="button" name="Button2" value="Closeout & Print Manifest" onClick="document.location.href='FedExWS_ManageShipmentsClose.asp?PackageInfo_ID=<%=pcv_intOrderID%>'" class="ibtnGrey"></td>
		</tr>
	</table>
<% end if %>

<table class="pcCPcontent">

	<form name="checkboxform" action="FedExWS_ManageShipmentsTrack.asp?id=<%=pcv_intOrderID%>&action=batch" method="post" class="pcForms">
		<tr> 
			<th nowrap>Shipped</th>
			<th nowrap>Tracking</th>
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
					pcv_strShipMethod = rs("pcPackageInfo_ShipMethod")
					pcv_strFDXRate=rs("pcPackageInfo_FDXRate")
					pcv_strFDXCarrierCode = rs("pcPackageInfo_FDXCarrierCode")

					select case pcv_strFDXCarrierCode
						case "FDXE"
							pcv_strCarrierCode = "FedEx Express"
						case "FDXG"
							pcv_strCarrierCode = "FedEx Ground"
						case "FXCC"
							pcv_strCarrierCode = "FedEx Cargo"
						case "FXFR"
							pcv_strCarrierCode = "FedEx Custom Critical"
						case "FXSP"
							pcv_strCarrierCode = "FedEx Freight"
					end select
					
					select case pcv_strShipMethod
						case "FedEx: PRIORITY_OVERNIGHT"
							pcv_ShipMethod = "FedEx Priority Overnight<sup>&reg;</sup>"
						case "FedEx: STANDARD_OVERNIGHT"
							pcv_ShipMethod = "FedEx Standard Overnight<sup>&reg;</sup>"
						case "FedEx: FEDEX_2_DAY"
							pcv_ShipMethod = "FedEx 2Day<sup>&reg;</sup>"
						case "FedEx: FEDEX_EXPRESS_SAVER"
							pcv_ShipMethod = "FedEx Express Saver<sup>&reg;</sup>"
						case "FedEx: FEDEX_GROUND"
							pcv_ShipMethod = "FedEx Ground<sup>&reg;</sup>"
						case "FedEx: GROUND_HOME_DELIVERY"
							pcv_ShipMethod = "FedEx Home Delivery<sup>&reg;</sup>"
						case "FedEx: INTERNATIONAL_FIRST"
							pcv_ShipMethod = "FedEx International First<sup>&reg;</sup>"
						case "FedEx: INTERNATIONAL_PRIORITY"
							pcv_ShipMethod = "FedEx International Priority<sup>&reg;</sup>"
						case "FedEx: INTERNATIONAL_ECONOMY"
							pcv_ShipMethod = "FedEx International Economy<sup>&reg;</sup>"
						case "FedEx: FEDEX_1_DAY_FREIGHT"
							pcv_ShipMethod = "FedEx 1Day<sup>&reg;</sup> Freight"
						case "FedEx: FEDEX_2_DAY_FREIGHT"
							pcv_ShipMethod = "FedEx 2Day<sup>&reg;</sup> Freight"
						case "FedEx: FEDEX_3_DAY_FREIGHT"
							pcv_ShipMethod = "FedEx 3Day<sup>&reg;</sup> Freight"
						case "FedEx: INTERNATIONAL_PRIORITY_FREIGHT"
							pcv_ShipMethod = "FedEx International Priority<sup>&reg;</sup> Freight"
						case "FedEx: INTERNATIONAL_ECONOMY_FREIGHT"
							pcv_ShipMethod = "FedEx International Economy<sup>&reg;</sup> Freight"
						case "FedEx: INTERNATIONAL_GROUND"
							pcv_ShipMethod = "FedEx International Ground<sup>&reg;</sup>"
						case "FedEx: FEDEX_FREIGHT"
							pcv_ShipMethod = "FedEx<sup>&reg;</sup> Freight"
						case "FedEx: FEDEX_NATIONAL_FREIGHT"
							pcv_ShipMethod = "FedEx National<sup>&reg;</sup> Freight"
						case "FedEx: SMART_POST"
							pcv_ShipMethod = "FedEx SmartPost<sup>&reg;</sup>"
						case "FedEx: EUROPE_FIRST_INTERNATIONAL_PRIORITY"
							pcv_ShipMethod = "FedEx Europe First - Int''l Priority"
					end select
					
					mcount=mcount+1 
					%>
												
					<tr valign="top"> 

						<td bgcolor="<%= strCol %>"><%=ShowDateFrmt(pidPackageShippedDate)%></td>
						<td bgcolor="<%= strCol %>">
							<a href="FedExWS_ManageShipmentsTrack.asp?id=<%=pcv_intOrder%>&PackageInfo_ID=<%=pcv_intPackageInfo_ID%>"><%=pcv_strTrackingNumber%></a>						</td>
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
							  <tr>
								<td align="right" valign="top">Carrier:</td>
								<td align="left" nowrap>
									<b>
										<%=pcv_strCarrierCode%>
									</b>
								</td>
							  </tr>
							  <tr>
								<td align="right" valign="top">Method:</td>
								<td align="left" nowrap>
									<b>
                                        <%=pcv_ShipMethod%>
									</b>
								</td>
							  </tr>
							  <tr>
								<td align="right" nowrap>Net Rate:</td>
								<td align="left">
								<b>
								<%
								if pcv_strFDXRate > 0 then
									response.write scCurSign&money(pcv_strFDXRate)
								else
									response.write "Alternate Payor"
								end if
								%>
								</b></td>
							  </tr>
						  </table>						</td>
						<td bgcolor="<%= strCol %>" nowrap>
							<a href="FedExWS_ManageShipmentsPrinting.asp?path=FedExLabels/Label<%=pcv_strTrackingNumber%>.PNG" target="_blank">View/ Print Label</a><br />
						</td>
						<td bgcolor="<%= strCol %>">
							<a href="FedExWS_ManageShipmentsCancel.asp?id=<%=pcv_intOrder%>&PackageInfo_ID=<%=pcv_intPackageInfo_ID%>">Cancel Shipment</a></td>
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
								<a href="FedEx_ManageShipmentsResults.asp?id=<%=pcv_intOrder%>&TypeSearch=<%=request.querystring("TypeSearch")%>&advquery=<%=request.querystring("advquery")%>&FromDate=<%=request.querystring ("FromDate") %>&ToDate=<%=request.querystring ("ToDate") %>&iPageCurrent=<%=iPageCurrent-1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/prev.gif" border="0" width="10" height="10"></a> 
							<% end If
							If iPageCount <> 1 then
								For I=1 To iPageCount
									If I=iPageCurrent Then %>
										<%=I%> 
									<% Else %>
										<a href="FedExWS_ManageShipmentsResults.asp?id=<%=pcv_intOrder%>&TypeSearch=<%=request.querystring("TypeSearch")%>&advquery=<%=request.querystring("advquery")%>&FromDate=<%=request.querystring ("FromDate") %>&ToDate=<%=request.querystring ("ToDate") %>&iPageCurrent=<%=I%>&order=<%=strORD%>&sort=<%=strSort%>"><%=I%></a> 
									<% End If %>
								<% Next %>
							<% end if %>
							<% if CInt(iPageCurrent) <> CInt(iPageCount) then
								'We are not at the end, show a next link %>
								<a href="FedExWS_ManageShipmentsResults.asp?id=<%=pcv_intOrder%>&TypeSearch=<%=request.querystring("TypeSearch")%>&advquery=<%=request.querystring("advquery")%>&FromDate=<%=request.querystring ("FromDate") %>&ToDate=<%=request.querystring ("ToDate") %>&iPageCurrent=<%=iPageCurrent+1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
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
			pcv_strPreviousPage = "Orddetails.asp?id=" & pcv_intOrder
			%>
				<input type="button" name="Button" value="Go Back To Order Details" onClick="document.location.href='<%=pcv_strPreviousPage%>'" class="ibtnGrey">
			<% 
			else 
			pcv_strPreviousPage = "Orddetails.asp?id=" & Request("id")
			%>
				<input type="button" name="Button" value="There are no Packages to Display >>> Go Back" onClick="document.location.href='<%=pcv_strPreviousPage%>'" class="ibtnGrey">
			<% end if %>		
			</td>
	</tr>
		<tr>
		  <td colspan="11" align="center"><br /><%= pcf_FedExWriteLegalDisclaimers %></td>
		  </tr>
</table>
<%
'// DESTROY THE FEDEX OBJECT
set objFedExClass = nothing
%>
<!--#include file="AdminFooter.asp"-->
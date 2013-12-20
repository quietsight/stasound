<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<%
Dim pageTitle, Section
pageTitle="Modify Shipped Package Information"
Section="orders"

Dim pcv_PackageID, connTemp, query, rs

pcv_PackageID=request("packID")
if pcv_PackageID="" then
	pcv_PackageID="0"
end if

call opendb()

Dim msg
msg=""

if request("info")="UPS" then
	pcv_warning="Modifying a UPS Shipment will not create a new label in ProductCart. <br>If you wish to generate a new label for this shipment, you will need to use the &quot;Reset Shipment&quot; link instead."
end if
if request("m")="del" then
	idOrder=request("id")
	query="UPDATE productsOrdered SET pcPackageInfo_ID=0, pcPrdOrd_Shipped=0 WHERE pcPackageInfo_ID="&pcv_PackageID&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	set rs=nothing
	call closeDb()
	response.redirect "OrdDetails.asp?id="&idOrder&"&ActiveTab=2"
end if

if (request("action")="upd") and (pcv_PackageID<>"0") then
	tmp_ShipMethod=request("pcv_method")
	if tmp_ShipMethod<>"" then
		tmp_ShipMethod=replace(tmp_ShipMethod,"'","''")
	end if
	tmp_TrackingNumber= request("pcv_tracking")
	tmp_ShippedDate=request("pcv_shippedDate")
	tmp_Comments=request("pcv_AdmComments")
	if tmp_Comments<>"" then
		tmp_Comments=replace(tmp_Comments,"'","''")
	end if
	tmp_MethodFlag=request("MethodFlag")
	dim dtShippedDate
	
	'// Reverse "International" Date Format for db entry
	DBInputArray=split(tmp_ShippedDate,"/")
	if scDateFrmt = "DD/MM/YY" then
		'DD/MM/YYYY
		dtInputDbD=DBInputArray(0)
		dtInputDbM=DBInputArray(1)
		dtInputDbY=DBInputArray(2)
	else
		'MM/DD/YYYY
		dtInputDbM=DBInputArray(0)
		dtInputDbD=DBInputArray(1)
		dtInputDbY=DBInputArray(2)
	end if
	if tmp_ShippedDate<>"" then
		if SQL_Format="1" then
			dtShippedDate=(dtInputDbD&"/"&dtInputDbM&"/"&dtInputDbY)
		else
			dtShippedDate=(dtInputDbM&"/"&dtInputDbD&"/"&dtInputDbY)
		end if
	end if	

	if tmp_ShippedDate<>"" then
		if scDB="SQL" then
			query="UPDATE pcPackageInfo SET pcPackageInfo_ShipMethod='" & tmp_ShipMethod & "',pcPackageInfo_TrackingNumber='" & tmp_TrackingNumber & "',pcPackageInfo_ShippedDate='" & dtShippedDate & "',pcPackageInfo_Comments='" & tmp_Comments & "' WHERE pcPackageInfo_ID=" & pcv_PackageID & ";"
		else
			query="UPDATE pcPackageInfo SET pcPackageInfo_ShipMethod='" & tmp_ShipMethod & "',pcPackageInfo_TrackingNumber='" & tmp_TrackingNumber & "',pcPackageInfo_ShippedDate=#" & dtShippedDate & "#,pcPackageInfo_Comments='" & tmp_Comments & "' WHERE pcPackageInfo_ID=" & pcv_PackageID & ";"
		end if
	else
		query="UPDATE pcPackageInfo SET pcPackageInfo_ShipMethod='" & tmp_ShipMethod & "',pcPackageInfo_TrackingNumber='" & tmp_TrackingNumber & "',pcPackageInfo_Comments='" & tmp_Comments & "' WHERE pcPackageInfo_ID=" & pcv_PackageID & ";"
	end if
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	set rs=nothing
	call closeDb()
	msg="Shipment Information was updated successfully!"
	msgtype=1
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<title>Modify Shipment Information</title>
	<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="background-image: none;">
<%if msg<>"" then%>
	<table class="pcCPcontent">
	<tr>
		<td>
			<% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
		</td>
	<tr>
		<td>
		<p align="center">
			<input type=button name="Close Window" value="Close Window" onClick="javascript:window.close();">
		</p>
		</td>
	</tr>
	</table>
<%else
	if pcv_warning<>"" then %>
        <table class="pcCPcontent">
        <tr>
            <td>
                <div class="pcCPmessage">
                    <%=pcv_warning%>
                </div>
            </td>
        </tr>
        </table>
    <% end if %>
    <form name="form1" method="post" action="modifyPackInfo.asp?action=upd" class="pcForms">
		<table class="pcCPcontent">
			<tr>
				<th colspan="2">Shipment information:</th>
			</tr>
			<%
			if validNum(pcv_PackageID) then
				call openDb()
				query="SELECT pcPackageInfo_ShipMethod, pcPackageInfo_TrackingNumber, pcPackageInfo_ShippedDate, pcPackageInfo_Comments, pcPackageInfo_MethodFlag FROM pcPackageInfo WHERE pcPackageInfo_ID=" & pcv_PackageID & ";"
				set rs=server.CreateObject("ADODB.Recordset")
				set rs=connTemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				if not rs.eof then
					tmp_ShipMethod= rs("pcPackageInfo_ShipMethod")
					tmp_TrackingNumber= rs("pcPackageInfo_TrackingNumber")
					tmp_ShippedDate=rs("pcPackageInfo_ShippedDate")
					tmp_Comments=rs("pcPackageInfo_Comments")
					tmp_MethodFlag=rs("pcPackageInfo_MethodFlag")
					%>
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<tr>
						<td>Shipment Method:</td>
						<td><input type="text" name="pcv_method" value="<%=tmp_ShipMethod%>" size="40"></td>
					</tr>
					<tr>
						<td>Tracking Number:</td>
						<td><input type="text" name="pcv_tracking" value="<%=tmp_TrackingNumber%>" size="40"></td>
					</tr>
			
					<tr>
						<td>Shipped Date:</td>
						<td><input type="text" name="pcv_shippedDate" value="<%=ShowDateFrmt(tmp_ShippedDate)%>" size="40"> <span class="pcCPnotes">Date Format: <%=scDateFrmt%></span></td>
					</tr>
					<tr>
						<td valign="top">Comments:</td>
						<td valign="top">
						<textarea name="pcv_AdmComments" size="40" rows="10" cols="50"><%=tmp_Comments%></textarea></td>
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td>
							<input type="submit" name="submit1" value="Update Shipment Information" class="submit2">
							&nbsp;<input type=button name="Close Window" value="Close Window" onClick="javascript:window.close();">
							<input type=hidden name="packID" value="<%=pcv_PackageID%>">						</td>
					</tr>
                    <% if tmp_MethodFlag="4" then %>
                    	<input type="hidden" name="MethodFlag" value="4">
					<% end if
				end if
			set rs=nothing
			call closeDb()
		else
		%>
        <tr>
        	<td colspan="2">Not a valid package ID</td>
        </tr>
        <%
		end if
		%>
	</table>
</form>
<%end if%>
</body>
</html>
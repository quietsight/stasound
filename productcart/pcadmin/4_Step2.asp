<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Canada Post Shipping Configuration" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<% dim query, rs, conntemp
call opendb ()

query="SELECT * FROM shipService WHERE serviceCode='2005';"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if rs.eof then

	query="INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee,  serviceLimitation) VALUES (0, '2005', 0, 'Canada Post Small Packets Surface US', 0, 0, 0, 0, 0, 0);"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	
	query="INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee,  serviceLimitation) VALUES (0, '2015', 0, 'Canada Post Small Packets Air US', 0, 0, 0, 0, 0, 0);"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	
	query="INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee,  serviceLimitation) VALUES (0, '2025', 0, 'Canada Post Expedited Commercial US', 0, 0, 0, 0, 0, 0);"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	
	query="INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee,  serviceLimitation) VALUES (0, '3005', 0, 'Canada Post Small Packets Surface International', 0, 0, 0, 0, 0, 0);"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	
	query="INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee,  serviceLimitation) VALUES (0, '3015', 0, 'Canada Post Small Packets Air International', 0, 0, 0, 0, 0, 0);"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	
	query="INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee,  serviceLimitation) VALUES (0, '3025', 0, 'Canada Post Xpresspost International', 0, 0, 0, 0, 0, 0);"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	
	query="INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee,  serviceLimitation) VALUES (0, '3050', 0, 'Canada Post Puropak International', 0, 0, 0, 0, 0, 0);"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	
end if

set rs=nothing
call closeDb()

dim strCPServiceArray
'//split array
strCPServiceArray="1010, 1020, 1030, 1040, 1120, 1130, 1220, 1230, 2005, 2010, 2015, 2020, 2025, 2030, 2040, 2050, 3005, 3010, 3015, 3020, 3025, 3040, 3050"

strSpltServiceArray=split(strCPServiceArray, ", ")


if request.form("submit")<>"" then
	CP_Service=request.form("CP_Service")
	Session("ship_CP_Service")=CP_Service
	if CP_Service="" then
		response.redirect "4_Step2.asp?msg="&Server.URLEncode("Select at least one service.")
		response.end
	end if
	freeshipStr=""
	handlingStr=""
	
	for aCnt=lbound(strSpltServiceArray) to ubound(strSpltServiceArray)
		If request.form("free"&strSpltServiceArray(aCnt))="YES" then
			freeamt=request.form("amt"&strSpltServiceArray(aCnt))
			freeshipStr=freeshipStr&strSpltServiceArray(aCnt)&"|"&replacecomma(freeamt)&","
		End if
		If request.form("handling"&strSpltServiceArray(aCnt))<>"0" AND request.form("handling"&strSpltServiceArray(aCnt))<>"" then
			If isNumeric(request.form("handling"&strSpltServiceArray(aCnt)))=true then
				handlingStr=handlingStr&strSpltServiceArray(aCnt)&"|"&replacecomma(request.form("handling"&strSpltServiceArray(aCnt)))&"|"&request.form("shfee"&strSpltServiceArray(aCnt))&","
			End If
		End if
	Next	

	Session("ship_CP_freeshipStr")=freeshipStr
	Session("ship_CP_handlingStr")=handlingStr
	response.redirect "4_Step3.asp"
	response.end
else %>
	<form name="form1" method="post" action="4_Step2.asp" class="pcForms">
		<table class="pcCPcontent">
            <tr>
                <td colspan="2" class="pcCPspacer">
                    <% ' START show message, if any %>
                        <!--#include file="pcv4_showMessage.asp"-->
                    <% 	' END show message %>
                </td>
            </tr>
			<tr> 
				<td colspan="2">Choose one or more shipping services to offer to your customers.</td>
			</tr>

			<% 	for bCnt=lbound(strSpltServiceArray) to ubound(strSpltServiceArray) 
				select case strSpltServiceArray(bcnt)
					case "1010"
						serviceName="Canada Post Regular" 
					case "1020"
						serviceName="Canada Post Expedited"
					case "1030"
						serviceName="Canada Post Xpresspost"
					case "1040"
						serviceName="Canada Post Priority Courier"
					case "1120"
						serviceName="Canada Post Expedited Evening"
					case "1130"
						serviceName="Canada Post Xpresspost Evening"
					case "1220"
						serviceName="Canada Post Expedited Saturday"
					case "1230"
						serviceName="Canada Post Xpresspost Saturday"
					case "2005"
						serviceName="Canada Post Small Packets Surface US"
					case "2010"
						serviceName="Canada Post Surface US"
					case "2015"
						serviceName="Canada Post Small Packets Air US"
					case "2020"
						serviceName="Canada Post Air US"
					case "2025"
						serviceName="Canada Post Expedited Commercial US"
					case "2030"
						serviceName="Canada Post Xpresspost US"
					case "2040"
						serviceName="Canada Post Puroloator US"
					case "2050"
						serviceName="Canada Post Puropak US"
					case "3005"
						serviceName="Canada Post Small Packets Surface International"
					case "3010"
						serviceName="Canada Post Surface International"
					case "3015"
						serviceName="Canada Post Small Packets Air International"
					case "3020"
						serviceName="Canada Post Air International"
					case "3025"
						serviceName="Canada Post Xpresspost International"
					case "3040"
						serviceName="Canada Post Puroloator International"
					case "3050"
						serviceName="Canada Post Puropak International"
				end select %>
                <tr>
                    <td colspan="2" class="pcCPspacer"></td>
                </tr>
				<tr bgcolor="#DDEEFF">
					<td align="right"> 
						<input type="checkbox" name="CP_Service" value="<%=strSpltServiceArray(bcnt)%>">
					</td>
					<td><font color="#3366CC"><b><%=serviceName%></b></font></td>
				</tr>
				<tr> 
					<td>&nbsp;</td>
					<td>
						<input name="free<%=strSpltServiceArray(bcnt)%>" type="checkbox" id="free<%=strSpltServiceArray(bcnt)%>" value="YES">
						Offer free shipping for orders over <%=scCurSign%> 
						<input name="amt<%=strSpltServiceArray(bcnt)%>" type="text" id="amt<%=strSpltServiceArray(bcnt)%>" size="10" maxlength="10">
					</td>
				</tr>
				<tr> 
					<td>&nbsp;</td>
					<td>Add Handling Fee <%=scCurSign%> 
					<input name="handling<%=strSpltServiceArray(bcnt)%>" type="text" id="handling<%=strSpltServiceArray(bcnt)%>" size="10" maxlength="10">
					</td>
				</tr>
				<tr> 
					<td>&nbsp;</td>
					<td>
					<input type="radio" name="shfee<%=strSpltServiceArray(bcnt)%>" value="-1" checked>
					Display as a &quot;Shipping &amp; Handling&quot; charge.<br>
					<input type="radio" name="shfee<%=strSpltServiceArray(bcnt)%>" value="0">
					Integrate into shipping rate.</td>
				</tr>
			<% next %>        

            <tr>
                <td colspan="2" class="pcCPspacer"><hr></td>
            </tr>
                            
			<tr> 
				<td>&nbsp;</td>
				<td>
				<input type="submit" name="Submit" value="Submit" class="submit2"></td>
			</tr>
		</table>
	</form>
<% end if %>
<!--#include file="AdminFooter.asp"-->
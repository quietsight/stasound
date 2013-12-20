<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% pageTitle="Purge PayFlow Pro Credit Card Numbers" %>
<% Section="" %>
<%PmAdmin=9%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/rc4.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="AdminHeader.asp"-->
<%dim query, rs, conntemp

dim intSuccessCnt
intSuccessCnt=0

dim intPurgeprocess
intPurgeprocess=0

if request.Form("PurgeNumbers")<>"" then
	intPurgeprocess=1
	dim strSuccessData
	strSuccessData=""
	'get the count
	pc_CardCnt=request.Form("iCnt")
	dim i
	for i=1 to pc_CardCnt
		'see if checkbox is checked
		if request.Form("idOrder"&i)="1" then
			tempOrderId=request.Form("ccOrderID"&i)
			call opendb()

			query="SELECT acct, pcSecurityKeyID FROM pfporders WHERE idpfporder="&tempOrderId&";"
			set rstemp=connTemp.execute(query)
			if NOT rstemp.EOF then
				cardnumber=rstemp("acct")
				tempSecurityKeyID=rstemp("pcSecurityKeyID")
			end if
			set rstemp=nothing
			
			tempfour=pcf_PurgeCardNumber(cardnumber,tempSecurityKeyID)

			query="UPDATE pfporders SET acct='"&tempfour&"' WHERE idpfporder="&tempOrderId&";"
			set rs=server.CreateObject("ADODB.RecordSet") 
			set rs=connTemp.execute(query)
			
			call closedb()
			intSuccessCnt=intSuccessCnt+1
			strSuccessData=strSuccessData&"Credit card number successfully purged for <strong>order #"&(scpre+int(tempOrderId))&"</strong>&nbsp;|&nbsp;<a href=Orddetails.asp?id="&int(tempOrderId)&">View Order</a><BR>"
		end if
	next
end if

if request.Form("search1")<>"" then
'show results 
	call opendb()
	pcv_fromMonth=request.Form("fromMonth")
	pcv_fromDay=request.Form("fromDay")
	pcv_fromYear=request.Form("fromYear")
	pcv_fromDate=pcv_fromMonth&"/"&pcv_fromDay&"/"&pcv_fromYear
	pcv_toMonth=request.Form("toMonth")
	pcv_toDay=request.Form("toDay")
	pcv_toYear=request.Form("toYear")
	pcv_toDate=pcv_toMonth&"/"&pcv_toDay&"/"&pcv_toYear
	pcv_captured=request.Form("captured")
	if scDB="SQL" then
		query="SELECT pfporders.idpfporder, pfporders.idOrder, orders.orderDate, pfporders.acct, pfporders.captured, pfporders.pcSecurityKeyID FROM pfporders INNER JOIN orders ON pfporders.idOrder = orders.idOrder WHERE (((orders.orderDate)='"&pcv_fromDate&"' OR (orders.orderDate)>'"&pcv_fromDate&"') AND ((orders.orderDate)<'"&pcv_toDate&"' OR (orders.orderDate)='"&pcv_toDate&"') AND ((pfporders.captured)="&pcv_captured&"));"
	else
		query="SELECT pfporders.idpfporder, pfporders.idOrder, orders.orderDate, pfporders.acct, pfporders.captured, pfporders.pcSecurityKeyID FROM pfporders INNER JOIN orders ON pfporders.idOrder = orders.idOrder WHERE (((orders.orderDate)=#"&pcv_fromDate&"# OR (orders.orderDate)>#"&pcv_fromDate&"#) AND ((orders.orderDate)<#"&pcv_toDate&"# OR (orders.orderDate)=#"&pcv_toDate&"#) AND ((pfporders.captured)="&pcv_captured&"));"
	end if
	set rs=server.CreateObject("ADODB.RecordSet") 
	set rs=connTemp.execute(query)
	
	'show results
	dim iCnt
	iCnt=0
	if NOT rs.eof then %>
		<form name="form1" method="post" action="PFPCCPurge.asp">
		<table width="94%" border="0" cellspacing="0" cellpadding="4" align="center">
			<tr bgcolor="#e5e5e5" class="normal"> 
				<td width="20" align="center" nowrap> 
					<div align="left"></div></td>
				<td width="115" nowrap><b>Date of Order</b></td>
		    <td width="484" nowrap><b>Order Number</b></td>
		</tr>
		<tr class="normal"> 
			<td height="1" colspan="3" background="images/pc_px.gif"></td>
		</tr>
			<% 
			do until rs.eof
				pcv_idpfporder=rs("idpfporder")
				pcv_idOrder=rs("idOrder")
				pcv_CCnumber=rs("acct")
				pcv_orderDate=rs("orderDate")
				pcv_captured=rs("captured")
				pcv_SecurityKeyID=rs("pcSecurityKeyID")
				
				if isNull(pcv_CCnumber) OR pcv_CCnumber="" then
					pcv_CCnumber="*"
				end if
				
				if instr(pcv_CCnumber,"*") then
				else
					pcv_SecurityPass = pcs_GetKeyUsed(pcv_SecurityKeyID)
					
					'decrypt CC data
					pcv_DecryptedCC=enDeCrypt(pcv_CCnumber, pcv_SecurityPass)
					iCnt=iCnt+1
					%>
					<tr class="normal">
						<td width="20"><input name="idOrder<%=iCnt%>" type="checkbox" value="1" checked>
						<input type="hidden" name="ccOrderID<%=iCnt%>" value="<%=pcv_idpfpOrder%>">
						<td class="normal"><%=pcv_orderDate%></td>
					  <td class="normal"><%=(scpre+int(pcv_idOrder))%></td>
					</tr>
				<% end if
				rs.movenext
			loop
			call closedb() %>
			<input name="iCnt" type="hidden" value="<%=iCnt%>">
			<% if iCnt=0 then %>
			<tr class="normal">
				<td colspan="3">No records found.</td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td colspan="2">&nbsp;</td>
			</tr>
			<% else %>
			<tr>
				<td>&nbsp;</td>
				<td colspan="2">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="3"><input name="PurgeNumbers" type="submit" class="ibtnGrey" value="Purge CC Numbers"></td>
			</tr>
			<% end if %>
		</table>
    </form>
	<% else %>
			<form name="form1" method="post" action="PFPCCPurge.asp">
			<table width="94%" border="0" cellspacing="0" cellpadding="4" align="center">
				<tr bgcolor="#e5e5e5" class="normal"> 
					<td width="20" align="center" nowrap> 
						<div align="left"></div></td>
					<td width="115" nowrap><b>Date of Order</b></td>
					<td width="484" nowrap><b>Order Number</b></td>
			</tr>
			<tr class="normal"> 
				<td height="1" colspan="3" background="images/pc_px.gif"></td>
			</tr>
			<tr class="normal">
				<td colspan="3">No records found.</td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td colspan="2">&nbsp;</td>
			</tr>
		</table>
		</form>
	<% end if %>
<% else
	if intPurgeprocess=0 then %>
		<table width="94%" border="0" cellspacing="0" cellpadding="4" align="center">
			<tr> 
				<td bgcolor="#e5e5e5" class="normal"> <b>Search Orders by Date</b></td>
			</tr>
			<tr class="normal"> 
				<td height="1" colspan="3" background="images/pc_px.gif"></td>
			</tr>
			<tr> 
				<td class="normal"> Enter a date range and select the status from the drop-menu below:</td>
			</tr>
			<tr> 
				<td class="normal">
			<form action="PFPCCPurge.asp" method="post"  name="CCPurgeSearch" align="center">
					<%
					FromDate=Date()
					FromDate=FromDate-13
					ToDate=Date()
					%>
					<table cellpadding="4" cellspacing="2">
						<tr class="normal">
							<td width="61">Date From:</td>
							<td width="525"  valign="top">Month: 
								<input type=text name="fromMonth" value="<%=month(FromDate)%>" size="2" maxlength="4">
								Day:
								<input type=text name="fromDay" value="<%=day(FromDate)%>" size="2" maxlength="4"> 
								Year:
								<select name="fromYear">
									<% Dim varYear
									varYear=year(now) %>
									<option value="<%=varYear-4%>"><%=varYear-4%></option>
									<option value="<%=varYear-3%>"><%=varYear-3%></option>
									<option value="<%=varYear-2%>"><%=varYear-2%></option> 
									<option value="<%=varYear-1%>"><%=varYear-1%></option>
									<option value="<%=varYear%>" selected><%=varYear%></option>
							</select></font></td>
						</tr>
						<tr class="normal">
							<td>Date To:</td>
							<td>Month:
								<input type=text name="toMonth" value="<%=month(ToDate)%>" size="2" maxlength="4">
								Day:
								<input type=text name="toDay" value="<%=day(ToDate)%>" size="2" maxlength="4">
								Year:
								<select name="toYear">
								<% 
								varYear=year(now) %>
									<option value="<%=varYear-4%>"><%=varYear-4%></option>
									<option value="<%=varYear-3%>"><%=varYear-3%></option>
									<option value="<%=varYear-2%>"><%=varYear-2%></option> 
									<option value="<%=varYear-1%>"><%=varYear-1%></option>
									<option value="<%=varYear%>" selected><%=varYear%></option>
								</select></font></td>
						</tr>
						<tr class="normal">
							<td>Status:</td>
						<td>
						Captured<input type="hidden" name="captured" value="1">
					</td>
						</tr>
						<tr class="normal">
							<td>&nbsp;</td>
							<td><input type="submit" name="search1" value="View " class="ibtnGrey"></td>
						</tr>
				</table>
				</form>
				</td>
			</tr>			
			<tr>
				<td class="normal">&nbsp;</td>
			</tr>			
		</table>
	<% else 
		if intSuccessCnt=0 then %>
		<% else %>
			<table width="94%" border="0" cellspacing="0" cellpadding="4" align="center">
				<tr class="normal">
					<td><font color="#FF0000"><strong><%=intSuccessCnt%></strong></font>&nbsp;Credit card numbers were successfully purged for the selected orders:<br>
						<% if strSuccessData<>"" then %>
							<br><%=strSuccessData%><br>
						<% end if %>
					</td>
				</tr>
				<tr class="normal">
					<td><p>&nbsp;</p>
					<p><a href="resultsAdvancedAll.asp?B1=View%2BAll&dd=1">Manage Orders</a></p></td>
				</tr>
			</table>
		<% end if %>
	<% end if %>
<% end if %>
<!--#include file="AdminFooter.asp"-->
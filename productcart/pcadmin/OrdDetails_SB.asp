<link type="text/css" rel="stylesheet" href="subscriptionBridge.css" />
<%
query="SELECT SB_GUID, SB_Terms FROM SB_Orders WHERE idOrder=" & qry_ID & ";"

Set rsSB=Server.CreateObject("ADODB.Recordset")
Set rsSB=connTemp.execute(query)
If NOT rsSB.eof Then
	pcv_strGUID = rsSB("SB_GUID")
	pcv_strTerms = rsSB("SB_Terms")
	%>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">SubscriptionBridge</th>
	</tr>	
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>	
	<tr>
		<td colspan="2">
			<a href="https://www.subscriptionbridge.com/MerchantCenter/" target="_blank">Manage Subscriptions</a>
		</td>
	</tr>
	<% if pcv_strGUID<>"" then %>
		<tr>
			<td colspan="2">
				Subscription ID: <strong><%=pcv_strGUID%></strong>
			</td>
		</tr>
		<tr>
			<td colspan="2">
				<%=pcv_strTerms%>
			</td>
		</tr>
	<% end if %>
    
<% End If %>
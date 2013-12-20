<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/bto_language.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/ShipFromSettings.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/SQLFormat.txt"-->
<% dim conntemp, query, rs

'Check to see if ARB has been turned off by admin, then display message
If scSBStatus="0" then
	response.redirect "msg.asp?message=212"
End If 

if scSSL = "1" then
	If (Request.ServerVariables("HTTPS") = "off") Then
		Dim xredir__, xqstr__
		xredir__ = "https://" & Request.ServerVariables("SERVER_NAME") & _
				   Request.ServerVariables("SCRIPT_NAME")
		xqstr__ = Request.ServerVariables("QUERY_STRING")
		if xqstr__ <> "" Then xredir__ = xredir__ & "?" & xqstr__
		Response.redirect xredir__
	End if
end if

call opendb()

'SB S
query="SELECT orders.idOrder, orders.orderDate, orders.total, orders.ord_OrderName, ProductsOrdered.idProductOrdered,ProductsOrdered.UnitPrice,ProductsOrdered.quantity, ProductsOrdered.pcSubscription_ID, ProductsOrdered.pcPO_SubAmount, ProductsOrdered.pcPO_SubActive, ProductsOrdered.pcPO_IsTrial, ProductsOrdered.pcPO_SubTrialAmount, ProductsOrdered.pcPO_SubStartDate, ProductsOrdered.pcPO_SubType FROM orders, productsordered WHERE orders.idCustomer=" & Session("idcustomer") &" AND orders.OrderStatus>1  And orders.idOrder = ProductsOrdered.idOrder and ProductsOrdered.pcSubscription_ID >0  ORDER BY ProductsOrdered.idProductOrdered DESC"
'SB E
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(query)
if err.number <> 0 then
	set rstemp=nothing
	call closeDb()
 	response.redirect "techErr.asp?error="& Server.Urlencode("Error in sb_CustViewSubs.asp: "&err.description) 
end If
if rstemp.eof then
	set rstemp=nothing
	call closeDb()
 	response.redirect "msg.asp?message=34"     
end if           
%> 

<!--#include file="header.asp"-->
<div id="pcMain">
	<table class="pcMainTable">   
		<tr>
			<td>
				<h1>
				<%
				if session("pcStrCustName") <> "" then
					response.write(session("pcStrCustName") & " -  Subscriptions")
					else
					response.write "Subscriptions"
				end if
				%>
				</h1>
			</td>
		</tr>
		<tr>
			<td>     
			<table class="pcShowContent">
				<tr>
			    <th nowrap><%response.write dictLanguage.Item(Session("language")&"_SB_10")%></th>
					<th nowrap><%response.write dictLanguage.Item(Session("language")&"_SB_11")%></th>
					<th nowrap><%response.write dictLanguage.Item(Session("language")&"_SB_12")%></th>
					<th nowrap>&nbsp;</th>
					<th>&nbsp;</th>
				</tr>
				<tr class="pcSpacer">
					<td colspan="5"></td>
				</tr>
				<%
				do while not rstemp.eof
					idorder = rstemp("idOrder")
					idProductOrdered = rstemp("idProductOrdered")
					pSubUnitPrice = rstemp("unitPrice")
					pSubQty = rstemp("quantity")
					pSubPrice = rstemp("pcPO_SubAmount")
					pSubTrial = rstemp("pcPO_IsTrial")
					pSubTrialAmount = rstemp("pcPO_SubTrialAmount")						 
					pSubStartDate = rstemp("pcPO_SubStartDate")
					pSubActive =rstemp("pcPO_SubActive") 
					pSubType=rstemp("pcPO_SubType")					
					
					'// Obtain Status
					Dim pvc_Status
					if pSubActive = "1"  Then
						pcv_Status = "Active"
					Elseif pSubActive = "2" then
						pcv_Status = "Pending"
					else
						pcv_Status = "<font color='#ff0000'>Not-Active</font>"
					End if 
					
					'// Obtain GUID and Email
					Dim pcv_strCustEmail, pcv_strGUID
					pcv_strCustEmail=""
					pcv_strGUID=""

					query = "SELECT customers.email, SB_Orders.SB_GUID FROM SB_Orders "
					query = query & "INNER JOIN ( orders INNER JOIN customers On orders.idcustomer = customers.idcustomer ) ON orders.idorder = SB_Orders.idorder "
					query = query & "WHERE orders.idorder = " & idorder

					set rsSB=Server.CreateObject("ADODB.Recordset")
					set rsSB=conntemp.execute(query)
					if err.number <> 0 then
						set rsSB=nothing
						call closeDb()
						response.redirect "techErr.asp?error="& Server.Urlencode("Error in sb_CustViewSubs.asp: "&err.description) 
					end If
					if NOT rsSB.eof then
						   pcv_strCustEmail = rsSb("email")
						   pcv_strGUID = rsSb("SB_GUID")
					end if  
					set rsSB=nothing 


					If len(pcv_strGUID)>0 Then

						query="SELECT Setting_APIUser,Setting_APIPassword,Setting_APIKey,Setting_RegSuccess FROM SB_Settings;"
						set rsAPI=connTemp.execute(query)
						if not rsAPI.eof then
							Setting_APIUser=rsAPI("Setting_APIUser")
							Setting_APIPassword=enDeCrypt(rsAPI("Setting_APIPassword"), scCrypPass)
							Setting_APIKey=enDeCrypt(rsAPI("Setting_APIKey"), scCrypPass)
						end if
						set rsAPI=nothing
						
						
						Set objSB = NEW pcARBClass
						
						objSB.GUID = pcv_strGUID
						If scSBLanguageCode<>"" Then
							objSB.CartLanguageCode = scSBLanguageCode
						Else
							objSB.CartLanguageCode = "en-EN"
						End If
			
						Dim result
	
						result = objSB.GetSubscriptionDetailsRequest(Setting_APIUser, Setting_APIPassword, Setting_APIKey)

						If SB_ErrMsg="" Then
							
							pcv_strGUID = objSB.pcf_GetNode(result, "Guid", "//GetSubscriptionDetailsResponse/Subscription/Identifiers")
							pcv_strStatus = objSB.pcf_GetNode(result, "Status", "//GetSubscriptionDetailsResponse/Subscription/Identifiers")
							pcv_strBalance = objSB.pcf_GetNode(result, "BalanceTotal", "//GetSubscriptionDetailsResponse/Subscription")
							pcv_strBillingAgreement = objSB.pcf_GetNode(result, "Terms", "//GetSubscriptionDetailsResponse/Subscription")

							if pcv_strBalance="" then
								pcv_strBalance = 0
							end if
							
							If len(pcv_strGUID)>0 Then
							%>
							<tr>
								<td>
								   <a href="CustviewPastD.asp?idOrder=<%response.write (scpre+int(IdOrder))%>"><%response.write (scpre+int(IdOrder))%></a>
								</td>
								<td>
									<a href="JavaScript:openManageSubscription('sb_CustSubDetails.asp?ID=<%=idorder%>&GUID=<%=pcv_strGUID%>')"><%=pcv_strGUID%></a>
                                    <br />
                                    <i><%response.write dictLanguage.Item(Session("language")&"_SB_31")%>: <%=pcv_strStatus%></i>
								</td>
								<td><%=pcv_strBillingAgreement%></td>
								<td>&nbsp;</td>
	
								<td nowrap>
										
										<a href="<%=gv_RootURL%>/CustomerCenter/AutoLogin.asp?ID=<%=pcv_strGUID%>&Email=<%=pcv_strCustEmail%>&mode=details" target="_blank"><%response.write dictLanguage.Item(Session("language")&"_SB_4")%></a>
											
								</td>
							</tr>
							<tr>
								<td colspan="5">
								
										<% If pcv_strBalance>0 AND scSSL = "1" AND lcase(pcv_strStatus)="active" Then %>
										
											<div align="center" class="pcErrorMessage">
												<%response.write dictLanguage.Item(Session("language")&"_SB_16")%>&nbsp; 
												<a href="JavaScript:openManageSubscription('sb_CustOneTimePayment.asp?ID=<%=idorder%>&GUID=<%=pcv_strGUID%>')"><%response.write dictLanguage.Item(Session("language")&"_SB_17")%></a> 
											</div>
											
											
										<% End If %>
								
										<div align="left" class="pcSmallText" style="display:none">
										
										
	
											<a href="<%=gv_RootURL%>/CustomerCenter/AutoLogin.asp?ID=<%=pcv_strGUID%>&Email=<%=pcv_strCustEmail%>&mode=history" target="_blank"><%response.write dictLanguage.Item(Session("language")&"_SB_5")%></a>
		
											<% if pSubActive = "1" Then %>
										
												  &nbsp;|&nbsp;
	
												  <a href="JavaScript:openManageSubscription('sb_CustUpdatePayment.asp?ID=<%=idorder%>')"><%response.write dictLanguage.Item(Session("language")&"_SB_14")%></a>
												  &nbsp;|&nbsp;
	
												  <a href="JavaScript:openManageSubscription('sb_CustCancelSub.asp?ID=<%=idorder%>&GUID=<%=pcv_strGUID%>')"><%response.write dictLanguage.Item(Session("language")&"_SB_8")%></a>
												  &nbsp;
				   
				  
										   <% elseif pSubActive = "2" then %>
												  
												  <a href="<%=gv_RootURL%>/CustomerCenter/AutoLogin.asp?ID=<%=pcv_strGUID%>&Email=<%=pcv_strCustEmail%>&mode=" target="_blank"><%response.write dictLanguage.Item(Session("language")&"_SB_9")%></a>&nbsp;
												 
										   <%else%>
												  
												  <%=pcv_status%>
												  
										   <%end if %>
										</div>
								
								</td>
							</tr>
							<tr>
								<td colspan="5"><hr></td>
							</tr>
							<%
							End If
						End If
						
					End If '// If len(pcv_strGUID)>0 Then
						
					rstemp.movenext
			  	loop
				%>
			</table>
			<% 
            set rstemp = nothing
            call closeDb()
            %>
			</td>
		</tr>
		<tr> 
            <td><a href="custPref.asp"><img src="<%=rslayout("back")%>"></a></td>
        </tr>
    </table>
    <% '// START: DIALOG %>
    <div id="Dialog" title="Manage Subscription" style="display:none; overflow:hidden">
        <div id="DialogLoader" style="overflow:hidden"><img src="images/ajax-loader1.gif" width="20" height="20" align="absmiddle"><%=dictLanguage.Item(Session("language")&"_opc_loadcontent")%></div>
        <iframe name="DialogFrame" id="DialogFrame" src="" width="600" height="650" frameborder="0" scrolling="auto"></iframe>
    </div>
    <% '// END: DIALOG %>
</div>
<script>
$(document).ready(function()
{
	// START: DIALOG
	$("#Dialog").dialog({
			bgiframe: true,
			autoOpen: false,
			width: 650,
			minHeight: 50,
			modal: true,
			open: function(event,ui)
			{
				$('#DialogLoader').show();
				$('#DialogFrame').hide();
			},
			close: function() {
				
			}
	});
	
	$('#DialogFrame').load( function() {
		$('#DialogLoader').hide();
		$('#DialogFrame').show();
	} );
	// END: DIALOG

});	

function openManageSubscription(a) {
	document.getElementById("DialogFrame").src=a;
	$("#Dialog").dialog("open");
}

function closeManageSubscription(a) {
	$("#Dialog").dialog("close");
	window.location = "sb_CustViewSubs.asp"
}
</script>
<!--#include file="footer.asp"-->
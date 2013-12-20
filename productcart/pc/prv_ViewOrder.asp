<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
' PRV41 start
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/ErrorHandler.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/securitysettings.asp" -->
<% 

Dim pIDOrder, pIDCustomer, pCustGuest, pcv_UniqueID
Dim pcv_PrdRevExc : pcv_PrdRevExc = "0"

' Check if the store is on. If store is turned off display store message
If scStoreOff="1" then
	response.redirect "msg.asp?message=83"
End If

%>
<!--#include file="pcStartSession.asp"-->
<%
Dim connTemp,rs
call opendb()
%>
<!--#include file="prv_getsettings.asp"-->
<%
if pcv_Active<>"1" then
	call closedb()
	response.redirect "default.asp"
end if

pcv_UniqueID=GetUserInput(request("UID"),0)
if len(pcv_UniqueID)<>36 then
	call closedb()
	response.redirect "msg.asp?message=210"
end If

query="SELECT pcRN_idCustomer, pcRN_idOrder, pcRN_DateLastViewed FROM pcReviewNotifications WHERE pcRN_UniqueID='" & pcv_UniqueID & "'"
set rs=connTemp.execute(query)

if rs.eof then
	set rs=nothing
	call closedb()
	response.redirect "msg.asp?message=210"
end If

' Update the date this order product review list was last viewed
If IsNull(rs("pcRN_DateLastViewed")) then
   query = "UPDATE pcReviewNotifications SET pcRN_DateLastViewed=" & formatDateForDB(now) & " WHERE pcRN_UniqueID='" & pcv_UniqueID & "'"
   connTemp.execute query
End if

pIdOrder = rs("pcRN_idOrder")
pIdCustomer = rs("pcRN_idCustomer")

query = "SELECT name, pcCust_Guest FROM customers WHERE idCustomer=" & pIDCustomer
Set rs = connTemp.execute(query)
If rs.eof Then
	set rs=nothing
	call closedb()
	response.redirect "msg.asp?message=210" ' Give the generic message to discourage script-kiddies
else
   pCustName = rs("name")
   pCustGuest = CLng(rs("pcCust_Guest"))
End If
rs.close

If pCustGuest<>0 Then
   session("CustomerGuest") = CStr(pCustGuest)
End if

Dim strRewardPrompt
Dim pRewardForReview, pRewardForReviewURL, pRewardForReviewFirstPts, pRewardForReviewAdditionalPts
pRewardForReview = 0
pRewardForReviewURL = ""
pRewardForReviewFirstPts = 0
pRewardForReviewAdditionalPts = 0

query = "SELECT pcRS_RewardForReview, pcRS_RewardForReviewURL, pcRS_RewardForReviewFirstPts, pcRS_RewardForReviewAdditionalPts, pcRS_RewardForReviewMaxPts FROM pcRevSettings"
Set rs = connTemp.execute(query)
If rs.eof = False Then
   pRewardForReview = rs("pcRS_RewardForReview")
   pRewardForReviewURL = rs("pcRS_RewardForReviewURL")
   pRewardForReviewFirstPts = rs("pcRS_RewardForReviewFirstPts")
   pRewardForReviewAdditionalPts = rs("pcRS_RewardForReviewAdditionalPts")
End If
rs.close

query="SELECT Orders.OrderDate,productsOrdered.idProductOrdered,products.description, products.sku, products.idProduct, products.active, products.removed, products.smallImageUrl FROM ProductsOrdered, products, orders WHERE orders.idorder=ProductsOrdered.idOrder AND ProductsOrdered.idProduct=products.idProduct AND orders.idCustomer=" & pIdCustomer & " AND orders.idOrder=" & pIdOrder
set rsOrdObj=conntemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rsOrdObj=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end If
%>
<!--#include file="header.asp"-->
<script>
function openbrowser(url) {
				self.name = "productPageWin";
				popUpWin = window.open(url,'rating','toolbar=0,location=0,directories=0,status=0,top=0,scrollbars=yes,resizable=1,width=705,height=535');
				if (navigator.appName == 'Netscape') {
								popUpWin.focus();
				}
}
</script>

<div id="pcMain">
	<table class="pcMainTable">   
	<tr>
		<td>
		   <p>
		   <%
	       strRewardPrompt = dictLanguage.Item(Session("language")&"_prv_24")
		   strRewardPrompt = Replace(strRewardPrompt,"<customer name>",ProperCase(pCustName))
		   response.write strRewardPrompt
		   %>
           </p>
           <%
           If pRewardForReview<>0 Then
		      strRewardPrompt = dictLanguage.Item(Session("language")&"_prv_22")
			  strRewardPrompt = Replace(strRewardPrompt,"<RFR_PAGE>",pRewardForReviewURL)
			  strRewardPrompt = Replace(strRewardPrompt,"<REWARD_POINTS_LABEL>",RewardsLabel)
		      response.write "<p>&nbsp;</p><p>" & strRewardPrompt & "</p><p>&nbsp;</p>"
           End If

			if pCustGuest=1 and pRewardForReview<>0 then%>
					<div id="PwdArea">
						<form id="PwdForm" name="PwdForm">
						<table class="pcShowContent">
						<tr>
							<th colspan="4" class="pcSectionTitle"><%=dictLanguage.Item(Session("language")&"_opc_common_2")%></th>
						</tr>
						<tr>
							<td colspan="4"><%
							strRewardPrompt = dictLanguage.Item(Session("language")&"_prv_27")
							strRewardPrompt = Replace(strRewardPrompt,"<REWARD_POINTS_LABEL>",RewardsLabel)
							response.write strRewardPrompt
							%></td>
						</tr>
						<tr>
							<td width="20%" nowrap><%=dictLanguage.Item(Session("language")&"_opc_6")%></td>
							<td width="30%"><input type="password" name="newPass1" id="newPass1" size="20"></td>
							<td width="20%" nowrap><%=dictLanguage.Item(Session("language")&"_opc_38")%></td>
							<td width="30%"><input type="password" name="newPass2" id="newPass2" size="20"></td>
						</tr>
						<tr>
							<td colspan="4" style="padding-top: 10px;"></td>
						</tr>
						<tr>
							<td colspan="4" style="padding-top: 10px;"><input type="button" name="PwdSubmit" id="PwdSubmit" value="<%=dictLanguage.Item(Session("language")&"_opc_common_4")%>" class="submit2"></td>
						</tr>
						</table>
						</form>
				</div>
				<div id="PwdLoader" style="display:none"></div>
		<script>
		$(document).ready(function()
		{
			jQuery.validator.setDefaults({
				success: function(element) {
					$(element).parent("td").children("input, textarea").addClass("success")
				}
			});
			
			//*Ajax Global Settings
			$("#GlobalAjaxErrorDialog").ajaxError(function(event, request, settings){
				$(this).dialog('open');
				$("#PwdLoader").hide();
				$("#ConLoader").hide();
			});
		
			
			//*Dialogs
			$("#GlobalAjaxErrorDialog").dialog({
					bgiframe: true,
					autoOpen: false,
					resizable: false,
					width: 450,
					height: 230,
					modal: true,
					buttons: {
						' OK ': function() {
								$(this).dialog('close');
							}
					}
			});
			
			<%if pCustGuest=1 then
			Session("SFStrRedirectUrl")="prv_ViewOrder.asp?uid=" & pcv_UniqueID %>
			//*Validate Password Form
			$("#PwdForm").validate({
				rules: {
					newPass1: 
					{
						required: true
					},
					newPass2:
					{
						required: true,
						equalTo: "#newPass1"
					}
				},
				messages: {
					newPass1: {
						required: "<%=dictLanguage.Item(Session("language")&"_opc_js_4")%>",
						minlength: "<%=dictLanguage.Item(Session("language")&"_opc_js_5")%>"
					},
					newPass2: {
						required: "<%=dictLanguage.Item(Session("language")&"_opc_js_47")%>",
						minlength: "<%=dictLanguage.Item(Session("language")&"_opc_js_5")%>",
						equalTo: "<%=dictLanguage.Item(Session("language")&"_opc_js_48")%>"
					}
				}
			})
			
			$('#PwdSubmit').click(function(){
				if ($('#PwdForm').validate().form())
				{
					$("#PwdLoader").html('<img src="images/ajax-loader1.gif" width="20" height="20" align="absmiddle"><%=dictLanguage.Item(Session("language")&"_opc_common_5")%>');
					$("#PwdLoader").show();	
					$.ajax({
						type: "POST",
						url: "opc_createacc.asp",
						data: $('#PwdForm').formSerialize() + "&action=create",
						timeout: 5000,
						success: function(data, textStatus){
							if (data=="SECURITY")
							{
								$("#PwdArea").html("");
								$("#PwdArea").hide();
								$("#PwdLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"><%=dictLanguage.Item(Session("language")&"_opc_common_6")%>');
								var callbackPwd=function (){setTimeout(function(){$("#PwdLoader").hide();},1000);}
								$("#PwdLoader").effect('pulsate',{},500,callbackPwd);
							}
							else
							{
							if ((data=="OK") || (data=="REG") || (data=="OKA") || (data=="REGA"))
							{
		
								if ((data=="OK") || (data=="OKA"))
								{
									$("#PwdLoader").html('<div class="pcSuccessMessage"><%=dictLanguage.Item(Session("language")&"_opc_common_7")%></div>');
								}
								else
								{
									$("#PwdLoader").html('<div class="pcSuccessMessage"><%=dictLanguage.Item(Session("language")&"_opc_common_8")%></div>');
								}
								if (data=="OKA")
								{
									$("#PwdArea").html("");
									$("#PwdArea").hide();
								}
								else
								{
									$("#PwdArea").html("");
									$("#PwdArea").hide();
								}
							}
							else
							{
								$("#PwdLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"> '+data);
								var callbackPwd=function (){setTimeout(function(){$("#PwdLoader").hide();},1000);}
								$("#PwdLoader").effect('pulsate',{},500,callbackPwd);
							}
							}
						}
					});
					return(false);
				}
				return(false);
			});
			<%end if%>
		
		
		});
		</script>
		
			<% end if %>
			<table class="pcShowContent">
				<tr>
					<th width="15%">
						<% response.write dictLanguage.Item(Session("language")&"_orderverify_26")%>
					</th>
					<th width="50%">
						<% response.write dictLanguage.Item(Session("language")&"_orderverify_27")%>
					</th>
					<th>
						<div align="right">
						&nbsp;
						</div>
					</th>
					<th>
						<% = dictLanguage.Item(Session("language")&"_CustviewPastD_14") %>
					</th>
					<th>
						<div align="right">
						&nbsp;
						</div>
					</th>
				</tr>
                <%
				Dim pSku, pdescription, pImage, pOrderDate, pIDProduct

				do while not rsOrdObj.eof
					pdescription=rsOrdObj("description")
					pSku=rsOrdObj("sku")
					pImage = rsOrdObj("smallImageURL")
					pOrderDate = DateValue(rsOrdObj("orderdate"))
					pidProduct = rsOrdObj("idProduct")
					pActive=rsOrdObj("active")
					pRemoved=rsOrdObj("removed")

					'// Check customer eligibility to write a review
					pcv_IPAddress=Request.ServerVariables("REMOTE_ADDR")
					prv_strDenied = "0"				
					Count=0	

					query="SELECT pcRev_IDReview FROM pcReviews where pcRev_IP='" & pcv_IPAddress & "' and pcRev_IDProduct=" & pidProduct						
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=connTemp.execute(query)			
					do while not rs.eof
						Count=Count+1
						rs.MoveNext
					loop
					set rs=nothing
					
					pcv_PrdRevExc = "0" 
					query="SELECT pcRE_IDProduct FROM pcRevExc WHERE pcRE_IDProduct = " & pidProduct
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=connTemp.execute(query)			
					if not rs.eof then
                        pcv_PrdRevExc = "1"
					end if
					set rs=nothing					
										
					Count1=getUserInput(Request.Cookies("Prd" & pcv_IDProduct),0)
					if Count1="" then
						Count1=0
					end if
					
					IF (clng(Count)>=clng(pcv_PostCount)) and (pcv_LockPost="0") THEN
						prv_strDenied = "1"
					END IF
					
					IF (clng(Count1)>=clng(pcv_PostCount)) and (pcv_LockPost="1") THEN
						prv_strDenied = "1"
					END IF
					
					IF ((clng(Count)>=clng(pcv_PostCount)) or (clng(Count1)>=clng(pcv_PostCount))) and (pcv_LockPost="2") THEN
						prv_strDenied = "1"
					END IF
                    %>
					<tr> 
						<td><p><%=pSku%></p></td>
						<td><p><a href="viewPrd.asp?idproduct=<%=pidProduct%>" target="_blank"><%=pdescription%></a></p></td>
						<td><p>
						<% If Len(Trim(pImage&""))>0 Then %>
                        <a href="viewPrd.asp?idproduct=<%=pidProduct%>" target="_blank"><img src="catalog/<%=pImage%>" width="35" border="0"></a>
                        <% End if %>
                        </p></td>
						<td><p><%=pOrderDate%></p></td>
						<td><p>
							  <%
							  if (pcv_PrdRevExc = "1") OR (pActive="0") OR (pRemoved<>"0") then
							     response.write ""
                              elseIf prv_strDenied = "1" Then
                                 response.write dictLanguage.Item(Session("language")&"_prv_26")
                              else
                                 response.write "<div id=""xrv" & rsOrdObj("idProductOrdered") & """><a href=""javascript:openbrowser('prv_postreview.asp?IDProduct=" & pIDProduct & "&idcustomer="& pIdCustomer &"&xrv=" & rsOrdObj("idProductOrdered") & "');"">" & dictLanguage.Item(Session("language")&"_prv_25") & "</a></div>"
                              end if
                              %>
                    	</p></td>
				    </tr>
					<%
					rsOrdObj.MoveNext
				loop
			%>
			</table>
		<td>
	<tr>
	</table>

</div>
<% call closedb() %>
<!--#include file="footer.asp"-->
<% 'PRV41 end %>
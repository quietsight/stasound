<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

response.Buffer=true
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/validation.asp" -->
<!--#include file="opc_contentType.asp" -->
<% On Error Resume Next
dim conntemp, query, rs

Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")

Call SetContentType()

if session("idCustomer")=0 OR session("idCustomer")="" then
	response.clear
	Call SetContentType()
	response.write "SECURITY"
	response.End
end if

dim pcTempIdPayment
pcTempIdPayment=request("idPayment")

if session("GWPaymentId")="" then
	session("GWPaymentId")=getUserInput(pcTempIdPayment,0)
else
	if pcTempIdPayment<>session("GWPaymentId") AND pcTempIdPayment<>"" then
		session("GWPaymentId")=getUserInput(pcTempIdPayment,0)
	end if
end if

If request("PaymentGWSubmitted")="Go" then
	
	pAccNum=URLDecode(getUserInput(request("AccNum"),0))

	' extract real idorder (without prefix)
	pTrueOrderId=(int(session("GWOrderId"))-scpre)

	pAccNum2=pAccNum

	Session("pcSFpAccNum2")=pAccNum2
	session("AccNum-" & session("GWPaymentId"))=pAccNum2

	' save account info
	call opendb()
	query="INSERT INTO offlinepayments (idorder, idPayment, AccNum) VALUES (" &pTrueOrderId& "," &session("GWPaymentId")& ",'" &pAccNum2& "')"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
	end if
	set rs=nothing
	
	call closedb()
	
	Response.write "OK"
	session("Entered-" & session("GWPaymentId"))="1"
	session("NeedToUpdatePay")="0"
ELSE
	session("NeedToUpdatePay")="0"
	
	pcGatewayDataIdOrder=session("GWOrderID")

	call opendb()

	query="SELECT paymentDesc,idPayment,terms,CReq,CPrompt FROM PayTypes WHERE idPayment=" & session("GWPaymentId")
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
		
	pcStrPaymentDesc=rs("paymentDesc")
	pcStrTerms=rs("terms")
	pcCReq=rs("CReq")
	pcCPrompt=rs("CPrompt")
	set rs=nothing
	call closedb()
	%>
	
			<table class="pcShowContent">
				<tr class="pcSectionTitle">
					<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_GateWay_5")%></p></td>
				</tr>
				<% if len(pcCustIpAddress)>0 AND CustomerIPAlert="1" then %>
				<tr>
					<td colspan="2"><div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_6")&pcCustIpAddress%></div></td>
				</tr>
				<% end if %>
				<tr>
					<td width="25%">
						<%response.write dictLanguage.Item(Session("language")&"_paymnta_c_2")%>
					</td>
					<td width="75%">
						<%=pcStrPaymentDesc%>
					</td>
				</tr>
				<tr> 
					<td>  
						<%response.write dictLanguage.Item(Session("language")&"_paymnta_c_3")%>
					</td>
					<td>
						<%=pcStrTerms%>
					</td>
				</tr>
				<% if pcCReq=-1 then
					if session("Entered-" & session("GWPaymentId"))<>"1" then
						session("NeedToUpdatePay")="1"
					end if %>
				<tr> 
					<td><%=pcCPrompt%></td>
					<td><input type="text" name="AccNum" class="required" <%if session("Entered-" & session("GWPaymentId"))="1" then%>value="<%=session("AccNum-" & session("GWPaymentId"))%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="2">
					<input type="image" name="PaySubmit" id="PaySubmit" src="<%=RSlayout("pcLO_Update")%>" border="0" style="display:none">
					<script>
					//*Submit Pay Form
						$('#PaySubmit').click(function(){
						if ($('#PayForm').validate().form())
						{
						{
							$.ajax({
								type: "POST",
								url: "opc_paymnta_c.asp",
								data: $('#PayForm').formSerialize() + "&PaymentGWSubmitted=Go",
								timeout: 450000,
								success: function(data, textStatus){
								if (data=="SECURITY")
								{
									// Session Expired
									window.location="msg.asp?message=1";
								}
								else
								{
									if (data=="OK")
									{
										$("#PayLoader").hide();
										NeedToUpdatePay=0;
										GetOrderInfo("","#PayLoader1",0,'');
										ValidateGroup2();
									}
									else
									{
										$("#PayLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"> '+data);
										$("#PayLoader").show();
										NeedToUpdatePay=1;
										btnShow1("Error","Pay");
									}
									}
								}
				 			});
							return(false);
						}
						}
						return(false);
						});
					</script>
					<% if session("Entered-" & session("GWPaymentId"))="1" then
						session("NeedToUpdatePay")="0" %>
						<script>NeedToUpdatePay=0;</script>
					<% end if %>
				</td>
				</tr>
				<% else %>
					<input type="image" name="PaySubmit" id="PaySubmit" src="<%=RSlayout("pcLO_Update")%>" border="0" style="display:none">
					<script>
						//*Submit Pay Form
						$('#PaySubmit').click(function(){	
														   
							  $("#PayLoader").hide();
							  NeedToUpdatePay=0;
							  GetOrderInfo("","#PayLoader1",0,'');
							  ValidateGroup2();
							  return(false);
					
						});
					</script>
				<% end If %>
			</table>
<% end if
conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing %>

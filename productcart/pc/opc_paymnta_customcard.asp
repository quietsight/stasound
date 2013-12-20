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
pcTempIdPayment=getUserInput(request("idPayment"),0)

if session("GWPaymentId")="" then
	session("GWPaymentId")=pcTempIdPayment
else
	if pcTempIdPayment<>session("GWPaymentId") AND pcTempIdPayment<>"" then
		session("GWPaymentId")=pcTempIdPayment
	end if
end if

pcGatewayDataIdOrder=session("GWOrderID")

IF request("PaymentGWSubmitted")="Go" THEN
	pcIntIdCustomerCardType=getUserInput(request("idCCT"),20)

	' extract real idorder (without prefix)
	pTrueOrderId=(int(session("GWOrderId"))-scpre)
	
	call opendb()

	query="SELECT customCardRules.idCustomCardRules, customCardRules.idCustomCardType, customCardRules.ruleName, customCardRules.intruleRequired, customCardRules.intlengthOfField, customCardRules.intmaxInput FROM customCardRules WHERE (((customCardRules.idCustomCardType)="&pcIntIdCustomerCardType&")) ORDER BY customCardRules.intOrder;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
	end if
	ErrCnt=0
	pcErrMsg=""
	do until rs.eof
		pIdCCR=rs("idCustomCardRules")
		pRuleName=rs("ruleName")
		pReq=rs("intruleRequired")
		session("admin"&pIdCCR)=URLDecode(getUserInput(request("customfield"&pIdCCR),0))
		session("admin-" & session("GWPaymentId") & "-" &pIdCCR)=URLDecode(getUserInput(request("customfield"&pIdCCR),0))
		if pReq<>"0" then
			if session("admin"&pIdCCR)="" then
				ErrCnt=ErrCnt+1 
				pcErrMsg=pcErrMsg & "<li>" & pRuleName & " is a required field.</li>"
			end if
		end if
		rs.moveNext
	loop
	set rs=nothing
	call closedb()
	
	IF pcErrMsg="" THEN
	
	' save custom info
	call opendb()
	
	query="SELECT customCardRules.idCustomCardRules, customCardRules.idCustomCardType, customCardRules.ruleName, customCardRules.intruleRequired, customCardRules.intlengthOfField, customCardRules.intmaxInput FROM customCardRules WHERE (((customCardRules.idCustomCardType)="&pcIntIdCustomerCardType&")) ORDER BY customCardRules.intOrder;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
	end if
	do until rs.eof
		
		pcv_strCustomCardRules=rs("idCustomCardRules")
		pcv_strRuleName=rs("ruleName")
		pIdCCR=replace(pcv_strCustomCardRules,"'","''")
		
		'// Create an amendum to the admin order email
		if len(session("admin"&pIdCCR))>0 then
			'check if this is a credit/debit card number
			strRuleValue=ShowLastFour(session("admin"&pIdCCR))
			ammendAdminEmail = ammendAdminEmail & pcv_strRuleName & ": " & strRuleValue & vbCrLf
			strRuleValue=""
		end if
		
		pRuleName=replace(pcv_strRuleName,"'","''")
		pReq=rs("intruleRequired")
		
		pcBillingTotal=0
		
		' extract real idorder (without prefix)
		pTrueOrderId=(int(session("GWOrderId"))-scpre)
		
		query="INSERT INTO customcardOrders (idorder, idcustomCardType, idcustomCardRules, strFormValue, intOrderTotal,strRuleName) VALUES (" &pTrueOrderId& "," &pcIntIdCustomerCardType& "," &pIdCCR& ",'" &replace(session("admin"&pIdCCR),"'","''")& "'," &pcBillingTotal& ",'"&pRuleName&"')"
		set rsCCObj=server.CreateObject("ADODB.RecordSet")
		set rsCCObj=conntemp.execute(query)
		set rsCCObj=nothing
		
		rs.moveNext
	loop
	set rs=nothing
	call closedb()
	
	'save ammendum
	Session("pcSFSpecialFields")=ammendAdminEmail
	END IF
	
	IF pcErrMsg="" THEN
		Response.write "OK"
		session("NeedToUpdatePay")="0"
		session("Entered-" & session("GWPaymentId"))="1"
	ELSE
		session("Entered-" & session("GWPaymentId"))=""
		pcErrMsg="Errors when saving payment details:<ul>"&pcErrMsg&"</ul>"
		Response.write pcErrMsg
	END IF
ELSE
	if session("Entered-" & session("GWPaymentId"))<>"1" then
	session("NeedToUpdatePay")="1"
	end if
	call opendb()

	query="SELECT payTypes.paymentDesc, customCardTypes.idcustomCardType FROM payTypes INNER JOIN customCardTypes ON payTypes.paymentDesc = customCardTypes.customCardDesc WHERE (((payTypes.idPayment)="&session("GWPaymentId")&"));"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
		
	pcStrPaymentDesc=rs("paymentDesc")
	pcIntIdCustomerCardType=rs("idcustomCardType")
	
	set rs=nothing
	call closedb()
	%>
		<input type="hidden" name="idCCT" value="<%=pcIntIdCustomerCardType%> ">
		<%if session("Entered-" & session("GWPaymentId"))<>"1" then%>
		<script>NeedToUpdatePay=1;</script>
		<%end if%>
		<table class="pcShowContent">
		<tr class="pcSectionTitle">
			<td colspan="2"><%=dictLanguage.Item(Session("language")&"_GateWay_5")%></td>
		</tr>
		<% if len(pcCustIpAddress)>0 AND CustomerIPAlert="1" then %>
        <tr>
            <td colspan="2"><div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_6")&pcCustIpAddress%></div></td>
        </tr>
        <% end if %>
			<% call opendb()
			query="SELECT idCustomCardRules, idCustomCardType, ruleName, intruleRequired, intlengthOfField, intmaxInput FROM customCardRules WHERE (((idCustomCardType)="&pcIntIdCustomerCardType&")) ORDER BY intOrder;"
			set rs=Server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)
			HaveFields=0
			do until rs.eof
				HaveFields=1
				pIdCCR=rs("idCustomCardRules")
				pcIntIdCustomerCardType=rs("idCustomCardType")
				pRuleName=rs("ruleName")
				pReq=rs("intruleRequired")
				pLOF=rs("intlengthOfField")
				pMInput=rs("intmaxInput")
				if pMInput="" or pMInput="0" then
					pMInput=pLOF
				end if
				%>
				<tr> 
					<td>
						<%=pRuleName%>
					</td>
					<td>
						<input name="customField<%=pIdCCR%>" type="text" <%if session("Entered-" & session("GWPaymentId"))="1" then%>value="<%=session("admin-" & session("GWPaymentId") & "-" &pIdCCR)%>"<%end if%> <% if pReq<>"0" then %>class="required"<%end if %> size="<%=pLOF%>" maxlength="<%=pMInput%>">
					</td>
				</tr>
				<% rs.movenext
			loop
			set rs=nothing
			call closedb()
			%>
			<%if HaveFields=1 then%>
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
								url: "opc_paymnta_customcard.asp",
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
					<%if session("Entered-" & session("GWPaymentId"))="1" then
					session("NeedToUpdatePay")="0"%>
					<script>NeedToUpdatePay=0;</script>
					<%end if%>
				</td>
			</tr>
			<%end if%>
		</table>
<%END IF

conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing%>

<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

response.Buffer=true

Response.Expires = 60
Response.Expiresabsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/rc4.asp"-->
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

If request("PaymentGWSubmitted")="Go" then
	pcErrMsg=""
	pCardType=getUserInput(request("cardType"),0)
	pCardNumber=getUserInput(request("cardNumber"),0)
	session("admin-" & session("GWPaymentId") & "-pCardType")=pCardType
	session("admin-" & session("GWPaymentId") & "-pCardNumber")=pCardNumber
	session("admin-" & session("GWPaymentId") & "-expMonth")=getUserInput(request("expMonth"),0)
	session("admin-" & session("GWPaymentId") & "-expYear")=getUserInput(request("expYear"),0)

	if request("expMonth")<>"" AND request("expYear")<>"" then
		pExpiration=getUserInput(request("expMonth"),0) & "/1/" & getUserInput(request("expYear"),0)
	else
		pcErrMsg= pcErrMsg & "<li>You did not enter Card Expiration Date</li>"
	end if

	if pCardType="" then
		pcErrMsg= pcErrMsg & "<li>You did not select the Card Type</li>"
	end if

	if pCardNumber="" then
		pcErrMsg= pcErrMsg & "<li>You did not enter Card Number</li>"
	end if

	if pExpiration="" then
		pcErrMsg= pcErrMsg & "<li>You did not enter Card Expiration Date</li>"
	end if

	IF pcErrMsg="" THEN
		' validates expiration
		if not IsCreditCard(pCardNumber, pCardType) then
			pcErrMsg=pcErrMsg & "<li>" & dictLanguage.Item(Session("language")&"_paymntb_o_5") & "</li>"
		else
			if DateDiff("d", Month(Now)&"/"&Year(now), request("expMonth")&"/"&request("expYear"))<=-1 then
				pcErrMsg= pcErrMsg & "<li>" & dictLanguage.Item(Session("language")&"_paymntb_o_6") & "</li>"
			end if
		end if
	END IF

	IF pcErrMsg="" THEN

		call opendb()

		pcv_SecurityPass = pcs_GetSecureKey
		pcv_SecurityKeyID = pcs_GetKeyID

		' encrypt CC data
		pCardNumber2=enDeCrypt(pCardNumber, pcv_SecurityPass)

		' extract real idorder (without prefix)
		pTrueOrderId=cLng(session("GWOrderId"))-cLng(scpre)

		' save credit card info
		if scDB="Access" then
			query="INSERT INTO creditcards (idorder, cardType, cardNumber, expiration, seqcode, pcSecurityKeyID) VALUES (" &pTrueOrderId& ",'" &pCardType& "','" &pCardNumber2& "',#" &pExpiration& "#, 'na', "&pcv_SecurityKeyID&")"
		else
			query="INSERT INTO creditcards (idorder, cardType, cardNumber, expiration, seqcode, pcSecurityKeyID) VALUES (" &pTrueOrderId& ",'" &pCardType& "','" &pCardNumber2& "','" &pExpiration& "','na', "&pcv_SecurityKeyID&")"
		end if
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)

		if err.number<>0 then
			call LogErrorToDatabase()
		end if

		set rs=nothing

		call closedb()

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
%>
		<script language="javascript">NeedToUpdatePay=1;</script>

			<table class="pcShowContent">
			<tr class="pcSectionTitle">
				<td colspan="2"><%response.write dictLanguage.Item(Session("language")&"_GateWay_5")%></td>
			</tr>
			<% if len(pcCustIpAddress)>0 AND CustomerIPAlert="1" then %>
			<tr>
				<td colspan="2"><div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_6")&pcCustIpAddress%></div></td>
			</tr>
			<% end if %>
			<tr>
				<td width="25%">
					<%response.write dictLanguage.Item(Session("language")&"_paymnta_o_2")%>
				</td>
				<td width="75%">
					<%
					call openDb()
					set rs=server.createobject("adodb.recordset")
					query="SELECT CCcode,CCType FROM CCTypes WHERE active=-1;"
					set rs=connTemp.execute(query)
					%>
					<select name="cardType" class="required">
					<% do until rs.eof
					CCcode=rs("CCcode")
					CCType=rs("CCType")  %>
						<option value="<%=CCcode%>" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if CCcode=session("admin-" & session("GWPaymentId") & "-pCardType") then%>selected<%end if%><%end if%>><%=CCType%></option>
					<% rs.moveNext
					loop
					set rs=nothing
					%>
					</select>
				</td>
			</tr>
			<tr>
				<td>
					<%response.write dictLanguage.Item(Session("language")&"_paymnta_o_3")%>
				</td>
				<td>
					<input type="text" name="cardNumber" size="30" <%if session("Entered-" & session("GWPaymentId"))="1" then%>value="<%=session("admin-" & session("GWPaymentId") & "-pCardNumber")%>"<%end if%>>
				</td>
			</tr>
			<tr>
				<td>
					<%response.write dictLanguage.Item(Session("language")&"_paymnta_o_4")%>
				</td>
				<td>
					<%response.write dictLanguage.Item(Session("language")&"_paymnta_o_5")%>
					<select name="expMonth">
						<option value="1" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expMonth")="1" then%>selected<%end if%><%end if%>>1</option>
						<option value="2" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expMonth")="2" then%>selected<%end if%><%end if%>>2</option>
						<option value="3" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expMonth")="3" then%>selected<%end if%><%end if%>>3</option>
						<option value="4" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expMonth")="4" then%>selected<%end if%><%end if%>>4</option>
						<option value="5" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expMonth")="5" then%>selected<%end if%><%end if%>>5</option>
						<option value="6" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expMonth")="6" then%>selected<%end if%><%end if%>>6</option>
						<option value="7" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expMonth")="7" then%>selected<%end if%><%end if%>>7</option>
						<option value="8" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expMonth")="8" then%>selected<%end if%><%end if%>>8</option>
						<option value="9" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expMonth")="9" then%>selected<%end if%><%end if%>>9</option>
						<option value="10" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expMonth")="10" then%>selected<%end if%><%end if%>>10</option>
						<option value="11" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expMonth")="11" then%>selected<%end if%><%end if%>>11</option>
						<option value="12" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expMonth")="12" then%>selected<%end if%><%end if%>>12</option>
					</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<%response.write dictLanguage.Item(Session("language")&"_paymnta_o_6")%>
					<select name="ExpYear">
						<% Dim varYear
						varYear=year(now) %>
						<option value="<%=varYear%>" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expYear")&""=clng(varYear) & "" then%>selected<%end if%><%end if%>><%=varYear%></option>
						<option value="<%=varYear+1%>" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expYear")&""=clng(varYear+1) & "" then%>selected<%end if%><%end if%>><%=varYear+1%></option>
						<option value="<%=varYear+2%>" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expYear")&""=clng(varYear+2) & "" then%>selected<%end if%><%end if%>><%=varYear+2%></option>
						<option value="<%=varYear+3%>" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expYear")&""=clng(varYear+3) & "" then%>selected<%end if%><%end if%>><%=varYear+3%></option>
						<option value="<%=varYear+4%>" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expYear")&""=clng(varYear+4) & "" then%>selected<%end if%><%end if%>><%=varYear+4%></option>
						<option value="<%=varYear+5%>" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expYear")&""=clng(varYear+5) & "" then%>selected<%end if%><%end if%>><%=varYear+5%></option>
						<option value="<%=varYear+6%>" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expYear")&""=clng(varYear+6) & "" then%>selected<%end if%><%end if%>><%=varYear+6%></option>
						<option value="<%=varYear+7%>" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expYear")&""=clng(varYear+7) & "" then%>selected<%end if%><%end if%>><%=varYear+7%></option>
						<option value="<%=varYear+8%>" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expYear")&""=clng(varYear+8) & "" then%>selected<%end if%><%end if%>><%=varYear+8%></option>
						<option value="<%=varYear+9%>" <%if session("Entered-" & session("GWPaymentId"))="1" then%><%if session("admin-" & session("GWPaymentId") & "-expYear")&""=clng(varYear+9) & "" then%>selected<%end if%><%end if%>><%=varYear+9%></option>
					</select>
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<input type="image" name="PaySubmit" id="PaySubmit" src="<%=RSlayout("pcLO_Update")%>" border="0" style="display:none">
					<script language="javascript">
					//*Validate Payment Form
					var validator4=$("#PayForm").validate({
						success: function(element) {
							$(element).parent("td").children("input, textarea").addClass("success")
						},
						rules: {
							cardNumber:
							{
								required: true,
								credit_card: true,
								remote: {
										url: "opc_checkCC.asp",
										type: "POST",
										data: {
												cardType: function() {
													return $("#cardType").val();
												}
										  }
								}
							}
						},
						messages: {
							cardNumber:
							{
								required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_44"))%>",
								minlength: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_45"))%>",
								remote: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_46"))%>"
							}
						}
					});
					//*Submit Pay Form
						$('#PaySubmit').click(function(){
						if ($('#PayForm').validate().form())
						{
						{
							$.ajax({
								type: "POST",
								url: "opc_paymnta_o.asp",
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
										ValidateGroup2();
										GetOrderInfo("","#PayLoader1",0,'');

									}
									else
									{
										$("#PayLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"> '+data);
										$("#PayLoader").show();
										NeedToUpdatePay=1;
										btnShow1("Error","Pay");
										validator4.resetForm();
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
						<script language="javascript">NeedToUpdatePay=0;</script>
					<%end if%>
				</td>
			</tr>
		</table>
<% END IF %>
<% function IsCreditCard(ByRef anCardNumber, ByRef asCardType)
	Dim lsNumber		' Credit card number stripped of all spaces, dashes, etc.
	Dim lsChar			' an individual character
	Dim lnTotal			' Sum of all calculations
	Dim lnDigit			' A digit found within a credit card number
	Dim lnPosition		' identifies a character position In a String
	Dim lnSum			' Sum of calculations For a specific Set

	' Default result is False
	IsCreditCard = False

	' ====
	' Strip all characters that are Not numbers.
	' ====

	' Loop through Each character inthe card number submited
	For lnPosition = 1 To Len(anCardNumber)
		' Grab the current character
		lsChar = Mid(anCardNumber, lnPosition, 1)
		' if the character is a number, append it To our new number
		if validNum(lsChar) Then lsNumber = lsNumber & lsChar

	Next ' lnPosition

	' ====
	' The credit card number must be between 13 and 16 digits.
	' ====
	' if the length of the number is less Then 13 digits, then Exit the routine
	if Len(lsNumber) < 13 Then Exit function

	' if the length of the number is more Then 16 digits, then Exit the routine
	if Len(lsNumber) > 16 Then Exit function

	' Choose action based on Type of card
	Select Case LCase(asCardType)
		' VISA
		Case "visa", "v", "V"
			' if first digit Not 4, Exit function
			if Not Left(lsNumber, 1) = "4" Then Exit function
		' American Express
		Case "american express", "americanexpress", "american", "ax", "A"
			' if first 2 digits Not 37, Exit function
			if Not Left(lsNumber, 2) = "37" AND Not Left(lsNumber, 2) = "34" Then Exit function
		' Mastercard
		Case "mastercard", "master card", "master", "M"
			' if first digit Not 5, Exit function
			if Not Left(lsNumber, 1) = "5" Then Exit function
		' Discover
		Case "discover", "discovercard", "discover card", "D"
			' if first digit Not 6, Exit function
			if Not Left(lsNumber, 1) = "6" Then Exit function

		Case Else
	End Select ' LCase(asCardType)

	' ====
	' if the credit card number is less Then 16 digits add zeros
	' To the beginning to make it 16 digits.
	' ====
	' Continue Loop While the length of the number is less Then 16 digits
	While Not Len(lsNumber) = 16

		' Insert 0 To the beginning of the number
		lsNumber = "0" & lsNumber

	Wend ' Not Len(lsNumber) = 16

	' ====
	' Multiply Each digit of the credit card number by the corresponding digit of
	' the mask, and sum the results together.
	' ====

	' Loop through Each digit
	For lnPosition = 1 To 16

		' Parse a digit from a specified position In the number
		lnDigit = Mid(lsNumber, lnPosition, 1)

		' Determine if we multiply by:
		'	1 (Even)
		'	2 (Odd)
		' based On the position that we are reading the digit from
		lnMultiplier = 1 + (lnPosition Mod 2)

		' Calculate the sum by multiplying the digit and the Multiplier
		lnSum = lnDigit * lnMultiplier

		' (Single digits roll over To remain single. We manually have to Do this.)
		' if the Sum is 10 or more, subtract 9
		if lnSum > 9 Then lnSum = lnSum - 9

		' Add the sum To the total of all sums
		lnTotal = lnTotal + lnSum

	Next ' lnPosition

	' ====
	' Once all the results are summed divide
	' by 10, if there is no remainder Then the credit card number is valid.
	' ====
	IsCreditCard = ((lnTotal Mod 10) = 0)

End function ' IsCreditCard
' ------------------------------------------------------------------------------
conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing%>
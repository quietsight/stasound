<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Sign-up for an Endicia Account" %>
<% response.Buffer=true %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/currencyformatinc.asp" -->
<!--#include file="../includes/EndiciaFunctions.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
'Require SSL Connection
tmpHTTPs=Request.ServerVariables("HTTPS")
IF UCase(tmpHTTPs)="OFF" THEN

msg="SSL is required to display this page. The reason is that in order to activate an Endicia account, you must provide payment information, which has to be transferred securely to Endicia. You can turn on SSL in the <a href='AdminSettings.asp'>Store Settings</a>"
msgType=0%>
<br><br>
<%' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>
<br><br>
<%ELSE%>
<%'// Initialize the Prototype.js files

Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<%Dim connTemp,rs,query

pcPageName="EDC_signup.asp"

call opendb()

call GetEDCSettings()

if EDCReg="1" OR ((request("action")<>"signup") AND (request("reg")<>"1")) then
	response.redirect "EDC_manage.asp"
end if

If intResetSessions=1 Then
	Session("pcAdminAdmComments")=""
End If

'// ORIGIN ADDRESS
if Session("pcAdminFromName") = "" OR intResetSessions=1 then
	pcv_strFromName = scOriginPersonName
	if pcv_strFromName="" then
		pcv_strFromName=scShipFromName
	end if
	Session("pcAdminFromName") = pcv_strFromName
end if

if instr(scOriginPersonName, " ") then
	pcv_FromNameArry=split(scOriginPersonName, " ")
	pcv_strFromFirstName=pcv_FromNameArry(0)
	pcv_strFromLastName=pcv_FromNameArry(1)
end if

if Session("pcAdminFromFirstName") = "" OR intResetSessions=1 then
	pcv_strFromFirstName = pcv_strFromFirstName
	if pcv_strFromFirstName="" then
		pcv_strFromFirstName=pcv_strFromName
	end if
	Session("pcAdminFromFirstName") = pcv_strFromFirstName
end if

if Session("pcAdminFromLastName") = "" OR intResetSessions=1 then
	pcv_strFromLastName = pcv_strFromLastName
	Session("pcAdminFromLastName") = pcv_strFromLastName
end if

if Session("pcAdminFromFirm") = "" OR intResetSessions=1 then
	pcv_strFromFirm = scShipFromName
	if pcv_strFromFirm="" then
		pcv_strFromFirm=scOriginPersonName
	end if
	Session("pcAdminFromFirm") = pcv_strFromFirm
end if

if Session("pcAdminFromPhone") = "" OR intResetSessions=1 then
	pcv_strFromPhone = scOriginPhoneNumber
	Session("pcAdminFromPhone") = pcv_strFromPhone
end if

if Session("pcAdminSenderEMail") = "" OR intResetSessions=1 then
	pcv_strSenderEMail = scFrmEmail
	Session("pcAdminSenderEMail") = pcv_strSenderEMail
end if

'// FROM ADDRESS	
if Session("pcAdminFromAddress1") = "" OR IntResetSessions=1 then
	pcv_strFromAddress1 = scShipFromAddress1
	Session("pcAdminFromAddress1") = pcv_strFromAddress1
end if
if Session("pcAdminCCAddress1")="" then
	Session("pcAdminCCAddress1")=Session("pcAdminFromAddress1")
end if
if Session("pcAdminFromAddress2") = "" OR intResetSessions=1 then
	pcv_strFromAddress2 = scShipFromAddress2
	Session("pcAdminFromAddress2") = pcv_strFromAddress2
end if
if Session("pcAdminFromCity") = "" OR intResetSessions=1 then
	pcv_strFromCity = scShipFromCity
	Session("pcAdminFromCity") = pcv_strFromCity
end if
If Session("pcAdminCCCity")="" then
	Session("pcAdminCCCity")=Session("pcAdminFromCity")
end if
if Session("pcAdminFromState") = "" OR intResetSessions=1 then
	pcv_strFromState = scShipFromState
	Session("pcAdminFromState") = pcv_strFromState
end if
if Session("pcAdminCCState")="" then
	Session("pcAdminCCState")=Session("pcAdminFromState")
end if
if Session("pcAdminFromZip5") = "" OR intResetSessions=1 then
	pcv_strFromZip5 = scShipFromPostalCode
	Session("pcAdminFromZip5") = pcv_strFromZip5
end if
if Session("pcAdminCCZip5")="" then
	Session("pcAdminCCZip5")=Session("pcAdminFromZip5")
end if

IF request("action")="signup" then
	Session("pcAdminWebPass")=request("edcwebpass")
	Session("pcAdminWebPass1")=request("edcwebpass1")
	Session("pcAdminPassP")=request("edcpassp")
	Session("pcAdminPassP1")=request("edcpassp1")
	Session("pcAdminQues")=request("edcques")
	Session("pcAdminAnswer")=request("edcanswer")
	Session("pcAdminFromFirm")=request("edcCompany")
	Session("pcAdminFromFirstName")=request("edcFName")
	Session("pcAdminFromMidName")=request("edcMName")
	Session("pcAdminFromLastName")=request("edcLName")
	Session("pcAdminFromTitle")=request("edcTitle")
	Session("pcAdminSenderEMail")=request("edcEmail")
	Session("pcAdminFromPhone")=request("edcPhone")
	Session("pcAdminFromPhoneExt")=request("edcPhoneExt")
	Session("pcAdminFromFax")=request("edcFax")
	Session("pcAdminFromAddress1")=request("edcAddr")
	Session("pcAdminFromCity")=request("edcCity")
	Session("pcAdminFromState")=request("edcState")
	Session("pcAdminFromZip5")=request("edcZip")
	Session("pcAdminPayType")=request("edcPayType")
	Session("pcAdminCCType")=request("edcCCType")
	Session("pcAdminCC")=request("edcCC")
	Session("pcAdminCCMonth")=request("edcCCMonth")
	Session("pcAdminCCYear")=request("edcCCYear")
	Session("pcAdminCCAddress1")=request("edcCCAddr")
	Session("pcAdminCCCity")=request("edcCCCity")
	Session("pcAdminCCState")=request("edcCCState")
	Session("pcAdminCCZip5")=request("edcCCZip")
	Session("pcAdminACHNum")=request("edcACHNum")
	Session("pcAdminACHRout")=request("edcACHRout")
	
	EDC_ErrMsg=""
	EDC_SuccessMsg=""
	msg=""
	
	if Session("pcAdminWebPass")<>Session("pcAdminWebPass1") then
		msg="Your 'Web Password' and 'Web Password Confirm' values are not the same"
		msgType=0
	else
		if Session("pcAdminPassP")<>Session("pcAdminPassP1") then
			msg="Your 'Pass Phrase' and 'Pass Phrase Confirm' values are not the same"
			msgType=0
		end if
	end if
	
	if msg="" then
		tmpWebEDC=EDCURLSpc & "&method=UserSignup"
		tmpXML=""
		tmpXML="<?xml version=""1.0"" encoding=""utf-8""?>"
		tmpXML=tmpXML & "<UserSignupRequest>"
		if DeveloperTest<>"" then
		tmpXML=tmpXML & "<Test>" & DeveloperTest & "</Test>"
		end if
		if Session("pcAdminFromFirm")<>"" then
			tmpXML=tmpXML & "<CompanyName>" & Session("pcAdminFromFirm") & "</CompanyName>"
		end if
		tmpXML=tmpXML & "<FirstName>" & Session("pcAdminFromFirstName") & "</FirstName>"
		if Session("pcAdminFromMidName")<>"" then
			tmpXML=tmpXML & "<MiddleInitial>" & Session("pcAdminFromMidName") & "</MiddleInitial>"
		end if
		tmpXML=tmpXML & "<LastName>" & Session("pcAdminFromLastName") & "</LastName>"
		if Session("pcAdminFromTitle")<>"" then
			tmpXML=tmpXML & "<Title>" & Session("pcAdminFromTitle") & "</Title>"
		end if
		tmpXML=tmpXML & "<EmailAddress>" & Session("pcAdminSenderEMail") & "</EmailAddress>"
		tmpXML=tmpXML & "<EmailConfirm>" & Session("pcAdminSenderEMail") & "</EmailConfirm>"
		tmpXML=tmpXML & "<PhoneNumber>" & Session("pcAdminFromPhone") & "</PhoneNumber>"
		if Session("pcAdminFromPhoneExt")<>"" then
			tmpXML=tmpXML & "<PhoneNumberExtension>" & Session("pcAdminFromPhoneExt") & "</PhoneNumberExtension>"
		end if
		tmpXML=tmpXML & "<ICertify>Y</ICertify>"
		if Session("pcAdminFromFax")<>"" then
			tmpXML=tmpXML & "<FaxNumber>" & Session("pcAdminFromFax") & "</FaxNumber>"
		end if
		tmpXML=tmpXML & "<PhysicalAddress>" & Session("pcAdminFromAddress1") & "</PhysicalAddress>"
		tmpXML=tmpXML & "<PhysicalCity>" & Session("pcAdminFromCity") & "</PhysicalCity>"
		tmpXML=tmpXML & "<PhysicalState>" & Session("pcAdminFromState") & "</PhysicalState>"
		tmpXML=tmpXML & "<PhysicalZipCode>" & Session("pcAdminFromZip5") & "</PhysicalZipCode>"
		tmpXML=tmpXML & "<WebPassword>" & Session("pcAdminWebPass") & "</WebPassword>"
		tmpXML=tmpXML & "<PassPhrase>" & Session("pcAdminPassP") & "</PassPhrase>"
		tmpXML=tmpXML & "<ChallengeQuestion>" & Session("pcAdminQues") & "</ChallengeQuestion>"
		tmpXML=tmpXML & "<ChallengeAnswer>" & Session("pcAdminAnswer") & "</ChallengeAnswer>"
		tmpXML=tmpXML & "<BillingType>T7</BillingType>"
		tmpXML=tmpXML & "<PartnerId>" & EDCPartnerID & "</PartnerId>"
		tmpXML=tmpXML & "<ProductType>LABELSERVER</ProductType>"
		tmpXML=tmpXML & "<PaymentType>" & Session("pcAdminPayType") & "</PaymentType>"
		if Session("pcAdminPayType")="CC" then
			tmpXML=tmpXML & "<CreditCardType>" & Session("pcAdminCCType") & "</CreditCardType>"
			tmpXML=tmpXML & "<CreditCardNumber>" & Session("pcAdminCC") & "</CreditCardNumber>"
			tmpXML=tmpXML & "<CreditCardExpMonth>" & Session("pcAdminCCMonth") & "</CreditCardExpMonth>"
			tmpXML=tmpXML & "<CreditCardExpYear>" & Session("pcAdminCCYear") & "</CreditCardExpYear>"
			tmpXML=tmpXML & "<CreditCardAddress>" & Session("pcAdminCCAddress1") & "</CreditCardAddress>"
			tmpXML=tmpXML & "<CreditCardCity>" & Session("pcAdminCCCity") & "</CreditCardCity>"
			tmpXML=tmpXML & "<CreditCardState>" & Session("pcAdminCCState") & "</CreditCardState>"
			tmpXML=tmpXML & "<CreditCardZipCode>" & Session("pcAdminCCZip5")& "</CreditCardZipCode>"
		else
			tmpXML=tmpXML & "<CheckingAccountNumber>" & Session("pcAdminACHNum") & "</CheckingAccountNumber>"
			tmpXML=tmpXML & "<CheckingAccountRoutingNumber>" & Session("pcAdminACHRout") & "</CheckingAccountRoutingNumber>"
		end if
		tmpXML=tmpXML & "<OverrideEmailCheck>N</OverrideEmailCheck>"
		tmpXML=tmpXML & "</UserSignupRequest>"
		tmpXML="XMLInput=" & Server.URLEncode(tmpXML)
		tmpWebEDC=tmpWebEDC & "&" & tmpXML
		result=ConnectServer(tmpWebEDC,"GET","","","")
		IF result="ERROR" or result="TIMEOUT" THEN
			msg="Cannot connect to Endicia Label Server"
			msgType=0
		ELSE
			tmpCode=FindStatusCode(result)
			if tmpCode="0" then
				tmpErrMsg=FindXMLValue(result,"ErrorMsg")
				if tmpErrMsg="" then
					if FindXMLValue(result,"ConfirmationNumber")<>"" then
						call opendb()
						query="DELETE FROM pcEDCSettings;"
						set rsQ=connTemp.execute(query)
						set rsQ=nothing
						tmpWebPass=enDeCrypt(Session("pcAdminWebPass"), scCrypPass)
						tmpPassP=enDeCrypt(Session("pcAdminPassP"), scCrypPass)
						query="INSERT INTO pcEDCSettings (pcES_WebPass,pcES_PassP,pcES_Reg,pcES_TestMode,pcES_LogTrans) VALUES ('" & tmpWebPass & "','" & tmpPassP & "',1,1,1);"
						set rsQ=connTemp.execute(query)
						set rsQ=nothing
						EDC_SuccessMsg="Account sign-up completed successfully!"
						call closedb()
						response.redirect "EDC_manage.asp?msg=" & EDC_SuccessMsg & "&s=1" 
					else
						msg="We were not able to complete the sign-up process. Unknown Error."
						msgType=0
					end if
				else
					msg=tmpErrMsg
					msgType=0
				end if
			else
				msg=FindXMLValue(result,"ErrorMsg")
				msgType=0
			end if
		END IF
	end if
END IF
%>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<div style="margin: 10px; padding: 10px; border: 1px dashed #CCC; font-size: 13px;">
  <p><a href="https://www.endicia.com/Products/Premium/" target="_blank"><img src="images/PoweredByEndicia_small.jpg" align="right" hspace="10" border="0"></a>Please note: through this sign up form you are requesting to sign up for the <strong>Endicia </strong> service, which is provided by <a href="http://www.endicia.com/CompanyInformation/" target="_blank">Endicia</a> under the terms and conditions listed at the bottom of this page.</p>
  <ul>
    <li>This service will allow you to print USPS shipping labels from ProductCart</li>
    <li>The cost of the service is <strong>$15.95/month</strong>. NOTE: this is the cost of the service. All postage is purchased separately and is not included in the cost of the service.</li>
    <li>You can cancel the service at any time.</li>
    <li>You will not be charged until the 2nd month of use (so if you cancel in the first 30 days, there will be no charge other than the charges associated with the postage that you will purchase for your shipments).</li>
    <li> <a href="http://wiki.earlyimpact.com/productcart/shipping-usps#integration_with_endicia" target="_blank">More information</a> on this service.</li>
  </ul>
</div>
<form name="form1" method="post" action="<%=pcPageName%>?action=signup" onsubmit="javascript: if (checkForm(this)) {pcf_Open_EndiciaPop();return(true);} else {return(false);}" class="pcForms">
<table class="pcCPcontent">
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">Endicia Log-in Information:</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr valign="top">
	<td width="30%">Web Password:</td>
	<td width="70%">
		<input type="password" name="edcwebpass" value="" size="40"> <img src="images/sample/pc_icon_required.gif" border="0">
		<div class="pcSmallText" style="padding-top: 5px;">You will use this to login to your account at <a href="https://www.endicia.com" target="_blank">endicia.com</a></div>
	</td>
</tr>
<tr valign="top">
	<td>Web Password Confirm:</td>
	<td>
		<input type="password" name="edcwebpass1" value="" size="40"> <img src="images/sample/pc_icon_required.gif" border="0">
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr valign="top">
	<td>Forgot Password Question:</td>
	<td>
		<input type="text" name="edcques" value="<%=Session("pcAdminQues")%>" size="40"> <img src="images/sample/pc_icon_required.gif" border="0">
	</td>
</tr>
<tr valign="top">
	<td>Forgot Password Answer:</td>
	<td>
		<input type="text" name="edcanswer" value="<%=Session("pcAdminAnswer")%>" size="40"> <img src="images/sample/pc_icon_required.gif" border="0">
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr valign="top">
	<td>Pass Phrase: <div class="pcSmallText" style="padding-top: 5px;">Use a long alphanumeric key</div>
</td>
	<td>
		<input type="password" name="edcpassp" value="" size="40"> <img src="images/sample/pc_icon_required.gif" border="0">
		<div class="pcSmallText" style="padding-top: 5px;">This will be used to create secure transactions between your store and the Endicia Label Server</div>
	</td>
</tr>
<tr valign="top">
	<td>Pass Phrase Confirm:</td>
	<td>
		<input type="password" name="edcpassp1" value="" size="40"> <img src="images/sample/pc_icon_required.gif" border="0">
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2"><b>Note:</b> Your will receive your <strong>Endicia Account ID</strong> in an e-mail confirmation after signing up. To complete the sign-up, you will need the Pass Phrase that you are entering here, so make sure to write it down.</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">Customer Information:</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr valign="top">
	<td>Company Name:</td>
	<td><input type="text" name="edcCompany" value="<%=Session("pcAdminFromFirm")%>" size="40"></td>
</tr>
<tr valign="top">
	<td>First Name:</td>
	<td><input type="text" name="edcFName" value="<%=Session("pcAdminFromFirstName")%>" size="40"> <img src="images/sample/pc_icon_required.gif" border="0"></td>
</tr>
<tr valign="top">
	<td>Middle Name:</td>
	<td><input type="text" name="edcMName" value="<%=Session("pcAdminFromMidName")%>" size="40"></td>
</tr>
<tr valign="top">
	<td>Last Name:</td>
	<td><input type="text" name="edcLName" value="<%=Session("pcAdminFromLastName")%>" size="40"> <img src="images/sample/pc_icon_required.gif" border="0"></td>
</tr>
<tr valign="top">
	<td>Title:</td>
	<td><input type="text" name="edcTitle" value="<%=Session("pcAdminFromTitle")%>" size="40"></td>
</tr>
<tr valign="top">
	<td>E-mail:</td>
	<td><input type="text" name="edcEmail" value="<%=Session("pcAdminSenderEMail")%>" size="40"> <img src="images/sample/pc_icon_required.gif" border="0"></td>
</tr>
<tr valign="top">
	<td>Phone:</td>
	<td><input type="text" name="edcPhone" value="<%=Session("pcAdminFromPhone")%>" size="40"> <img src="images/sample/pc_icon_required.gif" border="0"> Ext. <input type="text" name="edcPhoneExt" value="<%=Session("pcAdminFromPhoneExt")%>" size="5"></td>
</tr>
<tr valign="top">
	<td>Fax:</td>
	<td><input type="text" name="edcFax" value="<%=Session("pcAdminFromFax")%>" size="40"></td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">Billing Address:</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr valign="top">
	<td>Address:</td>
	<td><input type="text" name="edcAddr" value="<%=Session("pcAdminFromAddress1")%>" size="40"> <img src="images/sample/pc_icon_required.gif" border="0"></td>
</tr>
<tr valign="top">
	<td>City:</td>
	<td><input type="text" name="edcCity" value="<%=Session("pcAdminFromCity")%>" size="40"> <img src="images/sample/pc_icon_required.gif" border="0"></td>
</tr>
<% err.clear
err.number=0
dim rsStateObj, FromStateOptAry, ToStateOptAry
FromStateOptAry=""
FromStateOptAry1=""
call opendb()
query="SELECT stateCode, stateName FROM states WHERE pcCountryCode='US' ORDER BY stateName;"
set rsStateObj=server.CreateObject("ADODB.RecordSet")
set rsStateObj=conntemp.execute(query)
if NOT rsStateObj.eof then
	do until rsStateObj.eof
		strTempStateCode=rsStateObj("stateCode") 
		if Session("pcAdminFromState")=strTempStateCode then
			strSelectedValue="selected"
		else
			strSelectedValue=""
		end if
		FromStateOptAry=FromStateOptAry&"<option value="""&strTempStateCode&""" "&strSelectedValue&">"&rsStateObj("stateName")&"</option>"
		if Session("pcAdminCCState")=strTempStateCode then
			strSelectedValue1="selected"
		else
			strSelectedValue1=""
		end if
		FromStateOptAry1=FromStateOptAry1&"<option value="""&strTempStateCode&""" "&strSelectedValue1&">"&rsStateObj("stateName")&"</option>"
		rsStateObj.moveNext
	loop
end if
set rsStateObj=nothing%>	
<tr>
	<td>State:</td>
	<td>
		<select name="edcState" id="edcState">
			<%=FromStateOptAry%>
		</select> <img src="images/sample/pc_icon_required.gif" border="0">
	</td>
</tr>
<tr valign="top">
	<td>Postal Code:</td>
	<td><input type="text" name="edcZip" value="<%=Session("pcAdminFromZip5")%>" size="5"> <img src="images/sample/pc_icon_required.gif" border="0"></td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">Payment Information</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr valign="top">
	<td>Select a Payment Type:</td>
	<td>
		<script>
			function TurnOnOffOptions(tmpValue)
			{
				if (tmpValue=="CC")
				{
					document.getElementById("CCTable").style.display='';
					document.getElementById("ACHTable").style.display='none';
				}
				else
				{
					document.getElementById("CCTable").style.display='none';
					document.getElementById("ACHTable").style.display='';
				}
			}
		</script>
		<%if Session("pcAdminPayType")="" then
			Session("pcAdminPayType")="CC"
		end if%>
		<select name="edcPayType" id="edcPayType" onChange="javascript:TurnOnOffOptions(this.value);">
			<option value="CC" <%if Session("pcAdminPayType")="CC" then%>selected<%end if%>>Credit Card</option>
			<option value="ACH" <%if Session("pcAdminPayType")="ACH" then%>selected<%end if%>>Checking Account</option>
		</select> <img src="images/sample/pc_icon_required.gif" border="0">
	</td>
</tr>
<tr>
	<td colspan="2">
		<table id="CCTable" class="pcCPcontent" <%if Session("pcAdminPayType")<>"CC" then%>style="display:none"<%end if%>>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2"><h2>Credit Card Information</h2></td>
		</tr>
		<tr valign="top">
			<td width="30%" nowrap>Credit Card Type:</td>
			<td width="70%">
				<%if Session("pcAdminCCType")="" then
					Session("pcAdminCCType")="V"
				end if%>
				<select name="edcCCType">
					<option value="V" <%if Session("pcAdminCCType")="V" then%>selected<%end if%>>Visa Card</option>
					<option value="M" <%if Session("pcAdminCCType")="M" then%>selected<%end if%>>Master Card</option>
					<option value="A" <%if Session("pcAdminCCType")="A" then%>selected<%end if%>>American Express</option>
					<option value="B" <%if Session("pcAdminCCType")="B" then%>selected<%end if%>>Carte Blanche</option>
					<option value="N" <%if Session("pcAdminCCType")="N" then%>selected<%end if%>>Discover/Novis</option>
					<option value="D" <%if Session("pcAdminCCType")="D" then%>selected<%end if%>>Diners Club</option>
				</select> <img src="images/sample/pc_icon_required.gif" border="0">
			</td>
		</tr>
		<tr valign="top">
			<td nowrap>Credit Card Number:</td>
			<td><input type="text" name="edcCC" value="<%=Session("pcAdminCC")%>" size="40"> <img src="images/sample/pc_icon_required.gif" border="0"></td>
		</tr>
		<tr valign="top">
			<td>Expiration Date</td>
			<td>
				<%if Session("pcAdminCCMonth")="" then
					Session("pcAdminCCMonth")=Month(Date())
					if len(Session("pcAdminCCMonth"))=1 then
						Session("pcAdminCCMonth")="0" & Session("pcAdminCCMonth")
					end if
				end if%>
				Month: <select name="edcCCMonth">
				<option "01" <%if Session("pcAdminCCMonth")="01" then%>selected<%end if%>>01</option>
				<option "02" <%if Session("pcAdminCCMonth")="02" then%>selected<%end if%>>02</option>
				<option "03" <%if Session("pcAdminCCMonth")="03" then%>selected<%end if%>>03</option>
				<option "04" <%if Session("pcAdminCCMonth")="04" then%>selected<%end if%>>04</option>
				<option "05" <%if Session("pcAdminCCMonth")="05" then%>selected<%end if%>>05</option>
				<option "06" <%if Session("pcAdminCCMonth")="06" then%>selected<%end if%>>06</option>
				<option "07" <%if Session("pcAdminCCMonth")="07" then%>selected<%end if%>>07</option>
				<option "08" <%if Session("pcAdminCCMonth")="08" then%>selected<%end if%>>08</option>
				<option "09" <%if Session("pcAdminCCMonth")="09" then%>selected<%end if%>>09</option>
				<option "10" <%if Session("pcAdminCCMonth")="10" then%>selected<%end if%>>10</option>
				<option "11" <%if Session("pcAdminCCMonth")="11" then%>selected<%end if%>>11</option>
				<option "12" <%if Session("pcAdminCCMonth")="12" then%>selected<%end if%>>12</option>
				</select>  <img src="images/sample/pc_icon_required.gif" border="0">
				<%if Session("pcAdminCCYear")="" then
					Session("pcAdminCCYear")=Year(Date())
				end if%>
				&nbsp;&nbsp;Year: <select name="edcCCYear">
				<option "<%=Year(Date())%>" <%if Clng(Session("pcAdminCCYear"))=Clng(Year(Date())) then%>selected<%end if%>><%=Year(Date())%></option>
				<option "<%=Year(Date())+1%>" <%if Clng(Session("pcAdminCCYear"))=Clng(Year(Date())+1) then%>selected<%end if%>><%=Year(Date())+1%></option>
				<option "<%=Year(Date())+2%>" <%if Clng(Session("pcAdminCCYear"))=Clng(Year(Date())+2) then%>selected<%end if%>><%=Year(Date())+2%></option>
				<option "<%=Year(Date())+3%>" <%if Clng(Session("pcAdminCCYear"))=Clng(Year(Date())+3) then%>selected<%end if%>><%=Year(Date())+3%></option>
				<option "<%=Year(Date())+4%>" <%if Clng(Session("pcAdminCCYear"))=Clng(Year(Date())+4) then%>selected<%end if%>><%=Year(Date())+4%></option>
				<option "<%=Year(Date())+5%>" <%if Clng(Session("pcAdminCCYear"))=Clng(Year(Date())+5) then%>selected<%end if%>><%=Year(Date())+5%></option>
				<option "<%=Year(Date())+6%>" <%if Clng(Session("pcAdminCCYear"))=Clng(Year(Date())+6) then%>selected<%end if%>><%=Year(Date())+6%></option>
				<option "<%=Year(Date())+7%>" <%if Clng(Session("pcAdminCCYear"))=Clng(Year(Date())+7) then%>selected<%end if%>><%=Year(Date())+7%></option>
				<option "<%=Year(Date())+8%>" <%if Clng(Session("pcAdminCCYear"))=Clng(Year(Date())+8) then%>selected<%end if%>><%=Year(Date())+8%></option>
				<option "<%=Year(Date())+9%>" <%if Clng(Session("pcAdminCCYear"))=Clng(Year(Date())+9) then%>selected<%end if%>><%=Year(Date())+9%></option>
				<option "<%=Year(Date())+10%>" <%if Clng(Session("pcAdminCCYear"))=Clng(Year(Date())+10) then%>selected<%end if%>><%=Year(Date())+10%></option>
				</select>  <img src="images/sample/pc_icon_required.gif" border="0">
			</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2"><b>Billing Address</b> - <a href="javascript:copybill();">Click here</a> to copy from Billing Address
				<script>
					function copybill()
					{
					document.form1.edcCCAddr.value=document.form1.edcAddr.value;
					document.form1.edcCCCity.value=document.form1.edcCity.value;
					document.form1.edcCCState.value=document.form1.edcState.value;
					document.form1.edcCCZip.value=document.form1.edcZip.value;
					}
				</script>
            </td>
		</tr>

		<tr valign="top">
			<td>Address:</td>
			<td><input type="text" name="edcCCAddr" value="<%=Session("pcAdminCCAddress1")%>" size="40"> <img src="images/sample/pc_icon_required.gif" border="0"></td>
		</tr>
		<tr valign="top">
			<td>City:</td>
			<td><input type="text" name="edcCCCity" value="<%=Session("pcAdminCCCity")%>" size="40"> <img src="images/sample/pc_icon_required.gif" border="0"></td>
		</tr>
		<tr>
		<td>State:</td>
		<td>
			<select name="edcCCState" id="edcCCState">
			<%=FromStateOptAry1%>
			</select> <img src="images/sample/pc_icon_required.gif" border="0">
		</td>
		</tr>
		<tr valign="top">
			<td>Postal Code:</td>
			<td><input type="text" name="edcCCZip" value="<%=Session("pcAdminCCZip5")%>" size="5"> <img src="images/sample/pc_icon_required.gif" border="0"></td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>	
		</table>
		<table id="ACHTable" class="pcCPcontent" <%if Session("pcAdminPayType")<>"ACH" then%>style="display:none"<%end if%>>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2"><h2>Checking Account Information</h2></td>
		</tr>
		<tr valign="top">
			<td>Account Number:</td>
			<td><input type="text" name="edcACHNum" value="<%=Session("pcAdminACHNum")%>" size="40"> <img src="images/sample/pc_icon_required.gif" border="0"></td>
		</tr>
		<tr valign="top">
			<td>Routing Number:</td>
			<td><input type="text" name="edcACHRout" value="<%=Session("pcAdminACHRout")%>" size="40"> <img src="images/sample/pc_icon_required.gif" border="0"></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td colspan="2"><hr /></td>
</tr>
<tr>
	<td colspan="2">
    	<table class="pcCPcontent">
            <tr>
                <td valign="top" align="right" width="5%"><input type="checkbox" name="edcAgree" value="1" class="clearBorder"></td>
                <td width="95%">I certify that all information provided above is accurate and truthful. I also certify that I have read and understood the <a href="http://www.usps.com/privacyoffice/privacypolicy.htm#privacyact" target="_blank">United States Postal Service Privacy Act Statement</a>, <a href="http://www.usps.com/cpim/ftp/hand/as353/as353apdx_051.htm" target="_blank">PC Postage Privacy Principles</a>, <a href="https://www.endicia.com/SignUp/Notice/TermsAndConditions/vSmall.cfm" target="_blank">Endicia Terms and Conditions</a>, and <a href="https://www.endicia.com//SignUp/Notice/USPSRegardingShortpaidandUnpaidPostage/vSmall.cfm" target="_blank">USPS Policy Regarding Shortpaid and Unpaid Postage</a>.</td>
            </tr>
         </table>
     </td>
</tr>
<tr>
	<td colspan="2"><hr /></td>
</tr>
<tr>
	<td colspan="2" align="center"><input type="submit" name="submit1" value=" Sign-up " class="submit2"></td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
</table>
</form>
<script>
	function isDigit(s)
	{
		var test=""+s;
		if(test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
		{
			return(true);
		}
		return(false);
	}
	
	function allDigit(s)
	{
		var test=""+s ;
		for (var k=0; k <test.length; k++)
		{
			var c=test.substring(k,k+1);
			if (isDigit(c)==false)
			{
				return (false);
			}
		}
		return (true);
	}

	function checkForm(tmpForm)
	{
		if (tmpForm.edcwebpass.value=="")
		{
			alert("Please enter a value for 'Web Password' field");
			tmpForm.edcwebpass.focus();
			return(false);
		}
		if (tmpForm.edcwebpass1.value=="")
		{
			alert("Please enter a value for 'Web Password Confirm' field");
			tmpForm.edcwebpass1.focus();
			return(false);
		}
		if (tmpForm.edcwebpass1.value!=tmpForm.edcwebpass.value)
		{
			alert("'Web Password' and 'Web Password Confirm' values are not the same");
			tmpForm.edcwebpass1.focus();
			return(false);
		}
		if (tmpForm.edcpassp.value=="")
		{
			alert("Please enter a value for 'Pass Phrase' field");
			tmpForm.edcpassp.focus();
			return(false);
		}
		if (tmpForm.edcpassp1.value=="")
		{
			alert("Please enter a value for 'Pass Phrase Confirm' field");
			tmpForm.edcpassp1.focus();
			return(false);
		}
		if (tmpForm.edcpassp1.value!=tmpForm.edcpassp.value)
		{
			alert("'Pass Phrase' and 'Pass Phrase Confirm' values are not the same");
			tmpForm.edcpassp1.focus();
			return(false);
		}
		
		if (tmpForm.edcques.value=="")
		{
			alert("Please enter a value for 'Forgot Password Question' field");
			tmpForm.edcques.focus();
			return(false);
		}
		
		if (tmpForm.edcanswer.value=="")
		{
			alert("Please enter a value for 'Forgot Password Answer' field");
			tmpForm.edcanswer.focus();
			return(false);
		}
		
		if (tmpForm.edcFName.value=="")
		{
			alert("Please enter a value for 'First Name' field");
			tmpForm.edcFName.focus();
			return(false);
		}
		
		if (tmpForm.edcLName.value=="")
		{
			alert("Please enter a value for 'Last Name' field");
			tmpForm.edcLName.focus();
			return(false);
		}
		
		if (tmpForm.edcEmail.value=="")
		{
			alert("Please enter a value for 'E-mail' field");
			tmpForm.edcEmail.focus();
			return(false);
		}
		
		if (tmpForm.edcAddr.value=="")
		{
			alert("Please enter a value for 'Billing Address' field");
			tmpForm.edcAddr.focus();
			return(false);
		}
		
		if (tmpForm.edcCity.value=="")
		{
			alert("Please enter a value for 'Billing City' field");
			tmpForm.edcCity.focus();
			return(false);
		}
		
		if (tmpForm.edcZip.value=="")
		{
			alert("Please enter a value for 'Postal Code' field");
			tmpForm.edcZip.focus();
			return(false);
		}
		
		if (tmpForm.edcAgree.checked == false)
		{
			alert("In order to continue, you must certify that all information provided above is accurate and truthful, and that you read and understood the United States Postal Service Privacy Act Statement, PC Postage Privacy Principles, Endicia Terms and Conditions, and USPS Policy Regarding Shortpaid and Unpaid Postage.");
			tmpForm.edcAgree.focus();
			return(false);
		}
		
		if (tmpForm.edcPayType.value=="CC")
		{
			if (tmpForm.edcCC.value=="")
			{
				alert("Please enter a value for 'Credit Card Number' field");
				tmpForm.edcCC.focus();
				return(false);
			}
			if (allDigit(tmpForm.edcCC.value) == false)
			{
				alert("Please enter a valid number without spaces for 'Credit Card Number' field");
				tmpForm.edcCC.focus();
				return(false);
			}
			if (tmpForm.edcCCAddr.value=="")
			{
				alert("Please enter a value for 'Credit Card Address' field");
				tmpForm.edcCCAddr.focus();
				return(false);
			}
		
			if (tmpForm.edcCCCity.value=="")
			{
				alert("Please enter a value for 'Credit Card City' field");
				tmpForm.edcCCCity.focus();
				return(false);
			}
		
			if (tmpForm.edcCCZip.value=="")
			{
				alert("Please enter a value for 'Credit Card Postal Code' field");
				tmpForm.edcCCZip.focus();
				return(false);
			}
		}
		else
		{
			if (tmpForm.edcACHNum.value=="")
			{
				alert("Please enter a value for 'Account Number' field");
				tmpForm.edcACHNum.focus();
				return(false);
			}
		
			if (tmpForm.edcACHRout.value=="")
			{
				alert("Please enter a value for 'Routing Number' field");
				tmpForm.edcACHRout.focus();
				return(false);
			}
		}
		return(true);
	}
</script>
<%call closedb()%>
<%Response.write(pcf_ModalWindow("Connecting to Endicia Label Server... ","EndiciaPop", 300))%>
<%END IF 'SSL require%>
<!--#include file="AdminFooter.asp"-->
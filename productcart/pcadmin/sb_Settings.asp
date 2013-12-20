<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
Dim pageTitle, pageName, pageIcon, Section
pageTitle="SubscriptionBridge Integration Settings"
pageName="sb_Settings.asp"
pageIcon="pcv4_icon_sb.png"
Section="SB" 
%>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<!--#include file="sb_inc.asp"-->
<% dim conntemp, rs, query, paySubflag, SB_On

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// START: ON SUBMIT
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if request("SubmitSettings")<>"" then


	'/////////////////////////////////////////////////////
	'// START:  Get Gateways
	'/////////////////////////////////////////////////////
	SB_On = "0"
	
	Dim SB_GatewayCodeArr(4)
	Dim SB_GatewaySubArr(4)
	 
	SB_GatewayCodeArr(0) = "1"
	SB_GatewaySubArr(0) = "0" 
	SB_GatewayCodeArr(1) = "16"
	SB_GatewaySubArr(1) = "0"
	SB_GatewayCodeArr(2) = "999999"
	SB_GatewaySubArr(2) = "0"
	SB_GatewayCodeArr(3) = "46"
	SB_GatewaySubArr(3) = "0"
	SB_GatewayCodeArr(4) = "67"
	SB_GatewaySubArr(4) = "0"

	SB_PaymentType=request("SB_PaymentType")

	if SB_PaymentType = "SB_Authorize" Then 	  
	  SB_GatewaySubArr(0) = "1"
	End if 
	
	if SB_PaymentType = "SB_AuthorizeCheck" Then 	   
	   SB_GatewaySubArr(1) = "1"
	End if 
	
	if SB_PaymentType = "SB_EC" Then 	   
	   SB_GatewaySubArr(2) = "1"
	End if 
	
	if SB_PaymentType = "SB_WPP" Then 	   
	   SB_GatewaySubArr(3) = "1"
	End if 
	
	if SB_PaymentType = "SB_EIG" Then 	   
	   SB_GatewaySubArr(4) = "1"
	End if 
	
	call openDb()

	For i = 0 to 4
	  query="UPDATE payTypes SET pcPayTypes_Subscription=" & SB_GatewaySubArr(i) & " WHERE gwCode=" & SB_GatewayCodeArr(i) & ";"
	  set rs=Server.CreateObject("ADODB.Recordset")
	  set rs=conntemp.execute(query)
	  ' make sure a payment method is available
	  if SB_GatewaySubArr(i) ="1" Then
		SB_on = "1" 
	  end if  
	Next

	set rs=nothing
	call closeDb()
	
	'/////////////////////////////////////////////////////
	'// END:  Get Gateways
	'/////////////////////////////////////////////////////



	'/////////////////////////////////////////////////////
	'// START:  Form Values
	'/////////////////////////////////////////////////////
  	
	pcv_Status=request.Form("SB_Status")
	if pcv_Status="" or pcv_Status="0" then
		pcv_Status="0"
	else
	  if SB_on ="0" Then	
	    pcv_Status="0"
		msg="Recurring Billing must have at least one payment method." 	 
	  End if
	end if	
	pcv_ShowStartDate=request.Form("ShowStartDate")
	if pcv_ShowStartDate="" then
		pcv_ShowStartDate="2"
	end if	
	pcv_StartDateDesc=request.Form("StartDateDesc")
	pcv_ShowReoccurenceDate=request.Form("ShowReoccurenceDate")
	if pcv_ShowReoccurenceDate="" then
		pcv_ShowReoccurenceDate="2"
	end if	
	pcv_ReoccurenceDesc=request.Form("ReoccurenceDesc")
	pcv_ShowEOSDate=request.Form("ShowEOSDate")
	if pcv_ShowEOSDate="" then
		pcv_ShowEOSDate="2"
	end if	
	pcv_EOSDesc=request.Form("EOSDesc")
	pcv_ShowTrialDate=request.Form("ShowTrialDate")
	if pcv_ShowTrialDate="" then
		pcv_ShowTrialDate="2"
	end if	
	pcv_TrialDate=request.Form("TrialDate")
	pcv_ShowTrialPrice=request.Form("ShowTrialPrice")
	if pcv_ShowTrialPrice="" then
		pcv_ShowTrialPrice="2"
	end if
	pcv_TrialDesc=request.Form("TrialDesc")	
	pcv_ShowFreeTrial=request.Form("ShowFreeTrial")
	if pcv_FreeShowTrial="" then
		pcv_FreeShowTrial="2"
	end if
	pcv_FreeTrialDesc=request.Form("FreeTrialDesc")		
	pcv_ShowInstallment=request.Form("ShowInstallment")
	if pcv_ShowInstallment="" then
		pcv_ShowInstallment="2"
	end if	
	pcv_InstallmentDesc=request.Form("InstallmentDesc")	
	pcv_SBRegAgree=request.Form("SB_RegAgree")
	if pcv_SBRegAgree="" then
		pcv_SBRegAgree="2"
	end if

	pcs_ValidateHTMLField	"SB_PaymentPageText", False, 0
	pcv_PaymentPageText = Session("pcAdminSB_PaymentPageText")
	pcv_PaymentPageText=Replace(pcv_PaymentPageText, vbCrLf, "<br>")
	
	pcs_ValidateHTMLField	"SB_PaymentPageTrialText", False, 0
	pcv_PaymentPageTrialText = Session("pcAdminSB_PaymentPageTrialText")
	pcv_PaymentPageTrialText=Replace(pcv_PaymentPageTrialText, vbCrLf, "<br>")	
	
	pcv_SuccessPaymentEmail=replace(request.form("SB_SuccessPaymentEmail"),"""","&quot;")
	pcv_SuccessPaymentEmail=replace(pcv_SuccessPaymentEmail,"'","''")
	pcv_SuccessPaymentEmail=replace(pcv_SuccessPaymentEmail, vbCrLf, "<br>")
	pcv_PendingPaymentEmail=replace(request.form("SB_PendingPaymentEmail"),"""","&quot;")
	pcv_PendingPaymentEmail=replace(pcv_PendingPaymentEmail,"'","''")
	pcv_PendingPaymentEmail=replace(pcv_PendingPaymentEmail, vbCrLf, "<br>")
	pcv_UnSuccessPaymentEmail=replace(request.form("SB_UnSuccessPaymentEmail"),"""","&quot;")
	pcv_UnSuccessPaymentEmail=replace(pcv_UnSuccessPaymentEmail,"'","''")
	pcv_UnSuccessPaymentEmail=replace(pcv_UnSuccessPaymentEmail, vbCrLf, "<br>")
	pcv_CCExpireEmail=replace(request.form("SB_CCExpireEmail"),"""","&quot;")
	pcv_CCExpireEmail=replace(pcv_CCExpireEmail,"'","''")
	pcv_CCExpireEmail=replace(pcv_CCExpireEmail, vbCrLf, "<br>")
	pcv_EOTFirstPaymentEmail=replace(request.form("SB_EOTFirstPaymentEmail"),"""","&quot;")
	pcv_EOTFirstPaymentEmail=replace(pcv_EOTFirstPaymentEmail,"'","''")
	pcv_EOTFirstPaymentEmail=replace(pcv_EOTFirstPaymentEmail, vbCrLf, "<br>")

	pcs_ValidateHTMLField	"SB_OffMsgText", False, 0
	pcv_SBOffMsgText = Session("pcAdminSB_OffMsgText")
	pcv_SBOffMsgText=Replace(pcv_SBOffMsgText, vbCrLf, "<br>")	
	
	pcs_ValidateHTMLField	"SB_AgreeText", False, 0
	pcv_SBAgreeText = Session("pcAdminSB_AgreeText")
	pcv_SBAgreeText=Replace(pcv_SBAgreeText, vbCrLf, "<br>")	
	
	'// New Text Settings
	pcs_ValidateHTMLField	"SB_Lang1", False, 0
	pcv_SBLang1 = Session("pcAdminSB_Lang1")
	pcv_SBLang1=Replace(pcv_SBLang1, vbCrLf, "<br>")
	
	pcs_ValidateHTMLField	"SB_Lang2", False, 0
	pcv_SBLang2 = Session("pcAdminSB_Lang2")
	pcv_SBLang2=Replace(pcv_SBLang2, vbCrLf, "<br>")
	
	pcs_ValidateHTMLField	"SB_Lang3", False, 0
	pcv_SBLang3 = Session("pcAdminSB_Lang3")
	pcv_SBLang3=Replace(pcv_SBLang3, vbCrLf, "<br>")
	
	pcs_ValidateHTMLField	"SB_Lang4", False, 0
	pcv_SBLang4 = Session("pcAdminSB_Lang4")
	pcv_SBLang4=Replace(pcv_SBLang4, vbCrLf, "<br>")
	
	pcs_ValidateHTMLField	"SB_Lang5", False, 0
	pcv_SBLang5 = Session("pcAdminSB_Lang5")
	pcv_SBLang5=Replace(pcv_SBLang5, vbCrLf, "<br>")
	
	pcs_ValidateHTMLField	"SB_Lang6", False, 0
	pcv_SBLang6 = Session("pcAdminSB_Lang6")
	pcv_SBLang6=Replace(pcv_SBLang6, vbCrLf, "<br>")
	
	pcs_ValidateHTMLField	"SB_Lang7", False, 0
	pcv_SBLang7 = Session("pcAdminSB_Lang7")
	pcv_SBLang7=Replace(pcv_SBLang7, vbCrLf, "<br>")
	
	pcs_ValidateHTMLField	"SB_Lang8", False, 0
	pcv_SBLang8 = Session("pcAdminSB_Lang8")
	pcv_SBLang8=Replace(pcv_SBLang8, vbCrLf, "<br>")
	
	pcs_ValidateHTMLField	"SB_Lang9", False, 0
	pcv_SBLang9 = Session("pcAdminSB_Lang9")
	pcv_SBLang9=Replace(pcv_SBLang9, vbCrLf, "<br>")
	
	pcs_ValidateHTMLField	"SB_Lang10", False, 0
	pcv_SBLang10 = Session("pcAdminSB_Lang10")
	pcv_SBLang10=Replace(pcv_SBLang10, vbCrLf, "<br>")
	
	If len(pcv_SBLang1)=0 Then pcv_SBLang1 = "Terms & Conditions"
	If len(pcv_SBLang2)=0 Then pcv_SBLang2 = "Please read and Agree to the Terms (below)"
	If len(pcv_SBLang3)=0 Then pcv_SBLang3 = "I Agree"
	If len(pcv_SBLang4)=0 Then pcv_SBLang4 = "Review & Agree to Proceed"
	If len(pcv_SBLang5)=0 Then pcv_SBLang5 = "The shopping cart is currently in use for purchasing a subscription.<br /><br />At this time, you can not add additional products to the cart. If you would like to purchase additional items, please complete the current order and then place a new one (this can be done very quickly as your customer information will have already be entered).<br /><br />Alternatively, you can empty the shopping cart and add different products.<br /><br /><a href=viewCart.asp><b>View shopping cart</b></a>"
	If len(pcv_SBLang6)=0 Then pcv_SBLang6 = "The shopping cart is currently in use.<br /><br />At this time, you can not add a subscription product to the cart. If you would like to purchase subscription items, please first complete the current order and then place a new one (this can be done very quickly as your customer information will have already be entered).<br /><br />Alternatively, you can empty the shopping cart and add different products.<br /><br /><a href=viewCart.asp><b>View shopping cart</b></a>"
	If len(pcv_SBLang7)=0 Then pcv_SBLang7 = "Pay Now:  "
	If len(pcv_SBLang8)=0 Then pcv_SBLang8 = "Terms:  "
	If len(pcv_SBLang9)=0 Then pcv_SBLang9 = "Disclaimer:  "
	If len(pcv_SBLang10)=0 Then pcv_SBLang10 = "Trial Disclaimer:  "
	
	'/////////////////////////////////////////////////////
	'// END:  Form Values
	'/////////////////////////////////////////////////////



	'/////////////////////////////////////////////////////
	'// START:  Write all changes to pcSBSettings.asp file
	'/////////////////////////////////////////////////////
	Dim objFS
	Dim objFile

	function removeReplaceSQ(myString)
		if isNULL(myString) then
			removeReplaceSQ=""
		else
			removeReplaceSQ=replace(myString,"''","'")
			removeReplaceSQ=replace(removeReplaceSQ,"'","''")
		end if
	end function

	function removeSQ(myString)
		if isNULL(myString) then
			removeSQ=""
		else
			myString=replace(myString,"''","'")
			removeSQ=replace(myString,"""","&quot;")
		end if
	end function

	Set objFS = Server.CreateObject ("Scripting.FileSystemObject")
	
	if PPD="1" then
		pcStrFileName=Server.Mappath ("/"&scPcFolder&"/includes/pcSBSettings.asp")
	else
		pcStrFileName=Server.Mappath ("../includes/pcSBSettings.asp")
	end if
	
	Set objFile = objFS.OpenTextFile (pcStrFileName, 2, True, 0)
	objFile.WriteLine CHR(60)&CHR(37)&"'// Storewide Settings //" & vbCrLf
	objFile.WriteLine "private const scSBStatus = """&pcv_Status&"""" & vbCrLf
	objFile.WriteLine "private const scSBShowStartDate = """&pcv_ShowStartDate&"""" & vbCrLf
	objFile.WriteLine "private const scSBStartDateDesc = """&removeSQ(pcv_StartDateDesc)&"""" & vbCrLf
	objFile.WriteLine "private const scSBShowReoccurenceDate = """&pcv_ShowReoccurenceDate&"""" & vbCrLf
	objFile.WriteLine "private const scSBReoccurenceDesc = """&removeSQ(pcv_ReoccurenceDesc)&"""" & vbCrLf
	objFile.WriteLine "private const scSBShowEOSDate = """&pcv_ShowEOSDate&"""" & vbCrLf
	objFile.WriteLine "private const scSBEOSDesc = """&removeSQ(pcv_EOSDesc)&"""" & vbCrLf
	objFile.WriteLine "private const scSBShowTrialDate = """&pcv_ShowTrialDate&"""" & vbCrLf
	objFile.WriteLine "private const scSBTrialDate = """&pcv_TrialDate&"""" & vbCrLf
	objFile.WriteLine "private const scSBShowTrialPrice = """&pcv_ShowTrialPrice&"""" & vbCrLf
	objFile.WriteLine "private const scSBTrialDesc = """&removeSQ(pcv_TrialDesc)&"""" & vbCrLf
	objFile.WriteLine "private const scSBShowFreeTrial = """&pcv_ShowFreeTrial&"""" & vbCrLf
	objFile.WriteLine "private const scSBFreeTrialDesc = """&removeSQ(pcv_FreeTrialDesc)&"""" & vbCrLf
	objFile.WriteLine "private const scSBShowInstallment = """&pcv_ShowInstallment&"""" & vbCrLf
	objFile.WriteLine "private const scSBInstallmentDesc = """&removeSQ(pcv_InstallmentDesc)&"""" & vbCrLf
	objFile.WriteLine "private const scSBaymentPageText = """&removeSQ(pcv_PaymentPageText)&"""" & vbCrLf
	objFile.WriteLine "private const scSBPaymentPageTrialText = """&removeSQ(pcv_PaymentPageTrialText)&"""" & vbCrLf
	objFile.WriteLine "private const scSBSuccessPaymentEmail = """&removeSQ(pcv_SuccessPaymentEmail)&"""" & vbCrLf
	objFile.WriteLine "private const scSBPendingPaymentEmail = """&removeSQ(pcv_PendingPaymentEmail)&"""" & vbCrLf
	objFile.WriteLine "private const scSBUnSuccessPaymentEmail = """&removeSQ(pcv_UnSuccessPaymentEmail)&"""" & vbCrLf
	objFile.WriteLine "private const scSBCCExpireEmail = """&removeSQ(pcv_CCExpireEmail)&"""" & vbCrLf
	objFile.WriteLine "private const scSBEOTFirstPaymentEmail = """&removeSQ(pcv_EOTFirstPaymentEmail)&"""" & vbCrLf
	objFile.WriteLine "private const scSBOffMsg = """&removeSQ(pcv_SBOffMsgText)&"""" & vbCrLf
	objFile.WriteLine "private const scSBRegAgree = """&removeSQ(pcv_SBRegAgree)&"""" & vbCrLf
    objFile.WriteLine "private const scSBAgreeText = """&removeSQ(pcv_SBAgreeText)&"""" & vbCrLf	
	objFile.WriteLine "private const scSBLang1 = """&removeSQ(pcv_SBLang1)&"""" & vbCrLf
	objFile.WriteLine "private const scSBLang2 = """&removeSQ(pcv_SBLang2)&"""" & vbCrLf
	objFile.WriteLine "private const scSBLang3 = """&removeSQ(pcv_SBLang3)&"""" & vbCrLf
	objFile.WriteLine "private const scSBLang4 = """&removeSQ(pcv_SBLang4)&"""" & vbCrLf
	objFile.WriteLine "private const scSBLang5 = """&removeSQ(pcv_SBLang5)&"""" & vbCrLf
	objFile.WriteLine "private const scSBLang6 = """&removeSQ(pcv_SBLang6)&"""" & vbCrLf
	objFile.WriteLine "private const scSBLang7 = """&removeSQ(pcv_SBLang7)&"""" & vbCrLf
	objFile.WriteLine "private const scSBLang8 = """&removeSQ(pcv_SBLang8)&"""" & vbCrLf
	objFile.WriteLine "private const scSBLang9 = """&removeSQ(pcv_SBLang9)&"""" & vbCrLf
	objFile.WriteLine "private const scSBLang10 = """&removeSQ(pcv_SBLang10)&"""" & vbCrLf
	objFile.WriteLine "'// Storewide Settings // " &CHR(37)&CHR(62)& vbCrLf 
	objFile.Close
	set objFS=nothing
	set objFile=nothing

	msg="Settings saved successfully!"
	
	response.Redirect "SB_Settings.asp?s=1&msg=" & msg
	response.end 
	'/////////////////////////////////////////////////////
	'// END:  Write all changes to pcSBSettings.asp file
	'/////////////////////////////////////////////////////
	
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// END: ON SUBMIT
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// START: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'// Find activated payment types that support SB
dim pcv_PayTypeNotAvail
pcv_PayTypeNotAvail=0
pcv_GwAuthorize="0"
pcv_GwAuthorizeSBActive=""

call openDb()
query="SELECT idPayment, gwCode, pcPayTypes_Subscription FROM payTypes WHERE gwCode=1 or gwCode=16 or gwCode=999999 or gwCode=46 or gwCode=67;"
set rs=Server.CreateObject("ADODB.Recordset") 
set rs=conntemp.execute(query)
if rs.eof then
	pcv_PayTypeNotAvail=1
	msg = "Recurring Billing cannot be enable until you have activated a compatible gateway<BR><BR>"
	Msg = msg & "<ul><li><a href=""AddModRTPayment.asp?gwchoice=1"" >Authorize.Net</a></li></ul> "
	Msg = msg & "<ul><li><a href=""pcPaymentSelection.asp"" >PayPal</a></li></ul> "
	Msg = msg & "<ul><li><a href=""AddModRTPayment.asp?gwchoice=67"" >NetSource Commerce Gateway</a></li></ul> "
	set rs=nothing
	call closeDb()
else
	do until rs.eof
		pcv_idPayment=rs("idPayment")
		pcv_gwCode=rs("gwCode")
		pcv_SubscriptionFlag=rs("pcPayTypes_Subscription")
		select case pcv_gwCode
			case "1"
				pcv_GwAuthorize="1"
				if pcv_SubscriptionFlag="1" then
					pcv_GwAuthorizeSBActive="checked"
				end if
			case "16" 
			    pcv_GwAuthorize="16"
				if pcv_SubscriptionFlag="1" then
					pcv_GwAuthorizeCheckSBActive="checked"
				end if
			case "999999" 
			    pcv_GwEC="999999"
				if pcv_SubscriptionFlag="1" then
					pcv_GwECSBActive="checked"
				end if
			case "46" 
			    pcv_GwWPP="46"
				if pcv_SubscriptionFlag="1" then
					pcv_GwWPPSBActive="checked"
				end if
			case "67" 
			    pcv_GwEIG="67"
				if pcv_SubscriptionFlag="1" then
					pcv_GwEIGSBActive="checked"
				end if				
		end select
		rs.moveNext
	loop
	set rs=nothing
	call closeDb()
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// END: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>
<form name="form1" method="post" action="SB_Settings.asp" class="pcForms">
	<table class="pcCPcontent">
        <tr>
            <td class="pcCPspacer"></td>
        </tr>
		<tr>
			<th colspan="4">Recurring Billing Settings</th>
		</tr>
		<tr>
			<td colspan="4" class="pcCPspacer"></td>
		</tr>
		<tr>
		  <td colspan="4">Turn Recurring Billing: &nbsp;
            <INPUT type="radio" value="1" name="SB_status" <% if scSBStatus="1" then %>checked<%end if%> class="cleSBorder">On
            <INPUT type="radio" value="0" name="SB_status" <% if scSBStatus="0" then %>checked<%end if%> class="clearBorder">Off 
          </td>
		</tr>
		<tr>
			<td colspan="4" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="4">Select the Gateways you wish to use for orders containing Recurring Billing products:</th>
		</tr>
		<tr>
			<td colspan="4" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td colspan="4" valign="middle">
            
            	<table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="4%"><input type="radio" name="SB_PaymentType" id="SB_Authorize" class="clearBorder" value="SB_Authorize" <%=pcv_GwAuthorizeSBActive%>></td>
                    <td width="24%">Authorize.Net</td>
                    <td width="4%">&nbsp;</td>
                    <td width="68%">&nbsp;</td>
            	  </tr>
				  <tr>
                    <td width="4%"><input type="radio" name="SB_PaymentType" id="SB_EC" class="clearBorder" value="SB_EC" <%=pcv_GwECSBActive%>></td>
                    <td width="24%">PayPal Express Checkout</td>
                    <td width="4%">&nbsp;</td>
                    <td width="68%">&nbsp;</td>
            	  </tr>
				  <tr>
                    <td width="4%"><input type="radio" name="SB_PaymentType" id="SB_WPP" class="clearBorder" value="SB_WPP" <%=pcv_GwWPPSBActive%>></td>
                    <td width="24%">PayPal Payments Pro</td>
                    <td width="4%">&nbsp;</td>
                    <td width="68%">&nbsp;</td>
            	  </tr>
				  <tr>
                    <td width="4%"><input type="radio" name="SB_PaymentType" id="SB_EIG" class="clearBorder" value="SB_EIG" <%=pcv_GwEIGSBActive%>></td>
                    <td width="24%">NetSource Commerce Gateway</td>
                    <td width="4%">&nbsp;</td>
                    <td width="68%">&nbsp;</td>
            	  </tr>
                </table>   
                   		
        	</td>
		</tr>
		<tr> 
			<td colspan="4" class="pcCPspacer"></td>
        </tr>
		<tr>
			<th colspan="4">Descriptions:</th>
		</tr>
		<tr>
			<td colspan="4" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td colspan="4">
				Payment Page Disclaimer:
				<div style="padding:5px;">
                    <%
                    pcv_SBaymentPageText=Replace(scSBaymentPageText, "<br>", vbCrLf)
                    pcv_SBaymentPageText=Replace(pcv_SBaymentPageText, """""", """")
                    %>
					<textarea name="SB_PaymentPageText" cols="80" rows="5" wrap="virtual"><%=pcv_SBaymentPageText%></textarea>
				</div>			
        	</td>
        </tr>
        <tr> 
			<td colspan="4">
				Payment Page Disclaimer for Trials:
				<div style="padding:5px;">
                    <%
                    pcv_SBPaymentPageTrialText=Replace(scSBPaymentPageTrialText, "<br>", vbCrLf)
                    pcv_SBPaymentPageTrialText=Replace(pcv_SBPaymentPageTrialText, """""", """")
                    %>
					<textarea name="SB_PaymentPageTrialText" cols="80" rows="5" wrap="virtual"><%=pcv_SBPaymentPageTrialText%></textarea>
				</div>            
        	</td>
		</tr>
		<tr> 
			<td colspan="4" class="pcCPspacer"></td>
        </tr>
		<tr>
			<th colspan="4">Messages:</th>
		</tr>
		<tr>
			<td colspan="4" class="pcCPspacer"></td>
		</tr>
        <tr>
          <td colspan="4">Message to show if SB is off:
            	<div style="padding:5px;">
                    <%
                    pcv_SBOffMsg=Replace(scSBOffMsg, "<br>", vbCrLf)
                    pcv_SBOffMsg=Replace(pcv_SBOffMsg, """""", """")
                    %>
              		<textarea name="SB_OffMsgText" cols="60" rows="3" wrap="virtual"><%=pcv_SBOffMsg%></textarea>
            	</div>
        	</td>
        </tr>
		<tr>
			<th colspan="4">Individual SB Agreement:</th>
		</tr>
		<tr>
          	<td colspan="4">Show and Require Agreement Check Box for each Subscription:
            	<div style="padding:5px;">
             		<input type="checkbox" name="SB_RegAgree" value="1" <% if scSBRegAgree="1" Then %> checked <% end if %>/>
            	</div>
        	</td>
        </tr>		
        <tr>
          	<td colspan="4">General Terms & Conditions Agreement:
            	<div style="padding:5px;">
                    <%
                    pcv_SBAgreeText=Replace(scSBAgreeText, "<br>", vbCrLf)
                    pcv_SBAgreeText=Replace(pcv_SBAgreeText, """""", """")
                    %>
              		<textarea name="SB_AgreeText" cols="60" rows="6" wrap="virtual"><%=pcv_SBAgreeText%></textarea>
            	</div>
        	</td>
        </tr>
		<tr> 
			<td colspan="4" class="pcCPspacer"></td>
        </tr>
		<tr>
			<th colspan="4">Labels:</th>
		</tr>
		<tr>
			<td colspan="4" class="pcCPspacer"></td>
		</tr>
        <tr> 
			<td colspan="4">
				Agreement Label 1:
			  	<div style="padding:5px;">	
                    <%
                    pcv_SBLang1=Replace(scSBLang1, "<br>", vbCrLf)
                    pcv_SBLang1=Replace(pcv_SBLang1, """""", """")
                    %>		
                	<input name="SB_Lang1" type="text" value="<%=pcv_SBLang1%>" size="80">
				</div>            
        	</td>
		</tr>
        <tr> 
			<td colspan="4">
				Agreement Label 2:
				<div style="padding:5px;">
                    <%
                    pcv_SBLang2=Replace(scSBLang2, "<br>", vbCrLf)
                    pcv_SBLang2=Replace(pcv_SBLang2, """""", """")
                    %>
					<input name="SB_Lang2" type="text" value="<%=pcv_SBLang2%>" size="80">
				</div>            
        	</td>
		</tr>
        <tr> 
			<td colspan="4">
				Agreement Label 3:
				<div style="padding:5px;">
                    <%
                    pcv_SBLang3=Replace(scSBLang3, "<br>", vbCrLf)
                    pcv_SBLang3=Replace(pcv_SBLang3, """""", """")
                    %>
					<input name="SB_Lang3" type="text" value="<%=pcv_SBLang3%>" size="80">
				</div>            
        	</td>
		</tr>
        <tr> 
			<td colspan="4">
				Agreement Label 4:
				<div style="padding:5px;">
                    <%
                    pcv_SBLang4=Replace(scSBLang4, "<br>", vbCrLf)
                    pcv_SBLang4=Replace(pcv_SBLang4, """""", """")
                    %>
					<input name="SB_Lang4" type="text" value="<%=pcv_SBLang4%>" size="80">
				</div>            
        	</td>
		</tr>
        <tr> 
			<td colspan="4">
				Error Message 1:
				<div style="padding:5px;">
                    <%
                    pcv_SBLang5=Replace(scSBLang5, "<br>", vbCrLf)
                    pcv_SBLang5=Replace(pcv_SBLang5, """""", """")
                    %>
					<textarea name="SB_Lang5" cols="80" rows="5" wrap="virtual"><%=pcv_SBLang5%></textarea>
				</div>            
        	</td>
		</tr>
        <tr> 
			<td colspan="4">
				Error Message 2:
				<div style="padding:5px;">
                    <%
                    pcv_SBLang6=Replace(scSBLang6, "<br>", vbCrLf)
                    pcv_SBLang6=Replace(pcv_SBLang6, """""", """")
                    %>  
					<textarea name="SB_Lang6" cols="80" rows="5" wrap="virtual"><%=pcv_SBLang6%></textarea>
				</div>            
        	</td>
		</tr>
        <tr> 
			<td colspan="4">
				Payment Label 1:
				<div style="padding:5px;">
                    <%
                    pcv_SBLang7=Replace(scSBLang7, "<br>", vbCrLf)
                    pcv_SBLang7=Replace(pcv_SBLang7, """""", """")
                    %>
					<input name="SB_Lang7" type="text" value="<%=pcv_SBLang7%>" size="80">
				</div>            
        	</td>
		</tr>
        <tr> 
			<td colspan="4">
				Payment Label 2:
				<div style="padding:5px;">
                    <%
                    pcv_SBLang8=Replace(scSBLang8, "<br>", vbCrLf)
                    pcv_SBLang8=Replace(pcv_SBLang8, """""", """")
                    %>
					<input name="SB_Lang8" type="text" value="<%=pcv_SBLang8%>" size="80">
				</div>            
        	</td>
		</tr>
        <tr> 
			<td colspan="4">
				Payment Label 3:
				<div style="padding:5px;">
                    <%
                    pcv_SBLang9=Replace(scSBLang9, "<br>", vbCrLf)
                    pcv_SBLang9=Replace(pcv_SBLang9, """""", """")
                    %>
					<input name="SB_Lang9" type="text" value="<%=pcv_SBLang9%>" size="80">
				</div>            
        	</td>
		</tr>
        <tr> 
			<td colspan="4">
				Payment Label 4:
				<div style="padding:5px;">
                    <%
                    pcv_SBLang10=Replace(scSBLang10, "<br>", vbCrLf)
                    pcv_SBLang10=Replace(pcv_SBLang10, """""", """")
                    %>
					<input name="SB_Lang10" type="text" value="<%=pcv_SBLang10%>" size="80">
				</div>            
        	</td>
		</tr>
		<tr>
			<td colspan="4" class="pcCPspacer"></td>
		</tr>
        <tr>
			<td colspan="4"><hr></td>
		</tr>
		<tr> 
			<td colspan="4" class="pcCPspacer"></td>
        </tr>
		<tr> 
			<td colspan="4" align="center">
				<input type="submit" name="SubmitSettings" value="Update" class="submit2">&nbsp;
				<input type="button" name="back" value="Back to Main Menu" onclick="location='sb_Default.asp';">
            </td>
        </tr>
		<tr> 
			<td colspan="4" class="pcCPspacer"></td>
        </tr>		
	</table>
</form>
<!--#include file="AdminFooter.asp"-->
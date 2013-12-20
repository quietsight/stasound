<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
<!--#include file="opc_contentType.asp" -->
<!--#include file="inc_sb.asp"-->
<% dim conntemp,query,rs 

HaveSecurity=0
if session("idCustomer")=0 OR session("idCustomer")="" then
	HaveSecurity=1
end if

call openDb()

'Check to see if ARB has been turned off by admin, then display message
If scSBStatus="0" then
	response.redirect "msg.asp?message=212"
End If 

Call SetContentType()

IF HaveSecurity=0 THEN

	qry_GUID = getUserInput(Request("GUID"),0)
	qry_ID = getUserInput(Request("ID"),0)
	if not validNum(qry_ID) then
	   qry_ID=0
	end if

	if request("action")="add" then

		'// Not needed

	end if
	
END IF
%>
<html>
<head>
<TITLE><%=dictLanguage.Item(Session("language")&"_opc_gwa_title")%></TITLE>
<%
Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN  

Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")
%>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
<script language="JavaScript">
<!--
	
	function Form1_Validator(theForm)
	{
		return (true);
	}

//-->
</script>
</head>
<body style="margin: 0;">
<div id="pcMain">
<form method="post" name="BillingForm" id="BillingForm" action="sb_CustOneTimePayment.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
<input name="ID" type="hidden" value="<%=qry_ID%>">
<input name="GUID" type="hidden" value="<%=qry_GUID%>">
<table class="pcMainTable">
	<%IF HaveSecurity=1 THEN%>
        <tr>
            <td>
                <div class="pcErrorMessage"><%response.write dictLanguage.Item(Session("language")&"_SB_20")%></div>
            </td>
        </tr>
    <%ELSE%>
    
        <%IF UpdateSuccess="1" THEN%>
            <tr>
                <td>
                    <div class="pcSuccessMessage"><%response.write dictLanguage.Item(Session("language")&"_SB_22")%></div>
                    <script>
                        setTimeout(function(){parent.closeManageSubscription()},1000);
                    </script>
                </td>
            </tr>        
		<%ELSE%>            


				<% If SB_ErrMsg <> "" Then %>
                <tr>
                    <td>
                        <div class="pcErrorMessage"><%=SB_ErrMsg%></div>
                    </td>
                </tr>
				<% End If %>
                <tr>
                	<td><h2><%response.write dictLanguage.Item(Session("language")&"_SB_21")%></h2></td>
                </tr>
				<%
				query="SELECT Setting_APIUser,Setting_APIPassword,Setting_APIKey,Setting_RegSuccess FROM SB_Settings;"
				set rsAPI=connTemp.execute(query)
				if not rsAPI.eof then
					Setting_APIUser=rsAPI("Setting_APIUser")
					Setting_APIPassword=enDeCrypt(rsAPI("Setting_APIPassword"), scCrypPass)
					Setting_APIKey=enDeCrypt(rsAPI("Setting_APIKey"), scCrypPass)
				end if
				set rsAPI=nothing                  
				
				Set objSB = NEW pcARBClass                  
				objSB.GUID = qry_GUID
				If scSBLanguageCode<>"" Then
					objSB.CartLanguageCode = scSBLanguageCode
				Else
					objSB.CartLanguageCode = "en-EN"
				End If
			  
				result = objSB.GetSubscriptionDetailsRequest(Setting_APIUser, Setting_APIPassword, Setting_APIKey)
				
				If SB_ErrMsg="" Then 					
					pcv_strGUID = objSB.pcf_GetNode(result, "Guid", "//GetSubscriptionDetailsResponse/Subscription/Identifiers")
					pcv_strStatus = objSB.pcf_GetNode(result, "Status", "//GetSubscriptionDetailsResponse/Subscription/Identifiers")
					pcv_strBillingEmail = objSB.pcf_GetNode(result, "Email", "//GetSubscriptionDetailsResponse/Subscription/Customer")
					pcv_Description = objSB.pcf_GetNode(result, "Description", "//GetSubscriptionDetailsResponse/Subscription/SubscriptionDetails")
					pcv_NextBillingAmt = objSB.pcf_GetNode(result, "NextBillingAmt", "//GetSubscriptionDetailsResponse/Subscription/BillingDetails")
					pcv_NextBillingDate = objSB.pcf_GetNode(result, "NextBillingDate", "//GetSubscriptionDetailsResponse/Subscription/BillingDetails")
		
					pcv_PreviousPaymentDate = objSB.pcf_GetNode(result, "LastPaymentDate", "//GetSubscriptionDetailsResponse/Subscription/BillingDetails")
					pcv_PreviousPaymentAmount = objSB.pcf_GetNode(result, "LastPaymentAmount", "//GetSubscriptionDetailsResponse/Subscription/BillingDetails")
					pcv_StartDate = objSB.pcf_GetNode(result, "StartDate", "//GetSubscriptionDetailsResponse/Subscription/BillingDetails")
					pcv_EndDate = objSB.pcf_GetNode(result, "EndDate", "//GetSubscriptionDetailsResponse/Subscription/BillingDetails")			
					If instr(pcv_EndDate,"1/1/1900")>0 Then
						pcv_EndDate=dictLanguage.Item(Session("language")&"_SB_30")
					End If
					pcv_strBalance = objSB.pcf_GetNode(result, "Balance", "//GetSubscriptionDetailsResponse/Subscription/OutstandingBalance")
					pcv_strReason = objSB.pcf_GetNode(result, "Reason", "//GetSubscriptionDetailsResponse/Subscription/OutstandingBalance")                     
				End If
                %>  
                <tr>
                    <td class="pcSpacer"></td>
                </tr>	
                <tr>
                    <td>
                
                		<table class="pcShowContent">
                          	<tr>
                              	<th><%response.write dictLanguage.Item(Session("language")&"_SB_11")%></th>
                              	<th><%response.write dictLanguage.Item(Session("language")&"_SB_23")%></th>
                                <th><%response.write dictLanguage.Item(Session("language")&"_SB_24")%></th>
                                <th><%response.write dictLanguage.Item(Session("language")&"_SB_25")%></th>
                          	</tr> 
                          	<tr>
                              	<td nowrap><%=pcv_strGUID%></td>
                              	<td><%=pcv_Description%></td>
                                <td><%=money(pcv_NextBillingAmt)%></td>
                                <td><%=pcv_NextBillingDate%></td>
                          	</tr>                
                		</table>
                
                		<table class="pcShowContent">
                          	<tr>
                              	<th><%response.write dictLanguage.Item(Session("language")&"_SB_26")%></th>
                              	<th><%response.write dictLanguage.Item(Session("language")&"_SB_27")%></th>
                                <th><%response.write dictLanguage.Item(Session("language")&"_SB_28")%></th>
                                <th><%response.write dictLanguage.Item(Session("language")&"_SB_29")%></th>
                          	</tr> 
                          	<tr>
                              	<td nowrap><%=pcv_PreviousPaymentDate%></td>
                              	<td nowrap><%=money(pcv_PreviousPaymentAmount)%></td>
                                <td nowrap><%=pcv_StartDate%></td>
                                <td nowrap><%=pcv_EndDate%></td>
                          	</tr>                
                		</table>
                        
                		<table class="pcShowContent">
                          	<tr>
                              	<th><%response.write dictLanguage.Item(Session("language")&"_SB_31")%></th>
                              	<th><%response.write dictLanguage.Item(Session("language")&"_SB_32")%></th>
                                <th></th>
                                <th></th>
                          	</tr> 
                          	<tr>
                              	<td nowrap><%=pcv_strStatus%></td>
                              	<td nowrap><%=pcv_strBillingEmail%></td>
                                <td nowrap></td>
                                <td nowrap></td>
                          	</tr>                
                		</table> 
                        
                    </td>
                </tr>
                <tr>
                    <td class="pcSpacer"></td>
                </tr>	
        <%END IF%>
    <%END IF%>
</table>
</form>
</div>
</body>
</html>
<% call closedb() %>
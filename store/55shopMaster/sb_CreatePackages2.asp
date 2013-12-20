<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Create Subscription Package Link" %>
<% Section="mngAcc" %>
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
<!--#include file="sb_inc.asp"-->
<% dim conntemp, rs, query

Dim pcv_strPageName
pcv_strPageName="sb_CreatePackages2.asp"

call openDb()

on error goto 0

pcMessage = getUserInput(Request.QueryString("msg"),0)

If request("prdlist")="" Then
	response.redirect("sb_CreatePackages.asp")
End If

if (request("action")="apply") then
  
  	intIsLinked= "1" '//request("IsLinked")
	strLinkedPackage=request("LinkedPackage")
	strRefName=request("RefName")
	strBillingPeriod=request("BillingPeriod")
	intBillingFrequency=request("BillingFrequency")
	intBillingCycles=request("BillingCycles")
	strCurrencyCode=request("CurrencyCode")
	intTrialAmount=request("TrialAmount")
	strTrialBillingPeriod=request("TrialBillingPeriod")
	intTrialBillingFrequency=request("TrialBillingFrequency")
	intTrialBillingCycles=request("TrialBillingCycles")
	intStartsImmediately=request("StartsImmediately")
	strStartDate=request("StartDate")	
	
	If intIsLinked="1" Then
		TrialPrice=request("LinkedTrialPrice")
		intIsTrial=request("LinkedIsTrial")
	Else
		TrialPrice=request("TrialPrice")
		Price=request("Price")	
		intIsTrial=request("IsTrial")
	End If
	
	If intBillingFrequency="" Then intBillingFrequency=0
	If intBillingCycles="" Then intBillingCycles=0
	If intTrialBillingFrequency="" Then intTrialBillingFrequency=0
	If intTrialBillingCycles="" Then intTrialBillingCycles=0
	If intStartsImmediately="" Then intStartsImmediately=0
	If TrialPrice="" Then TrialPrice=0
	If Price="" Then Price=0
	If intIsTrial="" Then intIsTrial=0
  
  	'// Validate Option 1
	If intIsLinked="1" AND strLinkedPackage="" Then
	
		pcMessage = "You must select a Package, or fill out the details"
	
	Else

		Dim pcv_strSuccess 
		pcv_strSuccess = "0"
		If (request("prdlist")<>"") and (request("prdlist")<>",") then
			prdlist=split(request("prdlist"),",")
			For i=lbound(prdlist) to ubound(prdlist)
				id=prdlist(i)
				IF (id<>"0") AND (id<>"") THEN				

					query="INSERT INTO SB_Packages (idProduct, SB_LinkID, SB_IsLinked, SB_RefName, SB_Amount, SB_BillingPeriod, SB_BillingFrequency, SB_BillingCycles, SB_CurrencyCode, SB_IsTrial, SB_TrialAmount, SB_TrialBillingPeriod, SB_TrialBillingFrequency, SB_TrialBillingCycles, SB_StartsImmediately, SB_StartDate) values (" & id & ",'" & strLinkedPackage & "', " & intIsLinked & ", '" & strRefName & "', " & Price & ", '" & strBillingPeriod & "', " & intBillingFrequency & ", " & intBillingCycles & ", '" & strCurrencyCode & "', " & intIsTrial & ", " & TrialPrice & ", '" & strTrialBillingPeriod & "', " & intTrialBillingFrequency & ", " & intTrialBillingCycles & ", " & intStartsImmediately & ", '" & strStartDate & "' );"
					'response.Write(query)
					'response.End()
					Set rstemp=conntemp.execute(query)
					Set rstemp=nothing
					pcv_strSuccess = "1"
	
				END IF
			next
		End if 'have prdlist
	
	End If

End if 'action=apply


'if (request("action")="setup") then
  
	If (request("prdlist")<>"") and (request("prdlist")<>",") then
		prdlist=split(request("prdlist"),",")
		For i=lbound(prdlist) to ubound(prdlist)
			id=prdlist(i)
			IF (id<>"0") AND (id<>"") THEN				

				query="SELECT description FROM products WHERE idProduct = " & id & ";"
				Set rstemp=conntemp.execute(query)
				pcv_strProductName = rstemp("description")
				Set rstemp=nothing				
				
			END IF
		next
	End if 'have prdlist

'End if 'action=setup

' START show message, if any
If pcMessage <> "" Then %>
	<div class="pcCPmessage">
		<%=pcMessage%>
	</div>
<% 	
end if
' END show message 
%>
<form id="form1" name="form1" method="post" action="<%=pcv_strPageName%>" class="pcForms">
<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<% if (request("action")="apply") AND (pcv_strSuccess = "1") then %>
	<tr> 
		<td align="center"> 
			<div class="pcCPmessageSuccess">
            	Success! The selected product is now a subscription package.
				<a href="sb_ModPackSettings.asp?idmain=<%=prdlist(0)%>">Configure Settings  &gt;&gt;</a>
			</div>
		</td>
	</tr>
	<tr>
		<td align="center">
			<a href="sb_ViewPackages.asp">View/Modify Package Links</a> 
            &nbsp;|&nbsp;
            <a href="sb_Default.asp">Main Menu</a>  
		</td>
	</tr>
    <% else %>
	<tr>
		<th>You selected: <%=pcv_strProductName%></th>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
        	<p>Link the selected product to an existing SubscriptionBridge Package:</p>
            <ul>	
                <li>All details will be retrieved dynamically from your SubscriptionBridge Merchant Center</li> 
                <li>Customers will be able to add/remove optional features to/from the subscription (if any are present)</li>
            </ul>
        </td>
	</tr>
	<tr>
		<td>
			<!--#include file="sb_inc_Packages.asp" -->
        </td>
	</tr>
    <tr>
        <td>
        	<hr />
            <input name="prdlist" type="hidden" value="<%=request("prdlist")%>">
            <input name="action" type="hidden" value="apply">
            <input type="submit" name="submit1" value=" Create Package Link " class="submit2">&nbsp;
            <input type="button" name="back" value="Back to Main Menu" onclick="location='sb_Default.asp';">
        </td>
    </tr>
    <% end if %>
</table>
</form>
<%
call closedb()
%>
<!--#include file="AdminFooter.asp"-->
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
Dim pageTitle, pageName, pageIcon, Section
pageTitle="Modify Subscription Package Link"
pageName="sb_ModPackage.asp"
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
<!--#include file="sb_inc.asp"-->
<% dim conntemp, rs, query

Dim pcv_strPageName
pcv_strPageName="sb_ModPackage.asp"

call openDb()

pcMessage = getUserInput(Request.QueryString("msg"),0)

pcv_intIDMain = request("idmain")
If pcv_intIDMain="" Then
	response.redirect("sb_ViewPackages.asp")
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
		If (pcv_intIDMain<>"") then
			prdlist=split(pcv_intIDMain,",")
			For i=lbound(prdlist) to ubound(prdlist)
				id=prdlist(i)
				IF (id<>"0") AND (id<>"") THEN	

					query="UPDATE SB_Packages SET SB_LinkID='" & strLinkedPackage & "' ,"
					query=query&"SB_IsLinked=" & intIsLinked & " ,"
					query=query&"SB_RefName='" & strRefName & "' ,"
					query=query&"SB_Amount='" & Price & "' ,"
					query=query&"SB_BillingPeriod='" & strBillingPeriod & "' ,"
					query=query&"SB_BillingFrequency=" & intBillingFrequency & " ,"
					query=query&"SB_BillingCycles=" & intBillingCycles & " ,"
					query=query&"SB_CurrencyCode='" & strCurrencyCode & "' ,"
					query=query&"SB_IsTrial=" & intIsTrial & " ,"
					query=query&"SB_TrialAmount='" & TrialPrice & "' ,"
					query=query&"SB_TrialBillingPeriod='" & strTrialBillingPeriod & "' ,"
					query=query&"SB_TrialBillingFrequency=" & intTrialBillingFrequency & " ,"
					query=query&"SB_TrialBillingCycles=" & intTrialBillingCycles & " ,"
					query=query&"SB_StartsImmediately=" & intStartsImmediately & " ,"
					query=query&"SB_StartDate='" & strStartDate & "' "
					query=query&"WHERE idProduct = " & id 

					Set rstemp=conntemp.execute(query)
					Set rstemp=nothing
					pcv_strSuccess = "1"
	
				END IF
			next
		End if 'have prdlist
	
	End If

End if 'action=apply


'if (request("action")="setup") then
  
	If (request("idmain")<>"") and (request("idmain")<>",") then
		prdlist=split(request("idmain"),",")
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
            	Success! The selected product has been updated.
				<br /><br />
				<a href="sb_ModPackSettings.asp?idmain=<%=pcv_intIDMain%>">Modify Package Settings  &gt;&gt;</a>
                &nbsp;|&nbsp;
                <a href="sb_ModPackage.asp?idmain=<%=pcv_intIDMain%>">Edit it Again &gt;&gt;</a>
			</div>
		</td>
	</tr>
	<tr>
		<td align="center">
			<a href="sb_ViewPackages.asp">View/Modify Packages</a> 
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
			<!--#include file="sb_inc_Packages.asp" -->
        </td>
	</tr>
    <tr>
        <td>
        	<hr />            
            <input name="idmain" type="hidden" value="<%=request("idmain")%>">
            <input name="action" type="hidden" value="apply">
            <input type="submit" name="submit1" value=" Update Package " class="submit2">&nbsp;
            <input type="button" name="back" value=" Modify Package Settings " onclick="location='sb_ModPackSettings.asp?idmain=<%=pcv_intIDMain%>';">&nbsp;
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
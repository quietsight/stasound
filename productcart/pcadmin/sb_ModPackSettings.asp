<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
Dim pageTitle, pageName, pageIcon, Section
pageTitle="Subscription Package Settings"
pageName="sb_ModPackSettings.asp"
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
pcv_strPageName="sb_ModPackSettings.asp"

call openDb()

pcMessage = getUserInput(Request.QueryString("msg"),0)

pcv_intIDMain = request("idmain")
If pcv_intIDMain="" Then
	response.redirect("sb_ViewPackages.asp")
End If

if (request("action")="apply") then
  
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

	pcv_SBRegAgree=request.Form("SB_RegAgree")
	if pcv_SBRegAgree="" then
		pcv_SBRegAgree="0"
	end if
	
	pcv_SBAgreeText=replace(request.form("SB_AgreeText"),"""","&quot;")
	pcv_SBAgreeText=replace(pcv_SBAgreeText,"'","''")
	pcv_SBAgreeText=replace(pcv_SBAgreeText, vbCrLf, "<br>")
  
  	'// Validate Option 1
	'If intIsLinked="1" AND strLinkedPackage="" Then
	
	'	pcMessage = "You must select a Package, or fill out the details"
	
	'Else

		Dim pcv_strSuccess 
		pcv_strSuccess = "0"
		If (pcv_intIDMain<>"") then
			prdlist=split(pcv_intIDMain,",")
			For i=lbound(prdlist) to ubound(prdlist)
				id=prdlist(i)
				IF (id<>"0") AND (id<>"") THEN	
				
					query="UPDATE SB_Packages SET SB_TrialDesc='" & pcv_TrialDesc & "' ,"
					query=query&"SB_Agree=" & pcv_SBRegAgree & " ,"
					query=query&"SB_AgreeText='" & pcv_SBAgreeText & "' "
					query=query&"WHERE idProduct = " & id
					
					Set rstemp=conntemp.execute(query)
					Set rstemp=nothing
					pcv_strSuccess = "1"
	
				END IF
			next
		End if 'have prdlist
	
	'End If

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
            	Success! The selected package settings have been updated.
				<br /><br />
				<a href="sb_ModPackSettings.asp?idmain=<%=pcv_intIDMain%>">Modify Package Settings &gt;&gt;</a>
                &nbsp;|&nbsp;
                <a href="sb_ModPackage.asp?idmain=<%=pcv_intIDMain%>">Change Package &gt;&gt;</a>
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
        	<strong>The following settings for this package will override your global settings. </strong>
        </td>
	</tr>
	<tr>
		<td>
			
			<%
			if len(pcv_intIDMain)>0 then
  				query="SELECT * FROM SB_Packages WHERE idProduct="& pcv_intIDMain 
  				Set rstemp=conntemp.execute(query)
  				If NOT rstemp.eof Then
					pShowTrialPrice = rstemp("SB_ShowTrialPrice")
					pTrialDesc = rstemp("SB_TrialDesc")
					pShowFreeTrial = rstemp("SB_ShowFreeTrial")
					pShowStartDate = rstemp("SB_ShowStartDate")
					pStartDateDesc = rstemp("SB_StartDateDesc")
					pShowReoccurenceDate = rstemp("SB_ShowReoccurenceDate")
					pReoccurenceDesc = rstemp("SB_ReoccurenceDesc")
					pShowEOSDate = rstemp("SB_ShowEOSDate")
					pEOSDesc = rstemp("SB_EOSDesc")
					pShowTrialDate = rstemp("SB_ShowTrialDate")
					pTrialDate = rstemp("SB_TrialDate")
					pFreeTrialDesc = rstemp("SB_FreeTrialDesc") 
					pRegAgree = rstemp("SB_Agree")
					pAgreeText = rstemp("SB_AgreeText")
  				End If
  				Set rstemp=nothing
			end if
			%>
			<table class="pcCPcontent">
				<tr>
        			<td colspan="4" class="pcCPspacer"></td>
        		</tr>
				<tr>
					<th colspan="4">Terms &amp; Conditions Agreement:</th>
				</tr>
        		<tr>
        			<td colspan="4" class="pcCPspacer"></td>
        		</tr>
		 		<tr>
          			<td colspan="4">Require that customers agree to the following Agreement: <input type="checkbox" name="SB_RegAgree" value="1" <% if pRegAgree="1" Then %> checked <% end if %>/></td>
        		</tr>		
        		<tr>
          			<td colspan="4">Agreement to be displayed:
            			<div style="padding:5px;">
              				<textarea name="SB_AgreeText" cols="60" rows="6" wrap="virtual"><%=pAgreeText%></textarea>
            			</div>
                	</td>
        		</tr>
            </table>
            
        </td>
    </tr>
    <tr>
        <td>
        	<hr />            
            <input name="idmain" type="hidden" value="<%=request("idmain")%>">
            <input name="action" type="hidden" value="apply">
            <input type="submit" name="submit1" value=" Update Settings " class="submit2">&nbsp;
            <input type="button" name="back" value=" Modify Package " onclick="location='sb_ModPackage.asp?idmain=<%=pcv_intIDMain%>';">&nbsp;
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
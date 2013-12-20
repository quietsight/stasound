<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/ErrorHandler.asp"--> 
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="pcStartSession.asp"-->
<%

dim pTempIntPageID, pcv_IDPage

pTempIntPageID=session("idContentPageRedirect")
	if pTempIntPageID = "" then
		pTempIntPageID=getUserInput(request("idpage"),10)
	end if

	'// Validate Content Page ID
	if not validNum(pTempIntPageID) then
		response.redirect "default.asp"
	end if
	pcv_IDPage=pTempIntPageID
	session("idContentPageRedirect")=""
	
	'// Check for admin preview
	if scSeoURLs=1 then ' Retrieve additional querystring (if any) from session variable
		pcIntAdminPreview = InStr(lcase(session("strSeoQueryString")),"adminpreview=1")
	else
		pcIntAdminPreview = getUserInput(request("adminPreview"),10)
	end if
	if not validNum(pcIntAdminPreview) then pcIntAdminPreview=0
	if pcIntAdminPreview = 1 and session("admin") <> 0 then
		query1 = ""
	else
		query1 = " AND pcCont_InActive=0 AND pcCont_Published=1"
	end if
	
	
	'// Select pages compatible with customer type
	if session("customerCategory")<>0 then ' The customer belongs to a customer category
		' Load pages accessible by ALL, plus those accessible by the customer pricing category that the customer belongs to
		if session("customerType")=0 then
			' Customer category does NOT have wholesale privileges, so exclude those pages
			query2 = " AND (pcCont_CustomerType = 'ALL' OR pcCont_CustomerType='CC_" & session("customerCategory") &"')"
		else
			' Customer category HAS wholesale privileges, so include wholesale-only pages
			query2 = " AND (pcCont_CustomerType = 'ALL' OR pcCont_CustomerType = 'W' OR pcCont_CustomerType='CC_" & session("customerCategory") &"')"
		end if
	else
		if session("customerType")=0 then
			' Retail customer or customer not logged in: load pages accessible by ALL
			query2 = " AND pcCont_CustomerType = 'ALL'"
		else
			' Wholesale customer: load pages accessible by ALL and Wholesale customers only
			query2 = " AND (pcCont_CustomerType = 'ALL' OR pcCont_CustomerType = 'W')"
		end if
	end if
	

	Dim rs, connTemp
	call opendb()
	query="SELECT pcCont_PageName, pcCont_IncHeader, pcCont_MetaTitle, pcCont_Description, pcCont_Parent, pcCont_MetaDesc, pcCont_MetaKeywords, pcCont_PageTitle, pcCont_HideBackButton FROM pcContents WHERE pcCont_IDPage=" & pcv_IDPage & query1 & query2
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if rs.eof then
		set rs=nothing
		call closeDB()
		response.redirect "msg.asp?message=300"       
	end if

	pcv_PageNameH=rs("pcCont_PageName")
	pcv_IncHeader=rs("pcCont_IncHeader")
	if not pcv_IncHeader<>"" then
		pcv_IncHeader="0"
	end if
	pcv_DefaultTitle=rs("pcCont_MetaTitle")
	if isNull(pcv_DefaultTitle) or trim(pcv_DefaultTitle)="" then
		pcv_DefaultTitle=ClearHTMLTags2(pcv_PageNameH,0)
	end if
	pcv_DefaultTitle = pcv_DefaultTitle & " - " & scCompanyName
	pcv_Description=rs("pcCont_Description")
	
	pcInt_Parent=rs("pcCont_Parent")
	if not validNum(pcInt_Parent) then pcInt_Parent=0
	
	pcv_DefaultDescription=rs("pcCont_MetaDesc")
		if pcv_DefaultDescription="" or isNull(pcv_DefaultDescription) then
			pcv_DefaultDescription=pcv_DefaultTitle
		end if	
	pcv_DefaultKeywords=rs("pcCont_MetaKeywords")
	pcv_PageTitle=rs("pcCont_PageTitle")
	
	pcIntHideBackButton=rs("pcCont_HideBackButton")
	if not validNum(pcIntHideBackButton) then pcIntHideBackButton=0

	set rs=nothing	
	call closeDB()

' If this content page contains the header & footer, load them
' Otherwise just load the content page code itself.
%>
<!--#include file="header.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<!--#include file="inc_addThis.asp"-->
<div id="pcMain">
	<table class="pcMainTable">
	
	<%
	if pcv_PageTitle<>"" then
	%>
		<tr>
			<td> 
				<%
                '// PC v4.5 AddThis integration
                if scAddThisDisplay=1 then
					call openDb()
					pcs_AddThis
					call closeDb()
				end if
                %>
				<h1><%=pcv_PageTitle%></h1>
			</td>
		</tr>
	<%
	end if
	%>
		<tr>
			<td>
            
<%
if NOT pcv_IncHeader="1" then 
	response.Clear()
end if
%>

						<%=pcv_Description%>
                        
                        <%
						'// Back button
						if pcInt_Parent=0 then
							pcvPageLink="viewPages.asp"
						else
							pcvPageLink="viewPages.asp?idpage=" & pcInt_Parent
						end if
						
						if pcIntHideBackButton="0" then
						%>
                        <hr />
                        <div style="margin-top: 10px;">
                            <a href="<%=pcvPageLink%>"><img src="<%=rslayout("back")%>" title="alt="<%=dictLanguage.Item(Session("language")&"_viewPages_1")%>"" alt="<%=dictLanguage.Item(Session("language")&"_viewPages_1")%>"></a>
						</div>
<%
						end if
if pcv_IncHeader="1" then
%>
				</td>
			</tr>
		</table>
	</div>
	<!--#include file="footer.asp"-->
<%
end if
%>
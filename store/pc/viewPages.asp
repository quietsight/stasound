<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/opendb.asp"--> 
<!--#INCLUDE FILE="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="pcStartSession.asp"-->
<!--#include file="header.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<% 
dim query, conntemp, rs, query1, query2
dim pcPageId, pcvPageName, pcvProductsView, pcIntPageParent, pcvPageMetaTitle, pcvPageMetaDesc, pcvPageMetaKeywords, pcvPageThumbnail, pcIntHideBackButton

scCatTotal=(scCatRow*scCatRowsPerPage)
iPageSize=scCatTotal

If Request("page")="" Then
	iPageCurrent=1
Else
	iPageCurrent=CInt(Request("page"))
End If

'// Load data from Existing Pages - START

	'// Look for subBrand
	pcPageId=session("idParentContentPageRedirect")
	session("idParentContentPageRedirect")=""
	if not validNum(pcPageId) then pcPageId=trim(getUserInput(request("idpage"),10))
	if not validNum(pcPageId) then pcPageId=""
	if pcPageId<>"" then
		query1="pcCont_Parent="&pcPageId
		else
		query1="pcCont_Parent=0"
	end if

	call opendb()
	
	'// Select pages compatible with customer type
	if session("customerCategory")<>0 then ' The customer belongs to a customer category
		' Load pages accessible by ALL, plus those accessible by the customer pricing category that the customer belongs to
		query2 = " AND (pcCont_CustomerType = 'ALL' OR pcCont_CustomerType='CC_" & session("customerCategory") &"')"
	else
		if session("customerType")=0 then ' Retail customer or customer not logged in
			query2 = " AND pcCont_CustomerType = 'ALL'" ' Load pages accessible by ALL
		else
			query2 = " AND pcCont_CustomerType = 'W'" ' Load pages accessible by Wholesale customers only
		end if
	end if

	'// Check for admin preview
	pcIntAdminPreview = getUserInput(request("adminPreview"),2)
	if not validNum(pcIntAdminPreview) then pcIntAdminPreview=0
	if pcIntAdminPreview = 1 and session("admin") <> 0 then
		query3 = ""
	else
		query3 = " AND pcCont_InActive=0 AND pcCont_Published=1"
	end if
	
	query="SELECT pcCont_IDPage, pcCont_PageName, pcCont_IncHeader, pcCont_MetaTitle, pcCont_Description, pcCont_MetaDesc, pcCont_MetaKeywords, pcCont_Order, pcCont_Parent, pcCont_Published, pcCont_Thumbnail FROM pcContents WHERE " & query1 & query2 & query3 & " AND pcCont_MenuExclude=0 ORDER BY pcCont_Order, pcCont_PageName ASC;"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.PageSize=iPageSize
	rs.CacheSize=iPageSize
	rs.Open query, conntemp, adOpenStatic, adLockReadOnly, adCmdText
	dim iPageCount
	iPageCount=rs.PageCount
	If iPageCurrent > iPageCount Then iPageCurrent=iPageCount
	If iPageCurrent < 1 Then iPageCurrent=1
		
	If iPageCount=0 Then 
		' There are no pages
		set rs=nothing
		call closeDb()
		response.redirect "msg.asp?message=300"       
	End if
	
	rs.AbsolutePage=iPageCurrent

'// Load data from Existing Pages - END

'// Load Parent Page Information - START
if pcPageId<>"" then

	query="SELECT pcCont_IncHeader, pcCont_MetaTitle, pcCont_Description, pcCont_MetaDesc, pcCont_MetaKeywords, pcCont_Order, pcCont_Parent, pcCont_Published, pcCont_PageTitle, pcCont_HideBackButton FROM pcContents WHERE pcCont_InActive=0 AND pcCont_MenuExclude=0 AND pcCont_IDPage="&pcPageId
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)
	if rstemp.eof then
		set rstemp=nothing
		call closeDb()
		response.redirect "msg.asp?message=85"       
	End if
	
	parentPageTitle=pcf_PrintCharacters(rstemp("pcCont_PageTitle"))
	parentPageContent=pcf_PrintCharacters(rstemp("pcCont_Description"))
	
	pcIntPageParent=rstemp("pcCont_Parent")
	if not validNum(pcIntPageParent) then pcIntPageParent=0

	pcv_DefaultTitle=rstemp("pcCont_MetaTitle")
	if isNull(pcv_DefaultTitle) or trim(pcv_DefaultTitle)="" then
		pcv_DefaultTitle=ClearHTMLTags2(parentBrandName,0)
	end if
	pcv_DefaultTitle = pcv_DefaultTitle & " - " & scCompanyName
	pcv_DefaultDescription=rstemp("pcCont_MetaDesc")
	pcv_DefaultKeywords=rstemp("pcCont_MetaKeywords")
	
	pcIntHideBackButton=rstemp("pcCont_HideBackButton")
	if not validNum(pcIntHideBackButton) then pcIntHideBackButton=0
	
	set rstemp=nothing
	
end if
'// Load Parent Brand Information - END
%>
		<div id="pcMain">
		<table class="pcMainTable">
        	<% if pcPageId<>"" then %>
            <tr>
				<td> 
                	<%
					if parentPageTitle<>"" then
					%>
					<h1><%=parentPageTitle%></h1>
                    <%
					end if
					%>
                    <% if parentPageContent<> "" then %>
                    	<%=parentPageContent%>
                    <% end if %>
                </td>
			</tr>
            <tr>
            	<td><hr></td>
            </tr>
            <% end if %>        
			<tr>
				<td>
					<table class="pcShowContent">
						<tr>
							<% 
							i=0 
							iRecordsShown=0 
							Do While iRecordsShown < iPageSize And NOT rs.EOF
							
								pcIntIDPage=rs("pcCont_IDPage")
								pcvPageName=pcf_PrintCharacters(rs("pcCont_PageName"))
								pcvPageThumbnail=rs("pcCont_Thumbnail")
								pcvPageContent=pcf_PrintCharacters(rs("pcCont_Description"))
								pcIntPageParent=rs("pcCont_Parent")
									' Check for Sub Pages
									Dim pcIntSubPagesExist
									pcIntSubPagesExist=0
									query="SELECT pcCont_IDPage FROM pcContents WHERE pcCont_Parent="&pcIntIDPage
									set rstemp=Server.CreateObject("ADODB.RecordSet")
									set rstemp=conntemp.execute(query)
									if not rstemp.EOF then
										pcIntSubPagesExist=1
									end if
									set rstemp=nothing
								pcvPageMetaTitle=rs("pcCont_MetaTitle")
								pcvPageMetaDesc=rs("pcCont_MetaDesc")
								pcvPageMetaKeywords=rs("pcCont_MetaKeywords")
							
								If pcvPageThumbnail="" or isNull(pcvPageThumbnail) Then
									pcvPageThumbnail="pcDefaultThumb.gif"
								End if
								
								if not validNum(pcIntPageParent) then pcIntPageParent=0
								
								'Build Link
								Dim pcvPageLink, pcvPageType
									'// Call SEO Routine
									pcIntContentPageID=pcIntIDPage
     								pcvContentPageName=pcvPageName
									if pcIntSubPagesExist=1 then
										'// There are sub-pages -> let the routine know
										pcvPageType = "parent"
									else 
										'// No sub-pages
										pcvPageType = ""
									end if
									call pcGenerateSeoLinks
									'//
									pcvPageLink=pcStrCntPageLink

		
								Dim pcv_strWidth, pcv_strAlign
								pcv_strWidth = ""
								Select case scCatRow
									case "1": pcv_strWidth = "width=100%"
									case "2": pcv_strWidth = "width=50%"
									case "3": pcv_strWidth = "width=33%"
									case "4": pcv_strWidth = "width=25%"
									case "5": pcv_strWidth = "width=20%"
									case "6": pcv_strWidth = "width=16%"
									case "7": pcv_strWidth = "width=14%"
									case "8": pcv_strWidth = "width=12%"
								End Select
							%>
				
							<td <%=pcv_strWidth%>>
				
								<% 
								'// Use the same diplay setting used by categories to determine whether the thumbnail should be showm
								if scCatImages="0" then %>
								
									<table class="pcShowCategory">
										<tr>
											<td class="pcShowCategoryImage">
												<a href="<%=pcvPageLink%>"><img src="catalog/<%=pcvPageThumbnail%>" alt="<%=ClearHTMLTags2(pcvPageName,0)%>"></a>
											</td>
										</tr>
										<tr>
											<td class="pcShowCategoryInfo">
												<p>
													<a href="<%=pcvPageLink%>"><%=pcvPageName%></a>
												</p>
											</td>
										</tr>
									</table>
								
								<% else %>
		
									<p><a href="<%=pcvPageLink%>"><%=pcvPageName%></a></p>
		
								<% end if %>
							</td>
						<% i=i + 1
						If i > (scCatRow-1) then 
						response.write "</tr><tr>"
						i=0
						End If
						iRecordsShown=iRecordsShown + 1
						rs.movenext
						loop
						
						set rs=nothing
						call closeDb()
						%>
				</table>
				</td>
			</tr>
		</table>
	
		<% 
        iRecSize=10
        pcStrPageName="viewPages.asp"
        '*******************************
        ' START Page Navigation
        '*******************************
        
        If iPageCount>1 then %>
            <div class="pcPageNav">
            <%response.write(dictLanguage.Item(Session("language")&"_advSrcb_4") & iPageCurrent & dictLanguage.Item(Session("language")&"_advSrcb_5") & iPageCount)%>
            &nbsp;-&nbsp;
            <% if iPageCount>iRecSize then %>
                <% if cint(iPageCurrent)>iRecSize then %>
                    <a href="<%=pcStrPageName%>?page=1&amp;iPageSize=<%=iPageSize%>">First</a>&nbsp;
                <% end if %>
                <% if cint(iPageCurrent)>1 then
                    if cint(iPageCurrent)<iRecSize AND cint(iPageCurrent)<iRecSize then
                        iPagePrev=cint(iPageCurrent)-1
                    else
                        iPagePrev=iRecSize
                    end if %>
                    <a href="<%=pcStrPageName%>?page=<%=cint(iPageCurrent)-iPagePrev%>&amp;iPageSize=<%=iPageSize%>">Previous <%=iPagePrev%> Pages</a>
                <% end if
                if cint(iPageCurrent)+1>1 then
                    intPageNumber=cint(iPageCurrent)
                else
                    intPageNumber=1
                end if
            else
                intPageNumber=1
            end if
            
            if (cint(iPageCount)-cint(iPageCurrent))<iRecSize then
                iPageNext=cint(iPageCount)-cint(iPageCurrent)
            else
                iPageNext=iRecSize
            end if
        
            For pageNumber=intPageNumber To (cint(iPageCurrent) + (iPageNext))
                If Cint(pageNumber)=Cint(iPageCurrent) Then %>
                    <strong><%=pageNumber%></strong> 
                <% Else %>
                    <a href="<%=pcStrPageName%>?page=<%=pageNumber%>&amp;iPageSize=<%=iPageSize%>"><%=pageNumber%></a>
                <% End If 
            Next
            
            if (cint(iPageNext)+cint(iPageCurrent))=iPageCount then
            else
                if iPageCount>(cint(iPageCurrent) + (iRecSize-1)) then %>
                    <a href="<%=pcStrPageName%>?page=<%=cint(intPageNumber)+iPageNext%>&amp;iPageSize=<%=iPageSize%>">Next <%=iPageNext%> Pages</a>
                <% end if
            
                if cint(iPageCount)>iRecSize AND (cint(iPageCurrent)<>cint(iPageCount)) then %>
                    &nbsp;<a href="<%=pcStrPageName%>?page=<%=cint(iPageCount)%>&amp;iPageSize=<%=iPageSize%>">Last</a>
                <% end if 
            end if %>
            </div>
        <% end if
        '*******************************
        ' END Page Navigation
        '*******************************
        %>
        
        <%
		if pcPageId<>"" and pcIntHideBackButton="0" then ' Back button to top level pages
		%>
        <hr />
        <div style="margin-top: 10px;">
            <a href="viewPages.asp"><img src="<%=rslayout("back")%>" title="alt="<%=dictLanguage.Item(Session("language")&"_viewPages_1")%>"" alt="<%=dictLanguage.Item(Session("language")&"_viewPages_1")%>"></a>
        </div>
        <%
		end if
		%>
</div>
<!--#include file="footer.asp"-->
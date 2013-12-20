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
<% 
dim query, conntemp, rs
dim pcIntBrandID, pcvBrandsDescription, pcvBrandsSDescription, pcIntBrandsActive, pcIntSubBrandsView, pcvProductsView, pcIntBrandsParent, pcvBrandsMetaTitle, pcvBrandsMetaDesc, pcvBrandsMetaKeywords, pcvBrandsBrandLogoLg, pcIntResultsPerPage, pcIntCurrentPage, iPageSize, pcIntIDBrand

'Number of brands to show
pcIntResultsPerPage=getUserInput(request("results"),10)
if not validNum(pcIntResultsPerPage) then
	iPageSize=(scCatRow*scCatRowsPerPage)
	else
	iPageSize=pcIntResultsPerPage
end if

'Number of products displayed on the brands page
if not validNum(pcIntResultsPerPage) then
	pcIntResultsPerPage=(scPrdRow*scPrdRowsPerPage)
end if

pcIntCurrentPage=getUserInput(Request("page"),10)
If pcIntCurrentPage="" Then
	iPageCurrent=1
Else
	iPageCurrent=CInt(pcIntCurrentPage)
End If

'// Load Parent Brand Information - START
pcIntBrandID=trim(request("idbrand"))
if not validNum(pcIntBrandID) then pcIntBrandID=""

call opendb()

if pcIntBrandID<>"" then

	query="SELECT BrandName, pcBrands_SDescription, pcBrands_SubBrandsView, pcBrands_Parent, pcBrands_MetaTitle, pcBrands_MetaDesc, pcBrands_MetaKeywords, pcBrands_BrandLogoLg FROM Brands WHERE pcBrands_Active=1 AND idBrand="&pcIntBrandID
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)
	if rstemp.eof then
		set rstemp=nothing
		call closeDb()
		response.redirect "msg.asp?message=85"       
	End if
	parentBrandName=pcf_PrintCharacters(rstemp("BrandName"))
	parentBrandsSDescription=pcf_PrintCharacters(rstemp("pcBrands_SDescription"))
	parentIntSubBrandsView=rstemp("pcBrands_SubBrandsView")
	parentIntBrandsParent=rstemp("pcBrands_Parent")
	parentBrandLogoLg=rstemp("pcBrands_BrandLogoLg")

	pcv_DefaultTitle=rstemp("pcBrands_MetaTitle")
	if isNull(pcv_DefaultTitle) or trim(pcv_DefaultTitle)="" then
		pcv_DefaultTitle=ClearHTMLTags2(parentBrandName,0)
	end if
	pcv_DefaultTitle = pcv_DefaultTitle & " - " & scCompanyName
	pcv_DefaultDescription=rstemp("pcBrands_MetaDesc")
	pcv_DefaultKeywords=rstemp("pcBrands_MetaKeywords")
	
	set rstemp=nothing

	if not validNum(parentIntSubBrandsView) then parentIntSubBrandsView=0
	if not validNum(parentIntBrandsParent) then parentIntBrandsParent=0
	
end if
'// Load Parent Brand Information - END
%>
<!--#include file="header.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<%
'// Load data from Existing Brands - START

	'// Look for subBrand
	if pcIntBrandID<>"" then
		query1=" AND pcBrands_Parent="&pcIntBrandID
		else
		query1=" AND pcBrands_Parent=0"
	end if

	query="SELECT idbrand, BrandName, BrandLogo, pcBrands_Description, pcBrands_SDescription, pcBrands_SubBrandsView, pcBrands_ProductsView, pcBrands_Active, pcBrands_Parent, pcBrands_MetaTitle, pcBrands_MetaDesc, pcBrands_MetaKeywords, pcBrands_BrandLogoLg FROM Brands WHERE pcBrands_Active=1"&query1&" ORDER BY pcBrands_Order, BrandName ASC;"
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
		if pcIntBrandID<>"" then
			response.clear()
			response.redirect "msg.asp?message=303"     
		else
			response.clear()
			response.redirect "msg.asp?message=302"     
		end if  
	End if
	
	rs.AbsolutePage=iPageCurrent

'// Load data from Existing Brand - END
%>
		<div id="pcMain">
		<table class="pcMainTable">
        	<% if pcIntBrandID="" then %>
			<tr>
				<td> 
					<h1><%response.write dictLanguage.Item(Session("language")&"_titles_8")%></h1>
				</td>
			</tr>    
            <% else %>
            <tr>
				<td> 
					<h1><%=parentBrandName%></h1>
                    <% if parentBrandLogoLg<>"" then %>
                        <div style="width: 100%; text-align: center;"><img src="catalog/<%=parentBrandLogoLg%>" alt="<%=ClearHTMLTags2(parentBrandName,0)%>"></div>
                    <% end if %>
                    <% if parentBrandsSDescription<> "" then %>
                    	<div class="pcPageDesc"><%=parentBrandsSDescription%></div>
                    <% end if %>
                </td>
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

								pcIntIDBrand=rs("IdBrand")
								BrandName=pcf_PrintCharacters(rs("BrandName"))
								BrandLogo=rs("BrandLogo")
								pcvBrandsDescription=pcf_PrintCharacters(rs("pcBrands_Description"))
								pcvBrandsSDescription=pcf_PrintCharacters(rs("pcBrands_SDescription"))
								pcIntSubBrandsView=rs("pcBrands_SubBrandsView")
								pcvProductsView=rs("pcBrands_ProductsView")
								pcIntBrandsActive=rs("pcBrands_Active")
								pcIntBrandsParent=rs("pcBrands_Parent")
									' Check for SubBrands
									Dim pcIntSubBrandsExist
									pcIntSubBrandsExist=0
									query="SELECT idbrand FROM brands WHERE pcBrands_Parent="&pcIntIDBrand
									set rstemp=Server.CreateObject("ADODB.RecordSet")
									set rstemp=conntemp.execute(query)
									if not rstemp.EOF then
										pcIntSubBrandsExist=1
									end if
									set rstemp=nothing
								pcvBrandsMetaTitle=rs("pcBrands_MetaTitle")
								pcvBrandsMetaDesc=rs("pcBrands_MetaDesc")
								pcvBrandsMetaKeywords=rs("pcBrands_MetaKeywords")
								pcvBrandsBrandLogoLg=rs("pcBrands_BrandLogoLg")
							
								If pBrandLogo="" Then
									pBrandLogo="no_image.gif"
								End if
								if not validNum(pcIntSubBrandsView) then pcIntSubBrandsView=0
								if not validNum(pcIntBrandsActive) then pcIntBrandsActive=1
								if not validNum(pcIntBrandsParent) then pcIntBrandsParent=0
								
								'Build Link
								Dim pcvBrandsLink
								
								if pcIntSubBrandsExist=0 then
									pcvBrandsLink="showsearchresults.asp?IDBrand=" & pcIntIDBrand & "&iPageSize=" & pcIntResultsPerPage & "&pageStyle=" & pcvProductsView
									else
									pcvBrandsLink="viewBrands.asp?IDBrand=" & pcIntIDBrand
								end if

		
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
								if parentIntSubBrandsView="" or isNull(parentIntSubBrandsView) then
									pcBrandDisplaySetting = sBrandLogo
								else
									pcBrandDisplaySetting = parentIntSubBrandsView
								end if

								if pcBrandDisplaySetting="1" then %>
								
									<table class="pcShowCategory">
										<tr>
											<td class="pcShowCategoryImage">
												<a href="<%=pcvBrandsLink%>"><img src="catalog/<%=BrandLogo%>" alt="<%=ClearHTMLTags2(BrandName,0)%>"></a>
											</td>
										</tr>
										<tr>
											<td class="pcShowCategoryInfo">
												<p>
													<a href="<%=pcvBrandsLink%>"><%=BrandName%></a>
												</p>
											</td>
										</tr>
									</table>
								
								<% else %>
		
									<p><a href="<%=pcvBrandsLink%>"><%=BrandName%></a></p>
		
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
        pcStrPageName="viewbrands.asp"
        '*******************************
        ' START Page Navigation
        '*******************************
        
        If iPageCount>1 then %>
            <div class="pcPageNav">
            <%response.write(dictLanguage.Item(Session("language")&"_advSrcb_4") & iPageCurrent & dictLanguage.Item(Session("language")&"_advSrcb_5") & iPageCount)%>
            &nbsp;-&nbsp;
            <% if iPageCount>iRecSize then %>
                <% if cint(iPageCurrent)>iRecSize then %>
                    <a href="<%=pcStrPageName%>?page=1&amp;iPageSize=<%=iPageSize%>&amp;idbrand=<%=pcIntBrandID%>">First</a>&nbsp;
                <% end if %>
                <% if cint(iPageCurrent)>1 then
                    if cint(iPageCurrent)<iRecSize AND cint(iPageCurrent)<iRecSize then
                        iPagePrev=cint(iPageCurrent)-1
                    else
                        iPagePrev=iRecSize
                    end if %>
                    <a href="<%=pcStrPageName%>?page=<%=cint(iPageCurrent)-iPagePrev%>&amp;iPageSize=<%=iPageSize%>&amp;idbrand=<%=pcIntBrandID%>">Previous <%=iPagePrev%> Pages</a>
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
                    <a href="<%=pcStrPageName%>?page=<%=pageNumber%>&amp;iPageSize=<%=iPageSize%>&amp;idbrand=<%=pcIntBrandID%>"><%=pageNumber%></a>
                <% End If 
            Next
            
            if (cint(iPageNext)+cint(iPageCurrent))=iPageCount then
            else
                if iPageCount>(cint(iPageCurrent) + (iRecSize-1)) then %>
                    <a href="<%=pcStrPageName%>?page=<%=cint(intPageNumber)+iPageNext%>&amp;iPageSize=<%=iPageSize%>&amp;idbrand=<%=pcIntBrandID%>">Next <%=iPageNext%> Pages</a>
                <% end if
            
                if cint(iPageCount)>iRecSize AND (cint(iPageCurrent)<>cint(iPageCount)) then %>
                    &nbsp;<a href="<%=pcStrPageName%>?page=<%=cint(iPageCount)%>&amp;iPageSize=<%=iPageSize%>&amp;idbrand=<%=pcIntBrandID%>">Last</a>
                <% end if 
            end if %>
        	&nbsp;<a href="<%=pcStrPageName%>?results=9999&amp;iPageSize=<%=iPageSize%>&amp;idbrand=<%=pcIntBrandID%>" onClick="pcf_Open_viewAll();"><%=dictLanguage.Item(Session("language")&"_viewCategories_21")%></a>
            </div>
        <% end if
        '*******************************
        ' END Page Navigation
        '*******************************
		
		if pcIntBrandID<>"" then
		%>
        <div style="margin: 15px 0;"><a href="<%=pcStrPageName%>"><img src="<%=rslayout("back")%>"></a></div>
        <%
		end if        
        %>
</div>
<%
Response.Write(pcf_InitializePrototype())
response.Write(pcf_ModalWindow(dictLanguage.Item(Session("language")&"_viewCategories_22"), "viewAll", 200))
%>

<!--#include file="footer.asp"-->
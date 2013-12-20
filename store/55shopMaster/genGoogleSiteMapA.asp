<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin="1*3*"%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../pc/pcSeoFunctions.asp"-->
<% 
pageTitle="Generate XML sitemaps for Google, Bing, and Yahoo!" 
pageIcon="pcv4_icon_xml.gif"
section="layout"

Server.ScriptTimeout = 5400

Dim connTemp, rsTemp, query
Dim pcCatPrdArr,intCountCatPrd,intk
Dim LOCALtmpBody
Dim tmp1,tmp1a,tmp2,tmpC1,tmpC2,tmp3
Dim SPathInfo,SPathInfo1

'Looking for correct paths
SPathInfo=scStoreURL
SPathInfo1=scStoreURL
	
if Right(SPathInfo,1)="/" then
	if trim(scPcFolder)<>"" then
		SPathInfo=SPathInfo & scPcFolder & "/"
		SPathInfo1=SPathInfo1 & scPcFolder & "/"
	end if
else
	if trim(scPcFolder)<>"" then
		SPathInfo=SPathInfo & "/" & scPcFolder & "/"
		SPathInfo1=SPathInfo1 & "/" & scPcFolder & "/"
	end if
end if

if Right(SPathInfo,1)="/" then
	SPathInfo=SPathInfo & "pc/"		
else
	SPathInfo=SPathInfo & "/pc/"
end if

pcv_changeFreq=request("changeFreq")
TodayDateTime=Year(Date()) & "-" & FixDate(Month(Date())) & "-" & FixDate(Day(Date()))

tmp1="   <url>" & vbcrlf
tmp1=tmp1 & "      <loc>" & Server.HTMLEncode(SPathInfo)
tmp1a=Server.HTMLEncode("&IDCategory=")
tmp2="</loc>" & vbcrlf
tmp2=tmp2 & "      <lastmod>" & TodayDateTime & "</lastmod>" & vbcrlf
tmp2=tmp2 & "      <changefreq>" & pcv_changeFreq & "</changefreq>" & vbcrlf
tmp2=tmp2 & "   </url>" & vbcrlf
tmp2=tmp2 & "   <url>" & vbcrlf
tmp2=tmp2 & "      <loc>" & Server.HTMLEncode(SPathInfo)

tmpC1="   <url>" & vbcrlf
tmpC1=tmpC1 & "      <loc>" & Server.HTMLEncode(SPathInfo)
tmpC2="</loc>" & vbcrlf
tmpC2=tmpC2 & "      <lastmod>" & TodayDateTime & "</lastmod>" & vbcrlf
tmpC2=tmpC2 & "      <changefreq>" & pcv_changeFreq & "</changefreq>" & vbcrlf
tmpC2=tmpC2 & "   </url>" & vbcrlf
tmpC2=tmpC2 & "   <url>" & vbcrlf
tmpC2=tmpC2 & "      <loc>" & Server.HTMLEncode(SPathInfo)

tmp3="</loc>" & vbcrlf
tmp3=tmp3 & "      <lastmod>" & TodayDateTime & "</lastmod>" & vbcrlf
tmp3=tmp3 & "      <changefreq>" & pcv_changeFreq & "</changefreq>" & vbcrlf
tmp3=tmp3 & "   </url>" & vbcrlf

LOCALtmpBody=""

call opendb()

pcv_strSMBody=""

pcv_strDelCatList=request("catlist")


if (pcv_strDelCatList<>"") then
	pcv_strDelCatList=" " & pcv_strDelCatList & ","
end if

Dim pcv_strSMBody,SiteHeader1,SiteFooter1,SiteHeader2,SiteFooter2
Dim URLCount,FileCount
Dim TodayDateTime

URLCount=0
FileCount=0

SiteHeader1="<?xml version=""1.0"" encoding=""UTF-8""?>" & vbcrlf
SiteHeader1=SiteHeader1 & "<urlset xmlns=""http://www.sitemaps.org/schemas/sitemap/0.9"">" & vbcrlf

SiteFooter1="</urlset>"

SiteHeader2="<?xml version=""1.0"" encoding=""UTF-8""?>" & vbcrlf
SiteHeader2="<sitemapindex xmlns=""http://www.sitemaps.org/schemas/sitemap/0.9"">" & vbcrlf

SiteFooter2="</sitemapindex>"

pcv_strSMBody=""


Function FixDate(datevalue)
	Dim Tmp1,Tmp2
	Tmp2=datevalue
	Tmp1=Cstr(Tmp2)
	if cint(Tmp1)<10 then
		FixDate="0" & Tmp1
	else
		FixDate="" & Tmp1
	end if
End Function


Sub SaveSiteMap()

	if PPD="1" then
		If FileCount=1 then
			SavedFile = "/" & scPcFolder & "/" & "SiteMap.xml"
		Else
			SavedFile = "/" & scPcFolder & "/" & FileCount-1 & ".xml"
		End if
	else
		If FileCount=1 then
			SavedFile = "../SiteMap.xml"
		Else
			SavedFile = "../SiteMap" & FileCount-1 & ".xml"
		End if
	end if

	findit = Server.MapPath(Savedfile)
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	Set f = fso.CreateTextFile(findit,True)
	
	on error resume next
	f.WriteLine SiteHeader1 & pcv_strSMBody & SiteFooter1
	dim intShowErrMsg
	intShowErrMsg=0
	if err.number<>0 then
		intShowErrMsg=1
	else
		pcv_strSMBody=""
	end if
	f.close

End Sub


Sub SaveSiteMapIndex()

	pcv_strSMBody1=""
	
	For k=1 to FileCount
		If k=1 then
			SavedFile = "SiteMap.xml"
		Else
			SavedFile = "SiteMap" & k-1 & ".xml"
		End if
		pcv_strSMBody1=pcv_strSMBody1 & "   <sitemap>" & vbcrlf
		pcv_strSMBody1=pcv_strSMBody1 & "      <loc>" & Server.HTMLEncode(SPathInfo1 & SavedFile) & "</loc>" & vbcrlf
		pcv_strSMBody1=pcv_strSMBody1 & "      <lastmod>" & TodayDateTime & "</lastmod>" & vbcrlf
		pcv_strSMBody1=pcv_strSMBody1 & "   </sitemap>" & vbcrlf	
	Next
	
	if PPD = "1" then
		SavedFile = "/"&scPcFolder&"/" & "SiteMapIndex.xml"
	else
		SavedFile = "../SiteMapIndex.xml"
	end if
	
	findit = Server.MapPath(Savedfile)
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	Set f = fso.CreateTextFile(findit,True)
	on error resume next
	f.WriteLine SiteHeader2 & pcv_strSMBody1 & SiteFooter2
	dim intShowErrMsg
	intShowErrMsg=0
	if err.number<>0 then
		intShowErrMsg=1
	end if
	pcv_strSMBody2=""
	f.close

End Sub


Sub CheckSiteMap()
	if (URLCount>=10000) or (len(SiteHeader1 & pcv_strSMBody & SiteFooter1)>=10000000) then
		FileCount=FileCount+1
		if LOCALtmpBody<>"|~|" then
			LOCALtmpBody=LOCALtmpBody & "|!|"
			LOCALtmpBody=replace(LOCALtmpBody,"|~||$|",tmp1)
			LOCALtmpBody=replace(LOCALtmpBody,"|~||%|",tmpC1)
			LOCALtmpBody=replace(LOCALtmpBody,"|-||!|",tmp3)
			LOCALtmpBody=replace(LOCALtmpBody,"|-||$|",tmp2)
			LOCALtmpBody=replace(LOCALtmpBody,"|-||%|",tmpC2)
			LOCALtmpBody=replace(LOCALtmpBody,"|*|",tmp1a)
		else
			LOCALtmpBody=""
		end if
		pcv_strSMBody=pcv_strSMBody & LOCALtmpBody
		LOCALtmpBody="|~|"
		SaveSiteMap()
		URLCount=0
	end if
End Sub


Sub GenContent(incURL)

	'// SEO Links
	'// Build Navigation Category Link
	if scSeoURLs=1 then
		sdquery="SELECT pcCont_IDPage, pcCont_PageName FROM pcContents WHERE pcCont_IDPage=" & incURL
		set rsSideCatObj=Server.CreateObject("ADODB.RecordSet")					
		set rsSideCatObj=connTemp.execute(sdquery)
		if NOT rsSideCatObj.EOF then
			pcIntContentPageID=rsSideCatObj("pcCont_IDPage")
			pcvContentPageName=rsSideCatObj("pcCont_PageName")
			pcStrCntPageLink=pcvContentPageName & "-d" & pcIntContentPageID & ".htm"
			pcStrCntPageLink=removeChars(pcStrCntPageLink)
		end if
		set rsSideCatObj = nothing
	else
		pcStrCntPageLink="viewContent.asp?idpage=" & incURL
	end if

	pcv_strSMBody=pcv_strSMBody & "   <url>" & vbcrlf
    pcv_strSMBody=pcv_strSMBody & "      <loc>" & Server.HTMLEncode(SPathInfo & pcStrCntPageLink) & "</loc>" & vbcrlf
    pcv_strSMBody=pcv_strSMBody & "      <lastmod>" & TodayDateTime & "</lastmod>" & vbcrlf
    pcv_strSMBody=pcv_strSMBody & "      <changefreq>" & pcv_changeFreq & "</changefreq>" & vbcrlf
	pcv_strSMBody=pcv_strSMBody & "   </url>" & vbcrlf
	URLCount=URLCount+1
	CheckSiteMap()

End Sub


Sub GenURL(incURL)

	pcv_strSMBody=pcv_strSMBody & "   <url>" & vbcrlf
    pcv_strSMBody=pcv_strSMBody & "      <loc>" & Server.HTMLEncode(incURL) & "</loc>" & vbcrlf
    pcv_strSMBody=pcv_strSMBody & "      <lastmod>" & TodayDateTime & "</lastmod>" & vbcrlf
    pcv_strSMBody=pcv_strSMBody & "      <changefreq>" & pcv_changeFreq & "</changefreq>" & vbcrlf
	pcv_strSMBody=pcv_strSMBody & "   </url>" & vbcrlf
	URLCount=URLCount+1
	CheckSiteMap()

End Sub


Sub GenSMBody()
	Dim query,rs,pcArr,intCount,i,j,tmpCatList,saveCat

	tmpCatList=""
	tmpCatList=" (" & pcv_strDelCatList & "0)"
	if request("excWCats")="1" then
	  ExcWCatsStr=" AND categories.pccats_RetailHide<>1 "
	else
	  ExcWCatsStr=""
	end if
	
	if request("excNFSPrds")="1" then
	  ExcNFSPrds=" AND products.formQuantity=0 "
	else
	  ExcNFSPrds=""
	end if

	if scSeoURLs=1 then
		tmpQuery1="products.idproduct,products.description"
	else
		tmpQuery1="products.idproduct"
	end if
	query="SELECT DISTINCT " & tmpQuery1 & " FROM categories INNER JOIN (categories_products INNER JOIN products ON categories_products.idproduct=products.idproduct) ON (categories.idcategory=categories_products.idcategory AND categories.iBTOHide<>1" & ExcWCatsStr & " AND (NOT (categories.idcategory IN " & tmpCatList & "))) WHERE products.active = -1" & ExcNFSPrds & " ORDER BY products.idproduct ASC;"
	set rs=connTemp.execute(query)
	
	tmpCatList=""
	LOCALtmpBody="|~|"

	intCount=-1
	IF not rs.eof then
		pcArr=rs.getRows()
		intCount=ubound(pcArr,2)
		set rs=nothing
		saveCat=0
		For i=0 to intCount			
			if scSeoURLs=1 then
				LOCALtmpBody=LOCALtmpBody & "|$|" &  removeChars(pcArr(1,i) & "-" & "p" & pcArr(0,i)) & ".htm|-|"
			else
				LOCALtmpBody=LOCALtmpBody & "|$|" &  Server.HTMLEncode("viewPrd.asp?idproduct=" & pcArr(0,i)) & "|-|"
			end if
			URLCount=URLCount+1
			CheckSiteMap()
		Next
		
		tmpCatList=" AND (NOT (idcategory IN (" & pcv_strDelCatList & "0)))"
		
		if scSeoURLs=1 then
			tmpQuery1="idcategory,categoryDesc"
		else
			tmpQuery1="idcategory"
		end if
		
		query="SELECT DISTINCT " & tmpQuery1 & " FROM categories WHERE iBTOHide<>1" & ExcWCatsStr & tmpCatList & " ORDER BY categories.idcategory ASC;"
		set rs=connTemp.execute(query)
		intCount=-1
		IF not rs.eof then
			pcArr=rs.getRows()
			intCount=ubound(pcArr,2)
			set rs=nothing
			For i=0 to intCount
				if pcArr(0,i)<>"1" then					
					if scSeoURLs=1 then
						LOCALtmpBody=LOCALtmpBody & "|%|" & removeChars(pcArr(1,i) & "-c" & pcArr(0,i)) & ".htm|-|"
					else
						LOCALtmpBody=LOCALtmpBody & "|$|" &  Server.HTMLEncode("viewcategories.asp?idCategory=" & pcArr(0,i)) & "|-|"
					end if
				end if
				URLCount=URLCount+1
				CheckSiteMap()
			Next
		End if
		set rs=nothing
	Else
		intShowErrMsg=1
	End if
	set rs=nothing
	
	if LOCALtmpBody<>"|~|" then
		LOCALtmpBody=LOCALtmpBody & "|!|"
		LOCALtmpBody=replace(LOCALtmpBody,"|~||$|",tmp1)
		LOCALtmpBody=replace(LOCALtmpBody,"|~||%|",tmpC1)
		LOCALtmpBody=replace(LOCALtmpBody,"|-||!|",tmp3)
		LOCALtmpBody=replace(LOCALtmpBody,"|-||$|",tmp2)
		LOCALtmpBody=replace(LOCALtmpBody,"|-||%|",tmpC2)
		LOCALtmpBody=replace(LOCALtmpBody,"|*|",tmp1a)
	else
		LOCALtmpBody=""
	end if
	pcv_strSMBody=pcv_strSMBody & LOCALtmpBody

End Sub


Sub GenUrlList()

	AddList=request("addList")
	if AddList<>"" then
		AList=split(AddList,vbcrlf)
		For i=lbound(AList) to ubound(AList)
			if trim(AList(i))<>"" then
				Call GenURL(AList(i))
			end if
		Next
		Set fso=Server.CreateObject("Scripting.FileSystemObject")
		Set f=fso.CreateTextFile(server.MapPath(".") & "\importlogs\UrlList.txt",True)		
		if Err.number>0 then
			err.Description=""
			on error GoTo 0
		else
			f.Write(AddList)
			f.Close
		end if		
		set f=nothing
		set fso=nothing
	end if

End Sub


Sub GenPcContents()

	PCContentsQuery="SELECT pcCont_IDPage,pcCont_InActive FROM pcContents WHERE pcCont_InActive=0 ORDER BY pcCont_IDPage ASC"
	set rsPCConObj=server.CreateObject("ADODB.RecordSet")
	set rsPCConObj=connTemp.execute(PCContentsQuery)
	IF not rsPCConObj.eof then
		strPCConArray1=rsPCConObj.getRows()
		intPCConCnt1=ubound(strPCConArray1,2)
		set rsPCConObj=nothing
		For pcGCnt=0 to intPCConCnt1
			pcContIDPage=strPCConArray1(0,pcGCnt)
			Call GenContent(pcContIDPage)
		Next
	End If	

End Sub


'// Main Generator Script
GenSMBody()
GenUrlList()

if request("incContent")="1" then
	GenPcContents()
end if

if pcv_strSMBody<>"" then
	FileCount=FileCount+1
	SaveSiteMap()
end if

if FileCount>"1" then
	SaveSiteMapIndex()
	SiteMapName="SiteMapIndex.xml"
else
	SiteMapName="SiteMap.xml"
end if

%>
<% pageTitle="Generate XML sitemaps for Google, Bing, and Yahoo!" %>
<% section="layout" 
on error GoTo 0
%>
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
    <tr>
        <td colspan="3" class="pcCPspacer"></td>
    </tr>
<% if intShowErrMsg=0 then %>
	<%if mid(scStoreURL,len(scStoreURL),1)="/" then
		scSiteURL=scStoreURL
	else
		scSiteURL=scStoreURL & "/"
	end if
	scSiteURL=scSiteURL & scPcFolder & "/" & SiteMapName%>
	
	<tr>
		<td>
		
        	<div class="pcCPmessageSuccess">Your XML sitemap has been generated successfully!</div>
            <h2>Sitemap File Location</h2>
			<div style="padding-top: 10px; padding-bottom: 10px;">The sitemap file location is: <b><%=scSiteURL%></b> (<a href="<%=scSiteURL%>" target="_blank">View XML</a>).<br />Use the links below to notify the search engines.</div>
            <h2>Google</h2>
            <div>You have two ways to notify <span style="font-weight: bold">Google</span> of the existence of this site map (and you can do so whenever you update the site map). The first method requires a few more clicks, but allows you to keep track of your site map submissions.
                <ul>
                	<li><a href="https://www.google.com/webmasters/tools/siteoverview" target="_blank">Manually notify Google</a> by adding your sitemap to your Google Account. <a href="https://www.google.com/webmasters/tools/siteoverview" target="_blank">Click here to log in &gt;&gt;</a></li>
                	<li>Or you can <a href="genGoogleSiteMapB.asp?fn=<%=Server.URLEncode(scSiteURL)%>">automatically notify Google</a></li>
                </ul>
             </div>
             <h2>Yahoo! &amp; Bing</h2>
             <div>Now, you can also submit your sitemap to <span style="font-weight: bold">Yahoo!</span> and <span style="font-weight: bold">Bing</span>. Use the links below. If you need to copy and paste the sitemap URL, use the full URL shown above in bold. </div>
             <div style="padding-top: 10px;">
				<script>
                function openWin(file,window)
                {
                msgWindow=open(file,window,'');
                if (msgWindow.opener == null) msgWindow.opener = self;
                }
                </script>
                <form class="pcForms">
                <input type="button" name="btnSubmitFeed" value="Notify Yahoo!" onclick="javascript:openWin('https://siteexplorer.search.yahoo.com/submit?txtFeedUrl=<%=scSiteURL%>');"/>&nbsp;
                <input type="button" name="btnSubmitFeed1" value="Notify Bing" onclick="javascript:openWin('http://www.bing.com/webmaster/Crawl/Sitemaps/?url=<%=scStoreURL%>');"/>
                <input type="hidden" name="class" value="Search" />
                </form>
              </div>
				<div>&nbsp;</div>
            </td>
        </tr>
    <% else %>
        <tr>
            <td>
            	<div class="pcCPmessage">ProductCart was NOT able to create the Sitemap. Make sure that the <%=scPcFolder%> folder has "WRITE" permissions. <a href="genGoogleSiteMap.asp">Try again</a>.</div>
            </td>
        </tr>
    <% end if %>
</table>
<!--#include file="AdminFooter.asp"-->
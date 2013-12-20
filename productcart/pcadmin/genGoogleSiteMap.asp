<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin="1*3*"%>
<!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->

<% 
pageTitle="Generate XML sitemaps for Google, Live.com, and Yahoo!" 
pageIcon="pcv4_icon_xml.gif"
section="layout"

Dim connTemp, rsTemp, query
%>
<!--#include file="AdminHeader.asp"-->
<% on error resume next %>
<form method="post" name="form1" action="genGoogleSiteMapA.asp" class="pcForms">
<%
if mid(scStoreURL,len(scStoreURL),1)="/" then
	scSiteURL=scStoreURL
	scSiteURL2=scStoreURL
else
	scSiteURL=scStoreURL & "/"
	scSiteURL2=scStoreURL & "/"
end if
scSiteURL=scSiteURL & scPcFolder & "/" & "SiteMap.xml"
%>
<table class="pcCPcontent">
	<tr>
    	<td colspan="2" class="pcCPspacer"></td>
    </tr>
    <tr>
    	<td colspan="2">
        	<h2>About Sitemaps</h2>
            <div>Many search engines including Google, Yahoo! and Microsoft's Live.com have agreed on a common format that Web sites can use to submit an index of the pages that they include. To learn more about XML sitemaps and the sitemap protocol: <strong><a href="http://www.sitemaps.org/" target="_blank">http://www.sitemaps.org/</a></strong>.</div>
            <div style="padding-top:6px;">ProductCart can create an XML sitemap of your store, based on that protocol. <a href="http://wiki.earlyimpact.com/productcart/marketing-generate_sitemaps" target="_blank"><img src="images/pcv3_infoIcon.gif" width="16" height="16" alt="More information on this feature"></a></div>
        </td>
	</tr>
	<tr>
    	<td colspan="2" class="pcCPspacer"></td>
    </tr>
    <tr>
    	<td colspan="2">
        	<h2>Generating a Sitemap</h2>
            <div>To improve performance, you can limit the amount of categories included in the map. Select the categories that you would like to <u>exclude</u> from the map generation process. Use the CTRL button to select multiple categories.
        </td>
    </tr>
	<tr>
    	<td colspan="2" class="pcCPspacer"></td>
    </tr>
	<tr> 
		<td align="right" valign="top">
        Exclude these categories:
        </td>
        <td width="90%" valign="top">
            <select size="8" name="catlist" multiple>
            <% 
			call opendb()
			query="SELECT idcategory, idParentCategory, categorydesc FROM categories WHERE idcategory<>1 and iBTOHide<>1 ORDER BY categoryDesc ASC"
			set rstemp=Server.CreateObject("ADODB.Recordset")
            set rstemp=conntemp.execute(query)
            if err.number <> 0 then
				set rstemp=nothing
                call closeDb()
                response.redirect "techErr.asp?error="& Server.Urlencode("Error in retreiving categories from database on genGoogleSiteMap.asp: "&Err.Description) 
            end If
            if rstemp.eof then
                catnum="0"
			else
                pcArr=rstemp.getRows()
                set rstemp=nothing
				call closeDb()
                intCount=ubound(pcArr,2)
                For i=0 to intCount
                    idparentcategory=pcArr(1,i)
                    if idparentcategory="1" then %>
                        <option value='<%response.write pcArr(0,i)%>'><%=pcArr(2,i)%></option>
                    <%
					else
                    	For j=0 to intCount
                    		if Clng(pcArr(0,j))=Clng(idparentcategory) then
                        	parentDesc=pcArr(2,j)%>
                        		<option value='<%response.write pcArr(0,i)%>'><%response.write pcArr(2,i)&" - Parent: "&parentDesc %></option>
                    <%		exit for
                    		end if 
                    	Next
              		end if
                Next
            end if
            %>
            </select>
		</td>
	</tr>
    <tr>
        <td nowrap align="right" valign="top">Exclude wholesale-only categories</td>
        <td><input type="checkbox" name="excWCats" value="1" checked class="clearBorder"></td>
    </tr>
    <tr>
        <td align="right" valign="top" nowrap>Exclude &quot;not for sale&quot; products</td>
        <td><input type="checkbox" name="excNFSPrds" value="1" class="clearBorder"></td>
    </tr>
    <tr>
        <td align="right" valign="top" nowrap>URL Change Frequency:</td>
        <td valign="top">
        	How often will your pages change? Google considers this a hint and not a command to visit the URLs.
            <div style="padding-top: 6px">
            <select name="changeFreq">
                <option value="always">Always</option>
                <option value="hourly">Hourly</option>
                <option value="daily">Daily</option>
                <option value="weekly" selected>Weekly</option>
                <option value="monthly">Monthly</option>
                <option value="yearly">Yearly</option>
                <option value="never">Never</option>
            </select>
            </div>
        </td>
    </tr>
    <tr>
        <td nowrap align="right" valign="top">Include content pages</td>
        <td><input type="checkbox" name="incContent" value="1" checked class="clearBorder"></td>
    </tr>
    <tr>
        <td align="right" valign="top"><a name="add"></a>Add additional URLs</td>
        <td>
            For example, non-ProductCart pages. <a href="http://wiki.earlyimpact.com/productcart/marketing-generate_sitemaps" target="_blank">See the documentation</a> about the location of the Sitemap file.<br /><u>Note</u>: Enter one URL per line
            <div style="padding-top: 6px">
        	<%
            URLList=""
            FileTxt = "importlogs/UrlList.txt"
            findit = Server.MapPath(FileTxt)
            Set fso = server.CreateObject("Scripting.FileSystemObject")
            Err.number=0
            Set f = fso.OpenTextFile(findit, 1)
            if Err.number>0 then
            err.Description=""
            err.number=0
            URLList=""
            else
            URLList=f.ReadAll
            f.close
            end if
            
            set f=nothing
            set fso=nothing
        	%>
        	<textarea name="addList" cols="53" rows="10"><%=URLList%></textarea>
        </td>
    </tr>
    <tr> 
        <td colspan="2" class="pcCPspacer"></td>
    </tr>
    <tr> 
        <td colspan="2" align="center">  
        	<input name="submit" type="submit" class="submit2" value="Generate Site Map">
        </td>
    </tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->
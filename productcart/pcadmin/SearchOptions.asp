<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce, Icon. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%>
<!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/SearchConstants.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<% 
pageTitle="Search Options"
pageIcon="pcv4_icon_search.png"
section="layout"
%>
<% dim conntemp, rs, query

Dim pcv_strPageName
pcv_strPageName="SearchOptions.asp"

msg=Request("msg")

call opendb()

If request("action")="add" Then

	Session("SRCH_MAX")=Request("max")
	Session("SRCH_CSFON")=Request("csfOn")
	Session("SRCH_CSFRON")=Request("csfROn")
	Session("SRCH_WAITBOX")=Request("waitbox")
	Session("SRCH_SUBS")=Request("subcats")

	response.Redirect("../includes/PageCreateSearchConstants.asp")

End If
%>
<!--#include file="AdminHeader.asp"-->
<style>
.pcCPOverview {
	background-color: #F5F5F5;
	border: 1px solid #FF9900;
	margin: 5px;
	padding: 5px;
	color: #666666;
	font-size:11px;
	text-align: left;
}
.pcCodeStyle {
	font-family: "Courier New", Courier, monospace;
	color: #FF0000;
	font-size: 9;
}
</style>
<% If msg="success" Then %>

<%
Session("SNW_TYPE")=""
Session("SNW_CATEGORY")=""
Session("SNW_MAX")=""
Session("SNW_AFFILIATE")=""
%>
<table class="pcCPcontent">
<tr>
	<td align="center">
    <div class="pcCPmessageSuccess">
		<p>Search Options Saved Successfully!</p>
		<p>&nbsp;</p>
        <p style="font-weight: normal;">Try a search on the storefront and reduce your search results as needed. Stores with thousands of products should keep search results to a minimum to reduce page load times.</p>
        <p>&nbsp;</p>
        <p style="font-weight: normal;">Note: if the advanced search page take a while to load, you can speed it up by disabling the category drop-down, or only make it load top-level categories. This option is on the <a href="AdminSettings.asp?tab=4">Store Settings</a> page.</p>
        <p>&nbsp;</p>      
        <p style="font-weight: normal;"><a href="../pc/search.asp" target="_blank">Try a Search</a> | <a href="SearchOptions.asp">Modify Search Options Again</a></p>
	</div>
    
    <% if Session("SRCH_CSFON")="1" or Session("SRCH_CSFRON")="1" then %>
    <div class="pcCPmessage" style="margin-top: 15px;">
		<p>IMPORTANT NOTE</p>
		<p>&nbsp;</p>
		<p style="font-weight: normal;">Make sure that the file &quot;<strong>CategorySearchFields.asp</strong>&quot; has been included into either <em>header.asp</em> or <em>footer.asp</em>. <a href="http://wiki.earlyimpact.com/productcart/search_fields_widget#storefront" target="_blank">Instructions &gt;&gt;</a></p>
	</div>
    <% end if %>
	</td>
</tr>
<tr>
	<td class="pcCPspacer"></td>
</tr>
</table>

<% Else %>

<form method="post" name="form1" action="<%=pcv_strPageName%>?action=add" class="pcForms">
	<table class="pcCPcontent">
	<%if msg<>"" then%>
	<tr>
		<td>
			<div class="pcCPmessage"><%=msg%></div>
     	</td>
	</tr>
	<%end if%>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
			<p>These settings allow you to optimize your storefront search performance based on the number of products in your database. <a href="http://wiki.earlyimpact.com/productcart/settings-search_options" target="_blank"><img src="images/pcv3_infoIcon.gif" width="16" height="16" alt="More information on this feature" title="More information on this feature" border="0"></a></p>
      </td>
	</tr>
	<tr>
		<td class="pcCPspacer"><hr></td>
	</tr>
	<tr>
		<td>
        <p><strong>Sub-category Search</strong></p>
        <p>Include sub categories when searching in a specific category:
		    <input type="radio" name="subcats" value="1" class="clearBorder" <% if SRCH_SUBS="1" then response.Write("checked") %>> 
		    Yes&nbsp;&nbsp;
		    <input type="radio" name="subcats" value="0" class="clearBorder" <% if SRCH_SUBS="0" then response.Write("checked") %>> 
		    No 
        </p>
    	</td>
	</tr>
	<tr>
		<td>
        <p><strong>Category List on Advanced Search Page</strong></p>
        <p>You can limit the amount of categories loaded in the &quot;Category&quot; drop-down on the advanced search page. See the <a href="AdminSettings.asp?tab=4" target="_blank">Store Settings page</a>.</p>
    	</td>
	</tr>
	<tr>
		<td class="pcCPspacer"><hr></td>
	</tr>
	<tr>
		<td>
        	<p>
        	<span style="font-weight: bold">Max Search Results:</span> 
		    <input name="max" type="text" value="<%=SRCH_MAX%>" size="4" maxlength="4"> 
		    &nbsp;&nbsp;<i class="pcCPnotes">Enter "0" for unlimited results</i>
        </p>
    	</td>
	</tr>
	<tr>
		<td class="pcCPspacer"><hr></td>
	</tr>
	<tr>
		<td><p><span style="font-weight: bold">Enable Waiting Dialog Box:</span> 
		    <input type="radio" name="waitbox" value="1" class="clearBorder" <% if SRCH_WAITBOX="1" then response.Write("checked") %>> 
		    Yes&nbsp;&nbsp;
		    <input type="radio" name="waitbox" value="0" class="clearBorder" <% if SRCH_WAITBOX="0" then response.Write("checked") %>> 
		    No</p></td>
	</tr>
	<tr>
		<td class="pcCPspacer"><hr></td>
	</tr>
	<tr>
		<td>
        	<p><strong>Custom search field widget for categories</strong> - <a href="http://wiki.earlyimpact.com/productcart/search_fields_widget" target="_blank">Help on this topic</a>&nbsp;|&nbsp;<a href="ManageSearchFields.asp">Manage Search Fields</a></p>	
    	</td>
	</tr>
	<tr>
		<td>
        	<p>
                Show  on Search Results: 
                <input type="radio" name="csfROn" value="1" class="clearBorder" <% if SRCH_CSFRON="1" then response.Write("checked") %>> 
                Yes&nbsp;&nbsp;
                <input type="radio" name="csfROn" value="0" class="clearBorder" <% if SRCH_CSFRON="0" then response.Write("checked") %>> 
                No 
                &nbsp;
            </p>
        	<p>
                Show on Category Pages: 
                <input type="radio" name="csfOn" value="1" class="clearBorder" <% if SRCH_CSFON="1" then response.Write("checked") %>> 
                Yes&nbsp;&nbsp;
                <input type="radio" name="csfOn" value="0" class="clearBorder" <% if SRCH_CSFON="0" then response.Write("checked") %>> 
                No 
        	</p> 
    	</td>
	</tr>
	<tr>
		<td class="pcCPspacer"><hr></td>
	</tr>
	<tr> 
		<td style="text-align: center;">
			<input name="submit" type="submit" class="submit2" value="Save Search Options">&nbsp;
    </td>
	</tr>
	</table>
</form>
<%
End If
call closedb()%>
<!--#include file="AdminFooter.asp"-->
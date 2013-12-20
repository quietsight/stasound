<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin="1*3*"%>
<!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp" -->
<% pageTitle="Generate Store Map" %>
<% section="layout" %>
<%
Dim connTemp,rsTemp
call opendb()
%>
<!--#include file="AdminHeader.asp"-->
<%
'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<form method="post" name="form1" action="genStoreMapA.asp" class="pcForms">
<table class="pcCPcontent">
<tr>
<td colspan="2">ProductCart can generate an <span style="font-weight: bold">HTML "store map"</span> that will include a list of categories, their subcategories, and the products that they contain, organized in a <span style="font-style: italic">Site Map</span> layout. Use this page to help search engine spiders locate and crawl your shopping cart pages. Note that the page will be saved to the following location:<span style="font-weight: bold"> /pc/StoreMap.asp</span>. Make sure that the catalog folder has "write" permissions.</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
	<th colspan="2">General Options:</th>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
<td nowrap>Category Only (no product links):</td>
<td width="95%">  
<input type="checkbox" value="1" name="catonly" class="clearBorder">
</td>
</tr>
<tr> 
<td nowrap>Include short category description:</td>
<td width="95%">  
<input type="radio" value="1" name="catdesc" class="clearBorder">Yes
<input type="radio" value="0" name="catdesc" checked class="clearBorder">No</td>
</tr>
<tr> 
<td nowrap>Include short product description:</td>
<td>
<input type="radio" value="1" name="prodesc" class="clearBorder">Yes
<input type="radio" value="0" name="prodesc" checked class="clearBorder">No</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
<td colspan="2">To improve performance, you can limit the amount of categories included in the map. Please select the categories that you would like to <strong>exclude</strong> from the map generation process.</td>
</tr>
<tr> 
<td colspan="2">Exclude the following cateories (<em>use the CTRL button to select multiple categories</em>):
<div style="margin-top: 10px;">
<select size="8" name="catlist" multiple>
<% query="SELECT idcategory,idParentCategory, categorydesc FROM categories WHERE idcategory<>1 and iBTOHide<>1 ORDER BY categoryDesc ASC"
set rstemp=conntemp.execute(query)
if err.number <> 0 then
	call closeDb()
	response.redirect "techErr.asp?error="& Server.Urlencode("Error in retreiving categories from database: "&Err.Description) 
end If
if rstemp.eof then
	catnum="0"
end If
if catnum<>"0" then
	pcArr=rstemp.getRows()
	set rstemp=nothing
	intCount=ubound(pcArr,2)
	For i=0 to intCount
		idparentcategory=pcArr(1,i)
		if idparentcategory="1" then %>
		    <option value='<%response.write pcArr(0,i)%>'><%=pcArr(2,i)%></option>
	    <%else
		For j=0 to intCount
		if Clng(pcArr(0,j))=Clng(idparentcategory) then
			parentDesc=pcArr(2,j)%>
			<option value='<%response.write pcArr(0,i)%>'><%response.write pcArr(2,i)&" - Parent: "&parentDesc %></option>
		<%exit for
		end if 
		Next
		end if
	Next
End if
set rstemp=nothing%>
</select>
</div>
</td>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
<th colspan="2">Display Options:</th>
</tr>
<tr> 
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
<td nowrap>Include store header &amp;  footer:</td>
<td>  
	<span style="color:#666;">This option has been removed. To remove the header and footer edit the file &quot;pc/StoreMap.asp&quot;.</span>
</td>
</tr>
<tr> 
<td nowrap>Use hardcoded font tags?</td>
<td>  
<input type="radio" value="1" name="storefont" checked class="clearBorder"> Yes
<input type="radio" value="0" name="storefont" class="clearBorder"> No</td>
</tr>
<tr>
	<td colspan="2">If yes, select the following options:</td>
</tr>
<tr> 
<td align="right">Font Name:</td>
<td>  
<select size="1" name="fontname">
<option selected value="Arial">Arial</option>
<option value="Verdana">Verdana</option>
<option value="Tahoma">Tahoma</option>
</select></td>
</tr>
<tr>
<td align="right">Font Size:</td>
<td>  
<select size="1" name="fontsize">
<option selected value="10">10px</option>
<option value="12">12px</option>
<option value="14">14px</option>
<option value="18">18px</option>
<option value="22">22px</option>
<option value="24">24px</option>
</select></td>
</tr>
<tr> 
<td align="right">Font Color:</td>
<td>  
<input type="text" name="fontcolor" id="fontcolor" size="10" value="#000000"> <input type="button" value="Choose" id="Choose" onClick="PcjsColorChooser('Choose','fontcolor','value')" name="button"></td>
</tr>
<tr> 
<td align="right">Link Color:</td>
<td>  
<input type="text" name="linkcolor" id="linkcolor" size="10" value="#0000FF"> <input type="button" value="Choose" id="Choose2" onClick="PcjsColorChooser('Choose','linkcolor','value')" name="button"></td>
</tr>
<tr> 
<td colspan="2">&nbsp;</td>
</tr>
<tr>
<td>Use H tags (H1, H2, etc):</td>
<td>  
<input type="radio" value="1" name="htags" class="clearBorder"> Yes
<input type="radio" value="0" name="htags" checked class="clearBorder"> No</td>
</tr>
<tr> 
<td colspan="2"><hr></td>
</tr>
<tr> 
<td colspan="2">  
<input name="submit" type="submit" class="submit2" value="Generate Store Map" onClick="pcf_Open_genStoreMap();">
</td>
</tr>
</table>
</form>
<%
'// Loading Window
'	>> Call Method with OpenHS();
	response.Write(pcf_ModalWindow("Generating Store Map. Please wait...", "genStoreMap", 300))
%>
<!--#include file="AdminFooter.asp"-->
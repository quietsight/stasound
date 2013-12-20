<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin="1*2*3*"%>
<!--#include file="adminv.asp"-->   
<!--#include file="../includes/utilities.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../pc/pcSeoFunctions.asp"-->
<% 
pageTitle="Generate Storefront Category Navigation" 
pageIcon="pcv4_icon_www.gif"
section="layout"

Dim connTemp,rsTemp
Dim query,query1,rsC,pcv_cats1,i,intCount,pcv_exportType,pcv_numprd
pcv_cats1=""
set StringBuilderObj = new StringBuilder

call opendb()

Function genCatInfor(tmp_IDCAT,tmp_CatName)
Dim tmp_showIcon,k,pcv_first,pcv_prdcount,pcv_back,pcv_end,tsT,tmp_query,pcv_CatName,pcv_prds1,tmp1,tmp2,pcv_CatNameFull,ProNameFull
	tmp1=0
	tmp2=0
	tmp_showIcon=0
	pcv_first=0
	pcv_prdcount=0
	maxProNameL = 22
	pcv_CatName=ClearHTMLTags2(tmp_CatName,0)
	pcv_CatNameFull=pcv_CatName
	If maxProNameL<len(pcv_CatName) then
		pcv_CatName=trim(left(pcv_CatName,maxProNameL)) & "..."
	End If
	if (pcv_exportType="0" or pcv_exportType="2") then
		pcv_end=intCount
		tmp1=0
		For k=0 to intCount
			pcv_back=pcv_end-k
			if cint(pcv_cats1(2,k))=tmp_IDCAT then
				tmp1=1
				exit for
			end if
			if pcv_cats1(2,pcv_back)=tmp_IDCAT then
				tmp1=1
				exit for
			end if
		Next
	end if
	
	if (pcv_exportType="1" or pcv_exportType="2") then
		tmp_query="SELECT products.idProduct,products.description FROM products INNER JOIN categories_products ON (products.idProduct=categories_products.idProduct AND products.removed=0 AND products.active=-1 AND products.configOnly=0) WHERE categories_products.idcategory=" & tmp_IDCAT & " ORDER BY categories_products.POrder asc,products.description ASC;"
		set rsT=connTemp.execute(tmp_query)
		if not rsT.eof then
			tmp2=1
			pcv_prds1=rsT.getRows()
			set rsT=nothing
		else
			tmp2=0
		end if
		set rsT=nothing
	end if
	if ((tmp1=1) and (pcv_exportType="0" or pcv_exportType="2")) or ((tmp2=1) and (pcv_exportType="1" or pcv_exportType="2")) then
		tmp_showIcon=1
	end if
	if pcvInt_htmlType=0 then
		StringBuilderObj.append "<tr><td valign=""top"">"
	else
		StringBuilderObj.append "<li>"		
	end if
	
	if pcvInt_htmlType=0 then
		if tmp_showIcon=1 then
			StringBuilderObj.append "<img align=""absmiddle"" name=""IMGCAT" & tmp_IDCAT & """ border=""0"" src=""images/btn_expand.gif"" onclick=""javascript:UpDown(" & tmp_IDCAT & ");"">"
		else
			StringBuilderObj.append "&nbsp;"
		end if
	end if

	'// SEO Links
	'// Build Navigation Category Link
	pcStrNavCatLink=pcv_CatNameFull & "-c" & tmp_IDCAT & ".htm"
	pcStrNavCatLink=removeChars(pcStrNavCatLink)
	if scSeoURLs<>1 then
		pcStrNavCatLink="viewCategories.asp?idCategory=" & tmp_IDCAT
	end if
	if pcvInt_htmlType=0 then
		StringBuilderObj.append "</td><td width=""100%"" valign=""top"" class="""&pcv_cssclass&""">&nbsp;<a href=""" & tempURL & pcStrNavCatLink & """>" & pcv_CatName & "</a>" & vbcrlf
	else
		if tmp_showIcon=1 then
		StringBuilderObj.append "<a href=""" & tempURL & pcStrNavCatLink & """ class=""MenuBarItemSubmenu"">" & pcv_CatName & "</a>" & vbcrlf	
		else
		StringBuilderObj.append "<a href=""" & tempURL & pcStrNavCatLink & """>" & pcv_CatName & "</a>" & vbcrlf	
		end if		
	end if
	'//
	
	if tmp_showIcon=1 then
		if pcvInt_htmlType=0 then	
			StringBuilderObj.append "<table border=""0"" width=""100%"" cellpadding=""0"" cellspacing=""0"" id=""SUB" & tmp_IDCAT & """ style=""display:none"">" & vbcrlf
		else
			StringBuilderObj.append "<ul>" & vbcrlf
		end if
		if ((tmp1=1) and (pcv_exportType="0" or pcv_exportType="2")) then
			For k=0 to intCount
				if cint(pcv_cats1(2,k))=tmp_IDCAT then
					pcv_first=1
					call genCatInfor(pcv_cats1(0,k),pcv_cats1(1,k))
				else
					if pcv_first=1 then
						exit for
					end if
				end if
			Next
		end if
		if ((tmp2=1) and (pcv_exportType="1" or pcv_exportType="2")) then
			if ubound(pcv_prds1,2)>pcv_numprd-1 then
				pcv_prdcount=pcv_numprd-1
			else
				pcv_prdcount=ubound(pcv_prds1,2)
			end if
			For k=0 to pcv_prdcount
				ProName=ClearHTMLTags2(pcv_prds1(1,k),0)
				ProNameFull=ProName
				If maxProNameL<len(ProName) then
					ProName=trim(left(ProName,maxProNameL)) & "..."
				End If


				pIntPrdId=pcv_prds1(0,k)
				'// SEO Links
				'// Build Navigation Product Link
				pcStrNavPrdLink=ProNameFull & "-" & tmp_IDCAT & "p" & pIntPrdId & ".htm"
				pcStrNavPrdLink=removeChars(pcStrNavPrdLink)
				if scSeoURLs<>1 then
					pcStrNavPrdLink="viewPrd.asp?idcategory=" & tmp_IDCAT & "&idproduct=" & pcv_prds1(0,k)
				end if
				'//
				if pcvInt_htmlType=0 then	
					StringBuilderObj.append "<tr><td>&nbsp;</td><td class="""&pcv_cssclass&"""><a href=""" & tempURL & pcStrNavPrdLink & """>" & ProName & "</a></td></tr>" & vbcrlf
				else
					StringBuilderObj.append "<li><a href=""" & tempURL & pcStrNavPrdLink & """>" & ProName & "</a></li>" & vbcrlf				
				end if
			Next
			if ubound(pcv_prds1,2)>pcv_numprd-1 then
				'// SEO Links
				'// Build Navigation Category Link
				pcStrNavCatLink=pcv_CatNameFull & "-c" & tmp_IDCAT & ".htm"
				pcStrNavCatLink=removeChars(pcStrNavCatLink)
				if scSeoURLs<>1 then
					pcStrNavCatLink="viewCategories.asp?idCategory=" & tmp_IDCAT
				end if
				'//
				if pcvInt_htmlType=0 then	
					StringBuilderObj.append "<tr><td>&nbsp;</td><td class="""&pcv_cssclass&"""><a href=""" & tempURL & pcStrNavCatLink & """>More Products...</a></td></tr>" & vbcrlf
				else
					StringBuilderObj.append "<li><a href=""" & tempURL & pcStrNavCatLink & """>More Products...</a></li>" & vbcrlf				
				end if
			end if
		end if
		if pcvInt_htmlType=0 then		
			StringBuilderObj.append "</table>" & vbcrlf
		else
			StringBuilderObj.append "</ul>" & vbcrlf		
		end if
	end if
	if pcvInt_htmlType=0 then
		StringBuilderObj.append "</td></tr>" & vbcrlf
	else
		StringBuilderObj.append "</li>" & vbcrlf	
	end if
End Function
%>
<!--#include file="AdminHeader.asp"-->
<%
'// Initialize the Prototype.js files
Dim pcv_strPrototype, pcvInt_htmlType, pcvTempQuery
Response.Write(pcf_InitializePrototype())


IF request("action")="add" THEN

	pcv_exportType=request("exportType")
	if pcv_exportType="" then
		pcv_exportType="0"
	end if
	if pcv_exportType="3" then
		pcvTempQuery = "AND categories.idParentCategory=1 "
		else
		pcvTempQuery = ""
	end if
	
	pcv_numprd=request("prdcount")
	if pcv_numprd="" or pcv_numprd="0" then
		pcv_numprd=5
	end if
	
	pcvInt_htmlType=request("htmlType") ' 1 = unordered list, 2 = tables
		if not validNum(pcvInt_htmlType) then pcvInt_htmlType=0
		
	if pcvInt_htmlType=1 then		
		pcIntSpryNav = request("spryNav") ' 1 = horizontal, 2 = vertical
		pcv_topULid = request("spryNavID")
		if not validNum(pcIntSpryNav) or pcIntSpryNav=0  then
			session("pcIntSpryNav")=0
			'// Unordered list, own settings
			pcv_topULid=request("topULid")
			session("pcv_topULid")=pcv_topULid
			pcv_topULclass=request("topULclass")
			session("pcv_topULclass")=pcv_topULclass
		else
			session("pcIntSpryNav")=pcIntSpryNav
			if pcIntSpryNav = 1 then
				pcv_topULclass = "MenuBarHorizontal"
				else
				pcv_topULclass = "MenuBarVertical"
			end if					
		end if	
	else
		pcv_cssclass=request("cssclass")
		session("pcv_cssclass")=pcv_cssclass
		session("pcv_topULid")=""
		session("pcv_topULclass")=""
		session("pcIntSpryNav")=0
	end if	
	
	pcvStr_linkType=request("linkType")
		if pcvStr_linkType="absolute" then
			tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/"),"//","/")
			tempURL=replace(tempURL,"http:/","http://")
		else
			tempURL=""
		end if

	
	query="SELECT idCategory,categoryDesc,idParentCategory FROM categories WHERE categories.idCategory<>1 " & pcvTempQuery & "AND categories.iBTOhide=0 AND categories.pccats_RetailHide=0 ORDER BY categories.idParentCategory ASC,categories.priority ASC,categories.categoryDesc ASC;"
	set rsC=connTemp.execute(query)
	if not rsC.eof then
		pcv_cats1=rsC.GetRows()
		intCount=ubound(pcv_cats1,2)
		set rsC=nothing
		if pcvInt_htmlType=0 then
			StringBuilderObj.append "<table border=""0"" width=""100%"" cellpadding=""0"" cellspacing=""0"">" & vbcrlf
		else
			StringBuilderObj.append "<ul id="""&pcv_topULid&""" class="""&pcv_topULclass&""">" & vbcrlf
		end if		
		For i=0 to intCount
			if pcv_cats1(2,i)="1" then
				call genCatInfor(pcv_cats1(0,i),pcv_cats1(1,i))
			else
				exit for
			end if
		Next
		if pcvInt_htmlType=0 then
			StringBuilderObj.append "</table>"
		else
			StringBuilderObj.append "</ul>"			
		end if
	end if
	set rsC=nothing
	
	if PPD="1" then
		pcStrFolder=Server.Mappath ("/"&scPcFolder&"/pc")
	else
		pcStrFolder=server.MapPath("../pc")
	end if

	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	Set a=fs.CreateTextFile(pcStrFolder & "\inc_RetailCatMenu.inc",True)
	a.Write(StringBuilderObj.toString())
	a.Close
	Set a=Nothing
	Set fs=Nothing
	Set StringBuilderObj=Nothing
tmp1=0
tmp2=0
pcv_cats1=""
set StringBuilderObj = new StringBuilder

	query="SELECT idCategory,categoryDesc,idParentCategory FROM categories WHERE categories.idCategory<>1 AND categories.iBTOhide=0 ORDER BY categories.idParentCategory ASC,categories.priority ASC,categories.categoryDesc ASC;"
	set rsC=connTemp.execute(query)
	if not rsC.eof then
		pcv_cats1=rsC.GetRows()
		intCount=ubound(pcv_cats1,2)
		set rsC=nothing
		if pcvInt_htmlType=0 then
			StringBuilderObj.append "<table border=""0"" width=""100%"" cellpadding=""0"" cellspacing=""0"">" & vbcrlf
		else
			StringBuilderObj.append "<ul id="""&pcv_topULid&""" class="""&pcv_topULclass&""">" & vbcrlf	
		end if
		For i=0 to intCount
			if pcv_cats1(2,i)="1" then
				call genCatInfor(pcv_cats1(0,i),pcv_cats1(1,i))
			else
				exit for
			end if
		Next
		if pcvInt_htmlType=0 then
			StringBuilderObj.append "</table>"
		else
			StringBuilderObj.append "</ul>"
		end if
	end if
	set rsC=nothing

	if PPD="1" then
		pcStrFolder=Server.Mappath ("/"&scPcFolder&"/pc")
	else
		pcStrFolder=server.MapPath("../pc")
	end if

	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	Set a=fs.CreateTextFile(pcStrFolder & "\inc_WholeSaleCatMenu.inc",True)
	a.Write(StringBuilderObj.toString())
	a.Close
	Set a=Nothing
	Set fs=Nothing
	Set StringBuilderObj=Nothing
%>
<table class="pcCPcontent">
<tr>
	<td align="center">
		<div class="pcCPmessageSuccess">
			New Storefront Category Navigation was created successfully!
			<br /><br />
			<a href="../pc/viewcategories.asp" target="_blank">View Storefront</a>&nbsp;|&nbsp;
            <%
			if session("pcIntSpryNav")=1 then
			%>
                <a href="genCatSpryPreviewH.asp" target="_blank">Preview SPRY Horizontal Menu</a>
                &nbsp;|&nbsp;
                <%
				elseif session("pcIntSpryNav") = 2 then
				%>
                <a href="genCatSpryPreviewV.asp" target="_blank">Preview SPRY Vertical Menu</a>
                &nbsp;|&nbsp;
                <%
				else
			end if
			%>
            <a href="genCatNavigation.asp">Generate New Navigation</a>
		</div>
	</td>
</tr>
<tr>
	<td class="pcSpacer">&nbsp;</td>
</tr>
</table>

<%
session("pcIntSpryNav")=""
ELSE
%>

<form method="post" name="form1" action="genCatNavigation.asp?action=add" class="pcForms">
	<table class="pcCPcontent">
	<tr>
		<td colspan="2">
        	<h2>How it works</h2>
            <div>
			ProductCart will generate a <u>static</u> file to store your navigation links. This improves storefront performance (less database queries). Remember to rerun this feature when you add/edit categories (and products if included in the navigation). Your &quot;<strong>header.asp</strong>&quot; or &quot;<strong>footer.asp</strong>&quot; file must include the file &quot;<strong>inc_catsmenu.asp</strong>&quot; in order for the navigation to show.<a href="http://wiki.earlyimpact.com/how_to/add_category_navigation" target="_blank"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature" width="16" height="16" border="0"></a>
            </div>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcSpacer">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2">
        <h2>Create or Update your storefront category navigation</h2>
        <p>When creating your navigation file you have three options:</p></td>
	</tr>
	<tr>
		<td align="right" width="10%"><input type="radio" name="exportType" value="0" checked class="clearBorder"></td>
		<td width="90%">Include categories and sub-categories, but no products</td>
	</tr>
	<tr>
		<td align="right"><input type="radio" name="exportType" value="2" class="clearBorder"></td>
		<td>Include categories, sub-categories, and their products</td>
	</tr>
	<tr>
		<td align="right"><input type="radio" name="exportType" value="1" class="clearBorder"></td>
		<td>Include only top-level categories and their products</td>
	</tr>
	<tr>
		<td align="right"><input type="radio" name="exportType" value="3" class="clearBorder"></td>
		<td>Include only top-level categories</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td align="right"><input type="text" name="prdcount" size="3" value="5"></td>
		<td>Maximum of products per category (e.g. 5)<br />
		A &quot;More products...&quot; link is automatically added if there are more products in the category than the number specified here.</td>
	</tr>
	<tr>
		<td colspan="2" class="pcSpacer"><hr></td>
	</tr>
	<tr>
		<td colspan="2">Do you want to use relative or absolute links?&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=445')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
	</tr>
	<tr>
		<td align="right"><input type="radio" name="linkType" value="relative" class="clearBorder"></td>
        <td>Relative Links</td>
    </tr>
    <tr>
    	<td align="right"><input type="radio" name="linkType" value="absolute" class="clearBorder" checked="checked"></td>
        <td>Absolute Links</td>
	</tr>
	<tr>
		<td colspan="2" class="pcSpacer"><hr></td>
	</tr>
	<tr>
		<td colspan="2">What kind of HTML tags would you like to create?&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=462')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
	</tr>
	<tr>
		<td align="right" valign="top"><input type="radio" name="htmlType" value="0" class="clearBorder" onClick="document.getElementById('cssOptions').style.display='none';"></td>
        <td valign="top">Table rows and cells (<em>ProductCart v3 setting</em>)
        <div style="margin-top: 6px;">
        	<table class="pcCPcontent">
        	<tr><td>Style table cells with this CSS class: <input type="text" name="cssclass" size="20" value="<%=session("pcv_cssclass")%>">&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=444')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td></tr>
            </table>
        </div>
        </td>
    </tr>
    <tr>
    	<td align="right" valign="top"><input type="radio" name="htmlType" value="1" class="clearBorder" onClick="document.getElementById('cssOptions').style.display='';"></td>
        <td valign="top">Unordered lists (<em>e.g. for SPRY navigation</em>)
        <table class="pcCPcontent" id="cssOptions" style="display:none; margin-top: 6px;">
			<tr> 
				<td valign="top" colspan="2">Prepare for <strong>SPRY menu bar</strong>: <a href="http://labs.adobe.com/technologies/spry/samples/menubar/MenuBarSample.html" target="_blank">Examples</a>, <a href="http://labs.adobe.com/technologies/spry/articles/menu_bar/index.html" target="_blank">Documentation</a></td>
            </tr>
            <tr>
                <td colspan="2">
                <input type="radio" name="spryNav" value="1"> Spry Horizontal Menu Bar <br />
                <input type="radio" name="spryNav" value="2"> Spry Vertical Menu Bar <br />
                <input type="radio" name="spryNav" value="0" checked> None
                </td>
			</tr>
			<tr> 
				<td align="right">ID of SPRY menu bar:</td>
                <td>
                <input type="text" name="spryNavID" value="<%if session("pcvULID")<>"" then response.write(session("pcvULID")) else response.write "menubar99" end if%>" size="30"> <span class="pcSmallText">See SPRY documentation for details.</span>
                </td>
			</tr>
            <tr><td colspan="2"><img src="images/pc_admin.gif" width="85" height="19" alt="Alternative setting" vspace="10"></td></tr>
            <tr><td colspan="2">&nbsp;... or use <strong>your own CSS</strong>:</td></tr>
        	<tr><td align="right">Top UL tag ID:</td><td><input type="text" name="topULid" value="<%=session("pcv_topULid")%>" size="30"></td></tr>
            <tr><td align="right">Top UL tag Class:</td><td><input type="text" name="topULclass" value="<%=session("pcv_topULclass")%>" size="30"></td></tr>
            <tr><td colspan="2" class="pcCPspacer"></td></tr>
            <tr>
            	<td colspan="2">
                	<div class="pcCPmessage">NOTE: make sure that your store interface (header.asp, footer.asp) contains the CSS needed to style the unordered list. If you are using SPRY, you can reference the SPRY documents located in the <em>includes/spry</em> folder. See the <a href="http://wiki.earlyimpact.com/how_to/add_category_navigation" target="_blank">ProductCart WIKI</a> for details.</div>
                </td>
            </tr>
        </table>
        </td>
	</tr>
	<tr>
		<td colspan="2" class="pcSpacer"></td>
	</tr>
	<tr>
		<td colspan="2" class="pcSpacer"><hr></td>
	</tr>
	<tr> 
		<td colspan="2">
			<input name="submit" type="submit" class="submit2" value="Generate Storefront Category Navigation" onClick="pcf_Open_genCatNav();">
			<%
            '// Loading Window
            '	>> Call Method with OpenHS();
            response.Write(pcf_ModalWindow("Generating category navigation. Please wait...", "genCatNav", 300))
            %>
		</td>
	</tr>
	</table>
</form>
<%
END IF
call closedb()%>
<!--#include file="AdminFooter.asp"-->
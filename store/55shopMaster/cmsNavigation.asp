<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Section="layout" %>
<%PmAdmin=11%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../pc/pcSeoFunctions.asp"-->

<%
pageTitle="Generate Content Pages Navigation"
dim query, conntemp, rs, pcvParentPageName, pcArray, pcInt_Parent, pcInt_Published, pcInt_Inactive, pcInt_Active, pcv_IntNumrowsCount, pcv_IntNumrows, m, n
%>

<!--#include file="AdminHeader.asp"-->
<%
call openDb()

' Load Pages
query="SELECT pcCont_IDPage, pcCont_PageName FROM pcContents WHERE pcCont_Parent=0 AND pcCont_InActive=0 AND pcCont_MenuExclude=0 ORDER BY pcCont_Order, pcCont_Parent, pcCont_PageName ASC;"
set rstemp=server.CreateObject("ADODB.RecordSet")
set rstemp=conntemp.execute(query)

if err.number <> 0 then
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?error="& Server.Urlencode("Error loading parent content pages") 
end If
%>

<script language="JavaScript">
<!--
    function chgWin(file,window) {
    msgWindow=open(file,window,'scrollbars=yes,resizable=yes,width=500,height=500');
    if (msgWindow.opener == null) msgWindow.opener = self;
}
//-->
</script>

<%
'Check URL

	SPath1=Request.ServerVariables("PATH_INFO")
	mycount1=0
	do while mycount1<2
	if mid(SPath1,len(SPath1),1)="/" then
	mycount1=mycount1+1
	end if
	if mycount1<2 then
	SPath1=mid(SPath1,1,len(SPath1)-1)
	end if
	loop
	SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & SPath1
	
	if Right(SPathInfo,1)="/" then
	pcv_strViewContents=SPathInfo & "pc/viewContent.asp"
	else
	pcv_strViewContents=SPathInfo & "/pc/viewContent.asp"
	end if
%>

<form name="form1" action="cmsNavigation.asp" method="post" class="pcForms">
	<table class="pcCPcontent">
    
		<%
		'-------------------------
		' NO Content Pages Found
		'-------------------------
		IF rstemp.eof THEN
			set rstemp=nothing
		%>
			<tr> 
				<td align="center">
					<div class="pcCPmessage">No Content Pages Found. <a href="cmsAddEdit.asp">Add New</a>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=436')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></div>
				</td>
			</tr>                  
		<% 
		ELSE
		
		'-------------------------
		' NO FORM Submitted
		'-------------------------	
			
			IF request("submit")="" THEN
		%>
			<tr> 
				<td colspan="2">
                	<h2>Which pages are included</h2>
                	The system will generate an unordered list with all content pages that are:
                	<ul>
                    	<li>Active</li>
                        <li>Not excluded from the navigation</li>
                    </ul>
				</td>
			</tr>  
			<tr> 
				<td colspan="2">
                	<h2>SPRY Navigation</h2>
				</td>
			</tr>  
			<tr> 
				<td valign="top">Prepare for SPRY menu bar:<br /><a href="http://labs.adobe.com/technologies/spry/samples/menubar/MenuBarSample.html" target="_blank">Examples</a>, <a href="http://labs.adobe.com/technologies/spry/articles/menu_bar/index.html" target="_blank">Documentation</a></td>
                <td>
                <input type="radio" name="spryNav" value="1"> Spry Horizontal Menu Bar <br />
                <input type="radio" name="spryNav" value="2"> Spry Vertical Menu Bar <br />
                <input type="radio" name="spryNav" value="0" checked> None
                </td>
			</tr>
			<tr> 
				<td>ID of SPRY menu bar:</td>
                <td>
                <input type="text" name="spryNavID" value="menubar1" size="30"> <span class="pcSmallText">See SPRY documentation for details.</span>
                </td>
			</tr>
            <tr>
            	<td colspan="2" class="pcCPspacer"></td>
            </tr>
			<tr> 
				<td colspan="2">
                	<h2>Advanced Settings</h2>
                    You can assign a CSS class to the most relevant elements in the unordered list. These settings are <u>ignored</u> if using the SPRY option above.
				</td>
			</tr>  
			<tr> 
				<td nowrap>&lt;UL&gt; Tag ID</td>
                <td><input type="text" name="ulid" value="" size="30"> <span class="pcSmallText">This is often referenced in JavaScript used to activate the menu.</span>
			</tr>
			<tr> 
				<td nowrap>&lt;UL&gt; CSS Class</td>
                <td><input type="text" name="ulclass" value="" size="30">
			</tr> 
			<tr> 
				<td nowrap>Top-level &lt;LI&gt; CSS Class</td>
                <td><input type="text" name="liclass" value="" size="30">
			</tr> 
			<tr> 
				<td nowrap>Top-level &lt;LI&gt; with Sub-Items CSS Class</td>
                <td><input type="text" name="lisbclass" value="" size="30">
			</tr>  
            <tr>
            	<td colspan="2" class="pcCPspacer"></td>
            </tr>
            <tr>
            	<td colspan="2"><input type="submit" name="submit" value="Generate Content Pages Navigation" onclick="return(confirm('You are about to overwrite the existing Content Pages navigation with a new list of pages. Back up the file pc/cmsNavigationLinks.inc if you need to keep a copy of the existing navigation. Are you sure you want to continue?'));" class="submit2"></td>
            </tr>
		
		<%
			ELSE
			'-------------------------
			' BUILD Content Pages Navigation
			'-------------------------
			
			pcIntSpryNav = request("spryNav") ' 1 = horizontal, 2 = vertical
			pcvULID = request("spryNavID")
			if not validNum(pcIntSpryNav) or pcIntSpryNav=0  then
				pcvULclass = request("ulclass")
				pcvULID = request("ulid")
				pcvLIclass = request("liclass")
				pcvLIsbclass = request("lisbclass")
			else
				if pcIntSpryNav = 1 then
					pcvULclass = "MenuBarHorizontal"
					else
					pcvULclass = "MenuBarVertical"
				end if					
				pcvLIclass = ""
				pcvLIsbclass = ""
			end if
					
			strNavigationStart = "<ul id='" & pcvULID & "' class='" & pcvULclass & "'>"
			strNavigationSubMenu = "<ul>"
			strNavigationEnd = "</ul>"
			
			pcArray = rstemp.getRows()
			set rstemp=nothing
			
			pcv_IntNumrows = UBound(pcArray, 2)

			pcv_IntNumrowsCount=0
			
			FOR m = 0 to pcv_IntNumrows

				pcv_IntNumrowsCount=pcv_IntNumrowsCount+1
				pcv_lngIDPage= pcArray(0,m)
				pcv_strPageName = pcArray(1,m)
				
				'// Check to see if there are subpages
				query="SELECT pcCont_IDPage FROM pcContents WHERE pcCont_Parent=" & pcv_lngIDPage & " AND pcCont_InActive=0 AND pcCont_MenuExclude=0;"
				set rstemp = Server.CreateObject("ADODB.Recordset")
				set rstemp = conntemp.execute(query)
				if err.number <> 0 then
					set rstemp=nothing
					call closedb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error loading child pages") 
				end If
				if not rstemp.eof then
					pcIntHasSubPages = 1
					else
					pcIntHasSubPages = 0
				end if
				
				'// SEO Links
				'// Build Navigation Product Link
				if scSeoURLs=1 then
					pcStrCntPageLink=pcv_strPageName & "-d" & pcv_lngIDPage & ".htm"
					pcStrCntPageLink=removeChars(pcStrCntPageLink)
					pcStrCntPageLink=SPathInfo & "pc/" & pcStrCntPageLink
				else
					pcStrCntPageLink=pcv_strViewContents&"?idpage="&pcv_lngIDPage
				end if
				'//			
				
				
				if pcv_IntNumrowsCount=1 then strNavigation=strNavigationStart&strNavigation & Vbcrlf
				
				if pcIntHasSubPages = 1 then ' Don't close the list item if there is a submenu
					if pcIntSpryNav > 0 then
						strNavigation = strNavigation & "<li><a href=" & pcStrCntPageLink & " class='MenuBarItemSubmenu'>" & pcv_strPageName & "</a>" & Vbcrlf
					else
						strNavigation = strNavigation & "<li class='" & pcvLIclass & "'><a href=" & pcStrCntPageLink & ">" & pcv_strPageName & "</a>" & Vbcrlf
					end if
				else
					strNavigation = strNavigation & "<li class='" & pcvLIclass & "'><a href=" & pcStrCntPageLink & ">" & pcv_strPageName & "</a></li>" & Vbcrlf
				end if	
					
					'// If there are subpages, build menu subsection
					if pcIntHasSubPages = 1 then
					
						query="SELECT pcCont_IDPage, pcCont_PageName FROM pcContents WHERE pcCont_Parent=" & pcv_lngIDPage & " ORDER BY pcCont_Order, pcCont_PageName ASC;"
						set rstemp = Server.CreateObject("ADODB.Recordset")
						set rstemp = conntemp.execute(query)
						if err.number <> 0 then
							set rstemp=nothing
							call closedb()
							response.redirect "techErr.asp?error="& Server.Urlencode("Error loading child pages") 
						end If
						
						pcArraySubPages = rstemp.getRows()
						set rstemp=nothing
						
						pcv_IntNumrowsSP = UBound(pcArraySubPages, 2)
						pcv_IntNumrowsCountSP=0
						
						FOR n = 0 to pcv_IntNumrowsSP
			
							pcv_IntNumrowsCountSP=pcv_IntNumrowsCountSP+1
							pcv_lngIDPageSP= pcArraySubPages(0,n)
							pcv_strPageNameSP = pcArraySubPages(1,n)
							
							'// SEO Links
							'// Build Navigation Product Link
							if scSeoURLs=1 then
								pcStrCntPageLink=pcv_strPageNameSP & "-d" & pcv_lngIDPageSP & ".htm"
								pcStrCntPageLink=removeChars(pcStrCntPageLink)
								pcStrCntPageLink=SPathInfo & "pc/" & pcStrCntPageLink
							else
								pcStrCntPageLink=pcv_strViewContents&"?idpage="&pcv_lngIDPageSP
							end if
							'//			
							
							if pcv_IntNumrowsCountSP=1 then strNavigation=strNavigation & strNavigationSubMenu & Vbcrlf
							if pcIntSpryNav > 0 then
								strNavigation = strNavigation & "<li><a href=" & pcStrCntPageLink & ">" & pcv_strPageNameSP & "</a></li>" & Vbcrlf
								else
								strNavigation = strNavigation & "<li class='" & pcvLIsbclass & "'><a href=" & pcStrCntPageLink & ">" & pcv_strPageNameSP & "</a></li>" & Vbcrlf								
							end if
							'// Close submenu and add closing list item for parent menu item
							if (pcv_IntNumrowsCountSP-1)=pcv_IntNumrowsSP then strNavigation=strNavigation & strNavigationEnd & "</li>"
							
						NEXT
							
					end if ' End if there are subpages
				
				if (pcv_IntNumrowsCount-1)=pcv_IntNumrows then strNavigation=strNavigation & strNavigationEnd
				
				NEXT
				
				set rstemp = nothing
				call closeDb()
				
				'-------------------------
				' WRITE to FILE
				'-------------------------
				
				if PPD="1" then
					pcStrFolder=Server.Mappath ("/"&scPcFolder&"/pc")
				else
					pcStrFolder=server.MapPath("../pc")
				end if
			
				Set fs=Server.CreateObject("Scripting.FileSystemObject")
				Set a=fs.CreateTextFile(pcStrFolder & "\cmsNavigationLinks.inc",True)
				a.Write(strNavigation)
				a.Close
				Set a=Nothing
				Set fs=Nothing
				%>
            <tr>
                <td>
                <div class="pcCPmessageSuccess">Content Pages navigation successfully saved to &quot;cmsNavigationLinks.inc&quot;</div>
                The unordered list containing the Content Pages navigation has been saved to the file <strong>cmsNavigationLinks.inc</strong> located in the &quot;<strong>pc</strong>&quot; folder. There are many ways to use an ordered list to create a navigation menu.
                </td>
            </tr>
            <tr>
                <td class="pcCPspacer"></td>
            </tr>
            <tr>
                <td>
                You can find the raw HTML code for the unordered list (UL) that contains the selected Content Pages below:
                <br /><br />
                <textarea cols="80" rows="10"><%=strNavigation %></textarea>
                <br /><br />
                <%
				if pcIntSpryNav = 1 then
				%>
                <a href="cmsSpryPreviewH.asp" target="_blank">See SPRY Horizontal Menu Bar Preview</a>
                &nbsp;|&nbsp;
                <%
				elseif pcIntSpryNav = 2 then
				%>
                <a href="cmsSpryPreviewV.asp" target="_blank">See SPRY Vertical Menu Bar Preview</a>
                &nbsp;|&nbsp;
                <%
				else
				%>
                <a href="cmsPreview.asp" target="_blank">View HTML</a>
                &nbsp;|&nbsp;
                <%
				end if
				%>
                <a href="cmsNavigation.asp">Back</a>
                </td>
            </tr>
		<%
			END IF
		END IF
		%>					
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->
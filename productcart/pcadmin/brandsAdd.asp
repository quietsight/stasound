<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add New Brand" %>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/validation.asp" --> 
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
dim query, conntemp, rs
dim pcvBrandsDescription, pcvBrandsSDescription, pcIntBrandsActive, pcIntSubBrandsView, pcvProductsView, pcIntBrandsParent, pcvBrandsMetaTitle, pcvBrandsMetaDesc, pcvBrandsMetaKeywords, pcvBrandsBrandLogoLg

'// Process New Brand
if request("action")="add" then
	BrandName=pcf_SanitizeApostrophe(request.form("BrandName"))
	BrandLogo=request.form("image")
	pcvBrandsBrandLogoLg=request.form("largeimage")
	pcvBrandsDescription=pcf_SanitizeApostrophe(request.form("pcBrandsDescription"))
	pcvBrandsSDescription=pcf_SanitizeApostrophe(request.form("pcBrandsSDescription"))
	pcIntSubBrandsView=request.form("intSubBrandsView")
	pcvProductsView=request.form("pcProductsView")
	pcIntBrandsActive=request.form("pcBrandsActive")
	pcIntBrandsParent=request.form("pcBrandsParent")
	pcvBrandsMetaTitle=getUserInput(request.form("pcBrandsMetaTitle"),0)
	pcvBrandsMetaDesc=getUserInput(request.form("pcBrandsMetaDesc"),0)
	pcvBrandsMetaKeywords=getUserInput(request.form("pcBrandsMetaKeywords"),0)
	
	if not validNum(pcIntSubBrandsView) then pcIntSubBrandsView=0
	if not validNum(pcIntBrandsActive) then pcIntBrandsActive=1
	if not validNum(pcIntBrandsParent) then pcIntBrandsParent=0
	
	call opendb()
	query="INSERT INTO Brands (BrandName,BrandLogo,pcBrands_Description,pcBrands_SDescription,pcBrands_SubBrandsView,pcBrands_ProductsView,pcBrands_Active, pcBrands_Parent, pcBrands_MetaTitle, pcBrands_MetaDesc, pcBrands_MetaKeywords, pcBrands_BrandLogoLg) VALUES ('" & BrandName & "','" & BrandLogo & "','" & pcvBrandsDescription & "','" & pcvBrandsSDescription& "'," & pcIntSubBrandsView & ",'" & pcvProductsView& "'," & pcIntBrandsActive & "," & pcIntBrandsParent & ",'" & pcvBrandsMetaTitle & "','" & pcvBrandsMetaDesc & "','" & pcvBrandsMetaKeywords & "','" & pcvBrandsBrandLogoLg & "')"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error saving the new brand") 
	else
		set rs=nothing
		call closeDb()
		response.redirect "BrandsAdd.asp?s=1&msg="&Server.URLEncode("New brand added successfully.")
	end if
end if

'// Show Add New Brand Page
%>
<!--#include file="AdminHeader.asp"-->
<!--#include file="pcv4_showMessage.asp"-->

<link href="../includes/spry/SpryTabbedPanels-PP.css" rel="stylesheet" type="text/css" />
<script src="../includes/spry/SpryTabbedPanels.js" type="text/javascript"></script>
<script src="../includes/spry/SpryURLUtils.js" type="text/javascript"></script>
<script type="text/javascript"> var params = Spry.Utils.getLocationParamsAsObject(); </script>

<script language="JavaScript">
<!--
	function newWindow(file,window) {
			msgWindow=open(file,window,'resizable=no,width=400,height=500');
			if (msgWindow.opener == null) msgWindow.opener = self;
	}

	function Form1_Validator(theForm)
	{
		if (theForm.BrandName.value == "")
			{
				 alert("Please enter a name for the new brand.");
					theForm.BrandName.focus();
					return (false);
			}
	return (true);
	}

	function chgWin(file,window)
	{
		msgWindow=open(file,window,'scrollbars=yes,resizable=yes,width=500,height=500');
		if (msgWindow.opener == null) msgWindow.opener = self;
	}
//-->
</script> 

	<form action="BrandsAdd.asp?action=add" method="post" name="hForm" onSubmit="return Form1_Validator(this)" class="pcForms">
		<%
		'// TABBED PANELS - MAIN DIV START
		%>
	  <div id="TabbedPanels1" class="VTabbedPanels">
		
		<%
		'// TABBED PANELS - START NAVIGATION
		%>
			<ul class="TabbedPanelsTabGroup">
				<li class="TabbedPanelsTab" tabindex="100">Name, Parent &amp; Images</li>
				<li class="TabbedPanelsTab" tabindex="200">Descriptions</li>
				<li class="TabbedPanelsTab" tabindex="300">Display &amp; Other Settings</li>				
				<li class="TabbedPanelsTab" tabindex="400">Meta Tags</li>
				<li class="TabbedPanelsTabButtons" tabindex="1000">
					<input name="Submit" type="submit" value="Add" class="submit2"><br />
                    <div style="margin-top: 5px"><input type="button" value="Manage Brands" onClick="document.location.href='BrandsManage.asp';"></div>
				</li>
			</ul>
			
		<%
		'// TABBED PANELS - END NAVIGATION
		
		'// TABBED PANELS - START PANELS
		%>
		
			<div class="TabbedPanelsContentGroup">
			
			<%
			'// =========================================
			'// FIRST PANEL - START - Name, Descriptions, Images
			'// =========================================
			%>
				<div class="TabbedPanelsContent">
				
					<table class="pcCPcontent">
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Brand Name, Images &amp; Parent (if any)</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<td width="20%" align="right">Brand Name:</td>
							<td width="80%"><input name="BrandName" id="brandName" type="text" value="" size="40" tabindex="101"></td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td align="right">Small Brand Logo:</td>
							<td>
								<input type="text" name="image" value="" size="40" tabindex="102"><a href="#" onClick="chgWin('../pc/imageDir.asp?ffid=image&fid=hForm','window2')"><img src="images/search.gif" alt="Locate previously uploaded images" width="16" height="16" border=0 hspace="3"></a>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=439')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
							</td>
						</tr>
						<tr> 
							<td align="right">Large Brand Logo:</td>
							<td> 
				        		<input type="text" name="largeimage" value="" size="40" tabindex="103"><a href="#" onClick="chgWin('../pc/imageDir.asp?ffid=largeimage&fid=hForm','window2')"><img src="images/search.gif" alt="Locate previously uploaded images" width="16" height="16" border=0 hspace="3"></a>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=439')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
							</td>
						</tr>
						<tr> 
							<td>&nbsp;</td>
							<td><p> 
								<!--#include file="uploadresize/checkImgUplResizeObjs.asp"-->
								<%If HaveImgUplResizeObjs=1 then%>
								<%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_8")%>&nbsp;<a href="#" onClick="window.open('uploadresize/catResizea.asp','popup','toolbar=no,status=no,location=no,menubar=no,height=350,width=400,scrollbars=no'); return false;">click here</a>.
								<% Else %>
									<%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_9")%>&nbsp;<a href="#" onClick="window.open('imageuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')">click here</a>.
								<% End If %>
							</td>
						</tr>
						<tr>
							<td colspan="2"><hr></td>
						</tr> 
					<tr> 
						<td align="right" valign="top" nowrap>Parent Brand:</td>
						<td>
                        	<%
							call OpenDb()
								Dim pcBrandsParentExist
								query="SELECT idbrand, BrandName FROM Brands WHERE pcBrands_Parent=0 ORDER BY BrandName ASC"
								set rs=Server.CreateObject("ADODB.Recordset")
								set rs=connTemp.execute(query)
								if rs.EOF then
									pcBrandsParentExist=0
								else
									pcBrandsParentExist=1
									pcBrandsArr=rs.getRows()
								end if
								set rs=nothing
							call closeDb()
							if pcBrandsParentExist=0 then
							%>
                                No brands available.
                                <br />
                                First add a brand, then you can use it as a &quot;Parent&quot; of another brand.
                            <%
							else
							%>
                            	<select name="pcBrandsParent" tabindex="104">
                                	<option value="0">None</option>
                            <%
                                intCount=ubound(pcBrandsArr,2)
                                For m=0 to intCount %>
									<option value='<%response.write pcBrandsArr(0,m)%>'><%=pcBrandsArr(1,m)%></option>
                            <%
                                Next
                            %>
								</select>
                            <%
							end if
							%>
						</td>
					</tr>
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>

					</table>
					
				</div>
			<%
			'// =========================================
			'// FIRST PANEL - END
			'// =========================================

			'// =========================================
			'// SECOND PANEL - START - Descriptions
			'// =========================================
			%>
				<div class="TabbedPanelsContent">

					<table class="pcCPcontent">	
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Descriptions:&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=440')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr valign="top">
							<td>Short Description:
								<br />
								<input type="button" value="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_401")%>" onClick="newWindow('pop_HtmlEditor.asp?fi=pcBrandsSDescription','window2')" class="ibtnGrey">	
							</td>			
							<td>
								<textarea name="pcBrandsSDescription" cols="50" rows="6" tabindex="201"></textarea>
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr valign="top">
							<td>Long Description:
							<br />
							<input type="button" value="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_401")%>" onClick="newWindow('pop_HtmlEditor.asp?fi=pcBrandsDescription','window2')" class="ibtnGrey">
							</td>
							<td>
							<textarea name="pcBrandsDescription" cols="50" rows="6" tabindex="202"></textarea>						
							</td>
						</tr>						
					</table>
					
				</div>
			<%
			'// =========================================
			'// SECOND PANEL - END
			'// =========================================

			'// =========================================
			'// THIRD PANEL - START - Display settings
			'// =========================================
			%>
				<div class="TabbedPanelsContent">

					<table class="pcCPcontent">
                        <tr>
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr>
                            <th colspan="2">Display &amp; Other Settings</th>
                        </tr>
                        <tr>
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr>
                            <td width="20%" valign="top" nowrap>Display Sub-brands:</td>
                            <td>
                                <select name="intSubBrandsView" tabindex="301">
                                    <option value="2">Default (like brands)</option>
                                    <option value="0">List (no images)</option>
                                    <option value="1">Icons (small brand logos)</option>
                                </select>
                                &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=427')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
                            </td>
                        </tr>
                        <tr>
                            <td valign="top">Display Products:</td>
                            <td>
                                <select name="pcProductsView" tabindex="302">
                                    <option value=""><%=dictLanguageCP.Item(Session("language")&"_cpCommon_315")%></option>
                                    <option value="h"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_517")%></option>
                                    <option value="p"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_518")%></option>
                                    <option value="l"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_519")%></option>
                                    <option value="m"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_520")%></option>
                                </select>
                                &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=427')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
                            </td>
                        </tr>
                        <tr>
                            <td>Active:</td>
                            <td><input type="radio" name="pcBrandsActive" value="1" class="clearBorder" checked="checked" tabindex="303"> Yes <input type="radio" name="pcBrandsActive" value="0" class="clearBorder" tabindex="303"> No</td>
                        </tr>
					</table>
					
				</div>
			<%
			'// =========================================
			'// THIRD PANEL - END
			'// =========================================

			'// =========================================
			'// FOURTH PANEL - START - Meta Tags
			'// =========================================
			%>
				<div class="TabbedPanelsContent">

					<table class="pcCPcontent">	

						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>  
						<tr>
							<th colspan="2">Brand Meta Tags</th>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<td colspan="2">Enter Meta Tags specific to this brand.&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=204')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
						</tr>
						<tr>
							<td align="right" valign="top">Title</td>
							<td><textarea name="pcBrandsMetaTitle" cols="50" rows="2" tabindex="401"></textarea>
						</tr>
						<tr>
							<td align="right" valign="top">Description</td>
							<td><textarea name="pcBrandsMetaDesc" cols="50" rows="6" tabindex="402"></textarea>
						</tr>
						<tr>
							<td align="right" valign="top">Keywords</td>
							<td><textarea name="pcBrandsMetaKeywords" cols="50" rows="4" tabindex="403"></textarea>
						</tr>
					
					</table>
					
				</div>
			<%
			'// =========================================
			'// FOURTH PANEL - END
			'// =========================================
			
			%>
			
			</div>
			
		</div>
		<%
		'// TABBED PANELS - MAIN DIV END
		%>

	<div style="clear: both;">&nbsp;</div>
  <script type="text/javascript">
		<!--
		var TabbedPanels1 = new Spry.Widget.TabbedPanels("TabbedPanels1", {defaultTab: params.tab ? params.tab : 0});
		//-->
  </script>

</form>

<!--#include file="AdminFooter.asp"-->
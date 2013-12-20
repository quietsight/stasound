<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/stringfunctions.asp"-->
<%
dim f, query, conntemp, rs
call opendb()
%>
<% pageTitle=dictLanguageCP.Item(Session("language")&"_cpInstCat_0") %>
<% section="products" %>
<!--#include file="AdminHeader.asp"-->
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
		if (theForm.categoryDesc.value == "")
			{
				 alert("Please enter a name for this category.");
					theForm.categoryDesc.focus();
					return (false);
			}
	return (true);
	}

	function chgWin(file,window)
	{
		msgWindow=open(file,window,'scrollbars=yes,resizable=yes,width=500,height=500');
		if (msgWindow.opener == null) msgWindow.opener = self;
	}
//--></script>
<form action="instCatb.asp" method="post" name="hForm" onSubmit="return Form1_Validator(this)" class="pcForms">
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
				<li class="TabbedPanelsTab" tabindex="300">Display Settings</li>				
				<li class="TabbedPanelsTab" tabindex="400">Other Settings</li>
				<li class="TabbedPanelsTab" tabindex="500">Meta Tags</li>
				<li class="TabbedPanelsTabButtons" tabindex="1000">
					<input type="hidden" name="reqstr" value="<%=request.QueryString("reqstr")%>">
					<input name="Submit" type="submit" value="Add" class="submit2">
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
							<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_1")%></th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<td width="20%" align="right"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_157")%>:</td>
							<td width="80%"><input name="categoryDesc" type="text" value="" size="40" tabindex="101"></td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_2")%>:</td>
							<td>
								<input type="text" name="image" value="" size="40" tabindex="102"><a href="#" onClick="chgWin('../pc/imageDir.asp?ffid=image&fid=hForm','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=439')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
							</td>
						</tr>
						<tr> 
							<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_3")%>:</td>
							<td> 
				        <input type="text" name="largeimage" value="<%=plargeImage%>" size="40" tabindex="103"><a href="#" onClick="chgWin('../pc/imageDir.asp?ffid=largeimage&fid=hForm','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=439')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
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
						<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_158")%>:</td>
						<td>
							<%
							cat_DropDownName="idParentCategory"
							cat_Type="0"
							cat_DropDownSize="1"
							cat_MultiSelect="0"
							cat_ExcBTOHide="0"
							cat_StoreFront="0"
							cat_ShowParent="1"
							cat_DefaultItem=""
							cat_SelectedItems="1,"
							cat_ExcItems=""
						
							%>
							<!--#include file="../includes/pcCategoriesList.asp"-->
							<%call pcs_CatList()%>
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
							<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_5")%>:&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=440')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr valign="top">
							<td><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_7")%>:
								<div class="small"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_8")%></div>
								<br />
								<input type="button" value="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_401")%>" onClick="newWindow('pop_HtmlEditor.asp?fi=SDesc','window2')" class="ibtnGrey">	
							</td>			
							<td>
								<textarea name="SDesc" cols="50" rows="6"></textarea>
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr valign="top">
							<td><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_9")%>:
							<div class="small"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_10")%></div>
							<br />
							<input type="button" value="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_401")%>" onClick="newWindow('pop_HtmlEditor.asp?fi=LDesc','window2')" class="ibtnGrey">
							</td>
							<td>
							<textarea name="LDesc" cols="50" rows="6" tabindex="203"></textarea>						
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td align="right">
							<input type="checkbox" name="HideDesc" value="1" class="clearBorder" tabindex="205">
							</td>
							<td>Do not show category descriptions</td>
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
						<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_11")%></th>
					</tr>
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<tr> 
						<td colspan="2">
						<%=dictLanguageCP.Item(Session("language")&"_cpInstCat_12")%><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_13")%><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_14")%>
			
						</td>
					</tr>
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<tr>
						<td align="right" width="30%">Display Subcategories:</td>
						<td width="70%">
							<select name="intSubCategoryView" tabindex="301">
								<option value="3"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_315")%></option>
								<option value="2"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_507")%></option>
								<option value="0"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_506")%></option>
								<option value="1"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_505")%></option>
								<option value="4">Thumbnails only</option>
							</select>
							&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=427')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
						</td>
					</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
					<tr>
						<td colspan="2">The following settings apply when categories are not displayed in a drop-down (if empty or 0, the default <a href="AdminSettings.asp?tab=3">store-wide setting</a> is used):</td>
					</tr>
					<tr> 
						<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_508")%>:</td>
						<td align="left"><input type="text" name="intCategoryColumns" value="<%=intCategoryColumns%>" tabindex="302">
						</td>
					</tr>
					<tr> 
						<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_509")%>:</td>
						<td align="left"> 
						<input type="text" name="intCategoryRows" value="<%=intCategoryRows%>" tabindex="302">
						</td>
					</tr>
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<tr>
						<td colspan="2"><hr></td>
					</tr>
					<tr> 
						<td colspan="2">
						<%=dictLanguageCP.Item(Session("language")&"_cpInstCat_17")%><a href="editCategories.asp?nav=&lid=<%=pIdCategory%>" target="_blank"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_18")%></a><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_19")%>:
						</td>
					</tr>
					<tr>
						<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_15")%>:</td>
						<td>
							<select name="strPageStyle" tabindex="303">
								<option value=""><%=dictLanguageCP.Item(Session("language")&"_cpCommon_315")%></option>
								<option value="h"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_517")%></option>
								<option value="p"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_518")%></option>
								<option value="l"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_519")%></option>
								<option value="m"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_520")%></option>
							</select>
							&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=429')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
						</td>
					</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
					<tr>
						<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_16")%>:</td>
					</tr>
					<tr> 
						<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_514")%>:</td>
						<td align="left"><input type="text" name="intProductColumns" value="" tabindex="304">
						</td>
					</tr>
					<tr> 
						<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_509")%>:</td>
						<td align="left"> 
						<input type="text" name="intProductRows" value="" tabindex="304">
						</td>
					</tr>
					<tr>
						<td class="pcCPspacer" colspan="2" style="height: 25px;"></td>
					</tr>  
					<tr>
						<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_23")%>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=424')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></th>
					</tr>
					<tr>
						<td class="pcCPspacer" colspan="2"></td>
					</tr>
					<tr>
						<td colspan="2">Choose a display option for the product details page. It will apply to all products within this category. This option <strong>overrides</strong> the <a href="AdminSettings.asp?tab=3">correspondng storewide setting</a>. This is a setting that can also be defined at the product level when adding/editing products.</td>
					</tr>
					<tr> 
						<td colspan="2">  
						 <input type="radio" name="CatDisplayLayout" value="C" class="clearBorder" tabindex="305"> <%=dictLanguageCP.Item(Session("language")&"_cpCommon_502")%></td>
					</tr>
					<tr> 
						<td colspan="2">  
						 <input type="radio" name="CatDisplayLayout" value="L" class="clearBorder" tabindex="306"> <%=dictLanguageCP.Item(Session("language")&"_cpCommon_503")%></td>
					</tr>
					<tr> 
						<td colspan="2">  
						<input type="radio" name="CatDisplayLayout" value="O" class="clearBorder" tabindex="307"> <%=dictLanguageCP.Item(Session("language")&"_cpCommon_504")%></td>
					</tr>
                    <tr> 
                        <td colspan="2">  
                        <input type="radio" name="CatDisplayLayout" value="D" checked class="clearBorder" tabindex="312"> Use store's default value</td>
                    </tr>

					</table>
					
				</div>
			<%
			'// =========================================
			'// THIRD PANEL - END
			'// =========================================
			
			'// =========================================
			'// FOURTH PANEL - START - Other settings
			'// =========================================
			%>
				<div class="TabbedPanelsContent">

					<table class="pcCPcontent">	
					
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Other Settings</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td colspan="2">Restrict the visibility of this category (the products that it contains are also hidden):</td>
						</tr>
						<tr> 
							<td align="right">
								<input type="checkbox" name="iBTOhide" value="1" class="clearBorder" tabindex="401">
							</td>
							<td>Hide this category in the storefront</td>
						</tr>
						<tr> 
							<td align="right"><input type="checkbox" name="RetailHide" value="1" class="clearBorder" tabindex="402"></td>
							<td>Hide this category in the storefront from retail customers (wholesale customers can see it)</td>
						</tr>
						
					</table>
					
				</div>
			<%
			'// =========================================
			'// FOURTH PANEL - END
			'// =========================================
			
			'// =========================================
			'// FIFTH PANEL - START - Meta Tags
			'// =========================================
			%>
				<div class="TabbedPanelsContent">

					<table class="pcCPcontent">	

						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>  
						<tr>
							<th colspan="2">Category Meta Tags</th>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<td colspan="2">Enter Meta Tags specific to this category.&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=204')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
						</tr>
						<tr>
							<td align="right" valign="top">Title</td>
							<td><textarea name="CatMetaTitle" cols="50" rows="2" tabindex="501"></textarea>
						</tr>
						<tr>
							<td align="right" valign="top">Description</td>
							<td><textarea name="CatMetaDesc" cols="50" rows="6" tabindex="502"></textarea>
						</tr>
						<tr>
							<td align="right" valign="top">Keywords</td>
							<td><textarea name="CatMetaKeywords" cols="50" rows="4" tabindex="503"></textarea>
						</tr>
					
					</table>
					
				</div>
			<%
			'// =========================================
			'// FIFTH PANEL - END
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
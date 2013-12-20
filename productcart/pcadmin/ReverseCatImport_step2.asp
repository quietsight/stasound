<% pageTitle = "Reverse Category Import Wizard - Choose Export Fields" %>
<% section = "products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
if session("cp_revCatImport_catlist")="" then
	response.redirect "ReverseCatImport_step1.asp"
end if
%>
<!--#include file="AdminHeader.asp"-->
<FORM name="checkboxform" method="post" action="ReverseCatImport_step3.asp" class="pcForms">
<table class="pcCPcontent">
<tr><td align="right" width="5%"><input type="checkbox" name="C1" value="1" checked class="clearBorder"></td><td width="95%">Category ID</td></tr>
<tr><td align="right"><input type="checkbox" name="C2" value="1" checked class="clearBorder"></td><td>Category Name</td></tr>
<tr><td align="right"><input type="checkbox" name="C3" value="1" checked class="clearBorder"></td><td>Small Image</td></tr>
<tr><td align="right"><input type="checkbox" name="C4" value="1" checked class="clearBorder"></td><td>Large Image</td></tr>
<tr><td align="right"><input type="checkbox" name="C5" value="1" checked class="clearBorder"></td><td>Parent Category Name</td></tr>
<tr><td align="right"><input type="checkbox" name="C6" value="1" checked class="clearBorder"></td><td>Parent Category ID</td></tr>
<tr><td align="right"><input type="checkbox" name="C7" value="1" checked class="clearBorder"></td><td>Category Short Description</td></tr>
<tr><td align="right"><input type="checkbox" name="C8" value="1" checked class="clearBorder"></td><td>Category Long Description</td></tr>
<tr><td align="right"><input type="checkbox" name="C9" value="1" checked class="clearBorder"></td><td>Hide Category Description</td></tr>
<tr><td align="right"><input type="checkbox" name="C10" value="1" checked class="clearBorder"></td><td>Display Sub-Categories</td></tr>
<tr><td align="right"><input type="checkbox" name="C11" value="1" checked class="clearBorder"></td><td>Sub-Categories per Row</td></tr>
<tr><td align="right"><input type="checkbox" name="C12" value="1" checked class="clearBorder"></td><td>Sub-Category Rows per Page</td></tr>
<tr><td align="right"><input type="checkbox" name="C13" value="1" checked class="clearBorder"></td><td>Display Products</td></tr>
<tr><td align="right"><input type="checkbox" name="C14" value="1" checked class="clearBorder"></td><td>Products per Row</td></tr>
<tr><td align="right"><input type="checkbox" name="C15" value="1" checked class="clearBorder"></td><td>Product Rows per Page</td></tr>
<tr><td align="right"><input type="checkbox" name="C16" value="1" checked class="clearBorder"></td><td>Hide category</td></tr>
<tr><td align="right"><input type="checkbox" name="C17" value="1" checked class="clearBorder"></td><td>Hide category from retail customers</td></tr>
<tr><td align="right"><input type="checkbox" name="C18" value="1" checked class="clearBorder"></td><td>Product Details Page Display Option</td></tr>
<tr><td align="right"><input type="checkbox" name="C19" value="1" checked class="clearBorder"></td><td>Category Meta Tags - Title</td></tr>
<tr><td align="right"><input type="checkbox" name="C20" value="1" checked class="clearBorder"></td><td>Category Meta Tags - Description</td></tr>
<tr><td align="right"><input type="checkbox" name="C21" value="1" checked class="clearBorder"></td><td>Category Meta Tags - Keywords</td></tr>
<tr><td align="right"><input type="checkbox" name="C22" value="1" checked class="clearBorder"></td><td>Featured Sub-Category Name</td></tr>
<tr><td align="right"><input type="checkbox" name="C23" value="1" checked class="clearBorder"></td><td>Featured Sub-Category ID</td></tr>
<tr><td align="right"><input type="checkbox" name="C24" value="1" checked class="clearBorder"></td><td>Use Featured Sub-Category Image</td></tr>
<tr><td align="right"><input type="checkbox" name="C25" value="1" checked class="clearBorder"></td><td>Category Order</td></tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td class="cpLinksList" colspan="2">
			<a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a>
			<script language="JavaScript">
			<!--
				function checkAll() {
					var theForm, z = 0;
					theForm = document.checkboxform;
					 for(z=0; z<theForm.length;z++){
					  if(theForm[z].type == 'checkbox'){
					  theForm[z].checked = true;
					  }
					}
				}
				 
				function uncheckAll() {
					var theForm, z = 0;
					theForm = document.checkboxform;
					 for(z=0; z<theForm.length;z++){
					  if(theForm[z].type == 'checkbox'){
					  theForm[z].checked = false;
					  }
					}
				}
			//-->
			</script>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"><hr></td>
	</tr>
	<tr>
		<td colspan="2"><input type="submit" name="submit" value=" Export categories " class="submit2" onclick="javascript: if (testCheckBox()) { return(confirm('You are about to export the selected categories fields. Are you sure you want to complete this action?')); } else { return(false); }"></td>
	</tr>
</table>
</FORM>
<!--#include file="AdminFooter.asp"-->
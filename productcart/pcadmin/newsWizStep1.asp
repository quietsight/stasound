<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=7%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
dim rstemp, conntemp, query
call opendb()
%>
<% pageTitle="Newsletter Wizard: Select Customers" %>
<% section="mngAcc" %>
<!--#include file="AdminHeader.asp"-->
<script language="JavaScript" type="text/javascript" src="../includes/spry/xpath.js"></script>
<script language="JavaScript" type="text/javascript" src="../includes/spry/SpryData.js"></script>
<script type="text/javascript">
	var dscategories = new Spry.Data.XMLDataSet("pcSpryCategoriesXML.asp?CP=1&idRootCategory=<%=pcv_IdRootCategory%>", "categories/category",{sortOnLoad:"categoryDesc",sortOrderOnLoad:"ascending",useCache:false});
	dscategories.setColumnType("idCategory", "number");
	dscategories.setColumnType("idParentCategory", "number");
	var dsProducts = new Spry.Data.XMLDataSet("categories_productsxml.asp?x_idCategory={dscategories::idCategory}", "/root/row");
</script>

<form name="form1" method="post" action="newsWizStep1a.asp" class="pcForms">
<table class="pcCPcontent">
<tr>
	<td colspan="2">
		<table width="100%">
		<tr>
			<td width="5%" align="right"><img border="0" src="images/step1a.gif"></td>
			<td width="95%"><b>Select Customers</b></td>
		</tr>
		<tr>
			<td align="right"><img border="0" src="images/step2.gif"></td>
			<td><font color="#A8A8A8">Verify customers</font></td>
		</tr>
		<tr>
			<td align="right"><img border="0" src="images/step3.gif"></td>
			<td><font color="#A8A8A8">Enter message</font></td>
		</tr>
		<tr>
			<td align="right"><img border="0" src="images/step4.gif"></td>
			<td><font color="#A8A8A8">Test message</font></td>
		</tr>
		<tr>
			<td align="right"><img border="0" src="images/step5.gif"></td>
			<td><font color="#A8A8A8">Send message</font></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">Opted in</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2">These are customers that chose to receive information from you. Make sure to comply with local regulations to avoid sending messages that might be considered SPAM and could trigger fines (<a href="http://wiki.earlyimpact.com/productcart/customer-newsletters" target="_blank">More information</a>). If you wish to support opt-in/opt-out preferences on multiple lists (e.g. &quot;Technical Support&quot; vs. &quot;Promotions&quot;), consider upgrading to a professional e-mail marketing tool like <a href="http://www.earlyimpact.com/productcart/mailup/" target="_blank">MailUp</a>.</td>
</tr>
<tr>
	<td width="20%" align="right">Select one:</td>
	<td width="80%">
		<select name="SOptedIn">
			<option value="1">Opted in only</option>
			<option value="0">All customers</option>
		</select>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">Product purchased</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2">You can filter based on who did or did not purchase a specific product, or any product within a specific product category.</td>
</tr>
<tr>
	<td width="20%" align="right" valign="middle"><input type="radio" name="purchaseType" value="0" checked class="clearBorder" onClick="document.getElementById('selectProduct').style.display='none'; document.getElementById('selectCategory').style.display='none';"></td>
	<td width="80%" valign="middle">All customers</td>
</tr>
<tr>
	<td width="20%" align="right" valign="middle"><input type="radio" name="purchaseType" value="4" class="clearBorder" onClick="document.getElementById('selectProduct').style.display='none'; document.getElementById('selectCategory').style.display='none';"></td>
	<td width="80%" valign="middle">Customers who have not yet purchased anything</td>
</tr>
<tr>
	<td width="20%" align="right" valign="middle"><input type="radio" name="purchaseType" value="3" class="clearBorder" onClick="document.getElementById('selectProduct').style.display='none'; document.getElementById('selectCategory').style.display='none';"></td>
	<td width="80%" valign="middle">Customers who have purchased something (<em>regardless of what they purchased</em>)</td>
</tr>
<tr>
	<td width="20%" align="right" valign="middle"><input type="radio" name="purchaseType" value="1" class="clearBorder" onClick="document.getElementById('selectProduct').style.display=''; document.getElementById('selectCategory').style.display='';"></td>
	<td width="80%" valign="middle">Customers who purchased...</td>
</tr>
<tr>
	<td width="20%" align="right" valign="middle"><input type="radio" name="purchaseType" value="2" class="clearBorder" onClick="document.getElementById('selectProduct').style.display=''; document.getElementById('selectCategory').style.display='';"></td>
	<td width="80%" valign="middle">Customers who did not purchase...</td>
</tr>
<tr id="selectProduct" style="display: none;">
	<td width="20%" align="right" valign="top">Select a product:</td>
	<td width="80%" valign="middle">Narrow your product search by selecting the category first. Then select the Product from the drop-down.
    
    <div style="margin: 10px 0 10px 0;">
        <div spry:region="dscategories" id="categorySelector">
            <div spry:state="loading"><img src="images/pc_AjaxLoader.gif"/>&nbsp;Loading Categories...</div>
            <div spry:state="ready">
            <select spry:repeatchildren="dscategories" name="categorySelect" onchange="document.form1.SIDproduct.disabled = true; dscategories.setCurrentRowNumber(this.selectedIndex);">
                <option spry:if="{ds_RowNumber} == {ds_CurrentRowNumber}" value="{idCategory}" selected="selected">{categoryDesc} {pcCats_BreadCrumbs}</option>
                <option spry:if="{ds_RowNumber} != {ds_CurrentRowNumber}" value="{idCategory}">{categoryDesc} {pcCats_BreadCrumbs}</option>
            </select>
            </div>
        </div>
    </div>

    <div style="margin: 0 0 10px 0;">
        <div spry:region="dsProducts" id="productSelector">
            <div spry:state="loading"><img src="images/pc_AjaxLoader.gif"/>&nbsp;Loading Products...</div>
            <div spry:state="ready">
            <select spry:repeatchildren="dsProducts" id="productSelect" name="SIDproduct">
                <option spry:if="{ds_RowNumber} == {ds_CurrentRowNumber}" value="{idProduct}" selected="selected">{description}</option>
                <option spry:if="{ds_RowNumber} != {ds_CurrentRowNumber}" value="{idProduct}">{description}</option>
            </select>
            </div>
        </div>    
    </div>    
  </td>
</tr>
<tr id="selectCategory" style="display: none;">
	<td align="right">or a category: </td>
	<td>
		<select name="SIDCategory">
			<option value="0">Any</option>
			<%
			query="SELECT idcategory, idParentCategory, categorydesc FROM categories WHERE idcategory<>1 and iBTOHide<>1 ORDER BY categoryDesc ASC"
			set rstemp=conntemp.execute(query)
			if err.number <> 0 then
				set rstemp=nothing
				call closeDb()
				response.redirect "techErr.asp?error="& Server.Urlencode("Error in retreiving categories from database: "&Err.Description) 
			end If
			if rstemp.eof then
				catnum="0"
				rstemp=nothing
			end If
			if catnum<>"0" then
				pcArr=rstemp.getRows()
				set rstemp=nothing
				intCount=ubound(pcArr,2)
				For i=0 to intCount
					idparentcategory=pcArr(1,i)
					if idparentcategory="1" then %>
					    <option value="<%response.write pcArr(0,i)%>"><%=pcArr(2,i)%></option>
				    <%else
					For j=0 to intCount
					if Clng(pcArr(0,j))=Clng(idparentcategory) then
					parentDesc=pcArr(2,j)%>
						<option value="<%response.write pcArr(0,i)%>"><%response.write pcArr(2,i)&" ["&parentDesc&"]"%></option>
					<%
					exit for
					end if 
					Next
					end if
				Next
			End if
			%>
		</select>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">Customer type</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2">You can include all customers, only retail customers, or only wholesale customers.</td>
</tr>
<tr>
	<td width="20%" align="right">Select customer type:</td>
	<td width="80%">
		<select name="SCustType">
			<option value="0">Any</option>
			<option value="1">Only Retail Customers</option>
			<option value="2">Only Whoselale Customers</option>
			<% 'START CT ADD %>
					<% 'if there are PBP customer type categories - List them here 
					
					query="SELECT idcustomerCategory, pcCC_Name, pcCC_CategoryType FROM pcCustomerCategories;"
					SET rs=Server.CreateObject("ADODB.RecordSet")
					SET rs=conntemp.execute(query)
					if NOT rs.eof then 
						do until rs.eof 
							intIdcustomerCategory=rs("idcustomerCategory")
							strpcCC_Name=rs("pcCC_Name")
							%>
							<option value='CC_<%=intIdcustomerCategory%>'
							<%if Session("pcAdmincustomertype")="CC_"&intIdcustomerCategory then 
								response.write "selected"
							end if%>
							><%="Only " & strpcCC_Name%></option>
							<% rs.moveNext
						loop
					end if
					SET rs=nothing
					
					call closeDb()
					
			'END CT ADD %>
		</select>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">Date range</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2">You can include only customers that have made a purchase within a certain date range.</td>
</tr>
<tr>
	<td align="right">Start date:</td>
	<td><input type="text" name="SStartDate" size="20"> (mm/dd/yyyy)</td>
</tr>
<tr>
	<td align="right">End date:</td>
	<td><input type="text" name="SEndDate" size="20"> (mm/dd/yyyy)</td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr>
	<td colspan="2" align="center">
		<input type="submit" name="submit" value="Continue" class="submit2">
		&nbsp;<input type="button" name="back" value="Back" onClick="javascript:history.back()">
	</td>
</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->
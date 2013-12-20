<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=7%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% pageTitle="Newsletter Wizard - STEP 1: Create Targeted Group" %>
<% section="mngAcc" %>
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp"-->
<%
dim rstemp, conntemp, query
call opendb()
%>
<script language="JavaScript" type="text/javascript" src="../includes/spry/xpath.js"></script>
<script language="JavaScript" type="text/javascript" src="../includes/spry/SpryData.js"></script>
<script type="text/javascript">
	var dscategories = new Spry.Data.XMLDataSet("pcSpryCategoriesXML.asp?CP=1&idRootCategory=<%=pcv_IdRootCategory%>", "categories/category",{sortOnLoad:"categoryDesc",sortOrderOnLoad:"ascending",useCache:false});
	dscategories.setColumnType("idCategory", "number");
	dscategories.setColumnType("idParentCategory", "number");
	var dsProducts = new Spry.Data.XMLDataSet("categories_productsxml.asp?x_idCategory={dscategories::idCategory}", "/root/row");
</script>

<form name="form1" method="post" action="mu_newsWizStep1a.asp" class="pcForms">
<table class="pcCPcontent">
<tr>
	<td colspan="2">
		<img src="images/pc2008_MailUp_Wizard.gif" alt="Newsletter Wizard - MailUp Integration" style="margin-bottom: 10px;" />
	</td>
</tr>
<tr>
	<th colspan="2">MailUp Lists</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2">When you send a newsletter, you do so to an entire list or a subset of a list. A subset of a list is called a <strong>Group</strong>. <strong>First</strong> select the list that your message pertains to (e.g. don't send a promotion to the &quot;Technical Support&quot; list). </td>
</tr>
<tr>
	<td width="20%" align="right">Select a distribution list:</td>
	<td width="80%">
		<select name="MailUpListID">
		<%query="SELECT pcMailUpLists_ID,pcMailUpLists_ListName FROM pcMailUpLists WHERE pcMailUpLists_Active=1 AND pcMailUpLists_Removed=0 ORDER BY pcMailUpLists_ListName;"
		set rs=connTemp.execute(query)
		if not rs.eof then
			tmpArr=rs.getRows()
			intCount=ubound(tmpArr,2)
			For i=0 to intCount%>
			<option value="<%=tmpArr(0,i)%>"><%=tmpArr(1,i)%></option>
			<%Next
		end if
		set rs=nothing%>
		</select>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">Opt-in Status</th>
</tr>
<tr>
	<td colspan="2">This refers to the <u>opt-in status in ProductCart</u>. If customers opt-out using the opt-out link in messages sent via MailUp, they will be unsubscribed in your MailUp console and will not receive new messages sent to the list, regardless of whether you include them in the group exported to your MailUp console or not. If you are creating a list to use elsewhere (i.e. not with MailUp), make sure to comply with consumer privacy regulations (e.g. <a href="http://www.ftc.gov/bcp/conline/pubs/buspubs/canspam.shtm" target="_blank">CAN-SPAM Act</a>).</td>
</tr>
<tr>
	<td width="20%" align="right">Select one:</td>
	<td width="80%">
		<select name="SOptedIn">
			<option value="1">Opted in only</option>
			<option value="2">Not Opted in only</option>
			<option value="0">All customers</option>
		</select>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<%query="SELECT pcMailUpSavedGroups_ID,pcMailUpSavedGroups_Name FROM pcMailUpSavedGroups ORDER BY pcMailUpSavedGroups_Name;"
set rs=connTemp.execute(query)
if not rs.eof then
	tmpArr=rs.getRows()
	intCount=ubound(tmpArr,2)%>
	<tr>
	<th colspan="2">Previously Saved</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2">You can use a previously saved set of customers, or create a new set by using the filters below. Sets of customers that you have saved in the past are shown in the drop-down below.</td>
	</tr>
	<tr>
	<td colspan="2">Use a previously saved set of customers: <select name="savedGroupID">
		<%For i=0 to intCount%>
		<option value="<%=tmpArr(0,i)%>"><%=tmpArr(1,i)%></option>
		<%Next%>
	</select>
	&nbsp;<input type="button" name="loadsaved" value="Load customers" class="submit2" onclick="location='mu_loadsavedgroup.asp?idgroup='+document.form1.savedGroupID.value+'&listid='+document.form1.MailUpListID.value;">
	<br /><br />
	</td>
	</tr>
	<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<%end if
set rs=nothing%>
<tr>
	<th colspan="2">Filter Customers</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2" class="pcCPsectionTitle">Product purchased</td>
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
	<td colspan="2" class="pcCPsectionTitle">Customer type</td>
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
					
			'END CT ADD %>
		</select>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2" class="pcCPsectionTitle">Customer Location</td>
</tr>
<tr>
	<td colspan="2">You can filter customers by location (Country, State/Province, ZIP Code)</td>
</tr>
<tr>
	<%
						'///////////////////////////////////////////////////////////
						'// START: COUNTRY AND STATE/ PROVINCE CONFIG
						'///////////////////////////////////////////////////////////
						' 
						' 1) Place this section ABOVE the Country field
						' 2) Note this module is used on multiple pages. Transfer your local variable into this rountine via the section below.
						' 3) Additional Required Info
						
						'// #2 Transfer local variable into this rountine here. (Format: Required Variable = Local Variable)
						pcv_isStateCodeRequired =  false
						pcv_isProvinceCodeRequired =  false
						pcv_isCountryCodeRequired =  false
						
						'// #3 Additional Required Info
						pcv_strTargetForm = "form1" '// Name of Form
						pcv_strCountryBox = "CustCountryCode" '// Name of Country Dropdown
						pcv_strTargetBox = "CustStateCode" '// Name of State Dropdown
						pcv_strProvinceBox =  "CustProvince" '// Name of Province Field
						
						Session(pcv_strSessionPrefix&pcv_strCountryBox) = ""
						Session(pcv_strSessionPrefix&pcv_strTargetBox) = ""
						Session(pcv_strSessionPrefix&pcv_strProvinceBox) = ""
						%>					
						<!--#include file="../includes/javascripts/pcStateAndProvince.asp"-->
						<%
						'///////////////////////////////////////////////////////////
						'// END: COUNTRY AND STATE/ PROVINCE CONFIG
						'///////////////////////////////////////////////////////////
						%>
	
						<%
						'// Display the Country Dropdown (/includes/javascripts/pcStateAndProvince.asp)
						pcs_CountryDropdown
						%>
						<%
						'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
						pcs_StateProvince
						%>
	<td width="20%">Zip Code:</td>
	<td width="80%">&nbsp;<input type="text" name="CustZipCode" size="8">
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2" class="pcCPsectionTitle">Customer Creation Date</td>
</tr>
<tr>
	<td colspan="2">You can filter customers by their creation date.</td>
</tr>
<tr>
	<td align="right">Start date:</td>
	<td><input type="text" name="CustStartDate" size="20"> (mm/dd/yyyy)</td>
</tr>
<tr>
	<td align="right">End date:</td>
	<td><input type="text" name="CustEndDate" size="20"> (mm/dd/yyyy)</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2" class="pcCPsectionTitle">Total Amount Ordered</td>
</tr>
<tr>
	<td colspan="2">You can filter customers by how much they have ordered.</td>
</tr>
<tr>
	<td align="right">Total ordered in <%=scCurSign%>:</td>
	<td><select name="CustTOType">
		<option value="0">&gt;</option>
		<option value="1">&lt;</option>
		<option value="2">&gt;=</option>
		<option value="3">&lt;=</option>
		</select>
		<input type="text" name="CustTotalAmount" size="20">
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2" class="pcCPsectionTitle">Ordered Date range</td>
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
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2" align="center">
		<input type="submit" name="submit" value="Continue" class="ibtnGrey">
		&nbsp;<input type="button" name="back" value="Back" onClick="javascript:history.back()" class="ibtnGrey">
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
</table>
</form>
<%call closedb()%><!--#include file="AdminFooter.asp"-->
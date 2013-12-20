<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
if request("pagetype")="1" then
	pcv_PageType=1
	pcv_Title="Drop-Shipper"
	pcv_Table="pcDropShipper"
else
	pcv_PageType=0
	pcv_Title="Supplier"
	pcv_Table="pcSupplier"
end if

pageTitle="Contact " & pcv_Title & "s - STEP 1: Create Targeted Group" %>
<% Section="products" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="AdminHeader.asp"-->

<%Dim connTemp, query, rs%>

<p class="pcCPsectionTitle">Select available <%=pcv_Title%>s</p>
	<table id="FindProducts" class="pcCPcontent">
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
				<select name="MailUpListID" id="MailUpListID" onChange="javascript: updURL(this.value);">
				<%
				call opendb()
				query="SELECT pcMailUpLists_ID,pcMailUpLists_ListName FROM pcMailUpLists WHERE pcMailUpLists_Active=1 AND pcMailUpLists_Removed=0 ORDER BY pcMailUpLists_ListName;"
				set rs=connTemp.execute(query)
				if not rs.eof then
					tmpArr=rs.getRows()
					intCount=ubound(tmpArr,2)
					For i=0 to intCount%>
						<option value="<%=tmpArr(0,i)%>" <%if (tmpArr(0,i) & ""=request("MailUpListID") & "") OR (tmpArr(0,i) & ""=session("CP_NW_ListID") & "") then%>selected<%end if%>><%=tmpArr(1,i)%></option>
					<%Next
				end if
				set rs=nothing%>
				</select>
                <script>
					var tmpURL="<%="mu_sds_newsWizStep1.asp?pagetype=" & pcv_PageType%>";
					var tmpURL1="<%="mu_sds_newsWizStep1a.asp?action=add&pagetype=" & pcv_PageType%>";
					function updURL(tmpvalue)
					{
						document.ajaxSearch.src_FromPage.value=tmpURL + "&MailUpListID=" + tmpvalue;
						document.ajaxSearch.src_ToPage.value=tmpURL1 + "&MailUpListID=" + tmpvalue;
					}
				</script>
			</td>
		</tr>
		<tr>
			<td colspan="2">&nbsp;</td>
		</tr>
		<tr>
			<td colspan="2">
			<%
				src_FormTitle1="Find " & pcv_Title & "s"
				src_FormTitle2="Contact " & pcv_Title & "s"
				src_FormTips1="Use the following filters to look for " & pcv_Title & "s in your store."
				src_FormTips2="Select one or more " & pcv_Title & "s that you want to contact:"
				src_DisplayType=1
				src_ShowLinks=0
				src_FromPage="mu_sds_newsWizStep1.asp?pagetype=" & pcv_PageType
				src_ToPage="mu_sds_newsWizStep1a.asp?action=add&pagetype=" & pcv_PageType
				src_Button1=" Search "
				src_Button2=" Continue"
				src_Button3=" Back "
				src_PageSize=15
				UseSpecial=0
				session("srcSDS_from")=""
				session("srcSDS_where")=""
				src_PageType=pcv_PageType
			%>
				<!--#include file="inc_srcSDSs.asp"-->
                <script>
					updURL(document.getElementById("MailUpListID").value);
				</script>
			</td>
		</tr>
	</table>
<!--#include file="AdminFooter.asp"-->
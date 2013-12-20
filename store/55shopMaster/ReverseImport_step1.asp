<% pageTitle = "Reverse Import Wizard - Locate products" %>
<% section = "products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
totalrecords=0
pcv_intExportSize = 300
Dim connTemp,query
call opendb()
Set rs=Server.CreateObject("ADODB.Recordset")
query="SELECT idproduct FROM Products WHERE removed=0;"
rs.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

if not rs.eof then
	totalrecords=clng(rs.RecordCount)
end if
set rs=nothing
call closedb()%>
<FORM action="ReverseImport_step1a.asp" method="post" class="pcForms">
<table class="pcCPcontent">
<tr>
	<td colspan="2" class="pcCPspacer">
		<% ' START show message, if any %>
            <!--#include file="pcv4_showMessage.asp"-->
        <% 	' END show message %>
	</td>
</tr>
<tr>
	<th colspan="2">Step 1: Choose the Products to Export</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2"><b>Your store has <%=totalrecords%> products</b>. <% if totalrecords>1000 then %><br />We recommend that you don't select &quot;All products&quot; as your store has a high number of products and the script could time out. Specifically: the more fields you decide to export, and the higher the number of products, the stronger is the chance that you will run into a time-out problem.<% end if %></td>
</tr>
<tr>
	<td align="right">
		<input type=radio name="prdlist" value="ALL1" class="clearBorder">
	</td>
	<td>
		All products
	</td>
</tr>
<tr>
	<td align="right">
		<input type=radio name="prdlist" value="" class="clearBorder">
	</td>
	<td>
		Locate products that you would like to export
	</td>
</tr>
<tr>
	<td align="right">
		<input type=radio name="prdlist" value="ALL" checked class="clearBorder">
	</td>
	<td>
		All Products in a range: Record <select name="pagecurrent">
		<%pages=fix(totalrecords/pcv_intExportSize)
		if totalrecords>pages*pcv_intExportSize then
			pages=pages+1
		end if
		if (pages=1) then%>
			<option value="1">1 To&nbsp;<%=totalrecords%></option>
		<%else
		For i=1 to pages-1
			if i=1 then%>
				<option value="1">1 To <%=pcv_intExportSize%></option>
			<%else%>
				<option value="<%=i%>"><%=((i-1)*pcv_intExportSize)+1%> - <%=i*pcv_intExportSize%></option>
			<%end if
		Next%>
		<option value="<%=i%>"><%=((pages-1)*pcv_intExportSize)+1%> To <%=totalrecords%></option>
		<%end if%>
		</select>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></th>
</tr>
<tr>
	<th colspan="2">Step 2: Choose a Field Separator</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2">If creating a comma- or TAB-separated file generates issues (e.g. the separator exists in the product description, thus rendering the exported file improperly formatted), try using a different separator. When you import the file into an application such as MS Excel, you will be able to indicate which separator should be used to part it.</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></th>
</tr>
<tr>
	<td align="right">
		<input type=radio name="fseparator" value="0" class="clearBorder" checked>
	</td>
	<td>
		Comma
	</td>
</tr>
<tr>
	<td align="right">
		<input type=radio name="fseparator" value="1" class="clearBorder">
	</td>
	<td>
		TAB
	</td>
</tr>
<tr>
	<td align="right">
		<input type=radio name="fseparator" value="2" class="clearBorder">
	</td>
	<td>
		Other: <input type="text" name="cseparator" size="5" value="">&nbsp;(e.g. &quot;|&quot;)
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"><hr></td>
</tr>
<tr>
	<td colspan="2"><input type="submit" name="Submit" value="Continue" class="submit2"></td>
</tr>
</table>
<input name="ExportSize" type="hidden" value="<%=pcv_intExportSize%>">
</FORM>
<!--#include file="AdminFooter.asp"-->
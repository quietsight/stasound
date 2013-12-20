<% pageTitle = "Export to Yahoo! Search Marketing - Locate products" %>
<% section = "specials" %>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
totalrecords=0
Dim connTemp,query
call opendb()
Set rs=Server.CreateObject("ADODB.Recordset")
query="SELECT idproduct FROM Products WHERE removed=0 AND active<>0 AND configonly=0;"
rs.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

if not rs.eof then
	totalrecords=clng(rs.RecordCount)
end if
set rs=nothing
call closedb()
%>
<FORM action="pcYahoo_step1a.asp" method="post" class="pcForms">

<table class="pcCPcontent">
<tr>
	<td colspan="2">
		<table border="0" cellpadding="0" cellspacing="0" width="60%">
		<tr>
			<td width="16%" align="center"><img border="0" src="images/step1a.gif"></td>
			<td width="84%"><b>Locate products</b></td>
		</tr>
		<tr>
			<td width="16%" align="center"><img border="0" src="images/step2.gif"></td>
			<td width="84%"><font color="#A8A8A8">Export results</font></td>
		</tr>
	</table>
	<% ' START show message, if any %>
        <!--#include file="pcv4_showMessage.asp"-->
    <% 	' END show message %>
	</td>
</tr>
<tr>
	<td colspan="2">&nbsp;</td>
</tr>
<tr>
	<td colspan="2"><b>Your store has <%=totalrecords%> products</b></td>
</tr>
<tr>
	<td align="right">
		<input type=radio name="prdlist" value="ALL1" class="clearBorder">
	</td>
	<td>
		All products in your store
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
		<%pages=fix(totalrecords/200)
		if totalrecords>pages*200 then
			pages=pages+1
		end if
		if (pages=1) then%>
			<option value="1">1 To&nbsp;<%=totalrecords%></option>
		<%else
		For i=1 to pages-1
			if i=1 then%>
				<option value="1">1 To 200</option>
			<%else%>
				<option value="<%=i%>"><%=((i-1)*200)+1%> - <%=i*200%></option>
			<%end if
		Next%>
		<option value="<%=i%>"><%=((pages-1)*200)+1%> To <%=totalrecords%></option>
		<%end if%>
		</select>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td><input type="submit" name="Submit" value="Continue" class="submit2"></td>
</tr>
</FORM>
</table>
<!--#include file="AdminFooter.asp"-->
<% pageTitle = "Customer Import Wizard - Upload/Locate Data File" %>
<% section = "mngAcc" %>
<%PmAdmin=7%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
if request("action")="select" then
	if request("ways")="1" then
		response.redirect "custindex_import1.asp"
	else
		if request("ways")="2" then
			response.redirect "custindex_import2.asp"
		end if
	end if
end if
%>
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
	<tr>
    	<td colspan="2">
            <table class="pcCPcontent">
            <tr>
                <td colspan="2"><h2>Steps:</h2></td>
            </tr>
            <tr>
                <td width="5%" align="right"><img border="0" src="images/step1a.gif"></td>
                <td width="95%"><b>Select product data file</b></td>
            </tr>
            <tr>
                <td align="right"><img border="0" src="images/step2.gif"></td>
                <td><font color="#A8A8A8">Map fields</font></td>
            </tr>
            <tr>
                <td align="right"><img border="0" src="images/step3.gif"></td>
                <td><font color="#A8A8A8">Confirm mapping</font></td>
            </tr>
            <tr>
                <td align="right"><img border="0" src="images/step4.gif"></td>
                <td><font color="#A8A8A8">Import results</font></td>
            </tr>
            </table>
        
        <br /> 
        <% ' START show message, if any %>
            <!--#include file="pcv4_showMessage.asp"-->
        <% 	' END show message %>
        
    	<% if (request.querystring("nextstep")=1) then %>
        <form action="cwstep2.asp" method="post" class="pcForms">
        <table class="pcCPcontent">
		<tr>
			<td>
				<input type="radio" name="append" value="0" <%if session("append")<>"1" then%>checked<%end if%>>
				Import data to database
			</td>
		</tr>
		<tr>
			<td>
				<input type="radio" name="append" value="1" <%if session("append")<>"1" then
				else%>checked<%end if%>>
				Update current data if E-mail address is an exact match with an existing&nbsp;E-mail address<br><br>
			</td>
		</tr>
		<tr>
			<td align="right">
				<input type="submit" name="Go" value="Go to Step 2 >>" class="submit2">
			</td>
		</tr>
		</table>
		</form>
	<%else%>
		<p>Select how ProductCart should locate your data file. You can either upload the file now, or provide a location on your Web server if the file has already been uploaded. Please note that <strong>ONLY</strong> <strong>*.csv</strong> &amp; <strong>*.xls</strong> files are accepted. For more information on what fields can be imported and on how to prepare your *.csv &amp; *.xls files for import, please refer to the ProductCart <a href="http://wiki.earlyimpact.com/productcart/customer-import" target="_blank">User Guide</a>.<br /><br /><p>
		<form method="post" action="custindex_import.asp?action=select" class="pcForms">
		<table border="0" cellspacing="0" width="100%" cellpadding="4">
		<tr>
			<th colspan="2">Customer data file</th>
		</tr>
        <tr>
        	<td colspan="2" class="pcCPspacer"></td>
        </tr>
		<tr>
			<td width="5%">
				<input type=radio name="ways" value="1" checked>
			</td>
			<td width="95%">
				Use the Customer Import Wizard to upload the data file to the server
			</td>
		</tr>
		<tr>
			<td width="5%">
				<input type=radio name="ways" value="2">
			</td>
			<td width="95%">
				The data file has been manually uploaded to the server
			</td>
		</tr>
		<tr>
			<td width="498" colspan="2">
				<br>
				<input type="submit" name="submit" value="Select" class="submit2">
			</td>
		</tr>
	</table> 
	</form>
	<%end if%>        
	</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->